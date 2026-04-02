from __future__ import annotations

import csv
import io
import math
import re
import zipfile
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Iterable

import matplotlib
import pandas as pd
import xlsxwriter
from matplotlib import pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages
from matplotlib.ticker import MultipleLocator


matplotlib.use("Agg")

TIME_SUFFIX_RE = re.compile(r"\s*\(\d+\)\s*$")
EXCEL_NUMBER_RE = re.compile(r"^\d+\.0+$")
NON_FILENAME_RE = re.compile(r"[^A-Za-z0-9._-]+")

CHART_SPECS = [
    {
        "title": "Rundenzeiten",
        "y_axis": "Rundenzeit (s)",
        "column_key": "total_col",
        "value_attr": "lap_seconds",
        "anchor": "F3",
    },
    {
        "title": "Aufstiege",
        "y_axis": "Aufstieg (s)",
        "column_key": "ascent_col",
        "value_attr": "ascent_seconds",
        "anchor": "N3",
    },
    {
        "title": "Abfahrten",
        "y_axis": "Abfahrt (s)",
        "column_key": "descent_col",
        "value_attr": "descent_seconds",
        "anchor": "F22",
    },
]

ID_COLUMN_ALIASES = {
    "id",
    "rider id",
    "rider_id",
    "riderid",
    "athlete id",
    "athlete_id",
    "athleteid",
    "fahrer id",
    "fahrer_id",
    "bib",
    "startnummer",
}

NAME_COLUMN_ALIASES = {
    "name",
    "athlete",
    "athlete name",
    "athlete_name",
    "athletenname",
    "rider",
    "rider name",
    "rider_name",
    "fahrer",
    "fahrername",
}


@dataclass
class RawRecord:
    n: int
    athlete_id: str
    total_seconds: float | None
    ascent_seconds: float | None
    descent_seconds: float | None
    raw_total: str
    has_split: bool


@dataclass
class LapEntry:
    source_n: int
    lap_number: int
    lap_seconds: float | None
    ascent_seconds: float | None
    descent_seconds: float | None


@dataclass
class BlockData:
    index: int
    marker_n: int | None = None
    marker_seconds: float | None = None
    laps: list[LapEntry] = field(default_factory=list)

    @property
    def label(self) -> str:
        return f"Block {self.index}"

    @property
    def has_split_data(self) -> bool:
        return any(
            lap.ascent_seconds is not None or lap.descent_seconds is not None
            for lap in self.laps
        )


@dataclass
class AthleteData:
    athlete_id: str
    blocks: list[BlockData]
    athlete_name: str | None = None

    @property
    def lap_count(self) -> int:
        return sum(len(block.laps) for block in self.blocks)

    @property
    def split_blocks(self) -> int:
        return sum(1 for block in self.blocks if block.has_split_data)

    @property
    def display_name(self) -> str:
        return self.athlete_name or self.athlete_id

    @property
    def summary_label(self) -> str:
        if self.athlete_name:
            return f"{self.athlete_name} ({self.athlete_id})"
        return self.athlete_id


@dataclass
class FreelapDataset:
    session_label: str
    session_date: str | None
    athletes: list[AthleteData]


@dataclass
class MappingReport:
    matched: int
    unmatched_ids: list[str]
    unused_mapping_ids: list[str]


def decode_csv_bytes(file_bytes: bytes) -> str:
    for encoding in ("utf-8-sig", "cp1252", "latin-1"):
        try:
            return file_bytes.decode(encoding)
        except UnicodeDecodeError:
            continue
    raise UnicodeDecodeError("freelap", b"", 0, 1, "CSV encoding could not be detected")


def parse_seconds(value: str | None) -> float | None:
    if value is None:
        return None
    text = value.strip()
    if not text or text == "-":
        return None

    text = text.lstrip("+").replace(",", ".")
    text = TIME_SUFFIX_RE.sub("", text)
    parts = text.split(":")

    total = 0.0
    for part in parts:
        total = total * 60 + float(part)
    return round(total, 2)


def format_seconds(value: float | None) -> str:
    if value is None:
        return ""
    return f"{value:.2f}"


def parse_dataset(file_bytes: bytes) -> FreelapDataset:
    text = decode_csv_bytes(file_bytes)
    rows = list(csv.reader(io.StringIO(text), delimiter=";"))
    header_index = None

    for index, row in enumerate(rows):
        if len(row) >= 3 and row[:3] == ["N", "ID", "TEMPS"]:
            header_index = index
            break

    if header_index is None:
        raise ValueError("Die CSV-Struktur wurde nicht erkannt.")

    session_label = next(
        (row[0].strip() for row in rows[:header_index] if row and row[0].strip().startswith("Entra")),
        "Freelap Export",
    )
    session_date = _extract_session_date(rows[:header_index])

    grouped_records: dict[str, list[RawRecord]] = {}

    for row in rows[header_index + 1 :]:
        if not row or len(row) < 3:
            continue

        normalized = row + [""] * (6 - len(row))
        n_text, athlete_id, total_raw, _, ascent_raw, descent_raw = normalized[:6]
        athlete_id = athlete_id.strip()
        n_text = n_text.strip()

        if not athlete_id or not n_text.isdigit():
            continue

        record = RawRecord(
            n=int(n_text),
            athlete_id=athlete_id,
            total_seconds=parse_seconds(total_raw),
            ascent_seconds=parse_seconds(ascent_raw),
            descent_seconds=parse_seconds(descent_raw),
            raw_total=total_raw.strip(),
            has_split=bool(ascent_raw.strip() or descent_raw.strip()),
        )
        grouped_records.setdefault(athlete_id, []).append(record)

    athletes = [
        AthleteData(athlete_id=athlete_id, blocks=_build_blocks(records))
        for athlete_id, records in grouped_records.items()
    ]
    athletes = [athlete for athlete in athletes if athlete.lap_count]

    if not athletes:
        raise ValueError("Es wurden keine verwertbaren Runden gefunden.")

    return FreelapDataset(
        session_label=session_label,
        session_date=session_date,
        athletes=athletes,
    )


def parse_athlete_mapping(file_bytes: bytes, filename: str) -> dict[str, str]:
    suffix = Path(filename).suffix.lower()

    if suffix == ".csv":
        frame = pd.read_csv(io.BytesIO(file_bytes), sep=None, engine="python", dtype=str)
    elif suffix in {".xlsx", ".xls"}:
        frame = pd.read_excel(io.BytesIO(file_bytes), dtype=str)
    else:
        raise ValueError("Die Athletenliste muss als CSV oder Excel-Datei hochgeladen werden.")

    if frame.empty:
        raise ValueError("Die Athletenliste ist leer.")

    frame = frame.fillna("")
    id_column = _find_column(frame.columns, ID_COLUMN_ALIASES)
    name_column = _find_column(frame.columns, NAME_COLUMN_ALIASES)

    if id_column is None or name_column is None:
        raise ValueError(
            "Benötigte Spalten nicht gefunden. Erwartet werden z. B. 'Rider ID' und 'Athlete Name'."
        )

    mapping: dict[str, str] = {}
    for _, row in frame.iterrows():
        rider_id = _normalize_identifier(row.get(id_column, ""))
        athlete_name = _clean_name(row.get(name_column, ""))

        if rider_id and athlete_name:
            mapping[rider_id] = athlete_name

    if not mapping:
        raise ValueError("In der Athletenliste wurden keine gültigen Rider IDs mit Namen gefunden.")

    return mapping


def apply_athlete_names(dataset: FreelapDataset, athlete_mapping: dict[str, str]) -> FreelapDataset:
    athletes = [
        AthleteData(
            athlete_id=athlete.athlete_id,
            athlete_name=athlete_mapping.get(_normalize_identifier(athlete.athlete_id)),
            blocks=athlete.blocks,
        )
        for athlete in dataset.athletes
    ]
    return FreelapDataset(
        session_label=dataset.session_label,
        session_date=dataset.session_date,
        athletes=athletes,
    )


def build_mapping_report(dataset: FreelapDataset, athlete_mapping: dict[str, str]) -> MappingReport:
    matched_ids = {
        _normalize_identifier(athlete.athlete_id)
        for athlete in dataset.athletes
        if athlete_mapping.get(_normalize_identifier(athlete.athlete_id))
    }
    unmatched_ids = [
        athlete.athlete_id
        for athlete in dataset.athletes
        if _normalize_identifier(athlete.athlete_id) not in matched_ids
    ]
    unused_mapping_ids = sorted(set(athlete_mapping) - matched_ids)
    return MappingReport(
        matched=len(matched_ids),
        unmatched_ids=unmatched_ids,
        unused_mapping_ids=unused_mapping_ids,
    )


def build_workbook(dataset: FreelapDataset) -> bytes:
    output = io.BytesIO()
    workbook = xlsxwriter.Workbook(output, {"in_memory": True})

    formats = _build_formats(workbook)
    _write_overview_sheet(workbook, dataset, formats)

    used_sheet_names: set[str] = {"Uebersicht"}
    for athlete in dataset.athletes:
        sheet_name = _unique_sheet_name(athlete.display_name, used_sheet_names)
        used_sheet_names.add(sheet_name)
        worksheet = workbook.add_worksheet(sheet_name)
        _write_athlete_sheet(workbook, worksheet, athlete, dataset, formats)

    workbook.close()
    output.seek(0)
    return output.getvalue()


def build_athlete_pdf(dataset: FreelapDataset, athlete: AthleteData) -> bytes:
    output = io.BytesIO()
    with PdfPages(output) as pdf:
        any_chart = False
        for chart_spec in CHART_SPECS:
            figure = _build_chart_figure(athlete, dataset, chart_spec)
            if figure is None:
                continue
            any_chart = True
            pdf.savefig(figure, bbox_inches="tight")
            plt.close(figure)

        if not any_chart:
            figure, axis = plt.subplots(figsize=(10, 4))
            axis.axis("off")
            axis.text(
                0.5,
                0.5,
                f"Keine Diagrammdaten fuer {athlete.summary_label}.",
                ha="center",
                va="center",
                fontsize=14,
            )
            pdf.savefig(figure, bbox_inches="tight")
            plt.close(figure)

    output.seek(0)
    return output.getvalue()


def build_chart_pdf(dataset: FreelapDataset, athlete: AthleteData, chart_title: str) -> bytes:
    chart_spec = next((spec for spec in CHART_SPECS if spec["title"] == chart_title), None)
    if chart_spec is None:
        raise ValueError(f"Unbekanntes Diagramm: {chart_title}")

    figure = _build_chart_figure(athlete, dataset, chart_spec)
    if figure is None:
        raise ValueError(f"Keine Daten fuer {chart_title.lower()} von {athlete.summary_label}.")

    output = io.BytesIO()
    with PdfPages(output) as pdf:
        pdf.savefig(figure, bbox_inches="tight")
    plt.close(figure)
    output.seek(0)
    return output.getvalue()


def build_pdf_zip(dataset: FreelapDataset, export_mode: str) -> bytes:
    output = io.BytesIO()

    with zipfile.ZipFile(output, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for athlete in dataset.athletes:
            athlete_slug = _safe_filename(athlete.display_name)

            if export_mode == "per_athlete":
                archive.writestr(
                    f"{athlete_slug}.pdf",
                    build_athlete_pdf(dataset, athlete),
                )
                continue

            if export_mode == "per_chart":
                for chart_spec in CHART_SPECS:
                    try:
                        pdf_bytes = build_chart_pdf(dataset, athlete, str(chart_spec["title"]))
                    except ValueError:
                        continue
                    archive.writestr(
                        f"{athlete_slug}/{athlete_slug}_{_safe_filename(str(chart_spec['title']))}.pdf",
                        pdf_bytes,
                    )
                continue

            raise ValueError("Unbekannter PDF-Exportmodus.")

    output.seek(0)
    return output.getvalue()


def _build_formats(workbook: xlsxwriter.Workbook) -> dict[str, xlsxwriter.format.Format]:
    return {
        "title": workbook.add_format(
            {
                "bold": True,
                "font_size": 16,
                "font_name": "Aptos",
                "align": "left",
                "valign": "vcenter",
                "bg_color": "#D7EEF8",
                "border": 1,
            }
        ),
        "meta_label": workbook.add_format(
            {
                "bold": True,
                "font_name": "Aptos",
                "font_color": "#35586C",
            }
        ),
        "block_label": workbook.add_format(
            {
                "bold": True,
                "font_name": "Aptos",
                "bg_color": "#EAF6FB",
                "font_color": "#35586C",
            }
        ),
        "header": workbook.add_format(
            {
                "bold": True,
                "font_name": "Aptos",
                "align": "center",
                "valign": "vcenter",
                "bg_color": "#A9D8E9",
                "border": 1,
            }
        ),
        "cell": workbook.add_format(
            {
                "font_name": "Aptos",
                "border": 1,
                "align": "center",
                "valign": "vcenter",
            }
        ),
        "number": workbook.add_format(
            {
                "font_name": "Aptos",
                "border": 1,
                "align": "center",
                "valign": "vcenter",
                "num_format": "0.00",
            }
        ),
        "note": workbook.add_format(
            {
                "italic": True,
                "font_name": "Aptos",
                "font_color": "#6E7D86",
            }
        ),
        "overview_header": workbook.add_format(
            {
                "bold": True,
                "font_name": "Aptos",
                "bg_color": "#D7EEF8",
                "border": 1,
            }
        ),
        "overview_cell": workbook.add_format(
            {
                "font_name": "Aptos",
                "border": 1,
            }
        ),
    }


def _write_overview_sheet(
    workbook: xlsxwriter.Workbook,
    dataset: FreelapDataset,
    formats: dict[str, xlsxwriter.format.Format],
) -> None:
    worksheet = workbook.add_worksheet("Uebersicht")
    worksheet.set_zoom(90)
    worksheet.set_column("A:A", 28)
    worksheet.set_column("B:B", 16)
    worksheet.set_column("C:F", 16)

    worksheet.write("A1", "Freelap Excel Export", formats["title"])
    worksheet.write("A3", "Session", formats["meta_label"])
    worksheet.write("B3", dataset.session_label)
    worksheet.write("A4", "Datum", formats["meta_label"])
    worksheet.write("B4", dataset.session_date or "-")
    worksheet.write("A5", "Athleten", formats["meta_label"])
    worksheet.write("B5", len(dataset.athletes))

    header_row = 7
    headers = ["Athlet", "Rider ID", "Bloecke", "Runden", "Bloecke mit Split", "Sheet"]
    for column, label in enumerate(headers):
        worksheet.write(header_row, column, label, formats["overview_header"])

    for row_offset, athlete in enumerate(dataset.athletes, start=1):
        row = header_row + row_offset
        worksheet.write(row, 0, athlete.display_name, formats["overview_cell"])
        worksheet.write(row, 1, athlete.athlete_id, formats["overview_cell"])
        worksheet.write_number(row, 2, len(athlete.blocks), formats["overview_cell"])
        worksheet.write_number(row, 3, athlete.lap_count, formats["overview_cell"])
        worksheet.write_number(row, 4, athlete.split_blocks, formats["overview_cell"])
        worksheet.write(row, 5, athlete.display_name[:31], formats["overview_cell"])


def _write_athlete_sheet(
    workbook: xlsxwriter.Workbook,
    worksheet: xlsxwriter.worksheet.Worksheet,
    athlete: AthleteData,
    dataset: FreelapDataset,
    formats: dict[str, xlsxwriter.format.Format],
) -> None:
    worksheet.set_zoom(85)
    worksheet.freeze_panes(4, 0)
    worksheet.set_column("A:A", 10)
    worksheet.set_column("B:D", 14)
    worksheet.set_column("F:Q", 16)
    worksheet.set_row(0, 28)

    worksheet.merge_range("A1:D1", athlete.display_name, formats["title"])
    worksheet.write("A2", "Rider ID", formats["meta_label"])
    worksheet.write("B2", athlete.athlete_id)
    worksheet.write("C2", "Session", formats["meta_label"])
    worksheet.write("D2", dataset.session_label)
    worksheet.write("A3", "Datum", formats["meta_label"])
    worksheet.write("B3", dataset.session_date or "-")

    block_ranges = []
    current_row = 4
    for block in athlete.blocks:
        worksheet.merge_range(
            current_row,
            0,
            current_row,
            3,
            _block_heading(block),
            formats["block_label"],
        )
        current_row += 1

        headers = ["Tour", "Rundenzeit (s)", "Aufstieg (s)", "Abfahrt (s)"]
        for column, label in enumerate(headers):
            worksheet.write(current_row, column, label, formats["header"])

        data_start = current_row + 1
        data_end = current_row
        for lap in block.laps:
            data_end += 1
            worksheet.write_number(data_end, 0, lap.lap_number, formats["cell"])
            if lap.lap_seconds is not None:
                worksheet.write_number(data_end, 1, lap.lap_seconds, formats["number"])
            else:
                worksheet.write_blank(data_end, 1, None, formats["cell"])
            if lap.ascent_seconds is not None:
                worksheet.write_number(data_end, 2, lap.ascent_seconds, formats["number"])
            else:
                worksheet.write_blank(data_end, 2, None, formats["cell"])
            if lap.descent_seconds is not None:
                worksheet.write_number(data_end, 3, lap.descent_seconds, formats["number"])
            else:
                worksheet.write_blank(data_end, 3, None, formats["cell"])

        block_ranges.append(
            {
                "label": block.label,
                "first_row": data_start,
                "last_row": data_end,
                "lap_col": 0,
                "total_col": 1,
                "ascent_col": 2,
                "descent_col": 3,
                "has_split_data": block.has_split_data,
            }
        )
        current_row = data_end + 2

    _insert_charts(workbook, worksheet, athlete, block_ranges)


def _extract_session_date(rows: Iterable[list[str]]) -> str | None:
    row_list = list(rows)
    for index, row in enumerate(row_list):
        if len(row) >= 2 and row[0].strip() == "Date" and index + 1 < len(row_list):
            date_text = row_list[index + 1][0].strip()
            try:
                return datetime.strptime(date_text, "%m/%d/%Y").strftime("%Y-%m-%d")
            except ValueError:
                try:
                    return datetime.strptime(date_text, "%d/%m/%Y").strftime("%Y-%m-%d")
                except ValueError:
                    return date_text
    return None


def _find_column(columns: Iterable[object], aliases: set[str]) -> str | None:
    normalized_to_original = {_normalize_header(str(column)): str(column) for column in columns}
    for alias in aliases:
        original = normalized_to_original.get(_normalize_header(alias))
        if original:
            return original
    return None


def _clean_name(value: object) -> str:
    return str(value).strip()


def _normalize_header(value: str) -> str:
    return " ".join(re.split(r"[\s_-]+", value.strip().lower()))


def _normalize_identifier(value: object) -> str:
    text = str(value).strip()
    if not text:
        return ""
    if EXCEL_NUMBER_RE.fullmatch(text):
        text = text.split(".", 1)[0]
    return re.sub(r"\s+", "", text).lower()


def _start_block(blocks: list[BlockData], marker: RawRecord | None) -> BlockData:
    return BlockData(
        index=len(blocks) + 1,
        marker_n=marker.n if marker else None,
        marker_seconds=marker.total_seconds if marker else None,
    )


def _build_blocks(records: list[RawRecord]) -> list[BlockData]:
    blocks: list[BlockData] = []
    current_block: BlockData | None = None
    pending_marker: RawRecord | None = None
    previous_n: int | None = None

    for index, record in enumerate(records):
        next_record = records[index + 1] if index + 1 < len(records) else None
        is_marker_row = (
            not record.has_split
            and next_record is not None
            and next_record.has_split
            and next_record.n > record.n
        )

        if previous_n is not None and record.n <= previous_n and current_block and current_block.laps:
            current_block = None
            pending_marker = None

        if is_marker_row:
            pending_marker = record
            current_block = None
            previous_n = record.n
            continue

        if current_block is None:
            current_block = _start_block(blocks, pending_marker)
            blocks.append(current_block)
            pending_marker = None

        current_block.laps.append(
            LapEntry(
                source_n=record.n,
                lap_number=len(current_block.laps) + 1,
                lap_seconds=record.total_seconds,
                ascent_seconds=record.ascent_seconds,
                descent_seconds=record.descent_seconds,
            )
        )
        previous_n = record.n

    return [block for block in blocks if block.laps]


def _block_heading(block: BlockData) -> str:
    if block.marker_seconds is None:
        return block.label
    return f"{block.label} - Marker {format_seconds(block.marker_seconds)} s"


def _insert_charts(
    workbook: xlsxwriter.Workbook,
    worksheet: xlsxwriter.worksheet.Worksheet,
    athlete: AthleteData,
    block_ranges: list[dict[str, int | str | bool]],
) -> None:
    max_lap_number = max((lap.lap_number for block in athlete.blocks for lap in block.laps), default=1)

    for chart_spec in CHART_SPECS:
        values = []
        for block in athlete.blocks:
            for lap in block.laps:
                value = getattr(lap, str(chart_spec["value_attr"]))
                if value is not None:
                    values.append(value)

        anchor = str(chart_spec["anchor"])
        if not values:
            worksheet.write(anchor, f"Keine Daten fuer {str(chart_spec['title']).lower()}.")
            continue

        chart = workbook.add_chart({"type": "line"})
        chart.set_title(
            {
                "name": f"{chart_spec['title']} - {athlete.display_name}",
                "name_font": {"name": "Aptos", "size": 14},
            }
        )
        chart.set_legend({"position": "bottom", "font": {"name": "Aptos", "size": 9}})
        chart.set_size({"width": 640, "height": 320})
        chart.set_chartarea({"border": {"color": "#D5DDE0"}, "fill": {"color": "#FFFFFF"}})
        chart.set_plotarea({"border": {"color": "#D5DDE0"}, "fill": {"color": "#FFFFFF"}})
        chart.set_x_axis(
            {
                "name": "Runde",
                "name_font": {"name": "Aptos", "size": 10},
                "num_font": {"name": "Aptos", "size": 9},
                "min": 0,
                "max": max_lap_number,
                "major_unit": 1,
                "minor_unit": 1,
                "major_gridlines": {"visible": True, "line": {"color": "#D9E2E6"}},
            }
        )
        axis_config = _axis_bounds(values)
        chart.set_y_axis(
            {
                "name": chart_spec["y_axis"],
                "name_font": {"name": "Aptos", "size": 10},
                "num_font": {"name": "Aptos", "size": 9},
                "major_gridlines": {"visible": True, "line": {"color": "#D9E2E6"}},
                "minor_gridlines": {"visible": True, "line": {"color": "#EEF3F5"}},
                **axis_config,
            }
        )

        for block_range in block_ranges:
            column_key = str(chart_spec["column_key"])
            col_index = int(block_range[column_key])
            first_row = int(block_range["first_row"])
            last_row = int(block_range["last_row"])

            if column_key != "total_col" and not bool(block_range["has_split_data"]):
                continue

            chart.add_series(
                {
                    "name": str(block_range["label"]),
                    "categories": [worksheet.name, first_row, 0, last_row, 0],
                    "values": [worksheet.name, first_row, col_index, last_row, col_index],
                    "marker": {"type": "circle", "size": 5},
                    "line": {"width": 1.75},
                }
            )

        worksheet.insert_chart(anchor, chart)


def _chart_series(athlete: AthleteData, value_attr: str) -> list[dict[str, list[float] | str]]:
    series: list[dict[str, list[float] | str]] = []
    for block in athlete.blocks:
        laps = []
        values = []
        for lap in block.laps:
            value = getattr(lap, value_attr)
            if value is None:
                continue
            laps.append(float(lap.lap_number))
            values.append(float(value))

        if values:
            series.append({"label": block.label, "laps": laps, "values": values})
    return series


def _build_chart_figure(
    athlete: AthleteData,
    dataset: FreelapDataset,
    chart_spec: dict[str, str],
):
    series = _chart_series(athlete, str(chart_spec["value_attr"]))
    if not series:
        return None

    values = [value for item in series for value in item["values"]]
    axis_config = _axis_bounds(values)
    axis_min = axis_config["min"]
    axis_max = axis_config["max"]
    max_lap_number = max((int(max(item["laps"])) for item in series), default=1)

    figure, axis = plt.subplots(figsize=(10.5, 5.6))
    colors = ["#4F81BD", "#C0504D", "#9BBB59", "#8064A2", "#4BACC6", "#F79646"]

    for index, item in enumerate(series):
        axis.plot(
            item["laps"],
            item["values"],
            marker="o",
            linewidth=2,
            markersize=5,
            label=str(item["label"]),
            color=colors[index % len(colors)],
        )

    axis.set_title(
        f"{chart_spec['title']}",
        fontsize=18,
        fontweight="bold",
        color="#4F4F4F",
    )
    axis.set_xlabel("Runde", fontsize=11, color="#4F4F4F")
    axis.set_ylabel(str(chart_spec["y_axis"]))
    axis.set_xlim(0, max_lap_number)
    axis.set_ylim(axis_min, axis_max)
    axis.set_xticks(list(range(0, max_lap_number + 1)))
    axis.xaxis.set_major_locator(MultipleLocator(1))
    axis.yaxis.set_major_locator(MultipleLocator(axis_config["major_unit"]))
    axis.yaxis.set_minor_locator(MultipleLocator(axis_config["minor_unit"]))
    axis.grid(which="major", axis="both", color="#D9D9D9", linewidth=0.8)
    axis.grid(which="minor", axis="y", color="#EFEFEF", linewidth=0.6)
    axis.set_facecolor("white")
    figure.patch.set_facecolor("white")
    for spine in axis.spines.values():
        spine.set_color("#D9D9D9")
    axis.tick_params(axis="both", colors="#5A5A5A", labelsize=10)
    axis.legend(loc="upper center", bbox_to_anchor=(0.5, -0.16), ncol=3, frameon=False)
    figure.tight_layout()
    return figure


def _axis_bounds(values: list[float]) -> dict[str, float]:
    minimum = min(values)
    maximum = max(values)
    value_range = maximum - minimum

    major_unit = _y_major_unit(value_range)
    minor_unit = round(max(major_unit / 2, 0.05), 2)
    if math.isclose(minimum, maximum):
        padding = max(major_unit * 2, 0.5)
    else:
        padding = max(major_unit, value_range * 0.08)

    axis_min = math.floor(max(0, minimum - padding) / minor_unit) * minor_unit
    axis_max = math.ceil((maximum + padding) / minor_unit) * minor_unit

    return {
        "min": round(axis_min, 2),
        "max": round(axis_max, 2),
        "major_unit": round(major_unit, 2),
        "minor_unit": round(minor_unit, 2),
    }


def _y_major_unit(value_range: float) -> float:
    if value_range <= 1:
        return 0.2
    if value_range <= 2:
        return 0.25
    if value_range <= 4:
        return 0.5
    if value_range <= 8:
        return 1.0
    return 2.0


def _unique_sheet_name(base: str, used_names: set[str]) -> str:
    sanitized = base.replace("/", "-").replace("\\", "-")[:31] or "Athlet"
    if sanitized not in used_names:
        return sanitized

    index = 2
    while True:
        suffix = f"_{index}"
        candidate = f"{sanitized[:31 - len(suffix)]}{suffix}"
        if candidate not in used_names:
            return candidate
        index += 1


def _safe_filename(value: str) -> str:
    normalized = NON_FILENAME_RE.sub("_", value.strip()).strip("._")
    return normalized or "athlet"


def preview_rows(athlete: AthleteData) -> list[dict[str, str | int]]:
    rows: list[dict[str, str | int]] = []
    for block in athlete.blocks:
        for lap in block.laps:
            rows.append(
                {
                    "Block": block.index,
                    "Runde": lap.lap_number,
                    "Rundenzeit (s)": format_seconds(lap.lap_seconds),
                    "Aufstieg (s)": format_seconds(lap.ascent_seconds),
                    "Abfahrt (s)": format_seconds(lap.descent_seconds),
                }
            )
    return rows


def export_csv_to_excel(input_path: str | Path, output_path: str | Path) -> Path:
    input_path = Path(input_path)
    output_path = Path(output_path)
    dataset = parse_dataset(input_path.read_bytes())
    output_path.write_bytes(build_workbook(dataset))
    return output_path
