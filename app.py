from __future__ import annotations

from datetime import datetime

import streamlit as st

from freelap_report import (
    apply_athlete_names,
    build_athlete_pdf,
    build_mapping_report,
    build_pdf_zip,
    build_workbook,
    parse_athlete_mapping,
    parse_dataset,
    preview_rows,
)


st.set_page_config(page_title="Freelap Export Hub", layout="wide")

st.title("Freelap Export Hub")
st.write(
    "Freelap-CSV hochladen, Rider IDs mit Athletennamen matchen und direkt Excel- sowie PDF-Exporte herunterladen."
)

with st.sidebar:
    st.header("Uploads")
    uploaded_file = st.file_uploader("Freelap CSV", type=["csv"])
    mapping_file = st.file_uploader(
        "Athletenliste",
        type=["csv", "xlsx", "xls"],
        help="Erwartete Spalten: z. B. 'Rider ID' und 'Athlete Name'.",
    )

st.info(
    "Die App laeuft jetzt in Streamlit Cloud. Dateien hochladen und die Exporte direkt im Browser herunterladen."
)

if uploaded_file is not None:
    try:
        dataset = parse_dataset(uploaded_file.getvalue())
    except Exception as exc:
        st.error(f"Die Datei konnte nicht verarbeitet werden: {exc}")
    else:
        mapping_report = None
        if mapping_file is not None:
            try:
                athlete_mapping = parse_athlete_mapping(
                    mapping_file.getvalue(),
                    mapping_file.name,
                )
                dataset = apply_athlete_names(dataset, athlete_mapping)
                mapping_report = build_mapping_report(dataset, athlete_mapping)
            except Exception as exc:
                st.error(f"Die Athletenliste konnte nicht verarbeitet werden: {exc}")

        st.success(
            f"{len(dataset.athletes)} Athleten erkannt, "
            f"{sum(athlete.lap_count for athlete in dataset.athletes)} Runden verarbeitet."
        )

        col1, col2, col3 = st.columns(3)
        col1.metric("Athleten", len(dataset.athletes))
        col2.metric("Session", dataset.session_label)
        col3.metric("Datum", dataset.session_date or "-")

        if mapping_report is not None:
            st.subheader("Namens-Matching")
            match_col1, match_col2, match_col3 = st.columns(3)
            match_col1.metric("Gematchte Rider IDs", mapping_report.matched)
            match_col2.metric("Nicht gematcht", len(mapping_report.unmatched_ids))
            match_col3.metric("Nur in Athletenliste", len(mapping_report.unused_mapping_ids))

            if mapping_report.unmatched_ids:
                st.warning(
                    "Keine Namen gefunden fuer folgende Rider IDs: "
                    + ", ".join(mapping_report.unmatched_ids)
                )

        summary_rows = [
            {
                "Athlet": athlete.display_name,
                "Rider ID": athlete.athlete_id,
                "Bloecke": len(athlete.blocks),
                "Runden": athlete.lap_count,
                "Split-Bloecke": athlete.split_blocks,
            }
            for athlete in dataset.athletes
        ]
        st.subheader("Uebersicht")
        st.dataframe(summary_rows, use_container_width=True, hide_index=True)

        athlete_labels = {athlete.summary_label: athlete for athlete in dataset.athletes}
        selected_athlete_label = st.selectbox(
            "Vorschau fuer Athlet",
            list(athlete_labels),
        )
        selected_athlete = athlete_labels[selected_athlete_label]
        st.dataframe(preview_rows(selected_athlete), use_container_width=True, hide_index=True)

        timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M")
        workbook_bytes = build_workbook(dataset)
        pdf_zip_per_athlete = build_pdf_zip(dataset, export_mode="per_athlete")
        pdf_zip_per_chart = build_pdf_zip(dataset, export_mode="per_chart")
        selected_athlete_pdf = build_athlete_pdf(dataset, selected_athlete)

        st.subheader("Downloads")
        download_col1, download_col2 = st.columns(2)
        download_col1.download_button(
            label="Excel herunterladen",
            data=workbook_bytes,
            file_name=f"freelap_export_{timestamp}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
        download_col2.download_button(
            label=f"PDF fuer {selected_athlete.display_name}",
            data=selected_athlete_pdf,
            file_name=f"{selected_athlete.display_name}_{timestamp}.pdf",
            mime="application/pdf",
            use_container_width=True,
        )

        zip_col1, zip_col2 = st.columns(2)
        zip_col1.download_button(
            label="ZIP: ein PDF pro Athlet",
            data=pdf_zip_per_athlete,
            file_name=f"freelap_pdfs_pro_athlet_{timestamp}.zip",
            mime="application/zip",
            use_container_width=True,
        )
        zip_col2.download_button(
            label="ZIP: einzelne PDF-Charts pro Athlet",
            data=pdf_zip_per_chart,
            file_name=f"freelap_pdfs_pro_grafik_{timestamp}.zip",
            mime="application/zip",
            use_container_width=True,
        )
else:
    st.info("Dateien hochladen, dann werden Vorschau, Namens-Matching und Exporte erzeugt.")
