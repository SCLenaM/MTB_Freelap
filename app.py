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


def render_styles() -> None:
    st.markdown(
        """
        <style>
        @import url('https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@500;700&family=Source+Sans+3:wght@400;600;700&display=swap');

        :root {
            --bg-soft: #edf5f2;
            --bg-card: rgba(255, 255, 255, 0.82);
            --border-soft: rgba(11, 110, 79, 0.12);
            --text-strong: #17313B;
            --text-soft: #49626B;
            --accent: #0B6E4F;
            --accent-warm: #D97706;
        }

        .stApp {
            background:
                radial-gradient(circle at top left, rgba(217, 119, 6, 0.10), transparent 28%),
                radial-gradient(circle at top right, rgba(11, 110, 79, 0.10), transparent 22%),
                linear-gradient(180deg, #f8fcfb 0%, #eef6f3 100%);
        }

        .stApp, [data-testid="stSidebar"] {
            font-family: "Source Sans 3", sans-serif;
        }

        h1, h2, h3 {
            font-family: "Space Grotesk", sans-serif !important;
            letter-spacing: -0.02em;
            color: var(--text-strong);
        }

        [data-testid="stSidebar"] {
            background: linear-gradient(180deg, #eaf4f0 0%, #ddeee7 100%);
            border-right: 1px solid rgba(11, 110, 79, 0.08);
        }

        [data-testid="stSidebar"] .stFileUploader {
            background: rgba(255, 255, 255, 0.65);
            border: 1px solid rgba(11, 110, 79, 0.10);
            border-radius: 20px;
            padding: 0.8rem;
        }

        [data-testid="stMetric"] {
            background: var(--bg-card);
            border: 1px solid var(--border-soft);
            border-radius: 20px;
            padding: 1rem;
            box-shadow: 0 10px 30px rgba(23, 49, 59, 0.06);
        }

        [data-testid="stDataFrame"], div[data-baseweb="select"] > div {
            border-radius: 18px !important;
        }

        .hero-card {
            padding: 2rem 2.2rem;
            border-radius: 28px;
            background:
                linear-gradient(135deg, rgba(11, 110, 79, 0.95) 0%, rgba(23, 49, 59, 0.96) 100%);
            color: white;
            box-shadow: 0 22px 60px rgba(23, 49, 59, 0.18);
            margin-bottom: 1.2rem;
        }

        .hero-kicker {
            display: inline-block;
            padding: 0.3rem 0.7rem;
            border-radius: 999px;
            background: rgba(255, 255, 255, 0.12);
            font-size: 0.88rem;
            letter-spacing: 0.04em;
            text-transform: uppercase;
        }

        .hero-title {
            margin: 0.9rem 0 0.7rem 0;
            font-family: "Space Grotesk", sans-serif;
            font-size: clamp(2.2rem, 5vw, 4rem);
            line-height: 0.95;
            letter-spacing: -0.04em;
        }

        .hero-copy {
            font-size: 1.15rem;
            line-height: 1.5;
            max-width: 56rem;
            color: rgba(255, 255, 255, 0.88);
            margin: 0;
        }

        .mini-grid {
            display: grid;
            grid-template-columns: repeat(3, minmax(0, 1fr));
            gap: 0.9rem;
            margin: 1rem 0 1.3rem 0;
        }

        .mini-card {
            background: var(--bg-card);
            border: 1px solid var(--border-soft);
            border-radius: 20px;
            padding: 1rem 1.1rem;
            box-shadow: 0 12px 30px rgba(23, 49, 59, 0.05);
        }

        .mini-card strong {
            display: block;
            color: var(--text-strong);
            margin-bottom: 0.2rem;
            font-family: "Space Grotesk", sans-serif;
        }

        .mini-card span {
            color: var(--text-soft);
        }

        .section-label {
            margin-top: 1.2rem;
            margin-bottom: 0.5rem;
            color: var(--text-strong);
            font-family: "Space Grotesk", sans-serif;
            font-size: 1.25rem;
            font-weight: 700;
        }

        .info-band {
            background: linear-gradient(135deg, rgba(217, 119, 6, 0.12), rgba(11, 110, 79, 0.10));
            border: 1px solid rgba(217, 119, 6, 0.16);
            color: var(--text-strong);
            border-radius: 18px;
            padding: 1rem 1.1rem;
            margin-bottom: 1rem;
        }

        .download-card {
            background: var(--bg-card);
            border: 1px solid var(--border-soft);
            border-radius: 22px;
            padding: 1rem;
            box-shadow: 0 12px 28px rgba(23, 49, 59, 0.05);
        }

        .download-card p {
            color: var(--text-soft);
            margin-top: 0.1rem;
            margin-bottom: 0.9rem;
        }

        .stDownloadButton > button {
            border-radius: 14px;
            font-weight: 700;
            min-height: 3rem;
        }

        @media (max-width: 900px) {
            .mini-grid {
                grid-template-columns: 1fr;
            }

            .hero-card {
                padding: 1.4rem;
            }
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def section_title(label: str) -> None:
    st.markdown(f'<div class="section-label">{label}</div>', unsafe_allow_html=True)


render_styles()

st.markdown(
    """
    <div class="hero-card">
        <span class="hero-kicker">Freelap Analyseplattform</span>
        <h1 class="hero-title">Freelap Export Hub</h1>
        <p class="hero-copy">
            Freelap-Export hochladen, Rider IDs mit Athletennamen verknuepfen und
            Excel- sowie PDF-Versionen direkt in einer klaren, schnellen und
            teamtauglichen Oberflaeche herunterladen.
        </p>
    </div>
    """,
    unsafe_allow_html=True,
)

st.markdown(
    """
    <div class="mini-grid">
        <div class="mini-card">
            <strong>1. Import</strong>
            <span>Freelap-CSV in der Seitenleiste hochladen.</span>
        </div>
        <div class="mini-card">
            <strong>2. Zuordnung</strong>
            <span>Optional eine Athletenliste hochladen, um Rider IDs durch Namen zu ersetzen.</span>
        </div>
        <div class="mini-card">
            <strong>3. Export</strong>
            <span>Excel-Uebersichten und PDF-Grafiken mit einem Klick herunterladen.</span>
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

with st.sidebar:
    st.markdown("### Datei-Upload")
    st.caption("Zuerst die Session-Daten und bei Bedarf danach die Athletenliste hochladen.")
    uploaded_file = st.file_uploader("Freelap CSV", type=["csv"])
    mapping_file = st.file_uploader(
        "Athletenliste",
        type=["csv", "xlsx", "xls"],
        help="Erwartete Spalten: zum Beispiel 'Rider ID' und 'Athlete Name'.",
    )

st.markdown(
    '<div class="info-band">Die Anwendung laeuft auf Streamlit Cloud. '
    "Dateien hochladen und die Exporte direkt im Browser herunterladen.</div>",
    unsafe_allow_html=True,
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

        stats_col1, stats_col2, stats_col3 = st.columns(3)
        stats_col1.metric("Athleten", len(dataset.athletes))
        stats_col2.metric("Session", dataset.session_label)
        stats_col3.metric("Datum", dataset.session_date or "-")

        if mapping_report is not None:
            section_title("Namens-Matching")
            match_col1, match_col2, match_col3 = st.columns(3)
            match_col1.metric("Zugeordnete IDs", mapping_report.matched)
            match_col2.metric("Ohne Namen", len(mapping_report.unmatched_ids))
            match_col3.metric("Nicht verwendet", len(mapping_report.unused_mapping_ids))

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

        preview_col, details_col = st.columns([1.35, 1], gap="large")

        with preview_col:
            section_title("Uebersicht")
            st.dataframe(summary_rows, use_container_width=True, hide_index=True)

        athlete_labels = {athlete.summary_label: athlete for athlete in dataset.athletes}

        with details_col:
            section_title("Ausgewaehlter Athlet")
            selected_athlete_label = st.selectbox(
                "Athlet auswaehlen",
                list(athlete_labels),
                label_visibility="collapsed",
            )
            selected_athlete = athlete_labels[selected_athlete_label]
            st.markdown(
                f"""
                <div class="download-card">
                    <strong style="font-family:'Space Grotesk',sans-serif;font-size:1.1rem;">
                        {selected_athlete.display_name}
                    </strong>
                    <p>Rider ID: {selected_athlete.athlete_id}</p>
                </div>
                """,
                unsafe_allow_html=True,
            )

        section_title("Rundenvorschau")
        st.dataframe(preview_rows(selected_athlete), use_container_width=True, hide_index=True)

        timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M")
        workbook_bytes = build_workbook(dataset)
        pdf_zip_per_athlete = build_pdf_zip(dataset, export_mode="per_athlete")
        pdf_zip_per_chart = build_pdf_zip(dataset, export_mode="per_chart")
        selected_athlete_pdf = build_athlete_pdf(dataset, selected_athlete)

        section_title("Downloads")
        row1_col1, row1_col2 = st.columns(2, gap="large")
        row2_col1, row2_col2 = st.columns(2, gap="large")

        with row1_col1:
            st.markdown(
                """
                <div class="download-card">
                    <strong style="font-family:'Space Grotesk',sans-serif;">Excel-Workbook</strong>
                    <p>Komplette Datei mit Uebersicht, Athletenblaettern und Diagrammen.</p>
                </div>
                """,
                unsafe_allow_html=True,
            )
            st.download_button(
                label="Excel-Datei herunterladen",
                data=workbook_bytes,
                file_name=f"freelap_export_{timestamp}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

        with row1_col2:
            st.markdown(
                f"""
                <div class="download-card">
                    <strong style="font-family:'Space Grotesk',sans-serif;">Einzelnes PDF</strong>
                    <p>Komplettes PDF von {selected_athlete.display_name} mit allen Diagrammen.</p>
                </div>
                """,
                unsafe_allow_html=True,
            )
            st.download_button(
                label=f"PDF fuer {selected_athlete.display_name} herunterladen",
                data=selected_athlete_pdf,
                file_name=f"{selected_athlete.display_name}_{timestamp}.pdf",
                mime="application/pdf",
                use_container_width=True,
            )

        with row2_col1:
            st.markdown(
                """
                <div class="download-card">
                    <strong style="font-family:'Space Grotesk',sans-serif;">ZIP pro Athlet</strong>
                    <p>Ein komplettes PDF pro Athlet in einer gemeinsamen ZIP-Datei.</p>
                </div>
                """,
                unsafe_allow_html=True,
            )
            st.download_button(
                label="ZIP pro Athlet herunterladen",
                data=pdf_zip_per_athlete,
                file_name=f"freelap_pdfs_pro_athlet_{timestamp}.zip",
                mime="application/zip",
                use_container_width=True,
            )

        with row2_col2:
            st.markdown(
                """
                <div class="download-card">
                    <strong style="font-family:'Space Grotesk',sans-serif;">ZIP pro Grafik</strong>
                    <p>Eine ZIP-Datei mit getrennten PDF-Grafiken fuer jeden Athleten.</p>
                </div>
                """,
                unsafe_allow_html=True,
            )
            st.download_button(
                label="ZIP pro Grafik herunterladen",
                data=pdf_zip_per_chart,
                file_name=f"freelap_pdfs_pro_grafik_{timestamp}.zip",
                mime="application/zip",
                use_container_width=True,
            )
else:
    st.markdown(
        """
        <div class="download-card">
            <strong style="font-family:'Space Grotesk',sans-serif;">Bereit zum Start</strong>
            <p>
                Freelap-CSV in der Seitenleiste hochladen, um Analyse,
                Athletenzuordnung und herunterladbare Exporte anzuzeigen.
            </p>
        </div>
        """,
        unsafe_allow_html=True,
    )
