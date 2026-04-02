# Freelap Export Hub

Lokale und cloud-faehige Streamlit-App, die eine Freelap-CSV einliest, Rider IDs optional mit Athletennamen matched und daraus Excel- sowie PDF-Exporte erzeugt.

## Funktionen

- liest semikolon-getrennte Freelap-CSV-Dateien ein
- erkennt pro Athlet mehrere Bloecke
- erlaubt den Upload einer Athletenliste als CSV oder Excel
- matched Rider IDs mit Athletennamen fuer Vorschau, Excel-Sheets und Dateinamen
- erzeugt eine `.xlsx` mit Uebersicht und einem Sheet pro Athlet
- erzeugt PDF-Diagramme fuer Rundenzeiten, Aufstiege und Abfahrten
- bietet drei Downloads:
  - Excel-Datei
  - ein PDF pro Athlet
  - einzelne PDF-Charts pro Athlet als ZIP

## Start lokal

```bash
python3 -m venv .venv
.venv/bin/python -m pip install -r requirements.txt
.venv/bin/streamlit run app.py
```

## Format Athletenliste

Die Datei darf als `.csv`, `.xlsx` oder `.xls` hochgeladen werden.

Erwartete Spaltennamen:

- fuer die ID: `Rider ID`, `Athlete ID`, `ID`, `Bib`
- fuer den Namen: `Athlete Name`, `Name`, `Rider Name`, `Athlete`

Die Zuordnung erfolgt ueber die Rider ID. Wenn ein Match gefunden wird, erscheinen die Exporte direkt unter dem Athletennamen.

## Streamlit Cloud

Das Projekt ist fuer Streamlit Cloud vorbereitet:

- `app.py` ist der Einstiegspunkt
- `requirements.txt` enthaelt alle Python-Abhaengigkeiten
- `runtime.txt` legt Python 3.11 fest
- `.streamlit/config.toml` enthaelt die App-Konfiguration

Deployment:

1. Repository nach GitHub pushen.
2. In Streamlit Community Cloud das Repo verbinden.
3. Als Main file path `app.py` waehlen.
4. Deploy starten.

## Hinweis zur Blocklogik

Zeilen ohne `L1` und `L2`, auf die direkt wieder Split-Zeilen folgen, werden als Block-Marker interpretiert. Wenn eine Datei nur reine Rundenzeiten ohne Splits enthaelt, werden diese Zeilen als normale Runden behandelt.
