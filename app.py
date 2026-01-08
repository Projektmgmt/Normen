# -*- coding: utf-8 -*-
"""
PM Normen & Standards Explorer – Light Theme + saubere PDF-Ausgabe
- Kein Upload – Excel wird direkt aus dem Dateisystem gelesen (Data as DB)
- Filter: Titel, Art, Jahr (robust), Trägerorganisation, Kategorie 1–3
- Ergebnis als helle Tabelle (HTML), PDF-Export
- Farbschema: hellblau & weiß, Light Mode für Oberfläche/Sidebar
"""

import streamlit as st
import pandas as pd
from io import BytesIO
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from datetime import datetime
from pathlib import Path
from typing import List

# -------------------------------------------------------
# Seite & Styling
# -------------------------------------------------------
st.set_page_config(page_title="Recherche von Normen und Standards für das Projekt-, Programm- und Portfoliomanagement", layout="wide", initial_sidebar_state="expanded")

# Light UI (Seitenleiste & Hintergrund)
st.markdown(
    """
    <style>
      html, body, [data-testid="stAppViewContainer"], [data-testid="stHeader"], [data-testid="stToolbar"] {
        background-color: #ffffff !important; color: #0f172a !important;
      }
      .main { background-color: #ffffff !important; }
      section[data-testid="stSidebar"] {
        background-color: #e6f2ff !important; color: #0f172a !important;
      }
      section[data-testid="stSidebar"] * { color: #0f172a !important; }
      section[data-testid="stSidebar"] input,
      section[data-testid="stSidebar"] textarea,
      section[data-testid="stSidebar"] [data-baseweb="select"] > div {
        background-color: #ffffff !important; color: #0f172a !important;
      }
      section[data-testid="stSidebar"] .stMultiSelect div[data-baseweb="tag"] { color:#0f172a !important; }
      section[data-testid="stSidebar"] .stMultiSelect div[aria-selected="true"] { color:#0f172a !important; }

      /* Download-Button dauerhaft hell */
      .stDownloadButton > button {
        border-radius: 10px;
        background-color: #e6f2ff !important;
        color: #000000 !important;
        border: 1px solid #99c2ff !important;
      }
      .stDownloadButton > button:hover {
        background-color: #cce0ff !important;
        color: #000000 !important;
        border: 1px solid #4d94ff !important;
      }

      /* Eigene helle Tabelle (für HTML-Render) */
      table.pm-table { border-collapse: collapse; width: 100%; }
      table.pm-table th, table.pm-table td {
        border: 1px solid #e5e7eb; padding: 8px; text-align: left; color: #0f172a;
      }
      table.pm-table th { background: #e6f2ff; font-weight: 600; }
      table.pm-table tr:nth-child(even) td { background: #f9f9f9; }
      table.pm-table tr:nth-child(odd) td  { background: #ffffff; }
    </style>
    """,
    unsafe_allow_html=True,
)

# -------------------------------------------------------
# Autoren (nur im PDF; nicht in der UI)
# -------------------------------------------------------
NAME1 = "Prof. Dr. Michael Klotz"
NAME2 = "Prof. Dr. Susanne Marx"
NAME3 = "Benjamin Birkmann"

# -------------------------------------------------------
# Daten laden (ohne Upload)
# -------------------------------------------------------
EXPECTED_COLS = [
    "Titel", "Art", "Herausgabejahr", "Trägerorganisation",
    "Kategorie 1", "Kategorie 2", "Kategorie 3",
]

DATA_CANDIDATES = [
    Path("data") / "Tabellarische_Darstellung.xlsx",
]


def load_excel() -> pd.DataFrame:
    for p in DATA_CANDIDATES:
        if p.exists():
            return pd.read_excel(p)
    raise FileNotFoundError("Excel nicht gefunden. EXCEL_PATH anpassen oder Datei ins Repo (data/) legen.")

df = load_excel()

missing = [c for c in EXPECTED_COLS if c not in df.columns]
if missing:
    st.error("In der Excel fehlen erwartete Spalten: " + ", ".join(missing))
    st.stop()

# Jahr robust numerisch
df["Herausgabejahr"] = pd.to_numeric(df["Herausgabejahr"], errors="coerce")
import re

# Mehrfachwerte in Kategorie-Spalten in einzelne Tokens zerlegen
KAT_COLS = ["Kategorie 1", "Kategorie 2", "Kategorie 3"]

def split_tokens(val: str):
    if pd.isna(val):
        return []
    # an Kommas trennen, trimmen, Leereinträge entfernen
    return [t.strip() for t in re.split(r",", str(val)) if t.strip()]

# Zusätzliche Listen-Spalten für robustes Filtern
for c in KAT_COLS:
    tok_col = f"{c}__tokens"
    df[tok_col] = df[c].apply(split_tokens)

# -------------------------------------------------------
# UI & Filter
# -------------------------------------------------------
st.title("📘 Recherche von Normen und Standards für das Projekt-, Programm- und Portfoliomanagement")
st.subheader("Prof. Dr. Michael Klotz, Prof. Dr. Susanne Marx, Benjamin Birkmann")
st.markdown("""
Die hier aufgeführten Normen und Standards des Projekt-, Programm- und Portfoliomanagements wurden im Rahmen eines Arbeitspapiers von Prof. Dr. Michael Klotz und Prof. Dr. Susanne Marx ermittelt. Hierfür kamen Methoden der Dokumentenanalyse, der systematischen Literaturanalyse und der qualitativen Inhaltsanalyse zum Einsatz.
Insgesamt werden 37 PM-Normen und 54 PM-Standards, die von 29 Trägerorganisationen publiziert werden, beschrieben. Jede Norm und jeder Standard werden im Arbeitspapier einzeln systematisch dargestellt.
Die jeweilige Beschreibung enthält eine Inhaltsangabe, den formellen Status der Norm bzw. des Standards und Links für die eigene weiterführende Recherche.
Insofern soll dieses Arbeitspapier nicht nur eine aktuelle, systematische Zusammenstellung bieten, sondern es stellt auch eine Hilfestellung für ein schnelles Orientieren und Nachschlagen dar.
Hierfür wurden die PM- Normen und -Standards verschiedenen Kategorien zugeordnet, die ihre inhaltliche Ausrichtung signalisieren.

Durch Experten empfohlende weitere Standards bzw. Aktualisierungen nehmen wir nach Prüfung in die Übersicht auf. Diese werden mit dem Zusatz "added / neu aufgenommen" gekennzeichnet.

Das Arbeitspapier steht frei zum Download zur Verfügung:
https://doi.org/10.13140/RG.2.2.18483.54562
Ebenso sind als Zusammenfassung Präsentationen auf Deutsch (DOI: 10.13140/RG.2.2.14744.87047) und Englisch (DOI: 10.13140/RG.2.2.21455.75683) verfügbar.""")

st.sidebar.header("🔍 Filter")

titel_filter = st.sidebar.text_input("Titel enthält:")
art_filter = st.sidebar.multiselect("Art", sorted(df["Art"].dropna().astype(str).unique()))
traeger_filter = st.sidebar.multiselect("Trägerorganisation", sorted(df["Trägerorganisation"].dropna().astype(str).unique()))
def options_from_tokens(tok_series):
    # Alle Tokens aus Listen-Spalte einsammeln und sortieren
    s = set()
    for lst in tok_series:
        s.update(lst)
    return sorted(s)

kat1_filter = st.sidebar.multiselect("Kategorie 1",
    options_from_tokens(df["Kategorie 1__tokens"]))
kat2_filter = st.sidebar.multiselect("Kategorie 2",
    options_from_tokens(df["Kategorie 2__tokens"]))
kat3_filter = st.sidebar.multiselect("Kategorie 3",
    options_from_tokens(df["Kategorie 3__tokens"]))

year_series = df["Herausgabejahr"].dropna()
jahr_range = None
if not year_series.empty:
    jmin, jmax = int(year_series.min()), int(year_series.max())
    jahr_range = st.sidebar.slider("Herausgabejahr", min_value=jmin, max_value=jmax, value=(jmin, jmax))

# -------------------------------------------------------
# Filtern
# -------------------------------------------------------
filtered_df = df.copy()
if titel_filter:
    filtered_df = filtered_df[filtered_df["Titel"].astype(str).str.contains(titel_filter, case=False, na=False)]
if art_filter:
    filtered_df = filtered_df[filtered_df["Art"].astype(str).isin(art_filter)]
if traeger_filter:
    filtered_df = filtered_df[filtered_df["Trägerorganisation"].astype(str).isin(traeger_filter)]
def has_any_token(token_list, selected):
    return any(t in token_list for t in selected)

if kat1_filter:
    filtered_df = filtered_df[
        filtered_df["Kategorie 1__tokens"].apply(lambda lst: has_any_token(lst, kat1_filter))
    ]
if kat2_filter:
    filtered_df = filtered_df[
        filtered_df["Kategorie 2__tokens"].apply(lambda lst: has_any_token(lst, kat2_filter))
    ]
if kat3_filter:
    filtered_df = filtered_df[
        filtered_df["Kategorie 3__tokens"].apply(lambda lst: has_any_token(lst, kat3_filter))
    ]
if jahr_range is not None:
    filtered_df = filtered_df[
        (filtered_df["Herausgabejahr"] >= jahr_range[0]) & (filtered_df["Herausgabejahr"] <= jahr_range[1])
    ]

# -------------------------------------------------------
# Ergebnisse – HELLE Tabelle (HTML)
# -------------------------------------------------------
st.subheader("Gefilterte Ergebnisse")
if filtered_df.empty:
    st.info("Keine Ergebnisse für die gewählten Filter.")
else:
    cols = [c for c in EXPECTED_COLS if c in filtered_df.columns]
    display_df = filtered_df[cols].copy()

    # Jahr ohne ".0" anzeigen
    if "Herausgabejahr" in display_df.columns:
        s = pd.to_numeric(display_df["Herausgabejahr"], errors="coerce")
        display_df["Herausgabejahr"] = s.apply(lambda x: "" if pd.isna(x) else f"{int(x)}")

    html_table = display_df.to_html(index=False, classes="pm-table", border=0, justify="left")
    st.markdown(html_table, unsafe_allow_html=True)

# -------------------------------------------------------
# PDF Export – A4 landscape + Wortumbruch + Spaltenbreiten
# -------------------------------------------------------
def create_pdf(dataframe: pd.DataFrame, namen: List[str]) -> bytes:
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib import colors
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

    buffer = BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=landscape(A4),   # Querformat
        leftMargin=16, rightMargin=16, topMargin=18, bottomMargin=18
    )

    styles = getSampleStyleSheet()
    elements = [
        Paragraph("Gefilterte Normen und Standards", styles["Heading1"]),
        Paragraph(f"Erstellt am {datetime.now().strftime('%d.%m.%Y, %H:%M')}", styles["Normal"]),
        Spacer(1, 12),
        Paragraph("""Die hier aufgeführten Normen und Standards des Projekt-, Programm- und Portfoliomanagements wurden im Rahmen eines Arbeitspapiers von Prof. Dr. Michael Klotz und Prof. Dr. Susanne Marx ermittelt. Hierfür kamen Methoden der Dokumentenanalyse, der systematischen Literaturanalyse und der qualitativen Inhaltsanalyse zum Einsatz.<br/>
        Insgesamt werden 37 PM-Normen und 54 PM-Standards, die von 29 Trägerorganisationen publiziert werden, beschrieben. Jede Norm und jeder Standard werden im Arbeitspapier einzeln systematisch dargestellt.<br/>
        Die jeweilige Beschreibung enthält eine Inhaltsangabe, den formellen Status der Norm bzw. des Standards und Links für die eigene weiterführende Recherche.<br/>
        Insofern soll dieses Arbeitspapier nicht nur eine aktuelle, systematische Zusammenstellung bieten, sondern es stellt auch eine Hilfestellung für ein schnelles Orientieren und Nachschlagen dar.<br/>
        Hierfür wurden die PM- Normen und -Standards verschiedenen Kategorien zugeordnet, die ihre inhaltliche Ausrichtung signalisieren.<br/>
        Durch Experten empfohlende weitere Standards bzw. Aktualisierungen nehmen wir nach Prüfung in die Übersicht auf. Diese werden mit dem Zusatz "added / neu aufgenommen" gekennzeichnet.<br/>
        Das Arbeitspapier steht frei zum Download zur Verfügung:<br/>
        https://doi.org/10.13140/RG.2.2.18483.54562<br/>
        Ebenso sind als Zusammenfassung Präsentationen auf Deutsch (DOI: 10.13140/RG.2.2.14744.87047) und Englisch (DOI: 10.13140/RG.2.2.21455.75683) verfügbar.""", styles["Normal"]),
        Spacer(1, 12),

    ]

    # --- Tabelle vorbereiten ---
    cols = [c for c in EXPECTED_COLS if c in dataframe.columns]
    tdf = dataframe[cols].copy()

    # Jahr als ganze Zahl darstellen (ohne .0)
    if "Herausgabejahr" in tdf.columns:
        tdf["Herausgabejahr"] = pd.to_numeric(tdf["Herausgabejahr"], errors="coerce")
        tdf["Herausgabejahr"] = tdf["Herausgabejahr"].apply(lambda x: "" if pd.isna(x) else f"{int(x)}")

    tdf = tdf.fillna("").astype(str)

    # Paragraph-Styles: Header + Zellen (mit Wortumbruch)
    header_style = ParagraphStyle(name="header", fontName="Helvetica-Bold", fontSize=10, leading=12)
    cell_style   = ParagraphStyle(name="cell",   fontName="Helvetica",      fontSize=9,  leading=11, wordWrap="CJK")

    header = [Paragraph(h, header_style) for h in tdf.columns]
    body   = [[Paragraph(val, cell_style) for val in row] for row in tdf.values.tolist()]
    table_data = [header] + body

    # Spaltengewichte — Titel/Träger breiter
    weights = []
    for c in tdf.columns:
        cl = c.lower()
        if cl.startswith("titel"):
            weights.append(3.2)
        elif cl.startswith("träger") or cl.startswith("traeger"):
            weights.append(3.2)
        elif cl.startswith("kategorie"):
            weights.append(1.6)
        elif cl.startswith("herausgabe"):
            weights.append(1.2)
        else:
            weights.append(1.4)

    page_width = doc.pagesize[0] - doc.leftMargin - doc.rightMargin
    total_w    = sum(weights)
    col_widths = [w / total_w * page_width for w in weights]

    # Tabelle mit festen Breiten + Umbruch
    t = Table(table_data, repeatRows=1, colWidths=col_widths, hAlign="LEFT")
    t.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,0), colors.lightblue),
        ("TEXTCOLOR",  (0,0), (-1,0), colors.black),
        ("FONTNAME",   (0,0), (-1,0), "Helvetica-Bold"),
        ("ALIGN",      (0,0), (-1,-1), "LEFT"),
        ("VALIGN",     (0,0), (-1,-1), "TOP"),
        ("FONTSIZE",   (0,1), (-1,-1), 9),
        ("LEADING",    (0,1), (-1,-1), 11),
        ("BOTTOMPADDING", (0,0), (-1,0), 8),
        ("BACKGROUND", (0,1), (-1,-1), colors.whitesmoke),
        ("GRID",       (0,0), (-1,-1), 0.25, colors.grey),
    ]))
    elements.append(t)

    elements.append(Spacer(1, 16))
    names_line = ", ".join([n for n in namen if n.strip()]) or "Name1, Name2, Name3"
    elements.append(Paragraph(f"Erstellt von: {names_line}", styles["Normal"]))

    doc.build(elements)
    pdf = buffer.getvalue()
    buffer.close()
    return pdf

if not filtered_df.empty:
    pdf_bytes = create_pdf(filtered_df, [NAME1, NAME2, NAME3])
    st.download_button(
        "📄 PDF herunterladen",
        data=pdf_bytes,
        file_name=f"Normen_Standards_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf",
        mime="application/pdf"
    )



















