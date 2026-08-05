import streamlit as st
import pandas as pd
from io import BytesIO
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet

st.set_page_config(page_title="LinkedIn Excel to PDF", layout="wide")
st.title("LinkedIn Excel to PDF")
st.write("Carica i 3 file Excel e genera un PDF riassuntivo.")

uploaded_files = st.file_uploader(
    "Carica i 3 file Excel",
    type=["xls", "xlsx"],
    accept_multiple_files=True
)


def read_excel_file(uploaded_file):
    data = uploaded_file.read()
    buffer = BytesIO(data)
    try:
        return pd.read_excel(buffer)
    except Exception:
        buffer.seek(0)
        return pd.read_excel(buffer, engine="xlrd")


def build_pdf(dfs, names):
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=36, leftMargin=36, topMargin=36, bottomMargin=36)
    styles = getSampleStyleSheet()
    story = []

    story.append(Paragraph("LinkedIn Report Summary", styles["Title"]))
    story.append(Spacer(1, 12))
    story.append(Paragraph("Report generato dai 3 file Excel caricati.", styles["Normal"]))
    story.append(Spacer(1, 12))

    summary_data = [["File", "Righe", "Colonne"]]
    for name, df in zip(names, dfs):
        summary_data.append([name, str(len(df)), str(len(df.columns))])

    summary_table = Table(summary_data, hAlign="LEFT")
    summary_table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1f4e79")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("BACKGROUND", (0, 1), (-1, -1), colors.whitesmoke),
    ]))
    story.append(summary_table)
    story.append(Spacer(1, 18))

    for name, df in zip(names, dfs):
        story.append(Paragraph(f"File: {name}", styles["Heading2"]))
        story.append(Paragraph(f"Righe: {len(df)} | Colonne: {len(df.columns)}", styles["Normal"]))
        story.append(Spacer(1, 8))
        preview = df.head(10).fillna("").astype(str)
        table_data = [list(preview.columns)] + preview.values.tolist()
        t = Table(table_data, repeatRows=1)
        t.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#d9eaf7")),
            ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ]))
        story.append(t)
        story.append(Spacer(1, 18))

    doc.build(story)
    buffer.seek(0)
    return buffer


if uploaded_files:
    st.success(f"File caricati: {len(uploaded_files)}")

    if len(uploaded_files) < 3:
        st.warning("Carica tutti e 3 i file Excel per generare il PDF.")
    elif len(uploaded_files) > 3:
        st.warning("Hai caricato più di 3 file: userò i primi 3.")

    if st.button("Genera PDF"):
        try:
            selected = uploaded_files[:3]
            dfs = [read_excel_file(f) for f in selected]
            names = [f.name for f in selected]
            pdf_buffer = build_pdf(dfs, names)
            st.download_button(
                label="Scarica PDF",
                data=pdf_buffer,
                file_name="linkedin_report.pdf",
                mime="application/pdf"
            )
            st.success("PDF generato correttamente.")
        except Exception as e:
            st.error(f"Errore nella generazione del PDF: {e}")
