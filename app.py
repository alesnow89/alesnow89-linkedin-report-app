import streamlit as st
import pandas as pd
from io import BytesIO
import matplotlib.pyplot as plt
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak, Image
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER

st.set_page_config(page_title="LinkedIn Excel to PDF", layout="wide")
st.title("LinkedIn Report Builder")
st.write("Carica i 3 file Excel esportati da LinkedIn e genera un report PDF più leggibile.")

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


def clean_df(df):
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    df = df.loc[:, ~df.columns.str.contains("^Unnamed", case=False, na=False)]
    return df


def num(v):
    try:
        if pd.isna(v):
            return 0
        return float(v)
    except Exception:
        return 0


def fmt(n):
    try:
        if float(n).is_integer():
            return f"{int(n):,}".replace(",", ".")
        return f"{n:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return str(n)


def date_col(df):
    for c in df.columns:
        cl = c.lower()
        if cl == "data" or cl.startswith("data") or "date" in cl:
            return c
    return None


def find_col(df, keyword):
    for c in df.columns:
        if keyword.lower() in c.lower():
            return c
    return None


def safe_cell(v):
    s = "" if pd.isna(v) else str(v)
    if len(s) > 40:
        s = s[:37] + "..."
    return Paragraph(s.replace("\n", "<br/>"), ParagraphStyle("cell", fontName="Helvetica", fontSize=6.5, leading=8))


def extract_metrics(content_df, followers_df, visitors_df):
    out = {}

    if not content_df.empty:
        c = content_df.copy()
        dcol = date_col(c)
        if dcol:
            c[dcol] = pd.to_datetime(c[dcol], errors="coerce")
            c = c.sort_values(dcol)
        out["content_rows"] = len(c)
        imp_col = find_col(c, "Impressioni (totale)")
        click_col = find_col(c, "Clic (totale)")
        react_col = find_col(c, "Reazioni (totale)")
        comm_col = find_col(c, "Commenti (totale)")
        interest_col = find_col(c, "Percentuale di interesse (totale)")
        out["content_impressions"] = num(c[imp_col].sum()) if imp_col else 0
        out["content_clicks"] = num(c[click_col].sum()) if click_col else 0
        out["content_reactions"] = num(c[react_col].sum()) if react_col else 0
        out["content_comments"] = num(c[comm_col].sum()) if comm_col else 0
        out["content_average_interest"] = num(c[interest_col].mean()) if interest_col else 0
        out["content_best_day"] = str(c.loc[c[imp_col].idxmax(), dcol].date()) if (imp_col and dcol and not c[imp_col].dropna().empty) else "N/A"
        out["content_daily"] = c[[dcol, imp_col, click_col, react_col]].copy() if dcol and imp_col else pd.DataFrame()
    else:
        out.update({"content_rows": 0, "content_impressions": 0, "content_clicks": 0, "content_reactions": 0, "content_comments": 0, "content_average_interest": 0, "content_best_day": "N/A", "content_daily": pd.DataFrame()})

    if not followers_df.empty:
        f = followers_df.copy()
        dcol = date_col(f)
        if dcol:
            f[dcol] = pd.to_datetime(f[dcol], errors="coerce")
            f = f.sort_values(dcol)
        out["followers_rows"] = len(f)
        total_col = find_col(f, "Follower totali")
        organic_col = find_col(f, "Follower organici")
        out["followers_total"] = num(f[total_col].sum()) if total_col else 0
        out["followers_organic"] = num(f[organic_col].sum()) if organic_col else 0
        out["followers_daily"] = f[[dcol, total_col, organic_col]].copy() if dcol and total_col else pd.DataFrame()
    else:
        out.update({"followers_rows": 0, "followers_total": 0, "followers_organic": 0, "followers_daily": pd.DataFrame()})

    if not visitors_df.empty:
        v = visitors_df.copy()
        dcol = date_col(v)
        if dcol:
            v[dcol] = pd.to_datetime(v[dcol], errors="coerce")
            v = v.sort_values(dcol)
        out["visitors_rows"] = len(v)
        pv_col = find_col(v, "Totale visualizzazioni della pagina (totale)")
        uv_col = find_col(v, "Totale visitatori unici (totale)")
        out["page_views_total"] = num(v[pv_col].sum()) if pv_col else 0
        out["unique_visitors_total"] = num(v[uv_col].sum()) if uv_col else 0
        out["visitors_daily"] = v[[dcol, pv_col, uv_col]].copy() if dcol and pv_col else pd.DataFrame()
    else:
        out.update({"visitors_rows": 0, "page_views_total": 0, "unique_visitors_total": 0, "visitors_daily": pd.DataFrame()})

    return out


def make_chart(df, date_col_name, value_cols, title):
    fig, ax = plt.subplots(figsize=(8, 3))
    if df is not None and not df.empty and date_col_name in df.columns:
        x = pd.to_datetime(df[date_col_name], errors="coerce")
        for col in value_cols:
            if col and col in df.columns:
                ax.plot(x, pd.to_numeric(df[col], errors="coerce"), marker="o", linewidth=2, label=col)
        ax.set_title(title)
        ax.tick_params(axis='x', rotation=45)
        ax.legend(fontsize=8)
        ax.grid(True, alpha=0.25)
        fig.tight_layout()
    else:
        ax.text(0.5, 0.5, "Dati non disponibili", ha="center", va="center")
        ax.axis("off")
    return fig


def fig_to_buffer(fig):
    buf = BytesIO()
    fig.savefig(buf, format="png", dpi=160, bbox_inches="tight")
    plt.close(fig)
    buf.seek(0)
    return buf


def build_pdf(content_df, followers_df, visitors_df, metrics, names, charts):
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=28, leftMargin=28, topMargin=28, bottomMargin=28)
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle("CenterTitle", parent=styles["Title"], alignment=TA_CENTER, textColor=colors.HexColor("#1f4e79"))
    small = ParagraphStyle("Small", parent=styles["BodyText"], fontSize=9, leading=11)
    story = []

    story.append(Paragraph("LinkedIn Report", title_style))
    story.append(Spacer(1, 8))
    story.append(Paragraph("Executive summary dei tre file esportati da LinkedIn.", small))
    story.append(Spacer(1, 12))

    summary = [
        ["File", "Righe", "Colonne"],
        [names[0], str(metrics["content_rows"]), str(len(content_df.columns))],
        [names[1], str(metrics["followers_rows"]), str(len(followers_df.columns))],
        [names[2], str(metrics["visitors_rows"]), str(len(visitors_df.columns))],
    ]
    t = Table(summary, hAlign="LEFT", colWidths=[240, 70, 70])
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1f4e79")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("GRID", (0, 0), (-1, -1), 0.4, colors.grey),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("BACKGROUND", (0, 1), (-1, -1), colors.whitesmoke),
    ]))
    story.append(t)
    story.append(Spacer(1, 14))

    story.append(Paragraph("Executive summary", styles["Heading1"]))
    summary_text = [
        f"Nel periodo analizzato, i contenuti hanno generato {fmt(metrics['content_impressions'])} impressioni, {fmt(metrics['content_clicks'])} clic, {fmt(metrics['content_reactions'])} reazioni e {fmt(metrics['content_comments'])} commenti.",
        f"La crescita follower registrata è stata di {fmt(metrics['followers_total'])} unità complessive, con prevalenza di follower organici ({fmt(metrics['followers_organic'])}).",
        f"La pagina ha totalizzato {fmt(metrics['page_views_total'])} visualizzazioni e {fmt(metrics['unique_visitors_total'])} visitatori unici.",
        f"Il giorno con il maggiore volume di impressioni è stato {metrics['content_best_day']}."
    ]
    for s in summary_text:
        story.append(Paragraph(s, small))
        story.append(Spacer(1, 4))

    story.append(Spacer(1, 6))
    story.append(Paragraph("Key insights", styles["Heading1"]))
    insights = [
        f"Impressioni totali contenuti: {fmt(metrics['content_impressions'])}.",
        f"Clic totali contenuti: {fmt(metrics['content_clicks'])}.",
        f"Follower totali registrati: {fmt(metrics['followers_total'])}.",
        f"Visualizzazioni pagina totali: {fmt(metrics['page_views_total'])}.",
        f"Visitatori unici totali: {fmt(metrics['unique_visitors_total'])}.",
        f"Giorno con più impressioni: {metrics['content_best_day']}.",
    ]
    for line in insights:
        story.append(Paragraph(f"• {line}", small))
    story.append(Spacer(1, 10))

    story.append(Paragraph("Charts", styles["Heading1"]))
    for chart in charts:
        story.append(Paragraph(chart[0], styles["Heading2"]))
        story.append(Image(chart[1], width=500, height=185))
        story.append(Spacer(1, 8))

    story.append(PageBreak())
    story.append(Paragraph("Interpretation", styles["Heading1"]))
    story.append(Paragraph(
        f"I dati mostrano una base follower in crescita, con {fmt(metrics['followers_organic'])} follower organici complessivi. "
        f"Il traffico sulla pagina è stato pari a {fmt(metrics['page_views_total'])} visualizzazioni e {fmt(metrics['unique_visitors_total'])} visitatori unici, "
        f"segno di interesse sul profilo aziendale.",
        small
    ))
    story.append(Spacer(1, 6))
    story.append(Paragraph(
        f"Il contenuto ha prodotto {fmt(metrics['content_impressions'])} impressioni e {fmt(metrics['content_clicks'])} clic. "
        f"Il rapporto tra visibilità e interazione suggerisce che alcuni post hanno performato meglio di altri, con un picco il {metrics['content_best_day']}.",
        small
    ))
    story.append(Spacer(1, 12))

    def preview_section(title, df):
        story.append(PageBreak())
        story.append(Paragraph(title, styles["Heading1"]))
        preview = df.head(6).fillna("").astype(str)
        display_cols = list(preview.columns[:6])
        preview = preview[display_cols]
        header = [Paragraph(str(c), ParagraphStyle("hdr", fontSize=6.5, leading=7, textColor=colors.black)) for c in preview.columns]
        table_data = [header]
        for _, row in preview.iterrows():
            table_data.append([safe_cell(v) for v in row.tolist()])
        tt = Table(table_data, repeatRows=1, colWidths=[72] * len(display_cols))
        tt.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#d9eaf7")),
            ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("FONTSIZE", (0, 0), (-1, -1), 6),
            ("LEADING", (0, 0), (-1, -1), 7),
        ]))
        story.append(tt)

    preview_section(f"Preview: {names[0]}", content_df)
    preview_section(f"Preview: {names[1]}", followers_df)
    preview_section(f"Preview: {names[2]}", visitors_df)

    doc.build(story)
    buffer.seek(0)
    return buffer


if uploaded_files:
    st.success(f"File caricati: {len(uploaded_files)}")
    if len(uploaded_files) < 3:
        st.warning("Carica tutti e 3 i file Excel.")
    elif len(uploaded_files) > 3:
        st.warning("Hai caricato più di 3 file: userò i primi 3.")

    if st.button("Genera PDF"):
        try:
            selected = uploaded_files[:3]
            dfs = [clean_df(read_excel_file(f)) for f in selected]
            names = [f.name for f in selected]
            metrics = extract_metrics(dfs[0], dfs[1], dfs[2])

            c_d = metrics["content_daily"]
            f_d = metrics["followers_daily"]
            v_d = metrics["visitors_daily"]

            charts = []
            if not c_d.empty:
                charts.append(("Content: impressioni, clic e reazioni", fig_to_buffer(make_chart(c_d, c_d.columns[0], c_d.columns[1:4], "Content performance"))))
            if not f_d.empty:
                charts.append(("Followers: andamento giornaliero", fig_to_buffer(make_chart(f_d, f_d.columns[0], f_d.columns[1:3], "Followers"))))
            if not v_d.empty:
                charts.append(("Visitors: page views e visitatori unici", fig_to_buffer(make_chart(v_d, v_d.columns[0], v_d.columns[1:3], "Visitors"))))

            pdf_buffer = build_pdf(dfs[0], dfs[1], dfs[2], metrics, names, charts)
            st.download_button(
                label="Scarica PDF",
                data=pdf_buffer,
                file_name="linkedin_report.pdf",
                mime="application/pdf"
            )
            st.success("PDF generato correttamente.")
        except Exception as e:
            st.error(f"Errore nella generazione del PDF: {e}")
