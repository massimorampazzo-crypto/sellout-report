import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Inches
import matplotlib.pyplot as plt
import smtplib
from email.message import EmailMessage
import tempfile

st.set_page_config(page_title="Sellout Report Generator")

st.title("📊 Sell-In / Sell-Out Report Generator")

uploaded_file = st.file_uploader(
    "Carica file Excel o CSV",
    type=["csv", "xlsx"]
)

email_dest = st.text_input("Inserisci email destinatario")

GMAIL_USER = "selloutreportsvgs@gmail.com"
GMAIL_PASSWORD = "Max2026$$$"


def generate_report(df):

    numeric_cols = [
        'Q.ta Carichi',
        'Val Vendite',
        'Q.ta Vendite'
    ]

    for col in numeric_cols:
        if col in df.columns:
            df[col] = (
                df[col]
                .astype(str)
                .str.replace('.', '', regex=False)
                .str.replace(',', '.', regex=False)
            )

            df[col] = pd.to_numeric(
                df[col],
                errors='coerce'
            )

    # Pulizia nomi colonne
df.columns = df.columns.str.strip()

# Trova automaticamente colonne corrette
vendite_col = None
carichi_col = None

for col in df.columns:

    if 'Vendite' in col and 'Q.ta' in col:
        vendite_col = col

    if 'Carichi' in col and 'Q.ta' in col:
        carichi_col = col

# Calcolo sell through
df['Sell Through %'] = (
    df[vendite_col] /
    df[carichi_col] * 100
).round(1)

top = df.sort_values(
    'Sell Through %',
    ascending=False
).head(10)

plt.figure(figsize=(8,4))

    plt.bar(
        top.iloc[:,0]
        top['Sell Through %']
    )

    plt.xticks(rotation=45)

    chart_path = tempfile.NamedTemporaryFile(
        delete=False,
        suffix=".png"
    ).name

    plt.tight_layout()
    plt.savefig(chart_path)
    plt.close()

doc = Document()

    doc.add_heading(
        'Report Sell-In / Sell-Out',
        level=1
    )

    doc.add_paragraph(
        f"Sell-through medio: "
        f"{df['Sell Through %'].mean():.1f}%"
    )

    table = doc.add_table(
        rows=1,
        cols=3
    )

    hdr = table.rows[0].cells

    hdr[0].text = "Brand"
    hdr[1].text = "Sell Through"
    hdr[2].text = "Vendite"

    for _, row in top.iterrows():

        cells = table.add_row().cells

        cells[0].text = str(
            row[df.columns[4]]
        )

        cells[1].text = (
            f"{row['Sell Through %']}%"
        )

        cells[2].text = (
            f"€ {row['Val Vendite']:,.0f}"
        )

    doc.add_picture(chart_path, width=Inches(5))

    report_path = tempfile.NamedTemporaryFile(
        delete=False,
        suffix=".docx"
    ).name

    doc.save(report_path)

    return report_path


if st.button("Genera Report"):

    if uploaded_file and email_dest:

        if uploaded_file.name.endswith(".csv"):

            df = pd.read_csv(
                uploaded_file,
                sep=";"
            )

        else:

            df = pd.read_excel(
                uploaded_file
            )

        report = generate_report(df)

        msg = EmailMessage()

        msg['Subject'] = (
            'Report Sell-Out'
        )

        msg['From'] = GMAIL_USER

        msg['To'] = email_dest

        msg.set_content(
            'In allegato il report.'
        )

        with open(report, 'rb') as f:

            msg.add_attachment(
                f.read(),
                maintype='application',
                subtype='octet-stream',
                filename='report.docx'
            )

        with smtplib.SMTP_SSL(
            'smtp.gmail.com',
            465
        ) as smtp:

            smtp.login(
                GMAIL_USER,
                GMAIL_PASSWORD
            )

            smtp.send_message(msg)

        st.success(
            "✅ Report inviato!"
        )
