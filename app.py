import streamlit as st
import pandas as pd
import io
import datetime
import os
from fpdf import FPDF
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment

st.set_page_config(layout="wide")

# SIDEBAR
secili_sablon = st.sidebar.radio(
    "Şablon Seç",
    ["⚓ INNOMAR Özel Teklif", "📄 Proforma Fatura"]
)

# INITIAL DATA
if 'df' not in st.session_state:
    if "Teklif" in secili_sablon:
        st.session_state.df = pd.DataFrame({
            'INSPECTION': ['Bakım', 'Onarım'],
            'UNIT': ['2', '1'],
            'PRICE': [1000, 2000]
        })
    else:
        st.session_state.df = pd.DataFrame({
            'Açıklama': ['',''],
            'Birim Fiyat': [0.0,0.0],
            'Adet': [1,1],
            'Tutar': [0.0,0.0]
        })

df = st.session_state.df

st.title(secili_sablon)

# TABLO
df = st.data_editor(df, num_rows="dynamic")

# OTOMATİK TUTAR
if "Birim Fiyat" in df.columns:
    df["Tutar"] = (
        pd.to_numeric(df["Birim Fiyat"], errors='coerce').fillna(0) *
        pd.to_numeric(df["Adet"], errors='coerce').fillna(0)
    )

# HESAPLAMA
toplam = df.iloc[:,-1].sum()
kdv = toplam * 0.20
genel = toplam + kdv

st.metric("Ara Toplam", f"{toplam:.0f} ₺")
st.metric("KDV", f"{kdv:.0f} ₺")
st.metric("Genel", f"{genel:.0f} ₺")

tarih = datetime.date.today().strftime("%d.%m.%Y")

# WORD
def word_olustur(df):
    doc = Document()

    if "Fatura" in secili_sablon:
        doc.add_paragraph("PROFORMA FATURA").runs[0].bold = True
        doc.add_paragraph("info@innomarin.com")
        doc.add_paragraph("www.innomarin.com")
        doc.add_paragraph("Heybeliada / Istanbul")

        doc.add_paragraph("\nMüşteri Adı:")
        doc.add_paragraph("Adres:")
        doc.add_paragraph(f"Tarih: {tarih}")

    table = doc.add_table(rows=1, cols=len(df.columns))

    for i, col in enumerate(df.columns):
        table.rows[0].cells[i].text = col

    for _, row in df.iterrows():
        cells = table.add_row().cells
        for i, val in enumerate(row):
            cells[i].text = str(val)

    doc.add_paragraph(f"\nAra Toplam: {toplam:.0f}")
    doc.add_paragraph(f"KDV: {kdv:.0f}")
    doc.add_paragraph(f"Toplam: {genel:.0f}")

    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()

# PDF
def pdf_olustur(df):
    pdf = FPDF()
    pdf.add_page()

    if "Fatura" in secili_sablon:
        pdf.set_font("Arial","B",14)
        pdf.cell(0,10,"PROFORMA FATURA",0,1,"C")

        pdf.set_font("Arial","",10)
        pdf.cell(0,5,"info@innomarin.com",0,1)
        pdf.cell(0,5,"www.innomarin.com",0,1)
        pdf.cell(0,5,"Heybeliada / Istanbul",0,1)

        pdf.ln(5)
        pdf.cell(0,5,"Musteri Adi:",0,1)
        pdf.cell(0,5,"Adres:",0,1)
        pdf.cell(0,5,f"Tarih: {tarih}",0,1)

    pdf.ln(5)

    for col in df.columns:
        pdf.cell(40,10,col,1)
    pdf.ln()

    for _, row in df.iterrows():
        for val in row:
            pdf.cell(40,10,str(val),1)
        pdf.ln()

    pdf.cell(0,10,f"Toplam: {genel:.0f}",0,1)

    return pdf.output(dest="S").encode("latin-1")

# EXCEL
def excel_olustur(df):
    wb = Workbook()
    ws = wb.active

    ws["A1"] = "PROFORMA FATURA"
    ws["A2"] = "info@innomarin.com"
    ws["A3"] = "www.innomarin.com"

    row = 5

    for i, col in enumerate(df.columns,1):
        ws.cell(row=row, column=i).value = col

    row +=1

    for _, r in df.iterrows():
        for i, val in enumerate(r,1):
            ws.cell(row=row, column=i).value = val
        row +=1

    ws[f"A{row+1}"] = f"Toplam: {genel:.0f}"

    bio = io.BytesIO()
    wb.save(bio)
    return bio.getvalue()

# DOWNLOAD
st.download_button("WORD", word_olustur(df), "fatura.docx")
st.download_button("PDF", pdf_olustur(df), "fatura.pdf")
st.download_button("EXCEL", excel_olustur(df), "fatura.xlsx")
