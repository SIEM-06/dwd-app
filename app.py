import streamlit as st
import pandas as pd
import io
import datetime
import os

from fpdf import FPDF
from docx import Document
from docx.shared import Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH

from openpyxl import Workbook
from openpyxl.drawing.image import Image as xlImage
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill

# ==============================
# AYARLAR
# ==============================

ANTET_IMG = "antet.png"

st.set_page_config(layout="wide", page_title="Doküman Oluşturucu")

# ==============================
# BAŞLANGIÇ DATA
# ==============================

if "df" not in st.session_state:
    st.session_state.df = pd.DataFrame({
        "Açıklama": ["Örnek Hizmet"],
        "Adet": [1],
        "Birim Fiyat": [1000]
    })

# ==============================
# ARAYÜZ
# ==============================

st.title("📄 Teklif / Fatura Oluşturucu")

df = st.data_editor(st.session_state.df, num_rows="dynamic", use_container_width=True)

# ==============================
# HESAPLAMA
# ==============================

fiyatlar = pd.to_numeric(df["Birim Fiyat"], errors="coerce").fillna(0)
adet = pd.to_numeric(df["Adet"], errors="coerce").fillna(0)

tutar = fiyatlar * adet

ara_toplam = tutar.sum()
kdv = ara_toplam * 0.20
genel = ara_toplam + kdv

st.metric("Ara Toplam", f"{ara_toplam:,.0f}")
st.metric("KDV %20", f"{kdv:,.0f}")
st.metric("Genel Toplam", f"{genel:,.0f}")

# ==============================
# WORD OLUŞTUR
# ==============================

def word_olustur(df):

    doc = Document()

    # header'a antet koy
    section = doc.sections[0]
    header = section.header
    p = header.paragraphs[0]

    if os.path.exists(ANTET_IMG):
        p.add_run().add_picture(ANTET_IMG, width=Cm(21))

    doc.add_paragraph("\n")

    table = doc.add_table(rows=1, cols=len(df.columns))

    for i, col in enumerate(df.columns):
        table.rows[0].cells[i].text = col

    for _, row in df.iterrows():
        cells = table.add_row().cells
        for i, val in enumerate(row):
            cells[i].text = str(val)

    doc.add_paragraph()

    doc.add_paragraph(f"Ara Toplam: {ara_toplam:,.0f}")
    doc.add_paragraph(f"KDV %20: {kdv:,.0f}")
    doc.add_paragraph(f"Genel Toplam: {genel:,.0f}")

    bio = io.BytesIO()
    doc.save(bio)

    return bio.getvalue()

# ==============================
# PDF OLUŞTUR
# ==============================

def pdf_olustur(df):

    class PDF(FPDF):

        def header(self):

            if os.path.exists(ANTET_IMG):
                self.image(ANTET_IMG, x=0, y=0, w=210)

            self.set_y(40)

    pdf = PDF()
    pdf.add_page()

    pdf.set_font("Arial", "B", 12)

    for col in df.columns:
        pdf.cell(40, 10, col, 1)

    pdf.ln()

    pdf.set_font("Arial", "", 11)

    for _, row in df.iterrows():
        for val in row:
            pdf.cell(40, 10, str(val), 1)
        pdf.ln()

    pdf.ln(10)

    pdf.cell(0,10,f"Ara Toplam: {ara_toplam:,.0f}",0,1)
    pdf.cell(0,10,f"KDV %20: {kdv:,.0f}",0,1)
    pdf.cell(0,10,f"Genel Toplam: {genel:,.0f}",0,1)

    return pdf.output(dest="S").encode("latin-1")

# ==============================
# EXCEL OLUŞTUR
# ==============================

def excel_olustur(df):

    wb = Workbook()
    ws = wb.active

    # antet resmi
    if os.path.exists(ANTET_IMG):

        img = xlImage(ANTET_IMG)
        ws.add_image(img, "A1")

    row_start = 12

    # tablo
    for i, col in enumerate(df.columns,1):
        ws.cell(row=row_start, column=i).value = col
        ws.cell(row=row_start, column=i).font = Font(bold=True)

    row_start += 1

    for r,row in df.iterrows():
        for c,val in enumerate(row,1):
            ws.cell(row=row_start, column=c).value = val
        row_start += 1

    row_start += 2

    ws.cell(row=row_start,column=1).value=f"Ara Toplam"
    ws.cell(row=row_start,column=2).value=ara_toplam

    ws.cell(row=row_start+1,column=1).value=f"KDV %20"
    ws.cell(row=row_start+1,column=2).value=kdv

    ws.cell(row=row_start+2,column=1).value=f"Genel Toplam"
    ws.cell(row=row_start+2,column=2).value=genel

    bio = io.BytesIO()
    wb.save(bio)

    return bio.getvalue()

# ==============================
# İNDİRME BUTONLARI
# ==============================

st.write("---")

c1,c2,c3 = st.columns(3)

with c1:
    st.download_button(
        "📄 WORD",
        data=word_olustur(df),
        file_name="belge.docx"
    )

with c2:
    st.download_button(
        "📊 EXCEL",
        data=excel_olustur(df),
        file_name="belge.xlsx"
    )

with c3:
    st.download_button(
        "📕 PDF",
        data=pdf_olustur(df),
        file_name="belge.pdf"
    )
