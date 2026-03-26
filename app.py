import streamlit as st
import pandas as pd
import io
import datetime
import os
from fpdf import FPDF
from docx import Document
from docx.shared import Cm, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter

st.set_page_config(layout="wide", page_title="Innomarin Kurumsal Platform")

# ==========================================
# YARDIMCI FONKSİYONLAR
# ==========================================
def cevir_tr(metin):
    tr_map = {'ş':'s', 'Ş':'S', 'ı':'i', 'İ':'I', 'ğ':'g', 'Ğ':'G', 'ü':'u', 'Ü':'U', 'ö':'o', 'Ö':'O', 'ç':'c', 'Ç':'C'}
    for k, v in tr_map.items(): metin = metin.replace(k, v)
    return metin

def get_birim_col(df_columns):
    for col in df_columns:
        if "birim fiyat" in str(col).lower(): return col
    return None

# ==========================================
# PLATFORM AYARLARI
# ==========================================
st.sidebar.markdown("### ⚙️ Sistem Ayarları")
secili_sablon = st.sidebar.radio(
    "📝 Çalışma Şablonunu Seçin:",
    ["⚓ INNOMAR Özel Teklif", "📄 Standart Proforma Fatura"]
)

gizle_checkbox = st.sidebar.checkbox("🔒 Birim Fiyatını Çıktılarda Gizle (Sansürle)", value=False)

if 'aktif_sablon' not in st.session_state or st.session_state.aktif_sablon != secili_sablon:
    st.session_state.aktif_sablon = secili_sablon
    if secili_sablon == "⚓ INNOMAR Özel Teklif":
        data = {'INSPECTION REMARK': ['ANA MAKİNE BAKIMLARI'], 'UNIT': ['1 SET'], 'PRICE': [40000.0]}
    else: 
        data = {'Marka': ['Innomarin'], 'Açıklama': ['Bakım Hizmeti'], 'KDV': ['%20'], 'Adet': ['1'], 'Birim Fiyat': ['5000'], 'Toplam Fiyat': [5000.0]}
    st.session_state.veri_df = pd.DataFrame(data)
    st.session_state.not_alani = "Banka Hesap Bilgilerimiz:\nIBAN: TR00..."
    st.rerun()

st.markdown(f"<h2 style='text-align: center;'>{secili_sablon.upper()}</h2>", unsafe_allow_html=True)

# --- PANEL ---
col_t, col_kur = st.columns(2)
secilen_tarih = col_t.date_input("Belge Tarihi", datetime.date.today())
kur_secimi = col_kur.selectbox("Para Birimi", ["Euro (€)", "Dolar ($)", "Türk Lirası (₺)"])
sembol = "€" if "Euro" in kur_secimi else "$" if "Dolar" in kur_secimi else "₺"
kur_metin = "EURO" if "Euro" in kur_secimi else "USD" if "Dolar" in kur_secimi else "TL"

# --- TABLO ---
df = st.session_state.veri_df.copy()
son_sutun = df.columns[-1]
for col in df.columns:
    if col == son_sutun: df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0.0)
    else: df[col] = df[col].astype(str).replace(['nan', 'NaN'], '')

duzenlenmis_df = st.data_editor(df, use_container_width=True, num_rows="dynamic")

# --- HESAPLAR ---
f_toplam = pd.to_numeric(duzenlenmis_df[son_sutun], errors='coerce').sum()
kdv = f_toplam * 0.20
g_toplam = f_toplam + kdv

# ==========================================
# WORD ÇIKTI (ARKAPLANLI)
# ==========================================
def word_olustur(dataframe, a_str, k_str, g_str, tarih, notlar, kur_m, sablon, gizle):
    doc = Document()
    birim_sutun = get_birim_col(dataframe.columns)
    
    # Arkaplan Ekleme
    if os.path.exists("arkaplan.png"):
        section = doc.sections[0]
        header = section.header
        p = header.paragraphs[0]
        run = p.add_run()
        run.add_picture("arkaplan.png", width=section.page_width, height=section.page_height)
        
    doc.add_paragraph("\n" * 6) # Logo boşluğu
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.add_run(f"{sablon.upper()}\nTarih: {tarih}").bold = True
    
    table = doc.add_table(rows=1, cols=len(dataframe.columns)+1)
    table.style = 'Table Grid'
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    
    # Başlıklar
    headers = ['Sıra'] + list(dataframe.columns)
    for i, h in enumerate(headers): table.rows[0].cells[i].text = h
    
    for i, row in dataframe.iterrows():
        cells = table.add_row().cells
        cells[0].text = str(i+1)
        for c_idx, cname in enumerate(dataframe.columns, 1):
            val = row[cname]
            if gizle and cname == birim_sutun: val = "***"
            elif cname == dataframe.columns[-1]:
                try: val = f"{float(val):,.0f} {kur_m}"
                except: val = "-NIL-"
            cells[c_idx].text = str(val)
            
    doc.add_paragraph(f"\nGenel Toplam: {g_str}\n\nNotlar:\n{notlar}")
    
    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()

# ==========================================
# EXCEL ÇIKTI (ARKAPLANLI)
# ==========================================
def excel_olustur(dataframe, g_str, tarih, sablon, gizle):
    wb = Workbook()
    ws = wb.active
    ws.sheet_view.showGridLines = False
    
    if os.path.exists("arkaplan.png"):
        ws.add_background("arkaplan.png")
    
    birim_sutun = get_birim_col(dataframe.columns)
    row_idx = 10
    ws.cell(row=row_idx, column=2).value = f"{sablon} - {tarih}"
    row_idx += 2
    
    # Başlıklar
    headers = ['Sıra'] + list(dataframe.columns)
    for i, h in enumerate(headers, 2):
        cell = ws.cell(row=row_idx, column=i, value=h)
        cell.font = Font(bold=True)
        cell.fill = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
        ws.column_dimensions[get_column_letter(i)].width = 18
    
    row_idx += 1
    for i, row in dataframe.iterrows():
        ws.cell(row=row_idx, column=2).value = i + 1
        for c_idx, cname in enumerate(dataframe.columns, 3):
            val = row[cname]
            if gizle and cname == birim_sutun: val = "***"
            ws.cell(row=row_idx, column=c_idx).value = str(val)
        row_idx += 1
    
    ws.cell(row=row_idx+1, column=len(headers)+1).value = f"Toplam: {g_str}"
    
    bio = io.BytesIO()
    wb.save(bio)
    return bio.getvalue()

# ==========================================
# PDF ÇIKTI (MÜKEMMEL HİZALANMIŞ)
# ==========================================
def pdf_olustur(dataframe, a_str, k_str, g_str, tarih, notlar, kur_m, sablon, gizle):
    birim_sutun = get_birim_col(dataframe.columns)
    class PDF(FPDF):
        def header(self):
            if os.path.exists("arkaplan.png"): self.image("arkaplan.png", 0, 0, 210, 297)
            self.set_y(75)
            if self.page_no() == 1:
                self.set_font('Arial', 'B', 15)
                self.cell(0, 10, cevir_tr(sablon.upper()), 0, 1, 'C')

    pdf = PDF()
    pdf.set_margins(left=25, top=75, right=25)
    pdf.add_page()
    pdf.set_font('Arial', '', 10)
    pdf.cell(0, 10, f"Tarih: {tarih}", 0, 1, 'R')
    
    cols = ['Sıra'] + list(dataframe.columns)
    w = 160 / len(cols)
    pdf.set_font('Arial', 'B', 9)
    for col in cols: pdf.cell(w, 8, cevir_tr(str(col)), 1, 0, 'C', False)
    pdf.ln()
    
    pdf.set_font('Arial', '', 8)
    for i, row in dataframe.iterrows():
        pdf.cell(w, 7, str(i+1), 1, 0, 'C')
        for cname in dataframe.columns:
            val = row[cname]
            if gizle and cname == birim_sutun: val = "***"
            elif cname == dataframe.columns[-1]:
                try: val = f"{float(val):,.0f} {kur_m}"
                except: val = "-NIL-"
            pdf.cell(w, 7, cevir_tr(str(val)), 1, 0, 'C')
        pdf.ln()
    
    pdf.set_font('Arial', 'B', 9)
    pdf.cell(w*(len(cols)-1), 8, "GENEL TOPLAM:", 1, 0, 'R')
    pdf.cell(w, 8, g_str, 1, 1, 'C')
    
    return pdf.output(dest='S').encode('latin-1')

# --- BUTONLAR ---
st.markdown("### 📥 Çıktıları Al")
c1, c2, c3 = st.columns(3)
with c1:
    st.download_button("📊 EXCEL", data=excel_olustur(duzenlenmis_df, f"{g_toplam:,.0f} {kur_metin}", tarih_metni, secili_sablon, gizle_checkbox), file_name=f"Innomarin_{dosya_tarihi}.xlsx", type="primary", use_container_width=True)
with c2:
    st.download_button("📄 WORD", data=word_olustur(duzenlenmis_df, f"{ara_toplam:,.0f} {kur_metin}", f"{kdv:,.0f} {kur_metin}", f"{g_toplam:,.0f} {kur_metin}", tarih_metni, st.session_state.not_alani, kur_metin, secili_sablon, gizle_checkbox), file_name=f"Innomarin_{dosya_tarihi}.docx", type="primary", use_container_width=True)
with c3:
    st.download_button("📕 PDF", data=pdf_olustur(duzenlenmis_df, f"{ara_toplam:,.0f} {kur_metin}", f"{kdv:,.0f} {kur_metin}", f"{g_toplam:,.0f} {kur_metin}", tarih_metni, st.session_state.not_alani, kur_metin, secili_sablon, gizle_checkbox), file_name=f"Innomarin_{dosya_tarihi}.pdf", type="primary", use_container_width=True)
