import streamlit as st
import pandas as pd
import io
import datetime
import os
from fpdf import FPDF
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter

st.set_page_config(layout="wide", page_title="Innomar Doküman Platformu", initial_sidebar_state="expanded")

# --- PLATFORM ŞABLON SEÇİCİ ---
st.sidebar.markdown("### ⚙️ Sistem Ayarları")
secili_sablon = st.sidebar.radio(
    "📝 Çalışma Şablonunu Seçin:",
    ["⚓ INNOMAR Özel Teklif", "📄 Standart Proforma Fatura"]
)

gizle_checkbox = st.sidebar.checkbox("🔒 Birim Fiyatını Çıktılarda Gizle (Sansürle)", value=False)

if 'aktif_sablon' not in st.session_state or st.session_state.aktif_sablon != secili_sablon:
    st.session_state.aktif_sablon = secili_sablon
    if secili_sablon == "⚓ INNOMAR Özel Teklif":
        data = {
            'INSPECTION REMARK': ['ANA MAKİNE BAKIMLARI'],
            'UNIT': ['2 PIECES'],
            'PRICE': [40000.0]
        }
        st.session_state.not_alani = "* IMPORTANT NOTICE;\n- DURING MAINTENANCE IF DEFORMATION DETECTED..."
    else: 
        data = {
            'Marka': ['Örnek Marka', ''],
            'Açıklama': ['Hizmet Açıklaması', ''],
            'KDV': ['%20', '%20'],
            'Adet': ['1', '2'],
            'Birim Fiyat': ['1000', '500'],
            'Toplam Fiyat': [1000.0, 1000.0]
        }
        st.session_state.not_alani = "Banka Hesap Bilgilerimiz:\nIBAN: TR00 0000..."
    
    st.session_state.veri_df = pd.DataFrame(data)
    st.rerun()

st.markdown(f"<h2 style='text-align: center;'>{secili_sablon.upper()}</h2>", unsafe_allow_html=True)

# --- ÜST PANEL ---
col_t, col_kur = st.columns([1, 1])
secilen_tarih = col_t.date_input("Belge Tarihi", datetime.date.today())
tarih_metni = secilen_tarih.strftime("%d.%m.%Y")
dosya_tarihi = secilen_tarih.strftime("%d_%m_%Y")

kur_secimi = col_kur.selectbox("Para Birimi", ["Euro (€)", "Dolar ($)", "Türk Lirası (₺)"])
if "Euro" in kur_secimi: sembol, kur_metin = "€", "EURO"
elif "Dolar" in kur_secimi: sembol, kur_metin = "$", "USD"
else: sembol, kur_metin = "₺", "TL"

# --- TABLO DÜZENLEME ---
df = st.session_state.veri_df.copy()
son_sutun = df.columns[-1] 
df[son_sutun] = pd.to_numeric(df[son_sutun], errors='coerce').fillna(0.0)

col_config = {son_sutun: st.column_config.NumberColumn(son_sutun, format=f"%d {sembol}")}
duzenlenmis_df = st.data_editor(df, column_config=col_config, num_rows="dynamic", use_container_width=True)

# --- HESAPLAMALAR ---
fiyatlar = pd.to_numeric(duzenlenmis_df[son_sutun], errors='coerce').fillna(0)
ara_toplam = fiyatlar.sum()
kdv = ara_toplam * 0.20
genel_toplam = ara_toplam + kdv

ara_str = f"{ara_toplam:,.0f}".replace(",", ".") + f" {kur_metin}"
kdv_str = f"{kdv:,.0f}".replace(",", ".") + f" {kur_metin}"
genel_str = f"{genel_toplam:,.0f}".replace(",", ".") + f" {kur_metin}"

st.write("---")
c_a, c_b, c_c = st.columns(3)
c_a.metric("Ara Toplam", ara_str)
c_b.metric("KDV (%20)", kdv_str)
c_c.metric("Genel Toplam", genel_str)

st.subheader("📄 Belge Altı Notları")
st.text_area("Yazdırılacak Notlar:", key="not_alani", height=150)

# ==========================================
# YARDIMCI FONKSİYONLAR
# ==========================================
def cevir_tr(metin):
    tr_map = {'ş':'s', 'Ş':'S', 'ı':'i', 'İ':'I', 'ğ':'g', 'Ğ':'G', 'ü':'u', 'Ü':'U', 'ö':'o', 'Ö':'O', 'ç':'c', 'Ç':'C'}
    for k, v in tr_map.items(): metin = metin.replace(k, v)
    return metin

def get_birim_col(cols):
    for c in cols:
        if "birim fiyat" in str(c).lower(): return c
    return None

# ==========================================
# ÇIKTI MOTORLARI (İMZASIZ & YENİ ŞABLONLU)
# ==========================================
def word_olustur(dataframe, g_str, tarih, sablon, gizle):
    doc = Document()
    birim_sutun = get_birim_col(dataframe.columns)
    
    # Antet boşluğu
    doc.add_paragraph("\n" * 6)
    
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    baslik = "PROFORMA FATURA" if "Fatura" in sablon else "QUOTATION"
    p_title.add_run(f"{baslik}\n{tarih}").bold = True
    
    table = doc.add_table(rows=1, cols=len(dataframe.columns)+1)
    table.style = 'Table Grid'
    
    headers = ['Sıra'] + list(dataframe.columns)
    for i, h in enumerate(headers): table.rows[0].cells[i].text = str(h)
    
    for i, row in dataframe.iterrows():
        cells = table.add_row().cells
        cells[0].text = str(i+1)
        for c_idx, cname in enumerate(dataframe.columns, 1):
            val = row[cname]
            if gizle and cname == birim_sutun: val = "***"
            elif cname == dataframe.columns[-1]:
                try: val = f"{float(val):,.0f} {kur_metin}"
                except: val = str(val)
            cells[c_idx].text = str(val)
            
    doc.add_paragraph(f"\nGENEL TOPLAM: {g_str}")
    doc.add_paragraph(f"\nNotlar:\n{st.session_state.not_alani}")
    
    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()

def excel_olustur(dataframe, a_str, k_str, g_str, tarih, sablon, gizle):
    wb = Workbook()
    ws = wb.active
    ws.sheet_view.showGridLines = False
    
    birim_sutun = get_birim_col(dataframe.columns)
    # Yeni Excel Taslağına Göre Satır İndeksi (7. Satır)
    row_idx = 7 if "Fatura" in sablon else 9
    
    ws.cell(row=row_idx, column=2).value = "PROFORMA FATURA" if "Fatura" in sablon else "QUOTATION"
    ws.cell(row=row_idx, column=2).font = Font(bold=True, size=14)
    ws.cell(row=row_idx, column=len(dataframe.columns)+1).value = f"TARİH: {tarih}"
    row_idx += 2
    
    headers = ['Sıra'] + list(dataframe.columns)
    gray_fill = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
    
    for i, h in enumerate(headers, 2):
        cell = ws.cell(row=row_idx, column=i, value=h)
        cell.font = Font(bold=True)
        cell.fill = gray_fill
        cell.alignment = Alignment(horizontal="center")
        ws.column_dimensions[get_column_letter(i)].width = 18
        
    row_idx += 1
    for i, row in dataframe.iterrows():
        ws.cell(row=row_idx, column=2).value = i + 1
        for c_idx, cname in enumerate(dataframe.columns, 3):
            val = row[cname]
            if gizle and cname == birim_sutun: val = "***"
            ws.cell(row=row_idx, column=c_idx).value = str(val)
            ws.cell(row=row_idx, column=c_idx).alignment = Alignment(horizontal="center")
        row_idx += 1
        
    ws.cell(row=row_idx, column=len(headers)).value = "Ara Toplam:"
    ws.cell(row=row_idx, column=len(headers)+1).value = a_str
    row_idx += 1
    ws.cell(row=row_idx, column=len(headers)).value = "KDV % 20:"
    ws.cell(row=row_idx, column=len(headers)+1).value = k_str
    row_idx += 1
    ws.cell(row=row_idx, column=len(headers)).value = "GENEL TOPLAM:"
    ws.cell(row=row_idx, column=len(headers)+1).value = g_str
    ws.cell(row=row_idx, column=len(headers)+1).font = Font(bold=True)
    
    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()

def pdf_olustur(dataframe, a_str, k_str, g_str, tarih, sablon, gizle):
    birim_sutun = get_birim_col(dataframe.columns)
    class PDF(FPDF):
        def header(self):
            if os.path.exists("arkaplan.png"): self.image("arkaplan.png", 0, 0, 210, 297)
            # Fatura taslağına göre üst boşluk (70mm)
            self.set_y(70)
            if self.page_no() == 1:
                self.set_font('Arial', 'B', 15)
                self.cell(0, 10, cevir_tr("PROFORMA FATURA" if "Fatura" in sablon else "QUOTATION"), 0, 1, 'C')

    pdf = PDF()
    pdf.set_margins(left=25, top=70, right=25)
    pdf.add_page()
    pdf.set_font('Arial', '', 10)
    pdf.cell(0, 10, f"Tarih: {tarih}", 0, 1, 'R')
    
    cols = ['Sıra'] + list(dataframe.columns)
    w = 160 / len(cols)
    pdf.set_font('Arial', 'B', 9)
    for col in cols: pdf.cell(w, 8, cevir_tr(str(col)), 1, 0, 'C')
    pdf.ln()
    
    pdf.set_font('Arial', '', 8)
    for i, row in dataframe.iterrows():
        pdf.cell(w, 7, str(i+1), 1, 0, 'C')
        for cname in dataframe.columns:
            val = row[cname]
            if gizle and cname == birim_sutun: val = "***"
            elif cname == dataframe.columns[-1]:
                try: val = f"{float(val):,.0f} {kur_metin}"
                except: val = str(val)
            pdf.cell(w, 7, cevir_tr(str(val)), 1, 0, 'C')
        pdf.ln()
    
    pdf.set_font('Arial', 'B', 9)
    pdf.cell(w*(len(cols)-1), 8, "GENEL TOPLAM:", 1, 0, 'R')
    pdf.cell(w, 8, g_str, 1, 1, 'C')
    
    pdf.ln(10)
    pdf.set_font('Arial', '', 8)
    pdf.multi_cell(0, 5, cevir_tr(st.session_state.not_alani))
    
    return pdf.output(dest='S').encode('latin-1')

# --- BUTONLAR ---
st.markdown("### 📥 İndir")
c1, c2, c3 = st.columns(3)
with c1: st.download_button("📊 EXCEL", data=excel_olustur(duzenlenmis_df, ara_str, kdv_str, genel_str, tarih_metni, secili_sablon, gizle_checkbox), file_name=f"Innomar_{dosya_tarihi}.xlsx", type="primary", use_container_width=True)
with c2: st.download_button("📄 WORD", data=word_olustur(duzenlenmis_df, genel_str, tarih_metni, secili_sablon, gizle_checkbox), file_name=f"Innomar_{dosya_tarihi}.docx", type="primary", use_container_width=True)
with c3: st.download_button("📕 PDF", data=pdf_olustur(duzenlenmis_df, ara_str, kdv_str, genel_str, tarih_metni, secili_sablon, gizle_checkbox), file_name=f"Innomar_{dosya_tarihi}.pdf", type="primary", use_container_width=True)
