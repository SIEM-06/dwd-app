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

st.set_page_config(layout="wide", page_title="Doküman Oluşturucu Platform", initial_sidebar_state="expanded")

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
            'INSPECTION REMARK': ['ANA MAKİNE BAKIMLARI', 'SU YAPICI BAKIMLARI'],
            'UNIT': ['2 PIECES', '1 SET'],
            'PRICE': [40000.0, 12000.0]
        }
        st.session_state.not_alani = "* IMPORTANT NOTICE;\n- DURING MAINTENANCE IF DEFORMATION DETECTED ON WORKING SURFACE AND NEEDED TO RENEW COMPONENTS EACH PARTS WILL BE PRICED ADDITIONALLY."
    else: 
        # Yeni attığın Proforma Fatura taslağındaki sütun yapısı
        data = {
            'Marka': ['Innomarin', ''],
            'Açıklama': ['Bakım ve Onarım Hizmeti', ''],
            'KDV': ['%20', '%20'],
            'Adet': ['1', '1'],
            'Birim Fiyat': ['1000', '500'],
            'Toplam Fiyat': [1000.0, 500.0]
        }
        st.session_state.not_alani = "Banka Hesap Bilgilerimiz:\nIBAN: TR00 0000 0000 0000 0000 0000 00\nHesap Sahibi: INNOMAR MARİNA YAT LİMAN A.Ş."
    
    st.session_state.veri_df = pd.DataFrame(data)
    st.rerun()

st.markdown(f"<h2 style='text-align: center;'>{secili_sablon.upper()} SİSTEMİ</h2>", unsafe_allow_html=True)

# --- ÜST PANEL: TARİH VE PARA BİRİMİ ---
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
c1, c2, c3 = st.columns(3)
c1.metric("Ara Toplam", ara_str)
c2.metric("KDV (%20)", kdv_str)
c3.metric("Genel Toplam", genel_str)

st.subheader("📄 Belge Altı Notları")
st.text_area("Notları Düzenle:", key="not_alani", height=150)

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
# ÇIKTI MOTORLARI (TAM TASLAK UYUMLU & İMZASIZ)
# ==========================================

def excel_olustur(dataframe, a_str, k_str, g_str, tarih, sablon, gizle):
    wb = Workbook()
    ws = wb.active
    ws.sheet_view.showGridLines = False
    birim_sutun = get_birim_col(dataframe.columns)
    
    # Excel Taslağındaki Gibi Başlık ve Tarih Yerleşimi
    row_idx = 7
    ws.cell(row=row_idx, column=1).value = "PROFORMA FATURA" if "Fatura" in sablon else "ÖZEL TEKLİF"
    ws.cell(row=row_idx, column=1).font = Font(bold=True, size=14)
    ws.cell(row=row_idx+1, column=len(dataframe.columns)).value = "TARİH:"
    ws.cell(row=row_idx+1, column=len(dataframe.columns)+1).value = tarih
    
    row_idx = 10 # Tablo Başlangıcı
    headers = ['Sıra'] + list(dataframe.columns)
    gray_fill = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
    border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    for i, h in enumerate(headers, 1):
        cell = ws.cell(row=row_idx, column=i, value=h)
        cell.font = Font(bold=True)
        cell.fill = gray_fill
        cell.border = border
        cell.alignment = Alignment(horizontal="center")
        ws.column_dimensions[get_column_letter(i)].width = 18

    row_idx += 1
    for i, row in dataframe.iterrows():
        ws.cell(row=row_idx, column=1).value = i + 1
        ws.cell(row=row_idx, column=1).border = border
        for c_idx, cname in enumerate(dataframe.columns, 2):
            val = row[cname-2]
            if gizle and cname-2 == birim_sutun: val = "***"
            elif cname-2 == dataframe.columns[-1]:
                try: val = f"{float(val):,.0f} {kur_metin}"
                except: val = str(val)
            
            cell = ws.cell(row=row_idx, column=c_idx, value=str(val))
            cell.border = border
            cell.alignment = Alignment(horizontal="center")
        row_idx += 1

    # Toplamlar
    ws.cell(row=row_idx+1, column=len(headers)).value = "Ara Toplam:"
    ws.cell(row=row_idx+1, column=len(headers)+1).value = a_str
    ws.cell(row=row_idx+2, column=len(headers)).value = "KDV % 20:"
    ws.cell(row=row_idx+2, column=len(headers)+1).value = k_str
    ws.cell(row=row_idx+3, column=len(headers)).value = "GENEL TOPLAM:"
    ws.cell(row=row_idx+3, column=len(headers)+1).value = g_str
    ws.cell(row=row_idx+3, column=len(headers)+1).font = Font(bold=True)

    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()

def pdf_olustur(dataframe, a_str, k_str, g_str, tarih, notlar, kur_m, sablon, gizle):
    birim_sutun = get_birim_col(dataframe.columns)
    class PDF(FPDF):
        def header(self):
            if os.path.exists("arkaplan.png"): self.image("arkaplan.png", 0, 0, 210, 297)
            self.set_y(70) # Taslağa göre üst boşluk
            if self.page_no() == 1:
                self.set_font('Arial', 'B', 15)
                self.cell(0, 10, cevir_tr(sablon.upper()), 0, 1, 'C')

    pdf = PDF()
    pdf.set_margins(left=20, top=70, right=20)
    pdf.add_page()
    pdf.set_font('Arial', '', 10)
    pdf.cell(0, 10, f"Tarih: {tarih}", 0, 1, 'R')
    
    cols = ['Sıra'] + list(dataframe.columns)
    w = 170 / len(cols)
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
    pdf.multi_cell(0, 5, cevir_tr(notlar))
    return pdf.output(dest='S').encode('latin-1')

def word_olustur(dataframe, g_str, tarih, sablon, gizle):
    doc = Document()
    birim_sutun = get_birim_col(dataframe.columns)
    doc.add_paragraph("\n" * 6) # Üst boşluk payı
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.add_run(f"{sablon.upper()}\n{tarih}").bold = True
    
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

# --- BUTONLAR ---
st.markdown("### 📥 Çıktıları Al")
btn_excel, btn_word, btn_pdf = st.columns(3)
with btn_excel: st.download_button("📊 EXCEL İNDİR", data=excel_olustur(duzenlenmis_df, ara_str, kdv_str, genel_str, tarih_metni, secili_sablon, gizle_checkbox), file_name=f"Innomar_{dosya_tarihi}.xlsx", type="primary", use_container_width=True)
with btn_word: st.download_button("📄 WORD İNDİR", data=word_olustur(duzenlenmis_df, genel_str, tarih_metni, secili_sablon, gizle_checkbox), file_name=f"Innomar_{dosya_tarihi}.docx", type="primary", use_container_width=True)
with btn_pdf: st.download_button("📕 PDF İNDİR", data=pdf_olustur(duzenlenmis_df, ara_str, kdv_str, genel_str, tarih_metni, st.session_state.not_alani, kur_metin, secili_sablon, gizle_checkbox), file_name=f"Innomar_{dosya_tarihi}.pdf", type="primary", use_container_width=True)
