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
from openpyxl.drawing.image import Image as xlImage

st.set_page_config(layout="wide", page_title="Doküman Oluşturucu Platform", initial_sidebar_state="expanded")

# ==========================================
# ZIRHLI YARDIMCI FONKSİYONLAR
# ==========================================
def cevir_tr(metin):
    tr_map = {'ş':'s', 'Ş':'S', 'ı':'i', 'İ':'I', 'ğ':'g', 'Ğ':'G', 'ü':'u', 'Ü':'U', 'ö':'o', 'Ö':'O', 'ç':'c', 'Ç':'C'}
    for k, v in tr_map.items(): metin = metin.replace(k, v)
    return metin

# ==========================================
# PLATFORM ŞABLON SEÇİCİ
# ==========================================
st.sidebar.markdown("### ⚙️ Sistem Ayarları")
secili_sablon = st.sidebar.radio(
    "📝 Çalışma Şablonunu Seçin:",
    ["⚓ INNOMAR Özel Teklif", "📄 Standart Proforma Fatura"]
)

gizle_checkbox = st.sidebar.checkbox("🔒 Birim Fiyatını Çıktılarda Gizle (Sansürle)", value=False, help="İşaretlendiğinde, indirilen dosyalarda Birim Fiyat sütunu '***' olarak görünür.")

if 'aktif_sablon' not in st.session_state or st.session_state.aktif_sablon != secili_sablon:
    st.session_state.aktif_sablon = secili_sablon
    if secili_sablon == "⚓ INNOMAR Özel Teklif":
        data = {
            'INSPECTION REMARK': ['ANA MAKİNE BAKIMLARI', 'SU YAPICI BAKIMLARI'],
            'UNIT': ['2 PIECES', '1 SET'],
            'PRICE': [40000.0, 12000.0]
        }
        st.session_state.not_alani = "* IMPORTANT NOTICE;\n- DURING MAINTENANCE IF DEFORMATION DETECTED ON WORKING SURFACE AND NEEDED TO RENEW COMPONENTS EACH PARTS WILL BE PRICED ADDITIONALLY.\n\n* REMARKS;\n- DELIVERY TIME FOR THE JOB IS 35 DAYS,\n- A DETAILED REPORT WILL BE SUBMITTED TO YOUR SIDE UPON COMPLETION OF THE WORK,\n- PAYMENT WILL BE ACCEPTED AS BELOW;\n    - %50 BEFORE WORK BEGINS,\n    - %50 UPON COMPLETION OF THE WORK."
    else: 
        data = {
            'Marka': ['Örnek Marka', ''],
            'Açıklama': ['Örnek Açıklama', ''],
            'KDV': ['%20', ''],
            'Adet': ['1', '2'],
            'Birim Fiyat': ['1000', '500'],
            'Toplam Fiyat': [1000.0, 1000.0]
        }
        st.session_state.not_alani = "Banka Hesap Bilgilerimiz:\nBanka Adı: \nIBAN: \nHesap Sahibi: INNOMAR MARİNA YAT LİMAN A.Ş."
    
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

# --- DİNAMİK SÜTUN YÖNETİMİ ---
st.write("---")
st.caption("Aşağıdaki kutuya virgülle ayırarak istediğiniz kadar sütun ekleyebilir veya silebilirsiniz. **DİKKAT: Fiyat/tutar sütunu her zaman EN SONDA olmalıdır.**")

mevcut_sutunlar = ", ".join(st.session_state.veri_df.columns)
yeni_sutunlar_str = st.text_input("Tablo Sütunları:", mevcut_sutunlar)

yeni_sutunlar = [s.strip() for s in yeni_sutunlar_str.split(",") if s.strip()]
yeni_sutunlar = list(dict.fromkeys(yeni_sutunlar)) 

if len(yeni_sutunlar) < 2:
    st.warning("Lütfen tabloda en az 2 sütun bırakın.")
    yeni_sutunlar = st.session_state.veri_df.columns.tolist()

if yeni_sutunlar != st.session_state.veri_df.columns.tolist():
    eski_df = st.session_state.veri_df.copy()
    yeni_df = pd.DataFrame(columns=yeni_sutunlar)
    for col in yeni_sutunlar:
        if col in eski_df.columns:
            yeni_df[col] = eski_df[col]
        else:
            yeni_df[col] = "" 
            
    st.session_state.veri_df = yeni_df
    st.rerun()

# --- VERİ TİPİ DÜZENLEME (ÇÖKMEZ YAPI) ---
df = st.session_state.veri_df.copy()
son_sutun = df.columns[-1] 

for col in df.columns:
    if col == son_sutun:
        df[col] = pd.to_numeric(df[col].astype(str).str.replace(',', '.'), errors='coerce').fillna(0.0)
    else:
        df[col] = df[col].astype(str).replace(['nan', 'None', '<NA>', 'NaN'], '')

st.session_state.veri_df = df

col_config = {}
for col in df.columns[:-1]:
    col_config[col] = st.column_config.TextColumn(col)

col_config[son_sutun] = st.column_config.NumberColumn(son_sutun, format=f"%d {sembol}")

st.info("💡 Tablodaki hücrelerin üzerine tıklayıp değiştirebilirsiniz. Toplam Fiyat kısmını elle giriniz.")

tablo_key = f"editor_{secili_sablon}_{''.join(df.columns)}"
duzenlenmis_df = st.data_editor(df, column_config=col_config, num_rows="dynamic", use_container_width=True, key=tablo_key)

# --- BASİT HESAPLAMALAR ---
fiyatlar = pd.to_numeric(duzenlenmis_df[son_sutun].astype(str).str.replace(',', '.'), errors='coerce').fillna(0.0)
ara_toplam = fiyatlar.sum()
kdv = ara_toplam * 0.20
genel_toplam = ara_toplam + kdv

ara_str = f"{ara_toplam:,.0f}".replace(",", ".") + f" {kur_metin}"
kdv_str = f"{kdv:,.0f}".replace(",", ".") + f" {kur_metin}"
genel_str = f"{genel_toplam:,.0f}".replace(",", ".") + f" {kur_metin}"

st.write("---")
col_a, col_b, col_c = st.columns(3)
col_a.metric("Ara Toplam", ara_str)
col_b.metric("KDV (%20)", kdv_str)
col_c.metric("Genel Toplam", genel_str)
st.write("---")

st.subheader("📄 Belge Altı Notları")
st.text_area("Bu alana yazdığınız metin belgenin altına eklenecektir:", key="not_alani", height=150)

if st.button("🔄 Notları Sisteme Kaydet (İndirmeden Önce Basın)"):
    st.success("Notlarınız başarıyla hafızaya alındı! Çıktı alabilirsiniz.")
st.write("---")

# ==========================================
# AKILLI HİZALAMA VE GENİŞLİK ALGORİTMALARI
# ==========================================
def get_alignment(col_name):
    name = str(col_name).lower()
    if any(x in name for x in ['fiyat', 'price', 'tutar']): return 'R'
    if any(x in name for x in ['sıra', 'no', 'kdv', 'adet', 'unit']): return 'C'
    return 'L'

def get_pdf_widths(headers, total_w=150): # KANKA TABLOYU 150 BİRİME KÜÇÜLTTÜK
    widths = []
    for h in headers:
        name = str(h).lower()
        if any(x in name for x in ['sıra', 'no']): widths.append(10)
        elif any(x in name for x in ['kdv', 'adet', 'unit']): widths.append(15)
        elif any(x in name for x in ['fiyat', 'price', 'tutar']): widths.append(25)
        elif any(x in name for x in ['marka']): widths.append(25)
        elif any(x in name for x in ['açıklama', 'remark', 'işlem']): widths.append(50)
        else: widths.append(25)
    
    scale = total_w / sum(widths) if sum(widths) > 0 else 1
    return [w * scale for w in widths]

def set_excel_col_widths(ws, headers):
    # Excel sütun genişliklerini de biraz kıstık ki çok yayılmasın
    ws.column_dimensions['A'].width = 3 # Sol kenar boşluğu
    for i, header in enumerate(headers, 2):
        col_letter = get_column_letter(i)
        name = str(header).lower()
        if any(x in name for x in ['sıra', 'no', 'kdv', 'adet', 'unit']):
            ws.column_dimensions[col_letter].width = 8
        elif any(x in name for x in ['fiyat', 'price', 'tutar']):
            ws.column_dimensions[col_letter].width = 16
        elif any(x in name for x in ['marka']):
            ws.column_dimensions[col_letter].width = 15
        elif any(x in name for x in ['açıklama', 'remark', 'işlem']):
            ws.column_dimensions[col_letter].width = 35
        else:
            ws.column_dimensions[col_letter].width = 18

# ==========================================
# ÇIKTI MOTORLARI
# ==========================================
def get_birim_col(df_columns):
    for col in df_columns:
        if "birim fiyat" in str(col).lower(): return col
    return None

def word_olustur(dataframe, a_str, k_str, g_str, tarih, notlar, kur_m, sablon_tipi, gizle_aktif):
    doc = Document()
    
    # Word için yukarıdan boşluk bırakma (Antet payı)
    doc.add_paragraph()
    doc.add_paragraph()
    doc.add_paragraph()
    
    headers = ['Sıra' if sablon_tipi != "⚓ INNOMAR Özel Teklif" else 'ITEM NO'] + list(dataframe.columns)
    birim_sutun = get_birim_col(dataframe.columns)
    
    if sablon_tipi == "⚓ INNOMAR Özel Teklif":
        p_title = doc.add_paragraph()
        p_title.add_run(f"•   MY ADA DRY DOCK SERVICES QUOTATION;                                 * DATE: {tarih}").bold = True
    else:
        p_title = doc.add_paragraph()
        p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_title.add_run("PROFORMA FATURA").bold = True
        p_date = doc.add_paragraph()
        p_date.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p_date.add_run(f"TARİH: {tarih}").bold = True
    
    table = doc.add_table(rows=1, cols=len(headers))
    table.style = 'Table Grid'
    
    for idx, header in enumerate(headers):
        table.rows[0].cells[idx].text = str(header)
        table.rows[0].cells[idx].paragraphs[0].runs[0].font.bold = True
        align = get_alignment(header)
        if align == 'R': table.rows[0].cells[idx].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
        elif align == 'C': table.rows[0].cells[idx].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    for index, row in dataframe.iterrows():
        row_cells = table.add_row().cells
        row_cells[0].text = str(index + 1)
        row_cells[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        for c_idx, col_name in enumerate(dataframe.columns):
            val = row[col_name]
            align = get_alignment(col_name)
            
            if gizle_aktif and col_name == birim_sutun:
                row_cells[c_idx+1].text = "***"
                align = 'C'
            elif col_name == dataframe.columns[-1]: 
                try:
                    fiyat = float(str(val).replace(',', '.'))
                    row_cells[c_idx+1].text = "-NIL-" if pd.isna(fiyat) or fiyat <= 0 else f"{fiyat:,.0f}".replace(",", ".") + f" {kur_m}"
                except:
                    row_cells[c_idx+1].text = "-NIL-"
            else:
                row_cells[c_idx+1].text = str(val)
                
            if align == 'R': row_cells[c_idx+1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
            elif align == 'C': row_cells[c_idx+1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            
    doc.add_paragraph()
    
    tot_table = doc.add_table(rows=3, cols=2)
    tot_table.style = 'Table Grid'
    tot_table.alignment = WD_TABLE_ALIGNMENT.RIGHT
    
    tot_table.rows[0].cells[0].text = "TOTAL PRICE" if sablon_tipi == "⚓ INNOMAR Özel Teklif" else "Ara Toplam"
    tot_table.rows[0].cells[1].text = a_str
    tot_table.rows[1].cells[0].text = "VAT (20%)" if sablon_tipi == "⚓ INNOMAR Özel Teklif" else "KDV % 20"
    tot_table.rows[1].cells[1].text = k_str
    tot_table.rows[2].cells[0].text = "GRAND TOTAL" if sablon_tipi == "⚓ INNOMAR Özel Teklif" else "GENEL TOPLAM"
    tot_table.rows[2].cells[1].text = g_str
    
    for row in tot_table.rows:
        row.cells[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
        row.cells[1].paragraphs[0].runs[0].font.bold = True
                    
    doc.add_paragraph()
    for satir in notlar.split('\n'):
        doc.add_paragraph(satir)
        
    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()

def pdf_olustur(dataframe, a_str, k_str, g_str, tarih, notlar, kur_m, sablon_tipi, gizle_aktif):
    birim_sutun = get_birim_col(dataframe.columns)
    
    class PDF(FPDF):
        def header(self):
            # ARKA PLAN
            if os.path.exists("arkaplan.png"):
                self.image("arkaplan.png", 0, 0, 210, 297)
            
            if self.page_no() == 1:
                # YUKARIDAN BOŞLUK 70MM'YE ÇIKARILDI (Logodan tamamen kurtuldu)
                self.set_y(70) 
                if sablon_tipi != "⚓ INNOMAR Özel Teklif":
                    self.set_font('Arial', 'B', 16)
                    self.set_text_color(0, 0, 0)
                    self.cell(0, 10, 'PROFORMA FATURA', 0, 1, 'C')
                    self.ln(5)
            else:
                self.set_y(70)

        def footer(self):
            if os.path.exists("arkaplan.png"):
                self.set_font('Arial', '', 8)
                self.set_text_color(50, 50, 50)
                # İkonların Hizası
                self.set_xy(135, 260)
                self.cell(65, 4, cevir_tr('Heybeliada Mah. Kılavuz Sokak'), 0, 1, 'L')
                self.set_x(135)
                self.cell(65, 4, cevir_tr('Zarif Apt. No:16/6 Heybeliada/İST'), 0, 1, 'L')
                
                self.set_xy(135, 270)
                self.cell(65, 4, 'Phn: (+90) 536 763 1911', 0, 1, 'L')
                self.set_x(135)
                self.cell(65, 4, 'Mob: (+90) 541 552 1907', 0, 1, 'L')
                
                self.set_xy(135, 280)
                self.cell(65, 4, 'info@innomarin.com', 0, 1, 'L')
                self.set_x(135)
                self.cell(65, 4, 'www.innomarin.com', 0, 1, 'L')

    pdf = PDF()
    
    # KANKA TABLOYU DARALTTIĞIMIZ İÇİN SOLDAN 30MM BOŞLUK BIRAKTIK Kİ TAM ORTALANSIN
    pdf.set_margins(left=30, top=70, right=30) 
    pdf.set_auto_page_break(auto=True, margin=40)
    pdf.add_page()
    
    pdf.set_font('Arial', 'B', 10)
    pdf.set_text_color(0, 0, 0)
    if sablon_tipi == "⚓ INNOMAR Özel Teklif":
        pdf.cell(110, 10, chr(149) + '   MY ADA DRY DOCK SERVICES QUOTATION;', 0, 0, 'L')
        pdf.cell(40, 10, f'* DATE: {tarih}', 0, 1, 'R')
    else:
        pdf.cell(110, 10, '', 0, 0, 'L') 
        pdf.cell(40, 10, f'TARİH: {tarih}', 0, 1, 'R')
    pdf.ln(2)
    
    headers = ['Sıra' if sablon_tipi != "⚓ INNOMAR Özel Teklif" else 'NO'] + list(dataframe.columns)
    widths = get_pdf_widths(headers)
    
    pdf.set_draw_color(0, 0, 0)
    pdf.set_font('Arial', 'B', 9)
    for idx, header in enumerate(headers):
        pdf.cell(widths[idx], 8, cevir_tr(str(header)), 1, align=get_alignment(header))
    pdf.ln()
    
    pdf.set_font('Arial', '', 8)
    for index, row in dataframe.iterrows():
        pdf.cell(widths[0], 8, str(index + 1), 1, align='C')
        for c_idx, col_name in enumerate(dataframe.columns):
            val = row[col_name]
            align = get_alignment(col_name)
            
            if gizle_aktif and col_name == birim_sutun:
                yaz_fiyat = "***"
                align = 'C'
            elif col_name == dataframe.columns[-1]: 
                try:
                    fiyat = float(str(val).replace(',', '.'))
                    yaz_fiyat = "-NIL-" if pd.isna(fiyat) or fiyat <= 0 else f"{fiyat:,.0f}".replace(",", ".") + f" {kur_m}"
                except:
                    yaz_fiyat = "-NIL-"
            else:
                yaz_fiyat = "" if str(val) == "nan" else cevir_tr(str(val))
                
            pdf.cell(widths[c_idx+1], 8, yaz_fiyat, 1, align=align)
        pdf.ln()
        
    # Tablo genişliği 150 olduğu için toplam kısımları da orantılandı
    w_empty = sum(widths[:-2]) if len(widths) > 2 else 85
    w_label = widths[-2] if len(widths) > 2 else 40
    w_val = widths[-1]
    
    pdf.set_font('Arial', '', 9)
    pdf.cell(w_empty, 8, '', 0, 0)
    pdf.cell(w_label, 8, 'Ara Toplam' if sablon_tipi != "⚓ INNOMAR Özel Teklif" else 'TOTAL PRICE', 1, 0, 'L')
    pdf.set_font('Arial', 'B', 9)
    pdf.cell(w_val, 8, a_str, 1, 1, 'R')
    
    pdf.set_font('Arial', '', 9)
    pdf.cell(w_empty, 8, '', 0, 0)
    pdf.cell(w_label, 8, 'KDV % 20' if sablon_tipi != "⚓ INNOMAR Özel Teklif" else 'VAT (20%)', 1, 0, 'L')
    pdf.set_font('Arial', 'B', 9)
    pdf.cell(w_val, 8, k_str, 1, 1, 'R')
    
    pdf.set_font('Arial', '', 9)
    pdf.cell(w_empty, 8, '', 0, 0)
    pdf.cell(w_label, 8, 'GENEL TOPLAM' if sablon_tipi != "⚓ INNOMAR Özel Teklif" else 'GRAND TOTAL', 1, 0, 'L')
    pdf.set_font('Arial', 'B', 9)
    pdf.cell(w_val, 8, g_str, 1, 1, 'R')
    
    pdf.ln(10)
    pdf.set_font('Arial', '', 8)
    pdf.multi_cell(0, 5, cevir_tr(notlar))

    return pdf.output(dest='S').encode('latin-1')

def excel_olustur(dataframe, a_str, k_str, g_str, tarih, notlar, kur_m, sablon_tipi, gizle_aktif):
    wb = Workbook()
    ws = wb.active
    ws.title = "Belge"
    
    # EXCEL'İ A4 KAĞIDI GİBİ GÖSTERMEK İÇİN IZGARA ÇİZGİLERİNİ KAPATTIK
    ws.sheet_view.showGridLines = False 
    
    birim_sutun = get_birim_col(dataframe.columns)
    
    # Yukarıdan antet logoları için boşluk (10 satır)
    row_idx = 10 
    
    if sablon_tipi == "⚓ INNOMAR Özel Teklif":
        ws[f'C{row_idx}'] = "• MY ADA DRY DOCK SERVICES QUOTATION;"
        ws[f'C{row_idx}'].font = Font(bold=True)
        ws.cell(row=row_idx, column=len(dataframe.columns)+2).value = f"* DATE: {tarih}"
        ws.cell(row=row_idx, column=len(dataframe.columns)+2).font = Font(bold=True)
        row_idx += 2
    else:
        ws[f'C{row_idx}'] = "PROFORMA FATURA"
        ws[f'C{row_idx}'].font = Font(bold=True, size=16)
        row_idx += 1
        ws.cell(row=row_idx, column=len(dataframe.columns)+2).value = f"TARİH: {tarih}"
        ws.cell(row=row_idx, column=len(dataframe.columns)+2).font = Font(bold=True)
        row_idx += 2
    
    headers = ['Sıra' if sablon_tipi != "⚓ INNOMAR Özel Teklif" else 'ITEM NO'] + list(dataframe.columns)
    set_excel_col_widths(ws, headers)
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    gray_fill = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
    
    # Excel'de tabloyu B sütunundan başlattık (A sütunu sol boşluk)
    for col_num, header in enumerate(headers, 2):
        cell = ws.cell(row=row_idx, column=col_num)
        cell.value = str(header)
        cell.font = Font(bold=True)
        cell.border = thin_border
        cell.fill = gray_fill
        align = get_alignment(header)
        if align == 'R': cell.alignment = Alignment(horizontal="right")
        elif align == 'C': cell.alignment = Alignment(horizontal="center")
        else: cell.alignment = Alignment(horizontal="left")
    row_idx += 1
    
    for index, row in dataframe.iterrows():
        cell = ws.cell(row=row_idx, column=2)
        cell.value = index + 1
        cell.border = thin_border
        cell.alignment = Alignment(horizontal="center")
        
        for c_idx, col_name in enumerate(dataframe.columns):
            val = row[col_name]
            cell = ws.cell(row=row_idx, column=c_idx+3)
            align = get_alignment(col_name)
            
            if gizle_aktif and col_name == birim_sutun:
                cell.value = "***"
                cell.alignment = Alignment(horizontal="center")
            elif col_name == dataframe.columns[-1]: 
                try:
                    fiyat = float(str(val).replace(',', '.'))
                    if pd.isna(fiyat) or fiyat <= 0:
                        cell.value = "-NIL-"
                    else:
                        cell.value = f"{fiyat:,.0f}".replace(",", ".") + f" {kur_m}"
                except:
                    cell.value = "-NIL-"
                if align == 'R': cell.alignment = Alignment(horizontal="right")
                elif align == 'C': cell.alignment = Alignment(horizontal="center")
                else: cell.alignment = Alignment(horizontal="left")
            else:
                cell.value = str(val) if str(val) != "nan" else ""
                if align == 'R': cell.alignment = Alignment(horizontal="right")
                elif align == 'C': cell.alignment = Alignment(horizontal="center")
                else: cell.alignment = Alignment(horizontal="left")
                
            cell.border = thin_border
        row_idx += 1
        
    tot_col = len(dataframe.columns) + 1
    val_col = len(dataframe.columns) + 2
    
    ws.cell(row=row_idx, column=tot_col).value = "Ara Toplam" if sablon_tipi != "⚓ INNOMAR Özel Teklif" else "TOTAL PRICE"
    ws.cell(row=row_idx, column=tot_col).border = thin_border
    ws.cell(row=row_idx, column=val_col).value = a_str
    ws.cell(row=row_idx, column=val_col).font = Font(bold=True)
    ws.cell(row=row_idx, column=val_col).border = thin_border
    ws.cell(row=row_idx, column=val_col).alignment = Alignment(horizontal="right")
    row_idx += 1
    
    ws.cell(row=row_idx, column=tot_col).value = "KDV % 20" if sablon_tipi != "⚓ INNOMAR Özel Teklif" else "VAT (20%)"
    ws.cell(row=row_idx, column=tot_col).border = thin_border
    ws.cell(row=row_idx, column=val_col).value = k_str
    ws.cell(row=row_idx, column=val_col).font = Font(bold=True)
    ws.cell(row=row_idx, column=val_col).border = thin_border
    ws.cell(row=row_idx, column=val_col).alignment = Alignment(horizontal="right")
    row_idx += 1
    
    ws.cell(row=row_idx, column=tot_col).value = "GENEL TOPLAM" if sablon_tipi != "⚓ INNOMAR Özel Teklif" else "GRAND TOTAL"
    ws.cell(row=row_idx, column=tot_col).border = thin_border
    ws.cell(row=row_idx, column=val_col).value = g_str
    ws.cell(row=row_idx, column=val_col).font = Font(bold=True)
    ws.cell(row=row_idx, column=val_col).border = thin_border
    ws.cell(row=row_idx, column=val_col).alignment = Alignment(horizontal="right")
    row_idx += 2
    
    for satir in notlar.split('\n'):
        ws[f'C{row_idx}'] = satir
        row_idx += 1
        
    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()

# --- İNDİRME BUTONLARI ---
st.markdown("### 📥 Çıktı Al")
btn_word, btn_excel, btn_pdf = st.columns(3)

with btn_word:
    st.download_button("📄 WORD İNDİR", data=word_olustur(duzenlenmis_df, ara_str, kdv_str, genel_str, tarih_metni, st.session_state.not_alani, kur_metin, secili_sablon, gizle_checkbox), file_name=f"{secili_sablon.split()[1]}_{dosya_tarihi}.docx", type="primary", use_container_width=True)
with btn_excel:
    st.download_button("📊 EXCEL İNDİR", data=excel_olustur(duzenlenmis_df, ara_str, kdv_str, genel_str, tarih_metni, st.session_state.not_alani, kur_metin, secili_sablon, gizle_checkbox), file_name=f"{secili_sablon.split()[1]}_{dosya_tarihi}.xlsx", type="primary", use_container_width=True)
with btn_pdf:
    st.download_button("📕 PDF İNDİR", data=pdf_olustur(duzenlenmis_df, ara_str, kdv_str, genel_str, tarih_metni, st.session_state.not_alani, kur_metin, secili_sablon, gizle_checkbox), file_name=f"{secili_sablon.split()[1]}_{dosya_tarihi}.pdf", type="primary", use_container_width=True)
