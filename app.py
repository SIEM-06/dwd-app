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
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.drawing.image import Image as xlImage

st.set_page_config(layout="wide", page_title="Doküman Oluşturucu Platform", initial_sidebar_state="expanded")

# --- PLATFORM ŞABLON SEÇİCİ ---
st.sidebar.markdown("### ⚙️ Sistem Ayarları")
secili_sablon = st.sidebar.radio(
    "📝 Çalışma Şablonunu Seçin:",
    ["⚓ INNOMAR Özel Teklif", "📄 Standart Proforma Fatura"]
)

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
            'KDV': ['%20', '%20'],
            'Adet': ['1', '2'],
            'Birim Fiyat': ['1000', '500'],
            'Toplam Fiyat': [1000.0, 1000.0]
        }
        st.session_state.not_alani = "Banka Hesap Bilgilerimiz:\nBanka Adı: \nIBAN: \nHesap Sahibi: INNOMAR MARİNA YAT LİMAN A.Ş." # Banka bilgileri
    
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
st.caption("Aşağıdaki kutuya virgülle ayırarak istediğiniz kadar sütun ekleyebilir veya silebilirsiniz. **DİKKAT: Hesaplamaların doğru çalışması için fiyat/tutar sütunu her zaman EN SONDA olmalıdır.**")

mevcut_sutunlar = ", ".join(st.session_state.veri_df.columns)
yeni_sutunlar_str = st.text_input("Tablo Sütunları:", mevcut_sutunlar)

yeni_sutunlar = [s.strip() for s in yeni_sutunlar_str.split(",") if s.strip()]
yeni_sutunlar = list(dict.fromkeys(yeni_sutunlar)) 

if len(yeni_sutunlar) < 2:
    st.warning("Lütfen tabloda en az 2 sütun bırakın.")
    yeni_sutunlar = st.session_state.veri_df.columns.tolist()

if yeni_sutunlar != st.session_state.veri_df.columns.tolist():
    eski_df = st.session_state.veri_df
    yeni_df = pd.DataFrame(columns=yeni_sutunlar)
    for col in yeni_sutunlar:
        if col in eski_df.columns:
            yeni_df[col] = eski_df[col]
        else:
            yeni_df[col] = "" 
            
    son_sutun_adi = yeni_sutunlar[-1]
    yeni_df[son_sutun_adi] = pd.to_numeric(yeni_df[son_sutun_adi], errors='coerce').fillna(0.0)
    st.session_state.veri_df = yeni_df
    st.rerun()

df = st.session_state.veri_df
son_sutun = df.columns[-1] 
df[son_sutun] = pd.to_numeric(df[son_sutun], errors='coerce').fillna(0.0)

col_config = {}
for col in df.columns[:-1]:
    col_config[col] = st.column_config.TextColumn(col)
col_config[son_sutun] = st.column_config.NumberColumn(son_sutun, format=f"%d {sembol}")

st.info("💡 Tablodaki hücrelerin üzerine tıklayıp değiştirebilirsiniz. Yeni satır için tablonun en altını kullanın.")

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
col_a, col_b, col_c = st.columns(3)
col_a.metric("Ara Toplam", ara_str)
col_b.metric("KDV (%20)", kdv_str)
col_c.metric("Genel Toplam", genel_str)
st.write("---")

# --- ÖZELLEŞTİRİLEBİLİR ALT NOTLAR ---
st.subheader("📄 Belge Altı Notları")
st.text_area("Bu alana yazdığınız metin belgenin altına eklenecektir:", key="not_alani", height=150)

if st.button("🔄 Notları Sisteme Kaydet (İndirmeden Önce Basın)"):
    st.success("Notlarınız başarıyla hafızaya alındı! Çıktı alabilirsiniz.")
st.write("---")

# ==========================================
# ÇIKTI MOTORLARI
# ==========================================
def cevir_tr(metin):
    tr_map = {'ş':'s', 'Ş':'S', 'ı':'i', 'İ':'I', 'ğ':'g', 'Ğ':'G', 'ü':'u', 'Ü':'U', 'ö':'o', 'Ö':'O', 'ç':'c', 'Ç':'C'}
    for k, v in tr_map.items(): metin = metin.replace(k, v)
    return metin

def word_olustur(dataframe, a_str, k_str, g_str, tarih, notlar, kur_m, sablon_tipi):
    doc = Document()
    
    if os.path.exists("logo.png"):
        p_logo = doc.add_paragraph()
        p_logo.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r_logo = p_logo.add_run()
        r_logo.add_picture("logo.png", width=Cm(6))
        
    p_info = doc.add_paragraph()
    run_name = p_info.add_run("INNOMAR MARİNA YAT\nLİMAN TURİZM İŞLETMECİLİĞİ VE İNŞAAT SANAYİ VE TİCARET A.Ş.\n")
    run_name.bold = True
    run_name.font.color.rgb = RGBColor(0, 51, 153)
    run_name.font.size = Pt(10)
    
    run_addr = p_info.add_run("Bahçelievler Mah Şehit Fethi Cad. Duygu Sokak No.3 İç Kapı No. 7\nPendik - ISTANBUL/TURKEY\nPhn- (+90) 536 763 1911 | Mob- (+90) 541 552 1907\nEmail- info@inno-mar.com.tr | www.inno-mar.com.tr")
    run_addr.font.size = Pt(9)
    doc.add_paragraph("_" * 75)
    
    if sablon_tipi == "⚓ INNOMAR Özel Teklif":
        p_title = doc.add_paragraph()
        run_title = p_title.add_run(f"•   MY ADA DRY DOCK SERVICES QUOTATION;                                 * DATE: {tarih}")
        run_title.bold = True
        run_title.font.size = Pt(10)
    else:
        # PROFORMA WORD BAŞLIĞI
        p_title = doc.add_paragraph()
        p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run_title = p_title.add_run("PROFORMA FATURA")
        run_title.bold = True
        run_title.font.size = Pt(16)
        
        doc.add_paragraph()
        
        p_date = doc.add_paragraph()
        p_date.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        run_date = p_date.add_run(f"TARİH: {tarih}")
        run_date.bold = True
    
    table = doc.add_table(rows=1, cols=len(dataframe.columns) + 1)
    table.style = 'Table Grid'
    
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = 'ITEM NO' if sablon_tipi == "⚓ INNOMAR Özel Teklif" else 'SIRA'
    for idx, col_name in enumerate(dataframe.columns):
        hdr_cells[idx+1].text = str(col_name)
        
    for cell in hdr_cells:
        for paragraph in cell.paragraphs:
            for run in paragraph.runs:
                run.font.bold = True
    
    for index, row in dataframe.iterrows():
        row_cells = table.add_row().cells
        row_cells[0].text = str(index + 1)
        for idx, col_name in enumerate(dataframe.columns):
            val = row[col_name]
            if col_name == dataframe.columns[-1]: 
                fiyat = pd.to_numeric(val, errors='coerce')
                if pd.isna(fiyat) or fiyat <= 0:
                    row_cells[idx+1].text = "-NIL-"
                else:
                    row_cells[idx+1].text = f"{fiyat:,.0f}".replace(",", ".") + f" {kur_m}"
            else:
                row_cells[idx+1].text = str(val)
                
    if sablon_tipi == "⚓ INNOMAR Özel Teklif":
        mid_cols = dataframe.columns[:-1]
        w_mids = []
        if len(mid_cols) == 1: w_mids = [14.0]
        elif len(mid_cols) == 2: w_mids = [9.5, 4.5]
        else:
            w_first = 14.0 - (len(mid_cols)-1)*3.0
            w_mids = [w_first] + [3.0]*(len(mid_cols)-1)
            
        widths = [Cm(1.5)] + [Cm(w) for w in w_mids] + [Cm(3.5)]
        for row in table.rows:
            for idx, width in enumerate(widths):
                row.cells[idx].width = width
            
    doc.add_paragraph()
    
    tot_table = doc.add_table(rows=3, cols=2)
    tot_table.style = 'Table Grid'
    tot_table.alignment = WD_TABLE_ALIGNMENT.RIGHT
    
    tot_table.rows[0].cells[0].text = "TOTAL PRICE" if sablon_tipi == "⚓ INNOMAR Özel Teklif" else "ARA TOPLAM"
    tot_table.rows[0].cells[1].text = a_str
    tot_table.rows[1].cells[0].text = "VAT (20%)" if sablon_tipi == "⚓ INNOMAR Özel Teklif" else "KDV (20%)"
    tot_table.rows[1].cells[1].text = k_str
    tot_table.rows[2].cells[0].text = "GRAND TOTAL" if sablon_tipi == "⚓ INNOMAR Özel Teklif" else "GENEL TOPLAM"
    tot_table.rows[2].cells[1].text = g_str
    
    for row in tot_table.rows:
        row.cells[1].paragraphs[0].runs[0].font.bold = True
        
    tot_widths = [Cm(4.5), Cm(3.5)]
    for row in tot_table.rows:
        for idx, width in enumerate(tot_widths):
            row.cells[idx].width = width
                    
    doc.add_paragraph()
    for satir in notlar.split('\n'):
        doc.add_paragraph(satir)
        
    doc.add_paragraph("\nİlker TEKINKAYA | Managing Partner | INNOMAR MARİNA YAT\nLİMAN TURİZM İŞLETMECİLİĞİ VE İNŞAAT SANAYİ VE TİCARET A.Ş.").runs[0].bold = True # CE kaldırıldı
    doc.add_paragraph("Bahçelievler Mah Şehit Fethi Cad. Duygu Sokak No.3 İç Kapı No. 7\nPendik - ISTANBUL/TURKEY\nPhn- (+90) 536 763 1911 | Mob- (+90) 541 552 1907\nEmail- info@inno-mar.com.tr | www.inno-mar.com.tr")
    
    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()

def pdf_olustur(dataframe, a_str, k_str, g_str, tarih, notlar, kur_m, sablon_tipi):
    class PDF(FPDF):
        def header(self):
            if self.page_no() == 1:
                if os.path.exists("logo.png"):
                    self.image("logo.png", x=65, y=10, w=80)
                self.ln(25)
                
                self.set_font('Arial', 'B', 10)
                self.set_text_color(0, 51, 153)
                self.cell(0, 5, cevir_tr('INNOMAR MARİNA YAT'), 0, 1, 'L')
                self.cell(0, 5, cevir_tr('LİMAN TURİZM İŞLETMECİLİĞİ VE İNŞAAT SANAYİ VE TİCARET A.Ş.'), 0, 1, 'L')
                
                self.set_font('Arial', '', 9)
                self.set_text_color(0, 0, 0)
                self.cell(0, 5, cevir_tr('Bahçelievler Mah Şehit Fethi Cad. Duygu Sokak No.3 İç Kapı No. 7'), 0, 1, 'L')
                self.cell(0, 5, 'Pendik - ISTANBUL/TURKEY', 0, 1, 'L')
                self.cell(0, 5, 'Phn- (+90) 536 763 1911 | Mob- (+90) 541 552 1907', 0, 1, 'L')
                
                self.set_text_color(0, 51, 153)
                self.cell(0, 5, 'Email- info@inno-mar.com.tr | www.inno-mar.com.tr', 0, 1, 'L')
                
                self.set_draw_color(0, 51, 153)
                self.set_line_width(0.3)
                self.line(10, self.get_y()+2, 200, self.get_y()+2)
                self.ln(10)
                
                if sablon_tipi != "⚓ INNOMAR Özel Teklif":
                    self.set_font('Arial', 'B', 16)
                    self.set_text_color(0, 0, 0)
                    self.cell(0, 10, 'PROFORMA FATURA', 0, 1, 'C')
                    self.ln(2)
                    
                if os.path.exists("watermark.png"):
                    self.image("watermark.png", x=30, y=80, w=150)
            else:
                self.ln(15)

    pdf = PDF()
    pdf.add_page()
    pdf.set_auto_page_break(auto=True, margin=15)
    
    pdf.set_font('Arial', 'B', 10)
    pdf.set_text_color(0, 0, 0)
    if sablon_tipi == "⚓ INNOMAR Özel Teklif":
        pdf.cell(130, 10, chr(149) + '   MY ADA DRY DOCK SERVICES QUOTATION;', 0, 0, 'L')
    else:
        pdf.cell(130, 10, '', 0, 0, 'L') 
        
    pdf.cell(60, 10, f'* DATE/TARIH: {tarih}', 0, 1, 'R')
    pdf.ln(2)
    
    # Dinamik Sütun Genişlikleri Hesaplama (Optimized)
    cols = list(dataframe.columns)
    mid_cols = cols[:-1]
    last_col = cols[-1]
    w_item, w_price = 15, 35
    w_first = 140 - (len(mid_cols)-1)*25 # Genişletilmiş 'UNIT' sütunu için
    w_mids = [w_first] + [25]*(len(mid_cols)-1) if len(mid_cols) > 1 else [140]
    
    pdf.set_draw_color(0, 0, 0)
    pdf.set_font('Arial', 'B', 9)
    pdf.cell(w_item, 8, 'NO', 1)
    for idx, col_name in enumerate(mid_cols):
        pdf.cell(w_mids[idx], 8, cevir_tr(str(col_name)), 1)
    pdf.cell(w_price, 8, cevir_tr(str(last_col)), 1)
    pdf.ln()
    
    pdf.set_font('Arial', '', 8)
    for index, row in dataframe.iterrows():
        pdf.cell(w_item, 8, str(index + 1), 1)
        for idx, col_name in enumerate(mid_cols):
            pdf.cell(w_mids[idx], 8, cevir_tr(str(row[col_name])), 1)
            
        fiyat = pd.to_numeric(row[last_col], errors='coerce')
        fiyat_str = "-NIL-" if pd.isna(fiyat) or fiyat <= 0 else f"{fiyat:,.0f}".replace(",", ".") + f" {kur_m}"
        pdf.cell(w_price, 8, fiyat_str, 1)
        pdf.ln()
        
    pdf.set_font('Arial', '', 9)
    pdf.cell(110, 8, '', 0, 0)
    pdf.cell(45, 8, 'ARA TOPLAM' if sablon_tipi != "⚓ INNOMAR Özel Teklif" else 'TOTAL PRICE', 1, 0, 'L')
    pdf.set_font('Arial', 'B', 9)
    pdf.cell(35, 8, a_str, 1, 1, 'L')
    
    pdf.set_font('Arial', '', 9)
    pdf.cell(110, 8, '', 0, 0)
    pdf.cell(45, 8, 'KDV (20%)' if sablon_tipi != "⚓ INNOMAR Özel Teklif" else 'VAT (20%)', 1, 0, 'L')
    pdf.set_font('Arial', 'B', 9)
    pdf.cell(35, 8, k_str, 1, 1, 'L')
    
    pdf.set_font('Arial', '', 9)
    pdf.cell(110, 8, '', 0, 0)
    pdf.cell(45, 8, 'GENEL TOPLAM' if sablon_tipi != "⚓ INNOMAR Özel Teklif" else 'GRAND TOTAL', 1, 0, 'L')
    pdf.set_font('Arial', 'B', 9)
    pdf.cell(35, 8, g_str, 1, 1, 'L')
    
    pdf.ln(10)
    
    # Kullanıcının girdiği notları yazdır
    pdf.set_font('Arial', '', 8)
    pdf.multi_cell(0, 5, cevir_tr(notlar))
    
    pdf.ln(10)
    
    pdf.set_font('Arial', 'B', 8)
    pdf.cell(0, 4, cevir_tr('İlker TEKINKAYA | Managing Partner | INNOMAR MARINA YAT'), 0, 1, 'L') # CE kaldırıldı
    pdf.cell(0, 4, cevir_tr('LIMAN TURIZM ISLETMECILIGI VE INSAAT SANAYI VE TICARET A.S.'), 0, 1, 'L')
    pdf.set_font('Arial', '', 8)
    pdf.cell(0, 4, cevir_tr('Bahcelievler Mah Sehit Fethi Cad. Duygu Sokak No.3 Ic Kapi No. 7'), 0, 1, 'L')
    pdf.cell(0, 4, 'Pendik - ISTANBUL/TURKEY', 0, 1, 'L')
    pdf.cell(0, 4, 'Phn- (+90) 536 763 1911 | Mob- (+90) 541 552 1907', 0, 1, 'L')
    
    pdf.set_text_color(0, 51, 153)
    pdf.cell(0, 4, 'Email- info@inno-mar.com.tr | www.inno-mar.com.tr', 0, 1, 'L')
    
    return pdf.output(dest='S').encode('latin-1')

def excel_olustur(dataframe, a_str, k_str, g_str, tarih, notlar, kur_m, sablon_tipi):
    wb = Workbook()
    ws = wb.active
    ws.title = "Belge"
    row_idx = 1
    
    if os.path.exists("logo.png"):
        img = xlImage("logo.png")
        ws.add_image(img, 'B1')
        row_idx = 8
        
    ws[f'B{row_idx}'] = "INNOMAR MARİNA YAT LİMAN TURİZM İŞLETMECİLİĞİ VE İNŞAAT SANAYİ VE TİCARET A.Ş."
    ws[f'B{row_idx}'].font = Font(color="003399", bold=True, size=11)
    row_idx += 1
    ws[f'B{row_idx}'] = "Bahçelievler Mah Şehit Fethi Cad. Duygu Sokak No.3 İç Kapı No. 7 Pendik - ISTANBUL/TURKEY"
    row_idx += 1
    ws[f'B{row_idx}'] = "Phn- (+90) 536 763 1911 | Mob- (+90) 541 552 1907"
    row_idx += 1
    ws[f'B{row_idx}'] = "Email- info@inno-mar.com.tr | www.inno-mar.com.tr"
    ws[f'B{row_idx}'].font = Font(color="003399", size=10)
    row_idx += 2
    
    if sablon_tipi == "⚓ INNOMAR Özel Teklif":
        ws[f'B{row_idx}'] = "• MY ADA DRY DOCK SERVICES QUOTATION;"
        ws[f'B{row_idx}'].font = Font(bold=True)
        ws.cell(row=row_idx, column=len(dataframe.columns)+1).value = f"* DATE: {tarih}"
        ws.cell(row=row_idx, column=len(dataframe.columns)+1).font = Font(bold=True)
        row_idx += 2
    else:
        ws[f'C{row_idx}'] = "PROFORMA FATURA"
        ws[f'C{row_idx}'].font = Font(bold=True, size=16)
        row_idx += 2
        ws.cell(row=row_idx, column=len(dataframe.columns)+1).value = f"TARIH: {tarih}"
        ws.cell(row=row_idx, column=len(dataframe.columns)+1).font = Font(bold=True)
        row_idx += 2
    
    headers = ['SIRA' if sablon_tipi != "⚓ INNOMAR Özel Teklif" else 'ITEM NO'] + list(dataframe.columns)
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    
    # Sütun Genişliklerini Ayarla (Optimized)
    ws.column_dimensions['A'].width = 8
    ws.column_dimensions['B'].width = 52 # INSPECTION REMARK/Açıklama sütunu
    ws.column_dimensions['C'].width = 25 # UNIT/Adet sütunu (Genişletildi)
    
    for i in range(3, len(dataframe.columns)+1):
        ws.column_dimensions[chr(65+i)].width = 20
        
    for col_num, header in enumerate(headers, 1):
        cell = ws.cell(row=row_idx, column=col_num)
        cell.value = str(header)
        cell.font = Font(bold=True)
        cell.border = thin_border
    row_idx += 1
    
    for index, row in dataframe.iterrows():
        ws.cell(row=row_idx, column=1).value = index + 1
        ws.cell(row=row_idx, column=1).border = thin_border
        
        for c_idx, col_name in enumerate(dataframe.columns):
            val = row[col_name]
            if col_name == dataframe.columns[-1]: 
                fiyat = pd.to_numeric(val, errors='coerce')
                ws.cell(row=row_idx, column=c_idx+2).value = "-NIL-" if pd.isna(fiyat) or fiyat <= 0 else f"{fiyat:,.0f}".replace(",", ".") + f" {kur_m}"
            else:
                ws.cell(row=row_idx, column=c_idx+2).value = str(val)
                
            ws.cell(row=row_idx, column=c_idx+2).border = thin_border
        row_idx += 1
        
    tot_col = len(dataframe.columns)
    val_col = len(dataframe.columns) + 1
    
    ws.cell(row=row_idx, column=tot_col).value = "ARA TOPLAM" if sablon_tipi != "⚓ INNOMAR Özel Teklif" else "TOTAL PRICE"
    ws.cell(row=row_idx, column=tot_col).border = thin_border
    ws.cell(row=row_idx, column=val_col).value = a_str
    ws.cell(row=row_idx, column=val_col).font = Font(bold=True)
    ws.cell(row=row_idx, column=val_col).border = thin_border
    row_idx += 1
    
    ws.cell(row=row_idx, column=tot_col).value = "KDV (20%)"
    ws.cell(row=row_idx, column=tot_col).border = thin_border
    ws.cell(row=row_idx, column=val_col).value = k_str
    ws.cell(row=row_idx, column=val_col).font = Font(bold=True)
    ws.cell(row=row_idx, column=val_col).border = thin_border
    row_idx += 1
    
    ws.cell(row=row_idx, column=tot_col).value = "GENEL TOPLAM" if sablon_tipi != "⚓ INNOMAR Özel Teklif" else "GRAND TOTAL"
    ws.cell(row=row_idx, column=tot_col).border = thin_border
    ws.cell(row=row_idx, column=val_col).value = g_str
    ws.cell(row=row_idx, column=val_col).font = Font(bold=True)
    ws.cell(row=row_idx, column=val_col).border = thin_border
    row_idx += 2
    
    for satir in notlar.split('\n'):
        ws[f'B{row_idx}'] = satir
        row_idx += 1
        
    # İmza/Alt Bilgi (CE kaldırıldı)
    ws[f'B{row_idx}'] = "İlker TEKINKAYA | Managing Partner | INNOMAR MARİNA YAT LİMAN A.Ş."
    ws[f'B{row_idx}'].font = Font(bold=True)
    row_idx += 1
    ws[f'B{row_idx}'] = "Bahçelievler Mah Şehit Fethi Cad. Duygu Sokak No.3 İç Kapı No. 7 Pendik - ISTANBUL/TURKEY"
    row_idx += 1
    ws[f'B{row_idx}'] = "Phn- (+90) 536 763 1911 | Mob- (+90) 541 552 1907"
    row_idx += 1
    ws[f'B{row_idx}'] = "Email- info@inno-mar.com.tr | www.inno-mar.com.tr"
    ws[f'B{row_idx}'].font = Font(color="003399", size=10)
    
    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()

# --- İNDİRME BUTONLARI ---
st.markdown("### 📥 Çıktı Al")

btn_word, btn_excel, btn_pdf = st.columns(3)

with btn_word:
    st.download_button("📄 WORD İNDİR", data=word_olustur(duzenlenmis_df, ara_str, kdv_str, genel_str, tarih_metni, st.session_state.not_alani, kur_metin, secili_sablon), file_name=f"{secili_sablon.split()[1]}_{dosya_tarihi}.docx", type="primary", use_container_width=True)
with btn_excel:
    st.download_button("📊 EXCEL İNDİR", data=excel_olustur(duzenlenmis_df, ara_str, kdv_str, genel_str, tarih_metni, st.session_state.not_alani, kur_metin, secili_sablon), file_name=f"{secili_sablon.split()[1]}_{dosya_tarihi}.xlsx", type="primary", use_container_width=True)
with btn_pdf:
    st.download_button("📕 PDF İNDİR", data=pdf_olustur(duzenlenmis_df, ara_str, kdv_str, genel_str, tarih_metni, st.session_state.not_alani, kur_metin, secili_sablon), file_name=f"{secili_sablon.split()[1]}_{dosya_tarihi}.pdf", type="primary", use_container_width=True)
