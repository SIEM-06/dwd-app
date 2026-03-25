import streamlit as st
import pandas as pd
import io
import datetime
import os
from fpdf import FPDF
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.drawing.image import Image as xlImage

st.set_page_config(layout="wide", page_title="Innomar Teklif Portali", initial_sidebar_state="collapsed")

st.markdown("<h2 style='text-align: center;'>⚓ INNOMAR TEKLİF SİSTEMİ</h2>", unsafe_allow_html=True)

secilen_tarih = st.date_input("Teklif Tarihi Belirle", datetime.date.today())
tarih_metni = secilen_tarih.strftime("%d.%m.%Y")
dosya_tarihi = secilen_tarih.strftime("%d_%m_%Y")

st.info("Telefondan veri girerken tablodaki hücrelerin üzerine tıklayıp değiştirebilirsiniz. Yeni satır için tablonun en altını kullanın.")

if 'veri_df' not in st.session_state:
    data = {
        'İşlem (INSPECTION REMARK)': ['ANA MAKİNE BAKIMLARI', 'SU YAPICI BAKIMLARI', 'ZİNCİR GALVANİZ YAPIMI'],
        'Birim': ['2 PIECES', '1 SET', '1 SET'],
        'Fiyat (€)': [40000.0, 12000.0, 0.0]
    }
    st.session_state.veri_df = pd.DataFrame(data)

df = st.session_state.veri_df

duzenlenmis_df = st.data_editor(
    df,
    column_config={
        "Fiyat (€)": st.column_config.NumberColumn(format="%d €"),
    },
    num_rows="dynamic",
    use_container_width=True 
)

ara_toplam = duzenlenmis_df['Fiyat (€)'].sum()
kdv = ara_toplam * 0.20
genel_toplam = ara_toplam + kdv

st.write("---")
col_a, col_b, col_c = st.columns(3)
col_a.metric("Ara Toplam", f"{ara_toplam:,.0f} €")
col_b.metric("KDV (%20)", f"{kdv:,.0f} €")
col_c.metric("Genel Toplam", f"{genel_toplam:,.0f} €")
st.write("---")

# ==========================================
# WORD OLUŞTURMA MOTORU
# ==========================================
def word_olustur(dataframe, ara_t, kdv_t, genel_t, tarih):
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
    
    run_addr = p_info.add_run("Bahçelievler Mah Şehit Fethi Cad. Duygu Sokak No.3 İç Kapı No. 7\nPendik - ISTANBUL/TURKEY\nPhn- (+90) 536 763 1911 | Mob- (+90) 541 552 1907\n")
    run_addr.font.size = Pt(9)
    
    run_mail = p_info.add_run("Email- info@inno-mar.com.tr | www.inno-mar.com.tr")
    run_mail.font.color.rgb = RGBColor(0, 51, 153)
    run_mail.font.size = Pt(9)
    
    doc.add_paragraph("_" * 75)
    
    p_title = doc.add_paragraph()
    run_title = p_title.add_run(f"•   MY ADA DRY DOCK SERVICES QUOTATION;                                 * DATE: {tarih}")
    run_title.bold = True
    run_title.font.size = Pt(10)
    
    table = doc.add_table(rows=1, cols=4)
    table.style = 'Table Grid'
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text, hdr_cells[1].text, hdr_cells[2].text, hdr_cells[3].text = 'ITEM NO', 'INSPECTION REMARK', 'UNIT', 'PRICE'
    for cell in hdr_cells:
        for paragraph in cell.paragraphs:
            for run in paragraph.runs:
                run.font.bold = True
    
    for index, row in dataframe.iterrows():
        row_cells = table.add_row().cells
        row_cells[0].text = str(index + 1)
        row_cells[1].text = str(row['İşlem (INSPECTION REMARK)'])
        row_cells[2].text = str(row['Birim'])
        fiyat = row['Fiyat (€)']
        row_cells[3].text = f"{fiyat:,.0f} EURO" if fiyat > 0 else "-NIL-"
        
    doc.add_paragraph()
    
    tot_table = doc.add_table(rows=3, cols=2)
    tot_table.style = 'Table Grid'
    tot_table.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    tot_table.rows[0].cells[0].text, tot_table.rows[0].cells[1].text = "TOTAL PRICE", f"{ara_t:,.0f} EURO"
    tot_table.rows[1].cells[0].text, tot_table.rows[1].cells[1].text = "VAT (20%)", f"{kdv_t:,.0f} EURO"
    tot_table.rows[2].cells[0].text, tot_table.rows[2].cells[1].text = "GRAND TOTAL", f"{genel_t:,.0f} EURO"
    
    for row in tot_table.rows:
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    run.font.bold = True
                    
    doc.add_paragraph("\n* IMPORTANT NOTICE;").runs[0].bold = True
    doc.add_paragraph("- DURING MAINTENANCE IF DEFORMATION DETECTED ON WORKING SURFACE AND NEEDED TO RENEW\nCOMPONENTS EACH PARTS WILL BE PRICED ADDITIONALLY.")
    
    doc.add_paragraph("\n* REMARKS;").runs[0].bold = True
    doc.add_paragraph("- DELIVERY TIME FOR THE JOB IS 35 DAYS,\n- A DETAILED REPORT WILL BE SUBMITTED TO YOUR SIDE UPON COMPLETION OF THE WORK,\n- PAYMENT WILL BE ACCEPTED AS BELOW;\n    - %50 BEFORE WORK BEGINS,\n    - %50 UPON COMPLETION OF THE WORK.")
    
    doc.add_paragraph("\nCE İlker TEKINKAYA | Managing Partner | INNOMAR MARİNA YAT\nLİMAN TURİZM İŞLETMECİLİĞİ VE İNŞAAT SANAYİ VE TİCARET A.Ş.").runs[0].bold = True
    doc.add_paragraph("Bahçelievler Mah Şehit Fethi Cad. Duygu Sokak No.3 İç Kapı No. 7\nPendik - ISTANBUL/TURKEY\nPhn- (+90) 536 763 1911 | Mob- (+90) 541 552 1907\nEmail- info@inno-mar.com.tr | www.inno-mar.com.tr")
    
    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()

# ==========================================
# EXCEL OLUŞTURMA MOTORU
# ==========================================
def excel_olustur(dataframe, ara_t, kdv_t, genel_t, tarih):
    wb = Workbook()
    ws = wb.active
    ws.title = "Innomar Teklif"
    
    ws.column_dimensions['A'].width = 8
    ws.column_dimensions['B'].width = 55
    ws.column_dimensions['C'].width = 15
    ws.column_dimensions['D'].width = 20
    
    row_idx = 1
    
    if os.path.exists("logo.png"):
        img = xlImage("logo.png")
        ws.add_image(img, 'B1')
        row_idx = 8
        
    blue_font_bold = Font(color="003399", bold=True, size=11)
    black_font = Font(color="000000", size=10)
    blue_font = Font(color="003399", size=10)
    
    ws[f'B{row_idx}'] = "INNOMAR MARİNA YAT LİMAN TURİZM İŞLETMECİLİĞİ VE İNŞAAT SANAYİ VE TİCARET A.Ş."
    ws[f'B{row_idx}'].font = blue_font_bold
    row_idx += 1
    ws[f'B{row_idx}'] = "Bahçelievler Mah Şehit Fethi Cad. Duygu Sokak No.3 İç Kapı No. 7 Pendik - ISTANBUL/TURKEY"
    ws[f'B{row_idx}'].font = black_font
    row_idx += 1
    ws[f'B{row_idx}'] = "Phn- (+90) 536 763 1911 | Mob- (+90) 541 552 1907"
    ws[f'B{row_idx}'].font = black_font
    row_idx += 1
    ws[f'B{row_idx}'] = "Email- info@inno-mar.com.tr | www.inno-mar.com.tr"
    ws[f'B{row_idx}'].font = blue_font
    row_idx += 2
    
    ws[f'B{row_idx}'] = "• MY ADA DRY DOCK SERVICES QUOTATION;"
    ws[f'B{row_idx}'].font = Font(bold=True)
    ws[f'D{row_idx}'] = f"* DATE: {tarih}"
    ws[f'D{row_idx}'].font = Font(bold=True)
    row_idx += 2
    
    headers = ['ITEM NO', 'INSPECTION REMARK', 'UNIT', 'PRICE']
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    
    for col_num, header in enumerate(headers, 1):
        cell = ws.cell(row=row_idx, column=col_num)
        cell.value = header
        cell.font = Font(bold=True)
        cell.border = thin_border
    row_idx += 1
    
    for index, row in dataframe.iterrows():
        ws.cell(row=row_idx, column=1).value = index + 1
        ws.cell(row=row_idx, column=2).value = str(row['İşlem (INSPECTION REMARK)'])
        ws.cell(row=row_idx, column=3).value = str(row['Birim'])
        fiyat = row['Fiyat (€)']
        ws.cell(row=row_idx, column=4).value = f"{fiyat:,.0f} EURO" if fiyat > 0 else "-NIL-"
        
        for i in range(1, 5):
            ws.cell(row=row_idx, column=i).border = thin_border
        row_idx += 1
        
    ws.cell(row=row_idx, column=3).value = "TOTAL PRICE"
    ws.cell(row=row_idx, column=3).font = Font(bold=True)
    ws.cell(row=row_idx, column=3).border = thin_border
    ws.cell(row=row_idx, column=4).value = f"{ara_t:,.0f} EURO"
    ws.cell(row=row_idx, column=4).border = thin_border
    row_idx += 1
    
    ws.cell(row=row_idx, column=3).value = "VAT (20%)"
    ws.cell(row=row_idx, column=3).font = Font(bold=True)
    ws.cell(row=row_idx, column=3).border = thin_border
    ws.cell(row=row_idx, column=4).value = f"{kdv_t:,.0f} EURO"
    ws.cell(row=row_idx, column=4).border = thin_border
    row_idx += 1
    
    ws.cell(row=row_idx, column=3).value = "GRAND TOTAL"
    ws.cell(row=row_idx, column=3).font = Font(bold=True)
    ws.cell(row=row_idx, column=3).border = thin_border
    ws.cell(row=row_idx, column=4).value = f"{genel_t:,.0f} EURO"
    ws.cell(row=row_idx, column=4).border = thin_border
    row_idx += 2
    
    ws[f'B{row_idx}'] = "* IMPORTANT NOTICE;"
    ws[f'B{row_idx}'].font = Font(bold=True)
    row_idx += 1
    ws[f'B{row_idx}'] = "- DURING MAINTENANCE IF DEFORMATION DETECTED ON WORKING SURFACE AND NEEDED TO RENEW COMPONENTS EACH PARTS WILL BE PRICED ADDITIONALLY."
    row_idx += 2
    
    ws[f'B{row_idx}'] = "* REMARKS;"
    ws[f'B{row_idx}'].font = Font(bold=True)
    row_idx += 1
    ws[f'B{row_idx}'] = "- DELIVERY TIME FOR THE JOB IS 35 DAYS,"
    row_idx += 1
    ws[f'B{row_idx}'] = "- A DETAILED REPORT WILL BE SUBMITTED TO YOUR SIDE UPON COMPLETION OF THE WORK,"
    row_idx += 1
    ws[f'B{row_idx}'] = "- PAYMENT WILL BE ACCEPTED AS BELOW;"
    row_idx += 1
    ws[f'B{row_idx}'] = "    - %50 BEFORE WORK BEGINS,"
    row_idx += 1
    ws[f'B{row_idx}'] = "    - %50 UPON COMPLETION OF THE WORK."
    
    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()

# ==========================================
# PDF OLUŞTURMA MOTORU
# ==========================================
def cevir_tr(metin):
    tr_map = {'ş':'s', 'Ş':'S', 'ı':'i', 'İ':'I', 'ğ':'g', 'Ğ':'G', 'ü':'u', 'Ü':'U', 'ö':'o', 'Ö':'O', 'ç':'c', 'Ç':'C'}
    for k, v in tr_map.items(): metin = metin.replace(k, v)
    return metin

def pdf_olustur(dataframe, ara_t, kdv_t, genel_t, tarih):
    class PDF(FPDF):
        def header(self):
            # Sadece 1. sayfada antetli kağıt başlığını çiz
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
                
                # Filigranı da sadece ilk sayfaya atıyoruz
                if os.path.exists("watermark.png"):
                    self.image("watermark.png", x=30, y=80, w=150)
            else:
                # 2. ve sonraki sayfalarda temiz bir üst boşluk bırakıp devam et
                self.ln(15)

    pdf = PDF()
    pdf.add_page()
    pdf.set_auto_page_break(auto=True, margin=15)
    
    pdf.set_font('Arial', 'B', 10)
    pdf.set_text_color(0, 0, 0)
    pdf.cell(130, 10, chr(149) + '   MY ADA DRY DOCK SERVICES QUOTATION;', 0, 0, 'L')
    pdf.cell(60, 10, f'* DATE: {tarih}', 0, 1, 'R')
    pdf.ln(2)
    
    pdf.set_draw_color(0, 0, 0)
    pdf.set_font('Arial', 'B', 9)
    pdf.cell(15, 8, 'ITEM NO', 1)
    pdf.cell(100, 8, 'INSPECTION REMARK', 1)
    pdf.cell(30, 8, 'UNIT', 1)
    pdf.cell(45, 8, 'PRICE', 1)
    pdf.ln()
    
    pdf.set_font('Arial', '', 8)
    for index, row in dataframe.iterrows():
        pdf.cell(15, 8, str(index + 1), 1)
        pdf.cell(100, 8, cevir_tr(str(row['İşlem (INSPECTION REMARK)'])), 1)
        pdf.cell(30, 8, cevir_tr(str(row['Birim'])), 1)
        fiyat = row['Fiyat (€)']
        pdf.cell(45, 8, f"{fiyat:,.0f} EURO" if fiyat > 0 else "-NIL-", 1)
        pdf.ln()
        
    pdf.set_font('Arial', 'B', 9)
    pdf.cell(115, 8, '', 0, 0)
    pdf.cell(30, 8, 'TOTAL PRICE', 1, 0, 'L')
    pdf.cell(45, 8, f"{ara_t:,.0f} EURO", 1, 1, 'L')
    
    pdf.cell(115, 8, '', 0, 0)
    pdf.cell(30, 8, 'VAT (20%)', 1, 0, 'L')
    pdf.cell(45, 8, f"{kdv_t:,.0f} EURO", 1, 1, 'L')
    
    pdf.cell(115, 8, '', 0, 0)
    pdf.cell(30, 8, 'GRAND TOTAL', 1, 0, 'L')
    pdf.cell(45, 8, f"{genel_t:,.0f} EURO", 1, 1, 'L')
    
    pdf.ln(10)
    
    pdf.set_font('Arial', 'B', 8)
    pdf.cell(0, 5, '* IMPORTANT NOTICE;', 0, 1, 'L')
    pdf.set_font('Arial', '', 8)
    pdf.cell(0, 5, '- DURING MAINTENANCE IF DEFORMATION DETECTED ON WORKING SURFACE AND NEEDED TO RENEW', 0, 1, 'L')
    pdf.set_font('Arial', 'B', 8)
    pdf.cell(0, 5, 'COMPONENTS EACH PARTS WILL BE PRICED ADDITIONALLY.', 0, 1, 'L')
    pdf.ln(4)
    
    pdf.set_font('Arial', 'B', 8)
    pdf.cell(0, 5, '* REMARKS;', 0, 1, 'L')
    pdf.set_font('Arial', '', 8)
    pdf.cell(0, 5, '- DELIVERY TIME FOR THE JOB IS 35 DAYS,', 0, 1, 'L')
    pdf.cell(0, 5, '- A DETAILED REPORT WILL BE SUBMITTED TO YOUR SIDE UPON COMPLETION OF THE WORK,', 0, 1, 'L')
    pdf.cell(0, 5, '- PAYMENT WILL BE ACCEPTED AS BELOW;', 0, 1, 'L')
    pdf.cell(10, 5, '', 0, 0)
    pdf.cell(0, 5, '- %50 BEFORE WORK BEGINS,', 0, 1, 'L')
    pdf.cell(10, 5, '', 0, 0)
    pdf.cell(0, 5, '- %50 UPON COMPLETION OF THE WORK.', 0, 1, 'L')
    
    pdf.ln(10)
    
    pdf.set_font('Arial', 'B', 8)
    pdf.cell(0, 4, cevir_tr('CE Ilker TEKINKAYA | Managing Partner | INNOMAR MARINA YAT'), 0, 1, 'L')
    pdf.cell(0, 4, cevir_tr('LIMAN TURIZM ISLETMECILIGI VE INSAAT SANAYI VE TICARET A.S.'), 0, 1, 'L')
    pdf.set_font('Arial', '', 8)
    pdf.cell(0, 4, cevir_tr('Bahcelievler Mah Sehit Fethi Cad. Duygu Sokak No.3 Ic Kapi No. 7'), 0, 1, 'L')
    pdf.cell(0, 4, 'Pendik - ISTANBUL/TURKEY', 0, 1, 'L')
    pdf.cell(0, 4, 'Phn- (+90) 536 763 1911 | Mob- (+90) 541 552 1907', 0, 1, 'L')
    
    pdf.set_text_color(0, 51, 153)
    pdf.cell(0, 4, 'Email- info@inno-mar.com.tr | www.inno-mar.com.tr', 0, 1, 'L')
    
    return pdf.output(dest='S').encode('latin-1')

# --- İNDİRME BUTONLARI ---
st.markdown("### 📥 Çıktı Al")

btn_word, btn_excel, btn_pdf = st.columns(3)

with btn_word:
    st.download_button("📄 WORD İNDİR", data=word_olustur(duzenlenmis_df, ara_toplam, kdv, genel_toplam, tarih_metni), file_name=f"Teklif_{dosya_tarihi}.docx", type="primary", use_container_width=True)
with btn_excel:
    st.download_button("📊 EXCEL İNDİR", data=excel_olustur(duzenlenmis_df, ara_toplam, kdv, genel_toplam, tarih_metni), file_name=f"Teklif_{dosya_tarihi}.xlsx", type="primary", use_container_width=True)
with btn_pdf:
    st.download_button("📕 PDF İNDİR", data=pdf_olustur(duzenlenmis_df, ara_toplam, kdv, genel_toplam, tarih_metni), file_name=f"Teklif_{dosya_tarihi}.pdf", type="primary", use_container_width=True)
