import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io
import datetime
import os
from fpdf import FPDF

st.set_page_config(layout="wide", page_title="Innomar Teklif Portali", initial_sidebar_state="collapsed")

st.markdown("<h2 style='text-align: center;'>⚓ INNOMAR TEKLİF SİSTEMİ</h2>", unsafe_allow_html=True)
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

def word_olustur(dataframe, ara_t, kdv_t, genel_t):
    doc = Document()
    
    if os.path.exists("logo.png"):
        pic_para = doc.add_paragraph()
        pic_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run_pic = pic_para.add_run()
        run_pic.add_picture("logo.png", width=Cm(6))
        
    header = doc.add_paragraph()
    header.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = header.add_run("INNOMAR MARİNA YAT LİMAN TURİZM İŞLETMECİLİĞİ VE İNŞAAT SANAYİ VE TİCARET A.Ş.\n")
    run.bold = True
    header.add_run("Pendik - ISTANBUL/TURKEY | info@inno-mar.com.tr\n")
    
    doc.add_paragraph(f"DATE: {datetime.date.today().strftime('%d.%m.%Y')}").alignment = WD_ALIGN_PARAGRAPH.RIGHT
    doc.add_heading('MY ADA DRY DOCK SERVICES QUOTATION', level=1)
    
    table = doc.add_table(rows=1, cols=4)
    table.style = 'Table Grid'
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text, hdr_cells[1].text, hdr_cells[2].text, hdr_cells[3].text = 'NO', 'INSPECTION REMARK', 'UNIT', 'PRICE'
    
    for index, row in dataframe.iterrows():
        row_cells = table.add_row().cells
        row_cells[0].text = str(index + 1)
        row_cells[1].text = str(row['İşlem (INSPECTION REMARK)'])
        row_cells[2].text = str(row['Birim'])
        fiyat = row['Fiyat (€)']
        row_cells[3].text = f"{fiyat:,.0f} EURO" if fiyat > 0 else "-NIL-"
        
    doc.add_paragraph("\nTOTALS:")
    doc.add_paragraph(f"TOTAL PRICE: {ara_t:,.0f} EURO\nVAT (20%): {kdv_t:,.0f} EURO\nGRAND TOTAL: {genel_t:,.0f} EURO")
    
    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()

def excel_olustur(dataframe, ara_t, kdv_t, genel_t):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        dataframe.to_excel(writer, index=False, sheet_name='Teklif_Listesi')
        toplamlar = pd.DataFrame({
            'İşlem (INSPECTION REMARK)': ['ARA TOPLAM', 'KDV (%20)', 'GENEL TOPLAM'],
            'Birim': ['', '', ''],
            'Fiyat (€)': [ara_t, kdv_t, genel_t]
        })
        toplamlar.to_excel(writer, index=False, header=False, startrow=len(dataframe)+2, sheet_name='Teklif_Listesi')
    return output.getvalue()

# ==========================================
# 3. PDF OLUŞTURMA MOTORU (TIPATIP ŞABLON)
# ==========================================
def cevir_tr(metin):
    tr_map = {'ş':'s', 'Ş':'S', 'ı':'i', 'İ':'I', 'ğ':'g', 'Ğ':'G', 'ü':'u', 'Ü':'U', 'ö':'o', 'Ö':'O', 'ç':'c', 'Ç':'C'}
    for k, v in tr_map.items(): metin = metin.replace(k, v)
    return metin

def pdf_olustur(dataframe, ara_t, kdv_t, genel_t):
    from fpdf import FPDF
    import os
    import datetime

    class PDF(FPDF):
        def header(self):
            # 1. LOGO KISMI
            if os.path.exists("logo.png"):
                self.image("logo.png", x=65, y=10, w=80)
            self.ln(25)
            
            # 2. ŞİRKET BİLGİLERİ (Mavi ve Siyah Tonlar)
            self.set_font('Arial', 'B', 10)
            self.set_text_color(0, 51, 153) # Kurumsal Lacivert/Mavi
            self.cell(0, 5, cevir_tr('INNOMAR MARİNA YAT'), 0, 1, 'L')
            self.cell(0, 5, cevir_tr('LİMAN TURİZM İŞLETMECİLİĞİ VE İNŞAAT SANAYİ VE TİCARET A.Ş.'), 0, 1, 'L')
            
            self.set_font('Arial', '', 9)
            self.set_text_color(0, 0, 0) # Siyah
            self.cell(0, 5, cevir_tr('Bahçelievler Mah Şehit Fethi Cad. Duygu Sokak No.3 İç Kapı No. 7'), 0, 1, 'L')
            self.cell(0, 5, 'Pendik - ISTANBUL/TURKEY', 0, 1, 'L')
            self.cell(0, 5, 'Phn- (+90) 536 763 1911 | Mob- (+90) 541 552 1907', 0, 1, 'L')
            
            self.set_text_color(0, 51, 153) # Tekrar Mavi
            self.cell(0, 5, 'Email- info@inno-mar.com.tr | www.inno-mar.com.tr', 0, 1, 'L')
            
            # 3. İNCE MAVİ ÇİZGİ
            self.set_draw_color(0, 51, 153)
            self.set_line_width(0.3)
            self.line(10, self.get_y()+2, 200, self.get_y()+2)
            self.ln(10)

    pdf = PDF()
    pdf.add_page()
    
    # 4. ARKA PLAN FİLİGRANI (Opsiyonel)
    if os.path.exists("watermark.png"):
        pdf.image("watermark.png", x=30, y=80, w=150)
        
    # 5. BAŞLIK VE TARİH
    pdf.set_font('Arial', 'B', 10)
    pdf.set_text_color(0, 0, 0)
    # Aynı satırda hem sol başlık hem sağ tarih
    pdf.cell(130, 10, chr(149) + '   MY ADA DRY DOCK SERVICES QUOTATION;', 0, 0, 'L')
    pdf.cell(60, 10, f'* DATE: {datetime.date.today().strftime("%d.%m.%Y")}', 0, 1, 'R')
    pdf.ln(2)
    
    # 6. TABLO BAŞLIKLARI
    pdf.set_draw_color(0, 0, 0)
    pdf.set_font('Arial', 'B', 9)
    pdf.cell(15, 8, 'ITEM NO', 1)
    pdf.cell(100, 8, 'INSPECTION REMARK', 1)
    pdf.cell(30, 8, 'UNIT', 1)
    pdf.cell(45, 8, 'PRICE', 1)
    pdf.ln()
    
    # 7. TABLO İÇERİĞİ (Dinamik Veriler)
    pdf.set_font('Arial', '', 8)
    for index, row in dataframe.iterrows():
        pdf.cell(15, 8, str(index + 1), 1)
        pdf.cell(100, 8, cevir_tr(str(row['İşlem (INSPECTION REMARK)'])), 1)
        pdf.cell(30, 8, cevir_tr(str(row['Birim'])), 1)
        fiyat = row['Fiyat (€)']
        pdf.cell(45, 8, f"{fiyat:,.0f} EURO" if fiyat > 0 else "-NIL-", 1)
        pdf.ln()
        
    # 8. TOPLAMLAR TABLOSU (Sağ Alt Köşe)
    pdf.set_font('Arial', 'B', 9)
    pdf.cell(115, 8, '', 0, 0) # Sola boşluk bırakarak sağa yaslama
    pdf.cell(30, 8, 'TOTAL PRICE', 1, 0, 'L')
    pdf.cell(45, 8, f"{ara_t:,.0f} EURO", 1, 1, 'L')
    
    pdf.cell(115, 8, '', 0, 0)
    pdf.cell(30, 8, 'VAT (20%)', 1, 0, 'L')
    pdf.cell(45, 8, f"{kdv_t:,.0f} EURO", 1, 1, 'L')
    
    pdf.cell(115, 8, '', 0, 0)
    pdf.cell(30, 8, 'GRAND TOTAL', 1, 0, 'L')
    pdf.cell(45, 8, f"{genel_t:,.0f} EURO", 1, 1, 'L')
    
    pdf.ln(15)
    
    # 9. ALT NOTLAR (IMPORTANT NOTICE & REMARKS)
    pdf.set_font('Arial', 'B', 9)
    pdf.cell(0, 5, '* IMPORTANT NOTICE;', 0, 1, 'L')
    pdf.set_font('Arial', '', 9)
    pdf.cell(0, 5, '- DURING MAINTENANCE IF DEFORMATION DETECTED ON WORKING SURFACE AND NEEDED TO RENEW', 0, 1, 'L')
    pdf.set_font('Arial', 'B', 9)
    pdf.cell(0, 5, 'COMPONENTS EACH PARTS WILL BE PRICED ADDITIONALLY.', 0, 1, 'L')
    pdf.ln(5)
    
    pdf.set_font('Arial', 'B', 9)
    pdf.cell(0, 5, '* REMARKS;', 0, 1, 'L')
    pdf.set_font('Arial', '', 9)
    pdf.cell(0, 5, '- DELIVERY TIME FOR THE JOB IS 35 DAYS,', 0, 1, 'L')
    pdf.cell(0, 5, '- A DETAILED REPORT WILL BE SUBMITTED TO YOUR SIDE UPON COMPLETION OF THE WORK,', 0, 1, 'L')
    pdf.cell(0, 5, '- PAYMENT WILL BE ACCEPTED AS BELOW;', 0, 1, 'L')
    pdf.cell(10, 5, '', 0, 0) # Girinti
    pdf.cell(0, 5, '- %50 BEFORE WORK BEGINS,', 0, 1, 'L')
    pdf.cell(10, 5, '', 0, 0) # Girinti
    pdf.cell(0, 5, '- %50 UPON COMPLETION OF THE WORK.', 0, 1, 'L')
    
    pdf.ln(15)
    
    # 10. İMZA BLOĞU
    pdf.set_font('Arial', 'B', 9)
    pdf.cell(0, 5, cevir_tr('CE Ilker TEKINKAYA | Managing Partner | INNOMAR MARINA YAT'), 0, 1, 'L')
    pdf.cell(0, 5, cevir_tr('LIMAN TURIZM ISLETMECILIGI VE INSAAT SANAYI VE TICARET A.S.'), 0, 1, 'L')
    pdf.set_font('Arial', '', 8)
    pdf.cell(0, 5, cevir_tr('Bahcelievler Mah Sehit Fethi Cad. Duygu Sokak No.3 Ic Kapi No. 7'), 0, 1, 'L')
    pdf.cell(0, 5, 'Pendik - ISTANBUL/TURKEY', 0, 1, 'L')
    pdf.cell(0, 5, 'Phn- (+90) 536 763 1911 | Mob- (+90) 541 552 1907', 0, 1, 'L')
    
    return pdf.output(dest='S').encode('latin-1')

st.markdown("### 📥 Çıktı Al")

btn_word, btn_excel, btn_pdf = st.columns(3)
tarih_str = datetime.date.today().strftime('%d_%m_%Y')

with btn_word:
    st.download_button("📄 WORD İNDİR", data=word_olustur(duzenlenmis_df, ara_toplam, kdv, genel_toplam), file_name=f"Teklif_{tarih_str}.docx", type="primary", use_container_width=True)
with btn_excel:
    st.download_button("📊 EXCEL İNDİR", data=excel_olustur(duzenlenmis_df, ara_toplam, kdv, genel_toplam), file_name=f"Teklif_{tarih_str}.xlsx", type="primary", use_container_width=True)
with btn_pdf:
    st.download_button("📕 PDF İNDİR", data=pdf_olustur(duzenlenmis_df, ara_toplam, kdv, genel_toplam), file_name=f"Teklif_{tarih_str}.pdf", type="primary", use_container_width=True)
