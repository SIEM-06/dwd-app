import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io
import datetime
from fpdf import FPDF

# --- MOBİL UYUMLU SAYFA AYARLARI ---
st.set_page_config(layout="wide", page_title="Innomar Teklif Portali", initial_sidebar_state="collapsed")

st.markdown("<h2 style='text-align: center;'>⚓ INNOMAR TEKLİF SİSTEMİ</h2>", unsafe_allow_html=True)
st.info("📱 Telefondan veri girerken tablodaki hücrelerin üzerine tıklayıp değiştirebilirsiniz. Yeni satır için tablonun en altını kullanın.")

# --- VERİ SETİ ---
if 'veri_df' not in st.session_state:
    data = {
        'İşlem (INSPECTION REMARK)': ['ANA MAKİNE BAKIMLARI', 'SU YAPICI BAKIMLARI', 'ZİNCİR GALVANİZ YAPIMI'],
        'Birim': ['2 PIECES', '1 SET', '1 SET'],
        'Fiyat (€)': [40000.0, 12000.0, 0.0]
    }
    st.session_state.veri_df = pd.DataFrame(data)

df = st.session_state.veri_df

# --- MOBİL UYUMLU İNTERAKTİF TABLO ---
duzenlenmis_df = st.data_editor(
    df,
    column_config={
        "Fiyat (€)": st.column_config.NumberColumn(format="%d €"),
    },
    num_rows="dynamic",
    use_container_width=True # Telefonda ekrana tam oturmasını sağlar
)

# --- HESAPLAMALAR ---
ara_toplam = duzenlenmis_df['Fiyat (€)'].sum()
kdv = ara_toplam * 0.20
genel_toplam = ara_toplam + kdv

st.write("---")
# Telefonda bu 3 kutu alt alta çok şık durur
col1, col2, col3 = st.columns(3)
col1.metric("Ara Toplam", f"{ara_toplam:,.0f} €")
col2.metric("KDV (%20)", f"{kdv:,.0f} €")
col3.metric("Genel Toplam", f"{genel_toplam:,.0f} €")
st.write("---")

# ==========================================
# 1. WORD OLUŞTURMA MOTORU
# ==========================================
def word_olustur(dataframe, ara_t, kdv_t, genel_t):
    doc = Document()
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

# ==========================================
# 2. EXCEL OLUŞTURMA MOTORU
# ==========================================
def excel_olustur(dataframe, ara_t, kdv_t, genel_t):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        dataframe.to_excel(writer, index=False, sheet_name='Teklif_Listesi')
        # Alt kısma toplamları ekleme
        toplamlar = pd.DataFrame({
            'İşlem (INSPECTION REMARK)': ['ARA TOPLAM', 'KDV (%20)', 'GENEL TOPLAM'],
            'Birim': ['', '', ''],
            'Fiyat (€)': [ara_t, kdv_t, genel_t]
        })
        toplamlar.to_excel(writer, index=False, header=False, startrow=len(dataframe)+2, sheet_name='Teklif_Listesi')
    return output.getvalue()

# ==========================================
# 3. PDF OLUŞTURMA MOTORU
# ==========================================
# PDF'te telefonlardan dolayı çökme olmaması için Türkçe karakterleri İngilizceye çeviriyoruz
def cevir_tr(metin):
    tr_map = {'ş':'s', 'Ş':'S', 'ı':'i', 'İ':'I', 'ğ':'g', 'Ğ':'G', 'ü':'u', 'Ü':'U', 'ö':'o', 'Ö':'O', 'ç':'c', 'Ç':'C'}
    for k, v in tr_map.items(): metin = metin.replace(k, v)
    return metin

def pdf_olustur(dataframe, ara_t, kdv_t, genel_t):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font('Arial', 'B', 12)
    pdf.cell(0, 8, 'INNOMAR MARINA YAT LIMAN A.S.', 0, 1, 'C')
    pdf.set_font('Arial', '', 9)
    pdf.cell(0, 5, 'Pendik - ISTANBUL/TURKEY', 0, 1, 'C')
    pdf.ln(5)
    
    pdf.set_font('Arial', 'B', 12)
    pdf.cell(0, 10, 'MY ADA DRY DOCK SERVICES QUOTATION', 0, 1, 'C')
    pdf.set_font('Arial', '', 10)
    pdf.cell(0, 10, f'DATE: {datetime.date.today().strftime("%d.%m.%Y")}', 0, 1, 'R')
    
    pdf.set_font('Arial', 'B', 9)
    pdf.cell(10, 8, 'NO', 1); pdf.cell(115, 8, 'INSPECTION REMARK', 1); pdf.cell(25, 8, 'UNIT', 1); pdf.cell(35, 8, 'PRICE', 1)
    pdf.ln()
    
    pdf.set_font('Arial', '', 8)
    for index, row in dataframe.iterrows():
        pdf.cell(10, 8, str(index + 1), 1)
        pdf.cell(115, 8, cevir_tr(str(row['İşlem (INSPECTION REMARK)'])), 1)
        pdf.cell(25, 8, cevir_tr(str(row['Birim'])), 1)
        fiyat = row['Fiyat (€)']
        pdf.cell(35, 8, f"{fiyat:,.0f} EURO" if fiyat > 0 else "-NIL-", 1)
        pdf.ln()
        
    pdf.ln(5)
    pdf.set_font('Arial', 'B', 10)
    pdf.cell(150, 8, 'TOTAL PRICE:', 0, 0, 'R'); pdf.cell(35, 8, f"{ara_t:,.0f} EURO", 1, 1, 'R')
    pdf.cell(150, 8, 'VAT (20%):', 0, 0, 'R'); pdf.cell(35, 8, f"{kdv_t:,.0f} EURO", 1, 1, 'R')
    pdf.cell(150, 8, 'GRAND TOTAL:', 0, 0, 'R'); pdf.cell(35, 8, f"{genel_t:,.0f} EURO", 1, 1, 'R')
    
    return pdf.output(dest='S').encode('latin-1')

# --- MOBİL UYUMLU İNDİRME BUTONLARI ---
st.markdown("### 📥 Çıktı Al")
st.caption("Aşağıdaki butonları kullanarak teklifi istediğiniz formatta cihazınıza indirebilirsiniz.")

# Butonları yan yana dizer
btn1, btn2, btn3 = st.columns(3)

tarih_str = datetime.date.today().strftime('%d_%m_%Y')

with btn1:
    st.download_button("📄 WORD İNDİR", data=word_olustur(duzenlenmis_df, ara_toplam, kdv, genel_toplam), file_name=f"Teklif_{tarih_str}.docx", type="primary", use_container_width=True)
with btn2:
    st.download_button("📊 EXCEL İNDİR", data=excel_olustur(duzenlenmis_df, ara_toplam, kdv, genel_toplam), file_name=f"Teklif_{tarih_str}.xlsx", type="primary", use_container_width=True)
with btn3:
    st.download_button("📕 PDF İNDİR", data=pdf_olustur(duzenlenmis_df, ara_toplam, kdv, genel_toplam), file_name=f"Teklif_{tarih_str}.pdf", type="primary", use_container_width=True)
