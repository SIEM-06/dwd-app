import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io
import datetime

st.set_page_config(layout="wide", page_title="Innomar Teklif Oluşturucu")

st.title("⚓ INNOMAR | Hızlı Teklif Oluşturucu")
st.write("Aşağıdaki tabloya yapılacak işlemleri girin, program teklif formatını otomatik hazırlasın.")

if 'veri_df' not in st.session_state:
    data = {
        'İşlem (INSPECTION REMARK)': ['ANA MAKİNE BAKIMLARI', 'SU YAPICI BAKIMLARI', 'ZİNCİR GALVANİZ YAPIMI'],
        'Birim (UNIT)': ['2 PIECES', '1 SET', '1 SET'],
        'Fiyat (EURO)': [40000.0, 12000.0, 0.0]
    }
    st.session_state.veri_df = pd.DataFrame(data)

df = st.session_state.veri_df

st.subheader("Teklif Kalemleri")
duzenlenmis_df = st.data_editor(
    df,
    column_config={
        "Fiyat (EURO)": st.column_config.NumberColumn(format="%d €"),
    },
    num_rows="dynamic",
    use_container_width=True
)

ara_toplam = duzenlenmis_df['Fiyat (EURO)'].sum()
kdv = ara_toplam * 0.20
genel_toplam = ara_toplam + kdv

st.write("---")
col1, col2, col3 = st.columns(3)
col1.metric("Ara Toplam", f"{ara_toplam:,.2f} €")
col2.metric("KDV (%20)", f"{kdv:,.2f} €")
col3.metric("Genel Toplam", f"{genel_toplam:,.2f} €")

def teklif_olustur(dataframe, ara_t, kdv_t, genel_t):
    doc = Document()
    
    # Başlık ve Şirket Bilgileri
    header = doc.add_paragraph()
    header.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = header.add_run("INNOMAR MARİNA YAT LİMAN TURİZM İŞLETMECİLİĞİ VE İNŞAAT SANAYİ VE TİCARET A.Ş.\n")
    run.bold = True
    run.font.size = Pt(12)
    header.add_run("Bahçelievler Mah Şehit Fethi Cad. Duygu Sokak No.3 İç Kapı No. 7 Pendik - ISTANBUL/TURKEY\n")
    header.add_run("Phn- (+90) 536 763 1911 | Mob- (+90) 541 552 1907\nEmail- info@inno-mar.com.tr | www.inno-mar.com.tr\n")
    
    doc.add_paragraph(f"DATE: {datetime.date.today().strftime('%d.%m.%Y')}").alignment = WD_ALIGN_PARAGRAPH.RIGHT
    
    doc.add_heading('MY ADA DRY DOCK SERVICES QUOTATION', level=1)
    
    # Tablo Oluşturma
    table = doc.add_table(rows=1, cols=4)
    table.style = 'Table Grid'
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = 'ITEM NO'
    hdr_cells[1].text = 'INSPECTION REMARK'
    hdr_cells[2].text = 'UNIT'
    hdr_cells[3].text = 'PRICE'
    
    for index, row in dataframe.iterrows():
        row_cells = table.add_row().cells
        row_cells[0].text = str(index + 1)
        row_cells[1].text = str(row['İşlem (INSPECTION REMARK)'])
        row_cells[2].text = str(row['Birim (UNIT)'])
        fiyat = row['Fiyat (EURO)']
        row_cells[3].text = f"{fiyat:,.0f} EURO" if fiyat > 0 else "-NIL-"
        
    doc.add_paragraph()
    
    # Toplamlar Tablosu
    tot_table = doc.add_table(rows=3, cols=2)
    tot_table.rows[0].cells[0].text = "TOTAL PRICE"
    tot_table.rows[0].cells[1].text = f"{ara_t:,.2f} EURO"
    tot_table.rows[1].cells[0].text = "VAT (20%)"
    tot_table.rows[1].cells[1].text = f"{kdv_t:,.2f} EURO"
    tot_table.rows[2].cells[0].text = "GRAND TOTAL"
    tot_table.rows[2].cells[1].text = f"{genel_t:,.2f} EURO"
    
    # Alt Notlar
    doc.add_paragraph("\n* IMPORTANT NOTICE;")
    doc.add_paragraph("- DURING MAINTENANCE IF DEFORMATION DETECTED ON WORKING SURFACE AND NEEDED TO RENEW COMPONENTS EACH PARTS WILL BE PRICED ADDITIONALLY.")
    doc.add_paragraph("\n* REMARKS;")
    doc.add_paragraph("- DELIVERY TIME FOR THE JOB IS 35 DAYS,\n- A DETAILED REPORT WILL BE SUBMITTED TO YOUR SIDE UPON COMPLETION OF THE WORK,\n- PAYMENT WILL BE ACCEPTED AS BELOW;\n  - %50 BEFORE WORK BEGINS,\n  - %50 UPON COMPLETION OF THE WORK.")
    
    doc.add_paragraph("\nİlker TEKINKAYA | Managing Partner\nINNOMAR MARİNA YAT LİMAN A.Ş.").alignment = WD_ALIGN_PARAGRAPH.RIGHT

    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()

st.write("---")
word_dosyasi = teklif_olustur(duzenlenmis_df, ara_toplam, kdv, genel_toplam)

st.download_button(
    label="📄 TEKLİFİ İNDİR (WORD)",
    data=word_dosyasi,
    file_name=f"Innomar_Teklif_{datetime.date.today().strftime('%d_%m_%Y')}.docx",
    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    type="primary"
)
