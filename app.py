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

st.set_page_config(layout="wide", page_title="INNOMARİN Doküman Sistemi", initial_sidebar_state="expanded")

# --- PLATFORM ŞABLON SEÇİCİ ---
st.sidebar.markdown("### ⚙️ Sistem Ayarları")
secili_sablon = st.sidebar.radio(
    "📝 Çalışma Şablonunu Seçin:",
    ["⚓ INNOMAR Özel Teklif", "📄 Standart Proforma Fatura"]
)

gizle_checkbox = st.sidebar.checkbox("🔒 Birim Fiyatını Gizle", value=False)

if 'aktif_sablon' not in st.session_state or st.session_state.aktif_sablon != secili_sablon:
    st.session_state.aktif_sablon = secili_sablon
    if secili_sablon == "⚓ INNOMAR Özel Teklif":
        data = {
            'INSPECTION REMARK': ['ANA MAKİNE BAKIMLARI', 'SU YAPICI BAKIMLARI'],
            'UNIT': ['2 PIECES', '1 SET'],
            'PRICE': [40000.0, 12000.0]
        }
        st.session_state.not_alani = "* IMPORTANT NOTICE;\n- DURING MAINTENANCE IF DEFORMATION DETECTED...\n- PAYMENT: %50 BEFORE, %50 AFTER."
    else: 
        # Taslaktaki sütun yapısına göre güncellendi
        data = {
            'Açıklama': ['Örnek Hizmet/Ürün', ''],
            'Birim Fiyat': [1000.0, 0.0],
            'Adet': [1, 0],
            'Tutar': [1000.0, 0.0]
        }
        st.session_state.not_alani = "Banka Hesap Bilgilerimiz:\nBanka Adı: \nIBAN: \nHesap Sahibi: "
    
    st.session_state.veri_df = pd.DataFrame(data)
    st.rerun()

st.markdown(f"<h2 style='text-align: center;'>{secili_sablon.upper()}</h2>", unsafe_allow_html=True)

# --- ÜST PANEL ---
col_t, col_kur, col_fat = st.columns([1, 1, 1])
secilen_tarih = col_t.date_input("Belge Tarihi", datetime.date.today())
kur_secimi = col_kur.selectbox("Para Birimi", ["Türk Lirası (₺)", "Euro (€)", "Dolar ($)"])
fatura_no = col_fat.text_input("Fatura/Teklif No", "INV-2026-001")

sembol = "₺" if "Lira" in kur_secimi else ("€" if "Euro" in kur_secimi else "$")
kur_metin = "TL" if "Lira" in kur_secimi else ("EURO" if "Euro" in kur_secimi else "USD")

# Müşteri Bilgileri
col_m1, col_m2 = st.columns(2)
musteri_adi = col_m1.text_input("Müşteri Adı / Unvan")
musteri_adres = col_m2.text_area("Müşteri Adresi", height=68)

st.write("---")

# --- VERİ EDİTÖRÜ ---
df = st.session_state.veri_df
# Otomatik tutar hesaplama (Fatura şablonu için)
if secili_sablon == "📄 Standart Proforma Fatura":
    st.info("💡 Tutar sütunu, Birim Fiyat ve Adet girildiğinde otomatik hesaplanır.")
    
duzenlenmis_df = st.data_editor(df, num_rows="dynamic", use_container_width=True)

# Hesaplamaları Güncelle
if secili_sablon == "📄 Standart Proforma Fatura":
    duzenlenmis_df['Tutar'] = duzenlenmis_df['Birim Fiyat'] * duzenlenmis_df['Adet']
    ara_toplam = duzenlenmis_df['Tutar'].sum()
else:
    ara_toplam = duzenlenmis_df.iloc[:, -1].sum()

kdv = ara_toplam * 0.20
genel_toplam = ara_toplam + kdv

# --- ÖZET METRİKLER ---
c1, c2, c3 = st.columns(3)
c1.metric("Ara Toplam", f"{ara_toplam:,.2f} {sembol}")
c2.metric("KDV (%20)", f"{kdv:,.2f} {sembol}")
c3.metric("Genel Toplam", f"{genel_toplam:,.2f} {sembol}")

st.subheader("📄 Notlar ve Şartlar")
notlar = st.text_area("Belge altına eklenecek notlar:", value=st.session_state.not_alani, height=100)

# --- YARDIMCI FONKSİYONLAR ---
def cevir_tr(metin):
    tr_map = {'ş':'s', 'Ş':'S', 'ı':'i', 'İ':'I', 'ğ':'g', 'Ğ':'G', 'ü':'u', 'Ü':'U', 'ö':'o', 'Ö':'O', 'ç':'c', 'Ç':'C'}
    for k, v in tr_map.items(): metin = str(metin).replace(k, v)
    return metin

# --- PDF OLUŞTURMA (Görseldeki Taslağa Uygun) ---
def pdf_olustur():
    pdf = FPDF()
    pdf.add_page()
    pdf.set_auto_page_break(auto=True, margin=15)
    
    # Header - INNOMARIN Logo & Info
    pdf.set_font('Arial', 'B', 16)
    pdf.set_text_color(184, 134, 11) # Gold renk tonu
    pdf.cell(0, 10, "INNOMARIN", 0, 1, 'C')
    pdf.set_font('Arial', 'I', 8)
    pdf.cell(0, 5, "SAILING INTO THE FUTURE", 0, 1, 'C')
    
    pdf.ln(5)
    pdf.set_text_color(0, 0, 0)
    pdf.set_font('Arial', '', 9)
    pdf.cell(0, 4, "Klavuz Sok. No: 16/6 Heybeliada", 0, 1, 'C')
    pdf.cell(0, 4, "info@innomarin.com | www.innomarin.com", 0, 1, 'C')
    pdf.line(10, pdf.get_y()+2, 200, pdf.get_y()+2)
    
    pdf.ln(10)
    # Fatura Başlığı ve Bilgiler
    pdf.set_font('Arial', 'B', 12)
    pdf.cell(100, 10, "PROFORMA FATURA" if secili_sablon == "📄 Standart Proforma Fatura" else "TEKLIF / QUOTATION", 0, 0)
    pdf.set_font('Arial', '', 9)
    pdf.cell(0, 5, f"No: {fatura_no}", 0, 1, 'R')
    pdf.cell(0, 5, f"Tarih: {secilen_tarih.strftime('%d.%m.%Y')}", 0, 1, 'R')
    
    # Müşteri Bilgileri
    pdf.ln(5)
    pdf.set_font('Arial', 'B', 9)
    pdf.cell(0, 5, "Fatura Kesilen / Bill To:", 0, 1)
    pdf.set_font('Arial', '', 9)
    pdf.multi_cell(0, 5, f"Musteri: {cevir_tr(musteri_adi)}\nAdres: {cevir_tr(musteri_adres)}")
    
    pdf.ln(5)
    
    # Tablo
    headers = list(duzenlenmis_df.columns)
    col_width = 190 / len(headers)
    
    pdf.set_fill_color(240, 240, 240)
    pdf.set_font('Arial', 'B', 9)
    for h in headers:
        pdf.cell(col_width, 8, cevir_tr(h), 1, 0, 'C', True)
    pdf.ln()
    
    pdf.set_font('Arial', '', 8)
    for _, row in duzenlenmis_df.iterrows():
        for col in headers:
            val = row[col]
            if gizle_checkbox and "Birim" in col:
                text = "***"
            elif isinstance(val, float):
                text = f"{val:,.2f}"
            else:
                text = cevir_tr(val)
            pdf.cell(col_width, 7, text, 1, 0, 'C')
        pdf.ln()
        
    # Toplamlar
    pdf.ln(2)
    pdf.set_x(130)
    pdf.cell(35, 7, "Ara Toplam:", 1)
    pdf.cell(35, 7, f"{ara_toplam:,.2f} {kur_metin}", 1, 1, 'R')
    pdf.set_x(130)
    pdf.cell(35, 7, "KDV (%20):", 1)
    pdf.cell(35, 7, f"{kdv:,.2f} {kur_metin}", 1, 1, 'R')
    pdf.set_x(130)
    pdf.set_font('Arial', 'B', 9)
    pdf.cell(35, 7, "Toplam:", 1)
    pdf.cell(35, 7, f"{genel_toplam:,.2f} {kur_metin}", 1, 1, 'R')
    
    # Notlar
    pdf.ln(10)
    pdf.set_font('Arial', 'I', 8)
    pdf.multi_cell(0, 5, f"Notlar:\n{cevir_tr(notlar)}")
    
    return pdf.output(dest='S').encode('latin-1')

# --- DOWNLOAD ---
st.write("---")
if st.button("📕 PDF OLARAK İNDİR", use_container_width=True, type="primary"):
    pdf_data = pdf_olustur()
    st.download_button(
        label="DOSYAYI KAYDET",
        data=pdf_data,
        file_name=f"{fatura_no}.pdf",
        mime="application/pdf"
    )
