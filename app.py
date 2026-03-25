import streamlit as st
import pandas as pd
# PDF veya Excel oluşturma kütüphanelerini buraya ekleyeceğiz (Örn: reportlab, xlsxwriter)

# --- SAYFA AYARLARI VE LOGO ---
st.set_page_config(layout="wide", page_title="Bisiklet Teklif Portali")

# Sol üst köşeye logo placeholder (Örnek bisiklet logosu)
# Eniştenin gerçek logosunu buraya st.image("firma_logosu.png") ile ekleyeceğiz.
st.markdown("## 🚲 [ŞİRKET LOGO PLACEHOLDER]") 

st.title("BİSİKLET TEKLİF OLUŞTURUCU | Streamlit Portali")
st.write("---")

# --- VERİ SETİ (Örnek) ---
# Gerçek uygulamada bu veri Excel'den veya API'den gelecek.
if 'veri_df' not in st.session_state:
    data = {
        'Parça Adı': ['Bisiklet Tekerleği', 'Jel Sele', 'Karbon Gidon', 'Vites Arttırıcı', 'Fren Balatası'],
        'Adet': [2, 1, 1, 1, 4],
        'Birim Fiyat (TL)': [2000.0, 450.0, 1200.0, 850.0, 150.0]
    }
    st.session_state.veri_df = pd.DataFrame(data)

df = st.session_state.veri_df

# --- İNTERAKTİF TABLO VE HESAPLAMA ---
st.subheader("Teklif İçeriğini Düzenle")

# Streamlit'in veri düzenleme aracı (data_editor) tam olarak görseldeki işi yapar
duzenlenmis_df = st.data_editor(
    df,
    column_config={
        "Adet": st.column_config.NumberColumn(help="Miktarı girin"),
        "Birim Fiyat (TL)": st.column_config.NumberColumn(format="%.2f TL"),
    },
    disabled=["Parça Adı"], # Parça adları değiştirilemesin
    num_rows="dynamic" # Yeni satır eklenebilsin
)

# Arka planda otomatik hesaplama
duzenlenmis_df['Satır Toplamı (TL)'] = duzenlenmis_df['Adet'] * duzenlenmis_df['Birim Fiyat (TL)']
genel_toplam = duzenlenmis_df['Satır Toplamı (TL)'].sum()

# --- SONUÇLAR VE AKSİYON BUTONLARI ---
st.write("---")
col1, col2 = st.columns([3, 1])

with col2:
    st.markdown(f"### GENEL TOPLAM: **{genel_toplam:,.2f} TL**")

# Buton Alanı
st.write("")
col3, col4, col5 = st.columns([2, 2, 4])

with col3:
    # BU BUTON GÖRSELDEKİ ANA AKSİYON
    if st.button("DÖKÜMANI HAZIRLA (LOGO VE FORMATLI PDF)", type="primary"):
        st.info("🔄 PDF formatlanıyor ve logo ekleniyor... (Bu kısma kod eklenecek)")
        # BURAYA: duzenlenmis_df'i alıp, logoyu ekleyip PDF oluşturan fonksiyon gelecek.
        # Örn: pdf_olustur(duzenlenmis_df, logo_path, "teklif.pdf")

with col4:
    if st.button("Sıfırla"):
        del st.session_state.veri_df
        st.experimental_rerun()

# Alt kısımdaki bilgilendirme notu
st.write("---")
st.caption("ℹ️ Bu prototip, kullanıcıdan gelen tabloyu işleyip, otomatik logo ve özel şablon ile PDF dosyası oluşturacak sistemi simüle eder. Gerçekleştirilen her değişiklik PDF çıktısını da günceller.")
