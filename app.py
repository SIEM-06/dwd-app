import streamlit as st

# Sayfa ayarları
st.set_page_config(page_title="Innomarin - Bakım Modu", page_icon="🏗️", layout="centered")

# Görsel bir alan oluşturma
st.markdown("<br><br><br>", unsafe_allow_html=True)
st.markdown("<h1 style='text-align: center; color: #003399;'>🏗️ SİSTEM BAKIMDA</h1>", unsafe_allow_html=True)
st.markdown("---", unsafe_allow_html=True)

st.warning("Şu anda sistem üzerinde güncelleme ve bakım çalışmaları yapılmaktadır.")

st.info("""
**Yapılan İşlemler:**
- Taslak şablonları güncelleniyor.
- PDF ve Excel motorları optimize ediliyor.
- İndirme butonlarındaki hatalar gideriliyor.
""")

st.markdown("<br>", unsafe_allow_html=True)
st.markdown("<p style='text-align: center; color: gray;'>Anlayışınız için teşekkür ederiz. En kısa sürede tekrar yayında olacağız.</p>", unsafe_allow_html=True)

# Arka planda kimsenin işlem yapamaması için geri kalan tüm kodları buraya eklemiyoruz.
