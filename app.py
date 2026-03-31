import streamlit as st

# Sayfa ayarları
st.set_page_config(page_title="Sistem Bakımda", page_icon="🛠️", layout="centered")

# Görsel ve Metinler
st.markdown("<h1 style='text-align: center; color: #003399;'>🛠️ Sistem Bakımda</h1>", unsafe_allow_html=True)
st.markdown("<h3 style='text-align: center;'>Şu anda kısa süreli bir güncelleme yapıyoruz.</h3>", unsafe_allow_html=True)
st.write("---")
st.markdown("<p style='text-align: center; font-size: 18px;'>Sizlere daha iyi ve hatasız hizmet verebilmek için altyapımızı güncelliyoruz.<br>Lütfen birkaç dakika sonra sayfayı yenileyin.</p>", unsafe_allow_html=True)

# Uyarı kutusu
st.info("💡 Sabrınız ve anlayışınız için teşekkür ederiz. Innomarin Ekibi")
