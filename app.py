import streamlit as st
import pandas as pd
import io
import datetime
import os
from fpdf import FPDF
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image as xlImage

st.set_page_config(layout="wide", page_title="Doküman Oluşturucu Platform", initial_sidebar_state="expanded")

# =========================================================
# DOSYA AYARLARI
# =========================================================
ANTET_IMG = "antet.png"              # PDF + Excel için kullanılacak
WORD_TEMPLATE = "word_template.docx" # Word için kullanılacak hazır antetli şablon

# =========================================================
# YARDIMCI FONKSİYONLAR
# =========================================================
def cevir_tr(metin):
    tr_map = {
        'ş': 's', 'Ş': 'S',
        'ı': 'i', 'İ': 'I',
        'ğ': 'g', 'Ğ': 'G',
        'ü': 'u', 'Ü': 'U',
        'ö': 'o', 'Ö': 'O',
        'ç': 'c', 'Ç': 'C'
    }
    for k, v in tr_map.items():
        metin = metin.replace(k, v)
    return metin

def get_alignment(col_name):
    name = str(col_name).lower()
    if any(x in name for x in ['fiyat', 'price', 'tutar', 'total', 'amount']):
        return 'R'
    if any(x in name for x in ['sıra', 'no', 'kdv', 'adet', 'unit', 'qty', 'miktar']):
        return 'C'
    return 'L'

def get_pdf_widths(headers, total_w=190):
    widths = []
    for h in headers:
        name = str(h).lower()
        if any(x in name for x in ['sıra', 'no']):
            widths.append(10)
        elif any(x in name for x in ['kdv', 'adet', 'unit', 'qty']):
            widths.append(16)
        elif any(x in name for x in ['fiyat', 'price', 'tutar', 'amount', 'total']):
            widths.append(28)
        elif any(x in name for x in ['marka', 'brand']):
            widths.append(24)
        elif any(x in name for x in ['açıklama', 'remark', 'işlem', 'description']):
            widths.append(65)
        else:
            widths.append(24)

    total = sum(widths)
    scale = total_w / total if total > 0 else 1
    return [w * scale for w in widths]

def set_excel_col_widths(ws, headers):
    for i, header in enumerate(headers, 1):
        col_letter = get_column_letter(i)
        name = str(header).lower()

        if any(x in name for x in ['sıra', 'no']):
            ws.column_dimensions[col_letter].width = 8
        elif any(x in name for x in ['kdv', 'adet', 'unit', 'qty']):
            ws.column_dimensions[col_letter].width = 12
        elif any(x in name for x in ['fiyat', 'price', 'tutar', 'amount', 'total']):
            ws.column_dimensions[col_letter].width = 16
        elif any(x in name for x in ['açıklama', 'remark', 'işlem', 'description']):
            ws.column_dimensions[col_letter].width = 40
        else:
            ws.column_dimensions[col_letter].width = 20

def get_birim_col(df_columns):
    for col in df_columns:
        if "birim fiyat" in str(col).lower():
            return col
    return None

def para_formatla(v, kur_m):
    try:
        fiyat = float(str(v).replace(",", "."))
        if pd.isna(fiyat) or fiyat <= 0:
            return "-NIL-"
        return f"{fiyat:,.0f}".replace(",", ".") + f" {kur_m}"
    except:
        return "-NIL-"

def tablo_stil_hucre_word(cell, align='L', bold=False):
    p = cell.paragraphs[0]
    if align == 'R':
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    elif align == 'C':
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    else:
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT

    if p.runs:
        p.runs[0].bold = bold
        p.runs[0].font.size = Pt(9)

# =========================================================
# STREAMLIT UI
# =========================================================
st.sidebar.markdown("### ⚙️ Sistem Ayarları")
secili_sablon = st.sidebar.radio(
    "📝 Çalışma Şablonunu Seçin:",
    ["⚓ INNOMAR Özel Teklif", "📄 Standart Proforma Fatura"]
)

gizle_checkbox = st.sidebar.checkbox(
    "🔒 Birim Fiyatını Çıktılarda Gizle (Sansürle)",
    value=False,
    help="İşaretlendiğinde, indirilen dosyalarda Birim Fiyat sütunu '***' olarak görünür."
)

if 'aktif_sablon' not in st.session_state or st.session_state.aktif_sablon != secili_sablon:
    st.session_state.aktif_sablon = secili_sablon

    if secili_sablon == "⚓ INNOMAR Özel Teklif":
        data = {
            'INSPECTION REMARK': ['ANA MAKİNE BAKIMLARI', 'SU YAPICI BAKIMLARI'],
            'UNIT': ['2 PIECES', '1 SET'],
            'PRICE': [40000.0, 12000.0]
        }
        st.session_state.not_alani = (
            "* IMPORTANT NOTICE;\n"
            "- DURING MAINTENANCE IF DEFORMATION DETECTED ON WORKING SURFACE AND NEEDED TO RENEW COMPONENTS EACH PARTS WILL BE PRICED ADDITIONALLY.\n\n"
            "* REMARKS;\n"
            "- DELIVERY TIME FOR THE JOB IS 35 DAYS,\n"
            "- A DETAILED REPORT WILL BE SUBMITTED TO YOUR SIDE UPON COMPLETION OF THE WORK,\n"
            "- PAYMENT WILL BE ACCEPTED AS BELOW;\n"
            "    - %50 BEFORE WORK BEGINS,\n"
            "    - %50 UPON COMPLETION OF THE WORK."
        )
    else:
        data = {
            'Açıklama': ['Örnek Hizmet', ''],
            'Adet': ['1', '2'],
            'KDV': ['%20', '%20'],
            'Birim Fiyatı': ['1000', '500'],
            'Tutar': [1000.0, 1000.0]
        }
        st.session_state.not_alani = ""

    st.session_state.veri_df = pd.DataFrame(data)
    st.rerun()

st.markdown(f"<h2 style='text-align: center;'>{secili_sablon.upper()} SİSTEMİ</h2>", unsafe_allow_html=True)

# =========================================================
# ÜST PANEL
# =========================================================
col_t, col_kur = st.columns([1, 1])

secilen_tarih = col_t.date_input("Belge Tarihi", datetime.date.today())
tarih_metni = secilen_tarih.strftime("%d.%m.%Y")
dosya_tarihi = secilen_tarih.strftime("%d_%m_%Y")

kur_secimi = col_kur.selectbox("Para Birimi", ["Euro (€)", "Dolar ($)", "Türk Lirası (₺)"])
if "Euro" in kur_secimi:
    sembol, kur_metin = "€", "EURO"
elif "Dolar" in kur_secimi:
    sembol, kur_metin = "$", "USD"
else:
    sembol, kur_metin = "₺", "TL"

# =========================================================
# DİNAMİK SÜTUN YÖNETİMİ
# =========================================================
st.write("---")
st.caption("Aşağıdaki kutuya virgülle ayırarak istediğiniz kadar sütun ekleyebilir veya silebilirsiniz. Hesaplamanın doğru çalışması için toplam/tutar/fiyat sütunu en sonda olmalı.")

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

df = st.session_state.veri_df.copy()
son_sutun = df.columns[-1]
df[son_sutun] = pd.to_numeric(df[son_sutun], errors='coerce').fillna(0.0)

col_config = {}
for col in df.columns[:-1]:
    col_config[col] = st.column_config.TextColumn(col)
col_config[son_sutun] = st.column_config.NumberColumn(son_sutun, format=f"%d {sembol}")

st.info("💡 Tablodaki hücrelerin üzerine tıklayıp düzenleyebilirsin. Yeni satır için en alt satırı kullan.")
duzenlenmis_df = st.data_editor(df, column_config=col_config, num_rows="dynamic", use_container_width=True)

# =========================================================
# HESAPLAMALAR
# =========================================================
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

st.subheader("📄 Belge Altı Notları")
st.text_area(
    "Bu alana yazdığınız metin belgenin altına eklenecektir:",
    key="not_alani",
    height=150,
    placeholder="Buraya notlarınızı veya banka hesap bilgilerinizi girebilirsiniz..."
)

if st.button("🔄 Notları Sisteme Kaydet (İndirmeden Önce Basın)"):
    st.success("Notlar hafızaya alındı. Çıktı oluşturabilirsin.")

# =========================================================
# WORD OLUŞTUR
# =========================================================
def word_olustur(dataframe, a_str, k_str, g_str, tarih, notlar, kur_m, sablon_tipi, gizle_aktif):
    """
    Word için en sağlıklı yöntem: hazır antetli bir word_template.docx kullanmak.
    Eğer yoksa boş Document ile devam eder ama tam antet görünümü garanti olmaz.
    """
    if os.path.exists(WORD_TEMPLATE):
        doc = Document(WORD_TEMPLATE)
    else:
        doc = Document()
        p_warn = doc.add_paragraph()
        p_warn.add_run("UYARI: word_template.docx bulunamadı. Bu nedenle Word çıktısı tam antetli görünmeyebilir.").bold = True

    # Sayfa kenar boşlukları
    for section in doc.sections:
        section.top_margin = Cm(3.3)
        section.bottom_margin = Cm(2.8)
        section.left_margin = Cm(1.6)
        section.right_margin = Cm(1.6)

    # Başlık
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p_title.add_run("QUOTATION" if sablon_tipi == "⚓ INNOMAR Özel Teklif" else "PROFORMA FATURA")
    run.bold = True
    run.font.size = Pt(12)

    p_date = doc.add_paragraph()
    p_date.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r = p_date.add_run(f"Tarih / Date: {tarih}")
    r.bold = True
    r.font.size = Pt(10)

    headers = ['Sıra' if sablon_tipi != "⚓ INNOMAR Özel Teklif" else 'ITEM NO'] + list(dataframe.columns)
    birim_sutun = get_birim_col(dataframe.columns)

    table = doc.add_table(rows=1, cols=len(headers))
    table.style = 'Table Grid'
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    for idx, header in enumerate(headers):
        cell = table.rows[0].cells[idx]
        cell.text = str(header)
        tablo_stil_hucre_word(cell, align=get_alignment(header), bold=True)

    for index, row in dataframe.iterrows():
        row_cells = table.add_row().cells
        row_cells[0].text = str(index + 1)
        tablo_stil_hucre_word(row_cells[0], align='C')

        for c_idx, col_name in enumerate(dataframe.columns):
            val = row[col_name]
            align = get_alignment(col_name)

            if gizle_aktif and col_name == birim_sutun:
                row_cells[c_idx + 1].text = "***"
                tablo_stil_hucre_word(row_cells[c_idx + 1], align='C')
            elif col_name == dataframe.columns[-1]:
                row_cells[c_idx + 1].text = para_formatla(val, kur_m)
                tablo_stil_hucre_word(row_cells[c_idx + 1], align=align)
            else:
                row_cells[c_idx + 1].text = str(val)
                tablo_stil_hucre_word(row_cells[c_idx + 1], align=align)

    doc.add_paragraph()

    tot_table = doc.add_table(rows=3, cols=2)
    tot_table.style = 'Table Grid'
    tot_table.alignment = WD_TABLE_ALIGNMENT.RIGHT

    labels = ["TOTAL PRICE", "VAT (20%)", "GRAND TOTAL"] if sablon_tipi == "⚓ INNOMAR Özel Teklif" else ["Ara Toplam", "KDV %20", "GENEL TOPLAM"]
    values = [a_str, k_str, g_str]

    for i in range(3):
        tot_table.rows[i].cells[0].text = labels[i]
        tot_table.rows[i].cells[1].text = values[i]
        tablo_stil_hucre_word(tot_table.rows[i].cells[0], align='R', bold=(i == 2))
        tablo_stil_hucre_word(tot_table.rows[i].cells[1], align='R', bold=True)

    doc.add_paragraph()

    if notlar.strip():
        for satir in notlar.split('\n'):
            p = doc.add_paragraph(satir)
            p.paragraph_format.space_after = Pt(2)

    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()

# =========================================================
# PDF OLUŞTUR
# =========================================================
def pdf_olustur(dataframe, a_str, k_str, g_str, tarih, notlar, kur_m, sablon_tipi, gizle_aktif):
    birim_sutun = get_birim_col(dataframe.columns)

    class PDF(FPDF):
        def header(self):
            # Her sayfada aynı antet arkaplan
            if os.path.exists(ANTET_IMG):
                self.image(ANTET_IMG, x=0, y=0, w=210, h=297)
            self.set_y(28)

    pdf = PDF()
    pdf.set_auto_page_break(auto=True, margin=24)
    pdf.add_page()

    # İçerik üst boşluk
    pdf.set_y(38)

    # Başlık
    pdf.set_font("Arial", "B", 12)
    pdf.cell(0, 8, cevir_tr("MY ADA DRY DOCK SERVICES QUOTATION" if sablon_tipi == "⚓ INNOMAR Özel Teklif" else "PROFORMA FATURA"), 0, 1, 'C')
    pdf.set_font("Arial", "B", 10)
    pdf.cell(0, 7, cevir_tr(f"Tarih / Date: {tarih}"), 0, 1, 'R')
    pdf.ln(3)

    headers = ['Sıra' if sablon_tipi != "⚓ INNOMAR Özel Teklif" else 'NO'] + list(dataframe.columns)
    widths = get_pdf_widths(headers)

    # Tablo header
    pdf.set_font('Arial', 'B', 8)
    pdf.set_fill_color(235, 228, 213)
    for idx, header in enumerate(headers):
        pdf.cell(widths[idx], 8, cevir_tr(str(header)), 1, 0, 'C', fill=True)
    pdf.ln()

    # Tablo satırları
    pdf.set_font('Arial', '', 8)
    for index, row in dataframe.iterrows():
        pdf.cell(widths[0], 7, str(index + 1), 1, 0, 'C')
        for c_idx, col_name in enumerate(dataframe.columns):
            val = row[col_name]
            align = get_alignment(col_name)

            if gizle_aktif and col_name == birim_sutun:
                yazilacak = "***"
                align = 'C'
            elif col_name == dataframe.columns[-1]:
                yazilacak = para_formatla(val, kur_m)
            else:
                yazilacak = cevir_tr(str(val))

            pdf.cell(widths[c_idx + 1], 7, yazilacak, 1, 0, align)
        pdf.ln()

    pdf.ln(4)

    # Toplamlar sağ blok
    total_box_width = 70
    x_start = 210 - 10 - total_box_width
    current_y = pdf.get_y()

    labels = ["TOTAL PRICE", "VAT (20%)", "GRAND TOTAL"] if sablon_tipi == "⚓ INNOMAR Özel Teklif" else ["Ara Toplam", "KDV %20", "GENEL TOPLAM"]
    values = [a_str, k_str, g_str]

    for i in range(3):
        pdf.set_x(x_start)
        pdf.set_font('Arial', 'B' if i == 2 else '', 9)
        pdf.cell(35, 8, cevir_tr(labels[i]), 1, 0, 'R')
        pdf.set_font('Arial', 'B', 9)
        pdf.cell(35, 8, cevir_tr(values[i]), 1, 1, 'R')

    pdf.ln(6)

    # Notlar
    if notlar.strip():
        pdf.set_font('Arial', '', 8)
        pdf.multi_cell(0, 5, cevir_tr(notlar))

    return pdf.output(dest='S').encode('latin-1')

# =========================================================
# EXCEL OLUŞTUR
# =========================================================
def excel_olustur(dataframe, a_str, k_str, g_str, tarih, notlar, kur_m, sablon_tipi, gizle_aktif):
    wb = Workbook()
    ws = wb.active
    ws.title = "Belge"
    ws.sheet_view.showGridLines = False

    # Arka plan görseli gibi üstte tam antet
    if os.path.exists(ANTET_IMG):
        img = xlImage(ANTET_IMG)
        oran = 760 / img.width
        img.width = 760
        img.height = int(img.height * oran)
        ws.add_image(img, 'A1')

    # İçerik başlangıcı
    row_idx = 14
    birim_sutun = get_birim_col(dataframe.columns)

    # Başlık
    baslik = "MY ADA DRY DOCK SERVICES QUOTATION" if sablon_tipi == "⚓ INNOMAR Özel Teklif" else "PROFORMA FATURA"
    ws[f'B{row_idx}'] = baslik
    ws[f'B{row_idx}'].font = Font(bold=True, size=13)
    ws[f'F{row_idx}'] = f"Tarih / Date: {tarih}"
    ws[f'F{row_idx}'].font = Font(bold=True, size=11)
    row_idx += 2

    headers = ['Sıra' if sablon_tipi != "⚓ INNOMAR Özel Teklif" else 'ITEM NO'] + list(dataframe.columns)
    set_excel_col_widths(ws, headers)

    thin_border = Border(
        left=Side(border_style='thin', color='000000'),
        right=Side(border_style='thin', color='000000'),
        top=Side(border_style='thin', color='000000'),
        bottom=Side(border_style='thin', color='000000')
    )

    header_fill = PatternFill(start_color="EBE4D5", end_color="EBE4D5", fill_type="solid")

    # Header
    for col_num, header in enumerate(headers, 2):  # B sütunundan başlat
        cell = ws.cell(row=row_idx, column=col_num)
        cell.value = str(header)
        cell.font = Font(bold=True)
        cell.border = thin_border
        cell.fill = header_fill

        align = get_alignment(header)
        if align == 'R':
            cell.alignment = Alignment(horizontal="right", vertical="center")
        elif align == 'C':
            cell.alignment = Alignment(horizontal="center", vertical="center")
        else:
            cell.alignment = Alignment(horizontal="left", vertical="center")

    row_idx += 1

    # Data
    for index, row in dataframe.iterrows():
        no_cell = ws.cell(row=row_idx, column=2)
        no_cell.value = index + 1
        no_cell.border = thin_border
        no_cell.alignment = Alignment(horizontal="center", vertical="center")

        for c_idx, col_name in enumerate(dataframe.columns, start=3):
            cell = ws.cell(row=row_idx, column=c_idx)
            align = get_alignment(col_name)
            val = row[col_name]

            if gizle_aktif and col_name == birim_sutun:
                cell.value = "***"
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif col_name == dataframe.columns[-1]:
                cell.value = para_formatla(val, kur_m)
                if align == 'R':
                    cell.alignment = Alignment(horizontal="right", vertical="center")
                elif align == 'C':
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                else:
                    cell.alignment = Alignment(horizontal="left", vertical="center")
            else:
                cell.value = str(val)
                if align == 'R':
                    cell.alignment = Alignment(horizontal="right", vertical="center")
                elif align == 'C':
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                else:
                    cell.alignment = Alignment(horizontal="left", vertical="center")

            cell.border = thin_border

        row_idx += 1

    # Toplamlar
    row_idx += 1
    label_col = len(headers) + 1
    value_col = len(headers) + 2

    labels = ["TOTAL PRICE", "VAT (20%)", "GRAND TOTAL"] if sablon_tipi == "⚓ INNOMAR Özel Teklif" else ["Ara Toplam", "KDV %20", "GENEL TOPLAM"]
    values = [a_str, k_str, g_str]

    for i in range(3):
        c1 = ws.cell(row=row_idx + i, column=label_col)
        c2 = ws.cell(row=row_idx + i, column=value_col)

        c1.value = labels[i]
        c2.value = values[i]

        c1.border = thin_border
        c2.border = thin_border
        c1.alignment = Alignment(horizontal="right", vertical="center")
        c2.alignment = Alignment(horizontal="right", vertical="center")
        c1.font = Font(bold=(i == 2))
        c2.font = Font(bold=True)

    row_idx += 5

    # Notlar
    if notlar.strip():
        for satir in notlar.split('\n'):
            ws.cell(row=row_idx, column=2).value = satir
            row_idx += 1

    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()

# =========================================================
# DOSYA KONTROL UYARILARI
# =========================================================
st.write("---")
st.markdown("### 📎 Şablon Kontrolü")

if os.path.exists(ANTET_IMG):
    st.success("antet.png bulundu. PDF ve Excel anteti hazır.")
else:
    st.error("antet.png bulunamadı. PDF ve Excel çıktısında antet görünmeyecek.")

if os.path.exists(WORD_TEMPLATE):
    st.success("word_template.docx bulundu. Word anteti hazır.")
else:
    st.warning("word_template.docx bulunamadı. Word çıktısı yine oluşur ama tam antetli görünüm garanti olmaz.")

# =========================================================
# İNDİRME BUTONLARI
# =========================================================
st.markdown("### 📥 Çıktı Al")
btn_word, btn_excel, btn_pdf = st.columns(3)

with btn_word:
    st.download_button(
        "📄 WORD İNDİR",
        data=word_olustur(
            duzenlenmis_df, ara_str, kdv_str, genel_str,
            tarih_metni, st.session_state.not_alani,
            kur_metin, secili_sablon, gizle_checkbox
        ),
        file_name=f"{secili_sablon.split()[1]}_{dosya_tarihi}.docx",
        type="primary",
        use_container_width=True
    )

with btn_excel:
    st.download_button(
        "📊 EXCEL İNDİR",
        data=excel_olustur(
            duzenlenmis_df, ara_str, kdv_str, genel_str,
            tarih_metni, st.session_state.not_alani,
            kur_metin, secili_sablon, gizle_checkbox
        ),
        file_name=f"{secili_sablon.split()[1]}_{dosya_tarihi}.xlsx",
        type="primary",
        use_container_width=True
    )

with btn_pdf:
    st.download_button(
        "📕 PDF İNDİR",
        data=pdf_olustur(
            duzenlenmis_df, ara_str, kdv_str, genel_str,
            tarih_metni, st.session_state.not_alani,
            kur_metin, secili_sablon, gizle_checkbox
        ),
        file_name=f"{secili_sablon.split()[1]}_{dosya_tarihi}.pdf",
        type="primary",
        use_container_width=True
    )
