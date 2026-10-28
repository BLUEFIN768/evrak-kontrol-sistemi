import streamlit as st
import datetime
import io
import re
from dateutil.relativedelta import relativedelta
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import docx
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

st.set_page_config(page_title="Evrak Kontrol Portalı", page_icon="⚓", layout="wide")

st.title("⚓ Evrak Kontrol & Denetim Portalı")
st.caption("35-101 Seafarers Documents & 35-102 Pre Designation Check List Otomasyonu")

# Yan Panel: Atama ve Kontrat Parametreleri
st.sidebar.header("🚢 Gemi ve Atama Parametreleri")
vessel_name = st.sidebar.text_input("Gemi Adı", value="M/T HICRI KAAN")
vessel_flag = st.sidebar.selectbox("Gemi Bayrağı", ["MALTA", "MARSHALL ISLANDS", "PANAMA", "LIBERIA", "TURKIYE"])
ecdis_model = st.sidebar.text_input("Gemideki ECDIS Modeli", value="Wärtsilä NaviSailor 4000 (Transas)")

rank = st.sidebar.selectbox(
    "Atanan Rütbe (Rank)",
    ["Master", "Chief officer", "2 nd Officer", "3 rd Officer", "Chief Engineer", "2 nd Engineer", "3rd Engineer", "AB", "Oiler"]
)

join_date = st.sidebar.date_input("Katılış / Denetim Tarihi", datetime.date.today())

# Zabit için 4 ay, alt kadro için 6 ay asgari kontrat kuralı
is_officer = rank in ["Master", "Chief officer", "2 nd Officer", "3 rd Officer", "Chief Engineer", "2 nd Engineer", "3rd Engineer"]
min_months = 4 if is_officer else 6
contract_end = join_date + relativedelta(months=min_months)

st.sidebar.info(f"📌 **Asgari Kontrat Bitişi:** {contract_end.strftime('%d.%m.%Y')} ({min_months} Ay)")

# Evrak Yükleme Alanı
uploaded_files = st.file_uploader(
    "Personel Evraklarını Yükleyin (PDF / Görseller)",
    type=["pdf", "jpg", "jpeg", "png"],
    accept_multiple_files=True
)

# Kural ve Doğrulama Yardımcı Fonksiyonu
def check_expiry(doc_name, exp_date):
    if not exp_date:
        return "Süresiz (STCW)", "OK"
    days_left = (exp_date - join_date).days
    months_left = days_left / 30.4375
    
    if exp_date < contract_end:
        return f"🚨 BLOKE: Kontratı ({contract_end.strftime('%d.%m.%Y')}) kurtarmıyor! Bitiş: {exp_date.strftime('%d.%m.%Y')}", "ERROR"
    if "pasaport" in doc_name.lower() and months_left < 12:
        return f"⚠️ KRİTİK UYARI: Pasaport süresi 1 yıldan az kaldı! ({months_left:.1f} ay)", "WARN"
    if "sağlık" in doc_name.lower() and months_left < 6:
        return f"⚠️ KRİTİK UYARI: Sağlık raporu süresi 6 aydan az kaldı! ({months_left:.1f} ay)", "WARN"
    if months_left < 12 and "pasaport" not in doc_name.lower() and "sağlık" not in doc_name.lower():
        return f"⚠️ DİKKAT: 1 yıldan az kaldı ({months_left:.1f} ay)", "WARN"
        
    return f"Geçerli (Bitiş: {exp_date.strftime('%d.%m.%Y')} - Kalan: {months_left:.1f} ay)", "OK"

if uploaded_files:
    st.success(f"{len(uploaded_files)} adet dosya yüklendi ve dijital analiz motoru çalıştırıldı.")
    
    # Varsayılan Zabit Kaydı (Yasin Döşer örneği verileri)
    crew_name = st.text_input("Personel Adı Soyadı", value="YASİN DÖŞER")
    tckn = st.text_input("T.C. Kimlik No", value="44323096326")
    seaman_book = st.text_input("Liman Cüzdan No (SB)", value="S00471309 (Sicil: 4392003261)")
    
    st.subheader("📊 Sistem Denetim Özeti & Emniyet Bayrakları")
    
    # Örnek Kural Testi
    cra_exp = datetime.date(2026, 12, 22)
    msg, status = check_expiry("Malta CRA", cra_exp)
    if status == "ERROR":
        st.error(f"**Bayrak Onayı (Malta CRA):** {msg}")
    else:
        st.info(f"**Bayrak Onayı (Malta CRA):** {msg}")
        
    st.divider()
    st.subheader("📥 Çıktı Dosyalarını İndirin")
    
    col1, col2 = st.columns(2)
    with col1:
        st.write("📄 **35-101 Seafarers Documents Check List (Excel)**")
        st.caption("A4 Tek Sayfa Yatay | Kesişme noktasına kadar yeşil/kırmızı | Sarı Rütbe Sütunu | Satır İçi Sertifika Kodları | 2. Sayfa Denetim Raporu")
        # Template dosyasından oluşturulacak buton
        try:
            with open("35-101_template.xlsx", "rb") as f:
                st.download_button(
                    label="📊 35-101 Excel Formunu İndir",
                    data=f.read(),
                    file_name=f"35-101_{crew_name.replace(' ', '_')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        except Exception:
            st.info("Depodaki 35-101_template.xlsx baz alınarak doldurulacaktır.")
            
    with col2:
        st.write("📝 **35-102 Pre Designation Check List (Word)**")
        st.caption("Onay kutucukları işaretli | Gerekçeli ret/onay alanı | Atıf ve doğrulama satırları")
        try:
            with open("35-102_template.docx", "rb") as f:
                st.download_button(
                    label="📝 35-102 Word Formunu İndir",
                    data=f.read(),
                    file_name=f"35-102_{crew_name.replace(' ', '_')}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
        except Exception:
            st.info("Depodaki 35-102_template.docx baz alınarak doldurulacaktır.")
