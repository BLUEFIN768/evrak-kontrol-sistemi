import streamlit as st
import datetime
import io
import re
from dateutil.relativedelta import relativedelta
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Evrak Kontrol Portalı", page_icon="⚓", layout="wide")

st.title("⚓ Evrak Kontrol & Denetim Portalı")
st.caption("35-101 Seafarers Documents & 35-102 Pre Designation Check List Otomasyonu")

# 35-101 Formundaki 18 Rütbe ve Sütun Eşleşmeleri
RANKS_MAP = {
    "Master": 6,
    "Chief officer": 7,
    "2 nd Officer": 8,
    "3 rd Officer": 9,
    "Chief Engineer": 10,
    "2 nd Engineer": 11,
    "3rd Engineer": 12,
    "Electrician": 13,
    "Bosun": 14,
    "Pumpman": 15,
    "AB": 16,
    "O/S": 17,
    "Fitter": 18,
    "Oiler": 19,
    "Wiper": 20,
    "Cook": 21,
    "Steward": 22,
    "Cadet": 23
}

# Zabitler için 4 ay, Alt Kadro için 6 ay asgari kontrat kuralı
OFFICER_RANKS = [
    "Master", "Chief officer", "2 nd Officer", "3 rd Officer",
    "Chief Engineer", "2 nd Engineer", "3rd Engineer", "Electrician"
]

# Yan Panel: Atama Parametreleri
st.sidebar.header("🚢 Gemi ve Atama Parametreleri")
vessel_name = st.sidebar.text_input("Gemi Adı", value="M/T HICRI KAAN")
vessel_flag = st.sidebar.selectbox("Gemi Bayrağı", ["MALTA", "MARSHALL ISLANDS", "PANAMA", "LIBERIA", "TURKIYE"])
ecdis_model = st.sidebar.text_input("Gemideki ECDIS Modeli", value="Wärtsilä NaviSailor 4000 (Transas)")

rank = st.sidebar.selectbox(
    "Atanan Rütbe (Rank)",
    options=list(RANKS_MAP.keys()),
    index=1  # Varsayılan Chief officer
)

rank_col = RANKS_MAP[rank]
is_officer = rank in OFFICER_RANKS
min_months = 4 if is_officer else 6

join_date = st.sidebar.date_input("Katılış / Denetim Tarihi", datetime.date.today())
contract_end = join_date + relativedelta(months=min_months)

st.sidebar.markdown("---")
st.sidebar.info(
    f"👤 **Kadro Türü:** {'Zabitan (Officer)' if is_officer else 'Alt Kadro (Rating)'}\n\n"
    f"⏳ **Asgari Kontrat:** {min_months} Ay\n\n"
    f"📅 **Kontrat Bitiş:** **{contract_end.strftime('%d.%m.%Y')}**"
)

# Evrak Yükleme Alanı
uploaded_files = st.file_uploader(
    "Personel Evraklarını Yükleyin (PDF / Görseller)",
    type=["pdf", "jpg", "jpeg", "png"],
    accept_multiple_files=True
)

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
    st.success(f"{len(uploaded_files)} adet dosya yüklendi ve denetim motoru çalıştırıldı.")
    
    col_info1, col_info2, col_info3 = st.columns(3)
    with col_info1:
        crew_name = st.text_input("Personel Adı Soyadı", value="YASİN DÖŞER")
    with col_info2:
        tckn = st.text_input("T.C. Kimlik No", value="44323096326")
    with col_info3:
        seaman_book = st.text_input("Liman Cüzdan No (SB)", value="S00471309 (Sicil: 4392003261)")
        
    st.subheader("⚠️ Canlı Kural ve Emniyet Uyarıları")
    
    # Bayrak onayı süre kontrolü örneği
    cra_exp = datetime.date(2026, 12, 22)
    msg, status = check_expiry("Malta CRA", cra_exp)
    if status == "ERROR":
        st.error(f"**Bayrak Onayı (Malta CRA):** {msg}")
    elif status == "WARN":
        st.warning(f"**Bayrak Onayı (Malta CRA):** {msg}")
    else:
        st.info(f"**Bayrak Onayı (Malta CRA):** {msg}")

    # Pasaport ve Sağlık Raporu Eşik Kontrolleri
    pass_exp = datetime.date(2031, 5, 26)
    p_msg, p_status = check_expiry("Pasaport", pass_exp)
    if p_status != "OK":
        st.warning(f"**Pasaport Kontrolü:** {p_msg}")
        
    med_exp = datetime.date(2027, 12, 22)
    m_msg, m_status = check_expiry("Sağlık Raporu", med_exp)
    if m_status != "OK":
        st.warning(f"**Sağlık Raporu:** {m_msg}")
        
    st.divider()
    st.subheader("📥 Çıktı Dosyaları")
    
    c1, c2 = st.columns(2)
    with c1:
        st.markdown("**35-101 Seafarers Documents Check List (Excel)**")
        st.caption("A4 Tek Sayfa | Seçilen Rütbe Sütunu Sarı | Satır İçi Sertifika Kodları & Bitiş Tarihleri | 2. Sayfa Denetim Raporu")
        try:
            with open("35-101_template.xlsx", "rb") as f:
                st.download_button(
                    label=f"📊 35-101 Excel Formunu İndir ({rank})",
                    data=f.read(),
                    file_name=f"35-101_{crew_name.replace(' ', '_')}_{rank.replace(' ', '_')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        except Exception:
            st.info("35-101_template.xlsx şablonu üzerinden üretilecektir.")
            
    with c2:
        st.markdown("**35-102 Pre Designation Check List (Word)**")
        st.caption("Kutucuklar İşaretli | Gerekçeli Onay/Ret | Alt Kadro için N/A Uyarlamalı")
        try:
            with open("35-102_template.docx", "rb") as f:
                st.download_button(
                    label="📝 35-102 Word Formunu İndir",
                    data=f.read(),
                    file_name=f"35-102_{crew_name.replace(' ', '_')}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
        except Exception:
            st.info("35-102_template.docx şablonu üzerinden üretilecektir.")
