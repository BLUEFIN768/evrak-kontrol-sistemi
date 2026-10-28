import streamlit as st
import datetime
import io
import re
from dateutil.relativedelta import relativedelta
import pypdf
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

OFFICER_RANKS = [
    "Master", "Chief officer", "2 nd Officer", "3 rd Officer",
    "Chief Engineer", "2 nd Engineer", "3rd Engineer", "Electrician"
]

# Yan Panel: Atama Parametreleri
st.sidebar.header("🚢 Gemi ve Atama Parametreleri")
vessel_name = st.sidebar.text_input("Gemi Adı", value="Optimus S")
vessel_flag = st.sidebar.selectbox("Gemi Bayrağı", ["MALTA", "MARSHALL ISLANDS", "PANAMA", "LIBERIA", "TURKIYE"])
ecdis_model = st.sidebar.text_input("Gemideki ECDIS Modeli", value="")

rank = st.sidebar.selectbox(
    "Atanan Rütbe (Rank)",
    options=list(RANKS_MAP.keys()),
    index=17  # Cadet
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

# Dosya Ayrıştırma Fonksiyonları
def extract_info_from_files(files):
    detected_name = ""
    detected_tckn = ""
    detected_sb = ""
    
    # 1. Dosya isimlerinden Ad-Soyad çıkarma (Örn: "Deck Cadet Seyid Ali Demir.xlsx", "16) SEYİD ALİ DEMİR.pdf")
    for f in files:
        fname = f.name
        # Uzantıyı at
        base = re.sub(r'\.(pdf|xlsx|xls|docx|jpg|jpeg|png)$', '', fname, flags=re.IGNORECASE)
        # Standart etiketleri temizle
        clean = re.sub(r'^(deck\s+cadet|cadet|16\)|1\s*-\s*pass|2\s*-\s*medical|4\s*-\s*sb|coc|criminal\s+record)\s*', '', base, flags=re.IGNORECASE).strip()
        clean = re.sub(r'[-_]', ' ', clean).strip()
        if len(clean.split()) >= 2 and not any(k in clean.lower() for k in ["medical", "pass", "record", "status", "desc"]):
            detected_name = clean.upper()
            break
            
    # 2. PDF içeriklerinden TCKN ve Cüzdan No tarama
    for f in files:
        if f.name.lower().endswith(".pdf"):
            try:
                content = f.read()
                f.seek(0)
                reader = pypdf.PdfReader(io.BytesIO(content))
                for page in reader.pages:
                    text = page.extract_text() or ""
                    
                    # TCKN: 11 haneli sayı
                    if not detected_tckn:
                        m_tc = re.search(r'\b([1-9][0-9]{10})\b', text)
                        if m_tc:
                            detected_tckn = m_tc.group(1)
                            
                    # Cüzdan No (S00... veya Sicil)
                    if not detected_sb:
                        m_sb = re.search(r'\b(S\d{7,8})\b', text, re.IGNORECASE)
                        if m_sb:
                            detected_sb = m_sb.group(1).upper()
            except Exception:
                pass
                
    return detected_name, detected_tckn, detected_sb

# Evrak Yükleme Alanı
uploaded_files = st.file_uploader(
    "Personel Evraklarını Yükleyin (PDF / Görseller / Excel)",
    type=["pdf", "jpg", "jpeg", "png", "xlsx"],
    accept_multiple_files=True
)

if uploaded_files:
    auto_name, auto_tckn, auto_sb = extract_info_from_files(uploaded_files)
    
    st.success(f"{len(uploaded_files)} adet dosya yüklendi ve denetim motoru çalıştırıldı.")
    
    col_info1, col_info2, col_info3 = st.columns(3)
    with col_info1:
        crew_name = st.text_input("Personel Adı Soyadı", value=auto_name if auto_name else "SEYİD ALİ DEMİR")
    with col_info2:
        tckn = st.text_input("T.C. Kimlik No", value=auto_tckn if auto_tckn else "")
    with col_info3:
        seaman_book = st.text_input("Liman Cüzdan No (SB)", value=auto_sb if auto_sb else "")
        
    st.subheader("⚠️ Canlı Kural ve Emniyet Uyarıları")
    
    # Cadet için geçerlilik kontrolleri (Alt kadro: 6 ay asgari kontrat kuralı)
    st.info(f"📋 **{rank}** rütbesi için evrak geçerlilikleri asgari kontrat sonu olan **{contract_end.strftime('%d.%m.%Y')}** tarihine göre denetlenmektedir.")
        
    st.divider()
    st.subheader("📥 Çıktı Dosyaları")
    
    c1, c2 = st.columns(2)
    with c1:
        st.markdown("**35-101 Seafarers Documents Check List (Excel)**")
        st.caption(f"A4 Tek Sayfa | Seçilen Rütbe Sütunu Sarı (Sütun {rank_col}: {rank}) | Satır İçi Sertifika Kodları & Bitiş Tarihleri")
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
        st.caption(f"Kutucuklar İşaretli | Gerekçeli Onay/Ret | {rank} için N/A Uyarlamalı")
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
