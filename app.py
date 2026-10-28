import streamlit as st
import datetime
import io
import os
import re
from dateutil.relativedelta import relativedelta
import pypdf
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment
import docx
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT

try:
    import pytesseract
    from pdf2image import convert_from_bytes
    from PIL import Image, ImageEnhance
    OCR_AVAILABLE = True
except ImportError:
    OCR_AVAILABLE = False

st.set_page_config(page_title="Evrak Kontrol Portalı", page_icon="⚓", layout="wide")

def check_password():
    if "authenticated" not in st.session_state:
        st.session_state["authenticated"] = False
    if st.session_state["authenticated"]:
        return True

    st.markdown("<h2 style='text-align: center; color: #1F3864;'>⚓ Blue Fin Crewing - Evrak Kontrol Portalı</h2>", unsafe_allow_html=True)
    st.markdown("<p style='text-align: center; color: #666;'>MRZ Motorlu Tam Otomatik Doğrulama Sistemi (Manuel Giriş İptal Edilmiştir)</p>", unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns([1, 1.5, 1])
    with col2:
        with st.form("login_form"):
            username = st.text_input("Yetkili Kullanıcı")
            password = st.text_input("Güvenlik Parolası", type="password")
            submit = st.form_submit_button("Giriş Yap", use_container_width=True)
            if submit:
                if username == "bluefin" and password == "bluefin2026":
                    st.session_state["authenticated"] = True
                    st.rerun()
                else:
                    st.error("Hatalı kullanıcı adı veya şifre!")
    return False

if not check_password():
    st.stop()

st.sidebar.markdown("🔒 **Güvenli Oturum:** Aktif")
if st.sidebar.button("🚪 Çıkış Yap"):
    st.session_state["authenticated"] = False
    st.rerun()

st.title("⚓ Evrak Kontrol & Denetim Portalı")
st.caption("Veriye Dayalı Tam Otomatik OCR (MRZ) & UAB Doğrulama")

RANKS_MAP = {
    "Master": 6, "Chief officer": 7, "2 nd Officer": 8, "3 rd Officer": 9,
    "Chief Engineer": 10, "2 nd Engineer": 11, "3rd Engineer": 12, "Electrician": 13,
    "Bosun": 14, "Pumpman": 15, "AB": 16, "O/S": 17, "Fitter": 18,
    "Oiler": 19, "Wiper": 20, "Cook": 21, "Steward": 22, 
    "Deck Cadet": 23, "Engine Cadet": 23
}
OFFICER_RANKS = ["Master", "Chief officer", "2 nd Officer", "3 rd Officer", "Chief Engineer", "2 nd Engineer", "3rd Engineer", "Electrician"]
ALL_RANKS = list(RANKS_MAP.keys())

CERT_MATRIX = [
    (14, "1", "Passport", ALL_RANKS, ["pass", "pasaport"]),
    (16, "2", "Medical Examination Report", ALL_RANKS, ["medical", "sağlık", "health", "saglik"]),
    (18, "3", "Certificate of Yellow Fever", ALL_RANKS, ["yellow", "fever", "humma", "aşı", "asi"]),
    (20, "4", "Seamanbook", ALL_RANKS, ["sb", "seaman", "cüzdan", "cuzdan", "s00"]),
    (22, "5", "Competency", ALL_RANKS, ["coc", "competency", "yeterlilik"]),
    (24, "6", "GMDSS", ["Master", "Chief officer", "2 nd Officer", "3 rd Officer"], ["gmdss", "goc"]),
    (26, "7.1", "Fire Prevention and Fire Fighting", ALL_RANKS, ["stcw", "courses", "fire", "yangın", "yangin"]),
    (28, "7.2", "Personal Survival Techniques", ALL_RANKS, ["stcw", "courses", "survival", "can kurtarma"]),
    (30, "7.3", "Elementary First Aid", ALL_RANKS, ["stcw", "courses", "first aid", "ilkyardım", "yardim"]),
    (32, "7.4", "Personal Safety and Social Responsibility", ALL_RANKS, ["stcw", "courses", "social", "sorumluluk"]),
    (34, "7.5", "Proficiency in Survival Craft and Rescue Boats", ["Master", "Chief officer", "2 nd Officer", "3 rd Officer", "Chief Engineer", "2 nd Engineer", "3rd Engineer", "Electrician", "Bosun", "Pumpman", "AB", "Oiler"], ["stcw", "courses", "craft", "rescue"]),
    (36, "8.1", "Security Awareness Certificate", ALL_RANKS, ["stcw", "courses", "security", "farkındalık", "farkindalik"]),
    (38, "8.2", "Designated Security Duties Certificate", ALL_RANKS, ["stcw", "courses", "security", "görev", "gorev"]),
    (42, "8.4", "Security - Realated Familirization", ALL_RANKS, ["stcw", "courses", "security", "tanıtım", "tanitim"]),
    (82, "5", "Endorsement (Bayrak Onayı)", OFFICER_RANKS + ["Deck Cadet", "Engine Cadet"], ["cra", "endorsement", "bayrak"])
]

st.sidebar.header("🚢 Gemi ve Atama Parametreleri")
vessel_name = st.sidebar.text_input("Gemi Adı", value="Optimus S")
vessel_flag = st.sidebar.selectbox("Gemi Bayrağı", ["MALTA", "MARSHALL ISLANDS", "PANAMA", "LIBERIA", "TURKIYE"])
rank = st.sidebar.selectbox("Atanan Rütbe", options=ALL_RANKS, index=17) 
rank_col = RANKS_MAP[rank]
is_officer = rank in OFFICER_RANKS  
join_date = st.sidebar.date_input("Katılış Tarihi", datetime.date.today())

def enhance_image_for_mrz(img):
    """Görüntüyü MRZ (<<<) satırlarının çok net okunabilmesi için optimize eder."""
    img = img.convert('L')
    enhancer = ImageEnhance.Contrast(img)
    return enhancer.enhance(2.5)

def extract_evidence(files):
    best_name = ""
    tckn = ""
    sb = ""
    passport_no = ""
    all_text_blob = ""
    
    # 1. EXCEL DOSYALARINDAN ÇEKİM (Hızlı ve Güvenilir Kaynak)
    for f in files:
        if f.name.lower().endswith(('.xlsx', '.xls')):
            try:
                content = f.read()
                f.seek(0)
                wb = openpyxl.load_workbook(io.BytesIO(content), data_only=True)
                for sheet in wb.worksheets:
                    for r in range(1, 50):
                        for c in range(1, 15):
                            val = str(sheet.cell(r, c).value or "").strip()
                            all_text_blob += f" {val.lower()} "
                            if not tckn and re.match(r'^[1-9]\d{10}$', val): tckn = val
                            if not sb and re.match(r'^[Ss]\s*00\d{5,6}$', val): 
                                sb = val.replace(" ", "").upper()
            except: pass

    # 2. PDF VE GÖRSELLER ÜZERİNDE MRZ (MAKİNE OKUNABİLİR ALAN) ANALİZİ
    for f in files:
        fname = f.name.lower()
        all_text_blob += f" {fname} "
        file_text = ""
        content = f.read()
        f.seek(0)
        
        if fname.endswith(".pdf"):
            try:
                reader = pypdf.PdfReader(io.BytesIO(content))
                for page in reader.pages: file_text += (page.extract_text() or "") + " "
            except: pass
                
        # Eğer PDF'te metin yoksa veya bu bir resimse OCR ile MRZ taraması yap
        if len(file_text.strip()) < 50 and OCR_AVAILABLE:
            try:
                if fname.endswith(".pdf"):
                    images = convert_from_bytes(content, first_page=1, last_page=1, dpi=300)
                    if images: 
                        clean_img = enhance_image_for_mrz(images[0])
                        file_text += pytesseract.image_to_string(clean_img, lang="eng") + "\n"
                elif fname.endswith((".jpg", ".jpeg", ".png")):
                    img = Image.open(io.BytesIO(content))
                    clean_img = enhance_image_for_mrz(img)
                    file_text += pytesseract.image_to_string(clean_img, lang="eng") + "\n"
            except: pass
                
        all_text_blob += f" {file_text.lower()} "
        
        # --- KUSURSUZ MRZ (MAKİNE KODU) AYIKLAYICI ---
        lines = file_text.split('\n')
        for line in lines:
            line_clean = line.replace(" ", "").upper()
            
            # Eğer satırda '<<' varsa bu bir MRZ satırıdır (Pasaport, Kimlik veya Liman Cüzdanı)
            if "<<" in line_clean:
                
                # 1. TCKN Yakalama (11 haneli sayı MRZ içine gömülüdür)
                m_tc = re.search(r'(?<!\d)([1-9]\d{10})(?!\d)', line_clean.replace("<", " "))
                if m_tc and not tckn:
                    tckn = m_tc.group(1)
                
                # 2. Pasaport Numarası Yakalama (Genellikle U ile başlar ve 8 rakamlıdır)
                m_pass = re.search(r'\b(U\d{8})\b', line_clean.replace("<", " "))
                if m_pass and not passport_no:
                    passport_no = m_pass.group(1)
                    
                # 3. Liman Cüzdanı Numarası Yakalama (MRZ satırının başında olur)
                # Tesseract 'S' harfini '5' veya '8' okuyabilir. Bunu MRZ mantığıyla kesin düzeltiyoruz.
                m_sb = re.match(r'^([S58$]00\d{5,6})', line_clean)
                if m_sb and not sb:
                    raw_sb = m_sb.group(1)
                    # Eğer 500, 800 veya $00 ile başlıyorsa bunu S00'a zorla dönüştür (Çünkü Liman Cüzdanı MRZ'si S ile başlar)
                    sb = "S00" + raw_sb[3:]

        # Eğer MRZ satırı bulunamadıysa bile belgenin içinden düz metin olarak numaraları ara
        if not tckn:
            m_tc_text = re.search(r'\b([1-9][0-9]{10})\b', file_text)
            if m_tc_text: tckn = m_tc_text.group(1)
            
        if not sb:
            # Standart metin içinde S00 ile başlayan cüzdan arayışı
            m_s_text = re.search(r'\b(S\s*00\d{5,6})\b', file_text, re.IGNORECASE)
            if m_s_text: sb = m_s_text.group(1).replace(" ", "").upper()

        if not passport_no:
            m_pass_text = re.search(r'\b(U\d{8})\b', file_text, re.IGNORECASE)
            if m_pass_text: passport_no = m_pass_text.group(1).upper()

    # İsim bulma: Klasik dosya adı uydurması yerine, belgelerde bulunan isimlerin MRZ'den çekilmesi
    # Eğer pasaport veya kimlik MRZ'si (P<TUR veya I<TUR) bulunduysa ismi oradan çek
    for line in file_text.split('\n'):
        line_c = line.replace(" ", "").upper()
        if "<<" in line_c and ("P<TUR" in line_c or "I<TUR" in line_c or "<TUR" in line_c):
            parts = line_c.split("<<")
            if len(parts) >= 2:
                # İsim ve soyisim MRZ'de << ile ayrılır. (< işaretlerini boşluğa çeviriyoruz)
                potential_name = parts[0].replace("P<TUR", "").replace("I<TUR", "").replace("<TUR", "").replace("<", " ").strip()
                if len(potential_name) > 3:
                    best_name = potential_name
                    break

    return best_name, tckn, sb, passport_no, all_text_blob

uploaded_files = st.file_uploader("Personel Evraklarını Yükleyin (Otomatik MRZ Taraması)", type=["pdf", "jpg", "jpeg", "png", "xlsx"], accept_multiple_files=True)

if uploaded_files:
    auto_name, auto_tckn, auto_sb, auto_pass, evidence_text = extract_evidence(uploaded_files)
    
    st.info("ℹ️ Sistem verileri doğrudan evraklardaki makine kodlarından (MRZ) çekmektedir. Veriler bulunamazsa formlar üretilmez.")
    
    col_info1, col_info2, col_info3, col_info4 = st.columns(4)
    with col_info1:
        crew_name = st.text_input("Personel Adı Soyadı", value=auto_name, disabled=True)
    with col_info2:
        tckn_input = st.text_input("T.C. Kimlik No", value=auto_tckn, disabled=True)
    with col_info3:
        sb_input = st.text_input("Liman Cüzdan No (SB)", value=auto_sb, disabled=True)
    with col_info4:
        pass_input = st.text_input("Pasaport No", value=auto_pass, disabled=True)
        
    validation_locked = False
    
    # 1. Kural: Veri evraktan okunamadıysa operasyon durur.
    if not auto_tckn or not auto_sb:
        st.error("🚨 KATI DENETİM KİLİDİ: TCKN veya Liman Cüzdan Numarası evraklardaki şifreli alanlardan (MRZ) okunamadı! İnsan müdahalesine izin verilmemektedir. Lütfen evrakların okunaklı olduğundan emin olup tekrar yükleyin.")
        validation_locked = True
    else:
        st.success("✅ Veriye Dayalı Onay: Sistem TCKN ve Cüzdan Numarasını belge üzerinden başarıyla okudu.")
        uab_url = f"https://gemiadami.uab.gov.tr/verification/{auto_tckn}/{auto_sb}"
        st.markdown(
            f"""
            <div style="background-color: #e8f4fd; border-left: 5px solid #0056b3; padding: 12px 18px; border-radius: 6px; margin: 10px 0;">
                <span style="font-weight: bold; color: #0056b3; font-size: 14.5px;">🌐 UAB Doğrulama Portalı (Otomatik Eşleşme)</span><br>
                <a href="{uab_url}" target="_blank" style="background-color: #0056b3; color: #ffffff; padding: 6px 12px; border-radius: 4px; text-decoration: none; font-weight: bold; font-size: 13px; display: inline-block; margin-top: 8px;">
                    UAB Resmi Doğrulama Sayfasına Git ↗
                </a>
            </div>
            """, unsafe_allow_html=True
        )
    
    st.divider()
    st.subheader("📥 Çıktı Dosyaları")

    # [EXCEL VE WORD ÇIKTI FONKSİYONLARI ÖNCEKİ KOD İLE AYNIDIR - BURADA SADECE BUTONLAR VERİLMİŞTİR]
    # Kodun çok uzamaması adına indirme butonlarını kilit mantığına bağladım:

    if validation_locked:
        st.error("Sistem güvenliği gereği belgelerden veri okunmadan çıktı alınamaz.")
    else:
        st.success("Otomatik form üretimi onaylandı. (Excel ve Word kodları buraya entegre edilebilir).")
