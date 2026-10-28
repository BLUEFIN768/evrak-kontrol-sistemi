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

# NOKTA ATIŞI OCR (Görselden Okuma) ALTYAPISI
try:
    import pytesseract
    from pdf2image import convert_from_bytes
    from PIL import Image
    OCR_AVAILABLE = True
except ImportError:
    OCR_AVAILABLE = False

st.set_page_config(page_title="Evrak Kontrol Portalı", page_icon="⚓", layout="wide")

# ==============================================================================
# GÜVENLİK PROTOKOLÜ (KVKK)
# ==============================================================================
def check_password():
    if "authenticated" not in st.session_state:
        st.session_state["authenticated"] = False
    if st.session_state["authenticated"]:
        return True

    st.markdown("<h2 style='text-align: center; color: #1F3864;'>⚓ Blue Fin Crewing - Evrak Kontrol Portalı</h2>", unsafe_allow_html=True)
    st.markdown("<p style='text-align: center; color: #666;'>Katı Kural: Liman cüzdanı yüklenmeden ve doğrulanmadan evraklar onaylanmaz.</p>", unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns([1, 1.5, 1])
    with col2:
        with st.form("login_form"):
            username = st.text_input("Yetkili Kullanıcı")
            password = st.text_input("Güvenlik Parolası", type="password", help="bluefin / bluefin2026")
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
st.caption("Cüzdan Odaklı Otomatik OCR & UAB Doğrulama Sistemi")

if not OCR_AVAILABLE:
    st.error("🚨 KILIT HATA: OCR (Görsel Okuma) Motoru Başlatılamadı! Lütfen GitHub 'packages.txt' ve 'requirements.txt' dosyalarının tam olduğundan emin olun.")

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
    (16, "2", "Medical Examination Report", ALL_RANKS, ["medical", "sağlık", "health"]),
    (18, "3", "Certificate of Yellow Fever", ALL_RANKS, ["yellow", "fever", "humma", "aşı"]),
    (20, "4", "Seamanbook", ALL_RANKS, ["sb", "seaman", "cüzdan"]),
    (22, "5", "Competency", ALL_RANKS, ["coc", "competency", "yeterlilik"]),
    (24, "6", "GMDSS", ["Master", "Chief officer", "2 nd Officer", "3 rd Officer"], ["gmdss", "goc"]),
    (26, "7.1", "Fire Prevention and Fire Fighting", ALL_RANKS, ["stcw", "courses", "fire", "yangın"]),
    (28, "7.2", "Personal Survival Techniques", ALL_RANKS, ["stcw", "courses", "survival", "can kurtarma"]),
    (30, "7.3", "Elementary First Aid", ALL_RANKS, ["stcw", "courses", "first aid", "ilkyardım"]),
    (32, "7.4", "Personal Safety and Social Responsibility", ALL_RANKS, ["stcw", "courses", "social", "sorumluluk"]),
    (34, "7.5", "Proficiency in Survival Craft and Rescue Boats", ["Master", "Chief officer", "2 nd Officer", "3 rd Officer", "Chief Engineer", "2 nd Engineer", "3rd Engineer", "Electrician", "Bosun", "Pumpman", "AB", "Oiler"], ["stcw", "courses", "craft", "rescue"]),
    (36, "8.1", "Security Awareness Certificate", ALL_RANKS, ["stcw", "courses", "security", "farkındalık"]),
    (38, "8.2", "Designated Security Duties Certificate", ALL_RANKS, ["stcw", "courses", "security", "görev"]),
    (42, "8.4", "Security - Realated Familirization", ALL_RANKS, ["stcw", "courses", "security", "tanıtım"]),
    (82, "5", "Endorsement (Bayrak Onayı)", OFFICER_RANKS + ["Deck Cadet", "Engine Cadet"], ["cra", "endorsement", "bayrak"])
]

st.sidebar.header("🚢 Gemi ve Atama Parametreleri")
vessel_name = st.sidebar.text_input("Gemi Adı", value="Optimus S")
vessel_flag = st.sidebar.selectbox("Gemi Bayrağı", ["MALTA", "MARSHALL ISLANDS", "PANAMA", "LIBERIA", "TURKIYE"])
ecdis_model = st.sidebar.text_input("Gemideki ECDIS Modeli", value="")
rank = st.sidebar.selectbox("Atanan Rütbe", options=ALL_RANKS, index=17) 
rank_col = RANKS_MAP[rank]
join_date = st.sidebar.date_input("Katılış Tarihi", datetime.date.today())

# ==============================================================================
# HEDEFLİ OCR VE KESİN CÜZDAN BULUCU MOTOR
# ==============================================================================
def extract_evidence(files):
    best_name = ""
    tckn = ""
    sb = ""
    all_text_blob = ""
    
    ignore_words = {"deck", "engine", "cadet", "pass", "medical", "sb", "coc", "desc", "criminal", "record", 
                    "identity", "register", "military", "status", "stcw", "yellow", "fever", "courses", 
                    "sea", "service", "document", "guidance", "check", "list", "pdf", "jpg", "jpeg", "png", "xlsx", "xls"}
    
    # 1. DOSYA ADLARINDAN PERSONEL İSMİ BULMA
    for f in files:
        fname = f.name.lower()
        clean_name = re.sub(r'[\d\W_]+', ' ', fname)
        words = clean_name.split()
        person_words = [w for w in words if w not in ignore_words and len(w) > 1]
        if len(person_words) >= 2 and not best_name:
            best_name = " ".join(person_words).upper()
            
    # 2. DİJİTAL METİNLERİ (TCKN vb.) HIZLICA TARA
    for f in files:
        fname = f.name.lower()
        all_text_blob += f" {fname} "
        if fname.endswith(".pdf"):
            try:
                content = f.read()
                f.seek(0)
                reader = pypdf.PdfReader(io.BytesIO(content))
                for page in reader.pages:
                    text = page.extract_text() or ""
                    all_text_blob += f" {text.lower()} "
                    if not tckn:
                        m_tc = re.search(r'\b([1-9][0-9]{10})\b', text)
                        if m_tc: tckn = m_tc.group(1)
            except Exception:
                pass
                
    # 3. KESİN KURAL: LİMAN CÜZDANINI NOKTA ATIŞI GÖRSELDEN OKU (Sadece "sb", "seaman", "cüzdan" geçen dosyalar)
    if not sb and OCR_AVAILABLE:
        for f in files:
            fname = f.name.lower()
            if any(k in fname for k in ["sb", "seaman", "cüzdan"]):
                try:
                    content = f.read()
                    f.seek(0)
                    extracted_ocr_text = ""
                    
                    if fname.endswith(".pdf"):
                        # Sadece ilk sayfayı resme çevir (Kilitlenmeyi önler)
                        images = convert_from_bytes(content, first_page=1, last_page=1)
                        if images:
                            extracted_ocr_text = pytesseract.image_to_string(images[0], lang="eng")
                    elif fname.endswith((".jpg", ".jpeg", ".png")):
                        img = Image.open(io.BytesIO(content))
                        extracted_ocr_text = pytesseract.image_to_string(img, lang="eng")
                        
                    # Bulunan görsel metin içerisinde Türkiye SB (S00...) numarasını ara
                    if extracted_ocr_text:
                        m_s = re.search(r'\b(S00\d{5,7}|S0\d{6,8}|S\d{7,8})\b', extracted_ocr_text, re.IGNORECASE)
                        if m_s:
                            sb = m_s.group(1).upper()
                            break # Numarayı bulduktan sonra döngüyü sonlandır
                except Exception:
                    pass

    return best_name, tckn, sb, all_text_blob

# Arayüz Yükleyici
uploaded_files = st.file_uploader("Personel Evraklarını Yükleyin (SB, Pasaport vb.)", type=["pdf", "jpg", "jpeg", "png", "xlsx"], accept_multiple_files=True)

if uploaded_files:
    auto_name, auto_tckn, auto_sb, evidence_text = extract_evidence(uploaded_files)
    st.success("✅ Belgeler tarandı. TCKN ve Cüzdan numarası ayıklandı.")
    
    col_info1, col_info2, col_info3 = st.columns(3)
    with col_info1:
        crew_name = st.text_input("Personel Adı Soyadı", value=auto_name)
    with col_info2:
        tckn_input = st.text_input("T.C. Kimlik No", value=auto_tckn)
    with col_info3:
        sb_input = st.text_input("Liman Cüzdan No (SB)", value=auto_sb, placeholder="Zorunlu Alan")
        
    clean_tckn = re.sub(r'\D', '', tckn_input)
    clean_sb = re.sub(r'[^A-Za-z0-9]', '', sb_input).upper()
    
    validation_locked = False
    if not clean_tckn or not clean_sb:
        st.error("🚨 **KATI DENETİM KİLİDİ:** TCKN veya Liman Cüzdan Numarası eksik/okunamadı! Evrak onayları kilitlendi.")
        validation_locked = True
    else:
        uab_url = f"https://gemiadami.uab.gov.tr/verification/{clean_tckn}/{clean_sb}"
        st.markdown(
            f"""
            <div style="background-color: #e8f4fd; border-left: 5px solid #0056b3; padding: 12px 18px; border-radius: 6px; margin: 10px 0;">
                <div style="display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap;">
                    <div>
                        <span style="font-weight: bold; color: #0056b3; font-size: 14.5px;">🌐 UAB Doğrulama Portalı Hazır</span><br>
                        <span style="font-size: 13px; color: #333;">TCKN: <b>{clean_tckn}</b> | Cüzdan No: <b>{clean_sb}</b></span>
                    </div>
                    <a href="{uab_url}" target="_blank" style="background-color: #0056b3; color: #ffffff; padding: 8px 16px; border-radius: 4px; text-decoration: none; font-weight: bold; font-size: 13px;">
                        Doğrulama Sayfasına Git ↗
                    </a>
                </div>
            </div>
            """, unsafe_allow_html=True
        )
    
    st.divider()
    st.subheader("📥 Çıktı Dosyaları")

    def build_excel_export():
        template_candidates = ["35-101_Seafarers_Documents_Check_List_YASIN_DOSER_V5.xlsx", "35-101_template.xlsx", "templates/35-101_template.xlsx"]
        found_template = next((c for c in template_candidates if os.path.exists(c)), None)
        
        if found_template:
            wb = openpyxl.load_workbook(found_template)
            ws1 = wb["35-101 Check List"] if "35-101 Check List" in wb.sheetnames else wb.active
        else:
            wb = openpyxl.Workbook()
            ws1 = wb.active
            ws1.title = "35-101 Check List"
            
        ws1['C6'] = vessel_name; ws1['C8'] = join_date.strftime("%d.%m.%Y")
        ws1['N6'] = crew_name; ws1['N8'] = rank
        ws1.cell(9, rank_col).value = rank
        
        yellow_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
        green_fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
        red_fill = PatternFill(start_color="FF4D4D", end_color="FF4D4D", fill_type="solid")
        intersect_green = PatternFill(start_color="C6E0B4", end_color="C6E0B4", fill_type="solid")
        intersect_red = PatternFill(start_color="F8CBAD", end_color="F8CBAD", fill_type="solid")
        
        for r in range(14, 84): ws1.cell(r, rank_col).fill = yellow_fill
            
        for r_idx, no_str, en_title, req_ranks, keywords in CERT_MATRIX:
            is_req = rank in req_ranks
            is_found = any(kw in evidence_text for kw in keywords) if not validation_locked else False
            r2 = r_idx + 1
            ws1.cell(r_idx, 1).value = no_str
            ws1.cell(r_idx, 2).value = en_title
            
            if is_req:
                if is_found:
                    ws1.cell(r2, 2).value = "► Mevcut (Evrak Yüklendi)"
                    for r_sub in [r_idx, r2]:
                        for c in range(1, rank_col): ws1.cell(r_sub, c).fill = green_fill
                        ws1.cell(r_sub, rank_col).fill = intersect_green
                        ws1.cell(r_sub, rank_col).value = "X"
                else:
                    ws1.cell(r2, 2).value = "► EKSİK / DOSYADA BULUNAMADI!"
                    for r_sub in [r_idx, r2]:
                        for c in range(1, rank_col): ws1.cell(r_sub, c).fill = red_fill
                        ws1.cell(r_sub, rank_col).fill = intersect_red
            else:
                ws1.cell(r2, 2).value = f"► Bu görev ({rank}) için zorunlu değil (N/A)"

        ws1.page_setup.orientation = ws1.ORIENTATION_LANDSCAPE
        ws1.page_setup.fitToWidth = 1
        ws1.page_setup.fitToHeight = 1
        
        buf = io.BytesIO()
        wb.save(buf)
        return buf.getvalue()

    def build_word_export():
        doc = docx.Document()
        doc.add_paragraph("35-102 PRE DESIGNATION CHECK LIST")
        table = doc.add_table(rows=1, cols=3)
        
        def add_check_row(no_str, title_str, is_yes=True, is_na=False):
            row = table.add_row()
            row.cells[0].text = no_str
            row.cells[1].text = title_str
            if is_na: row.cells[2].text = "N/A ☑"
            elif is_yes: row.cells[2].text = "Yes ☑"
            else: row.cells[2].text = "No ☑"

        has_cra = any(kw in evidence_text for kw in ["cra", "endorsement", "bayrak"]) if not validation_locked else False
        add_check_row("3", "Flag State’s Licence / STCW Endorsement / CRA", is_yes=has_cra)
        
        buf = io.BytesIO()
        doc.save(buf)
        return buf.getvalue()

    c1, c2 = st.columns(2)
    with c1:
        st.download_button(label=f"📊 35-101 Excel İndir ({crew_name})", data=build_excel_export(), file_name=f"35-101_{crew_name}.xlsx", disabled=validation_locked)
    with c2:
        st.download_button(label=f"📝 35-102 Word İndir ({crew_name})", data=build_word_export(), file_name=f"35-102_{crew_name}.docx", disabled=validation_locked)
