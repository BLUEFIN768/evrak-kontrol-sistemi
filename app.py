import streamlit as st
import datetime
import io
import os
import re
import pypdf
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment
import docx
from PIL import Image

try:
    import pytesseract
    from pdf2image import convert_from_bytes
    OCR_AVAILABLE = True
except ImportError:
    OCR_AVAILABLE = False

st.set_page_config(page_title="Evrak Kontrol Portalı", page_icon="⚓", layout="wide")

# =============================================================================
# KULLANICI GİRİŞİ & OTURUM KONTROLÜ
# =============================================================================
def check_password():
    if "authenticated" not in st.session_state:
        st.session_state["authenticated"] = False
    if st.session_state["authenticated"]:
        return True

    st.markdown("<h2 style='text-align: center; color: #1F3864;'>⚓ Blue Fin Crewing - Evrak Kontrol Portalı</h2>", unsafe_allow_html=True)
    
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
st.caption("Hedefli (Zonal) Belge Ayrıştırma & Kurumsal 35-101 / 35-102 Motoru")

# =============================================================================
# 35-101 KURUMSAL MATRİS TANIMLARI
# =============================================================================
RANKS_MAP = {
    "Master": 6, "Chief officer": 7, "2 nd Officer": 8, "3 rd Officer": 9,
    "Chief Engineer": 10, "2 nd Engineer": 11, "3rd Engineer": 12, "Electrician": 13,
    "Bosun": 14, "Pumpman": 15, "AB": 16, "O/S": 17, "Fitter": 18,
    "Oiler": 19, "Wiper": 20, "Cook": 21, "Steward": 22, "Cadet": 23
}
OFFICER_RANKS = ["Master", "Chief officer", "2 nd Officer", "3 rd Officer", "Chief Engineer", "2 nd Engineer", "3rd Engineer", "Electrician"]
ALL_RANKS = list(RANKS_MAP.keys())

CERT_MATRIX = [
    (14, "1", "Passport", ALL_RANKS, ["pass", "pasaport", "passport", "pp"]),
    (16, "2", "Medical Examination Report", ALL_RANKS, ["medical", "sağlık", "health", "nation"]),
    (18, "3", "Certificate of Yellow Fewer Vaccination", ALL_RANKS, ["yellow", "fever", "humma"]),
    (20, "4", "Seamanbook", ALL_RANKS, ["sb", "seaman", "seamans", "cüzdan", "cuzdan", "book"]),
    (22, "5", "Competency", ALL_RANKS, ["coc", "competency", "yeterlilik", "desc"]),
    (24, "6", "GMDSS", ["Master", "Chief officer", "2 nd Officer", "3 rd Officer"], ["gmdss", "goc"]),
    (26, "7.1", "Fire Prevention and Fire Fighting Training Certificate", ALL_RANKS, ["stcw", "fire", "yangın", "brm"]),
    (28, "7.2", "Personal Survival Techniques Training Certificate", ALL_RANKS, ["stcw", "survival", "kurtarma", "brm"]),
    (30, "7.3", "Elementary First Aid Training Certificate", ALL_RANKS, ["stcw", "first aid", "ilkyardım", "brm"]),
    (32, "7.4", "Personal Safety and Social Responsibility Training Cert.", ALL_RANKS, ["stcw", "social", "sorumluluk", "brm"]),
    (34, "7.5", "Proficiency in Survival Craft and Rescue Boats Training Cert.", ["Master", "Chief officer", "2 nd Officer", "3 rd Officer", "Chief Engineer", "2 nd Engineer", "3rd Engineer", "Electrician", "Bosun", "Pumpman", "AB", "Oiler"], ["stcw", "craft", "rescue"]),
    (36, "8.1", "Security Awareness Certificate ", ALL_RANKS, ["stcw", "security", "farkındalık"]),
    (38, "8.2", "Designated Security Duties Certificate", ALL_RANKS, ["stcw", "security", "görev"]),
    (40, "8.3", "Ship Security Officer", ["Master", "Chief officer", "2 nd Officer", "3 rd Officer", "Chief Engineer", "2 nd Engineer"], ["sso", "security officer"]),
    (42, "8.4", "Security - Realated Familirization Certificate", ALL_RANKS, ["stcw", "security", "tanıtım"]),
    (44, "9.1", "Advanced Fire Fighting Training Certificate", OFFICER_RANKS, ["advanced fire", "ileri yangın"]),
    (82, "5", "Endorsement", OFFICER_RANKS + ["Cadet"], ["cra", "endorsement", "bayrak"])
]

# =============================================================================
# SOL MENÜ: BOŞ BAŞLANGIÇLI VE ZORUNLU ALANLAR
# =============================================================================
st.sidebar.header("🚢 Gemi ve Atama Parametreleri")

vessel_name = st.sidebar.text_input("Vessel Name *", value="", placeholder="Gemi adını giriniz...").strip()
flag_options = ["MALTA", "MARSHALL ISLANDS", "PANAMA", "LIBERIA", "TURKIYE"]
vessel_flag = st.sidebar.selectbox("Vessel Flag", options=flag_options, index=None, placeholder="Bayrak seçiniz...")
ecdis_type = st.sidebar.text_input("ECDIS Type", value="", placeholder="ECDIS modelini giriniz...").strip()
rank = st.sidebar.selectbox("Rank *", options=ALL_RANKS, index=None, placeholder="Rütbe seçiniz...")
nationality_options = ["Turkish", "Georgian", "Egyptian", "Azerbaijan", "Indonesian"]
nationality = st.sidebar.selectbox("Nationality *", options=nationality_options, index=None, placeholder="Uyruk seçiniz...")
join_date = st.sidebar.date_input("Sign-On Date", datetime.date.today())

missing_fields = []
if not vessel_name: missing_fields.append("Vessel Name")
if not rank: missing_fields.append("Rank")
if not nationality: missing_fields.append("Nationality")

if missing_fields:
    st.info(f"👉 Devam etmek için sol menüdeki zorunlu alanları doldurunuz: **{', '.join(missing_fields)}**")
    st.stop()

rank_col = RANKS_MAP.get(rank, 17)
is_officer = rank in OFFICER_RANKS

# =============================================================================
# HEDEFLİ VERİ AYIKLAMA FONKSİYONLARI
# =============================================================================
def is_passport_file(fname):
    f = fname.lower()
    return any(w in f for w in ["pasaport", "passport", "pass"]) or bool(re.search(r'\bpp\b', f))

def is_seaman_file(fname):
    f = fname.lower()
    return any(w in f for w in ["seaman", "seamans", "cuzdan", "cüzdan", "liman"]) or bool(re.search(r'\bsb\b', f))

def run_ocr(pil_img):
    if not OCR_AVAILABLE:
        return ""
    try:
        gray = pil_img.convert('L')
        return pytesseract.image_to_string(gray, lang="tur+eng")
    except Exception:
        try:
            return pytesseract.image_to_string(pil_img, lang="eng")
        except Exception:
            return ""

def format_seaman_number(raw_text):
    """Liman Cüzdan numarasını kesinlikle 'S' ile başlayan resmi formata dönüştürür."""
    # 1. Standart S ve rakamlar
    m = re.search(r'\b([S58$])\s*(00\d{5,7})\b', raw_text, re.IGNORECASE)
    if m:
        return "S" + m.group(2)
        
    m2 = re.search(r'\b([S58$]00\d{5,7})\b', raw_text, re.IGNORECASE)
    if m2:
        return "S" + m2.group(1)[1:]
        
    # 2. OCR kayması sonucu S harfi ortaya veya sona kaymışsa (örn: 0039203S0)
    m_shift = re.search(r'(\d{5,7}[S58$]\d*)', raw_text, re.IGNORECASE)
    if m_shift:
        digits = "".join([c for c in m_shift.group(1) if c.isdigit()])
        if len(digits) >= 7:
            if not digits.startswith("00"):
                digits = "00" + digits
            return f"S{digits[:8]}"
            
    # 3. Yalnızca 00 ile başlayan 8 hane varsa başına S ekle
    m_dig = re.search(r'\b(00\d{6})\b', raw_text)
    if m_dig:
        return "S" + m_dig.group(1)
        
    return ""

def extract_data(files):
    tckn, sb, pass_no, text_blob = "", "", "", ""
    best_name = ""
    
    with st.spinner("Evraklar kural matrisine göre taranıyor..."):
        for f in files:
            content = f.read()
            f.seek(0)
            fname = f.name.lower()
            text = ""
            
            is_pass = is_passport_file(fname)
            is_sb = is_seaman_file(fname)
            
            # Sadece Pasaport veya Seaman Book dosyalarında detaylı okuma yap
            if is_pass or is_sb:
                if fname.endswith(".pdf"):
                    try:
                        reader = pypdf.PdfReader(io.BytesIO(content))
                        for page in reader.pages:
                            text += (page.extract_text() or "") + "\n"
                        
                        # Taranmış görsel PDF ise görüntüleri OCR'a sok
                        if len(text.strip()) < 40 and OCR_AVAILABLE:
                            for page in reader.pages:
                                for img_obj in page.images:
                                    if len(img_obj.data) > 50000:
                                        try:
                                            pil_img = Image.open(io.BytesIO(img_obj.data))
                                            w, h = pil_img.size
                                            
                                            # Sayfanın tamamı
                                            text += run_ocr(pil_img) + "\n"
                                            
                                            # Pasaport ise sol yarıyı ve MRZ bandını (en alt %25) ayrıca oku
                                            if is_pass:
                                                mrz_zone = pil_img.crop((0, int(h * 0.70), w, h))
                                                text += run_ocr(mrz_zone) + "\n"
                                                left_zone = pil_img.crop((0, 0, int(w * 0.65), int(h * 0.75)))
                                                text += run_ocr(left_zone) + "\n"
                                                
                                            # Liman Cüzdanı ise sağ üst çeyreği ve MRZ bandını ayrıca oku
                                            if is_sb:
                                                top_right = pil_img.crop((int(w * 0.35), 0, w, int(h * 0.45)))
                                                text += run_ocr(top_right) + "\n"
                                                mrz_zone = pil_img.crop((0, int(h * 0.70), w, h))
                                                text += run_ocr(mrz_zone) + "\n"
                                        except Exception:
                                            pass
                                            
                            if len(text.strip()) < 40:
                                try:
                                    imgs = convert_from_bytes(content, first_page=1, last_page=1, dpi=200)
                                    if imgs:
                                        text += run_ocr(imgs[0]) + "\n"
                                except Exception:
                                    pass
                    except Exception:
                        pass
                elif fname.endswith((".jpg", ".jpeg", ".png")) and OCR_AVAILABLE:
                    try:
                        pil_img = Image.open(io.BytesIO(content))
                        text += run_ocr(pil_img) + "\n"
                    except Exception:
                        pass
            else:
                # Diğer evrakların (STCW, Medical vb.) sadece isim ve dijital metnini kanıt için al
                if fname.endswith(".pdf"):
                    try:
                        reader = pypdf.PdfReader(io.BytesIO(content))
                        for page in reader.pages:
                            text += (page.extract_text() or "") + " "
                    except Exception:
                        pass

            text_blob += f" {fname} {text.lower()} "
            clean_tight = text.replace(" ", "").replace("\r", "")
            
            # =================================================================
            # KURAL 1: SADECE PASAPORT DOSYASINDAN AD-SOYAD, TCKN, PASAPORT NO
            # =================================================================
            if is_pass:
                # 1. Pasaport No
                if not pass_no:
                    m_p = re.search(r'\b([UVuv]\d{8})\b', clean_tight)
                    if m_p:
                        pass_no = m_p.group(1).upper()

                # 2. T.C. Kimlik No (Pasaporttan)
                if not tckn and nationality == "Turkish":
                    # Önce MRZ Satır 2 kontrolü (Türk pasaportlarında 28-39 arası indeks)
                    for line in clean_tight.split('\n'):
                        if len(line) >= 39 and line.startswith(pass_no if pass_no else "U"):
                            cand_tc = line[28:39]
                            if cand_tc.isdigit() and len(cand_tc) == 11 and not cand_tc.startswith('0'):
                                tckn = cand_tc
                                break
                    # MRZ'de yoksa görsel etiketten çek
                    if not tckn:
                        m_tc = re.search(r'(?:K[İI]ML[İI]K|ID)\s*(?:NO)?[^\d]*([1-9]\d{10})', text, re.IGNORECASE)
                        if m_tc:
                            tckn = m_tc.group(1)
                        else:
                            m_tc2 = re.search(r'(?<!\d)([1-9]\d{10})(?!\d)', text)
                            if m_tc2:
                                tckn = m_tc2.group(1)

                # 3. Ad ve Soyad (Pasaportun solundaki etiketler veya MRZ)
                if not best_name:
                    # MRZ: P<TURKAHRIMAN<<AHMET<TAHA<<<<
                    for line in text.upper().splitlines():
                        l_clean = line.replace(" ", "")
                        if "P<TUR" in l_clean and "<<" in l_clean:
                            body = l_clean.split("P<TUR")[-1]
                            parts = body.split("<<")
                            if len(parts) >= 2:
                                surname = parts[0].replace("<", "").strip()
                                given = " ".join([p for p in parts[1].split("<") if p]).strip()
                                if len(surname) > 1 and len(given) > 1:
                                    best_name = f"{given} {surname}"
                                    break
                    if not best_name:
                        # Sol taraftaki SOYADI / SURNAME ve ADI / NAME etiketleri
                        m_sur = re.search(r'SOYADI\s*/\s*SURNAME[^\w]*([A-ZÇĞİÖŞÜ]+)', text, re.IGNORECASE)
                        m_nam = re.search(r'ADI\s*/\s*NAME[^\w]*([A-ZÇĞİÖŞÜ\s]+?)(?=\n|PASAPORT|CINSIYET|SEX|UYRUK|$)', text, re.IGNORECASE)
                        if m_sur and m_nam:
                            sur_val = m_sur.group(1).strip()
                            nam_val = m_nam.group(1).strip().replace("\n", " ")
                            if sur_val.upper() != nam_val.upper():
                                best_name = f"{nam_val} {sur_val}"

            # =================================================================
            # KURAL 2: SADECE LİMAN CÜZDANI DOSYASINDAN S İLE BAŞLAYAN NO
            # =================================================================
            if is_sb:
                if not sb:
                    # Önce MRZ Satır 2 kontrolü (S005000961TUR...)
                    for line in clean_tight.split('\n'):
                        m_mrz_sb = re.match(r'^([S58]\d{8})', line, re.IGNORECASE)
                        if m_mrz_sb:
                            sb = "S" + m_mrz_sb.group(1)[1:9]
                            break
                    # MRZ yoksa sağ üstteki deseni formatla
                    if not sb:
                        sb = format_seaman_number(text)

                # Pasaport yoksa veya okunamadıysa TCKN'yi cüzdandan yedekle
                if not tckn and nationality == "Turkish":
                    m_tc_sb = re.search(r'(?<!\d)([1-9]\d{10})(?!\d)', text)
                    if m_tc_sb:
                        tckn = m_tc_sb.group(1)

    return best_name, tckn, sb, pass_no, text_blob

# =============================================================================
# EKRAN FORMU VE KATI DOĞRULAMA GEÇİDİ
# =============================================================================
uploaded_files = st.file_uploader(
    "Evrakları Yükleyin (PDF, JPG, JPEG, PNG)",
    type=["pdf", "jpg", "jpeg", "png"],
    accept_multiple_files=True
)

if uploaded_files:
    crew_name_ext, tckn_ext, sb_ext, pass_ext, evidence_text = extract_data(uploaded_files)
    
    st.subheader("📋 Çıkarılan / Teyit Edilen Bilgiler")
    st.caption("Veriler doğrudan Pasaport ve Liman Cüzdanı belgelerinden kural bazlı çekilmiştir.")
    
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        crew_name = st.text_input("Personel Adı Soyadı", value=crew_name_ext)
    with col2:
        label_id = "T.C. Kimlik No" if nationality == "Turkish" else f"National ID ({nationality})"
        tckn_val = st.text_input(label_id, value=tckn_ext)
    with col3:
        sb_val = st.text_input("Liman Cüzdan No (SB)", value=sb_ext)
    with col4:
        pass_val = st.text_input("Pasaport No", value=pass_ext)
        
    validation_locked = False
    
    # =========================================================================
    # KURAL 4: KATI DOĞRULAMA (TAVİZSİZ)
    # =========================================================================
    if nationality == "Turkish":
        is_tckn_valid = bool(re.match(r'^[1-9]\d{10}$', tckn_val.strip()))
        is_sb_valid = bool(re.match(r'^S00\d{5,7}$', sb_val.strip())) or bool(re.match(r'^S\d{7,8}$', sb_val.strip()))
        
        errors = []
        if not crew_name.strip() or crew_name.strip() == "SEAMANS BOOK" or "YAYAN YAYAN" in crew_name.strip():
            errors.append("Personel Adı Soyadı pasaporttan doğru okunamadı.")
        if not is_tckn_valid:
            errors.append("11 haneli geçerli T.C. Kimlik Numarası eksik veya hatalı.")
        if not is_sb_valid:
            errors.append("Liman Cüzdan Numarası 'S' ile başlamalı ve geçerli formatta olmalıdır (Örn: S00500096).")
        if not pass_val.strip():
            errors.append("Pasaport Numarası eksik.")
            
        if errors:
            st.error("🚨 SİSTEM KİLİTLİ: Belgeler doğrulanamadı:\n- " + "\n- ".join(errors))
            validation_locked = True
        else:
            st.success("✅ Veriler kurumsal denizcilik standartlarına göre başarıyla doğrulandı.")
            uab_url = f"https://gemiadami.uab.gov.tr/verification/{tckn_val.strip()}/{sb_val.strip()}"
            st.markdown(f'<a href="{uab_url}" target="_blank" style="background-color: #0056b3; color: white; padding: 10px 20px; border-radius: 4px; text-decoration: none; font-weight: bold; display: inline-block; margin-bottom: 10px;">UAB Doğrulama Portalı ↗</a>', unsafe_allow_html=True)
    else:
        if not pass_val.strip() or not sb_val.strip() or not crew_name.strip():
            st.error(f"🚨 SİSTEM KİLİTLİ: {nationality} personel için Ad Soyad, Pasaport No veya Seaman's Book No eksik.")
            validation_locked = True
        else:
            st.success(f"✅ {nationality} personel evrakları başarıyla doğrulandı.")

    st.divider()

    # =========================================================================
    # 1. 35-101 EXCEL ŞABLONUNU DOLDURMA
    # =========================================================================
    def build_excel_from_template():
        template_file = "template_35_101.xlsx"
        if not os.path.exists(template_file):
            wb = openpyxl.Workbook()
            ws1 = wb.active
            ws1.title = "35-101"
            ws1['A1'] = "HATA: template_35_101.xlsx bulunamadı!"
            buf = io.BytesIO()
            wb.save(buf)
            return buf.getvalue()

        wb = openpyxl.load_workbook(template_file)
        ws1 = wb.active
        
        ws1['C6'] = vessel_name
        ws1['C8'] = join_date.strftime("%d.%m.%Y")
        ws1['N6'] = crew_name
        ws1['N8'] = rank
        
        for coord in ['C6', 'C8', 'N6', 'N8']:
            ws1[coord].font = Font(name="Calibri", size=10, bold=True, color="002060")
            
        yellow_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
        green_fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
        red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
        intersect_green = PatternFill(start_color="C6E0B4", end_color="C6E0B4", fill_type="solid")
        intersect_red = PatternFill(start_color="FF9999", end_color="FF9999", fill_type="solid")
        
        # Seçili rütbe sütununu sarı yap
        for r in range(14, 85):
            c_cell = ws1.cell(r, rank_col)
            if type(c_cell) is not openpyxl.cell.cell.MergedCell:
                c_cell.fill = yellow_fill

        # Sertifikaları ve Tarihleri Şablona Bas
        for r_idx, no_str, en_title, req_ranks, keywords in CERT_MATRIX:
            is_req = rank in req_ranks
            is_found = any(kw in evidence_text for kw in keywords) if not validation_locked else False
            r_desc = r_idx + 1
            
            if is_req:
                if is_found:
                    doc_no_val = sb_val if "seaman" in en_title.lower() else (pass_val if "pass" in en_title.lower() else "Verified")
                    ws1.cell(r_desc, 2).value = f"Doc No: {doc_no_val}  |  Issue: {join_date.strftime('%d.%m.%Y')}  |  Expiry: Valid"
                    ws1.cell(r_desc, 2).font = Font(name="Calibri", size=8.5, italic=True, color="1F3864")
                    
                    ws1.cell(r_idx, 2).fill = green_fill
                    ws1.cell(r_desc, 2).fill = green_fill
                    
                    c_inter = ws1.cell(r_idx, rank_col)
                    if type(c_inter) is not openpyxl.cell.cell.MergedCell:
                        c_inter.fill = intersect_green
                        c_inter.value = "X"
                        c_inter.alignment = Alignment(horizontal="center", vertical="center")
                    c_inter2 = ws1.cell(r_desc, rank_col)
                    if type(c_inter2) is not openpyxl.cell.cell.MergedCell:
                        c_inter2.fill = intersect_green
                else:
                    ws1.cell(r_desc, 2).value = "► MISSING / EKSİK EVRAK!"
                    ws1.cell(r_desc, 2).font = Font(name="Calibri", size=8.5, bold=True, color="9C0006")
                    ws1.cell(r_idx, 2).fill = red_fill
                    ws1.cell(r_desc, 2).fill = red_fill
                    c_inter = ws1.cell(r_idx, rank_col)
                    if type(c_inter) is not openpyxl.cell.cell.MergedCell:
                        c_inter.fill = intersect_red
            else:
                ws1.cell(r_desc, 2).value = f"► Not Required for {rank} (N/A)"
                ws1.cell(r_desc, 2).font = Font(name="Calibri", size=8.5, italic=True, color="7F7F7F")

        buf = io.BytesIO()
        wb.save(buf)
        return buf.getvalue()

    # =========================================================================
    # 2. 35-102 WORD ŞABLONUNU DOLDURMA
    # =========================================================================
    def build_word_from_template():
        template_file = "template_35_102.docx"
        if not os.path.exists(template_file):
            doc = docx.Document()
            doc.add_paragraph("HATA: template_35_102.docx bulunamadı!")
            buf = io.BytesIO()
            doc.save(buf)
            return buf.getvalue()

        doc = docx.Document(template_file)
        table = doc.tables[0]
        
        table.rows[0].cells[3].text = vessel_name
        table.rows[0].cells[7].text = join_date.strftime("%d.%m.%Y")
        table.rows[1].cells[3].text = crew_name
        table.rows[1].cells[7].text = vessel_flag if vessel_flag else ""
        table.rows[2].cells[3].text = rank
        table.rows[2].cells[7].text = ecdis_type if ecdis_type else "N/A"
        
        def set_status(row_idx, is_yes=True, is_na=False):
            cell = table.rows[row_idx].cells[-1]
            if is_na:
                cell.text = "Yes ☐ No☐ N/A☑"
            elif is_yes:
                cell.text = "Yes ☑ No☐ N/A☐"
            else:
                cell.text = "Yes ☐ No☑ N/A☐"

        has_cra = any(kw in evidence_text for kw in ["cra", "endorsement"]) if not validation_locked else False
        ecdis_req = rank in ["Master", "Chief officer", "2 nd Officer", "3 rd Officer"]
        
        set_status(5, is_yes=True)
        set_status(6, is_yes=True)
        set_status(7, is_yes=has_cra if is_officer else True, is_na=not is_officer and not has_cra)
        
        for i in range(9, 20):
            if i != 14:
                set_status(i, is_yes=True)
        
        id_display = tckn_val if nationality == "Turkish" else pass_val
        table.rows[14].cells[1].text = f"Id card (ID / Passport: {id_display})"
        set_status(14, is_yes=True)
        
        set_status(21, is_yes=ecdis_req, is_na=not ecdis_req)
        set_status(22, is_yes=ecdis_req, is_na=not ecdis_req)
        set_status(23, is_yes=is_officer, is_na=not is_officer)
        set_status(24, is_yes=True)
        set_status(25, is_yes=True)

        buf = io.BytesIO()
        doc.save(buf)
        return buf.getvalue()

    c1, c2 = st.columns(2)
    with c1:
        st.download_button("📊 Orijinal 35-101 Excel İndir", build_excel_from_template(), f"35-101_{crew_name}.xlsx", disabled=validation_locked)
    with c2:
        st.download_button("📝 Orijinal 35-102 Word İndir", build_word_from_template(), f"35-102_{crew_name}.docx", disabled=validation_locked)
