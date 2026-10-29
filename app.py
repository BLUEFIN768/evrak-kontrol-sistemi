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
# 35-101 MATRİS TANIMLARI
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
    (14, "1", "Passport", ALL_RANKS, ["pass", "pasaport", "pp"]),
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
# SOL MENÜ: BOŞ VE ZORUNLU PARAMETRELER
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
# HEDEFLİ (ZONAL) VERİ AYIKLAMA MOTORU
# =============================================================================
def clean_sb(raw_val):
    val = re.sub(r'[^a-zA-Z0-9]', '', raw_val).upper()
    val = val.replace('5', 'S', 1).replace('8', 'S', 1).replace('$', 'S', 1)
    if val.startswith('S') and not val.startswith('S00'):
        val = val.replace('S', 'S00', 1)
    return val

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

def extract_data(files):
    tckn, sb, pass_no, text_blob = "", "", "", ""
    best_name = ""
    
    with st.spinner("Evraklar hedeflenerek taranıyor..."):
        for f in files:
            content = f.read()
            f.seek(0)
            fname = f.name.lower()
            text = ""
            
            is_passport_file = any(k in fname for k in ["pass", "pasaport", "pp"])
            is_seaman_file = any(k in fname for k in ["sb", "seaman", "seamans", "cüzdan", "cuzdan"])
            
            # PDF Okuma
            if fname.endswith(".pdf"):
                try:
                    reader = pypdf.PdfReader(io.BytesIO(content))
                    for page in reader.pages:
                        text += (page.extract_text() or "") + " "
                    
                    # Taranmış PDF ise ve hedef dosyalardan biriyse OCR çalıştır
                    if len(text.strip()) < 40 and OCR_AVAILABLE and (is_passport_file or is_seaman_file):
                        for page in reader.pages:
                            for img_obj in page.images:
                                # CamScanner küçük filigran şeritlerini atla (50KB altı)
                                if len(img_obj.data) > 50000:
                                    try:
                                        pil_img = Image.open(io.BytesIO(img_obj.data))
                                        text += run_ocr(pil_img) + " "
                                    except Exception:
                                        pass
                        if len(text.strip()) < 40:
                            try:
                                imgs = convert_from_bytes(content, first_page=1, last_page=1, dpi=200)
                                if imgs:
                                    text += run_ocr(imgs[0]) + " "
                            except Exception:
                                pass
                except Exception:
                    pass
            
            # Resim Dosyaları (.jpg, .jpeg, .png)
            elif fname.endswith((".jpg", ".jpeg", ".png")) and (is_passport_file or is_seaman_file):
                if OCR_AVAILABLE:
                    try:
                        pil_img = Image.open(io.BytesIO(content))
                        text += run_ocr(pil_img) + " "
                    except Exception:
                        pass
            
            text_blob += f" {fname} {text.lower()} "
            clean_search = text.replace("<", " ").replace("\n", " ")
            clean_tight = text.replace(" ", "").replace("\n", "")
            
            # =================================================================
            # PASAPORT DOSYASI KONTROLÜ
            # =================================================================
            if is_passport_file:
                # 1. Pasaport No
                if not pass_no:
                    m_pass = re.search(r'\b([UuVv]\d{8})\b', clean_tight)
                    if m_pass:
                        pass_no = m_pass.group(1).upper().replace('V', 'U')

                # 2. TCKN (MRZ veya Etiket Üzerinden)
                if not tckn and nationality == "Turkish":
                    m_tc_mrz = re.search(r'([1-9]\d{10})<', clean_tight)
                    if m_tc_mrz:
                        tckn = m_tc_mrz.group(1)
                    else:
                        m_tc_lbl = re.search(r'(?:K[İI]ML[İI]K|ID)\s*(?:NO)?[\.\s:]*([1-9]\d{10})', text, re.IGNORECASE)
                        if m_tc_lbl:
                            tckn = m_tc_lbl.group(1)
                        else:
                            m_tc_gen = re.search(r'(?<!\d)([1-9]\d{10})(?!\d)', clean_search)
                            if m_tc_gen:
                                tckn = m_tc_gen.group(1)

                # 3. Ad Soyad (MRZ veya Etiket Üzerinden)
                if not best_name:
                    # MRZ: P<TURKAHRIMAN<<AHMET<TAHA
                    for line in text.upper().split('\n'):
                        line_c = line.replace(" ", "")
                        if "<<" in line_c and ("P<TUR" in line_c or "P<" in line_c):
                            parts = line_c.split("<<")
                            if len(parts) >= 2:
                                surname = re.sub(r'^[A-Z]<[A-Z]{3}', '', parts[0]).replace("<", "").strip()
                                name_parts = parts[1].split("<")
                                name = " ".join([n for n in name_parts if n]).strip()
                                if len(surname) > 1 and len(name) > 1:
                                    best_name = f"{name} {surname}"
                                    break
                    if not best_name:
                        m_sur = re.search(r'SURNAME\s*[\n\r:]*\s*([A-ZÇĞİÖŞÜ]+)', text.upper())
                        m_name = re.search(r'NAME\s*[\n\r:]*\s*([A-ZÇĞİÖŞÜ\s]+?)(?=\n|PASAPORT|DATE|DOĞUM|$)', text.upper())
                        if m_sur and m_name:
                            sur_val = m_sur.group(1).strip()
                            name_val = m_name.group(1).strip().replace("\n", " ")
                            if len(sur_val) > 1 and len(name_val) > 1:
                                best_name = f"{name_val} {sur_val}"

            # =================================================================
            # LİMAN CÜZDANI DOSYASI KONTROLÜ
            # =================================================================
            if is_seaman_file:
                # 1. Cüzdan No (S 00500096)
                if not sb:
                    m_sb = re.search(r'([S58$]\s*00\d{5,7})', text, re.IGNORECASE)
                    if not m_sb:
                        m_sb = re.search(r'([S58$]00\d{5,7})', clean_tight, re.IGNORECASE)
                    if m_sb:
                        sb = clean_sb(m_sb.group(1))

                # Pasaportta isim okunamadıysa Cüzdan MRZ'sinden çek
                if not best_name:
                    for line in text.upper().split('\n'):
                        line_c = line.replace(" ", "")
                        if "<<" in line_c and "G<TUR" in line_c:
                            parts = line_c.split("<<")
                            if len(parts) >= 2:
                                clean_p = parts[0].replace("G<TUR", "").split("<")
                                clean_p = [p for p in clean_p if p]
                                if len(clean_p) >= 2:
                                    best_name = f"{' '.join(clean_p[:-1])} {clean_p[-1]}"
                                    break

                # Pasaportta TCKN okunamadıysa Cüzdandan tamamla
                if not tckn and nationality == "Turkish":
                    m_tc_cuz = re.search(r'(?<!\d)([1-9]\d{10})(?!\d)', clean_search)
                    if m_tc_cuz:
                        tckn = m_tc_cuz.group(1)

    return best_name, tckn, sb, pass_no, text_blob

# =============================================================================
# EKRAN ARAYÜZÜ VE FORMLAR
# =============================================================================
uploaded_files = st.file_uploader(
    "Evrakları Yükleyin (PDF, JPG, JPEG, PNG)",
    type=["pdf", "jpg", "jpeg", "png"],
    accept_multiple_files=True
)

if uploaded_files:
    crew_name_ext, tckn_ext, sb_ext, pass_ext, evidence_text = extract_data(uploaded_files)
    
    st.subheader("📋 Çıkarılan / Teyit Edilen Bilgiler")
    st.caption("Veriler belgelerden otomatik çekilmiştir. Gerekirse elle müdahale edebilirsiniz.")
    
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
    
    if nationality == "Turkish":
        if not tckn_val or not sb_val:
            st.error("🚨 SİSTEM KİLİTLİ: Türk personel için TCKN veya Liman Cüzdan Numarası eksik. Lütfen kontrol ediniz.")
            validation_locked = True
        else:
            st.success("✅ Veriler doğrulandı.")
            uab_url = f"https://gemiadami.uab.gov.tr/verification/{tckn_val}/{sb_val}"
            st.markdown(f'<a href="{uab_url}" target="_blank" style="background-color: #0056b3; color: white; padding: 8px 16px; border-radius: 4px; text-decoration: none; font-weight: bold; display: inline-block; margin-bottom: 10px;">UAB Doğrulama Portalı ↗</a>', unsafe_allow_html=True)
    else:
        if not pass_val or not sb_val:
            st.error(f"🚨 SİSTEM KİLİTLİ: {nationality} personel için Pasaport No veya Seaman's Book No eksik.")
            validation_locked = True
        else:
            st.success(f"✅ {nationality} personel evrakları doğrulandı.")

    st.divider()

    # =========================================================================
    # 1. 35-101 EXCEL ŞABLON DOLDURMA
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
        
        for r in range(14, 85):
            c_cell = ws1.cell(r, rank_col)
            if type(c_cell) is not openpyxl.cell.cell.MergedCell:
                c_cell.fill = yellow_fill

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
    # 2. 35-102 WORD ŞABLON DOLDURMA
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
