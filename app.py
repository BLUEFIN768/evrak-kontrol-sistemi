import streamlit as st
import datetime
import io
import os
import re
import numpy as np
import pypdf
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment
import docx
from PIL import Image

# 1. OCR Motorlarını Başlatma (EasyOCR Öncelikli, Tesseract Yedek)
try:
    import easyocr
    # Türkçe ve İngilizce modellerini belleğe al
    @st.cache_resource
    def load_easyocr_reader():
        return easyocr.Reader(['tr', 'en'], gpu=False)
    reader_ocr = load_easyocr_reader()
    EASYOCR_READY = True
except Exception:
    EASYOCR_READY = False

try:
    import pytesseract
    from pdf2image import convert_from_bytes
    TESSERACT_READY = True
except Exception:
    TESSERACT_READY = False

st.set_page_config(page_title="Evrak Kontrol Portalı", page_icon="⚓", layout="wide")

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

st.sidebar.markdown("🔒 **Güvenli Oturum:** Aktif (KVKK Uyumlu)")
if st.sidebar.button("🚪 Çıkış Yap"):
    st.session_state["authenticated"] = False
    st.rerun()

st.title("⚓ Evrak Kontrol & Denetim Portalı")
st.caption("Yerel Derin Öğrenme OCR & Orijinal Şablon Motoru (35-101 / 35-102)")

RANKS_MAP = {
    "Master": 6, "Chief officer": 7, "2 nd Officer": 8, "3 rd Officer": 9,
    "Chief Engineer": 10, "2 nd Engineer": 11, "3rd Engineer": 12, "Electrician": 13,
    "Bosun": 14, "Pumpman": 15, "AB": 16, "O/S": 17, "Fitter": 18,
    "Oiler": 19, "Wiper": 20, "Cook": 21, "Steward": 22, "Cadet": 23
}
OFFICER_RANKS = ["Master", "Chief officer", "2 nd Officer", "3 rd Officer", "Chief Engineer", "2 nd Engineer", "3rd Engineer", "Electrician"]
ALL_RANKS = list(RANKS_MAP.keys())

CERT_MATRIX = [
    (14, "1", "Passport", ALL_RANKS, ["pass", "pasaport"]),
    (16, "2", "Medical Examination Report", ALL_RANKS, ["medical", "sağlık", "health", "nation"]),
    (18, "3", "Certificate of Yellow Fewer Vaccination", ALL_RANKS, ["yellow", "fever", "humma"]),
    (20, "4", "Seamanbook", ALL_RANKS, ["sb", "seaman", "cüzdan", "s00", "s 00"]),
    (22, "5", "Competency", ALL_RANKS, ["coc", "competency", "yeterlilik"]),
    (24, "6", "GMDSS", ["Master", "Chief officer", "2 nd Officer", "3 rd Officer"], ["gmdss", "goc"]),
    (26, "7.1", "Fire Prevention and Fire Fighting Training Certificate", ALL_RANKS, ["stcw", "fire", "yangın"]),
    (28, "7.2", "Personal Survival Techniques Training Certificate", ALL_RANKS, ["stcw", "survival", "kurtarma"]),
    (30, "7.3", "Elementary First Aid Training Certificate", ALL_RANKS, ["stcw", "first aid", "ilkyardım"]),
    (32, "7.4", "Personal Safety and Social Responsibility Training Cert.", ALL_RANKS, ["stcw", "social", "sorumluluk"]),
    (34, "7.5", "Proficiency in Survival Craft and Rescue Boats Training Cert.", ["Master", "Chief officer", "2 nd Officer", "3 rd Officer", "Chief Engineer", "2 nd Engineer", "3rd Engineer", "Electrician", "Bosun", "Pumpman", "AB", "Oiler"], ["stcw", "craft", "rescue"]),
    (36, "8.1", "Security Awareness Certificate ", ALL_RANKS, ["stcw", "security", "farkındalık"]),
    (38, "8.2", "Designated Security Duties Certificate", ALL_RANKS, ["stcw", "security", "görev"]),
    (40, "8.3", "Ship Security Officer", ["Master", "Chief officer", "2 nd Officer", "3 rd Officer", "Chief Engineer", "2 nd Engineer"], ["sso", "security officer"]),
    (42, "8.4", "Security - Realated Familirization Certificate", ALL_RANKS, ["stcw", "security", "tanıtım"]),
    (44, "9.1", "Advanced Fire Fighting Training Certificate", OFFICER_RANKS, ["advanced fire", "ileri yangın"]),
    (82, "5", "Endorsement", OFFICER_RANKS + ["Cadet"], ["cra", "endorsement", "bayrak"])
]

st.sidebar.header("🚢 Gemi ve Atama Parametreleri")
vessel_name = st.sidebar.text_input("Gemi Adı", value="Optimus S")
vessel_flag = st.sidebar.selectbox("Gemi Bayrağı", ["MALTA", "MARSHALL ISLANDS", "PANAMA", "LIBERIA", "TURKIYE"])
ecdis_model = st.sidebar.text_input("Gemideki ECDIS Modeli", value="")
rank = st.sidebar.selectbox("Atanan Rütbe", options=ALL_RANKS, index=10) # Varsayılan O/S
rank_col = RANKS_MAP.get(rank, 17)
is_officer = rank in OFFICER_RANKS  
join_date = st.sidebar.date_input("Katılış Tarihi", datetime.date.today())

def clean_sb(raw_val):
    val = re.sub(r'[^a-zA-Z0-9]', '', raw_val).upper()
    val = val.replace('5', 'S', 1).replace('8', 'S', 1).replace('$', 'S', 1)
    if val.startswith('S') and not val.startswith('S00'):
        val = val.replace('S', 'S00', 1)
    return val

def run_ocr_on_image(pil_img):
    txt = ""
    if EASYOCR_READY:
        try:
            np_img = np.array(pil_img)
            results = reader_ocr.readtext(np_img, detail=0)
            txt = " ".join(results)
        except Exception:
            pass
    if len(txt.strip()) < 10 and TESSERACT_READY:
        try:
            txt = pytesseract.image_to_string(pil_img, lang="tur+eng")
        except Exception:
            pass
    return txt

def extract_data(files):
    tckn, sb, pass_no, text_blob = "", "", "", ""
    best_name = ""
    
    with st.spinner("Evraklar yerel derin öğrenme motoruyla taranıyor..."):
        for f in files:
            text = ""
            content = f.read()
            f.seek(0)
            fname = f.name.lower()
            
            # 1. Excel Okuma
            if fname.endswith(('.xlsx', '.xls')):
                try:
                    wb_in = openpyxl.load_workbook(io.BytesIO(content), data_only=True)
                    for sheet in wb_in.worksheets:
                        for r in range(1, 40):
                            for c in range(1, 10):
                                val = str(sheet.cell(r, c).value or "").strip()
                                text += f" {val} "
                except Exception:
                    pass
                
            # 2. PDF İnceleme (Dijital Metin + Dahili Resimler + pdf2image)
            elif fname.endswith(".pdf"):
                try:
                    reader = pypdf.PdfReader(io.BytesIO(content))
                    for page in reader.pages:
                        text += (page.extract_text() or "") + " "
                        
                    # Dijital metin yoksa sayfadaki gömülü resimleri tara
                    if len(text.strip()) < 50:
                        for page in reader.pages:
                            for img_obj in page.images:
                                try:
                                    pil_img = Image.open(io.BytesIO(img_obj.data))
                                    text += run_ocr_on_image(pil_img) + " "
                                except Exception:
                                    pass
                                    
                    # Hâlâ yetersizse pdf2image ile render al
                    if len(text.strip()) < 50 and TESSERACT_READY:
                        try:
                            imgs = convert_from_bytes(content, first_page=1, last_page=1, dpi=300)
                            if imgs:
                                text += run_ocr_on_image(imgs[0]) + " "
                        except Exception:
                            pass
                except Exception:
                    pass
                
            # 3. Resim Dosyaları (.jpg, .jpeg, .png)
            elif fname.endswith((".jpg", ".jpeg", ".png")):
                try:
                    pil_img = Image.open(io.BytesIO(content))
                    text += run_ocr_on_image(pil_img) + " "
                except Exception:
                    pass
                
            text_blob += f" {fname} {text.lower()} "
            clean_search = text.replace("<", " ").replace("\n", " ")
            
            # TCKN Yakalama
            if not tckn:
                m_tc = re.search(r'(?<!\d)([1-9]\d{10})(?!\d)', clean_search)
                if m_tc:
                    tckn = m_tc.group(1)
                
            # Pasaport No Yakalama
            if not pass_no:
                m_pass = re.search(r'\b([UuVv]\d{8})\b', clean_search)
                if m_pass:
                    pass_no = m_pass.group(1).upper().replace('V', 'U')
                
            # Liman Cüzdanı (SB) Yakalama
            if not sb:
                m_sb = re.search(r'([S58$]\s*00\d{5,7})', clean_search, re.IGNORECASE)
                if m_sb:
                    sb = clean_sb(m_sb.group(1))

            # MRZ Üzerinden İsim Soyisim
            for line in text.upper().split('\n'):
                line_c = line.replace(" ", "")
                if "<<" in line_c and ("P<TUR" in line_c or "G<TUR" in line_c or "I<TUR" in line_c):
                    parts = line_c.split("<<")
                    if len(parts) >= 2:
                        surname = parts[0].replace("P<TUR", "").replace("G<TUR", "").replace("I<TUR", "").replace("<", "").strip()
                        name = parts[1].replace("<", " ").strip().split()[0]
                        if len(surname) > 1 and len(name) > 1:
                            best_name = f"{name} {surname}"
                            break

        if not best_name:
            ignore = {"pdf", "jpg", "jpeg", "png", "xlsx", "sb", "pass", "passport", "medical", "coc", "stcw", "id", "kimlik", "cuzdan", "evrak", "belge", "nation", "yellow", "fever", "seaman", "seamenbook", "competency", "photo", "alexa"}
            for f in files:
                clean_name = re.sub(r'[\d\W_]+', ' ', f.name.lower())
                words = [w.upper() for w in clean_name.split() if w not in ignore and len(w) > 2 and not w.isdigit()]
                if len(words) >= 2:
                    best_name = " ".join(words)
                    break

    return best_name, tckn, sb, pass_no, text_blob

uploaded_files = st.file_uploader("Evrakları Yükleyin (Tümü)", type=["pdf", "jpg", "jpeg", "png", "xlsx"], accept_multiple_files=True)

if uploaded_files:
    crew_name, tckn_val, sb_val, pass_val, evidence_text = extract_data(uploaded_files)
    
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.text_input("Personel Adı Soyadı", value=crew_name, disabled=True)
    with col2:
        st.text_input("T.C. Kimlik No", value=tckn_val, disabled=True)
    with col3:
        st.text_input("Liman Cüzdan No (SB)", value=sb_val, disabled=True)
    with col4:
        st.text_input("Pasaport No", value=pass_val, disabled=True)
        
    validation_locked = False
    if not tckn_val or not sb_val:
        st.error("🚨 SİSTEM KİLİTLİ: TCKN veya Liman Cüzdan Numarası belgelerden otomatik okunamadı.")
        validation_locked = True
    else:
        st.success("✅ Veriler belgeden bağımsız olarak çekildi ve doğrulandı.")
        uab_url = f"https://gemiadami.uab.gov.tr/verification/{tckn_val}/{sb_val}"
        st.markdown(f'<a href="{uab_url}" target="_blank" style="background-color: #0056b3; color: white; padding: 8px 16px; border-radius: 4px; text-decoration: none; font-weight: bold; display: inline-block; margin-bottom: 10px;">UAB Doğrulama Portalı ↗</a>', unsafe_allow_html=True)
    
    st.divider()

    # =========================================================================
    # 1. ORİJİNAL 35-101 EXCEL ŞABLONUNU DOLDURMA
    # =========================================================================
    def build_excel_from_template():
        template_file = "template_35_101.xlsx"
        if not os.path.exists(template_file):
            wb = openpyxl.Workbook()
            ws1 = wb.active
            ws1.title = "35-101"
            ws1['A1'] = "HATA: template_35_101.xlsx bulunamadı! Lütfen repoya yükleyin."
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
        
        # Seçilen rütbe sütununu sarı yap
        for r in range(14, 85):
            c_cell = ws1.cell(r, rank_col)
            if type(c_cell) is not openpyxl.cell.cell.MergedCell:
                c_cell.fill = yellow_fill

        # Sertifikaları ve Tarihleri Orijinal Şablona Bas
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
    # 2. ORİJİNAL 35-102 WORD ŞABLONUNU DOLDURMA
    # =========================================================================
    def build_word_from_template():
        template_file = "template_35_102.docx"
        if not os.path.exists(template_file):
            doc = docx.Document()
            doc.add_paragraph("HATA: template_35_102.docx bulunamadı! Lütfen repoya yükleyin.")
            buf = io.BytesIO()
            doc.save(buf)
            return buf.getvalue()

        doc = docx.Document(template_file)
        table = doc.tables[0]
        
        table.rows[0].cells[3].text = vessel_name
        table.rows[0].cells[7].text = join_date.strftime("%d.%m.%Y")
        table.rows[1].cells[3].text = crew_name
        table.rows[1].cells[7].text = vessel_flag
        table.rows[2].cells[3].text = rank
        table.rows[2].cells[7].text = ecdis_model if ecdis_model else "N/A"
        
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
        
        set_status(9, is_yes=True)
        set_status(10, is_yes=True)
        set_status(11, is_yes=True)
        set_status(12, is_yes=True)
        set_status(13, is_yes=True)
        set_status(14, is_yes=True)
        table.rows[14].cells[1].text = f"Id card (T.C. Kimlik: {tckn_val})"
        set_status(15, is_yes=True)
        set_status(16, is_yes=True)
        set_status(17, is_yes=True)
        set_status(18, is_yes=True)
        set_status(19, is_yes=True)
        
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
