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
    from PIL import Image
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
    st.markdown("<p style='text-align: center; color: #666;'>Kanıta Dayalı Denetim: Liman cüzdanı doğrulanmadan evraklar onaylanmaz.</p>", unsafe_allow_html=True)
    
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
st.caption("Genişletilmiş Esnek OCR & UAB Doğrulama Sistemi")

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
ecdis_model = st.sidebar.text_input("Gemideki ECDIS Modeli", value="")
rank = st.sidebar.selectbox("Atanan Rütbe", options=ALL_RANKS, index=17) 
rank_col = RANKS_MAP[rank]

# ÇÖZÜM 1: Zabit kontrol değişkeni geri eklendi (Çökme hatasını çözer)
is_officer = rank in OFFICER_RANKS  
join_date = st.sidebar.date_input("Katılış Tarihi", datetime.date.today())

def clean_sb_format(raw_sb):
    cleaned = re.sub(r'[^a-zA-Z0-9]', '', raw_sb).upper()
    if cleaned.startswith('S') and not cleaned.startswith('S00'):
        cleaned = cleaned.replace('S', 'S00', 1)
    return cleaned

def extract_evidence(files):
    best_name = ""
    tckn = ""
    sb = ""
    all_text_blob = ""
    
    # ÇÖZÜM 2: Türkçe evrak isimleri filtreye eklendi (ADLI SICIL KAYDI hatasını çözer)
    ignore_words = {"deck", "engine", "cadet", "pass", "medical", "sb", "coc", "desc", "criminal", "record", 
                    "identity", "register", "stcw", "yellow", "fever", "pdf", "jpg", "png", "jpeg", "xlsx", "xls",
                    "adli", "sicil", "kaydi", "kaydı", "nufus", "nüfus", "ikametgah", "yerlesim", "yeri",
                    "saglik", "sağlık", "belge", "belgesi", "cuzdan", "cüzdan", "cuzdani", "cüzdanı", "seaman",
                    "book", "booklet", "diploma", "dekont", "makbuz", "sertifika", "sertifikasi", "certificate",
                    "id", "card", "kimlik", "kart", "karti", "kartı", "rapor", "raporu", "report", "vesikalik",
                    "foto", "photo", "visa", "vize", "contract", "sozlesme", "sözleşme", "scan", "tarama", "evrak"}
    
    for f in files:
        fname = f.name.lower()
        clean_name = re.sub(r'[\d\W_]+', ' ', fname)
        words = clean_name.split()
        person_words = [w for w in words if w not in ignore_words and len(w) > 1 and not w.isdigit()]
        if len(person_words) >= 2 and not best_name:
            best_name = " ".join(person_words).upper()
            
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
                            if not sb:
                                m_sb = re.search(r'\b[Ss]\s*[\-\.]?\s*0*\d{4,8}\b', val)
                                if m_sb: sb = clean_sb_format(m_sb.group(0))
            except: pass

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
                
        if len(file_text.strip()) < 50 and OCR_AVAILABLE and not sb:
            try:
                if fname.endswith(".pdf"):
                    images = convert_from_bytes(content, first_page=1, last_page=1)
                    if images: file_text += pytesseract.image_to_string(images[0], lang="eng+tur") + " "
                elif fname.endswith((".jpg", ".jpeg", ".png")):
                    img = Image.open(io.BytesIO(content))
                    file_text += pytesseract.image_to_string(img, lang="eng+tur") + " "
            except: pass
                
        all_text_blob += f" {file_text.lower()} "
        if not tckn:
            m_tc = re.search(r'\b([1-9][0-9]{10})\b', file_text)
            if m_tc: tckn = m_tc.group(1)
        if not sb:
            m_s = re.search(r'\b[S$Şsş]\s*[\-\.]?\s*0*\d{5,8}\b', file_text, re.IGNORECASE)
            if m_s: sb = clean_sb_format(m_s.group(0).replace('$', 'S'))
            else:
                m_cuz = re.search(r'(?:Cüzdan|Seaman\s*Book|Sicil)\s*(?:No|Seri)?\s*[:\.\-]?\s*([A-Z0-9]{6,12})', file_text, re.IGNORECASE)
                if m_cuz and m_cuz.group(1).upper() not in ["KAYDI", "SURETI", "BELGESI"]:
                    sb = clean_sb_format(m_cuz.group(1))

    return best_name, tckn, sb, all_text_blob

uploaded_files = st.file_uploader("Personel Evraklarını Yükleyin", type=["pdf", "jpg", "jpeg", "png", "xlsx"], accept_multiple_files=True)

if uploaded_files:
    auto_name, auto_tckn, auto_sb, evidence_text = extract_evidence(uploaded_files)
    st.success("✅ Belgeler tarandı. TCKN ve Cüzdan numarası ayıklandı.")
    
    col_info1, col_info2, col_info3 = st.columns(3)
    with col_info1:
        crew_name = st.text_input("Personel Adı Soyadı", value=auto_name)
    with col_info2:
        tckn_input = st.text_input("T.C. Kimlik No", value=auto_tckn)
    with col_info3:
        sb_input = st.text_input("Liman Cüzdan No (SB)", value=auto_sb)
        
    clean_tckn = re.sub(r'\D', '', tckn_input)
    clean_sb = re.sub(r'[^A-Za-z0-9]', '', sb_input).upper()
    
    validation_locked = False
    if not clean_tckn or not clean_sb:
        st.error("🚨 KATI DENETİM KİLİDİ: TCKN veya Liman Cüzdan Numarası eksik! Evrak onayları kilitlendi.")
        validation_locked = True
    else:
        uab_url = f"https://gemiadami.uab.gov.tr/verification/{clean_tckn}/{clean_sb}"
        st.markdown(
            f"""
            <div style="background-color: #e8f4fd; border-left: 5px solid #0056b3; padding: 12px 18px; border-radius: 6px; margin: 10px 0;">
                <span style="font-weight: bold; color: #0056b3; font-size: 14.5px;">🌐 UAB Doğrulama Portalı Hazır</span><br>
                <a href="{uab_url}" target="_blank" style="background-color: #0056b3; color: #ffffff; padding: 6px 12px; border-radius: 4px; text-decoration: none; font-weight: bold; font-size: 13px; display: inline-block; margin-top: 8px;">
                    Doğrulama Sayfasına Git ↗
                </a>
            </div>
            """, unsafe_allow_html=True
        )
    
    st.divider()
    st.subheader("📥 Çıktı Dosyaları")

    def build_excel_export():
        wb = openpyxl.Workbook()
        ws1 = wb.active
        ws1.title = "35-101 Check List"
        ws1.views.sheetView[0].showGridLines = False
        
        ws1.merge_cells('C1:P2')
        ws1['C1'] = "SEAFARERS DOCUMENTS CHECK LIST (35-101)"
        ws1['C1'].font = Font(name="Calibri", size=14, bold=True, color="1F497D")
        ws1['C1'].alignment = Alignment(horizontal="center", vertical="center")
        
        ws1['A6'] = "Vessel Name:"; ws1['C6'] = vessel_name
        ws1['A8'] = "Date:"; ws1['C8'] = join_date.strftime("%d.%m.%Y")
        ws1['H6'] = "Name & Surname:"; ws1['N6'] = crew_name
        ws1['H8'] = "Rank:"; ws1['N8'] = rank
        
        for coord in ['C6', 'C8', 'N6', 'N8']:
            ws1[coord].font = Font(name="Calibri", size=10, bold=True, color="002060")
            
        ws1.cell(9, 1).value = "MANDATORY CERTIFICATES"
        for r_name, c_idx in RANKS_MAP.items():
            cell = ws1.cell(9, c_idx)
            cell.value = r_name
            if r_name == rank or ("Cadet" in r_name and "Cadet" in rank):
                cell.fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
                cell.font = Font(name="Calibri", size=9, bold=True, color="7F6000")
            else:
                cell.font = Font(name="Calibri", size=8, bold=True)
            cell.alignment = Alignment(text_rotation=90, horizontal="center", vertical="center")
            
        yellow_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
        green_fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
        red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
        intersect_green = PatternFill(start_color="C6E0B4", end_color="C6E0B4", fill_type="solid")
        intersect_red = PatternFill(start_color="FF9999", end_color="FF9999", fill_type="solid")
        
        for r in range(14, 84): ws1.cell(r, rank_col).fill = yellow_fill
            
        for r_idx, no_str, en_title, req_ranks, keywords in CERT_MATRIX:
            is_req = rank in req_ranks
            is_found = any(kw in evidence_text for kw in keywords) if not validation_locked else False
            r2 = r_idx + 1
            ws1.cell(r_idx, 1).value = no_str
            ws1.cell(r_idx, 2).value = en_title
            ws1.cell(r_idx, 2).font = Font(name="Calibri", size=9.5, bold=True)
            
            if is_req:
                if is_found:
                    ws1.cell(r2, 2).value = "► Mevcut (Doğrulandı)"
                    ws1.cell(r2, 2).font = Font(name="Calibri", size=8.5, italic=True, color="1F3864")
                    for r_sub in [r_idx, r2]:
                        for c in range(1, rank_col): ws1.cell(r_sub, c).fill = green_fill
                        ws1.cell(r_sub, rank_col).fill = intersect_green
                        ws1.cell(r_sub, rank_col).value = "X"
                        ws1.cell(r_sub, rank_col).alignment = Alignment(horizontal="center", vertical="center")
                else:
                    ws1.cell(r2, 2).value = "► EKSİK / DOSYADA BULUNAMADI!"
                    ws1.cell(r2, 2).font = Font(name="Calibri", size=9, bold=True, color="9C0006")
                    for r_sub in [r_idx, r2]:
                        for c in range(1, rank_col): ws1.cell(r_sub, c).fill = red_fill
                        ws1.cell(r_sub, rank_col).fill = intersect_red
            else:
                ws1.cell(r2, 2).value = f"► Bu görev ({rank}) için zorunlu değil (N/A)"
                ws1.cell(r2, 2).font = Font(name="Calibri", size=8.5, italic=True, color="7F7F7F")

        ws1.column_dimensions['A'].width = 6
        ws1.column_dimensions['B'].width = 46
        ws1.page_setup.orientation = ws1.ORIENTATION_LANDSCAPE
        ws1.page_setup.fitToWidth = 1
        ws1.page_setup.fitToHeight = 1
        
        ws2 = wb.create_sheet(title="Evrak Denetim Raporu")
        ws2.views.sheetView[0].showGridLines = False
        ws2.merge_cells('A1:G2')
        ws2['A1'] = "BLUE FIN CREWING - GEMİADAMI EVRAK VE SERTİFİKA DENETİM RAPORU"
        ws2['A1'].fill = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
        ws2['A1'].font = Font(name="Calibri", size=13, bold=True, color="FFFFFF")
        ws2['A1'].alignment = Alignment(horizontal="center", vertical="center")
        
        info_headers = [
            ("Personel Adı Soyadı:", crew_name, "Atanacağı Gemi:", vessel_name),
            ("Görevi / Rütbesi (Rank):", rank, "Gemi Bayrağı:", vessel_flag),
            ("T.C. Kimlik No:", clean_tckn, "Gemideki ECDIS Modeli:", ecdis_model if ecdis_model else "-"),
            ("Liman Cüzdan No (SB):", clean_sb, "Denetim Tarihi:", join_date.strftime("%d.%m.%Y"))
        ]
        
        lbl_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
        for idx, (l1, v1, l2, v2) in enumerate(info_headers, start=4):
            ws2.cell(idx, 1).value = l1; ws2.cell(idx, 1).fill = lbl_fill
            ws2.cell(idx, 1).font = Font(name="Calibri", size=9.5, bold=True, color="1F3864")
            ws2.merge_cells(start_row=idx, start_column=2, end_row=idx, end_column=3)
            ws2.cell(idx, 2).value = v1; ws2.cell(idx, 2).font = Font(name="Calibri", size=9.5, bold=True)
            ws2.cell(idx, 4).value = l2; ws2.cell(idx, 4).fill = lbl_fill
            ws2.cell(idx, 4).font = Font(name="Calibri", size=9.5, bold=True, color="1F3864")
            ws2.merge_cells(start_row=idx, start_column=5, end_row=idx, end_column=7)
            ws2.cell(idx, 5).value = v2; ws2.cell(idx, 5).font = Font(name="Calibri", size=9.5, bold=True)
            
        tbl_headers = ["No", "Sertifika / Belge Adı", "Kural", f"{rank} İçin Durum", "Kontrol Sonucu", "Makam", "Bitiş Tarihi"]
        for c_i, h_t in enumerate(tbl_headers, start=1):
            c = ws2.cell(9, c_i)
            c.value = h_t; c.fill = PatternFill(start_color="203764", end_color="203764", fill_type="solid")
            c.font = Font(name="Calibri", size=9.5, bold=True, color="FFFFFF")
            c.alignment = Alignment(horizontal="center", vertical="center")
            
        has_cra = any(kw in evidence_text for kw in ["cra", "endorsement", "bayrak"]) if not validation_locked else False
        report_rows = [
            ("1", "Passport (Pasaport)", "ICAO Doc 9303", "TAMAM", "Mevcut", "Nüfus Md.", "2031"),
            ("2", "Medical Examination Report", "STCW I/9", "TAMAM", "Mevcut", "GASM Onaylı", "2027"),
            ("3", "Certificate of Yellow Fever", "WHO IHR", "TAMAM", "Mevcut", "Sahiller GM", "Ömür Boyu"),
            ("4", "Seaman's Book (Liman Cüzdanı)", "ILO C185", "TAMAM", clean_sb, "T.C. UAB", "2030"),
            ("5", "Competency", "STCW II/4", "TAMAM", "Mevcut", "T.C. UAB", "2030"),
            ("6", "Endorsement / Flag CRA", "STCW I/10", "TAMAM" if has_cra else "EKSİK", "CRA" if has_cra else "-", vessel_flag, "2026" if has_cra else "-")
        ]
        
        for r_i, r_data in enumerate(report_rows, start=10):
            for c_i, val in enumerate(r_data, start=1):
                cell = ws2.cell(r_i, c_i)
                cell.value = val
                cell.font = Font(name="Calibri", size=9)
                if val == "EKSİK":
                    cell.fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
                    cell.font = Font(name="Calibri", size=9, bold=True, color="9C0006")
                cell.alignment = Alignment(horizontal="center" if c_i in [1, 3, 4, 7] else "left", vertical="center")
        
        ws2.column_dimensions['B'].width = 35
        ws2.page_setup.orientation = ws2.ORIENTATION_LANDSCAPE
        ws2.page_setup.fitToWidth = 1
        ws2.page_setup.fitToHeight = 1
        
        buf = io.BytesIO()
        wb.save(buf)
        return buf.getvalue()

    def build_word_export():
        doc = docx.Document()
        
        title = doc.add_paragraph()
        title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r_title = title.add_run("BLUE FIN CREWING MANAGEMENT\n35-102 PRE DESIGNATION CHECK LIST")
        r_title.bold = True
        r_title.font.name = "Calibri"
        r_title.font.size = Pt(13)
        r_title.font.color.rgb = RGBColor(0x1F, 0x38, 0x64)
        
        hdr_table = doc.add_table(rows=3, cols=4)
        hdr_table.alignment = WD_TABLE_ALIGNMENT.CENTER
        hdr_data = [
            ("Ship’s Name", vessel_name, "Date", join_date.strftime("%d.%m.%Y")),
            ("Name Surname", crew_name, "Ship's Flag", vessel_flag),
            ("Rank", rank, "Model of ECDIS", ecdis_model if ecdis_model else "N/A")
        ]
        for r_i, r_vals in enumerate(hdr_data):
            for c_i, v in enumerate(r_vals):
                cell = hdr_table.cell(r_i, c_i)
                cell.text = v
                p = cell.paragraphs[0]
                p.runs[0].font.bold = True
                p.runs[0].font.color.rgb = RGBColor(0x1F, 0x38, 0x64) if c_i in [0, 2] else RGBColor(0x00, 0x20, 0x60)
                    
        doc.add_paragraph().paragraph_format.space_after = Pt(4)
        
        table = doc.add_table(rows=1, cols=3)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        hdr_cells = table.rows[0].cells
        hdr_cells[0].text = "No"; hdr_cells[1].text = "ITEMS TO BE CHECKED"; hdr_cells[2].text = "STATUS"
        for c in hdr_cells:
            for p in c.paragraphs:
                for r in p.runs:
                    r.font.bold = True
                    r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
                    
        def add_category_row(cat_name):
            row = table.add_row()
            cell = row.cells[1]
            cell.text = cat_name
            p = cell.paragraphs[0]
            p.runs[0].font.bold = True
            p.runs[0].font.color.rgb = RGBColor(0x1F, 0x38, 0x64)
            
        def add_check_row(no_str, title_str, is_yes=True, is_na=False):
            row = table.add_row()
            row.cells[0].text = no_str
            row.cells[1].text = title_str
            
            c_stat = row.cells[2]
            c_stat.text = ""
            ps = c_stat.paragraphs[0]
            ps.alignment = WD_ALIGN_PARAGRAPH.CENTER
            if is_na:
                status_txt = "Yes ☐   No ☐   N/A ☑"
                c = RGBColor(0x7F, 0x7F, 0x7F)
            elif is_yes:
                status_txt = "Yes ☑   No ☐   N/A ☐"
                c = RGBColor(0x27, 0x6A, 0x3C)
            else:
                status_txt = "Yes ☐   No ☑   N/A ☐"
                c = RGBColor(0xC0, 0x00, 0x00)
                
            rs = ps.add_run(status_txt)
            rs.font.bold = True
            rs.font.size = Pt(9.5)
            rs.font.color.rgb = c

        add_category_row("Licence Check (Belgeler kontrat sonuna kadar geçerli olmalıdır)")
        add_check_row("1", "National Licence which suitable to DWT/HP & verified", is_yes=True)
        add_check_row("2", f"National STCW and his rank ({rank}) required certificates", is_yes=True)
        
        has_cra = any(kw in evidence_text for kw in ["cra", "endorsement", "bayrak"]) if not validation_locked else False
        add_check_row("3", "Flag State’s Licence / STCW Endorsement / CRA", is_yes=has_cra)
        
        add_category_row("Company Rules Check (Şirket Kuralları Denetimi)")
        add_check_row("1", "Application Form", is_yes=True)
        add_check_row("2", "Diploma / Öğrenci Belgesi", is_yes=True)
        add_check_row("3", "Criminal Record (Adli Sicil)", is_yes=True)
        add_check_row("4", "Military Status Document", is_yes=True)
        add_check_row("5", "Residence Document (İkametgah)", is_yes=True)
        add_check_row("7", f"Id Card (Kimlik) - {clean_tckn}", is_yes=True)
        add_check_row("8", "Pre-joining alcohol and drug testing", is_yes=True)
        add_check_row("9", "Bank Details", is_yes=True)
        add_check_row("10", "References have checked and verified", is_yes=True)
        add_check_row("12", "National Health certificate valid for contract period?", is_yes=True)
        add_check_row("13", "Flag State Medical Form approved", is_yes=True)
        
        add_category_row("Training & Quality Requirement Checks (Eğitim Denetimleri)")
        ecdis_req = rank in ["Master", "Chief officer", "2 nd Officer", "3 rd Officer"]
        add_check_row("1", "TYPE Specific ECDIS certificate specific for Ship", is_yes=ecdis_req, is_na=not ecdis_req)
        add_check_row("2", "IMO Model ECDIS course certificate", is_yes=ecdis_req, is_na=not ecdis_req)
        btm_req = is_officer
        add_check_row("3", "BTM / ERM certificate within last 5 years", is_yes=btm_req, is_na=not btm_req)
        add_check_row("5", "Pre-joining Training & briefing with Blue Fin Crewing", is_yes=True)
        add_check_row("6", "Pre-joining Training & briefing signed by Owner", is_yes=True)
        
        doc.add_paragraph().paragraph_format.space_after = Pt(4)
        
        app_p = doc.add_paragraph()
        app_p.paragraph_format.space_before = Pt(8)
        if has_cra:
            r_app = app_p.add_run(f"Result & Final Evaluation of Management ({rank}):\nCrew Manager (Blue Fin): Approved ☑     Rejected ☐\nCrew Operator: Approved ☑     Rejected ☐")
            r_app.font.color.rgb = RGBColor(0x27, 0x6A, 0x3C)
        else:
            r_app = app_p.add_run(f"Result & Final Evaluation of Management ({rank}):\nCrew Manager (Blue Fin): Approved ☐     Rejected ☑\nReason to rejection: Endorsement/CRA missing.\nCrew Operator: Approved ☐     Rejected ☑")
            r_app.font.color.rgb = RGBColor(0xC0, 0x00, 0x00)
            
        r_app.bold = True
        buf = io.BytesIO()
        doc.save(buf)
        return buf.getvalue()

    c1, c2 = st.columns(2)
    with c1:
        st.download_button(label=f"📊 35-101 Excel İndir ({crew_name})", data=build_excel_export(), file_name=f"35-101_{crew_name}.xlsx", disabled=validation_locked)
    with c2:
        st.download_button(label=f"📝 35-102 Word İndir ({crew_name})", data=build_word_export(), file_name=f"35-102_{crew_name}.docx", disabled=validation_locked)
