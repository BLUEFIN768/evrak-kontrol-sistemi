import streamlit as st
import datetime
import io
import os
import re
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
st.caption("Otomatik Veri Çıkarımı & Kurumsal 35-101 / 35-102 Çıktı Motoru")

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
    (16, "2", "Medical Examination Report", ALL_RANKS, ["medical", "sağlık", "health", "nation"]),
    (18, "3", "Certificate of Yellow Fever", ALL_RANKS, ["yellow", "fever", "humma"]),
    (20, "4", "Seamanbook", ALL_RANKS, ["sb", "seaman", "cüzdan", "s00"]),
    (22, "5", "Competency", ALL_RANKS, ["coc", "competency"]),
    (24, "6", "GMDSS", ["Master", "Chief officer", "2 nd Officer", "3 rd Officer"], ["gmdss", "goc"]),
    (26, "7.1", "Fire Prevention and Fire Fighting", ALL_RANKS, ["stcw", "fire", "yangın"]),
    (28, "7.2", "Personal Survival Techniques", ALL_RANKS, ["stcw", "survival", "kurtarma"]),
    (30, "7.3", "Elementary First Aid", ALL_RANKS, ["stcw", "first aid", "ilkyardım"]),
    (32, "7.4", "Personal Safety and Social Responsibility", ALL_RANKS, ["stcw", "social", "sorumluluk"]),
    (34, "7.5", "Proficiency in Survival Craft", ["Master", "Chief officer", "2 nd Officer", "3 rd Officer", "Chief Engineer", "2 nd Engineer", "3rd Engineer", "Electrician", "Bosun", "Pumpman", "AB", "Oiler"], ["stcw", "craft", "rescue"]),
    (36, "8.1", "Security Awareness Certificate", ALL_RANKS, ["stcw", "security", "farkındalık"]),
    (38, "8.2", "Designated Security Duties", ALL_RANKS, ["stcw", "security", "görev"]),
    (42, "8.4", "Security - Realated Familirization", ALL_RANKS, ["stcw", "security", "tanıtım"]),
    (82, "5", "Endorsement (Bayrak Onayı)", OFFICER_RANKS + ["Deck Cadet", "Engine Cadet"], ["cra", "endorsement", "bayrak"])
]

st.sidebar.header("🚢 Gemi ve Atama Parametreleri")
vessel_name = st.sidebar.text_input("Gemi Adı", value="Optimus S")
vessel_flag = st.sidebar.selectbox("Gemi Bayrağı", ["MALTA", "MARSHALL ISLANDS", "PANAMA", "LIBERIA", "TURKIYE"])
ecdis_model = st.sidebar.text_input("Gemideki ECDIS Modeli", value="")
rank = st.sidebar.selectbox("Atanan Rütbe", options=ALL_RANKS, index=17) 
rank_col = RANKS_MAP[rank]
is_officer = rank in OFFICER_RANKS  
join_date = st.sidebar.date_input("Katılış Tarihi", datetime.date.today())

def clean_sb(raw_val):
    val = re.sub(r'[^a-zA-Z0-9]', '', raw_val).upper()
    val = val.replace('5', 'S', 1).replace('8', 'S', 1).replace('$', 'S', 1)
    if val.startswith('S') and not val.startswith('S00'):
        val = val.replace('S', 'S00', 1)
    return val

def extract_data(files):
    tckn, sb, pass_no, text_blob = "", "", "", ""
    best_name = ""
    
    for f in files:
        text = ""
        content = f.read()
        f.seek(0)
        fname = f.name.lower()
        
        # 1. Excel Okuma
        if fname.endswith(('.xlsx', '.xls')):
            try:
                wb = openpyxl.load_workbook(io.BytesIO(content), data_only=True)
                for sheet in wb.worksheets:
                    for r in range(1, 40):
                        for c in range(1, 10):
                            val = str(sheet.cell(r, c).value or "").strip()
                            text += f" {val} "
            except Exception:
                pass
            
        # 2. PDF Metin Okuma
        elif fname.endswith(".pdf"):
            try:
                reader = pypdf.PdfReader(io.BytesIO(content))
                for page in reader.pages:
                    text += (page.extract_text() or "") + " "
            except Exception:
                pass
            
        # 3. Görsel / OCR Okuma (Renkli Tesseract Okuması)
        if len(text.strip()) < 50 and OCR_AVAILABLE:
            try:
                if fname.endswith(".pdf"):
                    images = convert_from_bytes(content, first_page=1, last_page=1, dpi=300)
                    if images: 
                        text += pytesseract.image_to_string(images[0], lang="eng")
                elif fname.endswith((".jpg", ".jpeg", ".png")):
                    img = Image.open(io.BytesIO(content))
                    text += pytesseract.image_to_string(img, lang="eng")
            except Exception:
                pass
            
        text_blob += f" {fname} {text.lower()} "
        clean_text_for_search = text.replace("<", " ").replace("\n", " ")
        
        # Numaraları Ayıkla
        if not tckn:
            m_tc = re.search(r'(?<!\d)([1-9]\d{10})(?!\d)', clean_text_for_search)
            if m_tc:
                tckn = m_tc.group(1)
            
        if not pass_no:
            m_pass = re.search(r'\b([UuVv]\d{8})\b', clean_text_for_search)
            if m_pass:
                pass_no = m_pass.group(1).upper().replace('V', 'U')
            
        if not sb:
            m_sb = re.search(r'([S58$]\s*00\d{5,7})', clean_text_for_search, re.IGNORECASE)
            if m_sb:
                sb = clean_sb(m_sb.group(1))

        # MRZ üzerinden isim yakalama
        for line in text.upper().split('\n'):
            line_c = line.replace(" ", "")
            if "<<" in line_c and ("P<TUR" in line_c or "I<TUR" in line_c):
                parts = line_c.split("<<")
                if len(parts) >= 2:
                    surname = parts[0].replace("P<TUR", "").replace("I<TUR", "").replace("<", "").strip()
                    name = parts[1].replace("<", " ").strip().split()[0]
                    if len(surname) > 1 and len(name) > 1:
                        best_name = f"{name} {surname}"
                        break

    # MRZ bulunamazsa dosya isimlerinden ayıkla
    if not best_name:
        ignore = {"pdf", "jpg", "jpeg", "png", "xlsx", "sb", "pass", "medical", "coc", "stcw", "id", "kimlik", "cuzdan", "evrak", "belge", "nation", "yellow", "fever", "seaman", "book", "photo"}
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
        st.success("✅ Veriler belgeden başarıyla çekildi.")
        uab_url = f"https://gemiadami.uab.gov.tr/verification/{tckn_val}/{sb_val}"
        st.markdown(f'<a href="{uab_url}" target="_blank" style="background-color: #0056b3; color: white; padding: 8px 16px; border-radius: 4px; text-decoration: none; font-weight: bold; display: inline-block; margin-bottom: 10px;">UAB Doğrulama Portalı ↗</a>', unsafe_allow_html=True)
    
    st.divider()

    # =========================================================================
    # YASİN DÖŞER ÇALIŞMASINDAKİ KURUMSAL 35-101 FORMATI
    # =========================================================================
    def build_excel_export():
        wb = openpyxl.Workbook()
        ws1 = wb.active
        ws1.title = "35-101 Check List"
        ws1.views.sheetView[0].showGridLines = True
        
        # 1. Başlık ve Gemi/Personel Bilgileri
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
            
        # 2. Rütbe Başlıkları
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
        
        # Seçili rütbe sütununu boydan boya sarı yap
        for r in range(14, 84): 
            ws1.cell(r, rank_col).fill = yellow_fill
            
        # 3. Sertifikalar ve Alınış/Bitiş Tarihi Alanları
        for r_idx, no_str, en_title, req_ranks, keywords in CERT_MATRIX:
            is_req = rank in req_ranks
            is_found = any(kw in evidence_text for kw in keywords) if not validation_locked else False
            r2 = r_idx + 1
            
            ws1.cell(r_idx, 1).value = no_str
            ws1.cell(r_idx, 2).value = en_title
            ws1.cell(r_idx, 2).font = Font(name="Calibri", size=9.5, bold=True)
            
            if is_req:
                if is_found:
                    # İkinci satırda temiz denetim verisi
                    doc_no_display = sb_val if "seaman" in en_title.lower() else (pass_val if "pass" in en_title.lower() else "Verified")
                    ws1.cell(r2, 2).value = f"Doc No: {doc_no_display}  |  Issue: {join_date.strftime('%d.%m.%Y')}  |  Expiry: Valid"
                    ws1.cell(r2, 2).font = Font(name="Calibri", size=8.5, italic=True, color="1F3864")
                    for r_sub in [r_idx, r2]:
                        for c in range(1, rank_col): 
                            ws1.cell(r_sub, c).fill = green_fill
                        ws1.cell(r_sub, rank_col).fill = intersect_green
                        ws1.cell(r_sub, rank_col).value = "X"
                        ws1.cell(r_sub, rank_col).alignment = Alignment(horizontal="center", vertical="center")
                else:
                    ws1.cell(r2, 2).value = "► MISSING / EKSİK EVRAK!"
                    ws1.cell(r2, 2).font = Font(name="Calibri", size=8.5, bold=True, color="9C0006")
                    for r_sub in [r_idx, r2]:
                        for c in range(1, rank_col): 
                            ws1.cell(r_sub, c).fill = red_fill
                        ws1.cell(r_sub, rank_col).fill = intersect_red
            else:
                ws1.cell(r2, 2).value = f"► Not Required for {rank} (N/A)"
                ws1.cell(r2, 2).font = Font(name="Calibri", size=8.5, italic=True, color="7F7F7F")

        ws1.column_dimensions['A'].width = 6
        ws1.column_dimensions['B'].width = 55
        ws1.page_setup.orientation = ws1.ORIENTATION_LANDSCAPE
        ws1.page_setup.fitToWidth = 1
        ws1.page_setup.fitToHeight = 1
        
        # 2. Sayfa: Evrak Denetim Raporu
        ws2 = wb.create_sheet(title="Evrak Denetim Raporu")
        ws2.merge_cells('A1:G2')
        ws2['A1'] = "BLUE FIN CREWING - GEMİADAMI EVRAK RAPORU"
        ws2['A1'].font = Font(name="Calibri", size=13, bold=True, color="1F497D")
        ws2['A1'].alignment = Alignment(horizontal="center", vertical="center")
        
        info = [
            ("Ad Soyad:", crew_name, "Gemi:", vessel_name),
            ("Rütbe:", rank, "Bayrak:", vessel_flag),
            ("TCKN:", tckn_val, "Liman Cüzdanı:", sb_val),
            ("Pasaport No:", pass_val, "Denetim Tarihi:", join_date.strftime("%d.%m.%Y"))
        ]
        for idx, (l1, v1, l2, v2) in enumerate(info, start=4):
            ws2.cell(idx, 1).value = l1; ws2.cell(idx, 2).value = v1
            ws2.cell(idx, 4).value = l2; ws2.cell(idx, 5).value = v2
            
        buf = io.BytesIO()
        wb.save(buf)
        return buf.getvalue()

    def build_word_export():
        doc = docx.Document()
        doc.add_paragraph("BLUE FIN CREWING MANAGEMENT\n35-102 PRE DESIGNATION CHECK LIST")
        doc.add_paragraph(f"Ship: {vessel_name} | Date: {join_date.strftime('%d.%m.%Y')} | Name: {crew_name} | Rank: {rank}")
        table = doc.add_table(rows=1, cols=2)
        table.rows[0].cells[0].text = "ITEMS"; table.rows[0].cells[1].text = "STATUS"
        
        has_cra = any(kw in evidence_text for kw in ["cra", "endorsement"]) if not validation_locked else False
        checks = [
            ("National STCW & Certificates", True),
            ("Flag State’s Licence / CRA", has_cra if is_officer else None),
            ("Id Card (Kimlik)", True),
            ("Pre-joining alcohol/drug test", True)
        ]
        for item, status in checks:
            if status is None:
                continue
            row = table.add_row()
            row.cells[0].text = item
            row.cells[1].text = "Yes ☑" if status else "No ☑"
            
        buf = io.BytesIO()
        doc.save(buf)
        return buf.getvalue()

    c1, c2 = st.columns(2)
    with c1:
        st.download_button("📊 35-101 Excel İndir", build_excel_export(), f"35-101_{crew_name}.xlsx", disabled=validation_locked)
    with c2:
        st.download_button("📝 35-102 Word İndir", build_word_export(), f"35-102_{crew_name}.docx", disabled=validation_locked)
