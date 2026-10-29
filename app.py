import streamlit as st
import datetime
import requests
from bs4 import BeautifulSoup
import re
import io
import os
from dateutil.relativedelta import relativedelta
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment
import docx

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
st.caption("Doğrudan UAB Entegrasyonu & 35-101 / 35-102 Raporlama Motoru")

# =============================================================================
# KURUMSAL 35-101 MATRİS TANIMLARI
# =============================================================================
RANKS_MAP = {
    "Master": 6, "Chief officer": 7, "2 nd Officer": 8, "3 rd Officer": 9,
    "Chief Engineer": 10, "2 nd Engineer": 11, "3rd Engineer": 12, "Electrician": 13,
    "Bosun": 14, "Pumpman": 15, "AB": 16, "O/S": 17, "Fitter": 18,
    "Oiler": 19, "Wiper": 20, "Cook": 21, "Steward": 22, "Cadet": 23
}
OFFICER_RANKS = ["Master", "Chief officer", "2 nd Officer", "3 rd Officer", "Chief Engineer", "2 nd Engineer", "3rd Engineer", "Electrician"]
ALL_RANKS = list(RANKS_MAP.keys())

# UAB Türkçesine tam uyumlu anahtar kelimeler
CERT_MATRIX = [
    (14, "1", "Passport", ALL_RANKS, ["passport", "pasaport", "verified"]),
    (16, "2", "Medical Examination Report", ALL_RANKS, ["sağlık", "medical", "yoklaması"]),
    (18, "3", "Certificate of Yellow Fewer Vaccination", ALL_RANKS, ["yellow", "humma", "fever"]),
    (20, "4", "Seamanbook", ALL_RANKS, ["seaman", "cüzdan", "book", "gemiadamı"]),
    (22, "5", "Competency", ALL_RANKS, ["yeterlik", "competency", "coc"]),
    (24, "6", "GMDSS", ["Master", "Chief officer", "2 nd Officer", "3 rd Officer"], ["gmdss", "goc", "telsiz"]),
    (26, "7.1", "Fire Prevention and Fire Fighting Training Certificate", ALL_RANKS, ["yangın önleme", "fire prevention"]),
    (28, "7.2", "Personal Survival Techniques Training Certificate", ALL_RANKS, ["kişisel can", "survival techniques"]),
    (30, "7.3", "Elementary First Aid Training Certificate", ALL_RANKS, ["temel ilkyardım", "first aid"]),
    (32, "7.4", "Personal Safety and Social Responsibility Training Cert.", ALL_RANKS, ["sosyal sorumluluk", "social responsibility"]),
    (34, "7.5", "Proficiency in Survival Craft and Rescue Boats Training Cert.", ["Master", "Chief officer", "2 nd Officer", "3 rd Officer", "Chief Engineer", "2 nd Engineer", "3rd Engineer", "Electrician", "Bosun", "Pumpman", "AB", "Oiler"], ["cankurtarma araçlarını", "survival craft"]),
    (36, "8.1", "Security Awareness Certificate ", ALL_RANKS, ["farkındalık", "awareness"]),
    (38, "8.2", "Designated Security Duties Certificate", ALL_RANKS, ["güvenlik görevleri", "security duties"]),
    (40, "8.3", "Ship Security Officer", ["Master", "Chief officer", "2 nd Officer", "3 rd Officer", "Chief Engineer", "2 nd Engineer"], ["gemi güvenlik zabiti", "ship security officer"]),
    (42, "8.4", "Security - Realated Familirization Certificate", ALL_RANKS, ["tanıtım belgesi", "familiarization"]),
    (44, "9.1", "Advanced Fire Fighting Training Certificate", OFFICER_RANKS, ["ileri yangın", "advanced fire"]),
    (82, "5", "Endorsement", OFFICER_RANKS + ["Cadet"], ["endorsement", "onay belgesi"])
]

# =============================================================================
# SOL MENÜ: ZORUNLU ALANLAR
# =============================================================================
st.sidebar.header("🚢 Gemi ve Atama Parametreleri")
vessel_name = st.sidebar.text_input("Vessel Name *", value="", placeholder="Örn: Theo G").strip()
flag_options = ["MALTA", "MARSHALL ISLANDS", "PANAMA", "LIBERIA", "TURKIYE"]
vessel_flag = st.sidebar.selectbox("Vessel Flag", options=flag_options, index=None)
ecdis_type = st.sidebar.text_input("ECDIS Type", value="").strip()
rank = st.sidebar.selectbox("Rank *", options=ALL_RANKS, index=None)
nationality = st.sidebar.selectbox("Nationality *", options=["Turkish", "Georgian", "Egyptian", "Azerbaijan", "Indonesian"], index=0)
join_date = st.sidebar.date_input("Sign-On Date", datetime.date.today())

missing_fields = []
if not vessel_name: missing_fields.append("Vessel Name")
if not rank: missing_fields.append("Rank")

if missing_fields:
    st.info(f"👉 Sol menüdeki zorunlu alanları doldurunuz: **{', '.join(missing_fields)}**")
    st.stop()

rank_col = RANKS_MAP.get(rank, 17)
is_officer = rank in OFFICER_RANKS
required_contract_months = 4 if is_officer else 6

st.sidebar.markdown("---")
st.sidebar.markdown(f"⏱ **Asgari Kontrat Emniyeti:** `{required_contract_months} Ay`")

def locate_template(pattern_str, extension):
    target_clean = pattern_str.lower().replace("-", "").replace("_", "").replace(" ", "")
    for root, dirs, files in os.walk("."):
        for f in files:
            if f.lower().endswith(extension):
                f_clean = f.lower().replace("-", "").replace("_", "").replace(" ", "")
                if target_clean in f_clean:
                    return os.path.join(root, f)
    return None

template_excel_path = locate_template("35101", ".xlsx")
template_word_path = locate_template("35102", ".docx")

# =============================================================================
# ANA EKRAN: UAB DOĞRUDAN SORGULAMA
# =============================================================================
st.markdown("### 🔍 UAB Doğrudan Veri Çekme")
st.caption("PDF yüklemeye gerek yok. TC ve Cüzdan No girin, sistem UAB'den verileri otomatik çekip eşleştirsin.")

colA, colB, colC, colD = st.columns(4)
with colA:
    crew_name = st.text_input("Personel Adı Soyadı", value="")
with colB:
    tckn_val = st.text_input("T.C. Kimlik No *", value="")
with colC:
    sb_val = st.text_input("Liman Cüzdan No (SB) *", value="")
with colD:
    pass_val = st.text_input("Pasaport No", value="")

# Session state için verileri sakla
if "uab_data" not in st.session_state:
    st.session_state["uab_data"] = []
if "uab_raw_text" not in st.session_state:
    st.session_state["uab_raw_text"] = ""

def fetch_uab_data(tc, sb):
    url = f"https://gemiadami.uab.gov.tr/verification/{tc.strip()}/{sb.strip()}"
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/114.0.0.0 Safari/537.36'
    }
    extracted_rows = []
    all_text = ""
    try:
        response = requests.get(url, headers=headers, timeout=10)
        if response.status_code == 200:
            soup = BeautifulSoup(response.text, 'html.parser')
            all_text = soup.get_text(separator=" ", strip=True).upper()
            
            # Tablolardaki tüm satırları yakala
            for tr in soup.find_all('tr'):
                row_text = " | ".join([td.get_text(strip=True) for td in tr.find_all(['td', 'th'])])
                if row_text:
                    extracted_rows.append(row_text.upper())
            return extracted_rows, all_text, True, "Başarılı"
        else:
            return [], "", False, f"Sunucu Hatası: {response.status_code}"
    except Exception as e:
        return [], "", False, f"Bağlantı Hatası: {str(e)}"

if st.button("🌐 UAB Sisteminden Sorgula ve Raporla", type="primary"):
    if not tckn_val or not sb_val:
        st.error("TCKN ve Liman Cüzdan Numarası girmek zorunludur.")
    else:
        with st.spinner("UAB Sunucusuna bağlanılıyor ve veriler çekiliyor..."):
            rows, raw_text, success, msg = fetch_uab_data(tckn_val, sb_val)
            
            if success:
                # UAB doğrulama ekranında TC ve Cüzdan geçersizse genelde tablo boş döner
                if not rows or len(raw_text) < 100:
                    st.error("❌ UAB sunucusuna ulaşıldı ancak kayıt bulunamadı. Lütfen TC ve Cüzdan No'yu kontrol ediniz.")
                else:
                    st.success("✅ UAB Verileri Başarıyla Çekildi!")
                    st.session_state["uab_data"] = rows
                    st.session_state["uab_raw_text"] = raw_text
            else:
                st.error(f"❌ UAB sistemine otomatik bot girişi şu an engelleniyor. Hata: {msg}")
                st.info("Eğer bu uyarıyı alıyorsanız, devlet sitesi bot koruması aktif etmiştir. Bu durumda sistemin UAB sayfasını doğrudan tarayıcınızdan okuyabilmesi için 'Chrome Eklentisi (Extension)' yöntemine geçmemiz gerekecek.")

# =============================================================================
# EXCEL / WORD DOLDURMA (UAB VERİLERİYLE)
# =============================================================================
if st.session_state.get("uab_data") or st.session_state.get("uab_raw_text"):
    
    st.divider()
    st.markdown("### 📥 Çıktıları İndir")
    
    # 1. 35-101 EXCEL ŞABLONUNU DOLDURMA
    def build_excel_from_template():
        if not template_excel_path or not os.path.exists(template_excel_path):
            wb = openpyxl.Workbook()
            ws1 = wb.active
            ws1['A1'] = "HATA: Repoda '35-101' içeren Excel şablonu bulunamadı!"
            buf = io.BytesIO()
            wb.save(buf)
            return buf.getvalue()

        wb = openpyxl.load_workbook(template_excel_path)
        ws1 = wb.active
        
        ws1['C6'] = vessel_name
        ws1['C8'] = join_date.strftime("%d.%m.%Y")
        ws1['N6'] = crew_name if crew_name else "BİLGİ GİRİLMEDİ"
        ws1['N8'] = rank
        
        for coord in ['C6', 'C8', 'N6', 'N8']:
            ws1[coord].font = Font(name="Calibri", size=10, bold=True, color="002060")
            
        yellow_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
        green_fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
        red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
        intersect_green = PatternFill(start_color="C6E0B4", end_color="C6E0B4", fill_type="solid")
        intersect_red = PatternFill(start_color="FF9999", end_color="FF9999", fill_type="solid")
        intersect_yellow = PatternFill(start_color="FFE699", end_color="FFE699", fill_type="solid")
        
        for r in range(14, 85):
            c_cell = ws1.cell(r, rank_col)
            if type(c_cell) is not openpyxl.cell.cell.MergedCell:
                c_cell.fill = yellow_fill

        uab_rows = st.session_state["uab_data"]
        raw_text = st.session_state["uab_raw_text"]

        for r_idx, no_str, en_title, req_ranks, keywords in CERT_MATRIX:
            is_req = rank in req_ranks
            r_desc = r_idx + 1
            
            if is_req:
                # UAB Satırlarında Sertifikayı Ara
                matching_row = None
                for row_text in uab_rows:
                    if any(kw in row_text for kw in keywords):
                        matching_row = row_text
                        break
                
                # Cüzdan, Pasaport veya Sağlık Raporu UAB tablosunda listelenmemişse metin genelinde var mı bak
                is_found = False
                if matching_row:
                    is_found = True
                else:
                    if any(kw in raw_text for kw in keywords):
                        is_found = True
                    # Cüzdan veya pasaport nosu ekrandan girilmişse var kabul et
                    if "seaman" in en_title.lower() and sb_val: is_found = True
                    if "pass" in en_title.lower() and pass_val: is_found = True

                if is_found:
                    doc_no_val = "Verified in UAB"
                    if "seaman" in en_title.lower(): doc_no_val = sb_val
                    if "pass" in en_title.lower() and pass_val: doc_no_val = pass_val
                    
                    issue_str, expiry_str = "", ""
                    remaining_months = 60
                    
                    # Eğer UAB'de tablo satırı olarak bulunduysa, içinden tarihleri ve belge nosunu çek (Örn: 22.07.2024)
                    if matching_row:
                        # Tarih formati arama (GG.AA.YYYY veya GG/AA/YYYY)
                        dates = re.findall(r'\b\d{2}[./]\d{2}[./]\d{4}\b', matching_row)
                        if len(dates) >= 2:
                            issue_str = dates[0].replace("/", ".")
                            expiry_str = dates[1].replace("/", ".")
                        elif len(dates) == 1:
                            issue_str = dates[0].replace("/", ".")
                            expiry_str = "Süresiz"
                            
                        # Belge numarası tahminleme (Tarih harici harf+rakam dizisi)
                        tokens = re.findall(r'\b[A-Z0-9]{5,15}\b', matching_row.replace(".", ""))
                        tokens = [t for t in tokens if not re.match(r'^\d{8}$', t)] # Sadece tarih olanları ele
                        if tokens:
                            doc_no_val = tokens[0]

                    if expiry_str and expiry_str != "Süresiz":
                        try:
                            exp_date = datetime.datetime.strptime(expiry_str, "%d.%m.%Y").date()
                            remaining_months = (exp_date.year - join_date.year) * 12 + (exp_date.month - join_date.month)
                        except: pass

                    if remaining_months < required_contract_months:
                        status_str = f"Doc No: {doc_no_val} | Exp: {expiry_str} | ► YETERSİZ SÜRE! ({remaining_months} Ay Kaldı)"
                        chosen_fill = red_fill
                        inter_fill = intersect_red
                    elif remaining_months < 12:
                        status_str = f"Doc No: {doc_no_val} | Exp: {expiry_str} | DİKKAT: <1 YIL ({remaining_months} Ay)"
                        chosen_fill = yellow_fill
                        inter_fill = intersect_yellow
                    else:
                        status_str = f"Doc No: {doc_no_val} | Exp: {expiry_str} | Geçerli ({remaining_months} Ay)"
                        if expiry_str == "Süresiz":
                             status_str = f"Doc No: {doc_no_val} | Exp: Süresiz | Geçerli"
                        chosen_fill = green_fill
                        inter_fill = intersect_green

                    ws1.cell(r_desc, 2).value = status_str
                    ws1.cell(r_desc, 2).font = Font(name="Calibri", size=8.5, italic=True, color="1F3864")
                    
                    ws1.cell(r_idx, 2).fill = chosen_fill
                    ws1.cell(r_desc, 2).fill = chosen_fill
                    
                    c_inter = ws1.cell(r_idx, rank_col)
                    if type(c_inter) is not openpyxl.cell.cell.MergedCell:
                        c_inter.fill = inter_fill
                        c_inter.value = "X"
                        c_inter.alignment = Alignment(horizontal="center", vertical="center")
                    c_inter2 = ws1.cell(r_desc, rank_col)
                    if type(c_inter2) is not openpyxl.cell.cell.MergedCell:
                        c_inter2.fill = inter_fill
                else:
                    ws1.cell(r_desc, 2).value = "► MISSING / EKSİK EVRAK (UAB Sisteminde Bulunamadı)!"
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

    # 2. 35-102 WORD ŞABLONUNU DOLDURMA
    def build_word_from_template():
        if not template_word_path or not os.path.exists(template_word_path):
            doc = docx.Document()
            doc.add_paragraph("HATA: Repoda '35-102' içeren Word şablonu bulunamadı!")
            buf = io.BytesIO()
            doc.save(buf)
            return buf.getvalue()

        doc = docx.Document(template_word_path)
        table = doc.tables[0]
        
        table.rows[0].cells[3].text = vessel_name
        table.rows[0].cells[7].text = join_date.strftime("%d.%m.%Y")
        table.rows[1].cells[3].text = crew_name if crew_name else "BİLGİ GİRİLMEDİ"
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

        has_cra = "ENDORSEMENT" in st.session_state["uab_raw_text"] or "ONAY BELGESI" in st.session_state["uab_raw_text"]
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
        st.download_button("📊 35-101 Excel İndir", build_excel_from_template(), f"35-101_{crew_name}.xlsx", disabled=False)
    with c2:
        st.download_button("📝 35-102 Word İndir", build_word_from_template(), f"35-102_{crew_name}.docx", disabled=False)
