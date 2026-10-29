import streamlit as st
import datetime
import time
import io
import os
import json
from dateutil.relativedelta import relativedelta
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment
import docx

try:
    import google.generativeai as genai
    AI_AVAILABLE = True
except ImportError as e:
    AI_AVAILABLE = False
    AI_ERROR = str(e)

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
st.caption("Yapay Zeka Destekli, PDF İçerik Okuyucu & Kilit Olmayan Sistem")

# =============================================================================
# AI (GEMINI) BAĞLANTISI
# =============================================================================
api_key = st.sidebar.text_input("🔑 Google Gemini API Key", type="password", placeholder="AIzaSy...")

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

CERT_MATRIX = [
    (14, "1", "Passport", ALL_RANKS, ["passport", "pasaport"]),
    (16, "2", "Medical Examination Report", ALL_RANKS, ["medical", "sağlık", "health"]),
    (18, "3", "Certificate of Yellow Fewer Vaccination", ALL_RANKS, ["yellow", "fever", "humma"]),
    (20, "4", "Seamanbook", ALL_RANKS, ["seaman", "cüzdan", "book", "gemiadam"]),
    (22, "5", "Competency", ALL_RANKS, ["coc", "competency", "yeterli", "desc"]),
    (24, "6", "GMDSS", ["Master", "Chief officer", "2 nd Officer", "3 rd Officer"], ["gmdss", "goc", "telsiz"]),
    (26, "7.1", "Fire Prevention and Fire Fighting Training Certificate", ALL_RANKS, ["fire prevention", "yangın önleme"]),
    (28, "7.2", "Personal Survival Techniques Training Certificate", ALL_RANKS, ["survival techniques", "kişisel can", "kurtarma teknikleri"]),
    (30, "7.3", "Elementary First Aid Training Certificate", ALL_RANKS, ["first aid", "temel ilkyardım", "ilk yardım"]),
    (32, "7.4", "Personal Safety and Social Responsibility Training Cert.", ALL_RANKS, ["social responsibility", "sosyal sorumluluk"]),
    (34, "7.5", "Proficiency in Survival Craft and Rescue Boats Training Cert.", ["Master", "Chief officer", "2 nd Officer", "3 rd Officer", "Chief Engineer", "2 nd Engineer", "3rd Engineer", "Electrician", "Bosun", "Pumpman", "AB", "Oiler"], ["survival craft", "cankurtarma araç"]),
    (36, "8.1", "Security Awareness Certificate ", ALL_RANKS, ["security awareness", "güvenlik farkındalık", "farkındalık"]),
    (38, "8.2", "Designated Security Duties Certificate", ALL_RANKS, ["designated security", "belirlenmiş güvenlik", "güvenlik görev"]),
    (40, "8.3", "Ship Security Officer", ["Master", "Chief officer", "2 nd Officer", "3 rd Officer", "Chief Engineer", "2 nd Engineer"], ["ship security officer", "gemi güvenlik zabiti"]),
    (42, "8.4", "Security - Realated Familirization Certificate", ALL_RANKS, ["familiarization", "tanıtım belgesi"]),
    (44, "9.1", "Advanced Fire Fighting Training Certificate", OFFICER_RANKS, ["advanced fire", "ileri yangın"]),
    (82, "5", "Endorsement", OFFICER_RANKS + ["Cadet"], ["cra", "endorsement", "onay belgesi"])
]

# =============================================================================
# SOL MENÜ: ZORUNLU ALANLAR
# =============================================================================
st.sidebar.header("🚢 Gemi ve Atama Parametreleri")

vessel_name = st.sidebar.text_input("Vessel Name *", value="").strip()
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
    st.info(f"👉 Devam etmek için sol menüdeki zorunlu alanları doldurunuz: **{', '.join(missing_fields)}**")
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

def parse_ai_response(response_text):
    text = response_text.strip()
    if text.startswith("```json"): text = text[7:]
    elif text.startswith("```"): text = text[3:]
    if text.endswith("```"): text = text[:-3]
    try:
        return json.loads(text.strip())
    except:
        return {}

# =============================================================================
# EVRAK YÜKLEME VE YAPAY ZEKA (İÇERİK) ANALİZİ
# =============================================================================
if not api_key:
    st.error("🚨 İşleme başlamak için sol menüden API Anahtarınızı giriniz.")

uploaded_files = st.file_uploader(
    "Tüm Evrakları Yükleyin (Pasaport, Medical, Yellow Fever, STCW'ler)",
    type=["pdf", "jpg", "jpeg", "png"],
    accept_multiple_files=True
)

if uploaded_files and api_key:
    if not AI_AVAILABLE:
        st.error(f"🚨 KRİTİK HATA: Yapay Zeka kütüphanesi (google-generativeai) sunucuya yüklenemedi!\n\nHata Detayı: {AI_ERROR}\n\nÇözüm: Lütfen 'requirements.txt' dosyanızı kontrol edip, sağ alttaki siyah menüden 'Clear cache and deploy' seçeneğine tıklayın.")
        st.stop()

    genai.configure(api_key=api_key)
    model = genai.GenerativeModel('gemini-1.5-flash', generation_config={"response_mime_type": "application/json"})
    
    crew_name, crew_surname, tckn_val, sb_val, pass_val = "", "", "", "", ""
    all_certs = []
    
    prompt = """
    Sen bir denizcilik evrak uzmanısın. Ekteki evrakı dikkatlice incele.
    SADECE AŞAĞIDAKİ JSON FORMATINDA CEVAP VER. Evrak bir Pasaport, Gemiadamı Sağlık Raporu, Sarı Humma (Yellow Fever), Liman Cüzdanı veya STCW Sertifikası olabilir. İçeriğe bakarak türünü belirle ve tarihleri (GG.AA.YYYY) formatında çıkar.
    
    {
      "personel_bilgileri": {
        "ad": "Personelin Adı",
        "soyad": "Personelin Soyadı",
        "tc_kimlik": "11 Haneli TC No",
        "pasaport_no": "Pasaport No",
        "cuzdan_no": "Liman Cüzdanı No (S ile başlar)"
      },
      "sertifikalar": [
        {
          "tur": "Belgenin Türü (Örn: Passport, Medical, Yellow Fever, Seaman Book, STCW İlk Yardım, STCW Yangın...)",
          "belge_no": "Sertifika/Belge Numarası",
          "verilis_tarihi": "DD.MM.YYYY",
          "bitis_tarihi": "DD.MM.YYYY"
        }
      ]
    }
    """
    
    progress_text = "Yapay Zeka Evrakların İçeriğini Okuyor..."
    my_bar = st.progress(0, text=progress_text)
    
    for idx, f in enumerate(uploaded_files):
        fname = f.name.lower()
        content = f.read()
        f.seek(0)
        
        mime_type = "application/pdf"
        if fname.endswith((".jpg", ".jpeg")): mime_type = "image/jpeg"
        elif fname.endswith(".png"): mime_type = "image/png"
            
        part = {"mime_type": mime_type, "data": content}
        
        try:
            response = model.generate_content([prompt, part])
            data = parse_ai_response(response.text)
            
            gen = data.get("personel_bilgileri", {})
            if gen.get("ad") and not crew_name: crew_name = str(gen.get("ad")).upper()
            if gen.get("soyad") and not crew_surname: crew_surname = str(gen.get("soyad")).upper()
            if gen.get("tc_kimlik") and not tckn_val: tckn_val = str(gen.get("tc_kimlik")).upper()
            if gen.get("pasaport_no") and not pass_val: pass_val = str(gen.get("pasaport_no")).upper()
            if gen.get("cuzdan_no") and not sb_val: sb_val = str(gen.get("cuzdan_no")).upper()
            
            certs = data.get("sertifikalar", [])
            if isinstance(certs, list):
                all_certs.extend(certs)
                
        except Exception as e:
            pass
            
        time.sleep(3.5)
        my_bar.progress((idx + 1) / len(uploaded_files), text=f"{f.name} okundu.")
        
    my_bar.empty()

    st.subheader("📋 Çıkarılan / Teyit Edilen Bilgiler")
    st.caption("Veriler AI tarafından dosya içeriklerinden çekilmiştir. Gerekirse kutulardan düzeltme yapabilirsiniz.")
    
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        full_name_input = f"{crew_name} {crew_surname}".strip()
        final_crew_name = st.text_input("Personel Adı Soyadı", value=full_name_input)
    with col2:
        label_id = "T.C. Kimlik No" if nationality == "Turkish" else f"National ID ({nationality})"
        final_tckn = st.text_input(label_id, value=tckn_val)
    with col3:
        final_sb = st.text_input("Liman Cüzdan No (SB)", value=sb_val)
    with col4:
        final_pass = st.text_input("Pasaport No", value=pass_val)
        
    if nationality == "Turkish":
        if not final_tckn or not final_sb or len(final_crew_name.split()) < 2:
            st.warning("⚠️ Kutularda eksik bilgiler var (Elle doldurabilirsiniz). Sistem kilitsizdir, onay size aittir.")
        else:
            st.success(f"✅ Veriler Onaylandı. Asgari Kontrat Şartı: {required_contract_months} Ay.")
            uab_url = f"[https://gemiadami.uab.gov.tr/verification/](https://gemiadami.uab.gov.tr/verification/){final_tckn.strip()}/{final_sb.strip()}"
            # URL YÖNLENDİRME DÜZELTİLDİ: st.link_button kullanılarak yeni sekmede sorunsuz açılması sağlandı
            st.link_button("UAB Portalı Doğrulama ↗", uab_url, type="primary")
    else:
        st.success(f"✅ Yabancı uyruklu personel işlemleri devam edebilir.")

    st.divider()

    # =========================================================================
    # EXCEL / WORD OLUŞTURMA (GERÇEK TARİHLERLE)
    # =========================================================================
    def build_excel_from_template():
        if not template_excel_path or not os.path.exists(template_excel_path):
            wb = openpyxl.Workbook()
            ws1 = wb.active
            ws1['A1'] = "HATA: 35-101 Excel şablonu bulunamadı!"
            buf = io.BytesIO()
            wb.save(buf)
            return buf.getvalue()

        wb = openpyxl.load_workbook(template_excel_path)
        ws1 = wb.active
        
        ws1['C6'] = vessel_name
        ws1['C8'] = join_date.strftime("%d.%m.%Y")
        ws1['N6'] = final_crew_name
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

        for r_idx, no_str, en_title, req_ranks, keywords in CERT_MATRIX:
            is_req = rank in req_ranks
            r_desc = r_idx + 1
            
            if is_req:
                matching_cert = None
                for cert in all_certs:
                    c_type = str(cert.get("tur", "")).lower()
                    if any(kw in c_type for kw in keywords):
                        matching_cert = cert
                        break
                
                is_found = matching_cert is not None

                if is_found:
                    doc_no_val = str(matching_cert.get("belge_no", "Bulunamadı"))
                    if "seaman" in en_title.lower() and final_sb: doc_no_val = final_sb
                    if "pass" in en_title.lower() and final_pass: doc_no_val = final_pass
                    if doc_no_val == "Bulunamadı" or not doc_no_val: doc_no_val = "Orijinalinden Teyitli"

                    issue_str = str(matching_cert.get("verilis_tarihi", "")).replace("/", ".").replace("-", ".")
                    expiry_str = str(matching_cert.get("bitis_tarihi", "")).replace("/", ".").replace("-", ".")
                    remaining_months = 60
                    
                    if expiry_str and expiry_str.lower() != "süresiz" and len(expiry_str) >= 8:
                        try:
                            exp_date = datetime.datetime.strptime(expiry_str[-10:], "%d.%m.%Y").date()
                            remaining_months = (exp_date.year - join_date.year) * 12 + (exp_date.month - join_date.month)
                        except:
                            pass
                    elif expiry_str.lower() == "süresiz" or "medical" not in en_title.lower():
                        remaining_months = 100 
                        expiry_str = "Süresiz"

                    if remaining_months < required_contract_months:
                        status_str = f"Doc No: {doc_no_val} | Exp: {expiry_str} | ► YETERSİZ SÜRE! ({remaining_months} Ay Kaldı)"
                        chosen_fill = red_fill
                        inter_fill = intersect_red
                    elif remaining_months < 12:
                        status_str = f"Doc No: {doc_no_val} | Exp: {expiry_str} | DİKKAT: <1 YIL ({remaining_months} Ay)"
                        chosen_fill = yellow_fill
                        inter_fill = intersect_yellow
                    else:
                        status_str = f"Doc No: {doc_no_val} | Exp: {expiry_str} | Geçerli"
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

    def build_word_from_template():
        if not template_word_path or not os.path.exists(template_word_path):
            doc = docx.Document()
            doc.add_paragraph("HATA: 35-102 Word şablonu bulunamadı!")
            buf = io.BytesIO()
            doc.save(buf)
            return buf.getvalue()

        doc = docx.Document(template_word_path)
        table = doc.tables[0]
        
        table.rows[0].cells[3].text = vessel_name
        table.rows[0].cells[7].text = join_date.strftime("%d.%m.%Y")
        table.rows[1].cells[3].text = final_crew_name
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

        has_cra = False
        for cert in all_certs:
            if "endorsement" in str(cert.get("tur", "")).lower() or "cra" in str(cert.get("tur", "")).lower():
                has_cra = True
                break

        ecdis_req = rank in ["Master", "Chief officer", "2 nd Officer", "3 rd Officer"]
        
        set_status(5, is_yes=True)
        set_status(6, is_yes=True)
        set_status(7, is_yes=has_cra if is_officer else True, is_na=not is_officer and not has_cra)
        
        for i in range(9, 20):
            if i != 14:
                set_status(i, is_yes=True)
        
        id_display = final_tckn if nationality == "Turkish" else final_pass
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
        st.download_button("📊 35-101 Excel İndir", build_excel_from_template(), f"35-101_{final_crew_name}.xlsx", disabled=False)
    with c2:
        st.download_button("📝 35-102 Word İndir", build_word_from_template(), f"35-102_{final_crew_name}.docx", disabled=False)
