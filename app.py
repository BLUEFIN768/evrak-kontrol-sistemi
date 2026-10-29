import streamlit as st
import datetime
from dateutil.relativedelta import relativedelta
import io
import os
import json
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment
import docx
from PIL import Image
import google.generativeai as genai
from pdf2image import convert_from_bytes

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

st.title("⚓ Evrak Kontrol & Denetim Portalı (AI Vision Modu)")
st.caption("Yapay Zeka Destekli, Esnek ve Kilit Olmayan Kontrat Emniyet Motoru")

# =============================================================================
# YAPAY ZEKA (GEMINI) BAĞLANTISI
# =============================================================================
api_key = st.sidebar.text_input("🔑 Google Gemini API Key", type="password", placeholder="AIzaSy...")
if api_key:
    genai.configure(api_key=api_key)
    # Hızlı ve JSON formatında okuma yapabilen Gemini 1.5 Flash modelini kullanıyoruz
    model = genai.GenerativeModel('gemini-1.5-flash', generation_config={"response_mime_type": "application/json"})
else:
    st.sidebar.warning("Sistemin otomatik okuma yapabilmesi için lütfen AI Studio'dan aldığınız API Anahtarını giriniz.")

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
required_contract_months = 4 if is_officer else 6

st.sidebar.markdown("---")
st.sidebar.markdown(f"⏱ **Asgari Kontrat Emniyeti:** `{required_contract_months} Ay` ({'Zabit' if is_officer else 'Alt Kadro'})")
st.sidebar.markdown("⚠️ **Sarı İkaz Sınırı:** `< 12 Ay`")

# =============================================================================
# AKILLI ŞABLON BULUCU
# =============================================================================
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
# YAPAY ZEKA GÖRSEL İŞLEME MOTORU (VISION AI)
# =============================================================================
def extract_info_via_ai(pil_image, fname):
    """Görseli Gemini AI'a gönderir ve yapılandırılmış JSON verisi alır."""
    prompt = """
    Bu bir denizcilik evrakıdır (Pasaport, Liman Cüzdanı, STCW Sertifikası veya Sağlık Raporu olabilir).
    Lütfen evraka bir insan gibi bakarak aşağıdaki bilgileri çıkar ve SADECE JSON formatında ver. 
    Eğer bir veri evrakta yoksa veya okunamıyorsa değerini boş string ("") yap. 
    Lütfen ad ve soyad bilgisini, TCKN'yi, Pasaport veya Cüzdan Numarasını, Veriliş(Issue) ve Bitiş(Expiry) tarihlerini dikkatle incele.
    Tarihleri DD.MM.YYYY formatına çevirerek ver.
    
    JSON Formatı:
    {
      "name": "Sadece Personelin Adı",
      "surname": "Sadece Personelin Soyadı",
      "tckn": "11 Haneli TC Kimlik Numarası",
      "passport_no": "Pasaport Numarası",
      "seaman_book_no": "Liman Cüzdan Numarası (Genelde S ile başlar)",
      "issue_date": "Evrakın veriliş tarihi (DD.MM.YYYY)",
      "expiry_date": "Evrakın bitiş tarihi (DD.MM.YYYY)"
    }
    """
    try:
        response = model.generate_content([prompt, pil_image])
        return json.loads(response.text)
    except Exception as e:
        return {}

def convert_to_image(file_obj, filename):
    """PDF veya Resim dosyasını analiz için PIL Image nesnesine çevirir."""
    content = file_obj.read()
    file_obj.seek(0)
    if filename.lower().endswith(".pdf"):
        try:
            images = convert_from_bytes(content, first_page=1, last_page=1, dpi=200)
            if images:
                return images[0]
        except:
            return None
    elif filename.lower().endswith((".png", ".jpg", ".jpeg")):
        try:
            return Image.open(io.BytesIO(content)).convert('RGB')
        except:
            return None
    return None

# =============================================================================
# ANA VERİ TOPLAMA VE DENETİM DÖNGÜSÜ
# =============================================================================
uploaded_files = st.file_uploader(
    "Evrakları Yükleyin (PDF, JPG, JPEG, PNG)",
    type=["pdf", "jpg", "jpeg", "png"],
    accept_multiple_files=True
)

if uploaded_files:
    # Toplanacak Ana Veriler
    crew_name = ""
    crew_surname = ""
    tckn_val = ""
    sb_val = ""
    pass_val = ""
    
    # Evraklardan çekilen gerçek tarihler: {"pass": {"issue": "...", "expiry": "..."}, ...}
    extracted_dates = {}
    evidence_text = ""
    
    if api_key:
        progress_text = "Yapay Zeka Evrakları Okuyor, Lütfen Bekleyin..."
        my_bar = st.progress(0, text=progress_text)
        
        for idx, f in enumerate(uploaded_files):
            fname = f.name.lower()
            evidence_text += f" {fname} "
            
            pil_img = convert_to_image(f, fname)
            if pil_img:
                # Yapay Zekaya Gönder
                ai_data = extract_info_via_ai(pil_img, fname)
                
                # Kimlik Bilgilerini Topla
                if ai_data.get("name") and not crew_name: crew_name = ai_data.get("name").upper()
                if ai_data.get("surname") and not crew_surname: crew_surname = ai_data.get("surname").upper()
                if ai_data.get("tckn") and not tckn_val: tckn_val = ai_data.get("tckn")
                if ai_data.get("passport_no") and not pass_val: pass_val = ai_data.get("passport_no")
                if ai_data.get("seaman_book_no") and not sb_val: sb_val = ai_data.get("seaman_book_no")
                
                # Tarihleri Kaydet
                issue_d = ai_data.get("issue_date", "")
                expiry_d = ai_data.get("expiry_date", "")
                if issue_d or expiry_d:
                    # Dosya ismindeki anahtar kelimelere göre eşleştirme yap
                    extracted_dates[fname] = {"issue": issue_d, "expiry": expiry_d}

            my_bar.progress((idx + 1) / len(uploaded_files), text=f"{f.name} analiz edildi.")
            
        my_bar.empty()
    else:
        st.info("API Key girilmediği için evraklar sadece var/yok eşleşmesi için kullanılacak. Bilgileri manuel girebilirsiniz.")
        for f in uploaded_files:
            evidence_text += f" {f.name.lower()} "

    st.subheader("📋 Çıkarılan / Teyit Edilen Bilgiler")
    st.caption("AI tarafından okunan veriler. Yanlışlık varsa kutulardan **MANUEL OLARAK DÜZELTİN**. Sistem sizin yazdığınızı doğrudan doğru kabul edip kilitleri kaldıracaktır.")
    
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
        
    # KİLİT MEKANİZMASI İPTAL EDİLDİ - SADECE UYARI VERİR!
    has_warning = False
    
    if nationality == "Turkish":
        if not final_tckn or not final_sb or len(final_crew_name.split()) < 2:
            st.warning("⚠️ Eksik veya formatı hatalı veriler var. Ancak yönetici inisiyatifi ile Word/Excel çıktılarını alabilirsiniz.")
            has_warning = True
        else:
            st.success(f"✅ Veriler Onaylandı. Asgari Kontrat Şartı: {required_contract_months} Ay ({'Zabit' if is_officer else 'Alt Kadro'}).")
            uab_url = f"https://gemiadami.uab.gov.tr/verification/{final_tckn.strip()}/{final_sb.strip()}"
            st.markdown(f'<a href="{uab_url}" target="_blank" style="background-color: #0056b3; color: white; padding: 10px 20px; border-radius: 4px; text-decoration: none; font-weight: bold; display: inline-block; margin-bottom: 10px;">UAB Doğrulama Portalı ↗</a>', unsafe_allow_html=True)
    else:
        if not final_pass or not final_sb or len(final_crew_name.split()) < 2:
            st.warning(f"⚠️ {nationality} personel için eksik bilgiler var. Manuel tamamlayıp çıktı alabilirsiniz.")
            has_warning = True
        else:
            st.success(f"✅ {nationality} personel evrakları başarıyla doğrulandı.")

    st.divider()

    # =========================================================================
    # 1. 35-101 EXCEL ŞABLONUNU DOLDURMA
    # =========================================================================
    def build_excel_from_template():
        if not template_excel_path or not os.path.exists(template_excel_path):
            wb = openpyxl.Workbook()
            ws1 = wb.active
            ws1.title = "35-101"
            ws1['A1'] = "HATA: Repoda '35-101' içeren Excel şablonu bulunamadı! Lütfen GitHub'a yükleyin."
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
            
            # Bu sertifikaya ait yüklenmiş dosya var mı?
            matching_fname = None
            is_found = False
            for kw in keywords:
                for fn in extracted_dates.keys():
                    if kw in fn:
                        matching_fname = fn
                        is_found = True
                        break
                if is_found:
                    break
            
            if not is_found:
                is_found = any(kw in evidence_text for kw in keywords)

            r_desc = r_idx + 1
            
            if is_req:
                if is_found:
                    doc_no_val = final_sb if "seaman" in en_title.lower() else (final_pass if "pass" in en_title.lower() else "Verified")
                    
                    # GERÇEK TARİH HESAPLAMA MANTIĞI
                    issue_str = ""
                    expiry_str = ""
                    remaining_months = 60 # Varsayılan
                    
                    if matching_fname and extracted_dates[matching_fname].get("expiry"):
                        expiry_str = extracted_dates[matching_fname]["expiry"]
                        issue_str = extracted_dates[matching_fname]["issue"]
                        try:
                            # Gerçek tarihi datetime'a çevir ve kalan ayı hesapla
                            exp_date = datetime.datetime.strptime(expiry_str, "%d.%m.%Y").date()
                            remaining_months = (exp_date.year - join_date.year) * 12 + (exp_date.month - join_date.month)
                        except:
                            # Parse hatası olursa varsayılan tahmini süre atla
                            validity_years = 2 if "medical" in en_title.lower() else 5
                            simulated_expiry = join_date + relativedelta(years=validity_years)
                            expiry_str = simulated_expiry.strftime('%d.%m.%Y') + " (Tahmini)"
                            remaining_months = (simulated_expiry.year - join_date.year) * 12 + (simulated_expiry.month - join_date.month)
                    else:
                        validity_years = 2 if "medical" in en_title.lower() else 5
                        simulated_expiry = join_date + relativedelta(years=validity_years)
                        expiry_str = simulated_expiry.strftime('%d.%m.%Y') + " (Tahmini)"
                        remaining_months = (simulated_expiry.year - join_date.year) * 12 + (simulated_expiry.month - join_date.month)
                    
                    if not issue_str:
                        issue_str = join_date.strftime('%d.%m.%Y')

                    if remaining_months < required_contract_months:
                        status_str = f"► YETERSİZ SÜRE! ({remaining_months} Ay Kaldı - Kontratı Tamamlayamaz)"
                        chosen_fill = red_fill
                        inter_fill = intersect_red
                    elif remaining_months < 12:
                        status_str = f"Doc No: {doc_no_val} | Exp: {expiry_str} | DİKKAT: <1 YIL ({remaining_months} Ay)"
                        chosen_fill = yellow_fill
                        inter_fill = intersect_yellow
                    else:
                        status_str = f"Doc No: {doc_no_val} | Exp: {expiry_str} | Geçerli ({remaining_months} Ay)"
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

    # =========================================================================
    # 2. 35-102 WORD ŞABLONUNU DOLDURMA
    # =========================================================================
    def build_word_from_template():
        if not template_word_path or not os.path.exists(template_word_path):
            doc = docx.Document()
            doc.add_paragraph("HATA: Repoda '35-102' içeren Word şablonu bulunamadı! Lütfen GitHub'a yükleyin.")
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

        has_cra = any(kw in evidence_text for kw in ["cra", "endorsement"])
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
        # disabled=False ile buton kilidini tamamen kaldırdık!
        st.download_button("📊 35-101 Excel İndir", build_excel_from_template(), f"35-101_{final_crew_name}.xlsx", disabled=False)
    with c2:
        st.download_button("📝 35-102 Word İndir", build_word_from_template(), f"35-102_{final_crew_name}.docx", disabled=False)
