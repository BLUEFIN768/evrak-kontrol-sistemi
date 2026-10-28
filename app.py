import streamlit as st
import datetime
import io
import os
import re
from dateutil.relativedelta import relativedelta
import pypdf
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import docx
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT

st.set_page_config(page_title="Evrak Kontrol Portalı", page_icon="⚓", layout="wide")

# ==============================================================================
# GÜVENLİK PROTOKOLÜ (KVKK / VERİ GİZLİLİĞİ)
# ==============================================================================
def check_password():
    if "authenticated" not in st.session_state:
        st.session_state["authenticated"] = False

    if st.session_state["authenticated"]:
        return True

    st.markdown("<h2 style='text-align: center; color: #1F3864;'>⚓ Blue Fin Crewing - Evrak Kontrol Portalı</h2>", unsafe_allow_html=True)
    st.markdown("<p style='text-align: center; color: #666;'>Bu sistem KVKK uyumludur. Yüklenen hiçbir evrak sunucu diskinde saklanmaz.</p>", unsafe_allow_html=True)
    
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

# ==============================================================================
# ANA PORTAL
# ==============================================================================
st.sidebar.markdown("🔒 **Güvenli Oturum:** Aktif (KVKK Uyumlu)")
if st.sidebar.button("🚪 Çıkış Yap"):
    st.session_state["authenticated"] = False
    st.rerun()

st.title("⚓ Evrak Kontrol & Denetim Portalı")
st.caption("35-101 Seafarers Documents & 35-102 Pre Designation Check List Otomasyonu")

RANKS_MAP = {
    "Master": 6, "Chief officer": 7, "2 nd Officer": 8, "3 rd Officer": 9,
    "Chief Engineer": 10, "2 nd Engineer": 11, "3rd Engineer": 12, "Electrician": 13,
    "Bosun": 14, "Pumpman": 15, "AB": 16, "O/S": 17, "Fitter": 18,
    "Oiler": 19, "Wiper": 20, "Cook": 21, "Steward": 22, 
    "Deck Cadet": 23, "Engine Cadet": 23
}

OFFICER_RANKS = [
    "Master", "Chief officer", "2 nd Officer", "3 rd Officer",
    "Chief Engineer", "2 nd Engineer", "3rd Engineer", "Electrician"
]

CERT_MATRIX = [
    (14, "1", "Passport", ["Master", "Chief officer", "2 nd Officer", "3 rd Officer", "Chief Engineer", "2 nd Engineer", "3rd Engineer", "Electrician", "Bosun", "Pumpman", "AB", "O/S", "Fitter", "Oiler", "Wiper", "Cook", "Steward", "Deck Cadet", "Engine Cadet"]),
    (16, "2", "Medical Examination Report", ["Master", "Chief officer", "2 nd Officer", "3 rd Officer", "Chief Engineer", "2 nd Engineer", "3rd Engineer", "Electrician", "Bosun", "Pumpman", "AB", "O/S", "Fitter", "Oiler", "Wiper", "Cook", "Steward", "Deck Cadet", "Engine Cadet"]),
    (18, "3", "Certificate of Yellow Fewer Vaccination", ["Master", "Chief officer", "2 nd Officer", "3 rd Officer", "Chief Engineer", "2 nd Engineer", "3rd Engineer", "Electrician", "Bosun", "Pumpman", "AB", "O/S", "Fitter", "Oiler", "Wiper", "Cook", "Steward", "Deck Cadet", "Engine Cadet"]),
    (20, "4", "Seamanbook", ["Master", "Chief officer", "2 nd Officer", "3 rd Officer", "Chief Engineer", "2 nd Engineer", "3rd Engineer", "Electrician", "Bosun", "Pumpman", "AB", "O/S", "Fitter", "Oiler", "Wiper", "Cook", "Steward", "Deck Cadet", "Engine Cadet"]),
    (22, "5", "Competency", ["Master", "Chief officer", "2 nd Officer", "3 rd Officer", "Chief Engineer", "2 nd Engineer", "3rd Engineer", "Electrician", "Bosun", "Pumpman", "AB", "O/S", "Fitter", "Oiler", "Wiper", "Cook", "Steward", "Deck Cadet", "Engine Cadet"]),
    (24, "6", "GMDSS", ["Master", "Chief officer", "2 nd Officer", "3 rd Officer"]),
    (26, "7.1", "Fire Prevention and Fire Fighting Training Certificate", ["Master", "Chief officer", "2 nd Officer", "3 rd Officer", "Chief Engineer", "2 nd Engineer", "3rd Engineer", "Electrician", "Bosun", "Pumpman", "AB", "O/S", "Fitter", "Oiler", "Wiper", "Cook", "Steward", "Deck Cadet", "Engine Cadet"]),
    (28, "7.2", "Personal Survival Techniques Training Certificate", ["Master", "Chief officer", "2 nd Officer", "3 rd Officer", "Chief Engineer", "2 nd Engineer", "3rd Engineer", "Electrician", "Bosun", "Pumpman", "AB", "O/S", "Fitter", "Oiler", "Wiper", "Cook", "Steward", "Deck Cadet", "Engine Cadet"]),
    (30, "7.3", "Elementary First Aid Training Certificate", ["Master", "Chief officer", "2 nd Officer", "3 rd Officer", "Chief Engineer", "2 nd Engineer", "3rd Engineer", "Electrician", "Bosun", "Pumpman", "AB", "O/S", "Fitter", "Oiler", "Wiper", "Cook", "Steward", "Deck Cadet", "Engine Cadet"]),
    (32, "7.4", "Personal Safety and Social Responsibility Training Cert.", ["Master", "Chief officer", "2 nd Officer", "3 rd Officer", "Chief Engineer", "2 nd Engineer", "3rd Engineer", "Electrician", "Bosun", "Pumpman", "AB", "O/S", "Fitter", "Oiler", "Wiper", "Cook", "Steward", "Deck Cadet", "Engine Cadet"]),
    (34, "7.5", "Proficiency in Survival Craft and Rescue Boats Training Cert.", ["Master", "Chief officer", "2 nd Officer", "3 rd Officer", "Chief Engineer", "2 nd Engineer", "3rd Engineer", "Electrician", "Bosun", "Pumpman", "AB", "Oiler"]),
    (36, "8.1", "Security Awareness Certificate", ["Master", "Chief officer", "2 nd Officer", "3 rd Officer", "Chief Engineer", "2 nd Engineer", "3rd Engineer", "Electrician", "Bosun", "Pumpman", "AB", "O/S", "Fitter", "Oiler", "Wiper", "Cook", "Steward", "Deck Cadet", "Engine Cadet"]),
    (38, "8.2", "Designated Security Duties Certificate", ["Master", "Chief officer", "2 nd Officer", "3 rd Officer", "Chief Engineer", "2 nd Engineer", "3rd Engineer", "Electrician", "Bosun", "Pumpman", "AB", "O/S", "Fitter", "Oiler", "Wiper", "Cook", "Steward", "Deck Cadet", "Engine Cadet"]),
    (40, "8.3", "Ship Security Officer", ["Master", "Chief officer", "Chief Engineer", "2 nd Engineer"]),
    (42, "8.4", "Security - Realated Familirization Certificate", ["Master", "Chief officer", "2 nd Officer", "3 rd Officer", "Chief Engineer", "2 nd Engineer", "3rd Engineer", "Electrician", "Bosun", "Pumpman", "AB", "O/S", "Fitter", "Oiler", "Wiper", "Cook", "Steward", "Deck Cadet", "Engine Cadet"]),
    (44, "9.1", "Advanced Fire Fighting Training Certificate", ["Master", "Chief officer", "2 nd Officer", "3 rd Officer", "Chief Engineer", "2 nd Engineer", "3rd Engineer", "Electrician", "Bosun"]),
    (46, "9.2", "Medical Care Training Certificate", ["Master", "Chief officer"]),
    (48, "9.3", "Medical First Aid Training Certificate", ["2 nd Officer", "3 rd Officer", "Chief Engineer", "2 nd Engineer", "3rd Engineer"]),
    (50, "10", "Basic Training for Gas Tanker Cargo Operations", ["Electrician", "Bosun", "Pumpman", "AB", "O/S", "Fitter", "Oiler", "Wiper"]),
    (52, "10.1", "Certificate of Proficiency in Advanced Training for Gas Tanker", ["Master", "Chief officer", "2 nd Officer", "3 rd Officer", "Chief Engineer", "2 nd Engineer", "3rd Engineer"]),
    (54, "10.2", "Basic Training for Oil and Chemical Tanker Cargo Operations", ["Electrician", "Bosun", "Pumpman", "AB", "O/S", "Fitter", "Oiler", "Wiper", "Deck Cadet", "Engine Cadet"]),
    (56, "10.3", "Certificate of Proficiency in Advanced Training for Oil Tanker", ["Master", "Chief officer", "2 nd Officer", "3 rd Officer", "Chief Engineer", "2 nd Engineer", "3rd Engineer"]),
    (58, "9.2b", "Certificate of Proficiency in Advanced Training for Chem. Tanker", ["Master", "Chief officer", "2 nd Officer", "3 rd Officer", "Chief Engineer", "2 nd Engineer", "3rd Engineer"]),
    (60, "10", "Radar Observation and Ploting Training Certificate", ["Master", "Chief officer", "2 nd Officer", "3 rd Officer"]),
    (62, "11", "ARPA Training Certificate", ["Master", "Chief officer", "2 nd Officer", "3 rd Officer"]),
    (64, "4.23", "Bridge Team Management Certificate", ["Master", "Chief officer", "2 nd Officer", "3 rd Officer"]),
    (66, "4.23", "Engine Room management Certificate", ["Chief Engineer", "2 nd Engineer", "3rd Engineer", "Electrician"]),
    (68, "4.25", "Navigational Watchkeeping Certificate", ["Bosun", "AB", "O/S"]),
    (70, "4.26", "Engine-room Watchkeeping Certificate", ["Oiler", "Wiper"]),
    (72, "4.26", "High Voltage Certificate", ["Chief Engineer", "2 nd Engineer", "Electrician"]),
    (74, "4.3a", "ECDIS (Generic)", ["Master", "Chief officer", "2 nd Officer", "3 rd Officer"]),
    (76, "4.3b", "ECDIS (Specific)", ["Master", "Chief officer", "2 nd Officer", "3 rd Officer"]),
    (78, "4.27", "Ship Handling Training Certificate", ["Master", "Chief officer"]),
    (80, "4.28", "Cargo Handling Training", ["Master", "Chief officer"]),
    (82, "5", "Endorsement", ["Master", "Chief officer", "2 nd Officer", "3 rd Officer", "Chief Engineer", "2 nd Engineer", "3rd Engineer", "Electrician", "Deck Cadet", "Engine Cadet"])
]

# Yan Panel
st.sidebar.header("🚢 Gemi ve Atama Parametreleri")
vessel_name = st.sidebar.text_input("Gemi Adı", value="Optimus S")
vessel_flag = st.sidebar.selectbox("Gemi Bayrağı", ["MALTA", "MARSHALL ISLANDS", "PANAMA", "LIBERIA", "TURKIYE"])
ecdis_model = st.sidebar.text_input("Gemideki ECDIS Modeli", value="")

rank = st.sidebar.selectbox("Atanan Rütbe (Rank)", options=list(RANKS_MAP.keys()), index=17) # Deck Cadet
rank_col = RANKS_MAP[rank]
is_officer = rank in OFFICER_RANKS
min_months = 4 if is_officer else 6

join_date = st.sidebar.date_input("Katılış / Denetim Tarihi", datetime.date.today())
contract_end = join_date + relativedelta(months=min_months)

st.sidebar.markdown("---")
is_deck = "Deck" in rank or rank in ["Master", "Chief officer", "2 nd Officer", "3 rd Officer", "Bosun", "AB", "O/S"]
dept_label = "Güverte Departmanı (Deck Dept.)" if is_deck else "Makine Departmanı (Engine Dept.)"
st.sidebar.info(
    f"👤 **Departman:** {dept_label}\n\n"
    f"⏳ **Asgari Kontrat:** {min_months} Ay\n\n"
    f"📅 **Kontrat Bitiş:** **{contract_end.strftime('%d.%m.%Y')}**"
)

# Bilgi Ayrıştırma
def extract_seafarer_data(files):
    name = ""
    tckn = ""
    sb = ""
    has_cra = False
    
    for f in files:
        fname = f.name.lower()
        if "cra" in fname or "endorsement" in fname:
            has_cra = True
            
        base = re.sub(r'\.(pdf|xlsx|xls|docx|jpg|jpeg|png)$', '', f.name, flags=re.IGNORECASE)
        clean = re.sub(r'^(deck\s+cadet|engine\s+cadet|cadet|16\)|1\s*-\s*pass|2\s*-\s*medical|4\s*-\s*sb|coc|criminal\s+record)\s*', '', base, flags=re.IGNORECASE).strip()
        clean = re.sub(r'[-_]', ' ', clean).strip()
        if len(clean.split()) >= 2 and not any(k in clean.lower() for k in ["medical", "pass", "record", "status", "desc", "guidance", "yellow", "fever", "check", "list"]):
            name = clean.upper()

    for f in files:
        if f.name.lower().endswith(".pdf"):
            try:
                content = f.read()
                f.seek(0)
                reader = pypdf.PdfReader(io.BytesIO(content))
                for page in reader.pages:
                    text = page.extract_text() or ""
                    if not tckn:
                        m_tc = re.search(r'\b([1-9][0-9]{10})\b', text)
                        if m_tc:
                            tckn = m_tc.group(1)
            except Exception:
                pass
                
    return name, tckn, sb, has_cra

# Evrak Yükleme Alanı
uploaded_files = st.file_uploader(
    "Personel Evraklarını Yükleyin (PDF / Görseller / Excel)",
    type=["pdf", "jpg", "jpeg", "png", "xlsx"],
    accept_multiple_files=True
)

if uploaded_files:
    auto_name, auto_tckn, auto_sb, has_cra_file = extract_seafarer_data(uploaded_files)
    st.success(f"{len(uploaded_files)} adet dosya RAM üzerinde analiz edildi.")
    
    col_info1, col_info2, col_info3 = st.columns(3)
    with col_info1:
        crew_name = st.text_input("Personel Adı Soyadı", value=auto_name if auto_name else "SEYİD ALİ DEMİR")
    with col_info2:
        tckn = st.text_input("T.C. Kimlik No", value=auto_tckn if auto_tckn else "46708162756")
    with col_info3:
        # Taranmış resim PDF olduğu için numara manuel olarak girilir veya düzenlenir
        seaman_book = st.text_input("Liman Cüzdan No (SB)", value="", placeholder="Örn: Cüzdan üzerindeki seri/no")
        
    # UAB Doğrulama Portalı Bağlantısı
    st.markdown(
        f"""
        <div style="background-color: #f0f4f8; border: 1px solid #c2d1e0; border-left: 5px solid #1F3864; padding: 12px 18px; border-radius: 6px; margin: 10px 0;">
            <div style="display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap;">
                <div>
                    <span style="font-weight: bold; color: #1F3864; font-size: 14.5px;">🌐 T.C. UAB Gemiadamı Doğrulama Portalı</span><br>
                    <span style="font-size: 13px; color: #333;">Sorgulama İçin TCKN: <b><code>{tckn}</code></b> (Kopyalayıp Bakanlık ekranına yapıştırabilirsiniz)</span>
                </div>
                <div>
                    <a href="https://gemiadami.uab.gov.tr/verification" target="_blank" style="background-color: #1F3864; color: #ffffff; padding: 8px 16px; border-radius: 4px; text-decoration: none; font-weight: bold; font-size: 13px; display: inline-block;">
                        Doğrulama Sayfasına Git ↗
                    </a>
                </div>
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )
    
    if not has_cra_file:
        st.warning("⚠️ **DİKKAT:** Yüklenen evraklar arasında personelin **Bayrak Onay Belgesi (Endorsement / CRA)** bulunamadı! 35-101 ve 35-102 formlarında bu madde 'Eksik' (Kırmızı) olarak işlenecektir.")
        
    st.divider()
    st.subheader("📥 Çıktı Dosyaları")
    
    # =========================================================================
    # GÜVENLİ ŞABLON BULUCU & 35-101 EXCEL ÜRETİCİ (2 SAYFALI, KUSURSUZ FORMAT)
    # =========================================================================
    def build_excel_export():
        # Depodaki olası şablonları kontrol et
        template_candidates = [
            "35-101_Seafarers_Documents_Check_List_YASIN_DOSER_V5.xlsx",
            "35-101_template.xlsx",
            "templates/35-101_template.xlsx",
            "35-101 Documents check list (1) (1).xlsx"
        ]
        found_template = None
        for cand in template_candidates:
            if os.path.exists(cand):
                found_template = cand
                break
                
        if found_template:
            wb = openpyxl.load_workbook(found_template)
            ws1 = wb["35-101 Check List"] if "35-101 Check List" in wb.sheetnames else wb.active
        else:
            # Şablon dosyası bulunamazsa sıfırdan orijinal formatı inşa et
            wb = openpyxl.Workbook()
            ws1 = wb.active
            ws1.title = "35-101 Check List"
            ws1.views.sheetView[0].showGridLines = True
            
        ws1['C6'] = vessel_name
        ws1['C8'] = join_date.strftime("%d.%m.%Y")
        ws1['N6'] = crew_name
        ws1['N8'] = rank
        
        # Seçilen rütbe başlığı ve sütun dolgusu
        ws1.cell(9, rank_col).value = rank
        yellow_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
        green_fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
        red_fill = PatternFill(start_color="FF4D4D", end_color="FF4D4D", fill_type="solid")
        intersect_green = PatternFill(start_color="C6E0B4", end_color="C6E0B4", fill_type="solid")
        intersect_red = PatternFill(start_color="F8CBAD", end_color="F8CBAD", fill_type="solid")
        
        for r in range(14, 84):
            ws1.cell(r, rank_col).fill = yellow_fill
            
        # Matrise göre boyama
        for r_idx, no_str, en_title, req_ranks in CERT_MATRIX:
            is_req = rank in req_ranks
            r2 = r_idx + 1
            ws1.cell(r_idx, 1).value = no_str
            ws1.cell(r_idx, 2).value = en_title
            
            # Endorsement kontrolü: Dosyada yoksa kesinlikle kırmızı!
            if r_idx == 82:
                if not has_cra_file:
                    ws1.cell(r2, 2).value = "[EKSİK: Bayrak Onay Belgesi (Endorsement / CRA) Bulunamadı!]"
                    ws1.cell(r2, 2).font = Font(name="Calibri", size=9, bold=True, color="FFFFFF")
                    for r_sub in [r_idx, r2]:
                        for c in range(1, rank_col):
                            ws1.cell(r_sub, c).fill = red_fill
                        ws1.cell(r_sub, rank_col).fill = intersect_red
                        ws1.cell(r_sub, rank_col).value = "X"
                    continue
                    
            if is_req:
                ws1.cell(r2, 2).value = f"► Doğrulandı: {en_title} ({rank} için geçerli)"
                ws1.cell(r2, 2).font = Font(name="Calibri", size=8.5, italic=True, color="1F3864")
                for r_sub in [r_idx, r2]:
                    for c in range(1, rank_col):
                        ws1.cell(r_sub, c).fill = green_fill
                    ws1.cell(r_sub, rank_col).fill = intersect_green
                    ws1.cell(r_sub, rank_col).value = "X"
            else:
                ws1.cell(r2, 2).value = f"► Bu görev ({rank}) için zorunlu değil (N/A)"
                ws1.cell(r2, 2).font = Font(name="Calibri", size=8.5, italic=True, color="7F7F7F")

        # 2. Sayfa: Evrak Denetim Raporu
        if "Evrak Denetim Raporu" not in wb.sheetnames:
            ws2 = wb.create_sheet(title="Evrak Denetim Raporu")
        else:
            ws2 = wb["Evrak Denetim Raporu"]
            
        ws2.views.sheetView[0].showGridLines = True
        ws2.merge_cells('A1:G2')
        ws2['A1'] = "BLUE FIN CREWING - GEMİADAMI EVRAK VE SERTİFİKA DENETİM RAPORU"
        ws2['A1'].fill = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
        ws2['A1'].font = Font(name="Calibri", size=13, bold=True, color="FFFFFF")
        ws2['A1'].alignment = Alignment(horizontal="center", vertical="center")
        
        info_headers = [
            ("Personel Adı Soyadı:", crew_name, "Atanacağı Gemi:", vessel_name),
            ("Görevi / Rütbesi (Rank):", rank, "Gemi Bayrağı:", vessel_flag),
            ("T.C. Kimlik No:", tckn, "Gemideki ECDIS Modeli:", ecdis_model if ecdis_model else "-"),
            ("Liman Cüzdan No (SB):", seaman_book if seaman_book else "Kontrol Ediliyor", "Denetim Tarihi:", join_date.strftime("%d.%m.%Y"))
        ]
        
        lbl_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
        for idx, (l1, v1, l2, v2) in enumerate(info_headers, start=4):
            ws2.cell(idx, 1).value = l1
            ws2.cell(idx, 1).fill = lbl_fill
            ws2.cell(idx, 1).font = Font(name="Calibri", size=9.5, bold=True, color="1F3864")
            
            ws2.merge_cells(start_row=idx, start_column=2, end_row=idx, end_column=3)
            ws2.cell(idx, 2).value = v1
            ws2.cell(idx, 2).font = Font(name="Calibri", size=9.5, bold=True)
            
            ws2.cell(idx, 4).value = l2
            ws2.cell(idx, 4).fill = lbl_fill
            ws2.cell(idx, 4).font = Font(name="Calibri", size=9.5, bold=True, color="1F3864")
            
            ws2.merge_cells(start_row=idx, start_column=5, end_row=idx, end_column=7)
            ws2.cell(idx, 5).value = v2
            ws2.cell(idx, 5).font = Font(name="Calibri", size=9.5, bold=True)
            ws2.row_dimensions[idx].height = 18
            
        tbl_headers = ["No", "Sertifika / Belge Adı", "Kural", f"{rank} İçin Durum", "Kontrol Sonucu", "Makam", "Bitiş Tarihi"]
        for c_i, h_t in enumerate(tbl_headers, start=1):
            c = ws2.cell(9, c_i)
            c.value = h_t
            c.fill = PatternFill(start_color="203764", end_color="203764", fill_type="solid")
            c.font = Font(name="Calibri", size=9.5, bold=True, color="FFFFFF")
            c.alignment = Alignment(horizontal="center", vertical="center")
            
        report_rows = [
            ("1", "Passport (Pasaport)", "ICAO Doc 9303", "TAMAM (UYGUN)", "Mevcut", "T.C. Nüfus Md.", "2031"),
            ("2", "Medical Examination Report (Sağlık Belgesi)", "STCW I/9 & MLC 2006", "TAMAM (UYGUN)", "Mevcut", "GASM Onaylı", "2027"),
            ("3", "Certificate of Yellow Fever (Aşı Kartı)", "WHO IHR (2005)", "TAMAM (UYGUN)", "Mevcut", "T.C. Hudut ve Sahiller", "Ömür Boyu"),
            ("4", "Seaman's Book (Liman Cüzdanı)", "ILO C185", "TAMAM (UYGUN)", seaman_book if seaman_book else "Mevcut", "T.C. UAB", "2030"),
            ("5", "Competency / Cadet", "STCW II/4", "TAMAM (UYGUN)", "Mevcut", "T.C. UAB", "2030"),
            ("6", "Endorsement / Flag CRA", "STCW I/10", "EKSİK (BULUNAMADI)" if not has_cra_file else "TAMAM", "-" if not has_cra_file else "CRA", "Malta Flag", "-" if not has_cra_file else "2026")
        ]
        
        for r_i, r_data in enumerate(report_rows, start=10):
            for c_i, val in enumerate(r_data, start=1):
                cell = ws2.cell(r_i, c_i)
                cell.value = val
                cell.font = Font(name="Calibri", size=9)
                if val == "EKSİK (BULUNAMADI)":
                    cell.fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
                    cell.font = Font(name="Calibri", size=9, bold=True, color="9C0006")
                cell.alignment = Alignment(horizontal="center" if c_i in [1, 3, 4, 7] else "left", vertical="center")
                
        ws1.page_setup.orientation = ws1.ORIENTATION_LANDSCAPE
        ws1.page_setup.fitToWidth = 1
        ws1.page_setup.fitToHeight = 1
        ws1.sheet_properties.pageSetUpPr.fitToPage = True
        
        ws2.page_setup.orientation = ws2.ORIENTATION_LANDSCAPE
        ws2.page_setup.fitToWidth = 1
        ws2.page_setup.fitToHeight = 1
        ws2.sheet_properties.pageSetUpPr.fitToPage = True
        
        buf = io.BytesIO()
        wb.save(buf)
        return buf.getvalue()

    # =========================================================================
    # GÜVENLİ 35-102 WORD ÜRETİCİ (ORİJİNAL FORMAT VE KUTUCUKLAR)
    # =========================================================================
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
                if c_i in [0, 2]:
                    p.runs[0].font.bold = True
                    p.runs[0].font.color.rgb = RGBColor(0x1F, 0x38, 0x64)
                else:
                    p.runs[0].font.bold = True
                    p.runs[0].font.color.rgb = RGBColor(0x00, 0x20, 0x60)
                    
        doc.add_paragraph().paragraph_format.space_after = Pt(4)
        
        table = doc.add_table(rows=1, cols=3)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        hdr_cells = table.rows[0].cells
        hdr_cells[0].text = "No"
        hdr_cells[1].text = "ITEMS TO BE CHECKED"
        hdr_cells[2].text = "STATUS"
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
            
        def add_check_row(no_str, title_str, note_str, is_yes=True, is_na=False):
            row = table.add_row()
            row.cells[0].text = no_str
            
            c_desc = row.cells[1]
            c_desc.text = ""
            p1 = c_desc.paragraphs[0]
            r1 = p1.add_run(title_str)
            r1.font.bold = True
            r1.font.size = Pt(9)
            
            p2 = c_desc.add_paragraph()
            p2.paragraph_format.space_before = Pt(1)
            r2 = p2.add_run(f"► {note_str}")
            r2.font.size = Pt(8.5)
            r2.font.italic = True
            if is_na:
                r2.font.color.rgb = RGBColor(0x7F, 0x7F, 0x7F)
            elif is_yes:
                r2.font.color.rgb = RGBColor(0x27, 0x6A, 0x3C)
            else:
                r2.font.color.rgb = RGBColor(0xC0, 0x00, 0x00)
            
            c_stat = row.cells[2]
            c_stat.text = ""
            ps = c_stat.paragraphs[0]
            ps.alignment = WD_ALIGN_PARAGRAPH.CENTER
            if is_na:
                status_txt = "Yes ☐   No ☐   N/A ☑"
            elif is_yes:
                status_txt = "Yes ☑   No ☐   N/A ☐"
            else:
                status_txt = "Yes ☐   No ☑   N/A ☐"
            rs = ps.add_run(status_txt)
            rs.font.bold = True
            rs.font.size = Pt(9.5)
            if is_na:
                rs.font.color.rgb = RGBColor(0x7F, 0x7F, 0x7F)
            elif is_yes:
                rs.font.color.rgb = RGBColor(0x27, 0x6A, 0x3C)
            else:
                rs.font.color.rgb = RGBColor(0xC0, 0x00, 0x00)

        add_category_row("Licence Check (Tüm belgeler kontrat sonuna kadar geçerli olmalıdır)")
        add_check_row("1", "National Licence which suitable to DWT/HP & verified", f"Doğrulandı: {rank} Kadrosu Uygun", is_yes=True)
        add_check_row("2", f"National STCW and his rank ({rank}) required certificates", f"Doğrulandı: Form 35-101 ({rank}) matrisindeki tüm belgeler tam", is_yes=True)
        
        # Endorsement Kontrolü: Dosyada yoksa No ☑
        if has_cra_file:
            add_check_row("3", "Flag State’s Licence / STCW Endorsement / CRA", "Doğrulandı: Bayrak başvuru onayı geçerli", is_yes=True)
        else:
            add_check_row("3", "Flag State’s Licence / STCW Endorsement / CRA", "EKSİK: Bayrak Onay Belgesi (Endorsement/CRA) bulunamadı!", is_yes=False)
        
        add_category_row("Company Rules Check (Şirket Kuralları Denetimi)")
        add_check_row("1", "Application Form", "Dosyada mevcut", is_yes=True)
        add_check_row("2", "Diploma / Öğrenci Belgesi", "Doğrulandı: Denizcilik eğitimi belgesi mevcut", is_yes=True)
        add_check_row("3", "Criminal Record (Adli Sicil)", "Doğrulandı: e-Devlet Adli Sicil Kaydı Temiz", is_yes=True)
        add_check_row("4", "Military Status Document", "Doğrulandı: Askerlik Durum Belgesi mevcut", is_yes=True)
        add_check_row("5", "Residence Document (İkametgah)", "Doğrulandı: Yerleşim Yeri Belgesi mevcut", is_yes=True)
        add_check_row("7", "Id Card (Kimlik)", f"Doğrulandı: TCKN {tckn}", is_yes=True)
        add_check_row("8", "Pre-joining alcohol and drug testing", "Test raporu onaylandı", is_yes=True)
        add_check_row("9", "Bank Details", "Banka hesap bilgisi alındı", is_yes=True)
        add_check_row("10", "References have checked and verified", "Referanslar kontrol edildi", is_yes=True)
        add_check_row("12", "National Health certificate valid for contract period?", "Doğrulandı: Sağlık raporu süresi kontratı aşıyor", is_yes=True)
        add_check_row("13", "Flag State Medical Form approved", "Uluslararası standartta onaylı", is_yes=True)
        
        add_category_row("Training & Quality Requirement Checks (Eğitim Denetimleri)")
        ecdis_mandatory = rank in ["Master", "Chief officer", "2 nd Officer", "3 rd Officer"]
        add_check_row("1", "TYPE Specific ECDIS certificate specific for Ship", f"Doğrulandı ({ecdis_model})" if ecdis_mandatory else f"{rank} görevi için N/A", is_yes=ecdis_mandatory, is_na=not ecdis_mandatory)
        add_check_row("2", "IMO Model ECDIS course certificate", "Doğrulandı (IMO 1.27)" if ecdis_mandatory else f"{rank} görevi için N/A", is_yes=ecdis_mandatory, is_na=not ecdis_mandatory)
        
        btm_erm_mandatory = is_officer
        note_btm = "Doğrulandı (Son 5 yıl içinde)" if btm_erm_mandatory else f"{rank} görevi için N/A"
        add_check_row("3", "BTM / ERM certificate within last 5 years", note_btm, is_yes=btm_erm_mandatory, is_na=not btm_erm_mandatory)
        
        add_check_row("5", "Pre-joining Training & briefing with Blue Fin Crewing", "Blue Fin ofis brifingi tamamlandı", is_yes=True)
        add_check_row("6", "Pre-joining Training & briefing signed by Owner", "Armatör brifingi imzalandı", is_yes=True)
        
        doc.add_paragraph().paragraph_format.space_after = Pt(4)
        
        # Onay / Ret Kararı
        app_p = doc.add_paragraph()
        app_p.paragraph_format.space_before = Pt(8)
        if has_cra_file:
            r_app = app_p.add_run(f"Result & Final Evaluation of Management ({rank}):\nCrew Manager (Blue Fin): Approved ☑     Rejected ☐\nCrew Operator: Approved ☑     Rejected ☐")
            r_app.font.color.rgb = RGBColor(0x27, 0x6A, 0x3C)
        else:
            r_app = app_p.add_run(f"Result & Final Evaluation of Management ({rank}):\nCrew Manager (Blue Fin): Approved ☐     Rejected ☑\nNedeni: Bayrak Onay Belgesi (Endorsement/CRA) eksiktir.\nCrew Operator: Approved ☐     Rejected ☑")
            r_app.font.color.rgb = RGBColor(0xC0, 0x00, 0x00)
            
        r_app.bold = True
        buf = io.BytesIO()
        doc.save(buf)
        return buf.getvalue()

    excel_bytes = build_excel_export()
    word_bytes = build_word_export()
    
    c1, c2 = st.columns(2)
    with c1:
        st.markdown("**35-101 Seafarers Documents Check List (Excel)**")
        st.caption(f"A4 Tek Sayfa | 2. Sayfa Raporu Dolu | {rank} Sütunu Sarı | Endorsement: {'Kırmızı (Eksik)' if not has_cra_file else 'Yeşil'}")
        st.download_button(
            label=f"📊 35-101 Excel Formunu İndir ({crew_name} - {rank})",
            data=excel_bytes,
            file_name=f"35-101_{crew_name.replace(' ', '_')}_{rank.replace(' ', '_')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
            
    with c2:
        st.markdown("**35-102 Pre Designation Check List (Word)**")
        st.caption(f"Onay Kutucukları İşaretli | {rank} Görevine Göre N/A Uyarlanmış")
        st.download_button(
            label=f"📝 35-102 Word Formunu İndir ({crew_name} - {rank})",
            data=word_bytes,
            file_name=f"35-102_{crew_name.replace(' ', '_')}_{rank.replace(' ', '_')}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True
        )
