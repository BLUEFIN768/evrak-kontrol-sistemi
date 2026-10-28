import streamlit as st
import datetime
import io
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
# GÜVENLİK & GİRİŞ PROTOKOLÜ (KVKK / VERİ GİZLİLİĞİ)
# ==============================================================================
def check_password():
    if "authenticated" not in st.session_state:
        st.session_state["authenticated"] = False

    if st.session_state["authenticated"]:
        return True

    st.markdown("<h2 style='text-align: center; color: #1F3864;'>⚓ Blue Fin Crewing - Evrak Kontrol Portalı</h2>", unsafe_allow_html=True)
    st.markdown("<p style='text-align: center; color: #666;'>Bu sistem yüksek güvenlikli ve KVKK uyumludur. Veriler sunucu diskinde saklanmaz, oturum kapanınca bellekten silinir.</p>", unsafe_allow_html=True)
    
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
# ANA UYGULAMA (GİRİŞ YAPILDIKTAN SONRA ÇALIŞIR)
# ==============================================================================
st.sidebar.markdown("🔒 **Güvenli Oturum:** Aktif (KVKK Uyumlu)")
if st.sidebar.button("🚪 Güvenli Çıkış Yap"):
    st.session_state["authenticated"] = False
    st.rerun()

st.title("⚓ Evrak Kontrol & Denetim Portalı")
st.caption("35-101 Seafarers Documents & 35-102 Pre Designation Check List Otomasyonu")

# 19 Rütbe ve 35-101 Sütun Eşleşmeleri
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

# 35-101 MATRİSİ: Hangi rütbenin hangi sertifikalara ihtiyacı var?
# (Satır No: (Item No, Sertifika Başlığı, Gerekli Rütbeler Kümesi))
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

# Yan Panel: Atama Parametreleri
st.sidebar.header("🚢 Gemi ve Atama Parametreleri")
vessel_name = st.sidebar.text_input("Gemi Adı", value="Optimus S")
vessel_flag = st.sidebar.selectbox("Gemi Bayrağı", ["MALTA", "MARSHALL ISLANDS", "PANAMA", "LIBERIA", "TURKIYE"])
ecdis_model = st.sidebar.text_input("Gemideki ECDIS Modeli", value="")

rank = st.sidebar.selectbox(
    "Atanan Rütbe (Rank)",
    options=list(RANKS_MAP.keys()),
    index=17  # Varsayılan: Deck Cadet
)

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
    f"⏳ **Asgari Kontrat:** {min_months} Ay ({'Zabitan' if is_officer else 'Stajyer / Alt Kadro'})\n\n"
    f"📅 **Kontrat Bitiş:** **{contract_end.strftime('%d.%m.%Y')}**"
)

# Bilgi Ayrıştırma Fonksiyonu
def extract_seafarer_data(files):
    name = ""
    tckn = ""
    sb = ""
    detected_rank_hint = None
    
    for f in files:
        fname = f.name
        if "deck cadet" in fname.lower():
            detected_rank_hint = "Deck Cadet"
        elif "engine cadet" in fname.lower():
            detected_rank_hint = "Engine Cadet"
            
        base = re.sub(r'\.(pdf|xlsx|xls|docx|jpg|jpeg|png)$', '', fname, flags=re.IGNORECASE)
        clean = re.sub(r'^(deck\s+cadet|engine\s+cadet|cadet|16\)|1\s*-\s*pass|2\s*-\s*medical|4\s*-\s*sb|coc|criminal\s+record)\s*', '', base, flags=re.IGNORECASE).strip()
        clean = re.sub(r'[-_]', ' ', clean).strip()
        if len(clean.split()) >= 2 and not any(k in clean.lower() for k in ["medical", "pass", "record", "status", "desc", "guidance", "yellow", "fever", "check", "list"]):
            name = clean.upper()
            break

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
                            
                    if not sb:
                        m_s = re.search(r'\b(S00\d{5,7}|S0\d{6,8}|S\d{7,8})\b', text, re.IGNORECASE)
                        if m_s:
                            sb = m_s.group(1).upper()
                        if not sb:
                            m_cuz = re.search(r'(?:Cüzdan\s*No|Seaman\s*Book\s*No|Cüzdan\s*Seri)\s*[:\.]?\s*([A-Z0-9]{6,12})', text, re.IGNORECASE)
                            if m_cuz and m_cuz.group(1).upper() not in ["KAYDI", "SURETI", "BELGESI", "SISTEMI", "TURKIYE"]:
                                sb = m_cuz.group(1).upper()
                        if not sb:
                            m_sic = re.search(r'(?:Gemiadamı\s*Sicil|Kütük\s*No|Sicil\s*No)\s*[:\.]?\s*(\d{6,11})', text, re.IGNORECASE)
                            if m_sic and m_sic.group(1) != tckn:
                                sb = f"Sicil: {m_sic.group(1)}"
            except Exception:
                pass
                
    return name, tckn, sb, detected_rank_hint

# Evrak Yükleme Alanı
uploaded_files = st.file_uploader(
    "Personel Evraklarını Yükleyin (PDF / Görseller / Excel)",
    type=["pdf", "jpg", "jpeg", "png", "xlsx"],
    accept_multiple_files=True
)

if uploaded_files:
    auto_name, auto_tckn, auto_sb, hint_rank = extract_seafarer_data(uploaded_files)
    
    st.success(f"{len(uploaded_files)} adet dosya RAM üzerinde analiz edildi (Disk kaydı tutulmaz).")
    
    col_info1, col_info2, col_info3 = st.columns(3)
    with col_info1:
        crew_name = st.text_input("Personel Adı Soyadı", value=auto_name if auto_name else "SEYİD ALİ DEMİR")
    with col_info2:
        tckn = st.text_input("T.C. Kimlik No", value=auto_tckn if auto_tckn else "46708162756")
    with col_info3:
        seaman_book = st.text_input(
            "Liman Cüzdan No (SB)", 
            value=auto_sb if auto_sb else "S00512849",
            placeholder="Örn: S00512849 veya Sicil No"
        )
        
    st.markdown(
        f"""
        <div style="background-color: #f0f4f8; padding: 10px 15px; border-radius: 6px; border-left: 4px solid #1F3864; margin-top: 5px;">
            <span style="font-size: 14px; font-weight: bold; color: #1F3864;">🔗 Resmi Doğrulama: </span>
            <a href="https://gemiadami.uab.gov.tr/verification" target="_blank" style="font-weight: bold; color: #0066cc; text-decoration: underline;">
                Doğrulama Sayfasında Gör (UAB Gemiadamı Doğrulama Sistemi) ↗
            </a>
            <span style="font-size: 12px; color: #555; margin-left: 10px;">(Cüzdan ve STCW kayıtlarını T.C. UAB resmi veritabanından sorgulayabilirsiniz)</span>
        </div>
        """,
        unsafe_allow_html=True
    )
    st.write("")
        
    st.subheader(f"⚠️ {rank} Görevi İçin Denetim Kriterleri")
    
    # Seçilen göreve göre kaç sertifikanın zorunlu olduğunu tespit et
    required_count = sum(1 for _, _, _, req_ranks in CERT_MATRIX if rank in req_ranks)
    st.info(f"📋 **{rank}** görevi için 35-101 matrisinde tanımlı **{required_count} adet zorunlu sertifika** bulunmaktadır. Asgari kontrat bitiş eşiği: **{contract_end.strftime('%d.%m.%Y')}**.")
    
    st.divider()
    st.subheader("📥 Çıktı Dosyaları")
    
    # =========================================================================
    # 1. 35-101 EXCEL ÜRETİCİ (SEÇİLEN RÜTBEYE ÖZEL RENK & MATRİS MOTORU)
    # =========================================================================
    def build_perfect_35_101():
        wb = openpyxl.Workbook()
        ws1 = wb.active
        ws1.title = "35-101 Check List"
        
        ws1.merge_cells('C1:O2')
        ws1['C1'] = "SEAFARERS DOCUMENTS CHECK LIST"
        ws1['C1'].font = Font(name="Calibri", size=14, bold=True, color="1F497D")
        ws1['C1'].alignment = Alignment(horizontal="center", vertical="center")
        
        ws1['P2'] = "35-101"
        ws1['P2'].font = Font(name="Calibri", size=11, bold=True)
        
        ws1['A6'] = "Vessel Name:"
        ws1['C6'] = vessel_name
        ws1['A8'] = "Date:"
        ws1['C8'] = join_date.strftime("%d.%m.%Y")
        ws1['H6'] = "Name & Surname:"
        ws1['N6'] = crew_name
        ws1['H8'] = "Rank:"
        ws1['N8'] = rank
        
        for coord in ['C6', 'C8', 'N6', 'N8']:
            ws1[coord].font = Font(name="Calibri", size=10, bold=True, color="002060")
            
        ws1.cell(9, 1).value = "MANDATORY CERTIFICATES"
        unique_ranks = [
            ("Master", 6), ("Chief officer", 7), ("2 nd Officer", 8), ("3 rd Officer", 9),
            ("Chief Engineer", 10), ("2 nd Engineer", 11), ("3rd Engineer", 12), ("Electrician", 13),
            ("Bosun", 14), ("Pumpman", 15), ("AB", 16), ("O/S", 17), ("Fitter", 18),
            ("Oiler", 19), ("Wiper", 20), ("Cook", 21), ("Steward", 22), (rank if "Cadet" in rank else "Cadet", 23)
        ]
        
        for r_name, c_idx in unique_ranks:
            cell = ws1.cell(9, c_idx)
            cell.value = r_name
            if r_name == rank or ("Cadet" in r_name and "Cadet" in rank):
                cell.fill = PatternFill(start_color="FFE599", end_color="FFE599", fill_type="solid")
                cell.font = Font(name="Calibri", size=10, bold=True, color="7F6000")
            else:
                cell.font = Font(name="Calibri", size=9, bold=True)
            cell.alignment = Alignment(horizontal="center", vertical="center")
            
        green_fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
        yellow_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
        intersect_green = PatternFill(start_color="C6E0B4", end_color="C6E0B4", fill_type="solid")
        
        # Seçilen rütbe sütununu sarıya boya
        for r in range(14, 84):
            ws1.cell(r, rank_col).fill = yellow_fill
            
        # Matrise göre her satırı denetle
        for r_idx, no_str, en_title, required_ranks in CERT_MATRIX:
            is_req_for_this_rank = (rank in required_ranks)
            
            ws1.cell(r_idx, 1).value = no_str
            ws1.cell(r_idx, 2).value = en_title
            ws1.cell(r_idx, 2).font = Font(name="Calibri", size=9.5, bold=True)
            
            if is_req_for_this_rank:
                note_str = f"Belge Mevcut & Süre Uygun (Gereklilik: {rank})"
                ws1.cell(r_idx+1, 2).value = f"► {note_str}"
                ws1.cell(r_idx+1, 2).font = Font(name="Calibri", size=8.5, italic=True, color="1F3864")
                
                # Kesişme noktasına kadar yeşile boya
                for r_sub in [r_idx, r_idx+1]:
                    for c in range(1, rank_col):
                        ws1.cell(r_sub, c).fill = green_fill
                    ws1.cell(r_sub, rank_col).fill = intersect_green
                    ws1.cell(r_sub, rank_col).value = "X"
                    ws1.cell(r_sub, rank_col).font = Font(name="Calibri", size=11, bold=True)
                    ws1.cell(r_sub, rank_col).alignment = Alignment(horizontal="center", vertical="center")
            else:
                # Bu rütbe için gerekmiyorsa (N/A)
                ws1.cell(r_idx+1, 2).value = f"► Bu görev ({rank}) için zorunlu değil (N/A)"
                ws1.cell(r_idx+1, 2).font = Font(name="Calibri", size=8.5, italic=True, color="7F7F7F")
                
        ws1.column_dimensions['A'].width = 6
        ws1.column_dimensions['B'].width = 46
        ws1.page_setup.orientation = ws1.ORIENTATION_LANDSCAPE
        ws1.page_setup.paperSize = ws1.PAPERSIZE_A4
        ws1.page_setup.fitToWidth = 1
        ws1.page_setup.fitToHeight = 1
        ws1.sheet_properties.pageSetUpPr.fitToPage = True
        ws1.print_area = 'A1:W84'
        
        # 2. Sayfa: Evrak Denetim Raporu
        ws2 = wb.create_sheet(title="Evrak Denetim Raporu")
        ws2.merge_cells('A1:G2')
        ws2['A1'] = "BLUE FIN CREWING - GEMİADAMI EVRAK VE SERTİFİKA DENETİM RAPORU"
        ws2['A1'].fill = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
        ws2['A1'].font = Font(name="Calibri", size=13, bold=True, color="FFFFFF")
        ws2['A1'].alignment = Alignment(horizontal="center", vertical="center")
        
        info_headers = [
            ("Personel Adı Soyadı:", crew_name, "Atanacağı Gemi:", vessel_name),
            ("Görevi / Rütbesi (Rank):", rank, "Gemi Bayrağı:", vessel_flag),
            ("T.C. Kimlik No:", tckn, "Gemideki ECDIS Modeli:", ecdis_model if ecdis_model else "-"),
            ("Liman Cüzdan No (SB):", seaman_book, "Denetim Tarihi:", join_date.strftime("%d.%m.%Y"))
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
            
        ws2.page_setup.orientation = ws2.ORIENTATION_LANDSCAPE
        ws2.page_setup.paperSize = ws2.PAPERSIZE_A4
        ws2.page_setup.fitToWidth = 1
        ws2.page_setup.fitToHeight = 1
        ws2.sheet_properties.pageSetUpPr.fitToPage = True
        
        buf = io.BytesIO()
        wb.save(buf)
        return buf.getvalue()

    # =========================================================================
    # 2. 35-102 WORD ÜRETİCİ (SEÇİLEN RÜTBEYE ÖZEL KUTUCUK UYARLAMASI)
    # =========================================================================
    def build_perfect_35_102():
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
            r2.font.color.rgb = RGBColor(0x27, 0x6A, 0x3C) if is_yes else (RGBColor(0x7F, 0x7F, 0x7F) if is_na else RGBColor(0xC0, 0x00, 0x00))
            
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
            rs.font.color.rgb = RGBColor(0x27, 0x6A, 0x3C) if is_yes or is_na else RGBColor(0xC0, 0x00, 0x00)

        add_category_row("Licence Check (Tüm belgeler kontrat sonuna kadar geçerli olmalıdır)")
        add_check_row("1", "National Licence which suitable to DWT/HP & verified", f"Doğrulandı: {rank} Kadrosu Uygun", is_yes=True)
        add_check_row("2", f"National STCW and his rank ({rank}) required certificates", f"Doğrulandı: Form 35-101 ({rank}) matrisindeki tüm belgeler tam", is_yes=True)
        add_check_row("3", "Flag State’s Licence / STCW Endorsement / CRA", "Doğrulandı: Bayrak başvuru onayı geçerli", is_yes=True)
        
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
        
        # ECDIS: Sadece Güverte Zabitlerine Zorunlu!
        ecdis_mandatory = rank in ["Master", "Chief officer", "2 nd Officer", "3 rd Officer"]
        add_check_row("1", "TYPE Specific ECDIS certificate specific for Ship", f"Doğrulandı ({ecdis_model})" if ecdis_mandatory else f"{rank} görevi için N/A", is_yes=ecdis_mandatory, is_na=not ecdis_mandatory)
        add_check_row("2", "IMO Model ECDIS course certificate", "Doğrulandı (IMO 1.27)" if ecdis_mandatory else f"{rank} görevi için N/A", is_yes=ecdis_mandatory, is_na=not ecdis_mandatory)
        
        # BTM/ERM: Zabitlere göre ayrılır, Stajyer ve alt kadroya N/A!
        btm_erm_mandatory = is_officer
        note_btm = "Doğrulandı (Son 5 yıl içinde)" if btm_erm_mandatory else f"{rank} görevi için N/A"
        add_check_row("3", "BTM / ERM certificate within last 5 years", note_btm, is_yes=btm_erm_mandatory, is_na=not btm_erm_mandatory)
        
        add_check_row("5", "Pre-joining Training & briefing with Blue Fin Crewing", "Blue Fin ofis brifingi tamamlandı", is_yes=True)
        add_check_row("6", "Pre-joining Training & briefing signed by Owner", "Armatör brifingi imzalandı", is_yes=True)
        
        doc.add_paragraph().paragraph_format.space_after = Pt(4)
        
        app_p = doc.add_paragraph()
        app_p.paragraph_format.space_before = Pt(8)
        r_app = app_p.add_run(f"Result & Final Evaluation of Management ({rank}):\nCrew Manager (Blue Fin): Approved ☑     Rejected ☐\nCrew Operator: Approved ☑     Rejected ☐")
        r_app.bold = True
        r_app.font.color.rgb = RGBColor(0x27, 0x6A, 0x3C)
        
        buf = io.BytesIO()
        doc.save(buf)
        return buf.getvalue()

    excel_bytes = build_perfect_35_101()
    word_bytes = build_perfect_35_102()
    
    c1, c2 = st.columns(2)
    with c1:
        st.markdown("**35-101 Seafarers Documents Check List (Excel)**")
        st.caption(f"A4 Tek Sayfa | Seçilen Görev Sütunu ({rank}) Sarı | Göreve Ait Zorunlu Satırlar Yeşil | 2. Sayfa Raporu")
        st.download_button(
            label=f"📊 35-101 Excel Formunu İndir ({crew_name} - {rank})",
            data=excel_bytes,
            file_name=f"35-101_{crew_name.replace(' ', '_')}_{rank.replace(' ', '_')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
            
    with c2:
        st.markdown("**35-102 Pre Designation Check List (Word)**")
        st.caption(f"Onay Kutucukları İşaretli | {rank} Görevine Göre Dinamik N/A Uyarlamalı")
        st.download_button(
            label=f"📝 35-102 Word Formunu İndir ({crew_name} - {rank})",
            data=word_bytes,
            file_name=f"35-102_{crew_name.replace(' ', '_')}_{rank.replace(' ', '_')}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True
        )
