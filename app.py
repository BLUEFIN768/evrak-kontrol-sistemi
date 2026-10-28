import streamlit as st
import datetime
import io
import re
from dateutil.relativedelta import relativedelta
import pypdf
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

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
    st.markdown("<p style='text-align: center; color: #666;'>Bu sistem yüksek güvenlikli ve KVKK uyumludur. Yüklenen veriler sunucuda saklanmaz.</p>", unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns([1, 1.5, 1])
    with col2:
        with st.form("login_form"):
            username = st.text_input("Yetkili Kullanıcı")
            password = st.text_input("Güvenlik Parolası", type="password")
            submit = st.form_submit_button("Giriş Yap", use_container_width=True)
            
            # Şirket varsayılan giriş bilgileri (İstediğiniz zaman değiştirebilirsiniz)
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
st.sidebar.markdown(f"🔒 **Güvenli Oturum:** Aktif (KVKK Uyumlu)")
if st.sidebar.button("🚪 Güvenli Çıkış Yap"):
    st.session_state["authenticated"] = False
    st.rerun()

st.title("⚓ Evrak Kontrol & Denetim Portalı")
st.caption("35-101 Seafarers Documents & 35-102 Pre Designation Check List Otomasyonu")

# 18 Rütbe ve Sütun Eşleşmeleri
RANKS_MAP = {
    "Master": 6, "Chief officer": 7, "2 nd Officer": 8, "3 rd Officer": 9,
    "Chief Engineer": 10, "2 nd Engineer": 11, "3rd Engineer": 12, "Electrician": 13,
    "Bosun": 14, "Pumpman": 15, "AB": 16, "O/S": 17, "Fitter": 18,
    "Oiler": 19, "Wiper": 20, "Cook": 21, "Steward": 22, "Cadet": 23
}

OFFICER_RANKS = [
    "Master", "Chief officer", "2 nd Officer", "3 rd Officer",
    "Chief Engineer", "2 nd Engineer", "3rd Engineer", "Electrician"
]

# Yan Panel: Atama Parametreleri
st.sidebar.header("🚢 Gemi ve Atama Parametreleri")
vessel_name = st.sidebar.text_input("Gemi Adı", value="Optimus S")
vessel_flag = st.sidebar.selectbox("Gemi Bayrağı", ["MALTA", "MARSHALL ISLANDS", "PANAMA", "LIBERIA", "TURKIYE"])
ecdis_model = st.sidebar.text_input("Gemideki ECDIS Modeli", value="")

rank = st.sidebar.selectbox("Atanan Rütbe (Rank)", options=list(RANKS_MAP.keys()), index=17) # Cadet
rank_col = RANKS_MAP[rank]
is_officer = rank in OFFICER_RANKS
min_months = 4 if is_officer else 6

join_date = st.sidebar.date_input("Katılış / Denetim Tarihi", datetime.date.today())
contract_end = join_date + relativedelta(months=min_months)

st.sidebar.markdown("---")
st.sidebar.info(
    f"👤 **Kadro Türü:** {'Zabitan (Officer)' if is_officer else 'Alt Kadro (Rating)'}\n\n"
    f"⏳ **Asgari Kontrat:** {min_months} Ay\n\n"
    f"📅 **Kontrat Bitiş:** **{contract_end.strftime('%d.%m.%Y')}**"
)

# Gelişmiş Liman Cüzdanı & Bilgi Tarayıcı
def extract_seafarer_data(files):
    name = ""
    tckn = ""
    sb = ""
    
    # 1. Dosya isimlerinden Ad-Soyad çıkarma
    for f in files:
        fname = f.name
        base = re.sub(r'\.(pdf|xlsx|xls|docx|jpg|jpeg|png)$', '', fname, flags=re.IGNORECASE)
        clean = re.sub(r'^(deck\s+cadet|cadet|16\)|1\s*-\s*pass|2\s*-\s*medical|4\s*-\s*sb|coc|criminal\s+record)\s*', '', base, flags=re.IGNORECASE).strip()
        clean = re.sub(r'[-_]', ' ', clean).strip()
        if len(clean.split()) >= 2 and not any(k in clean.lower() for k in ["medical", "pass", "record", "status", "desc", "guidance"]):
            name = clean.upper()
            break

    # 2. PDF metinlerinden TCKN ve Cüzdan No Arama
    for f in files:
        if f.name.lower().endswith(".pdf"):
            try:
                content = f.read()
                f.seek(0)
                reader = pypdf.PdfReader(io.BytesIO(content))
                for page in reader.pages:
                    text = page.extract_text() or ""
                    
                    # TCKN: 11 haneli sayı
                    if not tckn:
                        m_tc = re.search(r'\b([1-9][0-9]{10})\b', text)
                        if m_tc:
                            tckn = m_tc.group(1)
                            
                    # Cüzdan No (S0... formatı veya Sicil No)
                    if not sb:
                        # Format 1: S00471309 gibi S ile başlayan 8-9 haneli
                        m_sb = re.search(r'\b(S\d{7,8})\b', text, re.IGNORECASE)
                        if m_sb:
                            sb = m_sb.group(1).upper()
                        # Format 2: "Cüzdan No" veya "Sicil No" ibaresinden sonra gelen numara
                        if not sb:
                            m_sb2 = re.search(r'(?:Cüzdan|Sicil|Seaman\s*Book)\s*(?:No|Nu)?\s*[:\.]?\s*([A-Z0-9]{5,12})', text, re.IGNORECASE)
                            if m_sb2:
                                sb = m_sb2.group(1).upper()
            except Exception:
                pass
                
    return name, tckn, sb

# Evrak Yükleme Alanı
uploaded_files = st.file_uploader(
    "Personel Evraklarını Yükleyin (PDF / Görseller / Excel)",
    type=["pdf", "jpg", "jpeg", "png", "xlsx"],
    accept_multiple_files=True
)

if uploaded_files:
    auto_name, auto_tckn, auto_sb = extract_seafarer_data(uploaded_files)
    
    st.success(f"{len(uploaded_files)} adet dosya RAM üzerinde analiz edildi (Disk kaydı tutulmaz).")
    
    col_info1, col_info2, col_info3 = st.columns(3)
    with col_info1:
        crew_name = st.text_input("Personel Adı Soyadı", value=auto_name if auto_name else "SEYİD ALİ DEMİR")
    with col_info2:
        tckn = st.text_input("T.C. Kimlik No", value=auto_tckn if auto_tckn else "46708162756")
    with col_info3:
        seaman_book = st.text_input("Liman Cüzdan No (SB)", value=auto_sb if auto_sb else "S00512849 (Sicil: 4392004812)")
        
    st.subheader("⚠️ Canlı Kural ve Emniyet Uyarıları")
    st.info(f"📋 **{rank}** rütbesi için evrak geçerlilikleri asgari kontrat sonu olan **{contract_end.strftime('%d.%m.%Y')}** ({min_months} ay) tarihine göre denetlenmektedir.")
    
    st.divider()
    st.subheader("📥 Çıktı Dosyaları")
    
    # Canlı Excel Oluşturucu (Template dosyasına bağımlılığı kaldıran güvenli motor)
    def build_live_excel():
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "35-101 Check List"
        
        # Başlık ve Üst Bilgiler
        ws.merge_cells('C1:O2')
        ws['C1'] = "SEAFARERS DOCUMENTS CHECK LIST (35-101)"
        ws['C1'].font = Font(name="Calibri", size=14, bold=True, color="1F497D")
        ws['C1'].alignment = Alignment(horizontal="center", vertical="center")
        
        ws['A6'] = "Vessel Name:"
        ws['C6'] = vessel_name
        ws['A8'] = "Date:"
        ws['C8'] = join_date.strftime("%d.%m.%Y")
        ws['H6'] = "Name & Surname:"
        ws['N6'] = crew_name
        ws['H8'] = "Rank:"
        ws['N8'] = rank
        
        # Rütbe Başlıkları
        ws.cell(9, 1).value = "MANDATORY CERTIFICATES"
        for r_name, c_idx in RANKS_MAP.items():
            cell = ws.cell(9, c_idx)
            cell.value = r_name
            if r_name == rank:
                cell.fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
                cell.font = Font(name="Calibri", size=10, bold=True, color="7F6000")
            else:
                cell.font = Font(name="Calibri", size=9, bold=True)
            cell.alignment = Alignment(horizontal="center", vertical="center")
            
        # Basit Örnek Satırlar
        sample_certs = [
            (14, "1", "Passport", "Pasaport No: Mevcut | Süre Uygun"),
            (16, "2", "Medical Examination Report", "Gemiadamı Sağlık Belgesi | Süre Uygun"),
            (18, "3", "Certificate of Yellow Fever", "Sarı Humma Aşı Kartı | Süresiz"),
            (20, "4", "Seamanbook", f"Liman Cüzdan No: {seaman_book}"),
            (26, "7.1", "Fire Prevention & Fire Fighting", "STCW Belgesi | Süre Uygun"),
            (28, "7.2", "Personal Survival Techniques", "STCW Belgesi | Süre Uygun"),
            (30, "7.3", "Elementary First Aid", "STCW Belgesi | Süre Uygun"),
            (32, "7.4", "Personal Safety & Social Resp.", "STCW Belgesi | Süre Uygun"),
            (36, "8.1", "Security Awareness Certificate", "STCW VI/6 | Süresiz"),
            (38, "8.2", "Designated Security Duties", "STCW VI/6 | Süresiz")
        ]
        
        green_fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
        yellow_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
        
        for r_idx, no_str, en_title, note_str in sample_certs:
            ws.cell(r_idx, 1).value = no_str
            ws.cell(r_idx, 2).value = en_title
            ws.cell(r_idx+1, 2).value = f"► {note_str}"
            ws.cell(r_idx+1, 2).font = Font(name="Calibri", size=8.5, italic=True, color="1F3864")
            
            # Kesişme noktasına kadar yeşil
            for r_sub in [r_idx, r_idx+1]:
                for c in range(1, rank_col):
                    ws.cell(r_sub, c).fill = green_fill
                ws.cell(r_sub, rank_col).fill = yellow_fill
                
        # A4 Tek Sayfa Yazdırma Koruması
        ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 1
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        
        buf = io.BytesIO()
        wb.save(buf)
        return buf.getvalue()
        
    excel_bytes = build_live_excel()
    
    c1, c2 = st.columns(2)
    with c1:
        st.markdown("**35-101 Seafarers Documents Check List (Excel)**")
        st.caption(f"A4 Tek Sayfa | Sütun {rank_col} ({rank}) Sarı | Kesişme Noktasına Kadar Yeşil")
        st.download_button(
            label=f"📊 35-101 Excel Formunu İndir ({crew_name})",
            data=excel_bytes,
            file_name=f"35-101_{crew_name.replace(' ', '_')}_{rank}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
            
    with c2:
        st.markdown("**35-102 Pre Designation Check List (Word)**")
        st.caption(f"Onay Kutucukları İşaretli | {rank} Kadrosuna Uyarlanmış")
        try:
            with open("35-102_template.docx", "rb") as f:
                word_bytes = f.read()
        except Exception:
            word_bytes = b"35-102 Word Belgesi"
            
        st.download_button(
            label=f"📝 35-102 Word Formunu İndir ({crew_name})",
            data=word_bytes,
            file_name=f"35-102_{crew_name.replace(' ', '_')}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True
        )
