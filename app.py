import streamlit as st
import datetime
import io
import re
from dateutil.relativedelta import relativedelta
import pypdf
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
import docx
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

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
    st.markdown("<p style='text-align: center; color: #666;'>Bu sistem KVKK uyumludur. Yüklenen hiçbir evrak sunucuda saklanmaz.</p>", unsafe_allow_html=True)
    
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

# 19 Rütbe ve 35-101 Sütunları
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

# Yan Panel
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
    f"⏳ **Asgari Kontrat:** {min_months} Ay ({'Zabit' if is_officer else 'Stajyer / Alt Kadro'})\n\n"
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
                    if not sb:
                        m_s = re.search(r'\b(S00\d{5,7}|S0\d{6,8}|S\d{7,8})\b', text, re.IGNORECASE)
                        if m_s:
                            sb = m_s.group(1).upper()
                        if not sb:
                            m_cuz = re.search(r'(?:Cüzdan\s*No|Seaman\s*Book\s*No|Cüzdan\s*Seri)\s*[:\.]?\s*([A-Z0-9]{6,12})', text, re.IGNORECASE)
                            if m_cuz and m_cuz.group(1).upper() not in ["KAYDI", "SURETI", "BELGESI", "SISTEMI", "TURKIYE"]:
                                sb = m_cuz.group(1).upper()
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
    st.success(f"{len(uploaded_files)} adet dosya RAM üzerinde analiz edildi (Sıfır Disk Depolama).")
    
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
        
    # UAB Doğrulama Butonu & Hızlı Kopyalama Alanı
    st.markdown(
        f"""
        <div style="background-color: #f0f4f8; border: 1px solid #c2d1e0; border-left: 5px solid #1F3864; padding: 12px 18px; border-radius: 6px; margin: 10px 0;">
            <div style="display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap;">
                <div>
                    <span style="font-weight: bold; color: #1F3864; font-size: 14.5px;">🌐 T.C. UAB Gemiadamı Doğrulama Portalı</span><br>
                    <span style="font-size: 13px; color: #333;">Sorgulama İçin: <b>TCKN:</b> <code>{tckn}</code> | <b>Cüzdan No:</b> <code>{seaman_book}</code></span>
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
    
    # Endorsement Kontrolü
    if not has_cra_file:
        st.warning(f"⚠️ **DİKKAT:** Yüklenen evraklar arasında personelin **Bayrak Onay Belgesi (Endorsement / CRA)** bulunamadı! 35-101 ve 35-102 formlarında bu madde 'Eksik' (Kırmızı) olarak işlenecektir.")
        
    st.divider()
    st.subheader("📥 Çıktı Dosyaları")
    
    # =========================================================================
    # 1. ORİJİNAL ŞABLONDAN 35-101 EXCEL DOLDURUCU (KUSURSUZ FORMAT + 2. SAYFA)
    # =========================================================================
    def fill_excel_from_template():
        # Önceden kusursuz hazırladığımız V5 dosyasını doğrudan şablon olarak açar
        template_name = "35-101_Seafarers_Documents_Check_List_YASIN_DOSER_V5.xlsx"
        wb = openpyxl.load_workbook(template_name)
        
        # 1. SHEET: 35-101 Check List
        ws1 = wb["35-101 Check List"]
        ws1['C6'] = vessel_name
        ws1['C8'] = join_date.strftime("%d.%m.%Y")
        ws1['N6'] = crew_name
        ws1['N8'] = rank
        
        # Sütun 23 başlığını rütbeye göre güncelle (Cadet ise Deck Cadet/Engine Cadet)
        if "Cadet" in rank:
            ws1.cell(9, 23).value = rank
            
        # Sütun sarı dolgusunu ayarla
        yellow_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
        for r in range(14, 84):
            ws1.cell(r, rank_col).fill = yellow_fill
            
        # Liman Cüzdan Satırına (R20 / R21) personelin kendi cüzdan numarasını yaz
        ws1.cell(21, 2).value = f"Liman Cüzdan No: {seaman_book} | Bitiş: 2030"
        
        # Endorsement Satırı (R82 / R83): Dosyada CRA yoksa KESİNLİKLE KIRMIZI yap!
        if not has_cra_file:
            red_fill = PatternFill(start_color="FF4D4D", end_color="FF4D4D", fill_type="solid")
            ws1.cell(83, 2).value = "[EKSİK - DOSYADA BAYRAK ONAYI / CRA BULUNAMADI!]"
            ws1.cell(83, 2).font = Font(name="Calibri", size=9.5, bold=True, color="FFFFFF")
            for r_end in [82, 83]:
                for c in range(1, rank_col):
                    ws1.cell(r_end, c).fill = red_fill
                ws1.cell(r_end, rank_col).fill = PatternFill(start_color="F8CBAD", end_color="F8CBAD", fill_type="solid")
                
        # 2. SHEET: Evrak Denetim Raporu
        ws2 = wb["Evrak Denetim Raporu"]
        ws2.cell(4, 2).value = crew_name
        ws2.cell(4, 5).value = vessel_name
        ws2.cell(5, 2).value = rank
        ws2.cell(5, 5).value = vessel_flag
        ws2.cell(6, 2).value = tckn
        ws2.cell(6, 5).value = ecdis_model if ecdis_model else "-"
        ws2.cell(7, 2).value = seaman_book
        ws2.cell(7, 5).value = join_date.strftime("%d.%m.%Y")
        
        # Cüzdan satırını raporda güncelle (Row 13)
        ws2.cell(13, 5).value = seaman_book
        
        # Endorsement satırını raporda güncelle (Row 42)
        if not has_cra_file:
            ws2.cell(42, 4).value = "EKSİK (BULUNAMADI)"
            ws2.cell(42, 4).fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
            ws2.cell(42, 5).value = "-"
            ws2.cell(42, 7).value = "-"
            
        buf = io.BytesIO()
        wb.save(buf)
        return buf.getvalue()

    # =========================================================================
    # 2. ORİJİNAL ŞABLONDAN 35-102 WORD DOLDURUCU (KUSURSUZ ŞABLON)
    # =========================================================================
    def fill_word_from_template():
        template_name = "35-102_Pre_Designation_Check_List_YASIN_DOSER_FINAL.docx"
        doc = docx.Document(template_name)
        table = doc.tables[0]
        
        # Üst Bilgileri Güncelle
        table.rows[0].cells[0].text = vessel_name
        table.rows[0].cells[3].text = join_date.strftime("%d.%m.%Y")
        
        table.rows[1].cells[0].text = crew_name
        table.rows[1].cells[3].text = vessel_flag
        
        table.rows[2].cells[0].text = rank
        table.rows[2].cells[3].text = ecdis_model if ecdis_model else "N/A"
        
        # Rütbe Stajyer (Cadet) ise ECDIS ve BTM maddelerini N/A yap
        if not is_officer:
            # Row 21: Type Specific ECDIS -> N/A
            table.rows[21].cells[-1].text = "Yes ☐   No ☐   N/A ☑"
            table.rows[21].cells[1].paragraphs[1].runs[0].text = f"► {rank} görevi için N/A (Gerekmiyor)"
            
            # Row 22: IMO Model ECDIS -> N/A
            table.rows[22].cells[-1].text = "Yes ☐   No ☐   N/A ☑"
            table.rows[22].cells[1].paragraphs[1].runs[0].text = f"► {rank} görevi için N/A (Gerekmiyor)"
            
            # Row 23: BTM/ERM -> N/A
            table.rows[23].cells[-1].text = "Yes ☐   No ☐   N/A ☑"
            table.rows[23].cells[1].paragraphs[1].runs[0].text = f"► {rank} görevi için N/A (Gerekmiyor)"
            
        # Endorsement Belgesi (Row 7): Dosyada CRA yoksa No ☑ yap ve ret gerekçesine yaz
        if not has_cra_file:
            table.rows[7].cells[-1].text = "Yes ☐   No ☑   N/A ☐"
            table.rows[7].cells[1].paragraphs[1].runs[0].text = "► EKSİK: Bayrak Onay Belgesi (Endorsement / CRA) dosyada bulunamadı!"
            
            # Alt Onay Bölümü: Rejected
            rej_note = f"{rank} ataması için Bayrak Onay Belgesi (Endorsement/CRA) ve şirket içi evraklar eksiktir."
            for r_app in [29, 30]:
                table.rows[r_app].cells[-1].text = f"Approved ☐         Rejected ☑\nReason to rejection : {rej_note}"
                
        buf = io.BytesIO()
        doc.save(buf)
        return buf.getvalue()

    excel_bytes = fill_excel_from_template()
    word_bytes = fill_word_from_template()
    
    c1, c2 = st.columns(2)
    with c1:
        st.markdown("**35-101 Seafarers Documents Check List (Excel)**")
        st.caption(f"Orijinal Şablon Düzeninde | 2. Sayfa Raporu Dolu | {rank} Sütunu Sarı | Endorsement Durumu Kırmızı (Eksik)")
        st.download_button(
            label=f"📊 35-101 Excel Formunu İndir ({crew_name})",
            data=excel_bytes,
            file_name=f"35-101_{crew_name.replace(' ', '_')}_{rank.replace(' ', '_')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
            
    with c2:
        st.markdown("**35-102 Pre Designation Check List (Word)**")
        st.caption(f"Orijinal Şirket Şablonunda | Kutucuklar İşaretli | {rank} Görevine Göre N/A Uyarlanmış")
        st.download_button(
            label=f"📝 35-102 Word Formunu İndir ({crew_name})",
            data=word_bytes,
            file_name=f"35-102_{crew_name.replace(' ', '_')}_{rank.replace(' ', '_')}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True
        )
