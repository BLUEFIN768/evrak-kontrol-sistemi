import streamlit as st
import io
import os
from datetime import datetime
from crew_core import Evidence, evaluate, verify_identity
from ai_extract import process_document
from fillers import fill_excel, fill_word

# Madde 3.7: Şifre ve API koda gömülmez, secrets.toml'dan okunur
def check_password():
    if "password_correct" not in st.session_state:
        st.session_state["password_correct"] = False

    if not st.session_state["password_correct"]:
        pwd = st.text_input("Sistem Parolası:", type="password")
        if st.button("Giriş"):
            if pwd == st.secrets["APP_PASSWORD"]:
                st.session_state["password_correct"] = True
                st.rerun()
            else:
                st.error("Hatalı parola!")
        return False
    return True

if not check_password():
    st.stop()

st.title("Evrak Kontrol Portalı v2")

# Madde 2: 4 sekmeli arayüz
tab1, tab2, tab3, tab4 = st.tabs(["1. Personel", "2. Evrak Yükle", "3. Kontrol", "4. Çıktı Al"])

if "evidences" not in st.session_state:
    st.session_state.evidences = []
if "results" not in st.session_state:
    st.session_state.results = []

# --- SEKME 1: PERSONEL BİLGİLERİ ---
with tab1:
    st.header("Personel Kimlik & Atama Bilgileri")
    col1, col2 = st.columns(2)
    with col1:
        p_name = st.text_input("Ad Soyad")
        p_tckn = st.text_input("TCKN")
        p_cuzdan = st.text_input("Cüzdan No")
    with col2:
        p_rank = st.text_input("Rütbe (Şablonla eşleşen)")
        p_vessel = st.text_input("Gemi Adı")
        is_tanker = st.checkbox("Gemi Tanker mi?") # Madde 4
        
    operator_name = st.text_input("Operatör Adı (İmza için)")

# --- SEKME 2: EVRAK YÜKLEME VE YAPAY ZEKA ---
with tab2:
    st.header("Evrak Yükleme")
    
    # 1. UAB PDF Yükleme
    uab_file = st.file_uploader("UAB Çıktısı (PDF)", type=["pdf"])
    if uab_file and st.button("UAB Oku"):
        st.info("UAB Okuma Modülü tetiklendi. (PDF parse mantığı crew_core.py içine eklendiğinde burası aktifleşecek)")
        
    # 2. Yapay Zeka ile Manuel Evrak Yükleme
    ai_file = st.file_uploader("Yapay Zeka için Evrak (Resim)", type=["jpg", "png", "jpeg"])
    if ai_file and st.button("Yapay Zeka ile Analiz Et"):
        try:
            ai_provider = st.secrets["AI_PROVIDER"]
            api_key = st.secrets["GEMINI_API_KEY"]
            model_name = st.secrets.get("AI_MODEL", "gemini-1.5-pro-latest")
            
            file_bytes = ai_file.getvalue()
            mime_type = ai_file.type
            
            with st.spinner('Yapay Zeka evrakı analiz ediyor...'):
                evidence = process_document(file_bytes, mime_type, ai_provider, api_key, model_name)
                st.session_state.evidences.append(evidence)
                st.success(f"Analiz Başarılı! Bulunan: {evidence.code}. (Operatör onayı bekleniyor)")
        except Exception as e:
            st.error(f"AI Analiz Hatası: {str(e)}")

# --- SEKME 3: KONTROL VE ONAY ---
with tab3:
    st.header("Kanıtlar ve Sonuçlar")
    
    st.subheader("Okunan Kanıtlar (Operatör Doğrulaması)")
    for i, ev in enumerate(st.session_state.evidences):
        status = "✅ Doğrulandı" if ev.verified else "⚠️ Bekliyor"
        st.write(f"**{ev.code}** | Tarih: {ev.issue_date} | Durum: {status}")
        
        if not ev.verified:
            if st.button(f"Orijinali Gördüm, Doğrula: {ev.code}", key=f"verify_{i}"):
                st.session_state.evidences[i].verified = True
                st.rerun()
                
    st.divider()
    
    if st.button("Kuralları İşlet (Evaluate)", type="primary"):
        excel_template_path = "templates/35-101.xlsx"
        if not os.path.exists(excel_template_path):
            st.error(f"Hata: Kural kaynağı {excel_template_path} bulunamadı!")
        else:
            try:
                # crew_core.py içindeki evaluate fonksiyonunu doğrudan çağırıyoruz
                st.session_state.results = evaluate(
                    evidences=st.session_state.evidences, 
                    rank=p_rank, 
                    excel_path=excel_template_path, 
                    is_tanker=is_tanker
                )
                st.success("Değerlendirme başarıyla tamamlandı. Çıktı Al sekmesine geçebilirsiniz.")
                
                # Değerlendirme özetini ekrana bas
                for r in st.session_state.results:
                    if r.status == "OK":
                        st.write(f"🟢 **{r.code}**: {r.message}")
                    elif r.status in ["REVIEW", "EXPIRED", "MISSING"]:
                        st.write(f"🔴 **{r.code}**: {r.message} ({r.status})")
            except Exception as e:
                st.error(f"Değerlendirme sırasında hata oluştu: {str(e)}")

# --- SEKME 4: ÇIKTI AL ---
with tab4:
    st.header("Evrakları İndir")
    st.write("Doldurulmuş formları diske yazmadan güvenle indirebilirsiniz.")
    
    personel_info = {
        "name": p_name, 
        "rank": p_rank, 
        "vessel": p_vessel, 
        "date": datetime.now().strftime("%d.%m.%Y")
    }
    
    if st.button("Excel ve Word Dosyalarını Hazırla"):
        if not st.session_state.results:
            st.warning("Önce 'Kontrol' sekmesinden kuralları işletmeniz gerekiyor.")
        else:
            excel_template = "templates/35-101.xlsx"
            word_template = "templates/35-102.docx"
            
            if not os.path.exists(excel_template) or not os.path.exists(word_template):
                st.error("Hata: 'templates' klasöründe orijinal 35-101.xlsx veya 35-102.docx bulunamadı!")
            else:
                try:
                    # RAM üzerinde (BytesIO) dosyaları oluşturma - Disk kullanımını engeller
                    excel_io = io.BytesIO()
                    word_io = io.BytesIO()
                    
                    # fillers.py fonksiyonları çalışıyor
                    fill_excel(excel_template, excel_io, st.session_state.results, personel_info)
                    fill_word(word_template, word_io, st.session_state.results, personel_info, operator_name)
                    
                    st.success("Dosyalar başarıyla bellekte oluşturuldu! Aşağıdan indirebilirsiniz.")
                    
                    st.download_button("35-101 İndir (Excel)", data=excel_io.getvalue(), file_name=f"35-101_{p_name}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                    st.download_button("35-102 İndir (Word)", data=word_io.getvalue(), file_name=f"35-102_{p_name}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
                
                except Exception as e:
                    st.error(f"Dosya oluşturulurken hata: {str(e)}")
