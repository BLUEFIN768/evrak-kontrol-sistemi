import openpyxl
from docx import Document
from typing import List, Dict, Optional
from crew_core import Result

# --- WORD EŞLEŞTİRME SÖZLÜĞÜ (Madde 5) ---
# 35-102.docx tablosundaki satır indeksleri (0'dan başlar).
# Şablon değişirse bu indeksler kontrol edilmelidir.
word_answers = {
    "PASSPORT": 4,
    "SEAMAN_BOOK": 5,
    "MEDICAL_CERT": 6,
    "YELLOW_FEVER": 7,
    "BASIC_SAFETY": 10,
    # Diğer belgeler Word tablosundaki satırlarına göre eklenecek
}

def fill_excel(template_path: str, output_path: str, results: List[Result], personel_info: dict):
    """
    35-101 Excel dosyasını doldurur.
    Madde 3.2: Sadece C6, N6, C8, N8 ve sağdaki yeni sütunlara (Y:AD) yazar. 
    Türkçe açıklama satırlarına dokunmaz.
    """
    wb = openpyxl.load_workbook(template_path)
    ws = wb["DOC_MATRIX"] # Çalışma sayfası adını şablona göre ayarlayın

    # Temel personel bilgilerini belirli hücrelere yazıyoruz (Madde 3.2)
    ws["C6"] = personel_info.get("name", "")
    ws["N6"] = personel_info.get("rank", "")
    ws["C8"] = personel_info.get("date", "")
    ws["N8"] = personel_info.get("vessel", "")

    # Sonuçları Y:AD sütunlarına yazma mantığı
    # Gerçek şablondaki ITEM_SIGNATURES satırlarına göre ilgili Result'ları sağ sütunlara işleyeceğiz.
    # Örnek: ws[f"Y{row_index}"] = result.status
    
    wb.save(output_path)

def fill_word(template_path: str, output_path: str, results: List[Result], personel_info: dict, operator_name: str):
    """
    35-102 Word dosyasını doldurur.
    Madde 3.2: Yalnız değer hücreleri (satır 0-2, hücre 3 ve 9) ve Yes/No/N/A kutuları.
    Madde 3.3: Onay/imza alanlarına (satır 29-30) dokunulmaz, sadece operatör adı yazılır[cite: 1].
    """
    doc = Document(template_path)
    table = doc.tables[0] # Formun ilk tablosu

    # 1. Üst Bilgiler (Satır 0-2, Hücre 3 ve 9)[cite: 1]
    table.rows[0].cells[3].text = personel_info.get("name", "")
    table.rows[0].cells[9].text = personel_info.get("rank", "")
    table.rows[1].cells[3].text = personel_info.get("vessel", "")
    table.rows[1].cells[9].text = personel_info.get("date", "")
    
    # Sonuçları (Results) bir sözlüğe çevirelim ki kolay arayalım
    results_dict = {r.code: r for r in results}

    # 2. Belge Kontrol Kutuları (Yes / No / N/A)
    for code, row_index in word_answers.items():
        if row_index >= len(table.rows):
            continue # Şablon dışına çıkmayı engelle
            
        result = results_dict.get(code)
        
        # Madde 3.4: Kanıtsız "Yes" yok. Kanıt yoksa kutu boş kalır (None)[cite: 1].
        if not result or not result.evidence_used or not result.evidence_used.verified:
            # Kanıt yok veya doğrulanmamış -> Boş bırak
            continue
            
        if result.status == "OK":
            # Sadece onaylı ve geçerli belgelerde Yes sütununa (örneğin 3. hücre) "X" koyuyoruz
            table.rows[row_index].cells[3].text = "X"
        elif result.status == "MISSING":
            # No sütununa (örneğin 4. hücre) "X" koy
            table.rows[row_index].cells[4].text = "X"
            
        # N/A durumu için (Atamaya bağlı 'D' olup belge olmayanlar)
        if result and result.status == "OK" and result.evidence_used is None:
            table.rows[row_index].cells[5].text = "X" # N/A hücresi

    # 3. Alt Kısım - Onay ve İmza (Madde 3.3)[cite: 1]
    # Sadece operatör adını yaz, Approved/Rejected veya imza/tarih kısımlarına dokunma[cite: 1].
    # Örneğin 29. satırın operatör adı hücresi (şablona göre indeks güncellenebilir)
    if len(table.rows) > 29:
        table.rows[29].cells[3].text = operator_name

    doc.save(output_path)
