
import pandas as pd
from dataclasses import dataclass
from typing import Optional, List, Dict, Any
from datetime import datetime

# --- VERİ MODELLERİ ---

@dataclass
class Evidence:
    """UAB PDF'inden, AI'dan veya manuel girişten gelen ham belge verisi."""
    code: str  # Belge kodu (örn: "PASSPORT", "SEAMAN_BOOK")
    document_no: Optional[str] = None
    issue_date: Optional[datetime] = None
    expiry_date: Optional[datetime] = None
    verified: bool = False # Madde 3.5: AI çıktısı verified=False başlar
    raw_text: Optional[str] = None

@dataclass
class Result:
    """evaluate() fonksiyonundan çıkan sonuç objesi."""
    code: str
    status: str  # "OK", "MISSING", "REVIEW", "EXPIRED"
    message: str
    evidence_used: Optional[Evidence] = None

# --- ŞABLON SATIR İMZALARI (ITEM_SIGNATURES) ---
# Madde 3.6: Şablon satır yapısı doğrulanır. 35-101 Excel satırlarına karşılık gelen kodlar.
ITEM_SIGNATURES = {
    # Örnek satır indeksleri (Gerçek Excel şablonuna göre güncellenmelidir)
    10: "PASSPORT",
    11: "SEAMAN_BOOK",
    12: "MEDICAL_CERT",
    13: "YELLOW_FEVER",
    # Eğitimler
    15: "BASIC_SAFETY",
    # ... Diğer belgeler buraya eklenecek
}

def load_rules_from_excel(filepath: str, rank: str, vessel_type: str = "DEFAULT") -> Dict[str, str]:
    """
    Madde 3.1: Kural kaynağı 35-101 Excel'idir. Rütbeye göre 'X' (Zorunlu) veya 'D' (Atamaya bağlı) okur.
    """
    df = pd.read_excel(filepath, sheet_name="DOC_MATRIX") # Sayfa adı şablona göre uyarlanmalı
    # Excel'den okuma mantığı buraya inşa edilecek
    rules = {}
    return rules

def _check_dates(evidence: Evidence, rule_type: str) -> str:
    """
    Madde 5: Yeni tarih kuralı yalnız bu fonksiyon içinde yazılır.
    """
    if not evidence.issue_date:
        return "REVIEW" # Tarih okunamadıysa REVIEW (Madde 3.4)
    
    # Süresiz veya özel kurallar (örn: Sarı Humma - Madde 4)
    if evidence.code == "YELLOW_FEVER" and not evidence.expiry_date:
        return "OK" # Ömür boyu geçerli varsayımı
        
    if evidence.expiry_date and evidence.expiry_date < datetime.now():
        return "EXPIRED"
        
    return "OK"

def evaluate(evidences: List[Evidence], rank: str, excel_path: str) -> List[Result]:
    """
    Ana değerlendirme fonksiyonu.
    Veri akışı: UAB PDF / AI / elle tablo → list[Evidence] → evaluate() → list[Result]
    """
    rules = load_rules_from_excel(excel_path, rank)
    results = []
    
    # Her bir kural (X veya D) için evidence listesini kontrol et
    # Eşleşme, tarih ve doğrulama kontrolleri (maskeli isim/TCKN) burada yapılacak.
    
    return results
