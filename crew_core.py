import pandas as pd
from dataclasses import dataclass
from typing import Optional, List, Dict, Any
from datetime import datetime
import re

# ==========================================
# BÖLÜM 1: VERİ MODELLERİ VE ŞABLON YAPISI
# ==========================================

@dataclass
class Evidence:
    """UAB PDF'inden, AI'dan veya manuel girişten gelen ham belge verisi."""
    code: str  # Belge kodu (örn: "PASSPORT", "SEAMAN_BOOK")
    document_no: Optional[str] = None
    issue_date: Optional[datetime] = None
    expiry_date: Optional[datetime] = None
    verified: bool = False # Madde 3.5: AI çıktısı verified=False başlar[cite: 1]
    raw_text: Optional[str] = None

@dataclass
class Result:
    """evaluate() fonksiyonundan çıkan sonuç objesi."""
    code: str
    status: str  # "OK", "MISSING", "REVIEW", "EXPIRED"
    message: str
    evidence_used: Optional[Evidence] = None

# --- ŞABLON SATIR İMZALARI (ITEM_SIGNATURES) ---
# Madde 3.6: Şablon satır yapısı doğrulanır. 35-101 Excel satırlarına karşılık gelen kodlar[cite: 1].
ITEM_SIGNATURES = {
    10: "PASSPORT",
    11: "SEAMAN_BOOK",
    12: "MEDICAL_CERT",
    13: "YELLOW_FEVER",
    15: "BASIC_SAFETY",
    # Diğer belgeler Excel'deki satır numaralarına göre buraya eklenecek
}

def load_rules_from_excel(filepath: str, rank: str, vessel_type: str = "DEFAULT") -> Dict[str, str]:
    """
    Madde 3.1: Kural kaynağı 35-101 Excel'idir. Rütbeye göre 'X' (Zorunlu) veya 'D' (Atamaya bağlı) okur[cite: 1].
    """
    df = pd.read_excel(filepath, sheet_name="DOC_MATRIX")
    rules = {}
    # Excel okuma mantığı projenin ilerleyen aşamalarında buraya entegre edilecek
    return rules

def _check_dates(evidence: Evidence, rule_type: str) -> str:
    """
    Madde 5: Yeni tarih kuralı yalnız bu fonksiyon içinde yazılır[cite: 1].
    """
    if not evidence.issue_date:
        return "REVIEW" # Tarih okunamadıysa REVIEW (Madde 3.4)[cite: 1]
    
    if evidence.code == "YELLOW_FEVER" and not evidence.expiry_date:
        return "OK" # Sarı humma: bitiş yoksa ömür boyu (Madde 4)[cite: 1]
        
    if evidence.expiry_date and evidence.expiry_date < datetime.now():
        return "EXPIRED"
        
    return "OK"

# ==========================================
# BÖLÜM 2: KİMLİK DOĞRULAMA VE DEĞERLENDİRME
# ==========================================

def verify_identity(uab_tckn: str, uab_cuzdan_no: str, personel_tckn: str, personel_cuzdan_no: str) -> bool:
    """
    Madde 3.8: UAB çıktısında ad/soyad maskelidir (E**R). 
    İsim eşleştirmesine asla güvenme; TCKN ve cüzdan no kullan[cite: 1].
    """
    if not uab_tckn or not personel_tckn:
        return False
    
    tckn_match = str(uab_tckn).strip() == str(personel_tckn).strip()
    cuzdan_match = str(uab_cuzdan_no).strip().upper() == str(personel_cuzdan_no).strip().upper()
    
    return tckn_match and cuzdan_match

def parse_uab_pdf(pdf_text: str) -> List[Evidence]:
    """
    UAB "Gemi Adamları" doğrulama çıktısını (metin katmanlı PDF) okur[cite: 1].
    """
    evidences = []
    # RegEx ve ayrıştırma mantığı buraya eklenecek
    return evidences

def
