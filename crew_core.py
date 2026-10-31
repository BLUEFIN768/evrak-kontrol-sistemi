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
    code: str
    document_no: Optional[str] = None
    issue_date: Optional[datetime] = None
    expiry_date: Optional[datetime] = None
    verified: bool = False
    raw_text: Optional[str] = None

@dataclass
class Result:
    """evaluate() fonksiyonundan çıkan sonuç objesi."""
    code: str
    status: str
    message: str
    evidence_used: Optional[Evidence] = None

# --- ŞABLON SATIR İMZALARI (ITEM_SIGNATURES) ---
ITEM_SIGNATURES = {
    10: "PASSPORT",
    11: "SEAMAN_BOOK",
    12: "MEDICAL_CERT",
    13: "YELLOW_FEVER",
    15: "BASIC_SAFETY",
}

def load_rules_from_excel(filepath: str, rank: str, vessel_type: str = "DEFAULT") -> Dict[str, str]:
    """Madde 3.1: Kural kaynağı 35-101 Excel'idir."""
    df = pd.read_excel(filepath, sheet_name="DOC_MATRIX")
    rules = {}
    return rules

def _check_dates(evidence: Evidence, rule_type: str) -> str:
    """Madde 5: Yeni tarih kuralı yalnız bu fonksiyon içinde yazılır."""
    if not evidence.issue_date:
        return "REVIEW"
    
    if evidence.code == "YELLOW_FEVER" and not evidence.expiry_date:
        return "OK"
        
    if evidence.expiry_date and evidence.expiry_date < datetime.now():
        return "EXPIRED"
        
    return "OK"

# ==========================================
# BÖLÜM 2: KİMLİK DOĞRULAMA VE DEĞERLENDİRME
# ==========================================

def verify_identity(uab_tckn: str, uab_cuzdan_no: str, personel_tckn: str, personel_cuzdan_no: str) -> bool:
    """Madde 3.8: İsim eşleştirmesine asla güvenme; TCKN ve cüzdan no kullan."""
    if not uab_tckn or not personel_tckn:
        return False
    
    tckn_match = str(uab_tckn).strip() == str(personel_tckn).strip()
    cuzdan_match = str(uab_cuzdan_no).strip().upper() == str(personel_cuzdan_no).strip().upper()
    
    return tckn_match and cuzdan_match

def parse_uab_pdf(pdf_text: str) -> List[Evidence]:
    """UAB PDF okuyucu mock."""
    evidences = []
    return evidences

def evaluate(evidences: List[Evidence], rank: str, excel_path: str, is_tanker: bool = False) -> List[Result]:
    """Kuralları değerlendirir."""
    rules = load_rules_from_excel(excel_path, rank)
    results = []
    
    for doc_code, rule_type in rules.items():
        if "TANKER" in doc_code and not is_tanker:
            continue
            
        matched_evidence = next((e for e in evidences if e.code == doc_code), None)
        
        if not matched_evidence:
            if rule_type == 'D':
                results.append(Result(code=doc_code, status="OK", message="Atamaya Bağlı - Belge Yok", evidence_used=None))
            else:
                results.append(Result(code=doc_code, status="MISSING", message="Zorunlu Belge Eksik", evidence_used=None))
            continue
            
        date_status = _check_dates(matched_evidence, rule_type)
        
        if date_status != "OK":
            msg = "Belge Süresi Dolmuş" if date_status == "EXPIRED" else "Tarih Okunamadı veya Eksik"
            results.append(Result(code=doc_code, status=date_status, message=msg, evidence_used=matched_evidence))
            continue
            
        if not matched_evidence.verified:
            results.append(Result(code=doc_code, status="REVIEW", message="Orijinalden Doğrulanmalı (Kanıtsız Yes Yok)", evidence_used=matched_evidence))
            continue
            
        results.append(Result(code=doc_code, status="OK", message="Geçerli", evidence_used=matched_evidence))
        
    return results
