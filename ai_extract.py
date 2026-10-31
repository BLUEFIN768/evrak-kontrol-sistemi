import json
import google.generativeai as genai
from typing import List, Dict, Callable
from datetime import datetime
from crew_core import Evidence

# Madde 5: Yeni belge türü eklenecekse buradaki CODES listesi de güncellenir.
CODES = ["PASSPORT", "SEAMAN_BOOK", "MEDICAL_CERT", "YELLOW_FEVER", "BASIC_SAFETY"]

# AI'ın her sağlayıcıda aynı JSON şemasında çıktı vermesi için sabit istem (Madde 5).
PROMPT = f"""
Sen bir denizcilik evrak analiz asistanısın. Aşağıdaki belgeyi incele ve JSON formatında şu bilgileri çıkar:
- "code": Belgenin türü ({', '.join(CODES)} listesinden en uygun olanı seç).
- "document_no": Belge numarası.
- "issue_date": Veriliş tarihi (YYYY-MM-DD formatında).
- "expiry_date": Bitiş/Geçerlilik tarihi (YYYY-MM-DD formatında, yoksa null).

Sadece geçerli bir JSON formatında çıktı ver.
"""

def extract_with_gemini(data: bytes, mime: str, api_key: str, model_name: str) -> str:
    """Gemini API çağrısı yapar ve JSON string döndürür."""
    genai.configure(api_key=api_key)
    model = genai.GenerativeModel(model_name)
    
    prompt_parts = [
        PROMPT,
        {"mime_type": mime, "data": data}
    ]
    
    response = model.generate_content(prompt_parts)
    return response.text

# Madde 5: Yeni AI sağlayıcı ekleneceğinde fonksiyon (data, mime, api_key, model) -> str imzasını taşımalıdır.
PROVIDERS: Dict[str, Callable[[bytes, str, str, str], str]] = {
    "gemini": extract_with_gemini
    # İleride Anthropic eklendiğinde: "anthropic": extract_with_anthropic
}

def process_document(data: bytes, mime: str, provider_name: str, api_key: str, model_name: str) -> Evidence:
    """
    Seçilen sağlayıcı ile resmi okur, JSON'ı parse eder ve Evidence objesine çevirir.
    Madde 3.5: AI çıktısı verified=False başlar. Resmî doğrulama yalnız UAB PDF + operatör onayı ile olur[cite: 1].
    """
    if provider_name not in PROVIDERS:
        raise ValueError(f"Desteklenmeyen AI sağlayıcısı: {provider_name}")
        
    extract_func = PROVIDERS[provider_name]
    raw_json_str = extract_func(data, mime, api_key, model_name)
    
    # Markdown bloklarını temizle (```json ... ```)
    clean_json_str = raw_json_str.replace("```json", "").replace("```", "").strip()
    parsed_data = json.loads(clean_json_str)
    
    issue_dt = None
    if parsed_data.get("issue_date"):
        issue_dt = datetime.strptime(parsed_data["issue_date"], "%Y-%m-%d")
        
    expiry_dt = None
    if parsed_data.get("expiry_date"):
        expiry_dt = datetime.strptime(parsed_data["expiry_date"], "%Y-%m-%d")
    
    return Evidence(
        code=parsed_data.get("code", "UNKNOWN"),
        document_no=parsed_data.get("document_no"),
        issue_date=issue_dt,
        expiry_date=expiry_dt,
        verified=False, # Güvenlik kilidi (Madde 3.5)[cite: 1]
        raw_text=clean_json_str
    )
