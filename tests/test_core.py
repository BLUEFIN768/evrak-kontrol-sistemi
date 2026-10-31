import unittest
from datetime import datetime, timedelta
import sys
import os

# Testlerin ana dizindeki modülleri bulabilmesi için sys.path ayarı
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

from crew_core import Evidence, Result, _check_dates, verify_identity

class TestCrewCore(unittest.TestCase):
    
    def test_verify_identity(self):
        """Madde 3.8: İsim eşleşmesine güvenilmez, TCKN ve Cüzdan No test edilir."""
        # Başarılı eşleşme (Büyük/küçük harf ve boşluk toleransı)
        self.assertTrue(verify_identity("12345678901", " A123456 ", "12345678901", "a123456"))
        
        # Hatalı eşleşme
        self.assertFalse(verify_identity("12345678901", "A123456", "99999999999", "A123456"))
        self.assertFalse(verify_identity("", "", "12345678901", "A123456"))

    def test_check_dates(self):
        """Madde 3.4 ve Madde 4: Tarih geçerlilik ve özel durum (Sarı Humma) kontrolleri."""
        past_date = datetime.now() - timedelta(days=100)
        future_date = datetime.now() + timedelta(days=100)
        
        # 1. Veriliş tarihi yoksa -> REVIEW (Madde 3.4)
        e_no_date = Evidence(code="PASSPORT", issue_date=None)
        self.assertEqual(_check_dates(e_no_date, "X"), "REVIEW")
        
        # 2. Süresi dolmuş -> EXPIRED
        e_expired = Evidence(code="MEDICAL_CERT", issue_date=past_date, expiry_date=past_date)
        self.assertEqual(_check_dates(e_expired, "X"), "EXPIRED")
        
        # 3. Geçerli -> OK
        e_valid = Evidence(code="SEAMAN_BOOK", issue_date=past_date, expiry_date=future_date)
        self.assertEqual(_check_dates(e_valid, "X"), "OK")
        
        # 4. Sarı Humma ömür boyu varsayımı (Madde 4)
        e_yellow = Evidence(code="YELLOW_FEVER", issue_date=past_date, expiry_date=None)
        self.assertEqual(_check_dates(e_yellow, "X"), "OK")

if __name__ == '__main__':
    unittest.main()
