"""
Unit tests for nlp_processor.py module.

Covers:
- Text normalization (removing noise, symbols, redundant spaces)
- Positional heuristic tokenization
- Cautious OCR confusion corrections (O/0, I/1, B/8, S/5, Z/2, 4/A)
- Transparency: verifying corrections_applied list and is_uncertain flag
- Format validation (VALID_AP, NON_AP, UNCERTAIN, INVALID_FORMAT)
- Bharat Series (BH) detection
- Edge cases: empty text, punctuation-only, gibberish, single tokens
- Summary formatting and disclaimer verification
"""

import unittest
import sys
from pathlib import Path

# Add project root to path
PROJECT_ROOT = Path(__file__).resolve().parent.parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from nlp_processor import (
    normalize_ocr_text,
    tokenize_components,
    process_registration_text,
    format_registration_summary,
    RegistrationValidationStatus,
)


class TestNLPProcessor(unittest.TestCase):

    def test_normalize_ocr_text(self):
        """Test normalization removes noise, keeps alphanumeric, and formats cleanly."""
        self.assertEqual(normalize_ocr_text("AP-21-AB-1234"), "AP 21 AB 1234")
        self.assertEqual(normalize_ocr_text("  ap . 21 / ab _ 1234  "), "AP 21 AB 1234")
        self.assertEqual(normalize_ocr_text("HR 26 DA 2330"), "HR 26 DA 2330")
        self.assertEqual(normalize_ocr_text(""), "")
        self.assertEqual(normalize_ocr_text(None), "")
        self.assertEqual(normalize_ocr_text("!@#$%^&*()_+"), "")

    def test_tokenize_components(self):
        """Test positional component tokenization."""
        # Standard 4 components
        s, r, ser, num = tokenize_components("AP21AB1234")
        self.assertEqual((s, r, ser, num), ("AP", "21", "AB", "1234"))

        # Single letter series
        s, r, ser, num = tokenize_components("AP21A1234")
        self.assertEqual((s, r, ser, num), ("AP", "21", "A", "1234"))

        # Single digit RTO (legacy)
        s, r, ser, num = tokenize_components("AP2A1234")
        self.assertEqual((s, r, ser, num), ("AP", "2", "A", "1234"))

        # No series component
        s, r, ser, num = tokenize_components("AP211234")
        self.assertEqual((s, r, ser, num), ("AP", "21", "", "1234"))

        # Invalid short string
        self.assertIsNone(tokenize_components("AP"))
        self.assertIsNone(tokenize_components("1234"))

    def test_valid_ap_registration(self):
        """Test valid AP vehicle registration strings without OCR confusion."""
        result = process_registration_text("AP 21 AB 1234")
        self.assertEqual(result.validation_status, RegistrationValidationStatus.VALID_AP)
        self.assertTrue(result.is_ap)
        self.assertFalse(result.is_uncertain)
        self.assertEqual(result.entities.state_code, "AP")
        self.assertEqual(result.entities.rto_code, "21")
        self.assertEqual(result.entities.series, "AB")
        self.assertEqual(result.entities.vehicle_number, "1234")
        self.assertEqual(len(result.corrections_applied), 0)

        # Unified series AP 39
        res39 = process_registration_text("AP39EF5678")
        self.assertEqual(res39.validation_status, RegistrationValidationStatus.VALID_AP)
        self.assertTrue(res39.is_ap)
        self.assertEqual(res39.entities.rto_code, "39")

    def test_ocr_confusion_corrections(self):
        """Test cautious context-aware correction of optical character confusions."""
        # 'I' in RTO slot and vehicle number slot should become '1'
        res1 = process_registration_text("AP 2I AB I234")
        self.assertEqual(res1.validation_status, RegistrationValidationStatus.UNCERTAIN)
        self.assertTrue(res1.is_uncertain)
        self.assertEqual(res1.entities.rto_code, "21")
        self.assertEqual(res1.entities.vehicle_number, "1234")
        self.assertGreater(len(res1.corrections_applied), 0)

        # 'O' in RTO slot and number slot should become '0'
        res2 = process_registration_text("AP O2 AB 233O")
        self.assertEqual(res2.entities.rto_code, "02")
        self.assertEqual(res2.entities.vehicle_number, "2330")

        # '4' in State slot should become 'A'
        res3 = process_registration_text("4P 21 AB 1234")
        self.assertEqual(res3.entities.state_code, "AP")

        # '8' in Series slot should become 'B'
        res4 = process_registration_text("AP 21 A8 1234")
        self.assertEqual(res4.entities.series, "AB")

        # Verify raw text is retained alongside corrected text
        self.assertEqual(res1.raw_text, "AP 2I AB I234")
        self.assertIn("AP 21 AB 1234", res1.corrected_text)

    def test_non_ap_indian_registration(self):
        """Test non-AP registrations like HR (Haryana) or KA (Karnataka)."""
        res_hr = process_registration_text("HR 26 DA 2330")
        self.assertEqual(res_hr.validation_status, RegistrationValidationStatus.NON_AP)
        self.assertFalse(res_hr.is_ap)
        self.assertEqual(res_hr.entities.state_code, "HR")
        self.assertEqual(res_hr.entities.rto_code, "26")
        self.assertEqual(res_hr.entities.series, "DA")
        self.assertEqual(res_hr.entities.vehicle_number, "2330")

        res_ka = process_registration_text("KA 05 MN 9999")
        self.assertEqual(res_ka.validation_status, RegistrationValidationStatus.NON_AP)
        self.assertEqual(res_ka.entities.state_code, "KA")

    def test_bharat_series(self):
        """Test Bharat Series (BH) detection."""
        res_bh = process_registration_text("22 BH 1234 AA")
        self.assertEqual(res_bh.validation_status, RegistrationValidationStatus.NON_AP)
        self.assertFalse(res_bh.is_ap)
        self.assertEqual(res_bh.entities.state_code, "BH")
        self.assertEqual(res_bh.entities.vehicle_number, "1234")

    def test_invalid_and_edge_cases(self):
        """Test invalid strings, empty input, and non-plate text."""
        # Empty input
        res_empty = process_registration_text("")
        self.assertEqual(res_empty.validation_status, RegistrationValidationStatus.INVALID_FORMAT)
        self.assertIsNone(res_empty.entities.state_code)

        # None input
        res_none = process_registration_text(None)
        self.assertEqual(res_none.validation_status, RegistrationValidationStatus.INVALID_FORMAT)

        # Gibberish
        res_gibberish = process_registration_text("HELLO WORLD")
        self.assertEqual(res_gibberish.validation_status, RegistrationValidationStatus.INVALID_FORMAT)

        # Numbers only
        res_num = process_registration_text("12345678")
        self.assertEqual(res_num.validation_status, RegistrationValidationStatus.INVALID_FORMAT)

    def test_format_registration_summary(self):
        """Verify report summary formatting and inclusion of required disclaimer."""
        res = process_registration_text("AP 21 AB 1234")
        summary = format_registration_summary(res)
        self.assertIn("AUTOMATIC LICENSE PLATE RECOGNITION", summary)
        self.assertIn("AP 21 AB 1234", summary)
        self.assertIn("State Code:             AP", summary)
        self.assertIn("RTO Code:               21", summary)
        self.assertIn("DISCLAIMER:", summary)
        self.assertIn("current physical location", summary)


if __name__ == "__main__":
    unittest.main(verbosity=2)
