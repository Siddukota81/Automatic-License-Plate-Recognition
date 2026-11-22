"""
Unit tests for rto_lookup.py module.

Covers:
- Active current Andhra Pradesh RTO lookups (e.g. AP-21 Kurnool)
- Statewide unified series (AP-39 and AP-40)
- Historical pre-bifurcation RTOs transferred to Telangana (e.g. AP-09 Hyderabad)
- Category-specific series (e.g. AP-18 RTC/Police)
- Unknown AP RTO codes (e.g. AP-99)
- Non-AP state registrations (e.g. HR-26 Haryana, KA-05 Karnataka)
- Error handling for missing CSV dataset (FileNotFoundError)
- Data verification status and official source URL validation
- Required non-location tracking disclaimer
"""

import unittest
import sys
from pathlib import Path

# Add project root to path
PROJECT_ROOT = Path(__file__).resolve().parent.parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from rto_lookup import (
    APDirectory,
    lookup_rto,
    MappingStatus,
    DISCLAIMER_TEXT,
)


class TestRTOLookup(unittest.TestCase):

    @classmethod
    def setUpClass(cls):
        cls.directory = APDirectory()

    def test_active_current_ap_rto(self):
        """Test lookup of an active Andhra Pradesh RTO (e.g. Kurnool AP-21)."""
        res = self.directory.lookup("AP", "21")
        self.assertTrue(res.is_ap)
        self.assertEqual(res.status, MappingStatus.CURRENT_AP_RTO)
        self.assertEqual(res.rto_code, "21")
        self.assertEqual(res.office_name, "Kurnool RTO")
        self.assertEqual(res.district_or_region, "Kurnool")
        self.assertEqual(res.verification_status, "Verified")
        self.assertTrue(res.source_url.startswith("https://"))

        # Test single digit with zero padding (e.g. AP 07 -> Guntur)
        res_guntur = self.directory.lookup("AP", "7")
        self.assertEqual(res_guntur.rto_code, "07")
        self.assertEqual(res_guntur.office_name, "Guntur RTO")
        self.assertEqual(res_guntur.status, MappingStatus.CURRENT_AP_RTO)

    def test_statewide_unified_series(self):
        """Test AP 39 and AP 40 statewide unified series."""
        res39 = self.directory.lookup("AP", "39")
        self.assertTrue(res39.is_ap)
        self.assertEqual(res39.status, MappingStatus.STATEWIDE_AP_SERIES)
        self.assertEqual(res39.rto_code, "39")
        self.assertIn("Andhra Pradesh", res39.district_or_region)
        self.assertEqual(res39.verification_status, "Verified")

        res40 = self.directory.lookup("AP", "40")
        self.assertEqual(res40.status, MappingStatus.STATEWIDE_AP_SERIES)

    def test_historical_pre_bifurcation_rto(self):
        """Test historical codes transferred to Telangana after 2014 bifurcation."""
        # AP 09 -> Hyderabad Central
        res09 = self.directory.lookup("AP", "09")
        self.assertTrue(res09.is_ap)
        self.assertEqual(res09.status, MappingStatus.HISTORICAL_PRE_BIFURCATION_RTO)
        self.assertEqual(res09.office_name, "Hyderabad Central (Khairtabad)")
        self.assertIn("Telangana", res09.historical_or_current)
        self.assertEqual(res09.verification_status, "Verified")

        # AP 36 -> Warangal
        res36 = self.directory.lookup("AP", "36")
        self.assertEqual(res36.status, MappingStatus.HISTORICAL_PRE_BIFURCATION_RTO)
        self.assertEqual(res36.office_name, "Warangal RTO")

    def test_category_specific_series(self):
        """Test category-specific registrations like AP 08 (Guntur RTC / Additional)."""
        res08 = self.directory.lookup("AP", "08")
        self.assertTrue(res08.is_ap)
        self.assertEqual(res08.status, MappingStatus.CATEGORY_SPECIFIC_AP)
        self.assertIn("RTC", res08.office_name)

    def test_unknown_ap_rto(self):
        """Test unknown or unallocated AP RTO code."""
        res = self.directory.lookup("AP", "99")
        self.assertTrue(res.is_ap)
        self.assertEqual(res.status, MappingStatus.UNKNOWN_AP_RTO)
        self.assertIn("not present in the verified", res.status_description)

    def test_non_ap_state_lookup(self):
        """Test non-AP states like Haryana (HR) and Karnataka (KA)."""
        res_hr = self.directory.lookup("HR", "26")
        self.assertFalse(res_hr.is_ap)
        self.assertEqual(res_hr.status, MappingStatus.NON_AP_REGISTRATION)
        self.assertEqual(res_hr.state, "Haryana")
        self.assertIn("not applicable", res_hr.status_description)

        res_ka = self.directory.lookup("KA", "01")
        self.assertFalse(res_ka.is_ap)
        self.assertEqual(res_ka.state, "Karnataka")

    def test_invalid_state_input(self):
        """Test empty or None state code input."""
        res = self.directory.lookup("", "21")
        self.assertEqual(res.status, MappingStatus.INVALID_REGISTRATION)

        res_none = self.directory.lookup(None, None)
        self.assertEqual(res_none.status, MappingStatus.INVALID_REGISTRATION)

    def test_missing_csv_file_raises_error(self):
        """Test that instantiating APDirectory with a non-existent path raises FileNotFoundError."""
        with self.assertRaises(FileNotFoundError):
            APDirectory(csv_path="non_existent_rto_file.csv")

    def test_disclaimer_presence(self):
        """Test that non-location disclaimer text is present in the module and results."""
        self.assertIn("physical location", DISCLAIMER_TEXT)
        self.assertIn("residential address", DISCLAIMER_TEXT)
        res = lookup_rto("AP", "21")
        self.assertEqual(res.disclaimer, DISCLAIMER_TEXT)


if __name__ == "__main__":
    unittest.main(verbosity=2)
