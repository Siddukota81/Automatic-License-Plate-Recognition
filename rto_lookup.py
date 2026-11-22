"""
Andhra Pradesh RTO Information Lookup Module.

This module provides structured lookup capabilities for Andhra Pradesh (AP)
vehicle registration codes based on a verified dataset of RTO authorities,
historical pre-bifurcation jurisdictions, and statewide unified series.

IMPORTANT DISCLAIMER:
Vehicle registration codes reflect the original or issuing regional transport
office authority. They DO NOT infer or track the vehicle's current physical
location, current operational territory, or the owner's personal residential address.
"""

from dataclasses import dataclass, asdict
from enum import Enum
import csv
import logging
from pathlib import Path
from typing import Dict, Optional, Tuple, Any

logger = logging.getLogger(__name__)

# Standard Indian State / Union Territory 2-letter codes mapping for non-AP identification
INDIAN_STATE_CODES: Dict[str, str] = {
    "AN": "Andaman and Nicobar Islands",
    "AP": "Andhra Pradesh",
    "AR": "Arunachal Pradesh",
    "AS": "Assam",
    "BR": "Bihar",
    "CG": "Chhattisgarh",
    "CH": "Chandigarh",
    "DD": "Daman and Diu",
    "DL": "Delhi",
    "DN": "Dadra and Nagar Haveli",
    "GA": "Goa",
    "GJ": "Gujarat",
    "HP": "Himachal Pradesh",
    "HR": "Haryana",
    "JH": "Jharkhand",
    "JK": "Jammu and Kashmir",
    "KA": "Karnataka",
    "KL": "Kerala",
    "LA": "Ladakh",
    "LD": "Lakshadweep",
    "MH": "Maharashtra",
    "ML": "Meghalaya",
    "MN": "Manipur",
    "MP": "Madhya Pradesh",
    "MZ": "Mizoram",
    "NL": "Nagaland",
    "OD": "Odisha",
    "OR": "Odisha (Old)",
    "PB": "Punjab",
    "PY": "Puducherry",
    "RJ": "Rajasthan",
    "SK": "Sikkim",
    "TG": "Telangana",
    "TN": "Tamil Nadu",
    "TR": "Tripura",
    "TS": "Telangana",
    "UK": "Uttarakhand",
    "UP": "Uttar Pradesh",
    "UA": "Uttarakhand (Old)",
    "WB": "West Bengal",
    "BH": "Bharat Series (Pan-India)",
}


class MappingStatus(str, Enum):
    CURRENT_AP_RTO = "CURRENT_AP_RTO"
    STATEWIDE_AP_SERIES = "STATEWIDE_AP_SERIES"
    HISTORICAL_PRE_BIFURCATION_RTO = "HISTORICAL_PRE_BIFURCATION_RTO"
    CATEGORY_SPECIFIC_AP = "CATEGORY_SPECIFIC_AP"
    UNKNOWN_AP_RTO = "UNKNOWN_AP_RTO"
    NON_AP_REGISTRATION = "NON_AP_REGISTRATION"
    INVALID_REGISTRATION = "INVALID_REGISTRATION"


@dataclass
class RTORecord:
    state_code: str
    rto_code: str
    office_name: str
    district_or_region: str
    registration_category: str
    historical_or_current: str
    effective_period: str
    source_url: str
    verification_status: str


DISCLAIMER_TEXT = (
    "Registration codes identify the issuing RTO authority. "
    "They do NOT infer the vehicle's current physical location or owner's residential address."
)


@dataclass
class RTOLookupResult:
    state: str
    state_code: str
    rto_code: Optional[str]
    is_ap: bool
    status: MappingStatus
    status_description: str
    office_name: Optional[str] = None
    district_or_region: Optional[str] = None
    registration_category: Optional[str] = None
    historical_or_current: Optional[str] = None
    effective_period: Optional[str] = None
    source_url: Optional[str] = None
    verification_status: Optional[str] = None
    disclaimer: str = DISCLAIMER_TEXT

    def to_dict(self) -> Dict[str, Any]:
        data = asdict(self)
        data["status"] = self.status.value
        return data


class APDirectory:
    """
    In-memory registry and lookup service for Andhra Pradesh RTO records.
    Loads data from data/ap_rto_codes.csv.
    """

    def __init__(self, csv_path: Optional[Path] = None):
        if csv_path is None:
            # Default to data/ap_rto_codes.csv relative to project root
            base_dir = Path(__file__).resolve().parent
            csv_path = base_dir / "data" / "ap_rto_codes.csv"

        self.csv_path = Path(csv_path)
        self._records: Dict[str, RTORecord] = {}
        self.load_dataset()

    def load_dataset(self) -> None:
        """Loads and parses the AP RTO CSV dataset."""
        if not self.csv_path.exists():
            raise FileNotFoundError(f"AP RTO dataset file not found at: '{self.csv_path}'")

        loaded: Dict[str, RTORecord] = {}
        with open(self.csv_path, mode="r", encoding="utf-8") as f:
            reader = csv.DictReader(f)
            required_fields = {
                "state_code",
                "rto_code",
                "office_name",
                "district_or_region",
                "registration_category",
                "historical_or_current",
                "effective_period",
                "source_url",
                "verification_status",
            }
            if not reader.fieldnames or not required_fields.issubset(set(reader.fieldnames)):
                raise ValueError(
                    f"Invalid CSV structure. Missing required headers in '{self.csv_path}'"
                )

            for line_no, row in enumerate(reader, start=2):
                rto_code = (row.get("rto_code") or "").strip().zfill(2)
                state_code = (row.get("state_code") or "").strip().upper()

                if not rto_code or not state_code:
                    logger.warning("Skipping invalid CSV record at line %d: %s", line_no, row)
                    continue

                record = RTORecord(
                    state_code=state_code,
                    rto_code=rto_code,
                    office_name=(row.get("office_name") or "").strip(),
                    district_or_region=(row.get("district_or_region") or "").strip(),
                    registration_category=(row.get("registration_category") or "").strip(),
                    historical_or_current=(row.get("historical_or_current") or "").strip(),
                    effective_period=(row.get("effective_period") or "").strip(),
                    source_url=(row.get("source_url") or "").strip(),
                    verification_status=(row.get("verification_status") or "").strip(),
                )
                loaded[rto_code] = record

        self._records = loaded

    def lookup(self, state_code: str, rto_code: Optional[str]) -> RTOLookupResult:
        """
        Lookup regional registration information by state and RTO code.
        """
        clean_state = (state_code or "").strip().upper()
        clean_rto = (rto_code or "").strip().zfill(2) if rto_code else None

        if not clean_state:
            return RTOLookupResult(
                state="Unknown",
                state_code="",
                rto_code=clean_rto,
                is_ap=False,
                status=MappingStatus.INVALID_REGISTRATION,
                status_description="Missing or empty state code in registration.",
            )

        # Check if it is an Andhra Pradesh registration
        if clean_state == "AP":
            if not clean_rto or not clean_rto.isdigit():
                return RTOLookupResult(
                    state="Andhra Pradesh",
                    state_code="AP",
                    rto_code=clean_rto,
                    is_ap=True,
                    status=MappingStatus.INVALID_REGISTRATION,
                    status_description="Andhra Pradesh registration detected, but RTO code is missing or non-numeric.",
                )

            if clean_rto in self._records:
                rec = self._records[clean_rto]

                # Distinguish status types
                if "Statewide" in rec.historical_or_current or clean_rto in ("39", "40"):
                    status = MappingStatus.STATEWIDE_AP_SERIES
                    desc = (
                        f"Statewide unified series '{rec.rto_code}' ({rec.office_name}) "
                        "applicable across all districts in Andhra Pradesh."
                    )
                elif "Category" in rec.historical_or_current or "Category" in rec.registration_category:
                    status = MappingStatus.CATEGORY_SPECIFIC_AP
                    desc = f"Category-specific series ({rec.registration_category}) in {rec.district_or_region}."
                elif "Historical" in rec.historical_or_current:
                    status = MappingStatus.HISTORICAL_PRE_BIFURCATION_RTO
                    desc = (
                        f"Historical code '{rec.rto_code}' originally allocated to {rec.district_or_region} "
                        f"({rec.historical_or_current})."
                    )
                else:
                    status = MappingStatus.CURRENT_AP_RTO
                    desc = f"Active Andhra Pradesh Regional Transport Office in {rec.district_or_region}."

                return RTOLookupResult(
                    state="Andhra Pradesh",
                    state_code="AP",
                    rto_code=rec.rto_code,
                    is_ap=True,
                    status=status,
                    status_description=desc,
                    office_name=rec.office_name,
                    district_or_region=rec.district_or_region,
                    registration_category=rec.registration_category,
                    historical_or_current=rec.historical_or_current,
                    effective_period=rec.effective_period,
                    source_url=rec.source_url,
                    verification_status=rec.verification_status,
                )
            else:
                return RTOLookupResult(
                    state="Andhra Pradesh",
                    state_code="AP",
                    rto_code=clean_rto,
                    is_ap=True,
                    status=MappingStatus.UNKNOWN_AP_RTO,
                    status_description=(
                        f"Registration code 'AP-{clean_rto}' is not present in the verified "
                        "Andhra Pradesh RTO registry."
                    ),
                )

        # Non-AP Registration
        state_name = INDIAN_STATE_CODES.get(clean_state, f"State/UT Code '{clean_state}'")
        return RTOLookupResult(
            state=state_name,
            state_code=clean_state,
            rto_code=clean_rto,
            is_ap=False,
            status=MappingStatus.NON_AP_REGISTRATION,
            status_description=(
                f"Vehicle registered under state '{state_name}' ({clean_state}). "
                "Andhra Pradesh regional lookup is not applicable."
            ),
        )


# Global default instance
_default_directory: Optional[APDirectory] = None


def get_default_directory() -> APDirectory:
    global _default_directory
    if _default_directory is None:
        _default_directory = APDirectory()
    return _default_directory


def lookup_rto(state_code: str, rto_code: Optional[str], directory: Optional[APDirectory] = None) -> RTOLookupResult:
    """
    Convenience function to perform RTO lookup.
    """
    dir_inst = directory or get_default_directory()
    return dir_inst.lookup(state_code, rto_code)
