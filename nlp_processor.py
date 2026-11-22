"""
NLP Text Processing and Vehicle Registration Entity Extraction Module.

This module processes raw OCR text outputs from license plate recognition pipelines.
It provides:
1. Normalization: Cleans noise, special characters, and non-alphanumeric artifacts.
2. Positional Heuristic Tokenization: Breaks registration numbers into semantic components.
3. Cautious Context-Aware Error Correction: Identifies and suggests corrections for
   common optical character confusion pairs (e.g. O/0, I/1, B/8, S/5, Z/2) based on
   syntactic slot grammar (State [alpha], RTO [numeric], Series [alpha], Number [numeric]).
4. Entity Extraction: Extracts state code, RTO code, registration series, and vehicle number.
5. Format Validation: Validates Andhra Pradesh and pan-Indian registration structures,
   categorizing outputs into VALID_AP, NON_AP, UNCERTAIN, or INVALID_FORMAT.

IMPORTANT ARCHITECTURAL NOTE:
Text processing in this module combines rule-based tokenization, morphological
string normalization, and positional character grammar. Regular expressions and
grammatical heuristics are used as deterministic information extraction tools, not
as a substitute for full statistical language models.
"""

from dataclasses import dataclass, field, asdict
from enum import Enum
import re
from typing import Dict, List, Optional, Tuple, Any


class RegistrationValidationStatus(str, Enum):
    VALID_AP = "VALID_AP"
    NON_AP = "NON_AP"
    UNCERTAIN = "UNCERTAIN"
    INVALID_FORMAT = "INVALID_FORMAT"


@dataclass
class OCRCorrection:
    slot: str  # 'STATE', 'RTO', 'SERIES', 'NUMBER'
    original_char: str
    corrected_char: str
    position_index: int
    rule_name: str
    confidence: float  # Heuristic confidence in correction (0.0 to 1.0)


@dataclass
class RegistrationEntities:
    state_code: Optional[str] = None
    rto_code: Optional[str] = None
    series: Optional[str] = None
    vehicle_number: Optional[str] = None


@dataclass
class ProcessedRegistration:
    raw_text: str
    normalized_text: str
    corrected_text: str
    entities: RegistrationEntities
    validation_status: RegistrationValidationStatus
    is_ap: bool
    is_uncertain: bool
    corrections_applied: List[OCRCorrection] = field(default_factory=list)
    uncertainty_reasons: List[str] = field(default_factory=list)
    status_description: str = ""

    def to_dict(self) -> Dict[str, Any]:
        data = asdict(self)
        data["validation_status"] = self.validation_status.value
        return data


# Common OCR Confusions mapping
CHAR_TO_DIGIT_MAP = {
    "O": "0",
    "D": "0",
    "Q": "0",
    "I": "1",
    "L": "1",
    "Z": "2",
    "S": "5",
    "G": "6",
    "B": "8",
}

DIGIT_TO_CHAR_MAP = {
    "0": "O",
    "1": "I",
    "2": "Z",
    "5": "S",
    "8": "B",
}

# Standard 2-letter state codes in India
VALID_STATE_CODES = {
    "AP", "AR", "AS", "BR", "CG", "CH", "DD", "DL", "DN", "GA",
    "GJ", "HP", "HR", "JH", "JK", "KA", "KL", "LA", "LD", "MH",
    "ML", "MN", "MP", "MZ", "NL", "OD", "OR", "PB", "PY", "RJ",
    "SK", "TG", "TN", "TR", "TS", "UK", "UP", "UA", "WB", "BH"
}


def normalize_ocr_text(raw_text: Optional[str]) -> str:
    """
    Normalizes raw OCR text:
    - Converts to uppercase.
    - Strips leading and trailing whitespaces, newlines, tabs.
    - Removes non-alphanumeric clutter characters (e.g., '|', '!', '_', '-', '.', '/', '\\').
    - Consolidates internal spacing.
    """
    if not raw_text:
        return ""

    text = raw_text.upper().strip()
    # Remove unwanted punctuation and noise symbols common in OCR plate borders
    text = re.sub(r"[^A-Z0-9\s]", " ", text)
    # Collapse multiple whitespaces into a single space
    text = re.sub(r"\s+", " ", text).strip()
    return text


def clean_alphanumeric_compact(text: str) -> str:
    """Removes all whitespace to yield a compact alphanumeric token string."""
    return re.sub(r"\s+", "", text)


def tokenize_components(compact_text: str) -> Optional[Tuple[str, str, str, str]]:
    """
    Heuristically splits a compact registration string into 4 logical slots:
    (state_candidate, rto_candidate, series_candidate, number_candidate)

    Expected standard pattern:
    - State: 2 chars (e.g. AP)
    - RTO: 1 or 2 chars (e.g. 21, or legacy 2)
    - Series: 0 to 3 chars (e.g. AB, D, Z, or empty)
    - Number: 1 to 4 chars (e.g. 1234)
    """
    if len(compact_text) < 4:
        return None

    state_part = compact_text[:2]
    rest = compact_text[2:]

    # From the right end, vehicle registration number is 1 to 4 digits (or digit-confusions)
    confusion_chars = "".join(CHAR_TO_DIGIT_MAP.keys())
    match_num = re.search(r"([0-9" + confusion_chars + r"]{1,4})$", rest)
    if not match_num:
        match_num = re.search(r"([0-9A-Z]{1,4})$", rest)
        if not match_num:
            return None

    num_part = match_num.group(1)
    middle = rest[: len(rest) - len(num_part)]

    # If middle is empty (e.g., pure numeric stream like '211234')
    if not middle:
        if len(num_part) >= 3:
            rto_part = num_part[:2]
            ser_part = ""
            num_part = num_part[2:]
            return (state_part, rto_part, ser_part, num_part)
        return None

    # Middle contains RTO (1-2 chars) and optionally series (0-3 chars)
    # 1. Single character middle -> RTO is 1 char (legacy), no series
    if len(middle) == 1:
        return (state_part, middle, "", num_part)

    # 2. Two character middle
    if len(middle) == 2:
        # Check if 1st is digit and 2nd is a distinctive series letter (not an OCR digit confusion like O or I)
        if middle[0].isdigit() and middle[1].isalpha() and middle[1] not in ("O", "I", "D", "Q"):
            return (state_part, middle[0], middle[1], num_part)
        else:
            return (state_part, middle, "", num_part)

    # 3. Three or more characters in middle
    # If 1st is digit and 2nd is letter and NOT an OCR confusion, it's a 1-digit legacy RTO
    if middle[0].isdigit() and middle[1].isalpha() and middle[1] not in ("O", "I", "D", "Q", "B", "S", "Z"):
        return (state_part, middle[0], middle[1:], num_part)
    else:
        # Standard 2-character RTO
        return (state_part, middle[:2], middle[2:], num_part)


def apply_cautious_corrections(
    state_token: str,
    rto_token: str,
    series_token: str,
    number_token: str
) -> Tuple[RegistrationEntities, List[OCRCorrection], List[str]]:
    """
    Applies position-aware OCR error corrections with full transparency.
    Returns:
        (RegistrationEntities, corrections_applied, uncertainty_reasons)
    """
    corrections: List[OCRCorrection] = []
    uncertainty: List[str] = []

    # 1. State Slot (Expected: Alphabetic, 2 chars)
    corr_state = []
    for idx, ch in enumerate(state_token):
        if ch.isdigit():
            if ch in DIGIT_TO_CHAR_MAP:
                c_char = DIGIT_TO_CHAR_MAP[ch]
                corrections.append(
                    OCRCorrection(
                        slot="STATE",
                        original_char=ch,
                        corrected_char=c_char,
                        position_index=idx,
                        rule_name=f"Digit-to-Alpha ({ch}->{c_char}) in State slot",
                        confidence=0.85,
                    )
                )
                corr_state.append(c_char)
                uncertainty.append(f"Position {idx}: Ambiguous digit '{ch}' corrected to '{c_char}' in state code.")
            else:
                corr_state.append(ch)
                uncertainty.append(f"Position {idx}: Uncorrectable digit '{ch}' found in state code slot.")
        else:
            corr_state.append(ch)
    final_state = "".join(corr_state)

    # Specific common OCR confusion for AP state code: e.g. "4P" -> "AP"
    if len(final_state) == 2 and final_state[0] == "4" and final_state[1] == "P":
        corrections.append(
            OCRCorrection(
                slot="STATE",
                original_char="4",
                corrected_char="A",
                position_index=0,
                rule_name="Digit '4' corrected to 'A' in AP prefix",
                confidence=0.90,
            )
        )
        final_state = "AP"
        uncertainty.append("Position 0: OCR glyph '4' corrected to 'A' for state prefix.")

    # 2. RTO Slot (Expected: Numeric, 2 digits)
    corr_rto = []
    for idx, ch in enumerate(rto_token):
        pos = len(state_token) + idx
        if ch.isalpha():
            if ch in CHAR_TO_DIGIT_MAP:
                c_char = CHAR_TO_DIGIT_MAP[ch]
                corrections.append(
                    OCRCorrection(
                        slot="RTO",
                        original_char=ch,
                        corrected_char=c_char,
                        position_index=pos,
                        rule_name=f"Alpha-to-Digit ({ch}->{c_char}) in RTO slot",
                        confidence=0.90,
                    )
                )
                corr_rto.append(c_char)
                uncertainty.append(f"Position {pos}: OCR letter '{ch}' in RTO slot corrected to digit '{c_char}'.")
            else:
                corr_rto.append(ch)
                uncertainty.append(f"Position {pos}: Unexpected letter '{ch}' in numeric RTO slot.")
        else:
            corr_rto.append(ch)
    final_rto = "".join(corr_rto).zfill(2)

    # 3. Series Slot (Expected: Alphabetic, 0-3 chars)
    corr_series = []
    for idx, ch in enumerate(series_token):
        pos = len(state_token) + len(rto_token) + idx
        if ch.isdigit():
            if ch in DIGIT_TO_CHAR_MAP:
                c_char = DIGIT_TO_CHAR_MAP[ch]
                corrections.append(
                    OCRCorrection(
                        slot="SERIES",
                        original_char=ch,
                        corrected_char=c_char,
                        position_index=pos,
                        rule_name=f"Digit-to-Alpha ({ch}->{c_char}) in Series slot",
                        confidence=0.85,
                    )
                )
                corr_series.append(c_char)
                uncertainty.append(f"Position {pos}: Digit '{ch}' in registration series corrected to '{c_char}'.")
            else:
                corr_series.append(ch)
                uncertainty.append(f"Position {pos}: Unexpected digit '{ch}' in alphabetic series slot.")
        else:
            corr_series.append(ch)
    final_series = "".join(corr_series)

    # 4. Number Slot (Expected: Numeric, 1-4 digits)
    corr_num = []
    for idx, ch in enumerate(number_token):
        pos = len(state_token) + len(rto_token) + len(series_token) + idx
        if ch.isalpha():
            if ch in CHAR_TO_DIGIT_MAP:
                c_char = CHAR_TO_DIGIT_MAP[ch]
                corrections.append(
                    OCRCorrection(
                        slot="NUMBER",
                        original_char=ch,
                        corrected_char=c_char,
                        position_index=pos,
                        rule_name=f"Alpha-to-Digit ({ch}->{c_char}) in Number slot",
                        confidence=0.92,
                    )
                )
                corr_num.append(c_char)
                uncertainty.append(f"Position {pos}: OCR letter '{ch}' in vehicle number slot corrected to '{c_char}'.")
            else:
                corr_num.append(ch)
                uncertainty.append(f"Position {pos}: Unrecognized letter '{ch}' in vehicle number slot.")
        else:
            corr_num.append(ch)
    final_num = "".join(corr_num)

    entities = RegistrationEntities(
        state_code=final_state,
        rto_code=final_rto,
        series=final_series if final_series else None,
        vehicle_number=final_num if final_num else None,
    )
    return entities, corrections, uncertainty


def process_registration_text(raw_text: Optional[str]) -> ProcessedRegistration:
    """
    Main processing pipeline for vehicle license plate text.
    Executes normalization, tokenization, entity extraction,
    cautious error correction, and format validation.
    """
    normalized = normalize_ocr_text(raw_text)
    compact = clean_alphanumeric_compact(normalized)

    # Empty or severely truncated input
    if not compact:
        return ProcessedRegistration(
            raw_text=raw_text or "",
            normalized_text="",
            corrected_text="",
            entities=RegistrationEntities(),
            validation_status=RegistrationValidationStatus.INVALID_FORMAT,
            is_ap=False,
            is_uncertain=True,
            uncertainty_reasons=["Empty or non-alphanumeric OCR output."],
            status_description="No valid alphanumeric characters detected.",
        )

    # Check for Bharat Series (BH)
    # Format: YY BH #### XX (e.g., 22BH1234AA)
    bh_match = re.match(r"^(\d{2})BH(\d{4})([A-Z]{1,2})$", compact)
    if bh_match:
        year, num, ser = bh_match.groups()
        entities = RegistrationEntities(
            state_code="BH",
            rto_code=year,
            series=ser,
            vehicle_number=num,
        )
        return ProcessedRegistration(
            raw_text=raw_text or "",
            normalized_text=normalized,
            corrected_text=compact,
            entities=entities,
            validation_status=RegistrationValidationStatus.NON_AP,
            is_ap=False,
            is_uncertain=False,
            status_description="Bharat Series (BH) Pan-India vehicle registration format.",
        )

    # Tokenize standard format
    tokens = tokenize_components(compact)
    if not tokens:
        return ProcessedRegistration(
            raw_text=raw_text or "",
            normalized_text=normalized,
            corrected_text=compact,
            entities=RegistrationEntities(),
            validation_status=RegistrationValidationStatus.INVALID_FORMAT,
            is_ap=False,
            is_uncertain=True,
            uncertainty_reasons=["Text structure does not match standard Indian registration layout."],
            status_description="Unrecognized registration structure.",
        )

    s_tok, r_tok, ser_tok, num_tok = tokens
    entities, corrections, uncertainty = apply_cautious_corrections(s_tok, r_tok, ser_tok, num_tok)

    # Reconstruct corrected clean string
    series_part = f" {entities.series}" if entities.series else ""
    num_part = f" {entities.vehicle_number}" if entities.vehicle_number else ""
    rto_part = f" {entities.rto_code}" if entities.rto_code else ""
    corrected_str = f"{entities.state_code}{rto_part}{series_part}{num_part}".strip()

    is_ap = (entities.state_code == "AP")
    is_uncertain = len(corrections) > 0 or len(uncertainty) > 0

    # Format Validation
    if is_ap:
        # Check validity: State 'AP', numeric 2-digit RTO, valid number 1-4 digits
        if entities.rto_code and entities.rto_code.isdigit() and entities.vehicle_number and entities.vehicle_number.isdigit():
            if is_uncertain:
                status = RegistrationValidationStatus.UNCERTAIN
                desc = f"Andhra Pradesh format matched with {len(corrections)} position-aware OCR correction(s)."
            else:
                status = RegistrationValidationStatus.VALID_AP
                desc = "Valid Andhra Pradesh registration format."
        else:
            status = RegistrationValidationStatus.INVALID_FORMAT
            desc = "Malformed Andhra Pradesh registration components."
    else:
        # Non-AP vehicle
        if entities.state_code in VALID_STATE_CODES and entities.rto_code and entities.rto_code.isdigit() and entities.vehicle_number and entities.vehicle_number.isdigit():
            if is_uncertain:
                status = RegistrationValidationStatus.UNCERTAIN
                desc = f"Non-AP Indian state registration ({entities.state_code}) with OCR corrections."
            else:
                status = RegistrationValidationStatus.NON_AP
                desc = f"Standard registration for state '{entities.state_code}'."
        else:
            status = RegistrationValidationStatus.INVALID_FORMAT
            desc = f"Invalid or unsupported registration format for state code '{entities.state_code}'."

    return ProcessedRegistration(
        raw_text=raw_text or "",
        normalized_text=normalized,
        corrected_text=corrected_str,
        entities=entities,
        validation_status=status,
        is_ap=is_ap,
        is_uncertain=is_uncertain,
        corrections_applied=corrections,
        uncertainty_reasons=uncertainty,
        status_description=desc,
    )


def format_registration_summary(nlp_result: ProcessedRegistration, rto_result: Optional[Any] = None) -> str:
    """
    Format a clean, structured human-readable summary of the NLP extraction
    and RTO database lookup results.
    """
    lines = []
    lines.append("=" * 60)
    lines.append("AUTOMATIC LICENSE PLATE RECOGNITION - STRUCTURED REPORT")
    lines.append("=" * 60)
    lines.append(f"Original OCR Text:       {nlp_result.raw_text}")
    lines.append(f"Normalized Registration: {nlp_result.normalized_text}")
    if nlp_result.corrected_text != nlp_result.normalized_text:
        lines.append(f"Corrected Candidate:     {nlp_result.corrected_text}")

    lines.append("\nExtracted Entities:")
    e = nlp_result.entities
    state_disp = e.state_code if e.state_code else "N/A"
    lines.append(f"  State Code:             {state_disp}")
    lines.append(f"  RTO Code:               {e.rto_code or 'N/A'}")
    lines.append(f"  Series:                 {e.series or 'N/A'}")
    lines.append(f"  Vehicle Number:         {e.vehicle_number or 'N/A'}")
    lines.append(f"  Validation Status:      {nlp_result.validation_status.value} ({nlp_result.status_description})")

    if rto_result:
        lines.append("\nRegional Information (RTO Registry):")
        lines.append(f"  State:                  {rto_result.state}")
        if rto_result.is_ap:
            if rto_result.office_name:
                lines.append(f"  Office Name:            {rto_result.office_name}")
            if rto_result.district_or_region:
                lines.append(f"  District / Region:      {rto_result.district_or_region}")
            if rto_result.registration_category:
                lines.append(f"  Registration Category:  {rto_result.registration_category}")
            if rto_result.historical_or_current:
                lines.append(f"  Historical / Current:   {rto_result.historical_or_current}")
            if rto_result.effective_period:
                lines.append(f"  Effective Period:       {rto_result.effective_period}")
            lines.append(f"  Mapping Status:         {rto_result.status.value}")
            if rto_result.verification_status:
                lines.append(f"  Verification Status:    {rto_result.verification_status}")
            if rto_result.source_url:
                lines.append(f"  Source:                 {rto_result.source_url}")
        else:
            lines.append(f"  Mapping Status:         {rto_result.status.value}")
            lines.append(f"  Note:                   {rto_result.status_description}")

    if nlp_result.corrections_applied:
        lines.append("\nApplied OCR Corrections (Context-Aware):")
        for corr in nlp_result.corrections_applied:
            lines.append(
                f"  - [{corr.slot}] '{corr.original_char}' -> '{corr.corrected_char}' "
                f"(pos {corr.position_index}, rule: {corr.rule_name}, conf: {corr.confidence:.2f})"
            )

    if nlp_result.uncertainty_reasons:
        lines.append("\nUncertainty Warnings:")
        for reason in nlp_result.uncertainty_reasons:
            lines.append(f"  ! {reason}")

    lines.append("\nDISCLAIMER:")
    lines.append(
        "  Registration codes represent the administrative registering authority\n"
        "  and do not infer or verify the vehicle's current physical location\n"
        "  or the owner's personal residential address."
    )
    lines.append("=" * 60)
    return "\n".join(lines)

