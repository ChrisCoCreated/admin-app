from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Iterable


PLACEHOLDER_CATEGORIES = (
    "CLIENT",
    "STAFF",
    "RELATIVE",
    "FRIEND",
    "PROFESSIONAL",
    "CARE_HOME",
    "LOCATION",
    "ORGANISATION",
    "DATE",
    "PHONE",
    "EMAIL",
    "ADDRESS",
)

PLACEHOLDER_RE = re.compile(r"\[(?:[A-Z_]+)_\d{3}\]")
LOW_SIGNAL_PRONOUNS = {"he", "him", "his", "she", "her", "hers"}

EMAIL_RE = re.compile(r"\b[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}\b", re.IGNORECASE)

PHONE_RE = re.compile(
    r"(?<!\w)(?:"
    r"(?:\+44\s?7\d{3}|0?7\d{3})[\s.-]?\d{3}[\s.-]?\d{3}|"
    r"(?:\+44\s?|0)(?:1\d{3}|2\d|3\d{2})[\s.-]?\d{3,4}[\s.-]?\d{3,4}"
    r")(?!\w)"
)

EXACT_DATE_RE = re.compile(
    r"\b(?:\d{1,2}[/-]\d{1,2}[/-]\d{2,4}|"
    r"\d{1,2}\s+(?:jan(?:uary)?|feb(?:ruary)?|mar(?:ch)?|apr(?:il)?|may|jun(?:e)?|"
    r"jul(?:y)?|aug(?:ust)?|sep(?:tember)?|sept|oct(?:ober)?|nov(?:ember)?|dec(?:ember)?)"
    r"\s+\d{2,4})\b",
    re.IGNORECASE,
)

POSTCODE_RE = re.compile(
    r"\b(?:GIR\s?0AA|[A-PR-UWYZ][A-HK-Y]?\d[A-Z\d]?\s?\d[ABD-HJLNP-UW-Z]{2})\b",
    re.IGNORECASE,
)

STREET_ADDRESS_RE = re.compile(
    r"\b\d{1,5}\s+[A-Z][A-Za-z'.-]*(?:\s+[A-Z][A-Za-z'.-]*){0,5}\s+"
    r"(?:Street|St|Road|Rd|Avenue|Ave|Lane|Ln|Drive|Dr|Close|Court|Way|Gardens|"
    r"Grove|Place|Terrace|Crescent|Square)\b(?:,\s*" + POSTCODE_RE.pattern[2:-2] + r")?",
    re.IGNORECASE,
)

NUMBERED_FLAT_RE = re.compile(
    r"\b(?:Flat|House|Room)\s+\d+[A-Z]?(?:,\s*\d{1,5}\s+[A-Z][A-Za-z'.-]*(?:\s+[A-Z][A-Za-z'.-]*){0,5}\s+"
    r"(?:Street|St|Road|Rd|Avenue|Ave|Lane|Ln|Drive|Dr|Close|Court|Way|Gardens|Grove|Place|Terrace|Crescent|Square)\b)?",
    re.IGNORECASE,
)

NHS_RE = re.compile(r"\b(?:\d[ -]?){10}\b")

NI_RE = re.compile(
    r"\b(?!BG)(?!GB)(?!KN)(?!NK)(?!NT)(?!TN)(?!ZZ)"
    r"[A-CEGHJ-PR-TW-Z]{2}\s?\d{2}\s?\d{2}\s?\d{2}\s?[A-D]\b",
    re.IGNORECASE,
)

NAME_TOKEN = r"[A-Z][a-z]+(?:[-'][A-Z][a-z]+)?"
INITIAL_NAME = rf"[A-Z]\.?\s+{NAME_TOKEN}"
FULL_NAME = rf"{NAME_TOKEN}(?:\s+{NAME_TOKEN}){{1,2}}"
LABELLED_PERSON_VALUE = rf"(?:{FULL_NAME}|{INITIAL_NAME}|{NAME_TOKEN})"
TITLE_NAME_RE = re.compile(rf"\b((?:Mr|Mrs|Ms|Miss|Dr)\.?\s+{NAME_TOKEN})\b")
FULL_NAME_RE = re.compile(rf"\b({FULL_NAME})\b")
SINGLE_NAME_RE = re.compile(
    rf"\b({NAME_TOKEN})\b(?=\s+(?:is|was|said|says|visited|rang|called|emailed|asked|advised|confirmed|noted|discussed)\b)"
)

CLIENT_LABEL_RE = re.compile(
    rf"(?im)^\s*(?:[-*]\s*)?(?:\*\*)?\s*(?:client\s+name|client|resident|service\s+user|patient)"
    rf"\s*:\s*(?:\*\*)?\s*({LABELLED_PERSON_VALUE})\b"
)

STAFF_LABEL_RE = re.compile(
    rf"(?im)^\s*(?:[-*]\s*)?(?:\*\*)?\s*(?:carer|care\s+worker|support\s+worker|staff|key\s+worker|"
    rf"nurse|manager|senior)\s*:\s*(?:\*\*)?\s*({LABELLED_PERSON_VALUE})\b"
)

CLIENT_CONTEXT_RE = re.compile(
    rf"(?i:\b(?:resident|service user|client|patient)\s+(?:is\s+|was\s+|named\s+|called\s+)?)({FULL_NAME}|{NAME_TOKEN})\b"
)

RELATIVE_CONTEXT_RE = re.compile(
    rf"(?i:\b(?:daughter|son|brother|sister|mother|father|wife|husband|partner|"
    rf"step\s?sons?|stepsons?|granddaughter|grandson|niece|nephew|next of kin|nok)\s+"
    rf"(?:is\s+|was\s+|called\s+|named\s+)?)({FULL_NAME}|{NAME_TOKEN})\b"
)

FRIEND_CONTEXT_RE = re.compile(
    rf"(?i:\b(?:friend|neighbour|neighbor)\s+"
    rf"(?:is\s+|was\s+|called\s+|named\s+)?)({FULL_NAME}|{NAME_TOKEN})\b"
)

STAFF_CONTEXT_RE = re.compile(
    rf"(?i:\b(?:staff|carer|care worker|support worker|nurse|manager|senior|key worker)\s+"
    rf"(?:is\s+|was\s+|called\s+|named\s+)?)({FULL_NAME}|{NAME_TOKEN})\b"
)

PROFESSIONAL_CONTEXT_RE = re.compile(
    rf"(?i:\b(?:dr|doctor|gp|social worker|occupational therapist|ot|psychiatrist|"
    rf"consultant|solicitor|advocate|counsellor|pharmacist|speech therapist|district nurse|"
    rf"community nurse|professional)\s+)(?:is\s+|was\s+|called\s+|named\s+)?({FULL_NAME}|{NAME_TOKEN})\b"
)

PROFESSIONAL_ACTION_RE = re.compile(
    rf"(?i:\b(?:discussed|spoke with|spoke to|contacted|called|emailed|liaised with|reviewed with|"
    rf"updated)\s+)({FULL_NAME})\b"
)

CARE_HOME_CONTEXT_RE = re.compile(
    r"(?i:\b(?:feels safe at|safe at|returned to|back at|resident at|lives at|living at|staying at|"
    r"moved to|admitted to|discharged to|at)\s+)"
    r"([A-Z][A-Za-z'&-]*(?:\s+[A-Z][A-Za-z'&-]*){0,3})\b"
)

LOCATION_CONTEXT_RE = re.compile(
    r"(?i:\b(?:lives in|living in|based in|from|moved from|returned from|over in)\s+)"
    r"([A-Z][A-Za-z'&-]*(?:\s+[A-Z][A-Za-z'&-]*){0,3})\b"
)

ORGANISATION_RE = re.compile(
    r"\b([A-Z][A-Za-z'&.-]*(?:\s+[A-Z][A-Za-z'&.-]*){0,4}\s+"
    r"(?:Council|Hospital|Clinic|Practice|Team|Service|Agency|Trust|Housing|Police|"
    r"Ambulance|Pharmacy|Solicitors?|Homecare|Care))\b"
)

RISK_TERMS = (
    "alcohol risk",
    "self-neglect",
    "hoarding",
    "safeguarding",
    "falls",
    "fall risk",
    "capacity",
    "dols",
    "deprivation of liberty",
    "mobility issues",
)

CLINICAL_TERMS = (
    "dementia",
    "diabetes",
    "brain injury",
    "mobility",
    "medication",
    "care needs",
    "emotional wellbeing",
    "anxiety",
    "depression",
    "self neglect",
)

SINGLE_NAME_STOPWORDS = {
    "Care",
    "Discussed",
    "Returned",
    "Called",
    "Emailed",
    "Contacted",
    "Resident",
    "Client",
    "Patient",
    "Daughter",
    "Son",
    "Brother",
    "Sister",
    "Mother",
    "Father",
    "Manager",
    "Carer",
    "Staff",
}

PERSON_VALUE_STOPWORDS = {
    "Name",
    "Type",
    "Visit",
    "He",
    "Him",
    "His",
    "She",
    "Her",
    "Hers",
}

PERSON_CATEGORIES_FOR_FILTERING = {
    "PERSON",
    "CLIENT",
    "STAFF",
    "RELATIVE",
    "FRIEND",
    "PROFESSIONAL",
}

PERSON_PHRASE_STOPWORDS = {
    "Client Name",
    "Visit Type",
    "Morning Visit",
    "Lunch Visit",
    "Tea Visit",
    "Bedtime Visit",
    "Evening Visit",
    "Night Visit",
}

LOCATION_STOPWORDS = {
    "Monday",
    "Tuesday",
    "Wednesday",
    "Thursday",
    "Friday",
    "Saturday",
    "Sunday",
}

COUNTRY_NAMES = (
    "Australia",
    "Canada",
    "France",
    "Germany",
    "India",
    "Ireland",
    "Jamaica",
    "Nigeria",
    "Pakistan",
    "Poland",
    "Spain",
    "United States",
    "USA",
    "Wales",
    "Scotland",
)

FAMILY_COUNTRY_RE = re.compile(
    r"\b(?:daughter|son|brother|sister|mother|father|wife|husband|partner|family|children)"
    r"\s+(?:lives?|living|based|stays?|moved)\s+(?:in|overseas in|abroad in|to|from)\s+"
    rf"({'|'.join(re.escape(country) for country in COUNTRY_NAMES)})\b",
    re.IGNORECASE,
)

MIGRATION_HISTORY_RE = re.compile(
    r"\b(?:moved|came|arrived|migrated|relocated)\s+(?:to|from)\s+"
    rf"({'|'.join(re.escape(country) for country in COUNTRY_NAMES)})\b",
    re.IGNORECASE,
)

PROPERTY_DETAIL_RE = re.compile(
    r"\b(?:sell(?:ing)?|sold|owns?|owned|property|properties|house sale|flat sale|landlord|tenant)\b",
    re.IGNORECASE,
)

SMALL_SERVICE_RE = re.compile(r"\b\d{1,3}\s*beds?\b", re.IGNORECASE)

COMMUNITY_LINK_RE = re.compile(
    r"\b(?:church|choir|mosque|temple|synagogue|parish|community group|club)\b",
    re.IGNORECASE,
)

DISTINCTIVE_QUOTE_RE = re.compile(r"['\"].{35,}?['\"]")


@dataclass(frozen=True)
class CandidateFinding:
    entity_type: str
    start: int
    end: int
    score: float
    text: str


@dataclass(frozen=True)
class Finding:
    entity_type: str
    start: int
    end: int
    score: float
    replacement: str


@dataclass(frozen=True)
class ResidualIdentifier:
    value: str
    start: int
    end: int
    severity: str
    category: str
    reason: str


def is_valid_nhs_number(value: str) -> bool:
    digits = re.sub(r"\D", "", value)
    if len(digits) != 10:
        return False
    total = sum(int(digit) * weight for digit, weight in zip(digits[:9], range(10, 1, -1)))
    check = 11 - (total % 11)
    if check == 11:
        check = 0
    if check == 10:
        return False
    return check == int(digits[-1])


def detect_candidate_entities(text: str) -> list[CandidateFinding]:
    candidates: list[CandidateFinding] = []

    candidates.extend(_match_pattern(text, EMAIL_RE, "EMAIL", 0.99))
    candidates.extend(_match_pattern(text, PHONE_RE, "PHONE", 0.96))
    candidates.extend(_match_pattern(text, EXACT_DATE_RE, "DATE", 0.91))
    candidates.extend(_match_pattern(text, STREET_ADDRESS_RE, "ADDRESS", 0.97))
    candidates.extend(_match_pattern(text, NUMBERED_FLAT_RE, "ADDRESS", 0.88))
    candidates.extend(_match_pattern(text, POSTCODE_RE, "ADDRESS", 0.82))
    candidates.extend(_match_pattern(text, CLIENT_LABEL_RE, "CLIENT", 0.995, group=1))
    candidates.extend(_match_pattern(text, STAFF_LABEL_RE, "STAFF", 0.985, group=1))
    candidates.extend(_match_pattern(text, CLIENT_CONTEXT_RE, "CLIENT", 0.98, group=1))
    candidates.extend(_match_pattern(text, RELATIVE_CONTEXT_RE, "RELATIVE", 0.97, group=1))
    candidates.extend(_match_pattern(text, FRIEND_CONTEXT_RE, "FRIEND", 0.96, group=1))
    candidates.extend(_match_pattern(text, STAFF_CONTEXT_RE, "STAFF", 0.96, group=1))
    candidates.extend(_match_pattern(text, PROFESSIONAL_CONTEXT_RE, "PROFESSIONAL", 0.96, group=1))
    candidates.extend(_match_pattern(text, PROFESSIONAL_ACTION_RE, "PROFESSIONAL", 0.9, group=1))
    candidates.extend(_match_pattern(text, CARE_HOME_CONTEXT_RE, "CARE_HOME", 0.93, group=1))
    candidates.extend(_match_pattern(text, LOCATION_CONTEXT_RE, "LOCATION", 0.88, group=1))
    candidates.extend(_match_pattern(text, ORGANISATION_RE, "ORGANISATION", 0.9, group=1))
    candidates.extend(_match_pattern(text, TITLE_NAME_RE, "PERSON", 0.78, group=1))
    candidates.extend(_match_pattern(text, FULL_NAME_RE, "PERSON", 0.74, group=1))
    candidates.extend(_single_name_candidates(text))

    return remove_overlaps(candidates)


def remove_overlaps(findings: Iterable[CandidateFinding]) -> list[CandidateFinding]:
    ordered = sorted(findings, key=lambda item: (-item.score, item.start, -(item.end - item.start)))
    accepted: list[CandidateFinding] = []
    occupied: list[range] = []

    for finding in ordered:
        if any(finding.start < block.stop and finding.end > block.start for block in occupied):
            continue
        accepted.append(finding)
        occupied.append(range(finding.start, finding.end))

    return sorted(accepted, key=lambda item: item.start)


def detect_uk_social_care_pii(text: str) -> list[Finding]:
    findings: list[Finding] = []
    counters: dict[str, int] = {}
    values: dict[tuple[str, str], str] = {}

    for candidate in detect_candidate_entities(text):
        category = candidate.entity_type if candidate.entity_type != "PERSON" else "STAFF"
        key = (category, normalise_value(candidate.text))
        replacement = values.get(key)
        if replacement is None:
            counters[category] = counters.get(category, 0) + 1
            replacement = f"[{category}_{counters[category]:03d}]"
            values[key] = replacement
        findings.append(
            Finding(
                entity_type=category,
                start=candidate.start,
                end=candidate.end,
                score=candidate.score,
                replacement=replacement,
            )
        )

    return findings


def collect_preserved_details(text: str, pseudonymised_text: str) -> tuple[list[str], list[str]]:
    risks_preserved = _present_in_both(text, pseudonymised_text, RISK_TERMS)
    clinical_details_preserved = _present_in_both(text, pseudonymised_text, CLINICAL_TERMS)
    return risks_preserved, clinical_details_preserved


def scan_for_residual_identifiers(text: str) -> list[str]:
    return list(dict.fromkeys(item.reason for item in _scan_direct_identifiers(text, {})))


def assess_data_disclosure_risk(text: str, mapping: dict[str, str]) -> dict[str, object]:
    direct_items = _scan_direct_identifiers(text, mapping)
    indirect_items = _scan_indirect_identifiers(text)
    direct_identifiers = list(dict.fromkeys(item.value for item in direct_items))
    indirect_identifiers = list(dict.fromkeys(item.value for item in indirect_items))
    residual_spans = [
        {
            "start": item.start,
            "end": item.end,
            "severity": item.severity,
            "category": item.category,
            "reason": item.reason,
        }
        for item in _dedupe_residual_spans([*direct_items, *indirect_items])
    ]

    if direct_identifiers:
        return {
            "riskLevel": "HIGH",
            "safeToSend": False,
            "directIdentifiers": direct_identifiers,
            "indirectIdentifiers": indirect_identifiers,
            "residualSpans": residual_spans,
            "reason": "Unresolved direct identifier remains in the pseudonymised text.",
        }

    if len(indirect_identifiers) >= 3:
        return {
            "riskLevel": "MEDIUM",
            "safeToSend": True,
            "directIdentifiers": [],
            "indirectIdentifiers": indirect_identifiers,
            "residualSpans": residual_spans,
            "reason": "Three or more indirect identifiers remain; safe to send only with review.",
        }

    return {
        "riskLevel": "LOW",
        "safeToSend": True,
        "directIdentifiers": [],
        "indirectIdentifiers": indirect_identifiers,
        "residualSpans": residual_spans,
        "reason": "No unresolved direct identifiers found.",
    }


def normalise_value(value: str) -> str:
    return re.sub(r"\s+", " ", value.strip().lower())


def _match_pattern(
    text: str,
    pattern: re.Pattern[str],
    entity_type: str,
    score: float,
    *,
    group: int = 0,
) -> list[CandidateFinding]:
    candidates: list[CandidateFinding] = []
    for match in pattern.finditer(text):
        value = match.group(group).strip(" ,.")
        first_token = value.split()[0] if value.split() else ""
        if _is_low_signal_pronoun(value):
            continue
        if entity_type in PERSON_CATEGORIES_FOR_FILTERING and value in PERSON_VALUE_STOPWORDS:
            continue
        if entity_type == "PERSON" and (
            first_token in SINGLE_NAME_STOPWORDS or value in PERSON_PHRASE_STOPWORDS
        ):
            continue
        if not value or value in LOCATION_STOPWORDS:
            continue
        candidates.append(
            CandidateFinding(entity_type, match.start(group), match.end(group), score, value)
        )
    return candidates


def _single_name_candidates(text: str) -> list[CandidateFinding]:
    candidates: list[CandidateFinding] = []
    for match in SINGLE_NAME_RE.finditer(text):
        value = match.group(1)
        if _is_low_signal_pronoun(value):
            continue
        if value in SINGLE_NAME_STOPWORDS or value in LOCATION_STOPWORDS:
            continue
        candidates.append(CandidateFinding("PERSON", match.start(1), match.end(1), 0.68, value))
    return candidates


def _present_in_both(source_text: str, pseudonymised_text: str, terms: Iterable[str]) -> list[str]:
    source_lower = source_text.lower()
    output_lower = pseudonymised_text.lower()
    return [term for term in terms if term in source_lower and term in output_lower]


def _warn_matches(
    text: str,
    pattern: re.Pattern[str],
    prefix: str,
    *,
    group: int = 0,
) -> list[str]:
    warnings: list[str] = []
    for match in pattern.finditer(text):
        value = match.group(group).strip()
        if not value:
            continue
        warnings.append(f"{prefix}: {value}")
    return warnings


def _scan_direct_identifiers(text: str, mapping: dict[str, str]) -> list[ResidualIdentifier]:
    placeholder_ranges = _placeholder_ranges(text)
    direct: list[ResidualIdentifier] = []

    for placeholder, original in mapping.items():
        category = _placeholder_category(placeholder)
        if category in {"CLIENT", "STAFF", "RELATIVE", "FRIEND", "PROFESSIONAL"}:
            for part in _name_scan_parts(original):
                direct.extend(
                    _known_identifier_matches(
                        text,
                        placeholder_ranges,
                        part,
                        category,
                        f"Known {category.lower()} name remains",
                    )
                )
        elif category in {"CARE_HOME", "LOCATION", "ORGANISATION", "ADDRESS", "PHONE", "EMAIL"}:
            value = original.strip()
            direct.extend(
                _known_identifier_matches(
                    text,
                    placeholder_ranges,
                    value,
                    category,
                    f"Known {category.lower()} value remains",
                )
            )

    _collect_direct_matches(
        direct, text, placeholder_ranges, EMAIL_RE, "Likely email remains", "EMAIL"
    )
    _collect_direct_matches(
        direct, text, placeholder_ranges, PHONE_RE, "Likely phone number remains", "PHONE"
    )
    _collect_direct_matches(
        direct, text, placeholder_ranges, STREET_ADDRESS_RE, "Likely address remains", "ADDRESS"
    )
    _collect_direct_matches(
        direct, text, placeholder_ranges, POSTCODE_RE, "Likely postcode remains", "ADDRESS"
    )
    _collect_direct_matches(
        direct, text, placeholder_ranges, NHS_RE, "Likely NHS number remains", "NHS_NUMBER"
    )
    _collect_direct_matches(
        direct,
        text,
        placeholder_ranges,
        NI_RE,
        "Likely National Insurance number remains",
        "NI_NUMBER",
    )
    _collect_direct_matches(
        direct,
        text,
        placeholder_ranges,
        ORGANISATION_RE,
        "Likely organisation name remains",
        "ORGANISATION",
        group=1,
    )
    _collect_direct_matches(
        direct,
        text,
        placeholder_ranges,
        CARE_HOME_CONTEXT_RE,
        "Likely care home or named location remains",
        "CARE_HOME",
        group=1,
    )
    _collect_direct_matches(
        direct,
        text,
        placeholder_ranges,
        FULL_NAME_RE,
        "Likely personal name remains",
        "PERSON",
        group=1,
    )
    _collect_direct_matches(
        direct,
        text,
        placeholder_ranges,
        SINGLE_NAME_RE,
        "Likely named person remains",
        "PERSON",
        group=1,
    )

    return _dedupe_residual_spans(direct)


def _scan_indirect_identifiers(text: str) -> list[ResidualIdentifier]:
    placeholder_ranges = _placeholder_ranges(text)
    indirect: list[ResidualIdentifier] = []

    indirect.extend(
        _extract_residual_matches(
            text, placeholder_ranges, EXACT_DATE_RE, "DATE", "Exact date remains"
        )
    )
    indirect.extend(
        _extract_residual_matches(
            text,
            placeholder_ranges,
            FAMILY_COUNTRY_RE,
            "FAMILY_LOCATION",
            "Family location detail remains",
        )
    )
    indirect.extend(
        _extract_residual_matches(
            text,
            placeholder_ranges,
            MIGRATION_HISTORY_RE,
            "MIGRATION_HISTORY",
            "Migration history detail remains",
        )
    )
    indirect.extend(
        _extract_residual_matches(
            text,
            placeholder_ranges,
            PROPERTY_DETAIL_RE,
            "PROPERTY_DETAIL",
            "Property detail remains",
        )
    )
    indirect.extend(
        _extract_residual_matches(
            text, placeholder_ranges, SMALL_SERVICE_RE, "SERVICE_SIZE", "Service size remains"
        )
    )
    indirect.extend(
        _extract_residual_matches(
            text, placeholder_ranges, COMMUNITY_LINK_RE, "COMMUNITY_LINK", "Community link remains"
        )
    )
    indirect.extend(
        _extract_residual_matches(
            text,
            placeholder_ranges,
            DISTINCTIVE_QUOTE_RE,
            "DISTINCTIVE_QUOTE",
            "Distinctive quote remains",
        )
    )

    return _dedupe_residual_spans(indirect)


def _placeholder_category(placeholder: str) -> str:
    match = re.match(r"\[([A-Z_]+)_\d{3}\]", placeholder)
    return match.group(1) if match else ""


def _name_scan_parts(name: str) -> list[str]:
    cleaned = re.sub(r"\b(?:Mr|Mrs|Ms|Miss|Dr)\.?\s+", "", name.strip(), flags=re.IGNORECASE)
    parts = [part for part in re.split(r"\s+", cleaned) if part]
    values: list[str] = []
    if len(parts) >= 2:
        values.append(" ".join(parts))
    values.extend(parts)
    return list(dict.fromkeys(values))


def _contains_identifier(text: str, identifier: str) -> bool:
    escaped = re.escape(identifier.strip())
    if not escaped:
        return False
    pattern = re.compile(rf"(?<![A-Za-z]){escaped}(?:['’]s)?(?![A-Za-z])", re.IGNORECASE)
    return bool(pattern.search(text))


def _known_identifier_matches(
    text: str,
    placeholder_ranges: list[range],
    identifier: str,
    category: str,
    reason: str,
) -> list[ResidualIdentifier]:
    if _is_low_signal_pronoun(identifier):
        return []
    escaped = re.escape(identifier.strip())
    if not escaped:
        return []
    pattern = re.compile(rf"(?<![A-Za-z]){escaped}(?:['’]s)?(?![A-Za-z])", re.IGNORECASE)
    return [
        ResidualIdentifier(
            value=identifier,
            start=match.start(),
            end=match.end(),
            severity="direct",
            category=category,
            reason=reason,
        )
        for match in pattern.finditer(text)
        if not _inside_placeholder(match.start(), match.end(), placeholder_ranges)
    ]


def _collect_direct_matches(
    direct: list[ResidualIdentifier],
    text: str,
    placeholder_ranges: list[range],
    pattern: re.Pattern[str],
    reason: str,
    category: str,
    *,
    group: int = 0,
) -> None:
    for match in pattern.finditer(text):
        value = match.group(group).strip()
        if (
            value
            and not _is_low_signal_pronoun(value)
            and not _inside_placeholder(match.start(group), match.end(group), placeholder_ranges)
        ):
            direct.append(
                ResidualIdentifier(
                    value=value,
                    start=match.start(group),
                    end=match.end(group),
                    severity="direct",
                    category=category,
                    reason=reason,
                )
            )


def _is_low_signal_pronoun(value: str) -> bool:
    return value.strip().lower() in LOW_SIGNAL_PRONOUNS


def _extract_residual_matches(
    text: str,
    placeholder_ranges: list[range],
    pattern: re.Pattern[str],
    category: str,
    reason: str,
    *,
    group: int = 0,
) -> list[ResidualIdentifier]:
    return [
        ResidualIdentifier(
            value=match.group(group).strip(),
            start=match.start(group),
            end=match.end(group),
            severity="indirect",
            category=category,
            reason=reason,
        )
        for match in pattern.finditer(text)
        if match.group(group).strip()
        and not _inside_placeholder(match.start(group), match.end(group), placeholder_ranges)
    ]


def _placeholder_ranges(text: str) -> list[range]:
    return [range(match.start(), match.end()) for match in PLACEHOLDER_RE.finditer(text)]


def _inside_placeholder(start: int, end: int, placeholder_ranges: list[range]) -> bool:
    return any(
        start in placeholder_range or end - 1 in placeholder_range
        for placeholder_range in placeholder_ranges
    )


def _dedupe_residual_spans(items: Iterable[ResidualIdentifier]) -> list[ResidualIdentifier]:
    unique: dict[tuple[int, int, str, str], ResidualIdentifier] = {}
    for item in items:
        if item.start >= item.end:
            continue
        unique.setdefault((item.start, item.end, item.severity, item.reason), item)
    return sorted(unique.values(), key=lambda item: (item.start, item.end, item.severity))
