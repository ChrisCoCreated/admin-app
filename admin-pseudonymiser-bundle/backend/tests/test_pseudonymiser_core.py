from __future__ import annotations

from pseudonymiser_core.anonymiser import CareNoteAnonymiser
from pseudonymiser_core.recognizers import (
    assess_data_disclosure_risk,
    is_valid_nhs_number,
    scan_for_residual_identifiers,
)


def test_validates_nhs_checksum() -> None:
    assert is_valid_nhs_number("943 476 5919")
    assert not is_valid_nhs_number("943 476 5918")


def test_pseudonymises_sample_note_with_stable_role_placeholders() -> None:
    note = (
        "Paul Jones feels safe at Mossbank. Claire is brilliant. "
        "Discussed Brian Smith and selling her properties."
    )

    result = CareNoteAnonymiser().pseudonymise(note)

    assert (
        result.pseudonymised_text
        == "[CLIENT_001] feels safe at [CARE_HOME_001]. [STAFF_001] is brilliant. "
        "Discussed [PROFESSIONAL_001] and selling her properties."
    )
    assert result.mapping == {
        "[CLIENT_001]": "Paulette Crawley",
        "[CARE_HOME_001]": "Lindau",
        "[STAFF_001]": "Carrie",
        "[PROFESSIONAL_001]": "Martin Tyrell",
    }
    assert result.findings.entities_detected == [
        "[CLIENT_001]",
        "[CARE_HOME_001]",
        "[STAFF_001]",
        "[PROFESSIONAL_001]",
    ]
    assert result.findings.warnings == []


def test_restoration_replaces_all_placeholders_with_original_values() -> None:
    note = (
        "Paulette Crawley feels safe at Lindau. Carrie is brilliant. "
        "Discussed Martin Tyrell and selling her properties."
    )
    result = CareNoteAnonymiser().pseudonymise(note)

    restored = result.pseudonymised_text
    for placeholder, original in result.mapping.items():
        restored = restored.replace(placeholder, original)

    assert restored == note


def test_partial_client_aliases_map_to_same_placeholder() -> None:
    note = (
        "Paulette Crawley said Paulette was tired. Mrs Crawley later asked for reassurance."
    )

    result = CareNoteAnonymiser().pseudonymise(note)

    assert result.pseudonymised_text.count("[CLIENT_001]") == 3
    assert "Paulette" not in result.pseudonymised_text
    assert "Crawley" not in result.pseudonymised_text


def test_labelled_care_note_fields_do_not_pseudonymise_headings() -> None:
    note = (
        "**Client Name:** Jane Smith  \n"
        "**DOB:** 14/02/1941  \n"
        "**Date:** 14/05/2026  \n"
        "**Carer:** A. Brown  \n"
        "**Visit Type:** Morning, lunch, tea, and bedtime support\n\n"
        "**07:30 Morning Visit**  \n"
        "Arrived at property and knocked loudly before \n"
    )

    result = CareNoteAnonymiser().pseudonymise(note)

    assert "**Client Name:** [CLIENT_001]" in result.pseudonymised_text
    assert "**Carer:** [STAFF_001]" in result.pseudonymised_text
    assert "**Visit Type:** Morning, lunch, tea, and bedtime support" in result.pseudonymised_text
    assert "**07:30 Morning Visit**" in result.pseudonymised_text
    assert result.mapping["[CLIENT_001]"] == "Jane Smith"
    assert result.mapping["[STAFF_001]"] == "A. Brown"
    assert "Name" not in result.mapping.values()
    assert "Visit Type" not in result.mapping.values()
    assert "Morning Visit" not in result.mapping.values()


def test_friend_can_be_pseudonymised_as_own_identifier_type() -> None:
    note = "Friend Sarah Brown visited after lunch. Neighbour Alison called later."

    result = CareNoteAnonymiser().pseudonymise(note)

    assert "[FRIEND_001]" in result.pseudonymised_text
    assert "[FRIEND_002]" in result.pseudonymised_text
    assert result.mapping["[FRIEND_001]"] == "Sarah Brown"
    assert result.mapping["[FRIEND_002]"] == "Alison"


def test_preserves_clinical_and_risk_meaning_in_messy_note() -> None:
    note = (
        "Paulette Crawley has dementa, diabetes and high falls risk. "
        "Daughter lives overseas and says 'I feel anxious at night'. "
        "Needs help with mobility and meds."
    )

    result = CareNoteAnonymiser().pseudonymise(note)
    output = result.pseudonymised_text.lower()

    assert "dementa" in output
    assert "diabetes" in output
    assert "falls risk" in output
    assert "daughter lives overseas" in output
    assert "mobility" in output
    assert "anxious at night" in output
    assert "falls" in result.findings.risks_preserved
    assert "diabetes" in result.findings.clinical_details_preserved


def test_safety_check_flags_unresolved_identifiers_before_llm_send() -> None:
    warnings = scan_for_residual_identifiers(
        "Client is [CLIENT_001]. Please ring 07911 123456 or email jane.bloggs@example.org."
    )

    assert any("Likely phone number remains" in warning for warning in warnings)
    assert any("Likely email remains" in warning for warning in warnings)


def test_known_client_first_name_remaining_forces_high_data_risk() -> None:
    risk = assess_data_disclosure_risk(
        "[CLIENT_001] feels safe. Eleanor is happy at [CARE_HOME_001]. "
        "Daughter lives in Canada.",
        {"[CLIENT_001]": "Eleanor Marsh"},
    )

    assert risk["riskLevel"] == "HIGH"
    assert risk["safeToSend"] is False
    assert "Eleanor" in risk["directIdentifiers"]
    assert "Daughter lives in Canada" in risk["indirectIdentifiers"]


def test_known_client_name_parts_match_case_punctuation_and_possessives() -> None:
    mapping = {"[CLIENT_001]": "Eleanor Marsh"}

    for text, expected in (
        ("[CLIENT_001] said Eleanor Marsh wanted tea.", "Eleanor Marsh"),
        ("[CLIENT_001] said Eleanor's room was quiet.", "Eleanor"),
        ("[CLIENT_001] said Eleanor’s room was quiet.", "Eleanor"),
        ("[CLIENT_001] said marsh was written on the folder.", "Marsh"),
        ("[CLIENT_001] said (Marsh) was written on the folder.", "Marsh"),
    ):
        risk = assess_data_disclosure_risk(text, mapping)

        assert risk["riskLevel"] == "HIGH"
        assert risk["safeToSend"] is False
        assert expected in risk["directIdentifiers"]


def test_direct_identifier_risk_includes_spans_for_common_patterns() -> None:
    text = (
        "Please ring 07911 123456, email jane.bloggs@example.org, "
        "or visit 12 Acacia Road, SW1A 1AA."
    )

    risk = assess_data_disclosure_risk(text, {})
    spans = risk["residualSpans"]

    assert risk["riskLevel"] == "HIGH"
    assert _span_texts(text, spans) >= {
        "07911 123456",
        "jane.bloggs@example.org",
        "12 Acacia Road, SW1A 1AA",
        "SW1A 1AA",
    }
    assert all(span["severity"] == "direct" for span in spans)


def test_known_identifier_risk_includes_spans_for_names_and_possessives() -> None:
    text = "[CLIENT_001] said Eleanor's room was quiet. Marsh was written on the folder."

    risk = assess_data_disclosure_risk(text, {"[CLIENT_001]": "Eleanor Marsh"})
    spans = risk["residualSpans"]

    assert risk["riskLevel"] == "HIGH"
    assert "Eleanor" in risk["directIdentifiers"]
    assert "Marsh" in risk["directIdentifiers"]
    assert _span_texts(text, spans) >= {"Eleanor's", "Marsh"}


def test_several_indirect_identifiers_without_direct_identifiers_are_medium() -> None:
    text = (
        "[CLIENT_001] has dementia, diabetes, medication support and falls risk. "
        "Daughter lives in Canada. Moved from Jamaica in the 1960s. "
        "Discussed selling her properties. The service has 31 beds."
    )
    risk = assess_data_disclosure_risk(text, {})

    assert risk["riskLevel"] == "MEDIUM"
    assert risk["safeToSend"] is True
    assert risk["directIdentifiers"] == []
    assert len(risk["indirectIdentifiers"]) >= 3
    assert "Daughter lives in Canada" in risk["indirectIdentifiers"]
    assert _span_texts(text, risk["residualSpans"]) >= {
        "Daughter lives in Canada",
        "Moved from Jamaica",
        "selling",
        "properties",
        "31 beds",
    }
    assert all(span["severity"] == "indirect" for span in risk["residualSpans"])


def _span_texts(text: str, spans: list[dict[str, object]]) -> set[str]:
    return {text[int(span["start"]) : int(span["end"])] for span in spans}
