from __future__ import annotations

from collections import Counter
from dataclasses import dataclass

from .recognizers import (
    CandidateFinding,
    assess_data_disclosure_risk,
    collect_preserved_details,
    detect_candidate_entities,
)
from .schemas import EntityCount, FindingResponse, PseudonymisationFindings, PseudonymisationResponse


PERSON_CATEGORIES = {"CLIENT", "STAFF", "RELATIVE", "FRIEND", "PROFESSIONAL"}


@dataclass
class ResolvedEntity:
    category: str
    placeholder: str
    original_value: str
    first_name: str | None
    surname: str | None


class CareNoteAnonymiser:
    def pseudonymise(self, text: str) -> PseudonymisationResponse:
        candidates = detect_candidate_entities(text)
        findings, mapping = self._resolve_candidates(candidates)
        pseudonymised_text = self._apply_replacements(text, findings)
        counts = [
            EntityCount(entity_type=entity_type, count=count)
            for entity_type, count in sorted(Counter(item.entity_type for item in findings).items())
        ]
        risks_preserved, clinical_details_preserved = collect_preserved_details(
            text, pseudonymised_text
        )
        disclosure_risk = assess_data_disclosure_risk(pseudonymised_text, mapping)
        warning_items = disclosure_risk["directIdentifiers"]
        if disclosure_risk["riskLevel"] != "LOW":
            warning_items = warning_items + disclosure_risk["indirectIdentifiers"]
        warnings = [str(item) for item in warning_items]

        return PseudonymisationResponse(
            pseudonymised_text=pseudonymised_text,
            mapping=mapping,
            findings=PseudonymisationFindings(
                entities_detected=list(mapping.keys()),
                risks_preserved=risks_preserved,
                clinical_details_preserved=clinical_details_preserved,
                warnings=warnings,
                riskLevel=str(disclosure_risk["riskLevel"]),
                safeToSend=bool(disclosure_risk["safeToSend"]),
                directIdentifiers=list(disclosure_risk["directIdentifiers"]),
                indirectIdentifiers=list(disclosure_risk["indirectIdentifiers"]),
                residualSpans=list(disclosure_risk["residualSpans"]),
                reason=str(disclosure_risk["reason"]),
            ),
            entities=findings,
            counts=counts,
        )

    def anonymise(
        self, text: str
    ) -> tuple[str, list[FindingResponse], list[EntityCount], dict[str, str], PseudonymisationFindings]:
        result = self.pseudonymise(text)
        return (
            result.pseudonymised_text,
            result.entities,
            result.counts,
            result.mapping,
            result.findings,
        )

    def _resolve_candidates(
        self, candidates: list[CandidateFinding]
    ) -> tuple[list[FindingResponse], dict[str, str]]:
        findings: list[FindingResponse] = []
        mapping: dict[str, str] = {}
        counters: dict[str, int] = {}
        entities: list[ResolvedEntity] = []

        for candidate in candidates:
            category = candidate.entity_type

            if category == "PERSON":
                entity = self._find_person_entity(candidate.text, entities)
                if entity is None:
                    category = self._infer_person_category(candidate, entities)
                    entity = self._create_entity(candidate.text, category, counters, entities)
                placeholder = entity.placeholder
            elif category in PERSON_CATEGORIES:
                entity = self._find_person_entity(candidate.text, entities, preferred_category=category)
                if entity is None:
                    entity = self._create_entity(candidate.text, category, counters, entities)
                placeholder = entity.placeholder
            else:
                placeholder = self._get_or_create_direct_placeholder(
                    candidate.text, category, counters, entities
                )

            mapping.setdefault(placeholder, candidate.text)
            findings.append(
                FindingResponse(
                    entity_type=placeholder[1:-5],
                    start=candidate.start,
                    end=candidate.end,
                    score=candidate.score,
                    replacement=placeholder,
                )
            )

        return findings, mapping

    def _find_person_entity(
        self,
        text: str,
        entities: list[ResolvedEntity],
        *,
        preferred_category: str | None = None,
    ) -> ResolvedEntity | None:
        first_name, surname = _split_person_name(text)

        exact_matches = [
            entity
            for entity in entities
            if entity.category in PERSON_CATEGORIES
            and (preferred_category is None or entity.category == preferred_category)
            and _names_match(first_name, surname, entity)
        ]
        if len(exact_matches) == 1:
            return exact_matches[0]

        if surname:
            surname_matches = [
                entity
                for entity in entities
                if entity.category in PERSON_CATEGORIES
                and (preferred_category is None or entity.category == preferred_category)
                and entity.surname == surname
            ]
            if len(surname_matches) == 1:
                return surname_matches[0]

        if first_name:
            first_matches = [
                entity
                for entity in entities
                if entity.category in PERSON_CATEGORIES
                and (preferred_category is None or entity.category == preferred_category)
                and entity.first_name == first_name
            ]
            if len(first_matches) == 1:
                return first_matches[0]

        return None

    def _infer_person_category(
        self, candidate: CandidateFinding, entities: list[ResolvedEntity]
    ) -> str:
        first_name, surname = _split_person_name(candidate.text)

        if surname and not any(entity.category == "CLIENT" for entity in entities):
            return "CLIENT"
        if surname:
            return "PROFESSIONAL"
        if first_name:
            return "STAFF"
        return "PROFESSIONAL"

    def _create_entity(
        self,
        original_text: str,
        category: str,
        counters: dict[str, int],
        entities: list[ResolvedEntity],
    ) -> ResolvedEntity:
        counters[category] = counters.get(category, 0) + 1
        first_name, surname = _split_person_name(original_text)
        entity = ResolvedEntity(
            category=category,
            placeholder=f"[{category}_{counters[category]:03d}]",
            original_value=original_text,
            first_name=first_name,
            surname=surname,
        )
        entities.append(entity)
        return entity

    def _get_or_create_direct_placeholder(
        self,
        value: str,
        category: str,
        counters: dict[str, int],
        entities: list[ResolvedEntity],
    ) -> str:
        normalised = _normalise_direct_value(value)
        for entity in entities:
            if entity.category == category and _normalise_direct_value(entity.original_value) == normalised:
                return entity.placeholder

        counters[category] = counters.get(category, 0) + 1
        entity = ResolvedEntity(
            category=category,
            placeholder=f"[{category}_{counters[category]:03d}]",
            original_value=value,
            first_name=None,
            surname=None,
        )
        entities.append(entity)
        return entity.placeholder

    @staticmethod
    def _apply_replacements(text: str, findings: list[FindingResponse]) -> str:
        output: list[str] = []
        cursor = 0

        for finding in findings:
            output.append(text[cursor : finding.start])
            output.append(finding.replacement)
            cursor = finding.end

        output.append(text[cursor:])
        return "".join(output)


def _split_person_name(value: str) -> tuple[str | None, str | None]:
    stripped = value.strip().replace(".", "")
    had_title = False
    for title in ("Mr ", "Mrs ", "Ms ", "Miss ", "Dr "):
        if stripped.startswith(title):
            stripped = stripped.removeprefix(title)
            had_title = True
            break
    parts = [part for part in stripped.split() if part]
    if not parts:
        return None, None
    if len(parts) == 1:
        if had_title:
            return None, parts[0].lower()
        return parts[0].lower(), None
    return parts[0].lower(), parts[-1].lower()


def _names_match(first_name: str | None, surname: str | None, entity: ResolvedEntity) -> bool:
    if first_name and surname:
        return entity.first_name == first_name and entity.surname == surname
    if surname:
        return entity.surname == surname
    if first_name:
        return entity.first_name == first_name
    return False


def _normalise_direct_value(value: str) -> str:
    return " ".join(value.strip().lower().split())
