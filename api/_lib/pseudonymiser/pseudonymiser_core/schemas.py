from __future__ import annotations

from dataclasses import dataclass, field


@dataclass
class PseudonymiseRequest:
    text: str


@dataclass
class FindingResponse:
    entity_type: str
    start: int
    end: int
    score: float
    replacement: str


@dataclass
class EntityCount:
    entity_type: str
    count: int


@dataclass
class ResidualSpan:
    start: int
    end: int
    severity: str
    category: str
    reason: str


@dataclass
class PseudonymisationFindings:
    entities_detected: list[str]
    risks_preserved: list[str]
    clinical_details_preserved: list[str]
    warnings: list[str]
    riskLevel: str = "LOW"
    safeToSend: bool = True
    directIdentifiers: list[str] = field(default_factory=list)
    indirectIdentifiers: list[str] = field(default_factory=list)
    residualSpans: list[ResidualSpan] = field(default_factory=list)
    reason: str = "No unresolved direct identifiers found."


@dataclass
class PseudonymisationResponse:
    pseudonymised_text: str
    mapping: dict[str, str]
    findings: PseudonymisationFindings
    entities: list[FindingResponse]
    counts: list[EntityCount]


@dataclass
class SafetyCheckRequest:
    text: str


@dataclass
class SafetyCheckResponse:
    warnings: list[str]
    riskLevel: str = "LOW"
    safeToSend: bool = True
    directIdentifiers: list[str] = field(default_factory=list)
    indirectIdentifiers: list[str] = field(default_factory=list)
    reason: str = "No unresolved direct identifiers found."
