from __future__ import annotations

from pydantic import BaseModel, Field


class PseudonymiseRequest(BaseModel):
    text: str = Field(..., min_length=1, max_length=250_000)


class FindingResponse(BaseModel):
    entity_type: str
    start: int
    end: int
    score: float
    replacement: str


class EntityCount(BaseModel):
    entity_type: str
    count: int


class ResidualSpan(BaseModel):
    start: int
    end: int
    severity: str
    category: str
    reason: str


class PseudonymisationFindings(BaseModel):
    entities_detected: list[str]
    risks_preserved: list[str]
    clinical_details_preserved: list[str]
    warnings: list[str]
    riskLevel: str = "LOW"
    safeToSend: bool = True
    directIdentifiers: list[str] = Field(default_factory=list)
    indirectIdentifiers: list[str] = Field(default_factory=list)
    residualSpans: list[ResidualSpan] = Field(default_factory=list)
    reason: str = "No unresolved direct identifiers found."


class PseudonymisationResponse(BaseModel):
    pseudonymised_text: str
    mapping: dict[str, str]
    findings: PseudonymisationFindings
    entities: list[FindingResponse]
    counts: list[EntityCount]


class SafetyCheckRequest(BaseModel):
    text: str = Field(..., min_length=1, max_length=250_000)


class SafetyCheckResponse(BaseModel):
    warnings: list[str]
    riskLevel: str = "LOW"
    safeToSend: bool = True
    directIdentifiers: list[str] = Field(default_factory=list)
    indirectIdentifiers: list[str] = Field(default_factory=list)
    reason: str = "No unresolved direct identifiers found."
