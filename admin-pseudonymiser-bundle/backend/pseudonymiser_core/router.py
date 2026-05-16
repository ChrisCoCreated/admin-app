from __future__ import annotations

from fastapi import APIRouter

from .anonymiser import CareNoteAnonymiser
from .privacy import configure_private_logging
from .recognizers import assess_data_disclosure_risk
from .schemas import (
    PseudonymisationResponse,
    PseudonymiseRequest,
    SafetyCheckRequest,
    SafetyCheckResponse,
)

configure_private_logging()

router = APIRouter(tags=["pseudonymiser"])
anonymiser = CareNoteAnonymiser()


@router.get("/health")
async def pseudonymiser_health() -> dict[str, str]:
    return {"status": "ok"}


@router.post("/pseudonymise", response_model=PseudonymisationResponse)
async def pseudonymise(request: PseudonymiseRequest) -> PseudonymisationResponse:
    return anonymiser.pseudonymise(request.text)


@router.post("/safety-check", response_model=SafetyCheckResponse)
async def safety_check(request: SafetyCheckRequest) -> SafetyCheckResponse:
    risk = assess_data_disclosure_risk(request.text, {})
    return SafetyCheckResponse(
        warnings=[
            str(item)
            for item in risk["directIdentifiers"] + risk["indirectIdentifiers"]
        ],
        riskLevel=str(risk["riskLevel"]),
        safeToSend=bool(risk["safeToSend"]),
        directIdentifiers=list(risk["directIdentifiers"]),
        indirectIdentifiers=list(risk["indirectIdentifiers"]),
        reason=str(risk["reason"]),
    )
