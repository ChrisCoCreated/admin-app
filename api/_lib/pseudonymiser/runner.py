from __future__ import annotations

import json
import sys
from dataclasses import asdict, is_dataclass
from typing import Any

from pseudonymiser_core.anonymiser import CareNoteAnonymiser
from pseudonymiser_core.recognizers import assess_data_disclosure_risk


MAX_TEXT_LENGTH = 250_000


def to_plain(value: Any) -> Any:
    if is_dataclass(value):
        return to_plain(asdict(value))
    if isinstance(value, dict):
        return {str(key): to_plain(item) for key, item in value.items()}
    if isinstance(value, list):
        return [to_plain(item) for item in value]
    return value


def read_payload() -> dict[str, Any]:
    raw = sys.stdin.read()
    if not raw.strip():
        return {}
    payload = json.loads(raw)
    if not isinstance(payload, dict):
        raise ValueError("Request body must be a JSON object.")
    return payload


def get_text(payload: dict[str, Any]) -> str:
    text = payload.get("text")
    if not isinstance(text, str) or not text.strip():
        raise ValueError("Text is required.")
    if len(text) > MAX_TEXT_LENGTH:
        raise ValueError("Text is too long.")
    return text


def run(action: str) -> dict[str, Any]:
    if action == "health":
        return {"status": "ok"}

    payload = read_payload()
    text = get_text(payload)

    if action == "pseudonymise":
        return to_plain(CareNoteAnonymiser().pseudonymise(text))

    if action == "safety-check":
        risk = assess_data_disclosure_risk(text, {})
        return {
            "warnings": [
                str(item)
                for item in risk["directIdentifiers"] + risk["indirectIdentifiers"]
            ],
            "riskLevel": str(risk["riskLevel"]),
            "safeToSend": bool(risk["safeToSend"]),
            "directIdentifiers": list(risk["directIdentifiers"]),
            "indirectIdentifiers": list(risk["indirectIdentifiers"]),
            "reason": str(risk["reason"]),
        }

    raise ValueError("Unknown pseudonymiser action.")


def main() -> int:
    action = sys.argv[1] if len(sys.argv) > 1 else ""
    try:
        print(json.dumps(run(action), ensure_ascii=False))
        return 0
    except Exception as error:
        print(json.dumps({"error": str(error) or "Pseudonymiser failed."}), file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
