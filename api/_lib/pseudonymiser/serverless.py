from __future__ import annotations

import json
import sys
import urllib.error
import urllib.request
from dataclasses import asdict, is_dataclass
from pathlib import Path
from typing import Any


CORE_ROOT = Path(__file__).resolve().parent
if str(CORE_ROOT) not in sys.path:
    sys.path.insert(0, str(CORE_ROOT))

from pseudonymiser_core.anonymiser import CareNoteAnonymiser
from pseudonymiser_core.recognizers import assess_data_disclosure_risk


MAX_TEXT_LENGTH = 250_000
CLIENT_DATA_ROLES = {
    "admin",
    "care_manager",
    "operations",
    "clients_only",
    "hr_clients",
    "time_clients",
    "time_hr_clients",
}


def to_plain(value: Any) -> Any:
    if is_dataclass(value):
        return to_plain(asdict(value))
    if isinstance(value, dict):
        return {str(key): to_plain(item) for key, item in value.items()}
    if isinstance(value, list):
        return [to_plain(item) for item in value]
    return value


def send_json(handler: Any, status_code: int, payload: dict[str, Any]) -> None:
    body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
    handler.send_response(status_code)
    handler.send_header("Content-Type", "application/json; charset=utf-8")
    handler.send_header("Cache-Control", "no-store")
    handler.send_header("Content-Length", str(len(body)))
    handler.end_headers()
    handler.wfile.write(body)


def read_json_body(handler: Any) -> dict[str, Any]:
    length = int(handler.headers.get("content-length") or "0")
    if length <= 0:
        return {}
    if length > 1024 * 1024:
        raise ValueError("Request body too large.")
    raw = handler.rfile.read(length).decode("utf-8").strip()
    if not raw:
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


def auth_me_url(handler: Any) -> str:
    host = handler.headers.get("x-forwarded-host") or handler.headers.get("host")
    proto = handler.headers.get("x-forwarded-proto") or "https"
    if not host:
        raise PermissionError("Missing host header.")
    return f"{proto}://{host}/api/auth/me"


def role_can_access_client_data(role: str) -> bool:
    normalised = (role or "").strip().lower()
    if normalised in CLIENT_DATA_ROLES:
        return True
    if normalised.startswith("pages:"):
        pages = {page.strip().lower() for page in normalised.removeprefix("pages:").split(",")}
        return "clientdata" in pages or "clients" in pages
    return False


def require_client_data_auth(handler: Any) -> None:
    auth_header = handler.headers.get("authorization") or ""
    if not auth_header.lower().startswith("bearer "):
        raise PermissionError("Missing bearer token.")

    request = urllib.request.Request(
        auth_me_url(handler),
        headers={
            "Accept": "application/json",
            "Authorization": auth_header,
        },
        method="GET",
    )

    try:
        with urllib.request.urlopen(request, timeout=10) as response:
            profile = json.loads(response.read().decode("utf-8") or "{}")
    except urllib.error.HTTPError as error:
        raise PermissionError("Unauthorized.") from error

    if not role_can_access_client_data(str(profile.get("role") or "")):
        raise PermissionError("Forbidden.")


def handle_health(handler: Any) -> None:
    send_json(handler, 200, {"status": "ok"})


def handle_post_action(handler: Any, action: str) -> None:
    try:
        require_client_data_auth(handler)
        payload = read_json_body(handler)
        text = get_text(payload)
    except PermissionError as error:
        message = str(error) or "Unauthorized."
        status_code = 403 if message == "Forbidden." else 401
        send_json(handler, status_code, {"error": message})
        return
    except ValueError as error:
        send_json(handler, 400, {"error": str(error)})
        return
    except json.JSONDecodeError:
        send_json(handler, 400, {"error": "Invalid JSON body."})
        return

    try:
        if action == "pseudonymise":
            send_json(handler, 200, to_plain(CareNoteAnonymiser().pseudonymise(text)))
            return
        if action == "safety-check":
            risk = assess_data_disclosure_risk(text, {})
            send_json(
                handler,
                200,
                {
                    "warnings": [
                        str(item)
                        for item in risk["directIdentifiers"] + risk["indirectIdentifiers"]
                    ],
                    "riskLevel": str(risk["riskLevel"]),
                    "safeToSend": bool(risk["safeToSend"]),
                    "directIdentifiers": list(risk["directIdentifiers"]),
                    "indirectIdentifiers": list(risk["indirectIdentifiers"]),
                    "reason": str(risk["reason"]),
                },
            )
            return
        send_json(handler, 404, {"error": "Not Found"})
    except Exception:
        send_json(handler, 500, {"error": "Pseudonymiser failed."})
