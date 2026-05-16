from __future__ import annotations

import sys
from http.server import BaseHTTPRequestHandler
from pathlib import Path


CORE_ROOT = Path(__file__).resolve().parents[1] / "_lib" / "pseudonymiser"
if str(CORE_ROOT) not in sys.path:
    sys.path.insert(0, str(CORE_ROOT))

from serverless import handle_health, send_json


class handler(BaseHTTPRequestHandler):
    def do_GET(self) -> None:
        handle_health(self)

    def do_POST(self) -> None:
        send_json(self, 405, {"error": "Method not allowed."})
