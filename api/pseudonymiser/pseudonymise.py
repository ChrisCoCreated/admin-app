from __future__ import annotations

import sys
from http.server import BaseHTTPRequestHandler
from pathlib import Path


CORE_ROOT = Path(__file__).resolve().parents[1] / "_lib" / "pseudonymiser"
if str(CORE_ROOT) not in sys.path:
    sys.path.insert(0, str(CORE_ROOT))

from serverless import handle_post_action, send_json


class handler(BaseHTTPRequestHandler):
    def do_POST(self) -> None:
        handle_post_action(self, "pseudonymise")

    def do_GET(self) -> None:
        send_json(self, 405, {"error": "Method not allowed."})
