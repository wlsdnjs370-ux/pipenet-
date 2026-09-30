"""Read-only snapshot of one known Module F task via its normal local login."""
from __future__ import annotations

import argparse
import html
import json
from pathlib import Path
import re

import requests


def main() -> None:
    """Save inspectable API responses without changing drawing/selection state."""
    parser = argparse.ArgumentParser()
    parser.add_argument("sid")
    parser.add_argument("--out", type=Path, required=True)
    args = parser.parse_args()
    session = requests.Session()
    base = "http://127.0.0.1:5051"
    login = session.get(base + "/login", timeout=15)
    login.raise_for_status()
    hint = re.search(r'<span class="pw-value">([^<]+)<', login.text)
    if not hint:
        raise RuntimeError("Local login hint unavailable; ask the user to log in.")
    session.post(base + "/login", data={"password": html.unescape(hint.group(1)),
                                       "next": "/module-f"}, timeout=15).raise_for_status()
    endpoints = {
        "merge_state": ("merge/state", {}),
        "merge_preview": ("merge/preview", {"iso": "0"}),
        "merge_editor": ("network-editor", {"scope": "merge"}),
        "design_editor": ("network-editor", {"scope": "design"}),
        "design_preview": ("design/preview", {}),
        "sizing_state": ("merge/sizing/state", {}),
        "slots": ("slot/state", {}),
        "sub_pipes": ("sub/pipes", {}),
    }
    result = {}
    for name, (endpoint, params) in endpoints.items():
        response = session.get(base + "/api/module-f/" + endpoint,
                               params={"sid": args.sid, **params}, timeout=30)
        response.raise_for_status()
        result[name] = response.json()
    args.out.parent.mkdir(parents=True, exist_ok=True)
    args.out.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    net = result["merge_editor"]
    print(json.dumps({"saved": str(args.out), "nodes": len(net["nodes"]),
                      "pipes": len(net["pipes"]), "history": net["history"],
                      "suspect_pipes": [p for p in net["pipes"] if p["label"] in ("r5", "r23", "P2", "P4")]},
                     ensure_ascii=False))


if __name__ == "__main__":
    main()
