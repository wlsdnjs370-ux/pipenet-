from __future__ import annotations

import base64
import json
import subprocess
import tempfile
import time
import urllib.request
from pathlib import Path

import websocket


ROOT = Path(__file__).resolve().parents[1]
OUT_DIR = ROOT / "docs" / "generated_reports" / "server_screenshots"
OUT_DIR.mkdir(parents=True, exist_ok=True)

CHROME = Path(r"C:\Program Files\Google\Chrome\Application\chrome.exe")
URL = "http://127.0.0.1:5050/"
REPORT = ROOT / "data" / "uploads" / "201_3.docx"
SDF = ROOT / "data" / "uploads" / "1-1._201_3F___-RV03_NEW.sdf"


class CDP:
    def __init__(self, ws_url: str):
        self.ws = websocket.create_connection(ws_url, timeout=60)
        self.next_id = 1

    def call(self, method: str, params: dict | None = None):
        msg_id = self.next_id
        self.next_id += 1
        self.ws.send(json.dumps({"id": msg_id, "method": method, "params": params or {}}))
        while True:
            msg = json.loads(self.ws.recv())
            if msg.get("id") == msg_id:
                if "error" in msg:
                    raise RuntimeError(f"{method}: {msg['error']}")
                return msg.get("result", {})

    def close(self):
        self.ws.close()


def wait_for_debugger(port: int = 9223, timeout: float = 12.0) -> str:
    deadline = time.time() + timeout
    last = None
    while time.time() < deadline:
        try:
            with urllib.request.urlopen(f"http://127.0.0.1:{port}/json", timeout=1) as resp:
                pages = json.loads(resp.read().decode("utf-8"))
                if pages:
                    return pages[0]["webSocketDebuggerUrl"]
        except Exception as exc:
            last = exc
        time.sleep(0.25)
    raise RuntimeError(f"Chrome debugger did not start: {last}")


def wait(cdp: CDP, expr: str, timeout: float = 15.0):
    deadline = time.time() + timeout
    while time.time() < deadline:
        result = cdp.call("Runtime.evaluate", {"expression": expr, "returnByValue": True})
        if result.get("result", {}).get("value"):
            return True
        time.sleep(0.4)
    return False


def screenshot(cdp: CDP, name: str):
    result = cdp.call("Page.captureScreenshot", {"format": "png", "captureBeyondViewport": False})
    data = base64.b64decode(result["data"])
    path = OUT_DIR / f"{name}.png"
    path.write_bytes(data)
    print(path, path.stat().st_size)


def query_node(cdp: CDP, selector: str) -> int:
    root = cdp.call("DOM.getDocument", {"depth": 1})["root"]["nodeId"]
    return cdp.call("DOM.querySelector", {"nodeId": root, "selector": selector})["nodeId"]


def click(cdp: CDP, selector: str):
    cdp.call("Runtime.evaluate", {"expression": f"document.querySelector({selector!r})?.click()"})
    time.sleep(0.8)


def main():
    user_data = tempfile.mkdtemp(prefix="pipenet_chrome_")
    proc = subprocess.Popen(
        [
            str(CHROME),
            "--headless=new",
            "--disable-gpu",
            "--hide-scrollbars",
            "--remote-debugging-port=9223",
            "--remote-allow-origins=*",
            "--window-size=1440,960",
            f"--user-data-dir={user_data}",
            URL,
        ],
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
    )
    try:
        ws_url = wait_for_debugger()
        cdp = CDP(ws_url)
        try:
            cdp.call("Page.enable")
            cdp.call("DOM.enable")
            cdp.call("Runtime.enable")
            cdp.call("Page.navigate", {"url": URL})
            wait(cdp, "document.readyState === 'complete'", 10)
            time.sleep(1)
            screenshot(cdp, "01_main_upload")

            if REPORT.exists() and SDF.exists():
                cdp.call("DOM.setFileInputFiles", {"nodeId": query_node(cdp, "#report-file"), "files": [str(REPORT)]})
                cdp.call("DOM.setFileInputFiles", {"nodeId": query_node(cdp, "#sdf-file"), "files": [str(SDF)]})
                click(cdp, "#upload-form button[type='submit']")
                wait(cdp, "document.querySelector('#status') && !document.querySelector('#status').textContent.includes('중')", 25)
                time.sleep(2)
                screenshot(cdp, "02_after_validation")

            panels = [
                ("workspace-panel", "03_validation_network"),
                ("insight-panel", "04_optimization_guide"),
                ("tables-panel", "05_result_table"),
                ("stats-panel", "06_statistics"),
                ("cad-compare-panel", "07_cad_compare"),
            ]
            for panel, name in panels:
                cdp.call(
                    "Runtime.evaluate",
                    {
                        "expression": f"""
                        Array.from(document.querySelectorAll('.menu-panel')).forEach(p => p.classList.add('hidden'));
                        document.getElementById({panel!r})?.classList.remove('hidden');
                        Array.from(document.querySelectorAll('.menu-btn')).forEach(b => b.classList.toggle('active', b.dataset.panel === {panel!r}));
                        window.scrollTo(0,0);
                        """
                    },
                )
                time.sleep(1.2)
                screenshot(cdp, name)
        finally:
            cdp.close()
    finally:
        proc.terminate()
        try:
            proc.wait(timeout=5)
        except subprocess.TimeoutExpired:
            proc.kill()


if __name__ == "__main__":
    main()
