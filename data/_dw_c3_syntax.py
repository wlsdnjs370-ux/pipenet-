# -*- coding: utf-8 -*-
"""템플릿 <script> 본문만 떼어 node --check 로 문법을 본다. 문법일 뿐 실행은 아니다."""
import re
import subprocess
import sys
from pathlib import Path

src = Path("templates/design_workbench.html").read_text("utf-8")
body = re.search(r"<script>(.*)</script>", src, re.S).group(1)
out = Path("data/_dw_c3_script.js")
out.write_text(body, "utf-8")
sys.exit(subprocess.call(["node", "--check", str(out)], shell=True))
