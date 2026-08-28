# -*- coding: utf-8 -*-
"""[F-10b] `/edit/worst` 의 몸통을 두 라우트가 나눠 쓰도록 뽑아낸다.

원클릭(anchor-click)은 「두 픽 + 최불리」를 한 잡에서 한다. 최불리 계산을
베껴 쓰면 두 길이 언젠가 갈린다 — 같은 함수를 쓰게 만든다.

뽑아낸 함수의 반환 규약은 «(요약, 실패응답)» 이다. 실패 반환이 열 곳이라
손으로 고치면 한 곳은 틀리므로 프로그램이 한다.

    python scripts/_refactor_worst_extract.py [--check]
"""
from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
SRC = ROOT / "routes" / "module_f" / "api_edit.py"
HEAD = "    def _unused_worst_body(sess, body):\n"


def main() -> int:
    sys.stdout.reconfigure(errors="replace")
    ap = argparse.ArgumentParser()
    ap.add_argument("--check", action="store_true")
    a = ap.parse_args()

    text = SRC.read_text(encoding="utf-8").replace("\r\n", "\n")
    i = text.index(HEAD)
    # 함수 끝 = 다음 라우트 데코레이터
    j = text.index("\n    @app.", i)
    body = text[i:j]

    n_fail = len(re.findall(r"^(\s+)return _fail\(", body, re.M))
    n_json = len(re.findall(r"^(\s+)return jsonify\(", body, re.M))
    body2 = re.sub(r"^(\s+)return _fail\(", r"\1return None, _fail(", body, flags=re.M)
    body2 = re.sub(r"^(\s+)return jsonify\(", r"\1return None, jsonify(",
                   body2, flags=re.M)
    body2 = body2.replace(HEAD, DOC)

    print(f"■ 실패 반환 {n_fail + n_json}곳을 «(None, 실패응답)» 으로")
    print(f"    _fail {n_fail} · jsonify {n_json}")
    if a.check:
        print("  --check 라 쓰지 않았다.")
        return 0
    SRC.write_text(text[:i] + body2 + text[j:], encoding="utf-8", newline="\n")
    print(f"  썼다 — {SRC.name}")
    return 0


DOC = '''    def _compute_worst(sess, body):
        """[Remote 30] 최불리 K 헤드와 경로 — «두 라우트가 나눠 쓰는» 몸통.

        `/edit/worst`(사람이 K·영역을 정하고 누름)와 `/edit/anchor-click`
        (알람밸브 원클릭이 뒤이어 자동 실행)이 **같은 계산**을 써야 한다.
        베껴 두면 언젠가 한쪽만 고쳐지고, 그 어긋남은 산출로만 드러난다.

        반환: (요약 dict, None) 또는 (None, 실패응답). 실패응답은 그대로
        돌려주면 되는 Flask 응답이다 — 상태 코드가 자리마다 다르므로
        문장만 넘기지 않는다.
        """
'''


if __name__ == "__main__":
    sys.exit(main())
