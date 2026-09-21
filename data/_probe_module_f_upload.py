# -*- coding: utf-8 -*-
"""모듈 F 업로드 «지연» 이 어디에 있나 — 구간별 실측.

올린 뒤 도면이 뜨기까지를 다섯 토막으로 갈라 각각을 잰다:

    ① 브라우저→서버 전송(여기서는 못 잰다 — 바이트 수로 갈음)
    ② waitress 가 몸통을 임시파일로 받는 비용
    ③ werkzeug 멀티파트 파싱
    ④ `_save_upload` 의 read()+write_bytes()
    ⑤ 같은 일을 스트리밍으로 했을 때

⑤ 가 ④ 보다 뚜렷이 싸면 서버가 범인이고, ① 이 압도적이면 압축이 답이다.
"""
from __future__ import annotations

import gzip
import io
import os
import shutil
import sys
import tempfile
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
TARGET = ROOT / "data" / "uploads" / "B1F 현장조사 소화설비 평면도.dxf"


def _mb(n: int) -> str:
    return f"{n / 1024 / 1024:,.1f} MB"


def main() -> None:
    src = Path(sys.argv[1]) if len(sys.argv) > 1 else TARGET
    if not src.is_file():
        raise SystemExit(f"파일 없음: {src}")
    size = src.stat().st_size
    print(f"대상 : {src.name}  ({_mb(size)})")

    tmpdir = Path(tempfile.mkdtemp(prefix="mf_upload_probe_"))
    try:
        # ④ 현행 _save_upload — 통째로 읽어 통째로 쓴다
        t0 = time.perf_counter()
        raw = src.read_bytes()
        t_read = time.perf_counter() - t0
        t0 = time.perf_counter()
        (tmpdir / "a.dxf").write_bytes(raw)
        t_write = time.perf_counter() - t0
        print(f"④ 현행  read_bytes {t_read:6.2f}s + write_bytes {t_write:6.2f}s "
              f"= {t_read + t_write:6.2f}s   (RAM 최고 {_mb(len(raw))})")
        del raw

        # ⑤ 스트리밍 복사
        t0 = time.perf_counter()
        with src.open("rb") as fi, (tmpdir / "b.dxf").open("wb") as fo:
            shutil.copyfileobj(fi, fo, 1024 * 1024)
        t_stream = time.perf_counter() - t0
        print(f"⑤ 스트림 copyfileobj                      = {t_stream:6.2f}s   "
              f"(RAM 1 MB)")

        # 압축 — 브라우저 CompressionStream 은 zlib level 6 근처다
        t0 = time.perf_counter()
        with src.open("rb") as fi:
            buf = io.BytesIO()
            with gzip.GzipFile(fileobj=buf, mode="wb", compresslevel=6) as gz:
                shutil.copyfileobj(fi, gz, 1024 * 1024)
        gz_size = buf.tell()
        t_gzip = time.perf_counter() - t0
        print(f"압축   gzip(6) {t_gzip:6.2f}s → {_mb(gz_size)} "
              f"({size / max(gz_size, 1):.1f}배 작아짐)")

        # 서버측 해제 — 통째 해제 vs 스트리밍 해제
        blob = buf.getvalue()
        t0 = time.perf_counter()
        out = gzip.decompress(blob)
        t_dec = time.perf_counter() - t0
        print(f"해제   gzip.decompress                    = {t_dec:6.2f}s   "
              f"(RAM {_mb(len(blob))}+{_mb(len(out))})")
        del out
        t0 = time.perf_counter()
        with gzip.GzipFile(fileobj=io.BytesIO(blob)) as gi, \
                (tmpdir / "c.dxf").open("wb") as fo:
            shutil.copyfileobj(gi, fo, 1024 * 1024)
        t_dec_s = time.perf_counter() - t0
        print(f"해제   스트리밍 해제→파일                 = {t_dec_s:6.2f}s   "
              f"(RAM {_mb(len(blob))}+1 MB)")

        # ③ werkzeug 멀티파트 파싱 — 실제 요청과 같은 몸통을 만들어 통과시킨다
        try:
            from werkzeug.datastructures import FileStorage
            from werkzeug.test import EnvironBuilder
            t0 = time.perf_counter()
            with src.open("rb") as fi:
                b = EnvironBuilder(
                    method="POST", path="/x",
                    data={"kind": "plan",
                          "dxf_file": FileStorage(fi, filename=src.name)})
                env = b.get_environ()
            t_body = time.perf_counter() - t0
            body_len = int(env.get("CONTENT_LENGTH") or 0)
            print(f"② 몸통 만들기(=waitress 수신 대역)        = {t_body:6.2f}s   "
                  f"({_mb(body_len)})")

            from werkzeug.wrappers import Request
            t0 = time.perf_counter()
            req = Request(env)
            fs = req.files.get("dxf_file")
            n = 0
            if fs is not None:
                fs.stream.seek(0, os.SEEK_END)
                n = fs.stream.tell()
                fs.stream.seek(0)
            t_parse = time.perf_counter() - t0
            print(f"③ werkzeug 멀티파트 파싱                 = {t_parse:6.2f}s   "
                  f"(꺼낸 파일 {_mb(n)})")
        except Exception as exc:  # noqa: BLE001
            print(f"③ 멀티파트 실측 실패: {type(exc).__name__}: {exc}")

        print()
        print("— 전송량 —")
        print(f"  현행(원본 그대로) {_mb(size)}")
        print(f"  압축 후          {_mb(gz_size)}")
        for mbps, label in ((100, "100 Mbps(무선/외부)"),
                            (1000, "1 Gbps(유선 LAN)")):
            a = size * 8 / (mbps * 1e6)
            c = gz_size * 8 / (mbps * 1e6)
            print(f"  {label:<18} 원본 {a:6.1f}s → 압축 {c:5.1f}s "
                  f"(+압축 {t_gzip:.1f}s = {c + t_gzip:5.1f}s)")
    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)


if __name__ == "__main__":
    main()
