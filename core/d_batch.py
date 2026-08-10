# -*- coding: utf-8 -*-
"""모듈 D — D6: 폴더 하나를 넣으면 세트마다 PDF 한 벌이 나온다.

D5 의 계산서와 D3 의 ISO 도면을 세트 단위로 묶어 낸다. 여기서 새로 계산하는 값은
하나도 없다 — 어느 파일을 어느 조각에 넘길지 정하고, 실패를 격리하고, 무엇이
되고 무엇이 안 됐는지 적는 것이 전부다.

**격리** — 지시서 D6 수용 기준은 "1개가 실패해도 나머지 29개가 완료된다" 이다.
세트끼리는 물론이고 한 세트 안의 조각끼리도 갈라 둔다. 계산서가 안 나와도 도면
두 장은 나오고, 나온 것만 묶는다. 못 나온 것은 리포트에 사유와 함께 남는다.

**프리셋** — 압력본은 스프링클러용(``Pipe pressure difference``)과 옥내소화전용
(``Pipe bore``)이 갈리는데, 어느 쪽인지 파일에는 적혀 있지 않다. 참조 코퍼스
260 세트를 재 보면 SDF 가 들고 있는 표시 스킴은 마지막 편집 상태일 뿐이라
설비 종류와 무관하고(옥내소화전 안에서만 ``Pipe type`` 80 · ``None`` 55 ·
``Pipe vol. flow`` 22 로 흩어진다), 노즐 유무로도 갈리지 않는다(260 세트 전량이
노즐을 가진다). 그래서 추정하지 않고 호출자가 고르게 둔다 (지시서 7-4).
"""
from __future__ import annotations

import argparse
import sys
import time
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Callable, Sequence

from pypdf import PdfWriter

from core.d_display_model import load_display_model
from core.d_iso_renderer import PRESETS, render_iso
from core.d_report_typeset import render_report
from core.d_result_binder import bind_results

# 지시서 D6 — "구간별 [계산서 + ISO 유량본 + ISO 압력본] 생성 후 결합".
REPORT_PART = "계산서"
DEFAULT_PRESETS: tuple[str, ...] = ("유량본", "압력본")


# ── 세트 인식 ───────────────────────────────────────────────────────────────


@dataclass(frozen=True)
class FileSet:
    """같은 자리에 같은 이름으로 놓인 입력 한 벌."""
    stem: str
    sdf: Path
    xml: Path | None = None
    slf: Path | None = None

    @property
    def missing(self) -> tuple[str, ...]:
        return tuple(ext for ext, path in ((".xml", self.xml), (".slf", self.slf))
                     if path is None)


def _sibling(sdf: Path, ext: str) -> Path | None:
    """대소문자를 가리지 않는 짝 찾기. 없으면 None."""
    for candidate in (sdf.with_suffix(ext), sdf.with_suffix(ext.upper())):
        if candidate.exists():
            return candidate
    return None


def discover(root: str | Path) -> list[FileSet]:
    """폴더(또는 SDF 한 개) → 파일 세트 목록. SDF 가 있는 자리만 세트로 본다."""
    base = Path(root)
    if base.is_file():
        found = [base] if base.suffix.lower() == ".sdf" else []
    else:
        found = sorted(base.rglob("*.sdf"))
    # 오피스 임시본(~$…)은 입력이 아니다.
    return [FileSet(stem=sdf.stem, sdf=sdf, xml=_sibling(sdf, ".xml"),
                    slf=_sibling(sdf, ".slf"))
            for sdf in found if not sdf.name.startswith("~$")]


# ── 결과 ────────────────────────────────────────────────────────────────────


@dataclass
class PartResult:
    name: str
    path: Path | None = None
    pages: int = 0
    error: str = ""
    warnings: tuple[str, ...] = ()
    seconds: float = 0.0

    @property
    def ok(self) -> bool:
        return not self.error and self.path is not None


@dataclass
class SetResult:
    file_set: FileSet
    parts: list[PartResult] = field(default_factory=list)
    combined: Path | None = None
    error: str = ""
    notes: list[str] = field(default_factory=list)
    seconds: float = 0.0

    @property
    def done(self) -> list[PartResult]:
        return [p for p in self.parts if p.ok]

    @property
    def broken(self) -> list[PartResult]:
        return [p for p in self.parts if not p.ok]

    @property
    def ok(self) -> bool:
        return not self.error and bool(self.parts) and len(self.done) == len(self.parts)

    @property
    def partial(self) -> bool:
        return not self.error and bool(self.done) and len(self.done) < len(self.parts)

    def line(self) -> str:
        if self.error:
            state = f"실패 — {self.error}"
        elif self.partial:
            state = f"부분 {len(self.done)}/{len(self.parts)}"
        elif self.ok:
            state = f"성공 {len(self.parts)} 조각"
        else:
            state = "실패 — 나온 조각이 없다"
        pages = sum(p.pages for p in self.done)
        return f"{self.file_set.sdf.name:<40} {state} · {pages}쪽 · {self.seconds:.1f}s"


@dataclass
class BatchReport:
    root: Path
    out_dir: Path
    results: list[SetResult] = field(default_factory=list)
    seconds: float = 0.0

    @property
    def succeeded(self) -> list[SetResult]:
        return [r for r in self.results if r.ok]

    @property
    def partial(self) -> list[SetResult]:
        return [r for r in self.results if r.partial]

    @property
    def failed(self) -> list[SetResult]:
        return [r for r in self.results if not r.ok and not r.partial]

    def text(self) -> str:
        lines = [f"입력 {self.root}", f"출력 {self.out_dir}",
                 f"세트 {len(self.results)} · 성공 {len(self.succeeded)} · "
                 f"부분 {len(self.partial)} · 실패 {len(self.failed)} · "
                 f"{self.seconds:.1f}s", ""]
        for result in self.results:
            lines.append(result.line())
            for note in result.notes:
                lines.append(f"    · {note}")
            for part in result.broken:
                lines.append(f"    × {part.name} — {part.error}")
        return "\n".join(lines)


# ── 한 세트 ─────────────────────────────────────────────────────────────────


def _reason(exc: BaseException) -> str:
    return f"{type(exc).__name__}: {exc}".replace("\n", " ")[:200]


def _attempt(name: str, out: Path,
             action: Callable[[Path], tuple[int, Sequence[str]]]) -> PartResult:
    """조각 하나. 여기서 터진 예외는 이 조각 밖으로 나가지 않는다."""
    started = time.perf_counter()
    try:
        out.parent.mkdir(parents=True, exist_ok=True)
        pages, warnings = action(out)
    except Exception as exc:  # noqa: BLE001 — 세트 격리가 목적이다
        return PartResult(name, error=_reason(exc),
                          seconds=time.perf_counter() - started)
    return PartResult(name, path=out, pages=pages, warnings=tuple(warnings),
                      seconds=time.perf_counter() - started)


def _combine(parts: Sequence[PartResult], out: Path) -> None:
    """나온 조각만 한 권으로. 조각 이름이 그대로 책갈피가 된다."""
    writer = PdfWriter()
    try:
        for part in parts:
            writer.append(str(part.path), outline_item=part.name)
        out.parent.mkdir(parents=True, exist_ok=True)
        with out.open("wb") as fh:
            writer.write(fh)
    finally:
        writer.close()


def process_set(file_set: FileSet, dest: str | Path, *,
                presets: Sequence[str] = DEFAULT_PRESETS,
                combine: bool = True,
                generated: datetime | None = None) -> SetResult:
    """세트 하나 → ``dest/<stem>/`` 에 조각, ``dest/<stem>.pdf`` 에 합본."""
    started = time.perf_counter()
    result = SetResult(file_set=file_set)
    dest = Path(dest)
    parts_dir = dest / file_set.stem

    unknown = [name for name in presets if name not in PRESETS]
    if unknown:
        result.error = f"모르는 프리셋 {unknown} — {sorted(PRESETS)}"
        return result

    try:
        model = load_display_model(file_set.sdf)
    except Exception as exc:  # noqa: BLE001
        result.error = _reason(exc)
        result.seconds = time.perf_counter() - started
        return result
    result.notes.extend(model.warnings)

    bound = None
    if file_set.xml is None:
        result.notes.append("결과 XML 이 없다 — 계산서를 건너뛰고 도면만 그린다")
    else:
        try:
            bound = bind_results(model, file_set.xml)
        except Exception as exc:  # noqa: BLE001
            result.notes.append(f"결과 XML 을 결합하지 못했다 — {_reason(exc)}")

    if bound is not None:
        def _report(out: Path) -> tuple[int, Sequence[str]]:
            written = render_report(file_set.sdf, file_set.xml, out, generated=generated)
            return written.pages, written.warnings

        result.parts.append(_attempt(REPORT_PART, parts_dir / f"{REPORT_PART}.pdf",
                                     _report))

    for name in presets:
        def _iso(out: Path, name: str = name) -> tuple[int, Sequence[str]]:
            drawn = render_iso(bound or model, out, preset=name,
                               section=file_set.stem)
            return 1, drawn.warnings

        result.parts.append(_attempt(name, parts_dir / f"{name}.pdf", _iso))

    if combine and result.done:
        merged = dest / f"{file_set.stem}.pdf"
        try:
            _combine(result.done, merged)
            result.combined = merged
        except Exception as exc:  # noqa: BLE001
            result.notes.append(f"합본을 만들지 못했다 — {_reason(exc)}")

    result.seconds = time.perf_counter() - started
    return result


# ── 일괄 ────────────────────────────────────────────────────────────────────


def run_batch(root: str | Path, out_dir: str | Path, *,
              presets: Sequence[str] = DEFAULT_PRESETS,
              combine: bool = True,
              generated: datetime | None = None,
              sets: Sequence[FileSet] | None = None,
              progress: Callable[[int, int, SetResult], None] | None = None
              ) -> BatchReport:
    """폴더 하나를 무인으로 훑는다. 세트 하나가 터져도 나머지는 계속 간다."""
    base, out = Path(root), Path(out_dir)
    found = list(sets) if sets is not None else discover(base)
    report = BatchReport(root=base, out_dir=out)
    started = time.perf_counter()

    for index, file_set in enumerate(found, start=1):
        # 같은 이름이 동마다 있다 — 입력 폴더 구조를 그대로 살려 덮어쓰기를 막는다.
        relative = file_set.sdf.parent.relative_to(base) if base.is_dir() else Path()
        result = process_set(file_set, out / relative, presets=presets,
                             combine=combine, generated=generated)
        report.results.append(result)
        if progress is not None:
            progress(index, len(found), result)

    report.seconds = time.perf_counter() - started
    return report


# ── CLI ─────────────────────────────────────────────────────────────────────


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        prog="python -m core.d_batch",
        description="PIPENET 없이 SDF+XML 세트를 계산서·ISO 도면 PDF 로 일괄 변환한다.")
    parser.add_argument("root", help="입력 폴더 (또는 SDF 한 개)")
    parser.add_argument("out_dir", help="출력 폴더")
    parser.add_argument("--preset", action="append", dest="presets",
                        choices=sorted(PRESETS),
                        help=f"ISO 프리셋. 되풀이 지정. 기본 {' '.join(DEFAULT_PRESETS)}")
    parser.add_argument("--no-combine", action="store_true", help="합본을 만들지 않는다")
    parser.add_argument("--limit", type=int, help="앞에서 N 세트만")
    parser.add_argument("--report", help="리포트를 적을 텍스트 파일")
    args = parser.parse_args(argv)

    found = discover(args.root)
    if args.limit is not None:
        found = found[: args.limit]
    if not found:
        print(f"세트를 찾지 못했다 — {args.root}", file=sys.stderr)
        return 2

    def show(index: int, total: int, result: SetResult) -> None:
        print(f"[{index}/{total}] {result.line()}", flush=True)

    report = run_batch(args.root, args.out_dir, sets=found,
                       presets=tuple(args.presets or DEFAULT_PRESETS),
                       combine=not args.no_combine, progress=show)
    print()
    print(report.text())
    if args.report:
        Path(args.report).write_text(report.text(), encoding="utf-8")
    return 0 if not report.failed else 1


if __name__ == "__main__":
    raise SystemExit(main())
