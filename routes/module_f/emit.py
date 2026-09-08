# -*- coding: utf-8 -*-
"""[H-6] 결합망 → 입력파일 (특허 S750 · S760 · S770).

특허 도 9 의 주석이 이 파일의 규범이다::

    S750 에서 생성된 입력파일에는 호칭경 대조용 표준 자료를 함께 동봉하며,
    S760 의 변환은 **별도의 산출이 아니라 S750 의 결과 파일 자체를 원본으로**
    삼으므로 모든 형식이 항상 같은 배관망을 가리킨다.

그래서 순서가 정해져 있다:

    ① SDF 를 쓴다                       (S750 — 권위 있는 원본)
    ② 같은 폴더에 SLF 가 따라 나온다     (호칭경↔내경 대조 자료 · 동봉)
    ③ **그 SDF 파일을 읽어** KFP·HAS 를 만든다 (S760)
    ④ 넷을 하나로 압축한다               (S770)

③ 이 핵심이다. 결합망 객체에서 형식마다 따로 뽑으면 «같은 배관망» 이라는 보장이
사라진다 — 형식별로 다른 반올림·다른 누락이 들어가고, 그 어긋남은 솔버를 돌려
봐야 드러난다. 파일 하나를 원본으로 삼으면 그럴 수가 없다.

KFP·HAS 실패는 SDF 를 막지 않는다(A 의 `_emit_subnetwork_bundle` 과 같은 판단).
다만 **조용히 넘기지 않는다** — 무엇이 빠졌는지 반환에 남긴다.
"""
from __future__ import annotations

import zipfile
from pathlib import Path


def emit_merged(combined, out_dir, *, title: str = "모듈 F 통합",
                stem: str = "module_f_merged",
                coord_scale: float = 1.0,
                iso_nodes: list | None = None) -> dict:
    """결합망 하나 → {sdf, slf, kfp, has, zip, warnings}. 값은 절대경로.

    `combined` 는 `stitch_riser_and_heads` 산출(`CombinedTables`)이다.

    `iso_nodes` 를 주면 **아이소매트릭 좌표 한 벌**을 더 낸다(`<stem>_iso.*`).
    사용자 요청(2026-09-08: 「저번처럼 나오던 아이소매트릭 형태 위상으로
    .sdf 파일이 출력되었으면 좋겠는데, 그거 되게 잘 그려졌어서」) — 모듈 A 의
    통합이 `combined_<id>_iso.sdf` 를 함께 내는 것과 같은 규약이다.

    ★두 벌은 «좌표만» 다르다. 길이·관경·표고·노즐은 같은 표에서 나오므로
      수리계산 값은 한 글자도 안 바뀐다(좌표는 PIPENET 캔버스 표시용이고,
      좌표가 선언 길이를 덮지 못하게 하는 잠금은 `parse_sdf` 에 서 있다).
      절점 좌표는 화면 미리보기가 쓰는 그 함수(`merge.bake_combined_iso`)가
      만든 것을 그대로 받는다 — 여기서 다시 셈하면 화면과 파일이 갈린다.
    """
    from remote30_full_network import ProjectContext, emit_full_sdf

    out = Path(out_dir)
    out.mkdir(parents=True, exist_ok=True)
    warnings: list[str] = []

    # ① S750 — 권위 있는 원본.
    sdf = out / f"{stem}.sdf"
    emit_full_sdf(combined, sdf, ctx=ProjectContext.titled(title))

    # ② 호칭경 대조 자료 — emit_full_sdf 가 같은 폴더에 함께 낸다.
    #    PIPENET 은 .sdf 와 .slf 가 같은 폴더에 있어야 내경을 찾는다.
    slf = out / f"{stem}.slf"
    if not slf.is_file():
        warnings.append("SLF(호칭경 대조 자료)가 생성되지 않았습니다 — "
                        "PIPENET 에서 관경이 Unset 으로 보일 수 있습니다.")

    # ③ S760 — 별도 산출이 아니라 위 SDF **파일** 을 원본으로 변환한다.
    kfp = out / f"{stem}.kfp"
    try:
        from remote30_prototype import emit_kfp
        emit_kfp(sdf, kfp, coord_scale=float(coord_scale))
    except Exception as exc:  # noqa: BLE001 — SDF 출력을 막지 않는다
        warnings.append(f"KFP 변환 실패: {type(exc).__name__}: {exc}")

    has = out / f"{stem}.has"
    try:
        from remote30_prototype import emit_has
        emit_has(sdf, has)
    except Exception as exc:  # noqa: BLE001
        warnings.append(f"HAS 변환 실패: {type(exc).__name__}: {exc}")

    # ④ S770 — 형식별 파일 + 대조 자료를 하나로.
    zip_path = out / f"{stem}.zip"
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for p in (sdf, slf, kfp, has):
            if p.is_file():
                zf.write(p, arcname=p.name)

    out_files = {
        "sdf": str(sdf),
        "slf": str(slf) if slf.is_file() else None,
        "kfp": str(kfp) if kfp.is_file() else None,
        "has": str(has) if has.is_file() else None,
        "zip": str(zip_path),
        "warnings": warnings,
    }

    # ⑤ 아이소매트릭 한 벌 — 좌표만 갈아 끼운 사본으로 같은 길을 한 번 더 탄다.
    if iso_nodes:
        import copy as _copy
        iso_tbl = _copy.copy(combined)
        iso_tbl.nodes = list(iso_nodes)
        iso_stem = f"{stem}_iso"
        iso_sdf = out / f"{iso_stem}.sdf"
        try:
            emit_full_sdf(iso_tbl, iso_sdf,
                          ctx=ProjectContext.titled(f"{title} (아이소)"))
            out_files["sdf_iso"] = str(iso_sdf)
            iso_slf = out / f"{iso_stem}.slf"
            if iso_slf.is_file():
                out_files["slf_iso"] = str(iso_slf)
            iso_kfp = out / f"{iso_stem}.kfp"
            try:
                from remote30_prototype import emit_kfp as _ek
                _ek(iso_sdf, iso_kfp, coord_scale=float(coord_scale))
                out_files["kfp_iso"] = str(iso_kfp)
            except Exception as exc:  # noqa: BLE001 — 아이소 실패가 본산출을 막지 않는다
                warnings.append(f"아이소 KFP 변환 실패: {type(exc).__name__}: {exc}")
            iso_has = out / f"{iso_stem}.has"
            try:
                from remote30_prototype import emit_has as _eh
                _eh(iso_sdf, iso_has)
                out_files["has_iso"] = str(iso_has)
            except Exception as exc:  # noqa: BLE001
                warnings.append(f"아이소 HAS 변환 실패: {type(exc).__name__}: {exc}")
            with zipfile.ZipFile(zip_path, "a", zipfile.ZIP_DEFLATED) as zf:
                for p in (iso_sdf, iso_slf, iso_kfp, iso_has):
                    if p.is_file():
                        zf.write(p, arcname=p.name)
        except Exception as exc:  # noqa: BLE001
            # ★본 산출(평면 좌표)은 이미 나와 있다 — 아이소가 실패해도 그것을
            #   버리지 않는다. 못 냈다는 사실만 올린다(S340: 조용히 메우지 않는다).
            warnings.append(f"아이소 SDF 생성 실패: {type(exc).__name__}: {exc}")
    return out_files


# 연장 비교 허용 오차 (m). 형식마다 소수 자릿수가 달라 완전 동일을 요구하지 않는다.
LENGTH_TOL_M = 0.01


def cross_check(files: dict) -> dict:
    """산출 3종이 **같은 배관망을 가리키는가** — S760 의 약속을 실측한다.

    ★절점·배관 «수» 로 견주면 안 된다. SDF→KFP/HAS 변환은 직선 위 통과절점을
      통합하므로(특허 S440·S443 · `kfp_sdf_converter.simplify_passthrough_nodes`)
      수는 정당하게 줄어든다 — 실측으로 라이저 체인 때문에 59절점이 11절점이
      됐다. 그것을 «어긋남» 으로 읽으면 멀쩡한 산출을 불량으로 보고하게 된다.

      통합이 보존해야 하는 것은 **총 연장과 노즐**이다(S443: 통합된 관로의
      길이는 합산하여 보존된다). 그 둘로 견준다. 수는 참고로만 싣는다.

    못 읽는 형식은 건너뛴다 — 비교 대상이 없는 것과 어긋나는 것은 다르다.
    """
    got: dict[str, dict] = {}

    def _measure(name, fn):
        path = files.get(name)
        if not path or not Path(path).is_file():
            return
        try:
            got[name] = fn(Path(path))
        except Exception as exc:  # noqa: BLE001
            got[name] = {"error": f"{type(exc).__name__}: {exc}"}

    def _from_net(net) -> dict:
        total = 0.0
        for p in net.pipes.values():
            try:
                total += float(getattr(p, "length_m", 0.0) or 0.0)
            except (TypeError, ValueError):
                pass
        nozzles = sum(
            1 for n in net.nodes.values()
            if str(getattr(n, "kind", "")).lower() in ("nozzle", "head"))
        return {"nodes": len(net.nodes), "pipes": len(net.pipes),
                "total_m": round(total, 3), "nozzles": nozzles}

    def _sdf(p):
        from kfp_sdf_converter import parse_sdf
        return _from_net(parse_sdf(str(p)))

    def _kfp(p):
        from kfp_sdf_converter import parse_kfp
        return _from_net(parse_kfp(str(p)))

    def _has(p):
        from has_converter import parse_has
        return _from_net(parse_has(str(p)))

    _measure("sdf", _sdf)
    _measure("kfp", _kfp)
    _measure("has", _has)

    ok = {k: v for k, v in got.items() if "error" not in v}
    agree = None
    detail = ""
    if ok:
        lengths = [v["total_m"] for v in ok.values()]
        nozzles = {v["nozzles"] for v in ok.values()}
        len_ok = (max(lengths) - min(lengths)) <= LENGTH_TOL_M
        nz_ok = len(nozzles) <= 1
        agree = len_ok and nz_ok
        if not len_ok:
            detail = f"연장 불일치 {min(lengths):.3f}~{max(lengths):.3f} m"
        elif not nz_ok:
            detail = f"노즐 수 불일치 {sorted(nozzles)}"
    return {"per_format": got, "agree": agree, "detail": detail,
            "compared": sorted(ok),
            "invariant": "총 연장 · 노즐 수 (절점 수는 통합으로 정당하게 줄어든다)"}
