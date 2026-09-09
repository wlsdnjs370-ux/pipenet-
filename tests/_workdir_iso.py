# -*- coding: utf-8 -*-
"""시험용 작업폴더 격리 — `conftest` 가 아니라 **전용 모듈**에 둔다.

★`from conftest import ...` 로 쓰면 pytest 가 올려 둔 **다른 폴더의**
  conftest(예: `tests/characterization/conftest.py`)를 집을 수 있다. 실제로
  그 ImportError 를 만났다 — conftest 는 이름이 겹치기 쉬운 자리다.
"""
from __future__ import annotations


# ─────────────────────────────────────────────────────────────────────────
class isolated_workdir:
    """작업폴더를 임시 사본으로 돌린다 — **반드시 되돌린다.**

    ★되돌리기가 이 헬퍼의 존재 이유다. 격리를 손으로 걸면 그 프로세스에 남아
      **다음 시험이 빈 폴더를 본다** — 실측으로 그렇게 「B1F reopen 실패
      (FileNotFoundError)」가 났다. 한 시험을 지키려던 격리가 옆 시험을 깬 것이라,
      한 판에 돌릴 수 없는 시험 묶음이 되어 버린다.

    `with` 로 쓴다. 나갈 때 원래 경로를 그대로 되돌리고 임시 폴더를 지운다.

        with isolated_workdir() as work:
            ...            # 이 안에서의 쓰기는 전부 work 로 간다
    """

    def __init__(self, prefix: str = "mf_iso_", copy_key: str | None = None):
        self.prefix = prefix
        self.copy_key = copy_key
        self._tmp = None
        self._saved = None

    def __enter__(self):
        import os
        import shutil
        import tempfile

        from services.cad_import.pipeline import handoff as hf

        src_pick, src_edit = hf.pick_out_dir(), hf.default_edits_dir()
        src_cache = hf.import_write_root()
        # ★주입점 하나로 못박는다(그리고 나갈 때 그 하나만 되돌린다) — 종전에는
        #   함수를 갈아끼우고 그것으로 안 따라오는 두 상수를 따로 덮었다.
        #
        # ★★되돌릴 값은 «그때 실제로 가리키던 폴더» 다. `_WRITE_ROOT_OVERRIDE`
        #   를 그대로 담아 두면, 아직 부팅 전이라 None 이던 경우에 None 으로
        #   되돌아가고 — `_boot()` 는 한 번만 도므로(`_booted`) 다시는 안 채워진다.
        #   그러면 그 뒤의 모든 시험이 cwd 상대경로 "docs/import" 를 보게 된다.
        #   (이 지시서가 없애려던 바로 그 사고다.)
        self._saved = src_cache

        self._tmp = tempfile.TemporaryDirectory(prefix=self.prefix)
        work = self._tmp.name
        hf.set_write_root(work)
        os.makedirs(hf.pick_out_dir(), exist_ok=True)
        os.makedirs(hf.default_edits_dir(), exist_ok=True)

        if self.copy_key:      # 입력이 필요한 시험은 «읽어서 복사» 한다
            for src, dst in ((src_pick, hf.pick_out_dir()),
                             (src_edit, hf.default_edits_dir())):
                if src and os.path.isdir(src):
                    for name in os.listdir(src):
                        if self.copy_key in name:
                            try:
                                shutil.copy2(os.path.join(src, name),
                                             os.path.join(dst, name))
                            except OSError:
                                pass
            if src_cache and os.path.isdir(src_cache):
                for name in os.listdir(src_cache):
                    if name.startswith("_edit_disp_cache_") \
                            and self.copy_key in name:
                        try:
                            shutil.copy2(os.path.join(src_cache, name),
                                         os.path.join(work, name))
                        except OSError:
                            pass
        return work

    def __exit__(self, *exc):
        from services.cad_import.pipeline import handoff as hf

        hf.set_write_root(self._saved)
        if self._tmp is not None:
            self._tmp.cleanup()
        return False
