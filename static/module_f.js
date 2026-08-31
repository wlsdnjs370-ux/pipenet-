"use strict";
(function () {
  const $ = (id) => document.getElementById(id);
  const PICK_PX = 10;                       // 모듈 E 의 pick.board.PICK_PX 와 같은 눈금

  const S = {
    sid: null, key: null, stage: "open",
    slot: "plan",                           // [H-0] 활성 도면 슬롯 (S650)
    zones: [],                              // 최불리 후보를 가둘 사각형 (A 의 zones)
    refCounts: [],                          // NFTC 103 표 2.1.1.1 (서버가 준다)
    boreColor: true,                        // 관경 근거로 배관 색 나누기
    // [H-2 · H-3] 계통도·기계실이 찍는 두 점 · 무엇을 찍는 중인지
    sub: { picks: [null, null], arm: null, summary: null },
    merge: null,                            // [H-5] 통합 상태 한 장
    // [A 방식] 자동 추출 — 방식·알람밸브·헤드 후보·완료 여부
    method: null, autoAlarm: null, autoArm: null,
    autoHeads: [], autoDone: false, autoSummary: null,
    autoView: null,                         // 뽑아낸 배관망(절점·배관)
    // [S270 · S310] 검출한 망 — 최불리를 고르기 «전» 의 것
    autoNet: null, autoNetView: null,
    showJunc: true,                         // 이음자리(티·교차) 표시
    autoPipe: [],                           // 「배관으로 취급」 지정 묶음
    world: null, pick: null, edit: null, design: null,
    suggest: null,                          // [F-5] 찍기 후보 (제안만)
    suggestOff: new Set(),                  // 반영에서 제외한 후보 index
    // [F-8] 정찰·채택 — 후보는 «제안» 이고 board 에 닿은 것은 클릭뿐이다.
    recon: null,                            // 정찰 요약 (수치만)
    adopted: null,                          // 채택된 후보 index
    ghosts: null,                           // 찍히지 못한 후보 index (유령)
    showLow: false,                         // 낮은 띠 후보 표시 토글
    handoff: null,                          // [F-8d] 자동 → 손질 이어받기 제안
    // 되돌리기 — 자동·계통도는 기록이 화면 쪽에만 있어 여기에 쌓는다.
    // (찍기·손질은 엔진이 제 기록을 들고 있다)
    undo: [],
    hidden: new Set(),                      // 숨긴 레이어 묶음 id
    view: { scale: 1, ox: 0, oy: 0 },
    poll: null, es: null,
  };

  // ── 캔버스 기본기 ───────────────────────────────────────────────
  const cv = $("cv");
  const ctx = cv.getContext("2d");

  function resize() {
    const r = cv.parentElement.getBoundingClientRect();
    const dpr = window.devicePixelRatio || 1;
    cv.width = Math.max(1, Math.floor(r.width * dpr));
    cv.height = Math.max(1, Math.floor(r.height * dpr));
    ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
    draw();
  }
  window.addEventListener("resize", resize);

  function cssSize() {
    const r = cv.parentElement.getBoundingClientRect();
    return { w: r.width, h: r.height };
  }

  // 세계좌표 → 화면좌표. DXF 는 y 가 위로 자라므로 뒤집는다.
  function sx(x) { return (x - S.view.ox) * S.view.scale; }
  function sy(y) { return cssSize().h - (y - S.view.oy) * S.view.scale; }
  function wx(px) { return px / S.view.scale + S.view.ox; }
  function wy(py) { return (cssSize().h - py) / S.view.scale + S.view.oy; }

  function fit(bounds) {
    if (!bounds) return;
    const { w, h } = cssSize();
    const bw = Math.max(1e-6, bounds.maxx - bounds.minx);
    const bh = Math.max(1e-6, bounds.maxy - bounds.miny);
    const sc = Math.min(w / bw, h / bh) * 0.92;
    S.view.scale = sc;
    S.view.ox = bounds.minx - (w / sc - bw) / 2;
    S.view.oy = bounds.miny - (h / sc - bh) / 2;
    draw();
  }

  // 뽑아낸 배관망이 놓인 자리 — 없으면 null.
  function autoNetBounds() {
    const ns = (S.autoView && S.autoView.nodes) || [];
    if (!ns.length) return null;
    const xs = ns.map((n) => n.x), ys = ns.map((n) => n.y);
    const bx = Math.min(...xs), by = Math.min(...ys);
    const bw = Math.max(...xs) - bx, bh = Math.max(...ys) - by;
    // 설계면적 둘레가 조금은 보여야 «어디서 뽑혔나» 를 읽는다.
    const pad = Math.max(bw, bh, 1000) * 0.25;
    return { minx: bx - pad, miny: by - pad,
             maxx: bx + bw + pad, maxy: by + bh + pad };
  }

  function curBounds() {
    if ((S.stage === "edit" || S.stage === "conv") && S.edit) return S.edit.bounds;
    // 추출이 끝났으면 「화면 맞춤」도 뽑은 망을 기준으로 삼는다 — 도면 전체로
    // 맞추면 27개짜리 설계면적이 화면에서 점 하나가 된다(실측 LH306: 도면
    // 971m 대 헤드군 25m).
    if (S.stage === "auto" && S.autoDone) {
      const b = autoNetBounds();
      if (b) return b;
    }
    if (S.world) return S.world.bounds;
    return null;
  }

  $("btn-fit").onclick = () => fit(curBounds());

  let drag = null;
  // 영역 지정 드래그 — 켜져 있을 때만 왼쪽 버튼을 가로챈다(패닝은 그대로).
  let zoneDrag = null;
  cv.addEventListener("mousedown", (e) => {
    const armed = (S.stage === "edit" && $("ed-zone-arm").checked)
               || (S.stage === "auto" && $("au-zone-arm").checked);
    if (e.button === 0 && !e.shiftKey && armed) {
      zoneDrag = { x0: wx(e.offsetX), y0: wy(e.offsetY),
                   x1: wx(e.offsetX), y1: wy(e.offsetY) };
      e.preventDefault();
      return;
    }
    if (e.button === 1 || e.button === 2 || e.shiftKey) {
      drag = { x: e.offsetX, y: e.offsetY, ox: S.view.ox, oy: S.view.oy };
      e.preventDefault();
    }
  });
  cv.addEventListener("mousemove", (e) => {
    $("coord").textContent =
      `x ${wx(e.offsetX).toFixed(0)}  y ${wy(e.offsetY).toFixed(0)}`;
    if (zoneDrag) {
      zoneDrag.x1 = wx(e.offsetX);
      zoneDrag.y1 = wy(e.offsetY);
      draw();
      return;
    }
    if (!drag) return;
    S.view.ox = drag.ox - (e.offsetX - drag.x) / S.view.scale;
    S.view.oy = drag.oy + (e.offsetY - drag.y) / S.view.scale;
    draw();
  });
  window.addEventListener("mouseup", () => {
    drag = null;
    if (zoneDrag) {
      const z = zoneDrag;
      zoneDrag = null;
      // 손이 떨려 생긴 점은 영역이 아니다 — 화면에서 8px 넘게 끈 것만 받는다.
      const px = Math.abs(z.x1 - z.x0) * S.view.scale;
      const py = Math.abs(z.y1 - z.y0) * S.view.scale;
      if (px < 8 || py < 8) { draw(); return; }
      markUndo("영역 그리기");
      S.zones.push([Math.min(z.x0, z.x1), Math.min(z.y0, z.y1),
                    Math.max(z.x0, z.x1), Math.max(z.y0, z.y1)]);
      // 자동 경로는 영역이 서버의 필수 입력이라 곧바로 올린다(anchored 의
      // head_region). 수동은 최불리 선정을 누를 때 함께 보낸다.
      if (S.stage === "auto") pushAutoZones();
      else renderZones();
      draw();
    }
  });
  cv.addEventListener("contextmenu", (e) => e.preventDefault());
  cv.addEventListener("wheel", (e) => {
    e.preventDefault();
    const k = e.deltaY < 0 ? 1.15 : 1 / 1.15;
    const bx = wx(e.offsetX), by = wy(e.offsetY);
    S.view.scale *= k;
    S.view.ox = bx - e.offsetX / S.view.scale;
    S.view.oy = by - (cssSize().h - e.offsetY) / S.view.scale;
    draw();
  }, { passive: false });

  cv.addEventListener("click", (e) => {
    if (e.shiftKey) return;
    const x = wx(e.offsetX), y = wy(e.offsetY);
    const maxD = PICK_PX / S.view.scale;
    if (S.stage === "pick") pickClick(x, y, maxD);
    else if (S.stage === "edit") editClick(x, y, maxD);
    // [F-10e] 평면에서 보는 동안은 «그 자리에서» 고칠 수 있어야 한다. 손질과
    //   같은 클릭 경로를 그대로 태운다 — 새 길을 만들지 않는다(D-F10-6).
    else if (S.stage === "design" && planUnderlayOn() && S.edit) {
      editClick(x, y, maxD);
    }
    else if (S.stage === "sub") subClick(x, y);
    else if (S.stage === "auto") autoClick(x, y);
  });

  // ── 되돌리기 단축키 ────────────────────────────────────────────
  // Ctrl+Z 한 번 = 한 박자 되돌리기. 단계에 맞는 «되돌리기» 단추를 **그대로
  // 누른다** — 여기서 API 를 따로 부르면 단추와 단축키의 동작이 갈라진다
  // (한쪽만 고치는 사고가 난다).
  //
  // ★입력칸 안에서는 손대지 않는다. 변환 폼에 숫자를 치다 Ctrl+Z 를 누르면
  //   그건 «글자 되돌리기» 지 «손질 되돌리기» 가 아니다 — 브라우저에 맡긴다.
  //   Ctrl+Shift+Z(다시 실행)도 여기서 다루지 않는다(shiftKey 로 걸러진다).
  window.addEventListener("keydown", (e) => {
    if ((e.key || "").toLowerCase() !== "z") return;
    if (!(e.ctrlKey || e.metaKey) || e.altKey || e.shiftKey) return;

    const el = document.activeElement;
    const tag = el ? el.tagName : "";
    if (tag === "INPUT" || tag === "TEXTAREA" || tag === "SELECT"
        || (el && el.isContentEditable)) return;

    e.preventDefault();
    // 무거운 작업이 도는 중에 망을 되돌리면 그 작업과 엇갈린다.
    if (!$("busy").classList.contains("hidden")) {
      say("작업이 끝난 뒤에 되돌릴 수 있습니다.", "warn");
      return;
    }
    undoStep();
  });

  // ── 되돌리기 — 모든 단계에서 한 박자씩 ──────────────────────────
  // 찍기·손질은 엔진이 제 기록을 들고 있어 단추를 그대로 누른다. 자동·계통도는
  // 그 기록이 «화면 쪽» 에만 있으므로 여기서 직접 쌓는다. 눌러도 아무 일이
  // 없으면 사람은 프로그램이 고장 났다고 읽는다 — 실제로 그 말을 들었다.
  const UNDO_CAP = 60;

  function snapAuto() {
    return { alarm: S.autoAlarm ? S.autoAlarm.slice() : null,
             zones: S.zones.map((z) => z.slice()) };
  }
  function snapSub() {
    return { picks: S.sub.picks.map((p) => (p ? p.slice() : null)) };
  }

  // 되돌릴 «직전» 에 부른다 — 바꾸기 전 모습을 담아 둔다.
  function markUndo(label) {
    const st = S.stage;
    const snap = st === "auto" ? snapAuto() : st === "sub" ? snapSub() : null;
    if (!snap) return;
    S.undo.push({ stage: st, snap, label });
    if (S.undo.length > UNDO_CAP) S.undo.shift();
  }

  async function undoStep() {
    // 엔진이 기록을 들고 있는 단계는 그 단추를 그대로 누른다 — 여기서 API 를
    // 따로 부르면 단추와 단축키의 동작이 갈라진다.
    if (S.stage === "pick") { $("pk-undo").click(); return; }
    if (S.stage === "edit") { $("ed-undo").click(); return; }

    const i = [...S.undo].reverse().findIndex((u) => u.stage === S.stage);
    if (i < 0) { say("이 단계에는 되돌릴 것이 없습니다.", "warn"); return; }
    const item = S.undo.splice(S.undo.length - 1 - i, 1)[0];

    if (item.stage === "auto") {
      S.autoAlarm = item.snap.alarm ? item.snap.alarm.slice() : null;
      S.zones = item.snap.zones.map((z) => z.slice());
      S.autoArm = null;
      $("au-anchor").classList.remove("on");
      try {
        await post("/api/module-f/auto/anchor", S.autoAlarm
          ? { sid: S.sid, x: S.autoAlarm[0], y: S.autoAlarm[1] }
          : { sid: S.sid });
        await post("/api/module-f/auto/zones", { sid: S.sid, zones: S.zones });
      } catch (err) { say(err.message, "err"); return; }
      renderAuto(null);
      draw();
    } else if (item.stage === "sub") {
      S.sub.picks = item.snap.picks.map((p) => (p ? p.slice() : null));
      S.sub.arm = null;
      renderSubPicks();
      draw();
    }
    say(`되돌렸습니다 — ${item.label}`, "ok");
  }

  // ── 그리기 ─────────────────────────────────────────────────────
  function draw() {
    const { w, h } = cssSize();
    ctx.clearRect(0, 0, w, h);
    ctx.fillStyle = "#000";
    ctx.fillRect(0, 0, w, h);
    // 변환 단계에서도 손질한 망을 계속 보여준다 — 값만 채우는 동안 화면이
    // 검게 비면 무엇을 변환하는지 알 수 없다.
    if (S.stage === "design" && S.design) {
      // [F-10e] «평면에서 보기» — 밑그림(배경 도면)까지 깔아 손질 화면과 같은
      //   그림을 만든다. 여기서는 밑그림과 최불리망이 같은 세계 좌표라
      //   어긋날 수가 없다(아이소 아래에 깔지 않는 이유는 BLOCKED §17).
      if (planUnderlayOn() && S.edit) {
        if (editBgOn() && S.world) drawWorld(true, EDIT_BG_ALPHA);
        drawEdit();
        drawDesignMarks();
      }
      else if (designMarksOn() && S.edit) { drawEdit(); drawDesignMarks(); }
      else drawDesign();
      drawFocus();      // [F-10f] 「확인할 것」에서 고른 자리를 마지막에 덧그린다
    }
    else if ((S.stage === "edit" || S.stage === "conv") && S.edit) {
      // [F-10c] 배경 도면을 «밑에» 깐다 — 전사 17:53 · 23:06. corridor 만 뜨면
      //   어디서 뽑힌 망인지 안 보여 결과가 옳은지 판단할 수가 없다.
      //   board 보다 더 흐리게(0.07) 둬서 위계가 배경 → 비corridor → corridor
      //   순으로 읽히게 한다.
      if (editBgOn() && S.world) drawWorld(true, EDIT_BG_ALPHA);
      drawEdit();
      // [F-8d] 자동이 알던 자리 — «제안» 이라 점선 고리로만 그린다. 반영은
      // 사람이 단추를 눌러 기존 손질 클릭 경로로만 들어간다.
      if (S.stage === "edit") drawHandoffHints();
    }
    else if (S.world) {
      // 자동 추출이 끝났으면 도면 전체를 내리고 뽑아낸 망만 살린다.
      const focus = S.stage === "auto" && S.autoDone && S.autoView;
      drawWorld(focus);
      if (S.stage === "pick" && S.suggest) drawSuggest();
      // [H-2 · H-3] 찍은 두 점을 도면 위에 남긴다 — 어디를 찍었는지 안 보이면
      // 추출이 틀렸을 때 클릭이 문제인지 도면이 문제인지 가릴 수 없다.
      if (S.stage === "sub") drawSubPicks();
      if (S.stage === "auto") {
        drawZones();
        // 검출한 망(파랑)을 먼저 깔고, 뽑은 최불리(청록)를 그 위에 얹는다.
        drawAutoNetwork();
        if (focus) drawAutoNet();
        drawAuto(focus);
      }
    }
  }

  // `dim` — 뽑아낸 배관망을 돋보이게 하려고 나머지를 거의 투명한 점선으로
  // 내린다. 지우지는 않는다: 어디서 뽑혔는지 보이지 않으면 결과가 옳은지
  // 판단할 수가 없다.
  //
  // [F-10c] `alpha` 로 농도를 바꿀 수 있다 — 손질 밑그림은 board 보다 더
  // 흐려야 위계가 «배경 → 비corridor → corridor» 순으로 읽힌다. dim 경로로
  // 부르면 묶음만 그리고 일찍 끝나므로 찍기 하이라이트가 딸려오지 않는다.
  function drawWorld(dim, alpha) {
    ctx.lineWidth = 1;
    if (dim) {
      ctx.globalAlpha = (alpha === undefined ? 0.16 : alpha);
      ctx.setLineDash([2, 4]);
    }
    for (const b of S.world.bundles) {
      if (S.hidden.has(b.id)) continue;
      ctx.strokeStyle = b.css;
      ctx.beginPath();
      const sg = b.segs;
      for (let i = 0; i < sg.length; i += 4) {
        ctx.moveTo(sx(sg[i]), sy(sg[i + 1]));
        ctx.lineTo(sx(sg[i + 2]), sy(sg[i + 3]));
      }
      const cr = b.circles;
      for (let i = 0; i < cr.length; i += 3) {
        const r = cr[i + 2] * S.view.scale;
        if (r < 0.4) continue;
        ctx.moveTo(sx(cr[i]) + r, sy(cr[i + 1]));
        ctx.arc(sx(cr[i]), sy(cr[i + 1]), r, 0, Math.PI * 2);
      }
      const ar = b.arcs;
      for (let i = 0; i < ar.length; i += 5) {
        const r = ar[i + 2] * S.view.scale;
        if (r < 0.4) continue;
        // 화면은 y 를 뒤집으므로 각도 방향도 뒤집힌다.
        const a0 = -ar[i + 3] * Math.PI / 180;
        const a1 = a0 - ar[i + 4] * Math.PI / 180;
        ctx.moveTo(sx(ar[i]) + Math.cos(a0) * r, sy(ar[i + 1]) + Math.sin(a0) * r);
        ctx.arc(sx(ar[i]), sy(ar[i + 1]), r, a0, a1, true);
      }
      ctx.stroke();
    }
    if (dim) { ctx.globalAlpha = 1; ctx.setLineDash([]); return; }
    if (!S.pick) return;
    const hl = S.pick.highlight;
    ctx.lineWidth = 2.2;
    ctx.strokeStyle = "#b366ff";
    ctx.beginPath();
    for (const s of hl.pipe_segs) {
      ctx.moveTo(sx(s[0]), sy(s[1]));
      ctx.lineTo(sx(s[2]), sy(s[3]));
    }
    for (const s of hl.tri_segs) {
      ctx.moveTo(sx(s[0]), sy(s[1]));
      ctx.lineTo(sx(s[2]), sy(s[3]));
    }
    ctx.stroke();
    ctx.strokeStyle = "#ff5cf0";
    ctx.lineWidth = 2;
    ctx.beginPath();
    for (const c of hl.head_circles) {
      const r = Math.max(2.5, c[2] * S.view.scale);
      ctx.moveTo(sx(c[0]) + r, sy(c[1]));
      ctx.arc(sx(c[0]), sy(c[1]), r, 0, Math.PI * 2);
    }
    ctx.stroke();
    if (hl.last_click) {
      ctx.strokeStyle = "#ffffff";
      ctx.lineWidth = 1;
      const px = sx(hl.last_click[0]), py = sy(hl.last_click[1]);
      ctx.beginPath();
      ctx.moveTo(px - 9, py); ctx.lineTo(px + 9, py);
      ctx.moveTo(px, py - 9); ctx.lineTo(px, py + 9);
      ctx.stroke();
    }
  }

  // ── [F-10c] 시각 위계 — 반짝 · 페이드 · 동그라미 ───────────────
  //
  // 전사 06:57 「반짝반짝」 · 08:57 「도면은 그대로 있잖아」 · 23:45 「동그라미를
  // 치는 게 더 보기가 좋다」. 셋을 한꺼번에 만든다:
  //
  //   배경 도면      가장 흐리게(EDIT_BG_ALPHA) — 사라지지 않는다
  //   비corridor 망   흐린 점선 (`ed-worst-view` 기본 «dim»)
  //   corridor       굵게 + 담당 헤드 수 비례 굵기 + 은은한 펄스
  //   선정 헤드       흰 동그라미 · 앵커는 빨간 겹원
  //
  // ★펄스는 «몇 번 하고 멈춘다». 계속 깜빡이면 눈이 피로하고, 전사의 의도는
  //   강조지 점멸 지속이 아니다(지시서 F-10c 표의 단서 그대로).
  const EDIT_BG_ALPHA = 0.07;
  const PULSE_MS = 1800;        // 약 2~3회
  const PULSE_CYCLES = 2.5;
  let pulseT0 = 0;
  let pulseRAF = 0;

  const editBgOn = () => {
    const el = $("ed-bg");
    return !el || el.checked;
  };

  /** 0 이면 펄스 없음. 끝나면 스스로 0 으로 떨어진다. */
  function pulseAmt() {
    if (!pulseT0) return 0;
    const t = (performance.now() - pulseT0) / PULSE_MS;
    if (t >= 1) { pulseT0 = 0; return 0; }
    // 뒤로 갈수록 잦아든다 — 끝이 뚝 끊기지 않게.
    return (1 - t) * (0.5 - 0.5 * Math.cos(2 * Math.PI * PULSE_CYCLES * t));
  }

  function startPulse() {
    pulseT0 = performance.now();
    if (pulseRAF) return;
    const step = () => {
      draw();
      if (pulseT0) { pulseRAF = requestAnimationFrame(step); }
      else { pulseRAF = 0; draw(); }   // 마지막 한 장은 «정지» 상태로
    };
    pulseRAF = requestAnimationFrame(step);
  }

  function drawEdit() {
    const e = S.edit;
    // 최불리를 고른 뒤에는 «그것만» 보고 싶을 때가 있다. 표시만 바꾼다 —
    // 망은 그대로고, 산출물 범위는 「변환」 단계의 체크박스가 정한다.
    // (선정 전에는 걸지 않는다: 아무것도 안 보이는 화면이 되어 버린다.)
    const wv = e.worst ? $("ed-worst-view").value : "all";
    ctx.lineWidth = 1.4;
    if (wv !== "only") {
      ctx.save();
      if (wv === "dim") {
        ctx.globalAlpha = 0.22;
        ctx.lineWidth = 0.6;
        ctx.setLineDash([2, 5]);
      }
      for (const g of e.body_groups) {
        ctx.strokeStyle = g.css;
        ctx.beginPath();
        // segs 는 평평한 배열이다 — [x1,y1,x2,y2, x1,y1,x2,y2, …] (찍기 캔버스와 같은 규약)
        const sg = g.segs;
        for (let i = 0; i < sg.length; i += 4) {
          ctx.moveTo(sx(sg[i]), sy(sg[i + 1]));
          ctx.lineTo(sx(sg[i + 2]), sy(sg[i + 3]));
        }
        ctx.stroke();
      }
      ctx.restore();
      ctx.setLineDash([]);
    }
    if (e.wet_pipes.length && wv === "all") {
      ctx.strokeStyle = e.palette.wet;
      ctx.lineWidth = 2.6;
      ctx.beginPath();
      for (const s of e.wet_pipes) {
        ctx.moveTo(sx(s[0]), sy(s[1]));
        ctx.lineTo(sx(s[2]), sy(s[3]));
      }
      ctx.stroke();
    }
    // 자동 이음 «후보» — 아직 배관이 아니다. 실측 배관과 절대 같이 그리지 않는다:
    // 점선 + 다른 색. 틈이 수십 mm 라 화면 맞춤에서는 선이 1픽셀도 안 되므로,
    // 끊긴 자리마다 작은 고리를 하나 얹어 «여기가 그 자리» 를 보이게 한다.
    if (e.autojoin && e.autojoin.lines.length && wv === "all") {
      ctx.save();
      ctx.strokeStyle = "#ff9900";
      ctx.lineWidth = 1.6;
      ctx.setLineDash([5, 4]);
      ctx.beginPath();
      for (const s of e.autojoin.lines) {
        ctx.moveTo(sx(s[0]), sy(s[1]));
        ctx.lineTo(sx(s[2]), sy(s[3]));
      }
      ctx.stroke();
      ctx.setLineDash([]);
      ctx.globalAlpha = 0.85;
      ctx.beginPath();
      for (const s of e.autojoin.lines) {
        const mx = sx((s[0] + s[2]) / 2), my = sy((s[1] + s[3]) / 2);
        ctx.moveTo(mx + 3.5, my);
        ctx.arc(mx, my, 3.5, 0, Math.PI * 2);
      }
      ctx.stroke();
      ctx.restore();
    }
    // 선정 밖 헤드 — «최불리망만» 에서는 감추고, «비활성» 에서는 흐리게.
    // 선정된 30개는 아래에서 흰 고리로 다시 그려지므로 여기서 빠져도 보인다.
    if (wv !== "only") {
      ctx.save();
      if (wv === "dim") ctx.globalAlpha = 0.25;
      for (const hd of e.heads) {
        const r = Math.max(2, hd[2] * S.view.scale);
        ctx.fillStyle = hd[3];
        ctx.beginPath();
        ctx.arc(sx(hd[0]), sy(hd[1]), r, 0, Math.PI * 2);
        ctx.fill();
      }
      ctx.restore();
    }
    ctx.lineWidth = 2;
    ctx.strokeStyle = "#ffffff";
    for (const m of e.multi_heads) {
      const r = Math.max(3, m[2] * S.view.scale) + 2;
      ctx.beginPath();
      ctx.arc(sx(m[0]), sy(m[1]), r, 0, Math.PI * 2);
      ctx.stroke();
    }
    if (e.pending) {
      ctx.strokeStyle = "#ffffff";
      ctx.lineWidth = 3;
      ctx.beginPath();
      ctx.moveTo(sx(e.pending[0][0]), sy(e.pending[0][1]));
      ctx.lineTo(sx(e.pending[1][0]), sy(e.pending[1][1]));
      ctx.stroke();
    }
    if (e.selected_head) {
      ctx.strokeStyle = "#fbbf24";
      ctx.lineWidth = 2.4;
      const r = Math.max(4, e.selected_head[2] * S.view.scale) + 3;
      ctx.beginPath();
      ctx.arc(sx(e.selected_head[0]), sy(e.selected_head[1]), r, 0, Math.PI * 2);
      ctx.stroke();
    }
    // Remote 30 — 최불리 «배관망(corridor)». 균일 흰선이 아니라 담당 헤드
    // 수(load)에 따라 굵기가 다르다: 주배관은 굵고 말단 가지는 얇다. 이 굵기가
    // 곧 NFPC 별표1 의 관경 서열이라, 수리계산 대상 망이 한눈에 읽힌다.
    if (e.worst) {
      const wm = e.worst.max_load || 1;
      // [F-10c] 은은한 펄스 — 방금 뽑힌 corridor 를 몇 번 도드라지게 한 뒤
      //   가만히 둔다. 굵기와 밝기를 같이 살짝 올린다.
      const p = pulseAmt();
      ctx.strokeStyle = "#ffffff";
      ctx.lineCap = "round";
      for (const s of e.worst.corridor) {
        const load = s[4] || 1;
        ctx.globalAlpha = Math.min(1, (0.5 + 0.5 * (load / wm)) + 0.35 * p);
        ctx.lineWidth = (1.4 + 3.6 * Math.sqrt(load / wm)) * (1 + 0.45 * p);
        ctx.beginPath();
        ctx.moveTo(sx(s[0]), sy(s[1]));
        ctx.lineTo(sx(s[2]), sy(s[3]));
        ctx.stroke();
      }
      ctx.globalAlpha = 1;
      // 최원 유하거리 «경로» — 급수원 → 앵커. corridor 의 부분집합이지만
      // 굵기만으로는 그 거리가 어느 줄인지 읽을 수 없어 따로 덧그린다.
      // 기준압이 어느 관을 타고 오는지가 보여야 관경을 키울지 경로를 줄일지
      // 정할 수 있다.
      const ap = e.worst.anchor_path || [];
      if (ap.length > 1) {
        ctx.strokeStyle = "#ff3b3b";
        ctx.globalAlpha = 0.85;
        ctx.lineWidth = 2.2;
        ctx.setLineDash([9, 5]);
        ctx.beginPath();
        ctx.moveTo(sx(ap[0][0]), sy(ap[0][1]));
        for (let i = 1; i < ap.length; i++) ctx.lineTo(sx(ap[i][0]), sy(ap[i][1]));
        ctx.stroke();
        ctx.setLineDash([]);
        ctx.globalAlpha = 1;
      }
      ctx.strokeStyle = "#ffffff";
      ctx.lineCap = "butt";
      ctx.lineWidth = 2;
      // 「최불리망만」 에서는 도면 헤드가 통째로 감춰지므로 여기서 속을 채운다 —
      // 안 그러면 선정된 30개가 빈 고리로만 남아 헤드로 읽히지 않는다.
      ctx.fillStyle = "rgba(255,255,255,.30)";
      for (const h of e.worst.heads) {
        const r = Math.max(4, h[2] * S.view.scale) + 2;
        ctx.beginPath();
        ctx.arc(sx(h[0]), sy(h[1]), r, 0, Math.PI * 2);
        if (wv === "only") ctx.fill();
        ctx.stroke();
      }
      // 앵커 = 가장 불리한 지점(기준압을 잡는 헤드). 빨간 겹원으로 못박는다.
      if (e.worst.anchor) {
        const a = e.worst.anchor;
        const r = Math.max(5, a[2] * S.view.scale) + 4;
        ctx.strokeStyle = "#ff3b3b";
        ctx.lineWidth = 2.6;
        ctx.beginPath();
        ctx.arc(sx(a[0]), sy(a[1]), r, 0, Math.PI * 2);
        ctx.stroke();
        ctx.beginPath();
        ctx.arc(sx(a[0]), sy(a[1]), r + 4, 0, Math.PI * 2);
        ctx.stroke();
      }
    }
    drawZones();
    markers(e.sources, e.palette.source, 7);
    markers(e.valves, e.palette.valve, 6);
  }

  function markers(list, color, size) {
    ctx.fillStyle = color;
    ctx.strokeStyle = "#000";
    ctx.lineWidth = 1;
    for (const p of list) {
      const px = sx(p[0]), py = sy(p[1]);
      ctx.beginPath();
      ctx.rect(px - size, py - size, size * 2, size * 2);
      ctx.fill();
      ctx.stroke();
    }
  }

  // ── 통신 ───────────────────────────────────────────────────────
  async function api(path, opts) {
    const r = await fetch(path, opts);
    let d;
    try { d = await r.json(); }
    catch (err) { throw new Error(`서버 응답을 읽지 못했습니다 (HTTP ${r.status}).`); }
    if (!r.ok || d.ok === false) throw new Error(d.message || `HTTP ${r.status}`);
    return d;
  }
  const post = (path, body) => api(path, {
    method: "POST", headers: { "Content-Type": "application/json" },
    body: JSON.stringify(body),
  });

  function say(msg, cls) {
    const el = $("status");
    el.textContent = msg;
    el.className = cls || "";
  }
  function busy(on, text) {
    $("busy").classList.toggle("hidden", !on);
    if (text) $("busy-text").textContent = text;
  }

  // 슬롯마다 «실제로 밟는 단계» 가 다르다. 계통도·기계실은 찍을 재료가 없어
  // 찍기·손질·변환이 통째로 없다 — 두 점을 찍어 경로를 뽑으면 끝이다.
  // 통합(S700)은 슬롯에 딸리지 않는다: 세 슬롯의 결과를 모으는 자리라 어느
  // 슬롯에서 보든 맨 끝에 붙는다.
  // 평면도는 방식에 따라 밟는 단계가 통째로 다르다. 자동은 찍기·손질·변환이
  // 없다 — 알람밸브와 영역만 정하면 검출이 표까지 낸다.
  const STAGE_FLOW = {
    plan: ["open", "pick", "edit", "conv", "design"],
    plan_auto: ["open", "auto", "design"],
    system: ["open", "sub"],
    machineroom: ["open", "sub"],
  };
  const STAGE_LABEL = {
    open: "도면 열기", pick: "찍기", edit: "손질", conv: "변환",
    design: "수리계산", sub: "경로 추출", auto: "자동 추출", merge: "통합",
  };
  // 각 단계가 켜는 패널. 한 곳에 모아 둔다 — 예전에는 toggle 이 아홉 줄로
  // 흩어져 있어 패널을 하나 늘릴 때마다 빠뜨릴 자리가 늘었다.
  const STAGE_PANELS = {
    open: ["panel-open", "panel-resume"],
    // [F-10a] 「방식」 단계는 없앴다 — 물을 것이 없어졌다(D-F10-1).
    //   시작 배너와 고급은 찍기·손질 양쪽에 붙는다: 자동으로 흘러온 길이라
    //   사람이 고른 기억이 없으므로 화면이 무엇으로 시작했는지 말해 주고,
    //   설정(채택 기준)과 자동 차선 입구는 접힌 채로 곁에 둔다.
    pick: ["panel-start", "panel-pick", "panel-advanced", "panel-layers"],
    edit: ["panel-start", "panel-edit", "panel-advanced"],
    conv: ["panel-conv"],
    design: ["panel-design"],
    // 레이어 토글은 여기서도 쓴다 — 건축 배경을 못 끄면 두 점을 찍기 어렵다.
    sub: ["panel-sub", "panel-layers"],
    auto: ["panel-auto", "panel-layers"],
    merge: ["panel-merge"],
  };
  const ALL_PANELS = [...new Set(Object.values(STAGE_PANELS).flat())];

  function stageFlow() {
    // [F-10a] 방식을 묻지 않으므로 기본 흐름은 처음부터 정해져 있다. 다만
    //   도면을 아직 안 읽었으면 「도면 열기」 하나만 보인다 — 갈 수 있는
    //   곳이 그것뿐이라서다. (자동 차선을 고급에서 고르면 그때 갈린다.)
    if (S.slot === "plan" && !S.method) return ["open"];
    const key = (S.slot === "plan" && S.method === "auto") ? "plan_auto" : S.slot;
    return (STAGE_FLOW[key] || STAGE_FLOW.plan).concat(["merge"]);
  }

  // 지금 갈 수 있는 단계인가 — 재료가 없는 곳으로 보내면 빈 화면만 나온다.
  function stageReachable(name) {
    if (name === "open") return true;
    if (name === "merge") return !!S.sid;
    if (name === "sub") return !!S.world;
    if (name === "auto") return !!S.world;
    if (name === "pick") return !!S.world;
    if (name === "edit" || name === "conv") return !!S.edit;
    // 수리계산은 두 길 다에서 온다 — 자동은 표가 이미 나왔을 때만.
    if (name === "design") return S.method === "auto" ? !!S.autoDone : !!S.edit;
    return false;
  }

  function renderSteps() {
    const flow = stageFlow();
    const idx = flow.indexOf(S.stage);
    const box = $("steps");
    box.innerHTML = "";
    flow.forEach((k, i) => {
      const el = document.createElement("div");
      el.textContent = STAGE_LABEL[k];
      if (i === idx) el.classList.add("on");
      else if (idx >= 0 && i < idx) el.classList.add("done");
      const ok = stageReachable(k);
      el.style.cursor = ok ? "pointer" : "default";
      el.style.opacity = ok ? "" : ".45";
      if (ok) el.onclick = () => gotoStage(k);
      box.appendChild(el);
    });
  }

  // 단계바를 눌러 오갈 수 있게 한다 — 「어디로 가려면 어느 단추를 눌러야
  // 하는지」를 외우지 않아도 되게. 갈 수 없는 단계는 흐리게 두고 막는다.
  async function gotoStage(name) {
    if (name === S.stage || !stageReachable(name)) return;
    try {
      if (name === "merge") { await loadMerge(); return; }
      if (name === "sub") { await loadSub(); return; }
      if (name === "auto") { await loadAuto(); return; }
      if (name === "design") { setStage("design"); await designPreview(); return; }
      if (name === "conv") {
        setStage("conv");
        await loadFields();
        renderConvSummary();
        return;
      }
      setStage(name);
    } catch (err) { say(err.message, "err"); }
  }

  function setStage(name) {
    S.stage = name;
    const want = new Set(STAGE_PANELS[name] || []);
    for (const id of ALL_PANELS) {
      $(id).classList.toggle("hidden", !want.has(id));
    }
    // 「이어서 열기」는 평면도만의 개념이다(찍은 스펙 목록) — 계통도·기계실
    // 슬롯에서 띄우면 남의 도면 목록을 이 슬롯에 여는 것처럼 보인다.
    if (name === "open") {
      $("panel-resume").classList.toggle("hidden", S.slot !== "plan");
    }
    renderSteps();
    draw();
  }

  // ── 잡 진행 — SSE 우선, 미지원·오류 시 폴링 폴백 [F-6] ──────────
  //
  // 진행 묶음은 평소 접혀 있다. 다만 **작업이 도는 동안에는 저절로 편다** —
  // 큰 도면은 파싱이 십 분을 넘기고, 그때 화면이 아무 말도 안 하면 멈춘 것과
  // 구별되지 않는다. 끝나면 다시 접되 실패했을 때는 열어 둔다(무엇이 잘못됐는지
  // 읽어야 한다). 제목 옆 표에는 어느 단계가 도는지만 짧게 남긴다.
  let jobWasRunning = false;      // 시작하는 순간에만 편다 — 매 박자 열면
                                  // 사용자가 접어 둔 것을 계속 되돌린다
  function logFold(open) {
    const h2 = document.querySelector('h2.fold[data-fold="log-body"]');
    if (h2) toggleFold(h2, open);
  }
  function jobChip(text) {
    const chip = $("job-chip");
    chip.textContent = text || "—";
    chip.classList.toggle("hidden", !text);
  }
  // ★진행 «상태» 는 한 곳에서만 그린다. 예전에는 SSE 경로가 job-line 을 제
  //   손으로 다시 그려, jobRender 에 붙인 것이 스트림에서는 아예 돌지 않았다
  //   (실측: 진행 자동열림이 SSE 에서만 안 먹었다). 두 경로가 같은 것을 그리면
  //   한쪽만 고쳐지는 날이 반드시 온다.
  function jobStatus(j) {
    $("job-line").textContent =
      `${j.phase} · ${j.state === "run" ? "진행 중" : j.state} · ${j.elapsed}s`
      + (j.queued ? " (다른 작업이 끝나기를 기다리는 중)" : "");
    if (j.state === "run") {
      jobChip(`${j.phase} ${j.elapsed}s`);
      if (!jobWasRunning) logFold(true);
      jobWasRunning = true;
    }
  }

  // 폴링 전용 — 로그 «본문» 까지 서버 tail 로 채운다. 스트림은 줄이 생기는
  // 대로 쌓으므로(line 이벤트) 여기서 tail 로 덮으면 오히려 짧아진다.
  function jobRender(j) {
    $("log").textContent = j.lines && j.lines.length ? j.lines.join("\n") : "…";
    $("log").scrollTop = $("log").scrollHeight;
    jobStatus(j);
  }
  function jobFinish(j, onDone) {
    stopWatch();
    busy(false);
    jobWasRunning = false;
    if (j.state === "error") {
      jobChip("실패");
      logFold(true);            // 실패는 열어 둔다 — 읽어야 고친다
      say(j.error, "err");
      return;
    }
    jobChip("");
    logFold(false);
    // onDone 은 async 다. 여기서 잡지 않으면 실패가 조용히 삼켜진다.
    Promise.resolve()
      .then(() => onDone(j))
      .catch((err) => say(err.message || String(err), "err"));
  }

  function watch(onDone) {
    stopWatch();
    if (window.EventSource) {
      try { return watchStream(onDone); } catch (e) { /* 폴백 */ }
    }
    return watchPoll(onDone);
  }

  function watchStream(onDone) {
    // 진행 줄이 «생기는 순간» 흐른다 — 0.7초 폴링 박자를 기다리지 않는다.
    const es = new EventSource(`/api/module-f/job/stream?sid=${S.sid}`);
    S.es = es;
    const lines = [];
    es.addEventListener("line", (ev) => {
      lines.push(JSON.parse(ev.data));
      if (lines.length > 400) lines.splice(0, 200);
      $("log").textContent = lines.join("\n");
      $("log").scrollTop = $("log").scrollHeight;
    });
    es.addEventListener("state", (ev) => {
      const j = JSON.parse(ev.data);
      jobStatus(j);            // 진행 줄·표·자동열림 — 폴링과 같은 한 벌
      if (j.state === "run") { busy(true, `${j.phase} · ${j.elapsed}s`); return; }
      if (j.state === "idle") return;      // 잡이 아직 안 붙었다 — 서버가 기다린다
      jobFinish(j, onDone);
    });
    es.onerror = () => {
      // 프록시가 스트림을 못 넘기는 환경 — 조용히 폴링으로 돌아간다.
      if (S.es === es) {
        es.close();
        S.es = null;
        watchPoll(onDone);
      }
    };
  }

  function watchPoll(onDone) {
    S.poll = setInterval(async () => {
      let j;
      try { j = await api(`/api/module-f/job?sid=${S.sid}`); }
      catch (err) { stopWatch(); busy(false); say(err.message, "err"); return; }
      jobRender(j);
      if (j.state === "run") { busy(true, `${j.phase} · ${j.elapsed}s`); return; }
      jobFinish(j, onDone);
    }, 700);
  }

  function stopWatch() {
    if (S.poll) clearInterval(S.poll);
    S.poll = null;
    if (S.es) { S.es.close(); S.es = null; }
  }

  // ── 1. 열기 ────────────────────────────────────────────────────
  $("btn-open").onclick = async () => {
    const f = $("dxf").files[0];
    if (!f) { say("DXF 파일을 고르세요.", "warn"); return; }
    const fd = new FormData();
    fd.append("dxf_file", f);
    // [H-0] 활성 슬롯으로 넣는다. 세션이 이미 있으면 그 세션의 슬롯을 채운다
    // (S650 회귀 한 바퀴) — 없으면 새 세션이 이 종류로 시작한다.
    fd.append("kind", S.slot);
    if (S.sid) fd.append("sid", S.sid);
    busy(true, "업로드 중…");
    try {
      // 올리기까지다 — 읽기는 방식이 정해진 뒤(`/slot/read`).
      const d = await api("/api/module-f/slot/open", { method: "POST", body: fd });
      S.sid = d.sid;
      S.method = null;
      S.zones = []; S.autoAlarm = null; S.autoHeads = []; S.autoDone = false;
      S.world = null; S.edit = null; S.key = null;
      // 새 도면이다 — 앞 도면의 정찰·채택·이어받기 표시는 뜻이 없다.
      S.recon = null; S.suggest = null; S.ghosts = null; S.adopted = null;
      S.handoff = null; S.autoNet = null; S.autoNetView = null;
      S.autoView = null;
      // [F-10b] 화면 손질 모드도 도면에 딸린다 — 새 도면은 원클릭부터.
      S.emode = null;
      // ★되돌리기 기록도 함께 버린다. 남겨 두면 Ctrl+Z 가 «앞 도면의 좌표» 를
      //   되살려 이 도면에 씌운다.
      S.undo = [];
      $("pk-adopt-box").classList.add("hidden");
      renderHandoff();
      say(`${d.filename} 읽는 중…`);
      // 읽어서 화면에 띄우는 것까지는 방식과 무관하다 — 도면을 먼저 보여 준
      // 뒤에 어떻게 읽을지 묻는다.
      watch(async () => {
        await loadWorldRaw();
        fit(S.world.bounds);
        if (!d.needs_method) { await loadSub(); loadSlots(); return; }
        // 한 줄로 자르고 전체 이름은 툴팁에 — 좁은 옆판에서 제목이 토막나면
        // 어느 도면을 여는지가 안 읽힌다.
        const nm = `${S.key} · 선분 ${S.world.counts.segs.toLocaleString()}`;
        $("adv-file").textContent = nm;
        $("adv-file").title = nm;
        await loadRecon();          // [F-8c] 정찰 수치
        loadSlots();
        // [F-10a · D-F10-1] 여기서 묻지 않는다. 정찰이 성했으면 채택→조립까지
        //   흘려보내고, 못 쓰겠으면 «묻지 않고» 찍기 화면으로 내려간다.
        await autoStart();
      });
    } catch (err) { busy(false); say(err.message, "err"); }
  };

  // 어느 길로 갈지 정한다. 도면은 이미 읽혀 화면에 있다 — 수동은 더 읽을
  // 것이 없고(찍기판이 이미 섰다), 자동만 A 의 파서를 한 번 더 돌린다.
  async function readSlot(method) {
    busy(true, method === "auto"
      ? "자동 추출 준비 중… (처음 한 번만 오래 걸립니다)"
      : "여는 중…");
    try {
      const d = await post("/api/module-f/slot/read", { sid: S.sid, method });
      S.method = d.method || "manual";
      renderSteps();
      const go = async () => {
        if (S.method === "auto") { await loadAuto(); }
        else { await loadWorld(true); }   // 이미 받아 둔 도면을 그대로 쓴다
        loadSlots();
      };
      if (d.started) { watch(go); }      // 자동 — 파싱을 기다린다
      else { busy(false); await go(); }  // 수동 — 기다릴 것이 없다
    } catch (err) { busy(false); say(err.message, "err"); }
  }

  // ── [F-10a] 기본 흐름 — 묻지 않고 흐른다 ──────────────────────
  //
  // 차선은 코드로 살아 있다(엔드포인트·테스트 그대로). 사라진 것은 «질문»
  // 하나다: 업로드 시점에는 이 도면이 자동으로 될지 사람도 모르므로,
  // 「어떻게 추출할까요」는 답할 수 없는 질문이었다(D-F10-1).
  //
  // 대신 정찰 결과가 스스로 답한다. 정찰은 열기 잡 안에서 이미 돌았다(F-8a).
  const CONF_CHOICES = [
    [0.9, "높음 (≥0.9) — 기본"],
    [0.75, "중간 이상 (≥0.75)"],
    [0.0, "전부"],
  ];

  function fillConf() {
    const sel = $("adv-conf");
    if (sel.options.length) return;
    for (const [v, label] of CONF_CHOICES) {
      const o = document.createElement("option");
      o.value = String(v);
      o.textContent = label;
      sel.appendChild(o);
    }
    sel.value = "0.9";              // D-F8-4 — 기본값은 그대로다
  }

  const num = (v) => Number(v) || 0;

  // 고른 문턱으로 몇 개가 채택 대상인가.
  function reconPick(lo) {
    if (lo === undefined) lo = confMin();
    const b = (S.recon && S.recon.bands) || {};
    const hi = num(b["높음(≥0.9)"]), mid = num(b["중간(≥0.75)"]);
    const low = num(b["낮음"]);
    return lo >= 0.9 ? hi : lo >= 0.75 ? hi + mid : hi + mid + low;
  }

  // [F-11a · D-F11-2] 기본 채택 임계는 «도면 분포» 가 정한다(서버의 지배 띠
  //   규칙). 다만 사람이 고급에서 손으로 고르면 그것이 이긴다 — 규칙은 기본값
  //   이지 잠금이 아니다.
  function confMin() {
    fillConf();
    if (S.confManual) {
      const v = parseFloat($("adv-conf").value);
      if (Number.isFinite(v)) return v;
    }
    const a = (S.recon && S.recon.adopt) || null;
    if (a && a.conf_min !== null && a.conf_min !== undefined) {
      return Number(a.conf_min);
    }
    const v = parseFloat($("adv-conf").value);
    return Number.isFinite(v) ? v : 0.9;
  }

  // 정찰이 성해서 «자동으로 시작할 수 있는가». 이 판단이 곧 흐름의 갈림이고,
  // 사람에게 묻지 않는다 — 못 쓰겠으면 찍기 화면으로 내려가 사유를 적는다.
  //   ★배관 묶음이 0 이면 채택할 재료가 없다. 재료 없이 채택을 부르면 서버가
  //     「재료를 하나도 못 찍었습니다」로 끝나므로, 그 전에 갈라야 한다.
  function reconReady() {
    const r = S.recon;
    if (!r || r.state !== "ok") {
      return { ok: false, why: (r && r.state === "error")
        ? `자동 인식이 실패했습니다 — 색으로 직접 찍어 주세요. (${r.error || ""})`
        : "자동 인식 결과가 없습니다 — 색으로 직접 찍어 주세요." };
    }
    if (!num((r.bundles || {}).PIPE)) {
      return { ok: false,
        why: "자동 인식이 배관 레이어를 찾지 못했습니다 — 색으로 직접 찍어 주세요." };
    }
    // ★찍을 헤드가 하나도 없으면 조립이 죽는다(엔진이 `KeyError: 'heads'`).
    //   엔진은 읽기 전용이라 문 앞에서 가른다 — 이 게이트는 최후 방어로 남는다.
    //
    //   [F-11a · D-F11-2] 판단 기준이 «절대 0.9» 에서 «규칙이 고른 임계» 로
    //   바뀌었다. 예전에는 LH306(높음 0/42)이 여기서 막혔는데, 이제 지배 띠
    //   규칙이 중간까지 채택하므로 살아서 지나간다. 규칙조차 0 을 내는
    //   도면에서만 이 문이 닫힌다.
    if (!reconPick(confMin())) {
      const a = (r.adopt || {});
      return { ok: false,
        why: (a.why || "자동 인식이 찍을 만한 헤드 후보를 못 찾았습니다.")
          + " — 직접 찍거나, 고급에서 채택 기준을 낮춰 다시 채택하세요." };
    }
    return { ok: true };
  }

  function renderRecon() {
    const box = $("adv-recon");
    const r = S.recon;
    fillConf();
    if (!r || r.state === "none" || r.state === "error") {
      const bad = r && r.state === "error";
      box.innerHTML = `<div class="hint">자동 인식 `
        + (bad ? `<span class="warn">실패</span>` : "결과 없음") + "</div>";
      if (bad) box.title = r.error || "";
      $("adv-conf-row").classList.add("hidden");
      $("adv-conf-why").textContent = "";
      return;
    }
    const b = r.bands || {}, bd = r.bundles || {};
    const cell = (cls, label, n) =>
      `<div class="${cls}"><i>${label}</i><b>${num(n).toLocaleString()}</b></div>`;
    box.innerHTML =
      `<div class="bands">`
      + cell("hi", "높음 ≥0.9", b["높음(≥0.9)"])
      + cell("", "중간 ≥0.75", b["중간(≥0.75)"])
      + cell("lo", "낮음", b["낮음"])
      + `</div>`
      + `<div class="hint">헤드 후보 <b>${r.n.toLocaleString()}</b>개 · `
      + `배관 묶음 <b>${num(bd.PIPE)}</b>개`
      + (num(bd.HEAD) ? ` · 헤드 레이어 <b>${num(bd.HEAD)}</b>개` : "")
      + `</div>`
      // [F-11a] 어느 규칙이 발동했는지 카드에도 적는다 — 조용한 규칙 전환은
      //   새 은닉 오류다. 사람이 손으로 고른 뒤에는 그렇다고 말한다.
      + (r.adopt ? `<div class="hint">${S.confManual
            ? "채택 기준을 <b>직접 고른</b> 상태입니다."
            : r.adopt.why}</div>` : "");
    $("adv-conf-row").classList.remove("hidden");
    renderConfHint();
  }

  // 기준을 만족하는 후보가 몇 개인지 적는다. 0 이면 다시 채택을 잠근다 —
  // 눌러도 아무 일이 안 일어나는 단추는 고장으로 읽힌다. 실측으로 흔하다:
  // A 는 «알려진 블록 참조» 에만 0.95 를 주므로, 헤드를 레이어에 직접 그린
  // 도면은 높음 띠가 0 이 된다(LH306 0/42 · B1F 72/3,338).
  function renderConfHint() {
    if (!S.recon || S.recon.state !== "ok") return;
    const n = reconPick();
    $("adv-readopt").disabled = n === 0;
    $("adv-conf-why").innerHTML = n
      ? `이 기준으로 <b>${n.toLocaleString()}개</b>를 찍습니다.`
      : '<span class="warn">이 기준에 맞는 후보가 없습니다 — '
        + "기준을 낮춰 보세요.</span>";
  }

  // 사람이 손으로 고르는 순간 «수동» 이 된다 — 그 뒤로는 규칙이 안 이긴다.
  $("adv-conf").onchange = () => {
    S.confManual = true;
    // 카드도 다시 그린다 — 「규칙이 정했다」가 「직접 고른 상태」로 바뀌어야
    // 화면이 사실을 말한다.
    renderRecon();
  };

  async function loadRecon() {
    try {
      const d = await api(`/api/module-f/recon?sid=${S.sid}`);
      S.recon = d.recon || null;
    } catch (err) { S.recon = { state: "error", error: err.message }; }
    renderRecon();
  }

  // 시작 배너 — 무엇으로 시작했는지, 되돌릴 수 있는지 한 줄.
  function startNote(html, warn) {
    const box = $("start-note");
    box.innerHTML = html;
    box.classList.toggle("warn", !!warn);
  }

  // [D-F10-2] 자동 차선은 고급 안 한 줄로 남는다 — 엔드포인트·테스트·특허
  //   실시예는 그대로다. 화면에서만 «질문» 이 아니라 «선택» 이 되었다.
  $("adv-auto").onclick = (ev) => { ev.preventDefault(); readSlot("auto"); };

  // 인식 결과를 찍는다 — 채택까지. `to` 가 "edit" 이면 조립까지 이어서 간다.
  async function adoptRun(lo, to) {
    await post("/api/module-f/pick/adopt", {
      sid: S.sid, materials: true, heads: { conf_min: lo },
    });
    return new Promise((resolve) => {
      watch(async () => {
        const j = await api(`/api/module-f/convert/result?sid=${S.sid}`);
        const r = j.result || {};
        await loadWorld(true);            // 찍기 화면으로 (도면은 이미 있다)
        if (!r.ok) {
          startNote(r.error || "채택에 실패했습니다 — 직접 찍어 주세요.", true);
          say(r.error || "채택에 실패했습니다.", "err");
          resolve(false);
          return;
        }
        if (r.state) S.pick = r.state;
        await applyAdopt(r, lo);
        renderPick();
        draw();
        if (to !== "edit") { resolve(true); return; }
        // ★조립도 잡이다. 세션 잡은 한 번에 하나이므로 «끝난 뒤» 에 건다.
        busy(true, "배관망 구성 중…");
        try {
          await post("/api/module-f/pick/commit", { sid: S.sid });
          watch(async () => {
            await loadEdit();
            const g = (S.ghosts && S.ghosts.size) || 0;
            // [F-11a] 어느 규칙으로 채택했는지 배너에도 한 줄 — 사람이 고른
            //   기억이 없는 길이라 «왜 이만큼인가» 를 화면이 말해야 한다.
            const a = (S.recon && S.recon.adopt) || null;
            startNote(`자동 인식 결과로 시작했습니다 — 채택 `
              + `<b>${num(r.head_applied).toLocaleString()}</b>개`
              + (g ? ` · 유령 <b>${g.toLocaleString()}</b>개` : "")
              + ` · 단계바의 「찍기」로 내려가 고칠 수 있습니다.`
              + (a && !S.confManual ? `<br>${a.why}` : ""));
            resolve(true);
          });
        } catch (err) {
          busy(false);
          startNote(`배관망 구성에 실패했습니다 — 찍기에서 고쳐 주세요. `
            + `(${err.message})`, true);
          resolve(false);
        }
      });
    });
  }

  // [F-10a · D-F10-1] 업로드 뒤 «질문 0» 으로 손질까지. 못 가면 찍기에서 멈추되
  //   그것도 묻지 않는다 — 왜 멈췄는지 배너에 적을 뿐이다.
  async function autoStart() {
    const gate = reconReady();
    await post("/api/module-f/slot/read", { sid: S.sid, method: "manual" });
    S.method = "manual";
    renderSteps();
    if (!gate.ok) {
      busy(false);
      await loadWorld(true);
      startNote(gate.why, true);
      say(gate.why, "warn");
      return;
    }
    busy(true, "인식 결과를 찍는 중…");
    try {
      await adoptRun(confMin(), "edit");
    } catch (err) { busy(false); say(err.message, "err"); }
  }

  // 기준을 바꿔 다시 채택 — 찍기 화면에서 멈춘다(사람이 보고 판단할 자리다).
  $("adv-readopt").onclick = async () => {
    busy(true, "인식 결과를 다시 찍는 중…");
    try {
      await adoptRun(confMin(), "pick");
    } catch (err) { busy(false); say(err.message, "err"); }
  };

  // 채택 결과를 화면 상태로 — 후보 좌표는 여기서 처음 받아 온다(카드에서는
  // 수치만 받았다. 3천 점을 카드 그릴 때마다 내려보낼 이유가 없다).
  async function applyAdopt(r, lo) {
    const d = await api(`/api/module-f/recon?sid=${S.sid}&heads=1`);
    S.suggest = d.heads || [];
    S.suggestOff = new Set();
    S.ghosts = new Set((r.skipped_heads || []).map((g) => g.i));
    S.adopted = new Set();
    for (let i = 0; i < S.suggest.length; i++) {
      if (Number(S.suggest[i].conf) >= lo && !S.ghosts.has(i)) S.adopted.add(i);
    }
    $("pk-adopt-box").classList.remove("hidden");
    $("pk-suggest-apply").disabled = !S.suggest.length;
    $("pk-suggest-clear").disabled = !S.suggest.length;
    suggestInfo();
    $("pk-adopt-info").innerHTML =
      kv("재료", `${(r.mat_applied || []).length}묶음`
         + ((r.mat_skipped || []).length
            ? ` · 건너뜀 ${(r.mat_skipped || []).length}` : ""))
      + kv("헤드", `찍힘 ${num(r.head_applied)} · 이미 반영 ${num(r.head_already)}`
           + ` · <span class="${num(r.head_skipped) ? "warn" : ""}">`
           + `유령 ${num(r.head_skipped)}</span>`);
    const ghost = num(r.head_skipped);
    say(`인식 결과를 찍었습니다 — 재료 ${(r.mat_applied || []).length}묶음 · `
      + `헤드 ${num(r.head_applied) + num(r.head_already)}개.`
      + (ghost ? ` 점선 ${ghost}개는 그 자리에 찍을 도형이 없어 남겨 뒀습니다 `
                 + "— 직접 찍거나 무시하세요." : ""),
      ghost ? "warn" : "ok");
  }

  $("pk-show-low").onchange = (e) => { S.showLow = e.target.checked; draw(); };

  // 도면만 받아 캔버스에 올린다 — 찍기판이 없는 슬롯(계통도·기계실)도 쓴다.
  async function loadWorldRaw() {
    const d = await api(`/api/module-f/world?sid=${S.sid}`);
    S.world = d.world; S.key = d.key; S.pick = d.state;
    buildLayers();
    renderCats();
    return d;
  }

  // `reuse` — 이미 받아 둔 도면을 그대로 쓴다. 「불러오기」가 방금 받아 온
  // 것을 「수동으로 읽기」에서 또 받을 이유가 없다(실측 LH306 1.44 MB —
  // 내려받고 파싱하고 레이어 목록을 다시 짓는 값이 통째로 헛일이다).
  async function loadWorld(reuse) {
    if (!reuse || !S.world) await loadWorldRaw();
    setStage("pick");
    fit(S.world.bounds);
    renderPick();
    loadSlots();
    const c = S.world.counts, dr = S.world.dropped;
    let msg = `${S.key} · 선분 ${c.segs.toLocaleString()} · 원 ${c.circles.toLocaleString()}`
            + ` · 호 ${c.arcs.toLocaleString()}`;
    if (dr.segs || dr.circles || dr.arcs) {
      msg += ` — 화면에는 선분 ${dr.segs.toLocaleString()}·원 ${dr.circles.toLocaleString()}`
           + `·호 ${dr.arcs.toLocaleString()} 개를 뺐습니다(표시 상한). 찍기 판정은 전량 대상입니다.`;
      say(msg, "warn");
    } else { say(msg); }
  }

  // ── 접이식 묶음 ────────────────────────────────────────────────
  // `<h2 class="fold" data-fold="본문-id">` 하나로 걸린다. 제목이 곧 단추다.
  // 자잘한 단추가 늘 펼쳐져 있으면 정작 그 단계에서 할 일이 안 보인다.
  function toggleFold(h2, open) {
    const body = $(h2.dataset.fold);
    if (!body) return;
    const next = (open === undefined) ? body.classList.contains("hidden") : open;
    body.classList.toggle("hidden", !next);
    h2.classList.toggle("on", next);
  }

  for (const h2 of document.querySelectorAll("h2.fold")) {
    h2.onclick = () => toggleFold(h2);
    toggleFold(h2, false);          // 처음엔 접어 둔다
  }

  // ── 기준개수 (NFTC 103 표 2.1.1.1) ─────────────────────────────
  // 30 고정이 아니다. 용도·층수·부착높이에 따라 10 · 20 · 30 으로 갈리고,
  // 이 값은 설계면적 헤드 수뿐 아니라 수원량·펌프 유량까지 함께 정한다.
  // 표는 서버(core/nftc_rules.py)가 유일한 출처다 — 여기 옮겨 적지 않는다.
  // 표는 두 경로(수동 손질 · 자동 추출)가 같이 쓴다 — 한 번 받아 둘 다 채운다.
  const K_PICKERS = [
    { sel: "ed-k-preset", num: "ed-k", why: "ed-k-why" },
    { sel: "au-k-preset", num: "au-k", why: "au-k-why" },
  ];

  async function loadRefCounts() {
    let rows = [];
    try {
      const d = await api("/api/module-f/worst/reference-counts");
      rows = d.rows || [];
      if (d.message) say(d.message, "warn");
    } catch (err) { rows = []; }
    S.refCounts = rows;
    const html = '<option value="">— 직접 입력 —</option>'
      + rows.map((r, i) =>
          `<option value="${i}">${r.count}개 · ${r.label}</option>`).join("");
    for (const p of K_PICKERS) {
      const el = $(p.sel);
      if (el) el.innerHTML = html;
    }
  }

  for (const p of K_PICKERS) {
    const sel = $(p.sel);
    if (!sel) continue;
    sel.onchange = () => {
      const i = sel.value;
      if (i === "") { $(p.why).textContent = ""; return; }
      const row = S.refCounts[Number(i)];
      if (!row) return;
      $(p.num).value = row.count;
      $(p.why).innerHTML =
        `<b>${row.count}개</b> — ${row.label} <span class="tag">${row.rule_id}</span>`;
    };
    $(p.num).oninput = () => { sel.value = ""; };
  }

  // 표시 전용 토글 둘 — 서버에 아무것도 안 보낸다(세션 상태 불변).
  $("ed-worst-view").onchange = () => draw();
  $("ed-bg").onchange = () => draw();

  // ── 최불리 후보 영역 (모듈 A 의 zones) ─────────────────────────
  // 도면 장 나누기는 «자동으로 잰 경계» 라 실무에서 늘 맞지는 않는다. 한 층에
  // 방화구획이 여럿이거나 계산에서 빼야 할 구역이 섞이면 앵커가 그리로 튄다 —
  // 그때는 사람이 직접 가두는 수밖에 없다.
  function drawZones() {
    const live = zoneDrag
      ? [[Math.min(zoneDrag.x0, zoneDrag.x1), Math.min(zoneDrag.y0, zoneDrag.y1),
          Math.max(zoneDrag.x0, zoneDrag.x1), Math.max(zoneDrag.y0, zoneDrag.y1)]]
      : [];
    const all = S.zones.concat(live);
    if (!all.length) return;
    ctx.save();
    ctx.lineWidth = 1.5;
    for (let i = 0; i < all.length; i++) {
      const [x0, y0, x1, y1] = all[i];
      const px = sx(x0), py = sy(y1);
      const w = (x1 - x0) * S.view.scale, h = (y1 - y0) * S.view.scale;
      const live_i = i >= S.zones.length;
      ctx.strokeStyle = live_i ? "#facc15" : "#38bdf8";
      ctx.setLineDash(live_i ? [6, 4] : []);
      ctx.fillStyle = live_i ? "rgba(250,204,21,.10)" : "rgba(56,189,248,.10)";
      ctx.fillRect(px, py, w, h);
      ctx.strokeRect(px, py, w, h);
      if (!live_i) {
        ctx.fillStyle = "#38bdf8";
        ctx.font = "11px sans-serif";
        ctx.fillText(`영역 ${i + 1}`, px + 4, py + 13);
      }
    }
    ctx.restore();
    ctx.setLineDash([]);
  }

  function renderZones() {
    const box = $("ed-zones");
    if (!S.zones.length) { box.textContent = "영역 없음 · 도면 전체"; return; }
    let html = "";
    for (let i = 0; i < S.zones.length; i++) {
      const [x0, y0, x1, y1] = S.zones[i];
      html += kv(`영역 ${i + 1}`,
        `${(x1 - x0).toFixed(0)} × ${(y1 - y0).toFixed(0)} mm`);
    }
    box.innerHTML = html;
  }

  $("ed-zone-undo").onclick = () => { S.zones.pop(); renderZones(); draw(); };
  $("ed-zone-clear").onclick = () => { S.zones = []; renderZones(); draw(); };
  $("ed-zone-arm").onchange = () => {
    say($("ed-zone-arm").checked
      ? "캔버스를 드래그해 영역을 그리세요. (화면 이동은 Shift+드래그)"
      : "영역 그리기를 껐습니다.");
  };

  // ── 도면 슬롯 [H-0] 특허 S650 ──────────────────────────────────
  // 평면도 하나로 끝나지 않는다 — 계통도·기계실에 같은 절차를 반복하고
  // S700 이 셋을 결합한다. 여기서는 슬롯을 켜고 바꾸는 것까지가 전부다.
  function renderSlots(d) {
    const box = $("slots");
    box.innerHTML = "";
    if (!d || !d.slots) {
      const n = document.createElement("span");
      n.className = "note";
      n.textContent = "도면을 열면 슬롯이 켜집니다";
      box.appendChild(n);
      return;
    }
    // 슬롯이 바뀌면 밟는 단계도 바뀐다 — 단계바를 다시 그린다.
    const changed = S.slot !== d.active;
    S.slot = d.active;
    if (changed) renderSteps();
    for (const it of d.slots) {
      const b = document.createElement("button");
      b.className = (it.active ? "on " : "") + (it.opened ? "has" : "");
      const dot = document.createElement("span");
      dot.className = "dot";
      b.appendChild(dot);
      b.appendChild(document.createTextNode(
        it.key ? `${it.label} · ${it.key}` : it.label));
      b.onclick = () => switchSlot(it.kind);
      box.appendChild(b);
    }
    const note = document.createElement("span");
    note.className = "note";
    const n = d.slots.filter((s) => s.opened).length;
    note.textContent = `${n}/3 열림 — 계통도·기계실은 선택입니다`;
    box.appendChild(note);
  }

  async function loadSlots() {
    if (!S.sid) { renderSlots(null); return; }
    try { renderSlots(await api(`/api/module-f/slot/state?sid=${S.sid}`)); }
    catch (err) { renderSlots(null); }
  }

  async function switchSlot(kind) {
    if (!S.sid || kind === S.slot) return;
    busy(true, "도면 바꾸는 중…");
    try {
      const d = await post("/api/module-f/slot/switch", { sid: S.sid, kind });
      S.slot = kind;          // 단계바 흐름이 곧바로 이 슬롯 것이어야 한다
      renderSlots(d);
      const cur = d.slots.find((s) => s.active);
      // 슬롯마다 어디까지 갔는지가 다르다 — 그 단계로 되돌려 놓는다.
      S.sub = { picks: [null, null], arm: null, summary: null };
      S.zones = []; S.autoHeads = []; S.autoAlarm = null; S.autoDone = false;
      // 슬롯이 바뀌면 되돌릴 대상도 바뀐다 — 기록을 넘기면 남의 좌표가 온다.
      S.undo = []; S.autoView = null;
      S.autoNet = null; S.autoNetView = null;
      S.emode = null;                 // [F-10b] 슬롯마다 손질 모드도 새로

      // 방식도 슬롯에 딸린다 — 자동으로 연 평면도로 돌아오면 자동 화면이어야
      // 하고, 아직 안 연 슬롯이면 «안 고른» 상태 그대로여야 한다.
      const st = await api(`/api/module-f/auto/state?sid=${S.sid}`);
      S.method = st.method || null;
      if (!cur.opened) {
        S.world = null; S.edit = null; S.key = null;
        setStage("open");
        say(`${cur.label} — 아직 도면이 없습니다. DXF 를 여세요.`);
      } else if (kind === "plan" && !S.method) {
        // [F-10a] 읽어는 뒀는데 아직 길이 안 정해진 슬롯. 예전에는 여기서 방식을
        //   다시 물었다 — 이제 묻지 않고 열기 때와 같은 판단으로 흘려보낸다
        //   (새로고침 같은 이유로 흐름이 중간에 끊겼을 때 오는 자리다).
        await loadWorldRaw();
        fit(S.world.bounds);
        const nm = st.dxf_name || cur.key || "";
        $("adv-file").textContent = nm;
        $("adv-file").title = nm;
        await loadRecon();
        await autoStart();
      } else if (kind !== "plan") {
        // 계통도·기계실은 찍기·손질을 거치지 않는다 — 두 점 찍기로 바로 간다.
        await loadWorldRaw();
        await loadSub();
      } else if (S.method === "auto") {
        await loadWorldRaw();
        await loadAuto();
      } else if (cur.stage === "edit") { await loadEdit(); }
      else { await loadWorld(); }
    } catch (err) { say(err.message, "err"); }
    finally { busy(false); }
  }

  // ── [H-2 · H-3] 계통도 · 기계실 — 두 점을 찍어 경로를 뽑는다 ──────
  // 계통도는 펌프↔알람밸브, 기계실은 수원↔입상관 연결점. 두 점은 **사람이
  // 찍는다** — 도면마다 기호가 달라 자동 탐지가 조용히 틀리면 경로가 통째로
  // 다른 곳으로 간다(특허 S220 의 우선순위에서 «사용자 지정» 만 쓴다).
  const SUB_SPEC = {
    system: {
      title: "계통도", a: "펌프", b: "알람밸브",
      path: "/api/module-f/system/extract",
      keys: ["pump_x", "pump_y", "av_x", "av_y"],
      clean: true, ceiling: false,
    },
    machineroom: {
      title: "기계실", a: "수원(탱크 토출구)", b: "입상관 연결점",
      path: "/api/module-f/machineroom/extract",
      keys: ["source_x", "source_y", "conn_x", "conn_y"],
      clean: false, ceiling: true,
    },
  };

  function subSpec() { return SUB_SPEC[S.slot] || SUB_SPEC.system; }

  function renderSubPanel() {
    const sp = subSpec();
    $("sub-title").textContent = sp.title;
    $("sub-lab-a").textContent = sp.a;
    $("sub-lab-b").textContent = sp.b;
    $("sub-clean-row").classList.toggle("hidden", !sp.clean);
    $("sub-ceiling-row").classList.toggle("hidden", !sp.ceiling);
    renderSubPicks();
  }

  function renderSubPicks() {
    const sp = subSpec(), p = S.sub.picks;
    const fmt = (q) => q ? `${q[0].toFixed(0)}, ${q[1].toFixed(0)}` : "—";
    $("sub-picks").innerHTML =
      kv(`① ${sp.a}`, fmt(p[0])) + kv(`② ${sp.b}`, fmt(p[1]))
      + (S.sub.arm != null
         ? kv("지금", `<span class="warn">${S.sub.arm === 0 ? sp.a : sp.b}</span> 를 캔버스에서 찍으세요`)
         : "");
  }

  function subClick(x, y) {
    if (S.sub.arm == null) return;
    markUndo(`${S.sub.arm === 0 ? subSpec().a : subSpec().b} 찍기`);
    S.sub.picks[S.sub.arm] = [x, y];
    S.sub.arm = null;
    renderSubPicks();
    draw();
  }

  function armSub(i) {
    S.sub.arm = i;
    renderSubPicks();
    say(`${i === 0 ? subSpec().a : subSpec().b} 위치를 캔버스에서 찍으세요.`);
  }

  async function subExtract(clean) {
    const sp = subSpec();
    const body = { sid: S.sid,
                   snap_tolerance_mm: Number($("sub-snap").value || 2500) };
    if (clean) {
      body.clean = true;
    } else {
      if (!S.sub.picks[0] || !S.sub.picks[1]) {
        say("두 곳을 모두 찍어야 경로를 뽑을 수 있습니다.", "warn");
        return;
      }
      body[sp.keys[0]] = S.sub.picks[0][0];
      body[sp.keys[1]] = S.sub.picks[0][1];
      body[sp.keys[2]] = S.sub.picks[1][0];
      body[sp.keys[3]] = S.sub.picks[1][1];
      if (sp.ceiling) {
        const c = $("sub-ceiling").value;
        if (c !== "") body.ceiling_m = Number(c);
      }
    }
    busy(true, `${sp.title} 경로 추출 중…`);
    try {
      const d = await post(sp.path, body);
      S.sub.summary = d.summary;
      renderSubSummary(d);
      loadSlots();
      say(`${sp.title} 추출 완료 — 절점 ${d.summary.nodes} · 배관 ${d.summary.pipes}`
        + ` · 연장 ${d.summary.total_m} m`, "ok");
    } catch (err) {
      // 미도달을 성공으로 위장하지 않는다 — 특허 S340 의 규범이다.
      S.sub.summary = null;
      $("sub-summary").innerHTML =
        `<span class="err">추출 실패 — ${err.message}</span>`;
      say(err.message, "err");
    } finally { busy(false); }
  }

  function renderSubSummary(d) {
    const s = (d && d.summary) || null;
    if (!s) { $("sub-summary").textContent = "—"; return; }
    let html = kv("절점 / 배관", `${s.nodes} / ${s.pipes}`)
      + kv("연장", `${s.total_m} m`);
    if (s.av_node_label) html += kv("알람밸브 절점", s.av_node_label);
    if (s.source_node_label) html += kv("수원 절점", s.source_node_label);
    if (s.conn_node_label) html += kv("연결 절점", s.conn_node_label);
    if (d.mode) html += kv("방식", d.mode === "clean_network"
                           ? "깨끗한 배관망 통째" : "두 점 최단경로");
    if (s.bridges) html += kv('<span class="warn">추측 연결</span>',
                              `${s.bridges}곳 — 실측이 아닙니다`);
    if (s.elevation_unresolved) html += kv('<span class="warn">표고</span>',
                                           "천장고 미입력 — 첫 구간 미확정");
    $("sub-summary").innerHTML = html;
  }

  function drawSubPicks() {
    const sp = subSpec();
    const marks = [[S.sub.picks[0], "#3b82f6", sp.a],
                   [S.sub.picks[1], "#22c55e", sp.b]];
    ctx.lineWidth = 2;
    for (const [q, color, label] of marks) {
      if (!q) continue;
      const px = sx(q[0]), py = sy(q[1]);
      ctx.strokeStyle = color;
      ctx.beginPath(); ctx.arc(px, py, 10, 0, Math.PI * 2); ctx.stroke();
      ctx.beginPath();
      ctx.moveTo(px - 15, py); ctx.lineTo(px + 15, py);
      ctx.moveTo(px, py - 15); ctx.lineTo(px, py + 15);
      ctx.stroke();
      ctx.fillStyle = color;
      ctx.font = "11px sans-serif";
      ctx.fillText(label, px + 14, py - 8);
    }
    ctx.lineWidth = 1;
  }

  async function loadSub() {
    const d = await api(`/api/module-f/sub/state?sid=${S.sid}`);
    setStage("sub");
    renderSubPanel();
    renderSubSummary(d);
    if (S.world) fit(S.world.bounds);
    draw();
  }

  // ── [A 방식] 자동 추출 — 알람밸브 한 점 + 헤드 영역 ────────────────
  // 수동(E)이 색으로 배관·헤드를 직접 찍는 길이라면, 여기는 두 가지만 정하고
  // 헤드 검출·그래프 복원·앵커·최불리를 모듈 A 에 맡긴다.
  function renderAuto(d) {
    const a = (d && d.alarm) || S.autoAlarm;
    $("au-anchor-info").innerHTML = a
      ? kv("알람밸브", `${a[0].toFixed(0)}, ${a[1].toFixed(0)}`)
      : "—";
    $("au-zones").textContent = S.zones.length
      ? `영역 ${S.zones.length}곳`
      : "영역 없음 · 도면 전체";
    $("au-heads-info").innerHTML = S.autoHeads.length
      ? kv("검출된 헤드", `<span class="ok">${S.autoHeads.length.toLocaleString()}개</span>`
           + (S.zones.length ? ` · 영역 ${S.zones.length}곳 안` : " · 도면 전체"))
      : "—";
    // 어느 단계까지 왔는지 눈에 보이게 — 번호 칩이 초록으로 바뀐다. 순서가
    // 섞여 보인다는 지적을 받아, 「무엇을 이미 했나」를 화면이 직접 말한다.
    const mark = (id, on, txt) => {
      $(id).classList.toggle("done", !!on);
      $(`${id}-mark`).textContent = on ? (txt || "✓") : "";
    };
    mark("au-s1", !!a, "✓ 찍음");
    mark("au-s2", S.autoHeads.length > 0,
         S.autoHeads.length ? `✓ ${S.autoHeads.length.toLocaleString()}개` : "");
    const ns = S.autoNet;
    mark("au-s3", !!ns, ns ? `✓ 도달 ${ns.reached.toLocaleString()}` : "");
    // 범위는 «선택» 이다 — 안 그렸다고 미완으로 보이면 안 된다. 그린 곳이
    // 있으면 개수를, 없으면 «도면 전체» 임을 그 자리에서 말한다.
    mark("au-s4", S.zones.length > 0,
         S.zones.length ? `✓ ${S.zones.length}곳` : "");
    if (!S.zones.length) $("au-s4-mark").textContent = "도면 전체";
    mark("au-s5", S.autoDone, "✓ 추출됨");
    $("au-network").disabled = !a;
    renderAutoNet();
    renderJunctions();
    renderPipeLayers();
    $("au-zone-draw").classList.toggle("on", $("au-zone-arm").checked);
    $("au-zone-undo").disabled = !S.zones.length;
    $("au-zone-clear").disabled = !S.zones.length;
    // 영역은 «좁히는» 선택이다 — 알람밸브만 있으면 돌릴 수 있다.
    $("au-run").disabled = !a;
    $("au-heads").disabled = !a;
    $("au-to-design").disabled = !S.autoDone;
    // [F-8d] 이어받기는 «자동 결과를 본 뒤» 의 길이다 — 돌리기 전에는 뜻이 없다.
    $("au-handoff").disabled = !S.autoDone;
    if (d && d.summary) {
      S.autoSummary = d.summary;
      renderAutoSummary(d.summary);
    }
  }

  function renderAutoSummary(s) {
    if (!s) { $("au-summary").textContent = "—"; return; }
    let html = kv("설계면적 헤드", `<span class="ok">${s.k}개</span>`)
      + kv("최원 / 최근", `${s.far_m} m / ${s.near_m} m`)
      + kv("배관망", `절점 ${s.nodes} · 배관 ${s.pipes} · 노즐 ${s.nozzles}`
           + ` · 부속 ${s.fittings}`)
      + kv("범위", s.region_auto
           ? (s.sheet
              // 한 파일에 도면이 여러 장이면 알람밸브가 놓인 장으로 좁힌다 —
              // 전부를 범위로 잡으면 다른 장의 헤드까지 후보가 된다.
              ? `${s.sheet} <span class="dim">(알람밸브가 놓인 장)</span>`
              : "도면 전체 (검출된 헤드에서 자동)")
           : `영역 ${s.zones}곳`);
    // 급수원이 그래프에서 멀어 갈아탄 것은 숨기지 않는다.
    if (s.source_fallback) {
      html += kv('<span class="warn">급수원 대체</span>',
                 `${s.source_bridge_mm} mm 떨어져 최근접 절점으로 — `
                 + `알람밸브 위치를 확인하세요`);
    }
    $("au-summary").innerHTML = html;
  }

  // 뽑아낸 망을 받아 둔다 — 화면이 «무엇이 뽑혔나» 를 그릴 수 있어야 한다.
  // 실패해도 자동 경로는 그대로 돈다(그림이 없을 뿐이다).
  async function loadAutoView() {
    S.autoView = null;
    if (!S.autoDone) return;
    try {
      const d = await api(`/api/module-f/auto/preview?sid=${S.sid}`);
      S.autoView = d.view || null;
    } catch (err) { S.autoView = null; }
  }

  async function loadAuto() {
    setStage("auto");
    try {
      const d = await api(`/api/module-f/auto/state?sid=${S.sid}`);
      S.autoDone = !!d.done;
      if (d.alarm) S.autoAlarm = d.alarm;
      // 영역은 서버가 들고 있다 — 슬롯을 오갔다 와도 그대로 되살린다.
      if (Array.isArray(d.zones)) S.zones = d.zones.map((z) => z.slice());
      // 「배관으로 취급」 지정도 서버가 들고 있다 — 색·이름은 도면에서 되찾는다.
      if (Array.isArray(d.pipe_layers)) {
        const by = new Map((S.world ? S.world.bundles : [])
          .map((b) => [`${b.layer}|${b.color}`, b]));
        S.autoPipe = d.pipe_layers.map((p) => {
          const b = by.get(`${p.layer}|${p.color}`);
          return { id: b ? b.id : `${p.layer}${p.color}`, layer: p.layer,
                   color: p.color, css: b ? b.css : "#94a3b8",
                   name: b ? b.name : "" };
        });
      }
      await loadAutoNetView();   // 검출한 망을 먼저 되살린다(단계 표시가 쓴다)
      renderAuto(d);
      await loadAutoView();
      // 추출을 끝낸 슬롯으로 돌아오면 그 자리로 다시 맞춘다.
      const b = autoNetBounds();
      if (b) fit(b);
      else if (S.world) fit(S.world.bounds);
      draw();
    } catch (err) { say(err.message, "err"); }
  }

  function autoClick(x, y) {
    if (S.autoArm === "pipe") {
      S.autoArm = null;
      $("au-pipe-pick").classList.remove("on");
      const b = bundleAt(x, y, PICK_PX * 3 / S.view.scale);
      if (!b) { say("그 자리에서 선을 찾지 못했습니다.", "warn"); return; }
      S.autoPipe = S.autoPipe || [];
      if (S.autoPipe.some((q) => q.id === b.id)) {
        say(`이미 지정한 묶음입니다 — ${b.layer}`, "warn");
        return;
      }
      markUndo(`배관 지정 — ${b.layer}`);
      S.autoPipe.push({ id: b.id, layer: b.layer, color: b.color,
                        css: b.css, name: b.name });
      say(`${b.layer} 를 배관으로 취급합니다 — 헤드 검출부터 다시 하세요.`,
          "ok");
      pushPipeLayers();
      return;
    }
    if (S.autoArm !== "anchor") return;
    S.autoArm = null;
    $("au-anchor").classList.remove("on");
    markUndo("알람밸브 찍기");
    S.autoAlarm = [x, y];
    post("/api/module-f/auto/anchor", { sid: S.sid, x, y })
      .then(() => { renderAuto(null); draw(); })
      .catch((err) => say(err.message, "err"));
  }

  async function pushAutoZones() {
    try { await post("/api/module-f/auto/zones", { sid: S.sid, zones: S.zones }); }
    catch (err) { say(err.message, "err"); }
    renderAuto(null);
  }

  $("au-anchor").onclick = () => {
    S.autoArm = "anchor";
    // 단추가 «눌린 상태» 로 보여야 «지금 캔버스를 찍으면 된다» 가 읽힌다.
    $("au-anchor").classList.add("on");
    say("도면에서 알람밸브 자리를 클릭하세요 — 여기서 출발해 최불리 헤드군을 "
        + "찾습니다.");
  };
  $("au-anchor-clear").onclick = async () => {
    markUndo("알람밸브 지우기");
    S.autoAlarm = null; S.autoArm = null;
    $("au-anchor").classList.remove("on");
    try { await post("/api/module-f/auto/anchor", { sid: S.sid }); }
    catch (err) { say(err.message, "err"); }
    renderAuto(null); draw();
  };
  $("au-zone-undo").onclick = () => {
    if (!S.zones.length) return;
    markUndo("마지막 영역 지우기");
    S.zones.pop(); pushAutoZones(); draw();
  };
  $("au-zone-clear").onclick = () => {
    if (!S.zones.length) return;
    markUndo(`영역 ${S.zones.length}곳 전부 지우기`);
    S.zones = []; pushAutoZones(); draw();
  };
  // 단추가 무장 상태를 켜고 끈다 — 체크박스는 그 상태의 원천으로만 남는다
  // (캔버스 드래그 판정이 그 값을 읽는다).
  $("au-zone-draw").onclick = () => {
    const on = !$("au-zone-arm").checked;
    $("au-zone-arm").checked = on;
    $("au-zone-draw").classList.toggle("on", on);
    say(on ? "캔버스를 드래그해 뽑을 구역을 그리세요. (화면 이동은 Shift+드래그)"
           : "영역 그리기를 껐습니다.");
  };

  // ── [S270 · S310] 배관망 검출 ────────────────────────────────
  // 최불리는 «거리를 내림차순으로 자른 것» 이다. 그 거리를 어디서 어떻게 재는지
  // 가 안 보이면 사람은 결과만 받고 믿을 근거가 없다.
  function renderAutoNet() {
    const s = S.autoNet;
    if (!s) { $("au-net-info").textContent = "—"; return; }
    let html = kv("배관망", `절점 ${s.nodes.toLocaleString()} · `
                  + `배관 ${s.pipes.toLocaleString()} · ${s.len_m} m`)
      + kv("도달 헤드", `<span class="ok">${s.reached.toLocaleString()}</span>`
           + ` / ${s.detected.toLocaleString()}`
           + (s.unreached
              ? ` · <span class="warn">미도달 ${s.unreached.toLocaleString()}</span>`
              : ""))
      + kv("거리 (밸브→헤드)",
           `최근 ${s.near_m} · 중앙 ${s.mid_m} · <b>최원 ${s.far_m}</b> m`);
    if (s.pruned && s.cut_pipes) {
      html += kv("잘라낸 관로",
                 `${s.cut_pipes.toLocaleString()}개 · ${s.cut_m} m `
                 + `<span class="dim">(물 안 가는 배관)</span>`);
    }
    if (s.fragments) {
      html += kv('<span class="warn">미연결 조각</span>',
                 `${s.fragments.toLocaleString()}개 · ${s.frag_m} m`);
    }
    $("au-net-info").innerHTML = html;
  }

  async function loadAutoNetView() {
    S.autoNet = null; S.autoNetView = null;
    try {
      const d = await api(`/api/module-f/auto/network-view?sid=${S.sid}`);
      S.autoNet = d.summary || null;
      S.autoNetView = d.view || null;
    } catch (err) { /* 아직 안 돌렸다 — 빈 채로 둔다 */ }
  }

  $("au-network").onclick = async () => {
    busy(true, "배관망을 잇고 거리를 재는 중…");
    try {
      await post("/api/module-f/auto/network",
                 { sid: S.sid, prune: $("au-prune").checked });
      watch(async () => {
        await loadAutoNetView();
        renderAuto(null);
        if (S.autoNetView) fit(netBounds() || curBounds());
        draw();
        const s = S.autoNet;
        if (s) {
          say(`배관망 검출 완료 — 도달 헤드 ${s.reached.toLocaleString()}`
            + `/${s.detected.toLocaleString()} · 최원 ${s.far_m} m`
            + (s.cut_pipes ? ` · 물 안 가는 관로 ${s.cut_pipes}개 잘라냄` : ""),
              s.unreached ? "warn" : "ok");
        }
      });
    } catch (err) { busy(false); say(err.message, "err"); }
  };

  function netBounds() {
    const sg = S.autoNetView && S.autoNetView.segs;
    if (!sg || !sg.length) return null;
    let x0 = Infinity, y0 = Infinity, x1 = -Infinity, y1 = -Infinity;
    for (let i = 0; i < sg.length; i += 2) {
      if (sg[i] < x0) x0 = sg[i];
      if (sg[i] > x1) x1 = sg[i];
      if (sg[i + 1] < y0) y0 = sg[i + 1];
      if (sg[i + 1] > y1) y1 = sg[i + 1];
    }
    const pad = Math.max(x1 - x0, y1 - y0, 1000) * 0.06;
    return { minx: x0 - pad, miny: y0 - pad, maxx: x1 + pad, maxy: y1 + pad };
  }

  // 검출한 망 — 최불리(청록)와 구분되게 파랑으로 얇게 깐다.
  function drawAutoNetwork() {
    const v = S.autoNetView;
    if (!v || !v.segs || !v.segs.length) return;
    ctx.strokeStyle = "#60a5fa";
    ctx.lineWidth = 1.4;
    ctx.beginPath();
    for (let i = 0; i < v.segs.length; i += 4) {
      ctx.moveTo(sx(v.segs[i]), sy(v.segs[i + 1]));
      ctx.lineTo(sx(v.segs[i + 2]), sy(v.segs[i + 3]));
    }
    ctx.stroke();
    ctx.lineWidth = 1;
    drawJunctions(v);
  }

  // ── 이음자리 — «티» 와 «그냥 교차» ──────────────────────────────
  // 둘을 같은 크기·같은 자리에 그리되 «채움» 으로 가른다. 모양을 아주 다르게
  // 하면 비교가 안 되고, 같게 하면 구분이 안 된다.
  //
  //   티(분기)  ● 채운 원        물이 갈라진다 — 부속(티)이 서는 자리
  //   교차      ○ 빈 원 + 사선   평면 좌표로는 티인지 스쳐 지나감인지 못 가린다
  //
  // 「추정은 점선」이라는 저장소 규약을 여기에도 그대로 쓴다 — 확정 못한 쪽이
  // 점선이다.
  const JUNC_TEE = "#f59e0b";     // 분기 — 눈에 띄어야 한다
  const JUNC_X = "#94a3b8";       // 교차 — 판단 보류라 조용히

  // 무엇을 보고 있는지 숫자와 범례로 함께 말한다 — 색만으로는 못 읽는다.
  function renderJunctions() {
    const v = S.autoView || S.autoNetView;
    const box = $("au-junc-info");
    if (!v || (!v.tees && !v.crosses)) { box.innerHTML = ""; return; }
    const t = (v.tees || []).length, x = (v.crosses || []).length;
    box.innerHTML =
      `<span style="color:${JUNC_TEE}">●</span> 분기(티) <b>${t}</b>곳`
      + ` &nbsp; <span style="color:${JUNC_X}">◌</span> 교차 <b>${x}</b>곳`
      + (x ? `<br><span class="dim">교차는 평면 좌표만으로 티인지 층이 달라 `
             + `스쳐 지나가는지 못 가립니다 — 자르지 않고 표시만 합니다.</span>`
           : "");
  }

  $("au-junc").onchange = (e) => { S.showJunc = e.target.checked; draw(); };

  function drawJunctions(v) {
    if (!v || !S.showJunc) return;
    const R = 4.5;
    for (const p of (v.tees || [])) {
      ctx.beginPath();
      ctx.arc(sx(p[0]), sy(p[1]), R, 0, Math.PI * 2);
      ctx.fillStyle = JUNC_TEE;
      ctx.fill();
    }
    ctx.lineWidth = 1.4;
    ctx.setLineDash([2, 2]);
    for (const p of (v.crosses || [])) {
      const px = sx(p[0]), py = sy(p[1]);
      ctx.beginPath();
      ctx.arc(px, py, R, 0, Math.PI * 2);
      ctx.strokeStyle = JUNC_X;
      ctx.stroke();
      ctx.setLineDash([]);
      ctx.beginPath();                    // 사선 하나 — «가리지 못함» 의 표
      ctx.moveTo(px - R * 0.7, py + R * 0.7);
      ctx.lineTo(px + R * 0.7, py - R * 0.7);
      ctx.stroke();
      ctx.setLineDash([2, 2]);
    }
    ctx.setLineDash([]);
    ctx.lineWidth = 1;
  }

  // ── 「이 레이어를 배관으로 취급」 ─────────────────────────────
  // 수동 차선은 색으로 찍어 재료를 확정한다. 자동에는 그 길이 없어, 이름 사전이
  // OTHER 로 떨어뜨린 선은 손댈 방법이 아예 없었다(실측 B1F `현장조사#셔터`).
  // 지문 규칙을 늘리는 대신 사람이 찍게 한다 — 「선을 따라 헤드가 정렬」 지문은
  // 건축선(A-B1)에 28줄이 걸려 벽을 배관으로 먹는다.
  function segDist2(px, py, x0, y0, x1, y1) {
    const dx = x1 - x0, dy = y1 - y0;
    const L2 = dx * dx + dy * dy;
    let t = L2 > 0 ? ((px - x0) * dx + (py - y0) * dy) / L2 : 0;
    t = Math.max(0, Math.min(1, t));
    const cx = x0 + t * dx, cy = y0 + t * dy;
    return (px - cx) ** 2 + (py - cy) ** 2;
  }

  // 클릭 자리의 «레이어×색» 묶음 — 화면이 그리는 그 단위 그대로다.
  function bundleAt(x, y, maxD) {
    if (!S.world) return null;
    let best = null, bd = maxD * maxD;
    for (const b of S.world.bundles) {
      if (S.hidden.has(b.id)) continue;
      const sg = b.segs;
      for (let i = 0; i < sg.length; i += 4) {
        const d = segDist2(x, y, sg[i], sg[i + 1], sg[i + 2], sg[i + 3]);
        if (d < bd) { bd = d; best = b; }
      }
    }
    return best;
  }

  function renderPipeLayers() {
    const ls = S.autoPipe || [];
    const box = $("au-pipe-info");
    $("au-pipe-clear").disabled = !ls.length;
    if (!ls.length) {
      box.innerHTML = '<span class="dim">지정 없음 — 레이어 이름 사전이 '
        + "고른 배관만 씁니다.</span>";
      return;
    }
    box.innerHTML = ls.map((b) =>
      `<div class="kv"><b><i style="display:inline-block;width:9px;`
      + `height:9px;background:${b.css};margin-right:6px"></i>${b.layer}</b>`
      + `<span>${b.name || ""}</span></div>`).join("")
      + `<div class="dim" style="margin-top:5px">이 묶음을 배관으로 올려 `
      + `추출합니다.</div>`;
  }

  async function pushPipeLayers() {
    try {
      await post("/api/module-f/auto/pipe-layers", {
        sid: S.sid,
        layers: (S.autoPipe || []).map((b) => ({ layer: b.layer,
                                                 color: b.color })),
      });
    } catch (err) { say(err.message, "err"); return; }
    // 지정이 바뀌면 앞서 뽑은 것은 «다른 도면» 의 결과다 — 서버가 지웠으니
    // 화면도 같이 비운다.
    S.autoHeads = []; S.autoDone = false;
    S.autoNet = null; S.autoNetView = null; S.autoView = null;
    renderPipeLayers();
    renderAuto(null);
    draw();
  }

  $("au-pipe-pick").onclick = () => {
    S.autoArm = "pipe";
    $("au-pipe-pick").classList.add("on");
    say("배관으로 취급할 선을 도면에서 클릭하세요 — 그 레이어×색 묶음이 "
        + "통째로 배관이 됩니다.");
  };

  $("au-pipe-clear").onclick = () => {
    if (!(S.autoPipe || []).length) return;
    markUndo("배관 지정 지우기");
    S.autoPipe = [];
    pushPipeLayers();
  };

  $("au-heads").onclick = async () => {
    busy(true, "헤드 후보 찾는 중…");
    try {
      const d = await post("/api/module-f/auto/heads", { sid: S.sid });
      S.autoHeads = d.heads || [];
      renderAuto(null);
      let msg = d.n
        ? `헤드 ${d.n.toLocaleString()}개를 찾았습니다 — 그대로 추출하거나, `
          + `「범위 좁히기」로 구역을 지정하세요.`
        : "헤드를 찾지 못했습니다 — 알람밸브 위치나 도면 레이어를 확인하세요.";
      if (d.dropped) {
        msg += ` (화면에는 ${d.dropped.toLocaleString()}개를 뺐습니다 — 표시`
             + ` 상한. 추출은 전량 대상입니다.)`;
      }
      say(msg, d.n ? (d.dropped ? "warn" : "ok") : "warn");
      draw();
    } catch (err) { say(err.message, "err"); }
    finally { busy(false); }
  };

  $("au-run").onclick = async () => {
    const k = Math.max(1, Math.min(200, Number($("au-k").value || 30)));
    busy(true, "자동 추출 중…");
    try {
      await post("/api/module-f/auto/run", { sid: S.sid, k });
      watch(async () => {
        const d = await api(`/api/module-f/auto/state?sid=${S.sid}`);
        S.autoDone = !!d.done;
        renderAuto(d);
        $("dg-k").value = k;
        renderSteps();
        // 뽑힌 망을 받아 와야 나머지를 내리고 이것만 살릴 수 있다.
        await loadAutoView();
        // ★흐리게 내리는 것만으로는 안 드러난다 — 도면이 971 m 인데 설계면적은
        //   25 m 라, 화면을 도면 전체로 두면 결과가 점 하나로 남는다. 뽑은
        //   자리로 맞춰 준다(사람이 다시 「화면 맞춤」을 눌러도 여기로 온다).
        const b = autoNetBounds();
        if (b) fit(b); else draw();
        if (d.summary) {
          say(`자동 추출 완료 — 헤드 ${d.summary.k} · 최원 ${d.summary.far_m} m`
              + ` · 절점 ${d.summary.nodes}. 뽑은 자리로 화면을 맞추고 나머지`
              + " 도면은 흐리게 내렸습니다.", "ok");
        }
      });
    } catch (err) { busy(false); say(err.message, "err"); }
  };

  $("au-to-design").onclick = async () => {
    setStage("design");
    try { await designPreview(); }
    catch (err) { say(err.message, "err"); }
  };

  // 뽑아낸 배관망 — 도면을 내린 위에 이것만 밝게 얹는다.
  function drawAutoNet() {
    const v = S.autoView;
    if (!v || !v.pipes || !v.pipes.length) return;
    const at = {};
    for (const n of (v.nodes || [])) at[n.label] = n;
    ctx.lineWidth = 2.6;
    ctx.strokeStyle = "#22d3ee";
    ctx.beginPath();
    for (const p of v.pipes) {
      const a = at[p.a], b = at[p.b];
      if (!a || !b) continue;
      ctx.moveTo(sx(a.x), sy(a.y));
      ctx.lineTo(sx(b.x), sy(b.y));
    }
    ctx.stroke();
    // 뽑힌 헤드(노즐)와 급수 절점은 따로 찍는다 — 선만 보면 어디가 말단인지
    // 모른다.
    for (const n of (v.nodes || [])) {
      if (!n.head && !n.input) continue;
      ctx.fillStyle = n.input ? "#3b82f6" : "#22d3ee";
      ctx.beginPath();
      ctx.arc(sx(n.x), sy(n.y), n.input ? 5 : 3.2, 0, Math.PI * 2);
      ctx.fill();
    }
    ctx.lineWidth = 1;
    drawJunctions(v);      // 뽑은 망에도 티/교차를 갈라 표시한다
  }

  // 검출한 헤드는 «빨강» 이다. 신뢰도 색(초록/노랑/회색)은 어두운 도면 위에서
  // 티가 안 나 무엇이 잡혔는지 한눈에 안 들어왔다 — 띠는 옆 패널 숫자로 읽는다.
  const HEAD_MARK = "#ff3b30";

  function drawAuto(dim) {
    // 헤드 후보 — 추출이 끝나면 한 발 물러선다(뽑힌 망이 주인공이다).
    for (const h of (S.autoHeads || [])) {
      ctx.fillStyle = HEAD_MARK;
      ctx.globalAlpha = dim ? 0.3 : 0.9;
      ctx.beginPath();
      ctx.arc(sx(h.x), sy(h.y), dim ? 2.6 : 4, 0, Math.PI * 2);
      ctx.fill();
    }
    ctx.globalAlpha = 1;
    // 알람밸브 = 기준점. 손질 경로의 급수원과 같은 파란 겹원.
    if (S.autoAlarm) {
      const px = sx(S.autoAlarm[0]), py = sy(S.autoAlarm[1]);
      ctx.strokeStyle = "#3b82f6";
      ctx.lineWidth = 2.4;
      ctx.beginPath(); ctx.arc(px, py, 10, 0, Math.PI * 2); ctx.stroke();
      ctx.beginPath();
      ctx.moveTo(px - 15, py); ctx.lineTo(px + 15, py);
      ctx.moveTo(px, py - 15); ctx.lineTo(px, py + 15);
      ctx.stroke();
      ctx.fillStyle = "#3b82f6";
      ctx.font = "11px sans-serif";
      ctx.fillText("알람밸브", px + 14, py - 8);
      ctx.lineWidth = 1;
    }
  }

  // ── [H-4 · H-5 · H-6] 통합 — 특허 제5국면 S700 ────────────────────
  async function loadMergeModes() {
    const d = await api("/api/module-f/merge/modes");
    const box = $("mg-modes");
    box.innerHTML = "";
    for (const m of d.modes) {
      const lb = document.createElement("label");
      lb.className = "chk";
      const rb = document.createElement("input");
      rb.type = "radio"; rb.name = "mg-mode"; rb.value = m.key;
      rb.onchange = () => setMergeMode(m.key);
      const sp = document.createElement("span");
      sp.textContent = m.label;
      lb.appendChild(rb); lb.appendChild(sp);
      box.appendChild(lb);
    }
  }

  async function setMergeMode(key) {
    const body = { sid: S.sid, mode: key };
    const isPump = key === "hsp_pump";
    $("mg-drop-row").classList.toggle("hidden", !isPump);
    if (isPump) body.source_drop_m = Number($("mg-drop").value || 0);
    try {
      const d = await post("/api/module-f/merge/mode", body);
      say(`급수방식: ${d.label}`, "ok");
      await loadMergeState();
    } catch (err) { say(err.message, "err"); }
  }

  async function loadMergeState() {
    const d = await api(`/api/module-f/merge/state?sid=${S.sid}`);
    S.merge = d;
    let html = "";
    for (const k of ["plan", "system", "machineroom"]) {
      const ok = d.ready[k];
      const need = k === "plan";
      html += kv(d.labels[k],
        ok ? '<span class="ok">있음</span>'
           : (need ? '<span class="err">없음 — 필요</span>'
                   : '<span class="dim">없음 — 선택</span>'));
    }
    if (d.mode_label) html += kv("급수방식", d.mode_label);
    $("mg-ready").innerHTML = html;
    $("mg-build").disabled = !d.can_build;
    if (d.mode) {
      const rb = document.querySelector(`input[name=mg-mode][value=${d.mode}]`);
      if (rb) rb.checked = true;
      const isPump = d.mode === "hsp_pump";
      $("mg-drop-row").classList.toggle("hidden", !isPump);
    }
    if (d.summary) renderMergeSummary(d.summary);
    $("mg-emit").disabled = !(d.summary && d.summary.merged);
    return d;
  }

  function renderMergeSummary(s) {
    if (!s) { $("mg-summary").textContent = "—"; return; }
    let html = "";
    if (!s.merged) {
      html += kv('<span class="warn">결합 없음</span>',
                 "계통도가 없어 평면도 단독입니다");
    } else {
      html += kv("절점 / 배관", `${s.nodes} / ${s.pipes}`)
            + kv("노즐", `${s.nozzles}`)
            + kv("펌프 / 밸브", `${s.pumps || 0} / ${s.valves || 0}`)
            + kv("기계실 접속", s.attached
                 ? '<span class="ok">접속됨</span>'
                 : '<span class="warn">미접속</span>');
    }
    // 어느 단계가 실제로 돌았는지 — 「붙였다」고 말하려면 근거가 있어야 한다.
    for (const line of (s.steps || [])) html += kv("·", line);
    $("mg-summary").innerHTML = html;
  }

  $("mg-build").onclick = async () => {
    busy(true, "배관망 결합 중…");
    try {
      await post("/api/module-f/merge/build", { sid: S.sid });
      watch(async () => {
        const d = await loadMergeState();
        if (d.summary && d.summary.merged) {
          say(`결합 완료 — 절점 ${d.summary.nodes} · 배관 ${d.summary.pipes}`, "ok");
        } else {
          say("평면도 단독으로 지나갔습니다 (계통도 없음).", "warn");
        }
      });
    } catch (err) { busy(false); say(err.message, "err"); }
  };

  $("mg-emit").onclick = async () => {
    busy(true, "산출물 생성 중…");
    try {
      await post("/api/module-f/merge/emit", { sid: S.sid });
      watch(async () => {
        $("mg-download").disabled = false;
        say("산출 완료 — zip 으로 내려받으세요.", "ok");
      });
    } catch (err) { busy(false); say(err.message, "err"); }
  };

  $("mg-download").onclick = () => {
    window.location = `/api/module-f/merge/download?sid=${S.sid}&what=zip`;
  };

  $("mg-drop").onchange = () => {
    if (S.merge && S.merge.mode === "hsp_pump") setMergeMode("hsp_pump");
  };

  async function loadMerge() {
    setStage("merge");
    try {
      await loadMergeModes();
      await loadMergeState();
    } catch (err) { say(err.message, "err"); }
  }

  $("dg-to-merge").onclick = () => loadMerge();

  $("sub-pick-a").onclick = () => armSub(0);
  $("sub-pick-b").onclick = () => armSub(1);
  $("sub-clear").onclick = () => {
    if (!S.sub.picks.some(Boolean)) return;
    markUndo("찍은 점 지우기");
    S.sub.picks = [null, null]; S.sub.arm = null;
    renderSubPicks(); draw();
  };
  $("sub-extract").onclick = () => subExtract(false);
  $("sub-clean").onclick = () => subExtract(true);

  async function loadSaved() {
    try {
      const d = await api("/api/module-f/saved");
      const sel = $("saved");
      sel.innerHTML = "";
      if (!d.items.length) {
        sel.innerHTML = '<option value="">— 저장된 찍기 없음 —</option>';
        $("btn-reopen").disabled = true;
        return;
      }
      $("btn-reopen").disabled = false;
      for (const it of d.items) {
        const o = document.createElement("option");
        o.value = it.key;
        o.textContent = `${it.key}  ·  ${it.picked_at}`;
        sel.appendChild(o);
      }
    } catch (err) { say(err.message, "err"); }
  }

  $("btn-reopen").onclick = async () => {
    const key = $("saved").value;
    if (!key) { say("이어서 열 도면이 없습니다.", "warn"); return; }
    busy(true, "배관망 여는 중…");
    try {
      const d = await post("/api/module-f/reopen", { key });
      S.sid = d.sid; S.key = d.key;
      // 저장본은 찍기가 끝난 것이라 언제나 수동 경로다(손질부터 시작).
      S.method = "manual";
      renderSteps();
      watch(loadEdit);
    } catch (err) { busy(false); say(err.message, "err"); }
  };

  // ── 레이어 목록 ────────────────────────────────────────────────
  function buildLayers() {
    const box = $("layers");
    box.innerHTML = "";
    for (const b of S.world.bundles) {
      const lb = document.createElement("label");
      const cb = document.createElement("input");
      cb.type = "checkbox"; cb.checked = true; cb.dataset.id = b.id;
      cb.onchange = () => {
        if (cb.checked) S.hidden.delete(b.id); else S.hidden.add(b.id);
        draw();
      };
      const sw = document.createElement("span");
      sw.className = "sw"; sw.style.background = b.css;
      const ct = document.createElement("span");
      ct.className = `cat ${b.cat}`;
      ct.textContent = b.cat;
      const tx = document.createElement("span");
      tx.className = "nm";
      tx.textContent = `${b.layer} × ${b.name}`;
      // 잰 값만 적는다 — «배관다움» 같은 판정은 **안 붙인다**.
      //
      // ★붙이려다 실측이 막았다. 「긴 선분이 많으면 배관」이 그럴듯해 보이지만
      //   정반대로 작동한다: 평면도의 진짜 배관 레이어는 부속·꺾임 때문에 긴
      //   선분이 10~18% 뿐이고(대명동 `-소화(SP가지관)` 10% · LH306 `pipe`
      //   18%), 계통도의 층 구획선은 100% 다. 판정을 지어내면 이름 사전이
      //   틀린 자리에 «틀린 확신» 을 하나 더 얹게 된다.
      //
      // ★개수는 «도면에 있는 수»(n_all) 를 쓴다. 종전에는 그려 보낸 수를
      //   적었는데 그것은 상한에서 잘린 값이라 큰 도면에서 거짓말이 된다.
      const cn = document.createElement("span");
      cn.className = "cnt";
      const n = (b.n_all !== undefined ? b.n_all : b.n_seg);
      const cir = (b.n_circle_all !== undefined ? b.n_circle_all : b.n_circle);
      cn.textContent = `${num(b.len_m).toLocaleString()} m · ${n.toLocaleString()}`
        + (cir ? ` · ○${cir.toLocaleString()}` : "");
      cn.title = `총 연장 ${num(b.len_m).toLocaleString()} m`
        + ` · 선분 ${n.toLocaleString()}개`
        + ` · 중앙 선분 길이 ${num(b.len_mid).toLocaleString()} mm`
        + (cir ? ` · 원 ${cir.toLocaleString()}개` : "")
        + (b.n_arc_all ? ` · 호 ${b.n_arc_all.toLocaleString()}개` : "")
        + (b.n_seg < n ? `\n(화면에는 ${b.n_seg.toLocaleString()}개만 그립니다`
                         + " — 세는 것과 그리는 것은 다릅니다)" : "");
      lb.append(cb, sw, ct, tx, cn);
      box.appendChild(lb);
    }
  }
  $("ly-all").onclick = () => {
    S.hidden.clear();
    box_all(true); draw();
  };
  $("ly-none").onclick = () => {
    S.hidden = new Set(S.world.bundles.map((b) => b.id));
    box_all(false); draw();
  };
  function box_all(v) {
    for (const cb of $("layers").querySelectorAll("input")) cb.checked = v;
  }

  // ── 2. 찍기 ────────────────────────────────────────────────────
  function renderPick() {
    const p = S.pick;
    // 「상태」 다섯 줄은 뺐다 — 단추가 이미 같은 것을 말한다: 찍기가 켜졌는지는
    // 배관 선택이 눌린 모양으로, 재료가 찼는지는 선택 완료·배관망 구성의
    // 활성 여부로, 무엇을 찍었는지는 캔버스와 아래 상태줄로 보인다.
    for (const b of document.querySelectorAll(".slot")) {
      b.classList.toggle("on", b.dataset.slot === p.head_label && p.mode === "헤드");
    }
    $("pk-pipe").classList.toggle("on", p.mode === "재료" && p.armed);
    $("pk-done").disabled = p.materials.length === 0;
    $("pk-next").disabled = !p.mat_done;
    draw();
  }
  function kv(k, v) {
    return `<div class="kv"><b>${k}</b><span>${v}</span></div>`;
  }

  // ── 모듈 A 레이어 사전 추천 ────────────────────────────────────
  const CAT_ORDER = ["PIPE", "HEAD", "ALARM", "TEXT", "ARCH", "EXCLUDE", "OTHER"];
  function renderCats() {
    const cats = (S.world && S.world.cats) || {};
    $("pk-cats").innerHTML = CAT_ORDER
      .filter((c) => cats[c])
      .map((c) => `<span>${c}<b>${cats[c]}</b></span>`).join("");
    $("pk-auto-pipe").disabled = !cats.PIPE;
    $("pk-auto-head").disabled = !cats.HEAD;
  }

  async function pickAuto(cat) {
    try {
      const d = await post("/api/module-f/pick/auto", { sid: S.sid, cat });
      S.pick = d.state;
      renderPick();
      say(d.message + (d.applied.length ? ` — ${d.applied.join(", ")}` : ""),
          d.applied.length ? "ok" : "warn");
    } catch (err) { say(err.message, "err"); }
  }
  $("pk-auto-pipe").onclick = () => pickAuto("PIPE");
  $("pk-auto-head").onclick = () => pickAuto("HEAD");

  async function pickMode(action, slot) {
    try {
      const d = await post("/api/module-f/pick/mode", { sid: S.sid, action, slot });
      S.pick = d.state;
      renderPick();
      say(d.message, d.applied ? "" : "warn");
    } catch (err) { say(err.message, "err"); }
  }
  $("pk-pipe").onclick = () => pickMode("pipe");
  $("pk-done").onclick = () => pickMode("complete");
  for (const b of document.querySelectorAll(".slot")) {
    b.onclick = () => pickMode("slot", b.dataset.slot);
  }

  async function pickClick(x, y, maxD) {
    try {
      const d = await post("/api/module-f/pick/click",
                           { sid: S.sid, x, y, max_d: maxD });
      S.pick = d.state;
      renderPick();
      if (!d.report) {
        say(S.pick.armed
            ? "아무것도 잡히지 않았습니다. 더 가까이 클릭하세요."
            : "먼저 «배관 선택» 또는 헤드 «칸» 단추를 눌러 찍기를 켜세요.", "warn");
        return;
      }
      const r = d.report;
      say(`${r["모드"]} ${r["동작"]} — ${r["픽"]}`, "ok");
    } catch (err) { say(err.message, "err"); }
  }

  $("pk-undo").onclick = async () => {
    try {
      const d = await post("/api/module-f/pick/undo", { sid: S.sid });
      S.pick = d.state;
      renderPick();
      say(d.undone ? "한 단계 되돌렸습니다." : "되돌릴 클릭이 없습니다.",
          d.undone ? "" : "warn");
    } catch (err) { say(err.message, "err"); }
  };

  $("pk-next").onclick = async () => {
    busy(true, "배관망 구성 중…");
    try {
      await post("/api/module-f/pick/commit", { sid: S.sid });
      watch(loadEdit);
    } catch (err) { busy(false); say(err.message, "err"); }
  };

  // ── 3. 손질 ────────────────────────────────────────────────────
  // 서버는 «망이 바뀌었을 때만» 덩이 도형을 싣는다(B1F 실측 812KB/장).
  // 빈 배열은 «안 바뀜» 이라는 뜻이므로 들고 있던 사본을 그대로 쓴다.
  // 이 규약을 각 핸들러가 따로 구현하면 한 곳만 빠뜨려도 망이 화면에서 사라진다.
  function setEdit(state) {
    const prev = S.edit;
    S.edit = state;
    // state.keep 에 실린 이름은 «안 바뀜» 이다 — 들고 있던 사본을 그대로 쓴다.
    for (const name of state.keep || []) {
      if (prev && prev[name]) S.edit[name] = prev[name];
    }
    return S.edit;
  }

  async function loadEdit() {
    const d = await api(`/api/module-f/edit/state?sid=${S.sid}`);
    setEdit(d.state); S.key = d.key;
    // [F-10b] 손질에 들어오면 기본은 «알람밸브 원클릭» 이다 — 상무 시연이
    //   28분 내내 요구한 그 한 번이 첫 동작이 되게 한다. 이미 다른 모드를
    //   고른 뒤라면(모드 전환은 서버에 남는다) 그것을 지킨다.
    if (!S.emode) setUiMode(ONECLICK);
    loadSlots();
    setStage("edit");
    fit(S.edit.bounds);
    renderEdit();
    say(`${S.key} · 노드 ${S.edit.counts.pts.toLocaleString()}`
      + ` · 간선 ${S.edit.counts.edges.toLocaleString()}`
      + ` · 헤드 ${S.edit.counts.heads.toLocaleString()}`
      + ` · 덩이 ${S.edit.counts.bodies}`);
  }

  function renderEdit() {
    const e = S.edit;
    syncWorstSourceSelect(e);   // [F-1] 급수원 2곳 이상이면 기준 선택을 보인다
    const kinds = Object.entries(e.kinds)
      .map(([k, n]) => `${k} ${n}`).join(" · ") || "–";
    const undef = e.kinds["미지정"] || 0;
    // 물흐름을 돌리기 전에도 «급수원이 몇 개나 닿는지» 는 셀 수 있다.
    // 헤드 3,163개 중 264개만 닿는다는 사실을 나중에 아는 것은 너무 늦다.
    const bs = e.body_stat || {};
    const reachRow = bs.has_source
      ? kv("급수원이 닿는 헤드",
           `<span class="${bs.source_heads * 2 < bs.total_heads ? "warn" : "ok"}">`
           + `${bs.source_heads}</span> / ${bs.total_heads}`)
      : kv("가장 큰 덩이 헤드", `${bs.biggest_heads || 0} / ${bs.total_heads || 0}`);
    $("ed-info").innerHTML =
      kv("모드", e.mode) +
      kv("노드 / 간선", `${e.counts.pts} / ${e.counts.edges}`) +
      kv("덩이", `${e.counts.bodies}개`) +
      reachRow +
      kv("헤드 종류", kinds) +
      kv("급수 / 밸브", `${e.sources.length} / ${e.valves.length}`) +
      kv("이음 / 삭제", `${e.counts.joins} / ${e.counts.deletes}`) +
      (e.flowed ? kv("물 닿은 헤드",
        Object.entries(e.wet_counts).map(([k, n]) => `${k} ${n}`).join(" · ")) : "") +
      (e.worst ? kv("최불리망 <span class=\"tag\">설계면적</span>",
        `<span class="ok">${e.worst.k}개</span> · 앵커 ${e.worst.far_m} m`
        + ` · 폭 ${e.worst.span_m} m`
        + (e.worst.zones && e.worst.zones.length
           ? ` · 영역 ${e.worst.zones.length}곳` : "")
        + (e.worst.source ? ` · <b class="tag">${e.worst.source}</b> 기준` : ""))
      + kv('최원 유하거리 <span class="tag">경로</span>',
           `<span style="color:#ff3b3b">┈┈</span> ${e.worst.far_m} m`
           + ` · 절점 ${(e.worst.anchor_path || []).length}개`)
        + kv("배관 연장 / 주배관 부하",
          `${e.worst.total_m} m · <span class="ok">${e.worst.max_load}</span>개 담당`)
        : "") +
      (undef ? `<div class="kv"><b>변환 가능</b><span class="err">미지정 ${undef}개 — 막힘</span></div>`
             : `<div class="kv"><b>변환 가능</b><span class="ok">헤드 종류 확정</span></div>`);
    // [F-10d] 마지막 계산 뒤 고친 건수 — 0 이면 배지를 아예 감춘다. 늘 떠
    //   있으면 «지금 뭔가 밀려 있다» 는 신호가 아니라 장식이 된다.
    const nEdits = e.edits_since_worst || 0;
    $("ed-recalc-row").classList.toggle("hidden", !nEdits);
    $("ed-edits").textContent = `마지막 계산 후 수정 ${nEdits}건`;
    // [F-10e] 평면에서 보는 동안 고치면 그 배지도 같이 따라와야 한다 — 두
    //   화면이 같은 수를 보지 않으면 어느 쪽이 사실인지 알 수 없다.
    renderPlanUnderlay();

    // [F-10b] 화면 모드가 «원클릭» 이면 서버 모드로 덮지 않는다 — 원클릭은
    //   서버 모드가 아니라 둘을 한 번에 놓는 «행동» 이라 서버엔 이름이 없다.
    const uiMode = S.emode || e.mode;
    for (const b of document.querySelectorAll(".emode")) {
      b.classList.toggle("on", b.dataset.mode === uiMode);
    }
    $("ed-anchor-note").classList.toggle("hidden", uiMode !== ONECLICK);
    // 종류 단추의 점 색은 캔버스 헤드 색과 같은 표에서 온다 — 붙박이로 적으면
    // 색표를 고쳤을 때 그림과 도면이 어긋난다.
    const pal = (e.palette && e.palette.kinds) || {};
    for (const b of document.querySelectorAll(".ekind")) {
      b.disabled = !e.selected_head;
      const k = b.dataset.kind;
      b.querySelector(".dot").style.background = pal[k] || "";
      b.querySelector(".cnt").textContent = `${e.kinds[k] || 0}개`;
    }
    renderAutojoin(e);
    renderSheets(e);
    const none = $("ed-kind-none");
    none.querySelector(".dot").style.background = pal["미지정"] || "";
    none.querySelector(".cnt").textContent = `${undef}개`;
    none.classList.toggle("err", undef > 0);
    draw();
  }

  // ── 자동 이음 — A 의 실측 · E 의 판정 ──────────────────────────
  function renderAutojoin(e) {
    const aj = e.autojoin;
    $("ed-aj-apply").disabled = !(aj && aj.n);
    $("ed-aj-clear").disabled = !aj;
    const rep = e.autojoin_report;
    if (!aj) {
      $("ed-aj-eps-wrap").classList.add("hidden");
      $("ed-aj-info").innerHTML = rep
        ? kv("붙인 이음", `<span class="ok">${rep.made}</span>곳 · 막힘 `
             + `${rep.blocked} · 이미이어짐 ${rep.skipped}`)
          + kv("이음 모양", Object.entries(rep.kinds || {})
               .map(([k, n]) => `${k} ${n}`).join(" · ") || "–")
          + kv("덩이", `${rep.bodies_before} → ${rep.bodies_after}`)
        : "";
      return;
    }
    // 사다리 표를 그대로 보여준다 — 왜 이 여유를 골랐는지가 숫자에 남아야 한다.
    const here = aj.trials.find((t) => t.eps_mm === aj.eps_mm)
              || aj.trials[aj.trials.length - 1];
    const sel = $("ed-aj-eps");
    sel.innerHTML = aj.trials.map((t) =>
      `<option value="${t.eps_mm}"${t.eps_mm === aj.eps_mm ? " selected" : ""}>`
      + `${t.eps_mm} mm — 이을 곳 ${t.pairs} · 덩이 ${t.bodies}`
      + `${t.eps_mm === aj.auto_eps_mm ? " (도면 실측)" : ""}</option>`).join("");
    $("ed-aj-eps-wrap").classList.remove("hidden");
    $("ed-aj-info").innerHTML =
      kv("이음 여유", `${aj.eps_mm} mm`
        + (aj.eps_mm === aj.auto_eps_mm ? ' <span class="ok">실측</span>'
                                        : ' <span class="warn">직접</span>')) +
      kv("끊긴 관 끝", `${aj.ends.toLocaleString()}개`) +
      kv("방향이 맞는 짝", `${aj.kept} / ${aj.near}쌍`) +
      kv("이을 후보", `<span class="${aj.n ? "ok" : "warn"}">${aj.n}</span>군데 · `
        + (Object.entries(aj.by_kind || {}).map(([k, n]) => `${k} ${n}`)
           .join(" · ") || "–")
        + (aj.dropped ? ` <span class="err">(+${aj.dropped} 상한초과)</span>` : "")) +
      kv("덩이 예상", `${aj.bodies_before} → ${here.bodies}`);
  }

  async function autojoinScan(epsMm) {
    busy(true, "끊긴 곳을 재는 중…");
    try {
      const d = await post("/api/module-f/edit/autojoin/scan",
                           { sid: S.sid, eps_mm: epsMm || 0 });
      setEdit(d.state);
      renderEdit();
      say(d.message, S.edit.autojoin && S.edit.autojoin.n ? "ok" : "warn");
    } catch (err) { say(err.message, "err"); }
    finally { busy(false); }
  }

  $("ed-aj-scan").onclick = () => autojoinScan(0);
  $("ed-aj-eps").onchange = () => autojoinScan(Number($("ed-aj-eps").value));
  $("ed-aj-clear").onclick = async () => {
    try {
      const d = await post("/api/module-f/edit/autojoin/clear", { sid: S.sid });
      setEdit(d.state);
      renderEdit();
      say("자동 이음 후보를 지웠습니다.");
    } catch (err) { say(err.message, "err"); }
  };
  $("ed-aj-apply").onclick = async () => {
    busy(true, "자동 이음…");
    try {
      await post("/api/module-f/edit/autojoin/apply", { sid: S.sid });
      watch(async () => {
        await loadEdit();
        // 결과는 잡이 아니라 손질 상태에서 읽는다 — 잡 응답에는 진행 줄만 있다.
        const r = S.edit.autojoin_report || {};
        say(`자동 이음 — 붙임 ${r.made || 0} · 막힘 ${r.blocked || 0}`
          + ` · 덩이 ${r.bodies_before} → ${r.bodies_after}`
          + ` (여유 ${r.eps_mm} mm)`, r.made ? "ok" : "warn");
      });
    } catch (err) { busy(false); say(err.message, "err"); }
  };

  // ── 한 파일에 도면 여러 장 (모듈 A 규칙) ───────────────────────
  function renderSheets(e) {
    const sheets = e.sheets || [];
    const box = $("ed-sheets"), wrap = $("ed-sheet-wrap"), sel = $("ed-sheet");
    if (sheets.length < 2) {
      box.classList.add("hidden");
      wrap.classList.add("hidden");
      return;
    }
    box.classList.remove("hidden");
    box.innerHTML = `이 파일에는 도면이 <b>${sheets.length}장</b> 들어 있습니다.`
      + " 장을 고르지 않으면 최불리 30 이 <b>서로 다른 도면의 헤드를 섞어</b> 뽑습니다.";
    if (sel.options.length !== sheets.length + 1) {
      const keep = sel.value;
      sel.innerHTML = '<option value="0">전체 (섞임)</option>'
        + sheets.map((f) => `<option value="${f.index}">도면 ${f.index}`
          + ` — 헤드 ${f.head_count}개 · ${Math.round(f.size_mm[0] / 1000)}`
          + `×${Math.round(f.size_mm[1] / 1000)} m</option>`).join("");
      if (keep) sel.value = keep;
    }
    wrap.classList.remove("hidden");
  }

  // [F-10b · D-F10-4] 「알람밸브 원클릭」은 서버 모드가 아니라 **화면 모드** 다.
  //   서버의 손질 모드는 이음·삭제·급수시작위치·알람밸브위치 넷 그대로이고,
  //   원클릭은 그중 둘을 한 번에 놓는 «행동» 이다. 그래서 모드 전환을 서버에
  //   보내지 않고 여기서만 기억한다 — 엔진 계약을 늘리지 않는다.
  const ONECLICK = "원클릭";

  // 최불리 기준개수 — 원클릭과 「최불리 선정」이 같은 값을 써야 한다. 두 곳이
  // 갈리면 손질에서 본 개수와 표에 실린 K 가 달라진다.
  const edK = () => Math.max(1, Math.min(200, Number($("ed-k").value || 30)));

  function setUiMode(mode) {
    S.emode = mode;
    for (const b of document.querySelectorAll(".emode")) {
      b.classList.toggle("on", b.dataset.mode === mode);
    }
    $("ed-anchor-note").classList.toggle("hidden", mode !== ONECLICK);
  }

  for (const b of document.querySelectorAll(".emode")) {
    b.onclick = async () => {
      const mode = b.dataset.mode;
      if (mode === ONECLICK) { setUiMode(mode); say("알람밸브를 클릭하세요."); return; }
      try {
        const d = await post("/api/module-f/edit/mode",
                             { sid: S.sid, mode });
        setEdit(d.state);
        setUiMode(mode);
        renderEdit();
        say(`모드: ${mode}`);
      } catch (err) { say(err.message, "err"); }
    };
  }

  // 알람밸브 한 번 = 두 픽 + 최불리. 서버가 한 잡으로 한다(D-F10-4).
  async function anchorClick(x, y, maxD) {
    busy(true, "알람밸브 원클릭 — 두 자리를 놓고 최불리를 계산하는 중…");
    try {
      await post("/api/module-f/edit/anchor-click",
                 { sid: S.sid, x, y, max_d: maxD, k: edK() });
      watch(async () => {
        const j = await api(`/api/module-f/convert/result?sid=${S.sid}`);
        const r = j.result || {};
        if (r.state) { setEdit(r.state); }
        else { await loadEdit(); }
        renderEdit();
        draw();
        const s = r.summary;
        if (s) {
          startPulse();          // [F-10c] 방금 뜬 corridor 를 몇 번 도드라지게
          say(`최불리 ${s.k}개 · 최원 ${s.far_m} m · 담당 최대 ${s.max_load}개`
              + ` · 배관 ${s.path_edges}`, "ok");
        }
      });
    } catch (err) { busy(false); say(err.message, "err"); }
  }

  async function editClick(x, y, maxD) {
    if (S.emode === ONECLICK) { await anchorClick(x, y, maxD); return; }
    try {
      const d = await post("/api/module-f/edit/click",
                           { sid: S.sid, x, y, max_d: maxD });
      setEdit(d.state);
      renderEdit();
      if (!d.report) { say("아무것도 잡히지 않았습니다.", "warn"); return; }
      const r = d.report;
      let msg = r["동작"];
      if (r.kind) msg += ` · ${r.kind}`;
      if (r.made !== undefined) msg += ` · 만듦 ${r.made} 막힘 ${r.blocked}`;
      if (r.n !== undefined) msg += ` · ${r.n}개`;
      say(msg, "ok");
    } catch (err) { say(err.message, "err"); }
  }

  for (const b of document.querySelectorAll(".ekind")) {
    b.onclick = async () => {
      try {
        const d = await post("/api/module-f/edit/kind",
                             { sid: S.sid, kind: b.dataset.kind });
        setEdit(d.state);
        renderEdit();
        say(`헤드 종류를 ${b.dataset.kind} 으로 바꿨습니다.`, "ok");
      } catch (err) { say(err.message, "err"); }
    };
  }

  $("ed-undo").onclick = async () => {
    try {
      const d = await post("/api/module-f/edit/undo", { sid: S.sid });
      setEdit(d.state);
      renderEdit();
      say(d.undone ? "한 단계 되돌렸습니다." : "되돌릴 손질이 없습니다.",
          d.undone ? "" : "warn");
    } catch (err) { say(err.message, "err"); }
  };

  $("ed-flow").onclick = async () => {
    busy(true, "물흐름 계산 중…");
    try {
      const d = await post("/api/module-f/edit/flow", { sid: S.sid });
      setEdit(d.state);
      renderEdit();
      const w = d.water;
      say(`물 닿은 헤드 ${w.wet_heads}/${w.total_heads}`
        + ` · 젖은 간선 ${w.wet_edges} · 도달 노드 ${w.reach}`, "ok");
    } catch (err) { say(err.message, "err"); }
    finally { busy(false); }
  };

  // ── Remote 30 (모듈 A 개념 · E 그래프 위) ──────────────────────
  // [F-1] 어느 급수원 기준의 최불리인지 — kfp 변환과 같은 규약(Z1, Z2…).
  function syncWorstSourceSelect(e) {
    const sel = $("ed-src");
    if (!sel) return;
    const n = (e.sources || []).length;
    if (n < 2) { sel.classList.add("hidden"); sel.innerHTML = ""; return; }
    const keep = sel.value;
    sel.innerHTML = '<option value="">어느 급수원 기준?</option>'
      + e.sources.map((p, i) =>
          `<option value="Z${i + 1}">Z${i + 1} (${Math.round(p[0])}, ${Math.round(p[1])})</option>`
        ).join("");
    if (keep) sel.value = keep;
    sel.classList.remove("hidden");
  }

  async function runWorst(label) {
    busy(true, label);
    try {
      const sheet = Number(($("ed-sheet") || {}).value || 0);
      const k = edK();
      const body = { sid: S.sid, k, sheet };
      const src = ($("ed-src") || {}).value;
      if (src) body.source = src;
      if (S.zones.length) body.zones = S.zones;
      const d = await post("/api/module-f/edit/worst", body);
      setEdit(d.state);
      renderEdit();
      startPulse();              // [F-10c] 원클릭과 같은 연출 — 길만 다르다
      // 수리계산 단계도 같은 K 로 돈다 — 두 곳이 갈리면 손질에서 본 30개와
      // 표에 실린 K 가 달라져 「어느 쪽이 설계면적인가」 가 사라진다.
      $("dg-k").value = k;
      const s = d.summary;
      say(`최불리 ${s.k} 헤드 — 후보 ${s.candidates}개 중 · `
        + `최원 유하거리 ${s.far_m} m (경로 ${s.anchor_path_m} m)`
        + ` · ${s.k}번째 ${s.near_m} m`
        + (s.source ? ` · 급수원 ${s.source} 기준` : "")
        + (s.zones ? ` · 영역 ${s.zones}곳 안` : "")
        + (s.sheet ? ` · 도면 ${s.sheet}장 안` : ""),
          "ok");
      $("cv-worst-kfp").checked = true;
    } catch (err) { say(err.message, "err"); }
    finally { busy(false); }
  }

  $("ed-worst").onclick = () => runWorst("최불리 헤드 선정 중…");
  // [F-10d] 결과 위에서 고친 뒤 — 픽은 그대로 두고 최불리만 다시 돌린다.
  //   같은 몸통(`_compute_worst`)을 타므로 「최불리 선정」과 답이 같다.
  $("ed-recalc").onclick = () => runWorst("고친 망으로 최불리를 다시 계산 중…");

  $("ed-worst-clear").onclick = async () => {
    try {
      const d = await post("/api/module-f/edit/worst-clear", { sid: S.sid });
      setEdit(d.state);
      renderEdit();
      $("cv-worst-kfp").checked = false;
      say("최불리 선정을 해제했습니다.");
    } catch (err) { say(err.message, "err"); }
  };

  $("ed-save").onclick = async () => {
    try {
      const d = await post("/api/module-f/edit/save", { sid: S.sid });
      say(d.message, "ok");
    } catch (err) { say(err.message, "err"); }
  };

  $("ed-next").onclick = async () => {
    setStage("conv");
    await loadFields();
    const srcs = S.edit.sources;
    const wrap = $("src-wrap");
    const sel = $("conv-src");
    sel.innerHTML = "";
    if (srcs.length > 1) {
      wrap.classList.remove("hidden");
      srcs.forEach((p, i) => {
        const o = document.createElement("option");
        o.value = `Z${i + 1}`;
        o.textContent = `Z${i + 1} (${p[0]}, ${p[1]})`;
        sel.appendChild(o);
      });
    } else {
      wrap.classList.add("hidden");
    }
    // 모듈 E 는 이 단계에서 대화상자를 띄운다 — 같은 자리에서 같은 것을 묻는다.
    renderConvSummary();
    openConvModal();
    say("변환 값을 확인하고 실행하세요. 빈 칸은 기본값으로 갑니다.");
  };

  $("btn-back-edit").onclick = () => { setStage("edit"); renderEdit(); };

  // ── 4. 변환 ────────────────────────────────────────────────────
  let FIELDS = null;
  async function loadFields() {
    if (FIELDS) return;
    const d = await api("/api/module-f/convert/fields");
    FIELDS = d;
    const box = $("conv-fields");
    box.innerHTML = "";
    for (const g of d.groups) {
      const h = document.createElement("div");
      h.className = "grp";
      h.textContent = g.title;
      box.appendChild(h);
      // 칸 이름이 「① (m)」 뿐이다 — 어느 토막인지는 그림에만 적혀 있다.
      if (g.diagram) {
        const fig = document.createElement("div");
        fig.className = "grpfig";
        const im = document.createElement("img");
        im.className = "diagram";
        im.src = `/api/module-f/diagram/${g.diagram}`;
        im.alt = `${g.title} 배관 전개`;
        fig.appendChild(im);
        box.appendChild(fig);
      }
      for (const f of g.fields) {
        const lb = document.createElement("label");
        lb.className = "f";
        const sp = document.createElement("span");
        sp.textContent = f.label;
        const inp = document.createElement("input");
        inp.type = "text";
        inp.dataset.key = f.key;
        inp.placeholder = f.placeholder || "(비우면 기본값)";
        const dv = d.defaults[f.key];
        if (dv !== null && dv !== undefined) inp.value = String(dv);
        lb.append(sp, inp);
        box.appendChild(lb);
      }
    }
  }

  function readDto() {
    const dto = {};
    for (const inp of $("conv-fields").querySelectorAll("input")) {
      const raw = inp.value.trim();
      if (raw === "") { dto[inp.dataset.key] = null; continue; }
      const n = Number(raw);
      if (!Number.isFinite(n)) throw new Error(`${inp.dataset.key}: 숫자가 아닙니다 — "${raw}"`);
      dto[inp.dataset.key] = n;
    }
    return dto;
  }

  // ── 수직 전개 값 창 — 모듈 E 의 대화상자와 같은 자리 ────────────
  // 옆판에 펼쳐 넣으면 그림 다섯 장이 세로로 쌓여 다른 단추가 화면 밖으로
  // 밀린다. 창으로 띄우고, 옆판에는 채워진 값 요약만 남긴다.
  let convSnapshot = null;      // 취소하면 되돌릴 값

  function openConvModal() {
    convSnapshot = {};
    for (const inp of $("conv-fields").querySelectorAll("input")) {
      convSnapshot[inp.dataset.key] = inp.value;
    }
    $("conv-modal").classList.remove("hidden");
    const first = $("conv-fields").querySelector("input");
    if (first) first.focus();
  }

  function closeConvModal(revert) {
    if (revert && convSnapshot) {
      for (const inp of $("conv-fields").querySelectorAll("input")) {
        if (inp.dataset.key in convSnapshot) inp.value = convSnapshot[inp.dataset.key];
      }
    }
    convSnapshot = null;
    $("conv-modal").classList.add("hidden");
    renderConvSummary();
  }

  function renderConvSummary() {
    const box = $("conv-summary");
    const inputs = [...$("conv-fields").querySelectorAll("input")];
    if (!inputs.length) { box.textContent = "—"; return; }
    // 묶음별로 «몇 칸 중 몇 칸이 채워졌나» — 어느 그림의 값인지는 창에서 본다.
    const filled = inputs.filter((i) => i.value.trim() !== "").length;
    let html = kv("채운 칸", `${filled} / ${inputs.length}`);
    for (const g of (FIELDS ? FIELDS.groups : [])) {
      const vals = g.fields
        .map((f) => {
          const el = $("conv-fields").querySelector(`input[data-key="${f.key}"]`);
          return el && el.value.trim() !== "" ? el.value.trim() : null;
        })
        .filter((v) => v !== null);
      if (vals.length) html += kv(g.title, vals.join(" · ") + " m");
    }
    box.innerHTML = html;
  }

  $("btn-conv-fields").onclick = () => openConvModal();
  $("conv-ok").onclick = () => {
    try { readDto(); }        // 숫자가 아니면 닫지 않는다 — 창 안에서 고치게
    catch (err) { say(err.message, "err"); return; }
    closeConvModal(false);
  };
  $("conv-cancel").onclick = () => closeConvModal(true);
  $("conv-modal").addEventListener("click", (e) => {
    if (e.target === $("conv-modal")) closeConvModal(true);   // 바깥 클릭 = 취소
  });
  window.addEventListener("keydown", (e) => {
    if (e.key === "Escape" && !$("conv-modal").classList.contains("hidden")) {
      closeConvModal(true);
    }
  });

  $("btn-convert").onclick = async () => {
    let dto;
    try { dto = readDto(); }
    catch (err) { say(err.message, "err"); return; }
    const wrap = $("src-wrap");
    const selected = wrap.classList.contains("hidden") ? null : $("conv-src").value;
    for (const id of ["btn-download", "btn-download-worst",
                      "btn-download-design", "btn-download-set"]) {
      $(id).disabled = true;
    }
    $("conv-info").innerHTML = "";
    busy(true, "변환 중…");
    const outputs = {
      full_kfp: $("cv-full-kfp").checked,
      worst_kfp: $("cv-worst-kfp").checked,
      worst_sdf: $("cv-worst-sdf").checked,
    };
    if (!outputs.full_kfp && !outputs.worst_kfp && !outputs.worst_sdf) {
      busy(false);
      say("산출물을 하나도 고르지 않았습니다.", "err");
      return;
    }
    try {
      const d = await post("/api/module-f/convert/run", {
        sid: S.sid, dto, selected_source: selected, outputs,
      });
      if (d && d.ok === false && d.code === "worst_required") {
        // 막지 않는다 — 최불리 선정이 아직이면 수리계산 패널로 안내한다(D3).
        busy(false);
        say(d.message, "warn");
        setStage("design");
        return;
      }
      watch(showConvert);
    } catch (err) { busy(false); say(err.message, "err"); }
  };

  // 잡 상태(/job)에는 결과 본문이 없다 — 결과는 따로 받아 온다.
  async function showConvert() {
    let r;
    try {
      r = (await api(`/api/module-f/convert/result?sid=${S.sid}`)).result || {};
    } catch (err) { say(err.message, "err"); return; }
    // [정리 2026-08-31] `S.convResult = r` 을 지웠다 — 담아 두기만 하고 아무도
    //   안 읽었다. 상태에 «읽히지 않는 값» 이 있으면 다음 사람이 그것을 신뢰할
    //   수 있는 최신값으로 오해한다. 아래 코드는 지역 `r` 만 쓴다.
    if (!r.ok) {
      const rows = (r.blockers || [])
        .map((b) => `<tr><td>${b.code || ""}</td><td>${b.message || ""}</td></tr>`)
        .join("");
      $("conv-info").innerHTML =
        '<div class="kv"><b>결과</b><span class="err">변환 막힘</span></div>'
        + `<table class="rs">${rows}</table>`;
      say("변환이 막혔습니다. 사유를 확인하세요.", "err");
      return;
    }
    const s = r.summary || {}, st = r.stats || {};
    const pick = ["n_heads", "n_vert_head", "n_vert_branch", "n_tees",
                  "n_combo", "n_valve", "n_raised"];
    const rows = pick.filter((k) => k in st)
      .map((k) => `<tr><td>${k}</td><td>${st[k]}</td></tr>`).join("");
    let html = "";
    if (s.full) {
      html += kv("전체망 .kfp",
        `노드 ${s.full.nodes.toLocaleString()} · 배관 ${s.full.pipes.toLocaleString()}`
        + ` · ${(s.full.bytes / 1024).toFixed(0)} KB`);
    }
    if (s.worst) {
      html += kv(`최불리 .kfp <b class="tag">K${s.worst.k}</b>`,
        `노드 ${s.worst.nodes} · 배관 ${s.worst.pipes}`
        + ` · ${(s.worst.bytes / 1024).toFixed(0)} KB`);
    }
    if (s.design) {
      html += kv("최불리 .sdf",
        `<span class="ok">${s.design.sdf}</span>`
        + ` · ${(s.design.bytes / 1024).toFixed(0)} KB (+.slf)`);
    }
    $("conv-info").innerHTML = html + `<table class="rs">${rows}</table>`;
    $("btn-download").disabled = !s.full;
    $("btn-download-worst").disabled = !s.worst;
    $("btn-download-design").disabled = !s.design;
    $("btn-download-set").disabled = !(s.full || s.worst || s.design);
    const made = [s.full && "전체망 .kfp", s.worst && "최불리 .kfp",
                  s.design && "최불리 .sdf"].filter(Boolean).join(" · ");
    say(`변환 완료 — ${made}`, "ok");
  }

  const dl = (what) => {
    window.location.href = `/api/module-f/download?sid=${S.sid}&what=${what}`;
  };
  $("btn-download").onclick = () => dl("kfp");
  $("btn-download-worst").onclick = () => dl("worst-kfp");
  $("btn-download-design").onclick = () => dl("design");
  $("btn-download-set").onclick = () => dl("set");

  // ── 시작 ───────────────────────────────────────────────────────
  //
  // [검증 내보내기] 아래 블록의 `S.*` 는 **화면이 안 읽는다** — 브라우저 검증
  //   스크립트만 쓴다. 그래서 「쓰기만 하고 아무도 안 읽는 필드」 감사에서
  //   걸리는데, 그것이 정상이다. 감사기(`scripts/_probe_f_js_audit.py`)가
  //   이 표시를 보고 건너뛴다 — 표시를 지우면 감사가 다시 시끄러워진다.
  //
  // 상태와 좌표 환산을 밖으로 낸다 — 브라우저 검증 스크립트가 실제 화면 좌표로
  // 클릭해 보려면 세계↔화면 환산이 밖에서도 보여야 한다. 읽기 전용으로만 쓴다.
  window.__mf = S;
  S.toScreenX = sx;
  S.toScreenY = sy;
  S.toWorldX = wx;
  S.toWorldY = wy;
  // [F-11b] 직접 입력 화면은 «표가 확정된 뒤» 에만 서는 자리다 — 거기까지
  //   브라우저로 완주하려면 도면 한 장을 통째로 돌려야 한다(대명동 실측 수 분).
  //   그래서 그리는 함수도 같이 낸다. 검증은 진짜 표를 넣고 진짜 함수를 부른다.
  //   ★구문 검사만으로는 안 잡히는 회귀가 여기 산다: 함수-지역 헬퍼를 다른
  //     스코프에서 부르면 ReferenceError 로 화면만 조용히 빈다(저장소 규약).
  S.renderIssues = renderIssues;
  S.renderDesignTable = renderDesignTable;
  S.countFilled = countFilled;

  resize();

  // ── 5. 수리계산 입력 (설계) — G16 의 웹판 ──────────────────────────
  // 캔버스는 위의 기존 인프라(S.view·fit·sx/sy)를 그대로 쓴다. 좌표는
  // /design/preview 가 주는 «저장에 쓰는 그 값» 이다 — 여기서 다시 계산하는
  // 순간 미리보기가 거짓말이 된다.
  function designSettings() {
    return {
      k: Number($("dg-k").value || 30),
      schedule: $("dg-sched").value,
      iso: $("dg-iso").checked,
      iso_z_scale: Number($("dg-zscale").value || 1),
      canvas_units: Number($("dg-canvas").value || 3000),
      lift_ref: $("dg-ref").value,
      head_stub_pct: Number($("dg-stub").value || 2.5),
    };
  }

  // 수리계산 패널의 «표를 만드는» 입력은 수동 경로 것이다. 자동은 표가 이미
  // 나와 있어 그 단추가 「손질 세션이 없습니다」로 막히기만 한다.
  function syncDesignForMethod() {
    const auto = S.method === "auto";
    $("dg-build-inputs").classList.toggle("hidden", auto);
    $("dg-build-row").classList.toggle("hidden", auto);
    $("dg-back-auto-row").classList.toggle("hidden", !auto);
    // 자동은 「변환」 단계를 거치지 않으므로 그 산출물 단추도 뜻이 없다.
    $("dg-to-merge").classList.toggle("hidden", false);
  }

  $("dg-back-auto").onclick = () => loadAuto();

  // ── [F-8d] 탈출로 — 자동 결과를 손질로 이어받는다 ────────────────
  // 자동이 마음에 안 든다고 처음부터 다시 시작하게 두지 않는다. 같은 세션의
  // 찍기판은 살아 있다 — 채택 → 스펙 저장 → 손질 진입까지 서버 잡 하나다.
  $("au-handoff").onclick = async () => {
    busy(true, "인식 결과를 찍어 손질로 넘기는 중…");
    try {
      await post("/api/module-f/auto/handoff", { sid: S.sid });
      watch(async () => {
        const j = await api(`/api/module-f/convert/result?sid=${S.sid}`);
        const r = j.result || {};
        S.method = "manual";          // 자동 흐름을 떠난다 — 단계바가 갈린다
        S.handoff = r.alarm || r.source ? r : null;
        renderSteps();
        await loadEdit();             // 손질 화면 진입 (기존 경로)
        renderHandoff();
        draw();
        const g = Number(r.head_skipped) || 0;
        say(`손질로 이어받았습니다 — 재료 ${(r.mat_applied || []).length}묶음 · `
          + `헤드 ${(Number(r.head_applied) || 0)
                    + (Number(r.head_already) || 0)}개.`
          + (g ? ` 점선 ${g}개는 찍지 못했습니다.` : "")
          + " 알람밸브·급수 시작은 제안으로 표시했습니다 — 단추로 반영하세요.",
            g ? "warn" : "ok");
      });
    } catch (err) { busy(false); say(err.message, "err"); }
  };

  function renderHandoff() {
    const box = $("ed-handoff-box"), h = S.handoff;
    box.classList.toggle("hidden", !h);
    if (!h) return;
    $("ed-handoff-info").innerHTML =
      kv("찍은 것", `재료 ${(h.mat_applied || []).length}묶음 · `
         + `헤드 ${(Number(h.head_applied) || 0)
                   + (Number(h.head_already) || 0)}개`
         + (Number(h.head_skipped)
            ? ` · <span class="warn">유령 ${h.head_skipped}</span>` : ""))
      + kv("제안", h.alarm
           ? `알람밸브 · 급수 시작 (${h.alarm[0].toFixed(0)}, `
             + `${h.alarm[1].toFixed(0)})`
           : "자동이 알람밸브를 안 찍어 제안이 없습니다");
    $("ed-hint-alarm").disabled = !h.alarm;
    $("ed-hint-source").disabled = !h.source;
  }

  // 반영은 «기존 손질 클릭 경로» 로만 — 여기서도 주입은 없다(D-F8-3).
  async function applyHint(kind) {
    const h = S.handoff;
    const xy = kind === "알람밸브위치" ? (h && h.alarm) : (h && h.source);
    if (!xy) return;
    busy(true, "제안을 반영하는 중…");
    try {
      await post("/api/module-f/edit/mode", { sid: S.sid, mode: kind });
      const d = await post("/api/module-f/edit/click",
                           { sid: S.sid, x: xy[0], y: xy[1], max_d: 3000 });
      if (d.state) S.edit = d.state;
      renderEdit();
      draw();
      say(`${kind} 를 제안 자리에 반영했습니다.`, "ok");
    } catch (err) { say(err.message, "err"); }
    finally { busy(false); }
  }

  $("ed-hint-alarm").onclick = () => applyHint("알람밸브위치");
  $("ed-hint-source").onclick = () => applyHint("급수시작위치");

  function drawHandoffHints() {
    const h = S.handoff;
    if (!h) return;
    // 제안은 점선 고리다 — 확정된 것(실선)과 한눈에 갈려야 한다.
    ctx.setLineDash([4, 4]);
    ctx.lineWidth = 1.8;
    for (const [xy, color] of [[h.alarm, "#f97316"], [h.source, "#38bdf8"]]) {
      if (!xy) continue;
      ctx.beginPath();
      ctx.arc(sx(xy[0]), sy(xy[1]), 11, 0, Math.PI * 2);
      ctx.strokeStyle = color;
      ctx.stroke();
    }
    ctx.setLineDash([]);
    ctx.lineWidth = 1;
  }


  async function designPreview() {
    const cfg = designSettings();
    const q = new URLSearchParams({
      sid: S.sid, iso: cfg.iso ? "1" : "0",
      iso_z_scale: cfg.iso_z_scale, canvas_units: cfg.canvas_units,
      lift_ref: cfg.lift_ref, head_stub_pct: cfg.head_stub_pct,
    });
    const d = await api(`/api/module-f/design/preview?${q}`);
    S.design = { view: d.view, tables: d.tables, settings: d.settings,
                 marks: d.marks || {},
                 // [F-11d-2] 이번 계산에 «못 들어간» 직접 입력. 조용한 소실
                 //   금지 — 목록으로 올라가 사유까지 보인다.
                 ovMissed: d.ov_missed || [],
                 // [F-10f] 이상 목록이 부속·등가길이 수치를 여기서 읽는다.
                 summary: (S.design && S.design.summary) || null,
                 hilite: (S.design && S.design.hilite) || new Set() };
    // [§18] 고를 수 있는 부속 종류를 서버에서 받아 둔다 — 목록을 그리기 전에
    //   있어야 «고르세요» 칸이 빈 채로 뜨지 않는다.
    if (!S.fitKinds) await loadFitKinds();
    // [F-11c] 쓸 수 있는 호칭경도 서버에서 받아 둔다 — 화면이 따로 목록을 들면
    //   규격표가 바뀔 때 둘이 갈린다.
    if (!S.boreAllowed) await loadBoreOv();
    renderIssues();
    // ★«아직 확정 안 함» 은 오류가 아니라 상태다(서버가 200 · view:null 로
    //   답한다). 그릴 것이 없으면 여기서 조용히 멈춘다 — 화면은 「표 확정」
    //   단추가 선 채로 남는다.
    if (!d.view) { if (d.message) say(d.message); return; }
    const xs = d.view.nodes.map(n => n.x), ys = d.view.nodes.map(n => n.y);
    fit({ minx: Math.min(...xs), maxx: Math.max(...xs),
          miny: Math.min(...ys), maxy: Math.max(...ys) });
    syncDesignForMethod();
    renderDesignTable();
    renderBoreLegend();
    if (S.method === "auto") renderAutoDesignSummary();
    // ★미리보기가 떴다는 것은 표가 있다는 뜻이다 — 저장할 수 있다.
    //   예전에는 수동 「표 확정」 안에서만 풀어, 자동 경로에서는 산출 단추가
    //   영영 잠겨 있었다(계통도 없이 자동으로 뽑으면 저장할 길이 없었다).
    $("dg-emit").disabled = false;
    draw();
  }

  // 자동 경로는 「표 확정」을 거치지 않아 요약이 빈 채로 남는다 — 자동 추출이
  // 낸 수치를 그대로 옮긴다(두 경로가 같은 것을 말하게).
  function renderAutoDesignSummary() {
    const s = S.autoSummary;
    if (!s) { $("dg-summary").textContent = "—"; return; }
    let html = kv("설계면적", `<span class="ok">${s.k}개</span>`
                  + ` · 최원 ${s.far_m} m`)
      + kv("표", `절점 ${s.nodes} · 배관 ${s.pipes} · 노즐 ${s.nozzles}`
           + ` · 부속 ${s.fittings}`)
      + kv("경로", '자동 <span class="tag">MODULE A</span>');
    if (s.source_fallback) {
      html += kv('<span class="warn">급수원 대체</span>',
                 `${s.source_bridge_mm} mm 떨어져 최근접 절점으로`);
    }
    $("dg-summary").innerHTML = html;
  }

  function drawDesign() {
    // ★표 요약만 받고 미리보기는 아직인 상태가 있다(renderDesignSummary 가
    //   먼저 돈다). 그때 그리려 들면 「Cannot read properties of undefined」로
    //   화면이 멈춘다 — 그릴 것이 없으면 조용히 돌아간다.
    const v = S.design && S.design.view;
    if (!v || !v.nodes) return;
    const at = {};
    for (const n of v.nodes) at[n.label] = n;
    const maxLoad = Math.max(1, ...v.pipes.map(p => p.load || 0));
    for (const p of v.pipes) {
      const a = at[p.a], b = at[p.b];
      if (!a || !b) continue;
      const hot = S.design.hilite.has(p.label);
      // 관경을 무엇이 정했는지로 가른다. 규약으로만 정한 구간은 점선이다 —
      // 도면에서 읽은 실측과 규약 추정을 한 모양으로 그리면 구분이 안 된다.
      const st = S.boreColor ? BORE_STYLE[p.src] : null;
      ctx.strokeStyle = hot ? "#f97316" : (st ? st.color : "#94a3b8");
      ctx.setLineDash(hot || !st ? [] : st.dash);
      ctx.lineWidth = (1 + 4 * (p.load || 0) / maxLoad) + (hot ? 2 : 0);
      ctx.beginPath();
      ctx.moveTo(sx(a.x), sy(a.y));
      ctx.lineTo(sx(b.x), sy(b.y));
      ctx.stroke();
    }
    ctx.setLineDash([]);      // ★되돌린다 — 안 하면 아래 노드 기호까지 점선이 된다
    // 최원 유하거리 경로 — 손질 단계와 같은 빨간 점선. 두 단계가 같은 줄을
    // 가리켜야 «이 관을 키우면 그 압이 오른다» 가 이어진다.
    const dap = v.anchor_path || [];
    if (dap.length > 1) {
      ctx.strokeStyle = "#ff3b3b";
      ctx.lineWidth = 2.4;
      ctx.setLineDash([9, 5]);
      ctx.beginPath();
      let started = false;
      for (const lab of dap) {
        const n = at[lab];
        if (!n) continue;
        if (started) ctx.lineTo(sx(n.x), sy(n.y));
        else { ctx.moveTo(sx(n.x), sy(n.y)); started = true; }
      }
      if (started) ctx.stroke();
      ctx.setLineDash([]);
    }
    ctx.lineWidth = 1.2;
    for (const n of v.nodes) {
      const px = sx(n.x), py = sy(n.y);
      if (n.head) {
        // 상향 △ / 하향 ▽ — 방향 규칙은 베이크와 같다(표고 차 부호).
        const s = 7;
        ctx.beginPath();
        if (n.up) {
          ctx.moveTo(px, py - s);
          ctx.lineTo(px - s, py + s * 0.6);
          ctx.lineTo(px + s, py + s * 0.6);
        } else {
          ctx.moveTo(px, py + s);
          ctx.lineTo(px - s, py - s * 0.6);
          ctx.lineTo(px + s, py - s * 0.6);
        }
        ctx.closePath();
        ctx.fillStyle = "rgba(248,113,113,.45)";
        ctx.strokeStyle = "#ef4444";
        ctx.fill();
        ctx.stroke();
      }
      if (n.input) {
        ctx.beginPath();
        ctx.arc(px, py, 9, 0, Math.PI * 2);
        ctx.strokeStyle = "#3b82f6";
        ctx.lineWidth = 2;
        ctx.stroke();
        ctx.lineWidth = 1.2;
      }
      if (n.valve) {
        ctx.strokeStyle = "#22c55e";
        ctx.lineWidth = 2;
        ctx.strokeRect(px - 7, py - 7, 14, 14);
        ctx.lineWidth = 1.2;
      }
      // 앵커 = 기준압을 잡는 지점. 손질 단계와 같은 빨간 겹원.
      if (n.anchor) {
        ctx.strokeStyle = "#ff3b3b";
        ctx.lineWidth = 2.6;
        ctx.beginPath(); ctx.arc(px, py, 10, 0, Math.PI * 2); ctx.stroke();
        ctx.beginPath(); ctx.arc(px, py, 14, 0, Math.PI * 2); ctx.stroke();
        ctx.lineWidth = 1.2;
      }
    }
  }

  // 표 4종 — 저장될 값 그대로. 배관 표에는 관경 근거를 사람 말로 잇는다.
  const DG_COLS = { label: "이름", in: "시작", out: "끝", type: "관종",
    dia: "호칭경(mm)", length: "길이(m)", elev: "표고차(m)", c: "C",
    status: "상태", group: "그룹", dia_src: "관경 근거", elevation: "표고(m)",
    io_node: "입출력", x: "x", y: "y", flow_lmin: "유량(L/min)",
    count: "개수", pipe: "배관", pressure_pa: "압력(Pa)",
    flow_m3s: "유량(m³/s)", lib: "라이브러리", eq_len: "등가길이(m)",
    rel_pos: "위치", desc: "설명", off_tree: "루프 잔여" };
  const DG_SRC = { text: "도면 텍스트", nfpc_min: "별표1 보강",
                   nfpc_fallback: "별표1 폴백" };
  // 관경 근거별 캔버스 표시. 규약으로만 정한 것(별표1 폴백)은 점선 — 도면에서
  // 읽은 실측과 한 모양으로 그리지 않는다는 이 저장소의 규약을 따른다.
  const BORE_STYLE = {
    text:          { color: "#38bdf8", dash: [],     tip: "도면 치수 텍스트에서 읽음" },
    nfpc_min:      { color: "#facc15", dash: [],     tip: "도면 값이 별표1 최소보다 작아 안전측으로 올림" },
    nfpc_fallback: { color: "#64748b", dash: [6, 4], tip: "가까운 치수 텍스트가 없어 담당 헤드 수로 정함" },
  };

  function renderBoreLegend() {
    const box = $("dg-bore-legend");
    if (!S.design || !S.design.view) {
      box.textContent = "—";
      return;
    }
    // ★자동(A) 경로의 배관 행에는 관경 «근거» 칸이 없다(A 의 build_input_tables
    //   는 dia_src 를 남기지 않는다). 규칙은 수동과 같지만 배관별 근거를 기록
    //   하지 않을 뿐이다. 없는 것을 「전부 근거 없음」으로 그리면 도면에 치수가
    //   없다는 뜻으로 읽힌다 — 사실이 아니다. 그래서 비운다고 말한다.
    if (S.method === "auto") {
      box.innerHTML = kv("관경 근거",
        "자동 경로는 배관별 근거를 남기지 않습니다");
      $("dg-bore-color").disabled = true;
      $("dg-bore-color").checked = false;
      S.boreColor = false;
      return;
    }
    $("dg-bore-color").disabled = false;
    const pipes = S.design.view.pipes || [];
    const n = {};
    for (const p of pipes) n[p.src] = (n[p.src] || 0) + 1;
    const total = pipes.length || 1;
    let html = "";
    for (const k of ["text", "nfpc_min", "nfpc_fallback"]) {
      const st = BORE_STYLE[k];
      const c = n[k] || 0;
      const mark = `<span style="color:${st.color}" title="${st.tip}">`
        + `${st.dash.length ? "┈┈" : "━━"}</span>`;
      html += kv(`${mark} ${DG_SRC[k]}`,
                 `${c}개 · ${(c / total * 100).toFixed(1)}%`);
    }
    const unknown = pipes.filter(p => !BORE_STYLE[p.src]).length;
    if (unknown) html += kv("근거 없음", `${unknown}개`);
    box.innerHTML = html;
  }

  // [F-11b-3] 사람이 채운 자리는 표에서도 «다른 얼굴» 이어야 한다 — 자동이 낸
  //   값과 같은 얼굴로 두면 나중에 그 수치를 누가 정했는지 알 길이 없다.
  //
  //   ★엔진의 부속표에는 그런 칸이 없고, 이 항목에서 서버는 불변이다(지시서
  //     F-11b 수용기준). 그래서 화면이 «이미 받아 둔» `unresolved.applied` 를
  //     표에 겹쳐 놓는다 — 새 판정이 아니라 표시다. 개수는 여전히 엔진 한
  //     곳에서만 나오므로 둘이 어긋날 수 없다.
  //   ★맞춤은 (배관, 종류) 로 한다. 부속표 행에는 «어느 노드인지» 가 없어 더
  //     좁힐 수 없다 — 한 배관에 같은 종류가 둘이면 둘 다 표시된다. 넓게
  //     보이는 쪽이 «채운 걸 안 보여 주는» 쪽보다 정직하다.
  function overrideNoteOf(row, which) {
    const t = S.design.tables || {};
    // [F-11c] 관경은 «규칙 값도» 덮으므로 전·후가 함께 보여야 한다(D-F11-3).
    //   「직접 입력 80A — 사유 (원래 별표1 폴백 65A)」.
    if (which === "pipes") {
      const b = (t.bore_overrides || {})[String(row.label)];
      if (!b) return null;
      return `직접 입력 ${b.dia}A` + (b.note ? ` — ${b.note}` : "")
        + ` (원래 ${DG_SRC[b.orig_src] || b.orig_src} ${b.orig_dia}A)`;
    }
    if (which !== "fittings") return null;
    const app = (t.unresolved || {}).applied || [];
    if (!app.length) return null;
    const dia = ((t.pipes || []).find(
      (p) => String(p.label) === String(row.pipe)) || {}).dia;
    for (const a of app) {
      // 등가길이는 «(종류, 호칭경) 쌍» 이 단위라 배관이 아니라 그 쌍으로 맞춘다.
      const hit = a.what === "kind"
        ? (String(a.pipe) === String(row.pipe)
           && String(a.kind) === String(row.type))
        : (String(a.kind) === String(row.type)
           && Number(a.dia) === Number(dia));
      if (hit) {
        return (a.what === "kind" ? "직접 입력 — 부속" : "직접 입력 — 등가길이")
          + (a.note ? ` · ${a.note}` : "");
      }
    }
    return null;
  }

  function renderDesignTable() {
    if (!S.design) return;
    const which = $("dg-table").value;
    const rows = (S.design.tables[which] || []);
    const cols = [];
    for (const r of rows) {
      for (const k in r) if (!cols.includes(k)) cols.push(k);
    }
    // 채운 자리가 있을 때만 «근거» 칸을 덧붙인다 — 없으면 표를 안 건드린다.
    const notes = rows.map((r) => overrideNoteOf(r, which));
    const hasOv = notes.some(Boolean);
    let html = "<table><thead><tr>"
      + cols.map(c => `<th>${DG_COLS[c] || c}</th>`).join("")
      + (hasOv ? "<th>근거</th>" : "")
      + "</tr></thead><tbody>";
    rows.forEach((r, i) => {
      html += `<tr data-label="${r.label != null ? r.label : ""}">`
        + cols.map(c => {
            let v = r[c];
            if (c === "dia_src") v = DG_SRC[v] || v;
            return `<td>${v != null ? v : ""}</td>`;
          }).join("")
        + (hasOv ? `<td class="ovcell">${notes[i] || ""}</td>` : "")
        + "</tr>";
    });
    $("dg-grid").innerHTML = html + "</tbody></table>";
    // 행 → 캔버스 강조 (배관 표에서만 뜻이 있다)
    for (const tr of $("dg-grid").querySelectorAll("tr[data-label]")) {
      tr.onclick = () => {
        const lab = tr.dataset.label;
        if (which !== "pipes" || !lab) return;
        if (S.design.hilite.has(lab)) S.design.hilite.delete(lab);
        else S.design.hilite.add(lab);
        tr.classList.toggle("hl");
        draw();
        renderBoreOv();     // 고른 배관 수가 관경 덮기 자리에 바로 뜬다
      };
    }
    renderBoreOv();
  }

  // ── [F-11c] 관경 «직접 입력» ────────────────────────────────────
  //
  // 부속·등가길이(§18)와 문법은 같고 **범위만 다르다**(D-F11-3): 저 둘은 규칙이
  // 못 가린 자리에만 쓰지만 관경은 규칙이 낸 값도 덮는다 — 도면 치수가 틀렸거나
  // 설계 협의로 바뀌는 일이 실제로 있다. 그래서 «원래 얼마였나» 를 표에 항상
  // 같이 남긴다. 덮었다는 사실이 안 남으면 나중에 그 수치를 누가 정했는지
  // 알 길이 없다.
  //
  // ★자리를 가리키는 키는 **board 노드쌍** 이다(D-F11-4). 배관 라벨(P12)은 BFS
  //   순서로 매겨지므로 corridor 가 바뀌면 같은 이름이 다른 배관을 가리킨다 —
  //   사람이 80A 라고 적어 둔 자리가 조용히 옆 배관으로 옮겨간다.
  async function loadBoreOv() {
    try {
      const d = await api(`/api/module-f/design/bore-override?sid=${S.sid}`);
      S.boreAllowed = d.allowed || [];
      S.boreRows = d.rows || [];
      S.boreSchedule = d.schedule || "";
    } catch (err) { S.boreAllowed = S.boreAllowed || []; }
  }

  /** 배관 라벨 → board 노드쌍. 역참조가 없으면 null — 그 배관은 못 덮는다. */
  function boreRefOf(label) {
    const v = (S.design && S.design.view) || null;
    const p = v && (v.pipes || []).find(
      (x) => String(x.label) === String(label));
    return (p && p.ref) || null;
  }

  function renderBoreOv() {
    const row = $("dg-bore-row"), why = $("dg-bore-why");
    if (!row) return;
    const on = !!(S.design && S.design.view)
      && $("dg-table").value === "pipes";
    row.classList.toggle("hidden", !on);
    why.classList.toggle("hidden", !on);
    if (!on) { $("dg-bore-list").innerHTML = ""; return; }
    // 호칭경은 «서버가» 준 규격표 값만. 자유 숫자를 두면 SLF 에 없는 값이 들어가
    // PIPENET 이 그 배관을 못 푼다 — 「엘베」 교훈의 관경판이다.
    const sel = $("dg-bore-dia"), allow = S.boreAllowed || [];
    if (sel.options.length !== allow.length + 1) {
      sel.innerHTML = '<option value="">— 호칭경 —</option>'
        + allow.map((d) => `<option value="${d}">${d}A</option>`).join("");
    }
    // 대상은 표에서 누른 줄 그대로다 — 캔버스 강조와 같은 집합을 쓴다.
    const all = [...((S.design && S.design.hilite) || [])];
    const picked = all.filter((lab) => boreRefOf(lab));
    const noRef = all.length - picked.length;
    $("dg-bore-n").textContent = `고른 배관 ${picked.length}개`
      + (noRef ? ` · 못 덮는 것 ${noRef}개` : "");
    $("dg-bore-save").disabled = !picked.length || !sel.value;
    why.innerHTML = "표에서 배관 줄을 눌러 고른 뒤 호칭경을 정합니다 — "
      + `<b>${S.boreSchedule || "규격표"}</b> 에 있는 값만 쓸 수 있습니다. `
      + "부속과 달리 <b>규칙이 낸 값도 덮습니다</b> — 원래 값은 표에 남습니다. "
      + "덮은 뒤 <b>「표 확정」을 다시</b> 눌러야 산출에 들어갑니다."
      + (noRef ? " (도면에 그려진 선이 아닌 배관 — 헤드 접속관·가지 상승 — 은"
                 + " 가리킬 자리가 없어 못 덮습니다.)" : "");
    // 지금 덮어 둔 것 + 지우는 길. 갇히지 않게 하는 것이 «완결성» 이다.
    const rows = S.boreRows || [];
    $("dg-bore-list").innerHTML = rows.map((r, i) =>
      `<div class="ovrow"><span class="dim">노드 ${r.a}–${r.b} · `
      + `<b>${r.dia}A</b>${r.note ? ` — ${r.note}` : ""}</span>`
      + `<button class="ovdel" data-b="${i}">관경 덮기 지우기</button></div>`)
      .join("");
    for (const el of $("dg-bore-list").querySelectorAll(".ovdel")) {
      el.onclick = () => dropBoreOv(Number(el.dataset.b));
    }
  }

  /** 덮기 목록 «전체» 를 보낸다 — 서버가 그 목록을 그대로 세션에 둔다. */
  async function postBoreOv(rows, msg) {
    busy(true, "관경 직접 입력을 저장하는 중…");
    try {
      const d = await post("/api/module-f/design/bore-override",
                           { sid: S.sid, rows });
      S.boreRows = d.rows || [];
      // 배지 규약은 부속과 같다 — 재확정 전까지 「아직 안 들어갔다」를 말한다.
      S.ovDirty = !!d.needs_rebuild;
      busy(false);
      say(msg, "ok");
      $("dg-build").click();     // 값이 바뀌는 일이라 재확정까지 이어 준다
    } catch (err) { busy(false); say(err.message, "err"); }
  }

  $("dg-bore-dia").onchange = renderBoreOv;

  $("dg-bore-save").onclick = () => {
    const dia = Number($("dg-bore-dia").value || 0);
    if (!dia) { say("호칭경을 고르세요.", "warn"); return; }
    const note = String($("dg-bore-note").value || "").trim();
    // 이미 덮어 둔 것과 «합친다» — 한 번에 다 덮지 않아도 되게(§18 저장과 같다).
    const keep = new Map((S.boreRows || []).map((r) => [`${r.a}|${r.b}`, r]));
    let n = 0;
    for (const lab of ((S.design && S.design.hilite) || [])) {
      const ref = boreRefOf(lab);
      if (!ref) continue;
      keep.set(`${ref[0]}|${ref[1]}`, { a: ref[0], b: ref[1], dia, note });
      n += 1;
    }
    if (!n) { say("덮을 배관을 표에서 고르세요.", "warn"); return; }
    postBoreOv([...keep.values()],
               `관경 ${n}개를 ${dia}A 로 덮었습니다 — 표를 다시 확정합니다.`);
  };

  function dropBoreOv(i) {
    postBoreOv((S.boreRows || []).filter((_r, k) => k !== i),
               "관경 직접 입력을 지웠습니다 — 표를 다시 확정합니다.");
  }

  // ── [F-10f] 이상 표시 — 전수 검수 대신 ──────────────────────────
  //
  // 전사 27:36 「클릭을 다 클릭을 하는 것도 불편할 수 있거든」 · 27:41 「뭔가
  // 좀 이상하면 표시를 해서 확인을 해서 수정을 하고」. 집계 숫자로만 있던 것을
  // **항목** 으로 내린다. 목록이 0 이면 그것이 사람 검수의 완료 신호다.
  //
  // ★새 계산을 만들지 않는다 — 전부 이미 화면에 와 있는 자료다. 관경 근거는
  //   배관 행의 `src`, 제외 사유는 `marks`, 유령은 채택 결과. 그래서 서버를
  //   한 줄도 안 바꿨고, 산출물이 안 변한다는 것이 자명하다(D-F10-7).
  const ISSUE_CAP = 40;              // 조용히 자르지 않는다 — 남은 수를 적는다

  // ── [§18] 직접 입력 — 규칙이 못 가린 자리를 사람이 채운다 ──────
  //
  // ★고를 수 있는 종류는 «서버가» 준다. 자유 입력으로 두면 라이브러리에 없는
  //   이름이 들어와 부속 판정은 풀리지만 등가길이가 다시 미해결이 된다
  //   (실측: 「엘베」로 적었더니 판정 불가 3→2, 등가길이 0→1). 문제를 옮길 뿐이다.
  const kindLabel = (v) => {
    const hit = (S.fitKinds || []).find((k) => k.value === String(v));
    return hit ? hit.label : String(v);
  };

  async function loadFitKinds() {
    try {
      const d = await api(`/api/module-f/design/fitting-override?sid=${S.sid}`);
      S.fitKinds = d.kinds || [];
      S.fitOverrides = d.overrides || {};
    } catch (err) { S.fitKinds = S.fitKinds || []; }
  }

  function collectIssues() {
    const out = [];
    const s = (S.design && S.design.summary) || null;
    const v = (S.design && S.design.view) || null;

    // ① 관경 별표1 폴백 — 도면에 치수 텍스트가 없어 담당 헤드 수로 정한 배관.
    if (v && v.pipes) {
      const at = {};
      for (const n of v.nodes) at[String(n.label)] = n;
      const fb = v.pipes.filter((p) => p.src === "nfpc_fallback");
      if (fb.length) {
        out.push({
          key: "bore", color: "#64748b", n: fb.length,
          label: "관경 — 별표1 폴백 (도면 치수 없음)",
          items: fb.slice(0, ISSUE_CAP).map((p) => {
            const a = at[String(p.a)], b2 = at[String(p.b)];
            return {
              text: `${p.label} · ${p.dia}A · 담당 ${p.load}`,
              x: (a && b2) ? (a.x + b2.x) / 2 : null,
              y: (a && b2) ? (a.y + b2.y) / 2 : null,
              frame: "iso",
            };
          }),
        });
      }
    }

    // ② 유령 — 채택이 못 찍은 후보. 좌표는 mm(평면)다.
    if (S.ghosts && S.ghosts.size && S.suggest) {
      const gs = [...S.ghosts].filter((i) => S.suggest[i]);
      if (gs.length) {
        out.push({
          key: "ghost", color: "#f472b6", n: gs.length,
          label: "유령 — 채택이 못 찍은 헤드 후보",
          items: gs.slice(0, ISSUE_CAP).map((i) => ({
            text: `후보 #${i} · 신뢰도 ${S.suggest[i].conf}`,
            x: S.suggest[i].x, y: S.suggest[i].y, frame: "plan",
          })),
        });
      }
    }

    // ③ 제외 사유 — F-5 가 이미 갈라 둔 세 갈래. 좌표는 mm(평면)다.
    const m = (S.design && S.design.marks) || {};
    for (const [key, label, color] of [
      ["dry", "물길 미도달 헤드", "#64748b"],
      ["unattached", "이음 끊김 (부착 실패)", "#eab308"],
      ["unpicked", "찍히지 않음 (후보 제안 대비)", "#a855f7"],
    ]) {
      const g = m[key];
      if (!g || !g.n) continue;
      out.push({
        key, color, n: g.n, label,
        items: (g.xy || []).slice(0, ISSUE_CAP).map((p, i) => ({
          text: `${label} #${i + 1}`, x: p[0], y: p[1], frame: "plan",
        })),
      });
    }

    // ④ 부속 판정 불가 · 등가길이 미해결 — 이제 «어느 배관인지» 까지 온다.
    //    엔진이 세는 그 자리에서 목록도 함께 남기므로(§18) 개수와 어긋날 수
    //    없다. 자리를 아는 항목은 눌러서 그 배관으로 갈 수 있다.
    const un = (S.design && S.design.tables && S.design.tables.unresolved) || {};
    const nodeAt = {};
    if (v && v.nodes) for (const n of v.nodes) nodeAt[String(n.label)] = n;
    const pipeAt = {};
    if (v && v.pipes) for (const p of v.pipes) pipeAt[String(p.label)] = p;
    const mid = (pid) => {
      const p = pipeAt[String(pid)];
      const a = p && nodeAt[String(p.a)], b2 = p && nodeAt[String(p.b)];
      return (a && b2) ? [(a.x + b2.x) / 2, (a.y + b2.y) / 2] : [null, null];
    };

    const ki = un.kind_items || [];
    if (ki.length) {
      out.push({
        key: "fitting", color: "#f59e0b",
        n: ki.reduce((a, x) => a + (Number(x.n) || 0), 0),
        label: "부속 판정 불가 — 지어내지 않고 비워 둔 자리",
        items: ki.slice(0, ISSUE_CAP).map((x) => {
          const [mx, my] = mid(x.pipe);
          return {
            text: `${x.pipe} · ${x.where}`
              + (x.angle_deg !== undefined && x.angle_deg !== null
                 ? ` · 편향 ${x.angle_deg}°` : "")
              + ` (노드 ${x.node})`,
            x: mx, y: my, frame: "iso",
            // [§18] 이 자리를 채울 재료 — 자리(노드·배관)가 단위다.
            ov: { type: "kind", node: String(x.node), pipe: String(x.pipe) },
          };
        }),
      });
    }
    // 등가길이는 «(종류, 호칭경) 쌍» 이 채우기 단위다 — 한 번 채우면 그 쌍을
    // 쓰는 배관이 한꺼번에 풀린다. 그래서 배관이 아니라 쌍을 항목으로 세운다.
    const pairs = un.pairs || [];
    if (pairs.length) {
      const li = un.length_items || [];
      const firstPipe = (kind, dia) => {
        const hit = li.find((x) => String(x.kind) === String(kind)
                                && String(x.dia) === String(dia));
        return hit ? hit.pipe : null;
      };
      out.push({
        key: "eqlen", color: "#f59e0b",
        n: li.length,
        label: "등가길이 미해결 — 라이브러리에 그 호칭경 값이 없음",
        note: "한 쌍을 채우면 그 쌍을 쓰는 배관이 한꺼번에 풀립니다.",
        items: pairs.slice(0, ISSUE_CAP).map((p) => {
          const [mx, my] = mid(firstPipe(p.kind, p.dia));
          return {
            text: `${p.kind} · ${p.dia}A — ${p.n}건`,
            x: mx, y: my, frame: "iso",
            ov: { type: "eq_len", kind: String(p.kind), dia: Number(p.dia) },
          };
        }),
      });
    }
    // ★[F-11d-2] 적용 «못 한» 수정 — 조용한 소실 금지.
    //
    // 사람이 채운 값이 다음 계산에서 안 들어가는 일이 있다: 그 자리가 corridor
    // 에서 빠졌거나, 자동이 답을 내서 «판정 불가» 가 아니게 됐거나. 개수만 세면
    // 사람은 들어간 줄 안다. 사유와 함께 목록으로 올린다.
    const miss = (S.design && S.design.ovMissed) || [];
    if (miss.length) {
      out.push({
        key: "ovmiss", color: "#ef4444", n: miss.length,
        label: "적용 못 한 수정 — 직접 입력이 이번 산출에 안 들어갔습니다",
        note: "값은 지워지지 않았습니다. 자리가 돌아오면 다시 적용됩니다.",
        items: miss.slice(0, ISSUE_CAP).map((m) => {
          const [mx, my] = mid(m.pipe);
          const what = m.what === "eq_len"
            ? `${m.kind} ${m.dia}A · ${m.m} m`
            : `${m.pipe || "?"} · ${kindLabel(m.kind)}`;
          return {
            text: `${what} — ${m.why || "사유 없음"}`
              + (m.note ? ` (사유 「${m.note}」)` : ""),
            x: mx, y: my, frame: "iso",
          };
        }),
      });
    }
    // 채운 자리를 목록에 남긴다 — 값이 어디서 왔는지 나중에도 알 수 있어야 한다.
    const app = un.applied || [];
    if (app.length) {
      out.push({
        key: "applied", color: "#22c55e", n: app.length,
        label: "직접 입력 — 사람이 채운 자리",
        note: "표 확정에 이미 반영된 값입니다.",
        items: app.slice(0, ISSUE_CAP).map((a) => {
          const [mx, my] = mid(a.pipe);
          return {
            text: (a.what === "kind"
                   ? `${a.pipe} · ${kindLabel(a.kind)}`
                   : `${a.kind} ${a.dia}A · ${a.m} m`)
              + (a.note ? ` — ${a.note}` : ""),
            x: mx, y: my, frame: "iso",
            // [F-11b-4] 지우는 길 — 지우면 그 자리는 다시 미해결로 돌아간다.
            //   막다른 길을 만들지 않는 것이 «완결성» 이다(지시서 §0.1).
            del: (a.what === "kind"
                  ? { type: "kind", node: String(a.node), pipe: String(a.pipe) }
                  : { type: "eq_len", kind: String(a.kind), dia: Number(a.dia) }),
          };
        }),
      });
    }
    return out;
  }

  function renderIssues() {
    const box = $("dg-issues"), chip = $("dg-issues-n");
    if (!box) return;
    // ★일람은 **여기 맨 앞에서** 그린다. 아래에 조기 반환이 둘 있어(표 없음 ·
    //   이상 0건) 끝에 두면 그 두 경우에 일람이 안 그려진다 — 「이상 없음」인
    //   도면일수록 감사 화면이 필요한데 하필 그때 비는 셈이다.
    renderAudit();
    // ★표가 없으면 «없다» 가 아니라 «아직 모른다» 다. 안 재고 「이상 없음」이라
    //   적으면 그것은 완료 신호를 위조하는 것이다(저장소 규약: 정직한 진행 표시).
    if (!S.design || !S.design.view) {
      chip.textContent = "—";
      chip.classList.remove("ok");
      box.innerHTML = '<div class="hint">표를 확정하면 확인할 것이 '
        + "여기 모입니다.</div>";
      return;
    }
    const groups = collectIssues();
    const total = groups.reduce((a, g) => a + g.n, 0);
    chip.textContent = total ? `${total.toLocaleString()}건` : "없음";
    chip.classList.toggle("ok", !total);
    if (!total) {
      // 완료 신호 — 이것이 「다 봤다」의 뜻이다.
      box.innerHTML = '<div class="ok">확인할 이상 없음</div>';
      return;
    }
    box.innerHTML = groups.map((g, gi) => {
      const head = `<div class="kv"><b><span style="color:${g.color}">●</span> `
        + `${g.label}</b><span>${g.n.toLocaleString()}건</span></div>`
        // 덧말은 항목을 «가리지» 않는다 — 둘 다 필요하다(예: 채울 값 목록).
        + (g.note ? `<div class="hint">${g.note}</div>` : "");
      if (!g.items.length) return head;
      const rows = g.items.map((it, ii) => {
        const line = `<div class="issue" data-g="${gi}" data-i="${ii}">`
          + `${it.text}</div>`;
        // [F-11b-4] 채운 자리에는 지우는 단추를 단다.
        if (it.del) {
          return line + `<div class="ovrow"><button class="ovdel"`
            + ` data-id="${gi}-${ii}">직접 입력 지우기</button></div>`;
        }
        if (!it.ov) return line;
        // 채울 수 있는 자리에는 그 자리에서 바로 넣는 칸을 붙인다.
        const id = `${gi}-${ii}`;
        // 이름을 `box` 로 두면 바깥의 목록 상자를 가린다 — 지금은 안 쓰지만
        // 가려진 이름은 나중에 조용히 틀린다.
        const field = it.ov.type === "kind"
          ? `<select class="ovk" data-id="${id}">`
            + `<option value="">— 고르세요 —</option>`
            + (S.fitKinds || []).map((k) =>
                `<option value="${k.value}">${k.label}</option>`).join("")
            + `</select>`
          : `<input class="ovm" data-id="${id}" type="number" step="0.01"`
            + ` min="0" placeholder="등가길이 m">`;
        return line + `<div class="ovrow" data-id="${id}">${field}`
          + `<input class="ovn" data-id="${id}" maxlength="200"`
          + ` placeholder="사유 (어디서 확인했는지)"></div>`;
      }).join("");
      const rest = g.n - g.items.length;
      return head + rows
        + (rest > 0 ? `<div class="hint">… 그 외 ${rest.toLocaleString()}건`
                      + " (목록은 40건까지 보입니다)</div>" : "");
    }).join("");
    S.issues = groups;
    for (const el of box.querySelectorAll(".issue")) {
      el.onclick = () => {
        const g = S.issues[Number(el.dataset.g)];
        const it = g && g.items[Number(el.dataset.i)];
        if (it && it.x !== null && it.y !== null) focusIssue(it, g.color);
      };
    }
    // ★확인할 것이 «처음 생겼을 때» 한 번만 펴 준다. 매번 펴면 사람이 접어
    //   둔 것을 계속 되돌리게 된다(진행 표시가 쓰는 그 규약과 같다).
    if (total && !S.issuesOpened) {
      S.issuesOpened = true;
      const h2 = document.querySelector('h2.fold[data-fold="dg-issues-body"]');
      if (h2) toggleFold(h2, true);
    }
    // 채울 칸이 하나라도 있으면 저장 단추를 연다.
    const fillable = groups.some((g) => g.items.some((it) => it.ov));
    // [F-11b-2] 「표를 다시 확정해야 값이 들어간다」를 배지로 — F-10d 의
    //   «다시 계산» 배지와 같은 자리·같은 문법이다. 저장 뒤 재확정까지 자동으로
    //   이어 주지만, ★그 재확정이 실패하면 이 배지가 남아 사실을 말한다.
    //   그래서 채울 칸이 없어도(다 채웠어도) 배지 줄은 서 있어야 한다.
    $("dg-ov-row").classList.toggle("hidden", !fillable && !S.ovDirty);
    $("dg-ov-why").classList.toggle("hidden", !fillable);
    if (fillable) {
      $("dg-ov-why").innerHTML =
        "규칙이 <b>못 가린 자리에만</b> 쓰입니다 — 자동이 옳게 판정한 값은 "
        + "바뀌지 않습니다. 저장한 뒤 <b>「표 확정」을 다시</b> 눌러야 "
        + "산출에 들어갑니다.";
      for (const el of box.querySelectorAll(".ovk, .ovm, .ovn")) {
        el.onchange = countFilled;
        el.oninput = countFilled;
      }
    }
    countFilled();     // 배지·저장 단추는 «항상» 지금 사실에 맞춘다
    for (const el of box.querySelectorAll(".ovdel")) {
      el.onclick = () => {
        const [gi, ii] = el.dataset.id.split("-").map(Number);
        const d = S.issues[gi].items[ii].del;
        if (d) dropOverride(d);
      };
    }
  }

  // [F-11b-4] 직접 입력 하나를 지운다 — 지운 자리는 다시 미해결로 돌아가
  // 목록에 재등장한다. 잘못 채운 값에 갇히지 않는 것이 «완결성» 이다.
  async function dropOverride(d) {
    const prev = S.fitOverrides || {};
    const keep = (rows, key, gone) =>
      (rows || []).filter((r) => key(r) !== gone);
    const body = { sid: S.sid };
    if (d.type === "kind") {
      body.kind = keep(prev.kind, (r) => `${r.node}|${r.pipe}`,
                       `${d.node}|${d.pipe}`);
    } else {
      body.eq_len = keep(prev.eq_len, (r) => `${r.kind}|${r.dia}`,
                         `${d.kind}|${d.dia}`);
    }
    busy(true, "직접 입력을 지우는 중…");
    try {
      const r = await post("/api/module-f/design/fitting-override", body);
      S.fitOverrides = r.overrides || {};
      // ★저장은 «값이 바뀌는 일» 이지 표시가 아니다. 재확정 전까지는 산출이
      //   아직 옛 값이므로 그 사실을 배지로 든다(아래 저장 경로와 같은 규약).
      S.ovDirty = !!r.needs_rebuild;
      busy(false);
      say("직접 입력을 지웠습니다 — 표를 다시 확정합니다.", "ok");
      $("dg-build").click();     // 값이 바뀌는 일이라 재확정까지 이어 준다
    } catch (err) { busy(false); say(err.message, "err"); }
  }

  // ── [F-11d-3] 직접 입력 일람 — 감사 화면 ────────────────────────
  //
  // 「확인할 것」은 «지금 고칠 것» 을 보여 주는 자리다. 이쪽은 성격이 다르다 —
  // 나중에 「이 수치를 누가·왜 정했나」를 되짚는 자리라, 산출에 들어간 것과
  // 못 들어간 것을 **한자리에** 모아 둔다. 새로 계산하지 않는다: 전부 이미
  // 화면에 와 있는 자료다(applied · ovMissed · boreRows · fitOverrides).
  function renderAudit() {
    const box = $("dg-audit"), chip = $("dg-audit-n");
    if (!box) return;
    const t = (S.design && S.design.tables) || null;
    const un = (t && t.unresolved) || {};
    const rows = [];
    for (const a of (un.applied || [])) {
      rows.push({
        ok: true,
        what: a.what === "kind" ? "부속" : "등가길이",
        where: a.what === "kind"
          ? `${a.pipe} · 노드 ${a.node}`
          : `${a.kind} ${a.dia}A (그 쌍을 쓰는 배관 전부)`,
        val: a.what === "kind" ? kindLabel(a.kind) : `${a.m} m`,
        note: a.note || "",
      });
    }
    for (const [lab, b] of Object.entries((t && t.bore_overrides) || {})) {
      rows.push({
        ok: true, what: "관경", where: `${lab} · 노드 ${b.a}–${b.b}`,
        val: `${b.dia}A (원래 ${DG_SRC[b.orig_src] || b.orig_src} `
          + `${b.orig_dia}A)`,
        note: b.note || "",
      });
    }
    for (const m of ((S.design && S.design.ovMissed) || [])) {
      rows.push({
        ok: false,
        what: m.what === "eq_len" ? "등가길이" : "부속",
        where: m.what === "eq_len"
          ? `${m.kind} ${m.dia}A` : `${m.pipe || "?"} · 노드 ${m.node || "?"}`,
        val: m.what === "eq_len" ? `${m.m} m` : kindLabel(m.kind),
        note: m.note || "", why: m.why || "",
      });
    }
    chip.textContent = rows.length ? `${rows.length}건` : "없음";
    chip.classList.toggle("ok", !rows.length);
    if (!rows.length) {
      // ★«없다» 와 «아직 모른다» 를 가른다 — 표가 없으면 잰 적이 없는 것이다.
      box.innerHTML = t
        ? '<div class="hint">직접 입력한 값이 없습니다 — 전부 자동이 낸 값입니다.</div>'
        : '<div class="hint">표를 확정하면 여기 모입니다.</div>';
      return;
    }
    const n_ok = rows.filter((r) => r.ok).length;
    box.innerHTML =
      `<div class="kv"><b>산출에 들어감 ${n_ok}건</b>`
      + `<span>${rows.length - n_ok ? `못 들어감 ${rows.length - n_ok}건` : ""}`
      + `</span></div>`
      + rows.map((r) =>
          `<div class="issue"><span style="color:${r.ok ? "#22c55e" : "#ef4444"}">`
          + `●</span> <b>${r.what}</b> — ${r.where} → <b>${r.val}</b>`
          + (r.note ? ` · 사유 「${r.note}」` : " · <i>사유 없음</i>")
          + (r.why ? `<br><span class="dim">${r.why}</span>` : "")
          + `</div>`).join("");
  }

  /** 지금 몇 칸이 채워졌나 — 저장 전에 사람이 보고 안다.
   *
   * ★배지 자리는 하나다. 「표 확정 필요」와 「채운 칸 n」이 같은 칸을 쓰므로
   *   **여기 한 곳에서만** 쓴다. 두 곳에서 쓰면 나중 것이 앞 것을 덮는다 —
   *   실제로 그랬다: 배지를 세워 놓고 곧바로 「채운 칸 0」이 지워 버렸다.
   */
  function countFilled() {
    const box = $("dg-issues");
    let n = 0;
    for (const el of box.querySelectorAll(".ovk, .ovm")) {
      if (String(el.value || "").trim() !== "") n += 1;
    }
    const el = $("dg-ov-n");
    el.textContent = S.ovDirty
      ? "표 확정 필요 — 저장한 직접 입력이 아직 산출에 안 들어갔습니다"
      : `채운 칸 ${n}`;
    el.classList.toggle("warn", !!S.ovDirty);
    $("dg-ov-save").disabled = n === 0;
  }

  // 채운 것만 모아 보낸다. 빈 칸은 «안 정했다» 이지 «지운다» 가 아니다.
  $("dg-ov-save").onclick = async () => {
    const box = $("dg-issues");
    const note = (id) => {
      const el = box.querySelector(`.ovn[data-id="${id}"]`);
      return el ? String(el.value || "").trim() : "";
    };
    const kind = [], eq_len = [];
    for (const el of box.querySelectorAll(".ovk")) {
      const v = String(el.value || "").trim();
      if (!v) continue;
      const [gi, ii] = el.dataset.id.split("-").map(Number);
      const ov = S.issues[gi].items[ii].ov;
      kind.push({ node: ov.node, pipe: ov.pipe, kind: v, note: note(el.dataset.id) });
    }
    for (const el of box.querySelectorAll(".ovm")) {
      const v = String(el.value || "").trim();
      if (v === "") continue;
      const [gi, ii] = el.dataset.id.split("-").map(Number);
      const ov = S.issues[gi].items[ii].ov;
      eq_len.push({ kind: ov.kind, dia: ov.dia, m: Number(v),
                    note: note(el.dataset.id) });
    }
    if (!kind.length && !eq_len.length) { say("채운 칸이 없습니다.", "warn"); return; }
    try {
      // 이전에 저장한 것과 «합친다» — 한 번에 다 채우지 않아도 되게.
      const prev = S.fitOverrides || {};
      const merge = (old, add, key) => {
        const m = new Map((old || []).map((r) => [key(r), r]));
        for (const r of add) m.set(key(r), r);
        return [...m.values()];
      };
      const d = await post("/api/module-f/design/fitting-override", {
        sid: S.sid,
        kind: merge(prev.kind, kind, (r) => `${r.node}|${r.pipe}`),
        eq_len: merge(prev.eq_len, eq_len, (r) => `${r.kind}|${r.dia}`),
      });
      S.fitOverrides = d.overrides || {};
      // [F-11b-2] 아직 «저장만» 된 상태다 — 재확정이 끝나야 산출이 바뀐다.
      //   배지를 여기서 세우고, 재확정이 성공한 자리에서만 내린다. 재확정이
      //   실패하면 배지가 남아 「안 들어갔다」를 화면이 말한다.
      S.ovDirty = !!d.needs_rebuild;
      say(d.message || "직접 입력을 저장했습니다.", "ok");
      // 값이 바뀌는 일이라 표시만 고치고 끝내지 않는다 — 다시 확정한다.
      $("dg-build").click();
    } catch (err) { say(err.message, "err"); }
  };

  // 항목을 누르면 그 자리로 옮겨 가 강조한다. 좌표계가 둘이라(설계 vs 평면)
  // 먼저 «맞는 화면» 으로 돌린 뒤 옮긴다 — 안 그러면 엉뚱한 자리를 비춘다.
  function focusIssue(it, color) {
    const wantPlan = it.frame === "plan";
    const el = $("dg-plan");
    if (el && el.checked !== wantPlan) {
      el.checked = wantPlan;
      renderPlanUnderlay();
    }
    const bb = wantPlan
      ? (S.edit && S.edit.bounds)
      : null;
    let span;
    if (bb) span = Math.max(bb.maxx - bb.minx, bb.maxy - bb.miny);
    else if (S.design && S.design.view) {
      const xs = S.design.view.nodes.map((n) => n.x);
      const ys = S.design.view.nodes.map((n) => n.y);
      span = Math.max(Math.max(...xs) - Math.min(...xs),
                      Math.max(...ys) - Math.min(...ys));
    } else span = 1000;
    const pad = Math.max(span * 0.06, 1e-6);
    fit({ minx: it.x - pad, maxx: it.x + pad,
          miny: it.y - pad, maxy: it.y + pad });
    S.focus = { x: it.x, y: it.y, color: color || "#ff3b3b",
                frame: it.frame };
    draw();
    say(`${it.text} — 그 자리로 옮겼습니다.`);
  }

  // 강조 고리 — «어디를 보라» 는 표시다. 표시 전용이라 아무것도 안 바꾼다.
  function drawFocus() {
    const f = S.focus;
    if (!f) return;
    const onPlan = planUnderlayOn() || designMarksOn();
    if ((f.frame === "plan") !== onPlan) return;   // 지금 그 좌표계가 아니다
    ctx.save();
    ctx.strokeStyle = f.color;
    ctx.lineWidth = 2;
    for (const r of [10, 16]) {          // 겹고리 — 한 겹은 배경에 묻힌다
      ctx.beginPath();
      ctx.arc(sx(f.x), sy(f.y), r, 0, Math.PI * 2);
      ctx.stroke();
    }
    ctx.restore();
  }

  function renderDesignSummary(s) {
    S.design = S.design || {};
    S.design.summary = s;          // [F-10f] 이상 목록이 이 수치를 그대로 쓴다
    const b = s.bore_src || {};
    $("dg-summary").innerHTML =
      kv("설계면적", `<span class="ok">${s.k}개</span> · 앵커 ${s.far_m} m`
        + (s.source ? ` · <b class="tag">${s.source}</b> 기준` : ""))
      + kv("폭 / corridor", `${s.span_m} m / ${s.total_m} m`)
      + kv("주배관 담당", `${s.max_load}개`)
      + kv("표", `노드 ${s.counts.nodes} · 배관 ${s.counts.pipes}`
        + ` · 노즐 ${s.counts.nozzles} · 부속 ${s.counts.fittings}`)
      + kv("관경 근거", `텍스트 ${b.text} · 별표1 보강 ${b.nfpc_min}`
        + ` · 별표1 폴백 ${b.nfpc_fallback}`)
      + (S.design && S.design.view && S.design.view.anchor
         ? kv('최원 유하거리 <span class="tag">경로</span>',
              `<span style="color:#ff3b3b">┈┈</span> 앵커 절점`
              + ` ${S.design.view.anchor} · ${S.design.view.anchor_path_m} m`)
         : "")
      + kv("부속 판정 불가",
           `${s.fitting_unresolved} · 등가길이 미해결 ${s.eq_len_unresolved}`)
      + (s.excluded_heads
         ? kv('<span class="warn">제외 헤드</span>',
              `${s.excluded_heads.toLocaleString()}개 (후보 ${s.candidate_heads}`
              + ` / 도면 ${s.total_heads.toLocaleString()})`)
         : "")
      + (s.excluded_detail
         ? kv("제외 사유",
              [["dry", "물길 미도달"], ["unattached", "이음 끊김"],
               ["unpicked", "찍히지 않음"]]
                .filter(([k]) => k in s.excluded_detail)
                .map(([k, lab]) =>
                  `${lab} ${s.excluded_detail[k].toLocaleString()}`)
                .join(" · "))
         : "");
  }

  // ── [F-5] 찍기 후보 제안 — 표시는 여기, 반영은 기존 찍기 경로로만 ──
  function suggestColor(conf) {
    return conf >= 0.9 ? "#22c55e" : conf >= 0.75 ? "#eab308" : "#94a3b8";
  }

  function drawSuggest() {
    for (let i = 0; i < S.suggest.length; i++) {
      const c = S.suggest[i];
      // [F-8c] 낮은 띠는 접어 둔다 — 3천 점 위에 또 겹치면 아무것도 안 보인다.
      if (!S.showLow && Number(c.conf) < 0.75
          && !(S.ghosts && S.ghosts.has(i))) continue;
      const px = sx(c.x), py = sy(c.y);
      // 채택됨은 실선, 찍히지 못한 유령은 점선. 추정과 실측을 한 선으로 그리지
      // 않는다는 저장소 규약을 후보 표시에도 그대로 적용한다.
      const ghost = S.ghosts ? S.ghosts.has(i) : false;
      ctx.setLineDash(ghost ? [3, 3] : []);
      ctx.beginPath();
      ctx.arc(px, py, 6, 0, Math.PI * 2);
      ctx.strokeStyle = ghost ? "#ef4444" : suggestColor(c.conf);
      ctx.lineWidth = 1.6;
      ctx.stroke();
      ctx.setLineDash([]);
      if (S.adopted && S.adopted.has(i)) {
        ctx.beginPath();
        ctx.arc(px, py, 2, 0, Math.PI * 2);
        ctx.fillStyle = suggestColor(c.conf);
        ctx.fill();
      }
      if (S.suggestOff.has(i)) {
        ctx.beginPath();
        ctx.moveTo(px - 5, py - 5); ctx.lineTo(px + 5, py + 5);
        ctx.moveTo(px - 5, py + 5); ctx.lineTo(px + 5, py - 5);
        ctx.strokeStyle = "#ef4444";
        ctx.stroke();
      }
    }
    ctx.lineWidth = 1;
  }

  function suggestInfo() {
    const n = S.suggest ? S.suggest.length : 0;
    const off = S.suggestOff.size;
    $("pk-suggest-info").innerHTML = n
      ? kv("후보", `${n}개 · 제외 ${off}개 · 반영 예정 ${n - off}개`)
      : "";
  }

  $("pk-suggest").onclick = async () => {
    busy(true, "모듈 A 인식으로 후보를 찾는 중…");
    try {
      await post("/api/module-f/pick/suggest", { sid: S.sid });
      watch(async () => {
        try {
          const j = await api(`/api/module-f/convert/result?sid=${S.sid}`);
          const r = (j.result || {});
          if (!r.ok) throw new Error("후보 제안 실패 — 진행 로그를 확인하세요.");
          S.suggest = r.candidates || [];
          S.suggestOff = new Set();
          $("pk-suggest-apply").disabled = !S.suggest.length;
          $("pk-suggest-clear").disabled = !S.suggest.length;
          suggestInfo();
          draw();
          const b = r.bands || {};
          say(`후보 ${r.n}개 — ${Object.entries(b)
            .map(([k, v]) => `${k} ${v}`).join(" · ")}. 후보일 뿐, `
            + "확정은 사용자의 반영입니다.", "ok");
        } catch (err) { say(err.message, "err"); }
      });
    } catch (err) { busy(false); say(err.message, "err"); }
  };

  $("pk-suggest-clear").onclick = () => {
    S.suggest = null;
    S.suggestOff = new Set();
    // [F-8c] 채택 표시도 같이 걷는다 — 후보가 없는데 유령만 남으면 무엇을
    // 가리키는 점선인지 알 수 없다. 찍힌 것 자체는 board 에 그대로 있다.
    S.ghosts = null;
    S.adopted = null;
    $("pk-adopt-box").classList.add("hidden");
    $("pk-suggest-apply").disabled = true;
    $("pk-suggest-clear").disabled = true;
    suggestInfo();
    draw();
  };

  $("pk-suggest-apply").onclick = async () => {
    if (!S.suggest || !S.suggest.length) return;
    if (!S.pick || !S.pick.mat_done) {
      say("재료 선택을 먼저 완료해야 헤드를 반영할 수 있습니다.", "warn");
      return;
    }
    busy(true, "후보를 찍기 경로로 반영 중…");
    let okN = 0, dupN = 0, failN = 0, lastState = null;
    try {
      for (let i = 0; i < S.suggest.length; i++) {
        if (S.suggestOff.has(i)) continue;
        const c = S.suggest[i];
        // ★사람 클릭과 같은 API — E 의 확정 규칙이 그대로 심판한다.
        //   E 가 그 자리에서 헤드 표시를 못 찾으면 그 후보는 반영되지 않는다.
        //   찍기는 «문양 서명» 단위 토글이라, 이미 찍힌 서명 위의 후보는
        //   «취소» 로 응답한다 — 즉시 되클릭해 복원하고 «이미 반영» 으로 센다.
        const d = await post("/api/module-f/pick/click",
                             { sid: S.sid, x: c.x, y: c.y, max_d: 300 });
        const act = d.report && d.report["동작"];
        if (act === "추가") okN++;
        else if (act === "취소") {
          const d2 = await post("/api/module-f/pick/click",
                                { sid: S.sid, x: c.x, y: c.y, max_d: 300 });
          if (d2.state) lastState = d2.state;
          dupN++;
        }
        else failN++;
        if (d.state) lastState = d.state;
        if (i % 50 === 0) busy(true, `반영 중… ${i}/${S.suggest.length}`);
      }
      if (lastState) S.pick = lastState;
      say(`새 문양 ${okN}개 반영 · 이미 찍힌 문양 ${dupN}개 · E 가 거른 후보 `
        + `${failN}개 — 거른 것은 그 자리에 헤드 표시가 없다는 뜻입니다.`,
        failN ? "warn" : "ok");
      renderPick();
      draw();
    } catch (err) { say(err.message, "err"); }
    finally { busy(false); }
  };

  // 후보 클릭 → 반영 제외/복원 (찍기 단계, 후보 있을 때만)
  cv.addEventListener("mousedown", (e) => {
    if (S.stage !== "pick" || !S.suggest || e.button !== 0) return;
    for (let i = 0; i < S.suggest.length; i++) {
      const c = S.suggest[i];
      // ★[F-8c] 유령 위의 클릭은 가로채지 않는다. 유령은 «아직 안 찍힌 것» 이라
      //   사람이 직접 찍으려고 누르는 자리다 — 여기서 삼키면 그 길이 막힌다.
      //   화면에 안 그린 후보(낮은 띠 접힘)도 마찬가지다: 안 보이는 점이 클릭을
      //   먹으면 왜 안 찍히는지 알 도리가 없다.
      if (S.ghosts && S.ghosts.has(i)) continue;
      if (!S.showLow && Number(c.conf) < 0.75) continue;
      const dx = sx(c.x) - e.offsetX, dy = sy(c.y) - e.offsetY;
      if (dx * dx + dy * dy <= 64) {
        if (S.suggestOff.has(i)) S.suggestOff.delete(i);
        else S.suggestOff.add(i);
        suggestInfo();
        draw();
        e.stopImmediatePropagation();
        e.preventDefault();
        return;
      }
    }
  }, true);

  // ── [F-5] 설계 제외 사유 토글 ──
  function designMarksOn() {
    return $("dg-mk-dry").checked || $("dg-mk-unatt").checked
      || $("dg-mk-unpicked").checked;
  }

  function drawDesignMarks() {
    const m = (S.design && S.design.marks) || {};
    const draws = [
      ["dry", "dg-mk-dry", "#64748b"],
      ["unattached", "dg-mk-unatt", "#eab308"],
      ["unpicked", "dg-mk-unpicked", "#a855f7"],
    ];
    for (const [key, id, color] of draws) {
      if (!$(id).checked || !m[key]) continue;
      ctx.strokeStyle = color;
      ctx.lineWidth = 1.4;
      for (const [x, y] of m[key].xy) {
        const px = sx(x), py = sy(y);
        ctx.beginPath();
        ctx.arc(px, py, 5, 0, Math.PI * 2);
        ctx.stroke();
      }
    }
    ctx.lineWidth = 1;
  }

  // ── [F-10e] 평면에서 보기 — 밑그림 + 그 자리 수정 ────────────────
  //
  // 전사 24:47 의 요구는 「밑에 배관이 흐릿하게 보이고, 안 맞는 게 있으면 그
  // 자리에서 클릭해 고친다」이다. 그것을 **평면** 에서 만족시킨다.
  //
  // ★아이소 «아래» 에 깔지 않은 이유는 취향이 아니라 실측이다(BLOCKED §17):
  //   설계 좌표계는 board 의 변환이 아니라 빌드마다 새로 생성되는 스키매틱
  //   배치다. board→설계 전역 아핀이 없고(최대 잔차 도면의 9.3%), K 를 바꾸면
  //   같은 절점이 중앙값 11,281 만큼 옮겨진다. 겹쳐 그리면 어긋난 그림 위에서
  //   엉뚱한 배관을 고치게 된다 — 지시서 스스로 「의미가 없다」고 한 상태다.
  const planUnderlayOn = () => {
    const el = $("dg-plan");
    return !!(el && el.checked);
  };

  function fitDesignView() {
    if (planUnderlayOn() && S.edit && S.edit.bounds) { fit(S.edit.bounds); return; }
    if (designMarksOn() && S.edit && S.edit.bounds) { fit(S.edit.bounds); return; }
    // ★`S.design` 은 있는데 `view` 가 없을 수 있다 — 표 요약만 받고 미리보기는
    //   아직인 상태다(renderDesignSummary 가 먼저 돈다). 그때 `.view.nodes` 를
    //   읽으면 「Cannot read properties of undefined」로 화면이 멈춘다.
    if (!S.design || !S.design.view) return;
    const xs = S.design.view.nodes.map(n => n.x);
    const ys = S.design.view.nodes.map(n => n.y);
    fit({ minx: Math.min(...xs), maxx: Math.max(...xs),
          miny: Math.min(...ys), maxy: Math.max(...ys) });
  }

  function renderPlanUnderlay() {
    const on = planUnderlayOn();
    $("dg-plan-row").classList.toggle("hidden", !on);
    $("dg-plan-row2").classList.toggle("hidden", !on);
    const n = (S.edit && S.edit.edits_since_worst) || 0;
    $("dg-edits").textContent = `마지막 계산 후 수정 ${n}건`;
    const mode = (S.edit && S.edit.mode) || "";
    for (const b of document.querySelectorAll(".dgmode")) {
      b.classList.toggle("on", b.dataset.mode === mode);
    }
  }

  $("dg-plan").onchange = async () => {
    // 손질 상태를 안 들고 있으면 평면을 그릴 수 없다 — 한 번 받아 둔다.
    if (planUnderlayOn() && !S.edit) {
      try {
        const d = await api(`/api/module-f/edit/state?sid=${S.sid}`);
        setEdit(d.state);
      } catch (err) { say(err.message, "err"); }
    }
    renderPlanUnderlay();
    fitDesignView();
    draw();
  };

  for (const b of document.querySelectorAll(".dgmode")) {
    b.onclick = async () => {
      try {
        const d = await post("/api/module-f/edit/mode",
                             { sid: S.sid, mode: b.dataset.mode });
        setEdit(d.state);
        renderPlanUnderlay();
        draw();
        say(`모드: ${b.dataset.mode} — 흐린 배관을 클릭해 고치세요.`);
      } catch (err) { say(err.message, "err"); }
    };
  }

  // 다시 계산 → 표 확정 → 아이소 갱신. 셋이 한 단추다 — 「고쳤으니 다시」가
  // 사람 머릿속에서는 한 동작이기 때문이다. 자동 재실행은 여전히 없다(D-F10-5).
  $("dg-recalc").onclick = async () => {
    busy(true, "고친 망으로 최불리를 다시 계산 중…");
    try {
      // ★K 는 «이 화면의» 값(dg-k)을 쓴다. 손질의 ed-k 를 쓰면 표가 확정되는
      //   K 와 최불리 K 가 갈려 「어느 쪽이 설계면적인가」가 사라진다.
      const sheet = Number(($("ed-sheet") || {}).value || 0);
      const k = Math.max(1, Math.min(200, Number($("dg-k").value || 30)));
      const body = { sid: S.sid, k, sheet };
      const src = ($("ed-src") || {}).value;
      if (src) body.source = src;
      if (S.zones.length) body.zones = S.zones;
      const d = await post("/api/module-f/edit/worst", body);
      setEdit(d.state);
      renderPlanUnderlay();
      busy(false);
      $("dg-build").click();          // 표 확정 → designPreview → 아이소 갱신
    } catch (err) { busy(false); say(err.message, "err"); }
  };

  $("dg-bore-color").onchange = () => {
    S.boreColor = $("dg-bore-color").checked;
    draw();
  };

  for (const id of ["dg-mk-dry", "dg-mk-unatt", "dg-mk-unpicked"]) {
    $(id).onchange = () => {
      // 제외 사유는 손질 망(mm) 좌표다 — 켜면 그 좌표계로 화면을 맞춘다.
      fitDesignView();      // 같은 판단이 두 곳에 있으면 한쪽만 고쳐진다
      draw();
    };
  }

  $("btn-to-design").onclick = () => setStage("design");
  $("dg-back").onclick = () => setStage("conv");

  $("dg-build").onclick = async () => {
    busy(true, "최불리 선정과 표 확정 중…");
    try {
      const d = await post("/api/module-f/design/build",
                           { sid: S.sid, ...designSettings() });
      if (!d.ok) throw new Error(d.message || "확정 실패");
      watch(async () => {
        try {
          const j = await api(`/api/module-f/convert/result?sid=${S.sid}`);
          const sum = (j.result && j.result.summary) || null;
          if (!sum) throw new Error((j.result && j.result.error) || "확정 실패");
          // [F-11b-2] 재확정이 «성공한» 여기서만 배지를 내린다 — 위 throw 로
          //   빠지면 배지가 남아 「저장했지만 아직 안 들어갔다」를 말한다.
          S.ovDirty = false;
          renderDesignSummary(sum);
          await designPreview();
          $("dg-emit").disabled = false;
          say("표 확정 — 미리보기와 표는 저장될 값 그대로입니다.", "ok");
        } catch (err) { say(err.message, "err"); }
      });
    } catch (err) { busy(false); say(err.message, "err"); }
  };

  // 보기 설정이 바뀌면 preview 만 다시 — 최불리 재계산 없음(0.5초 규약).
  for (const id of ["dg-iso", "dg-zscale", "dg-canvas", "dg-ref", "dg-stub"]) {
    $(id).onchange = () => {
      if (S.design) designPreview().catch(err => say(err.message, "err"));
    };
  }
  $("dg-table").onchange = renderDesignTable;

  $("dg-emit").onclick = async () => {
    busy(true, ".sdf + .slf 저장 중…");
    try {
      const d = await post("/api/module-f/design/emit",
                           { sid: S.sid, ...designSettings() });
      if (!d.ok) throw new Error(d.message || "저장 실패");
      S.designDownload = d.download;
      $("dg-download").disabled = false;
      say(`저장 — ${d.sdf.name} (${d.sdf.bytes.toLocaleString()}B)`
        + ` + ${d.slf.name}. SDF 는 옆의 .slf 와 한 쌍입니다.`, "ok");
    } catch (err) { say(err.message, "err"); }
    finally { busy(false); }
  };
  $("dg-download").onclick = () => {
    if (S.designDownload) window.location.href = S.designDownload;
  };

  setStage("open");
  loadSaved();
  loadRefCounts();
})();

