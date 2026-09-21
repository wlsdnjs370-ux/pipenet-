
    // 레이어명은 업로드된 DXF 에서 온다 — 속성 보간 탈출을 막으려면 따옴표까지 5자 전부.
    function escHtml(s) {
      return String(s == null ? "" : s)
        .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;").replace(/'/g, "&#39;");
    }

    // ====== 상수 (지시서 §12.5) ======
    const ARCH_COLORS = {
      WALL: "#94a3b8", DOOR: "#fbbf24", WINDOW: "#67e8f9", COLUMN: "#a1a1aa",
      STAIR: "#c4b5fd", SHAFT: "#f472b6", ROOM_TEXT: "#22c55e", DIM: "#52525b",
      FURNITURE: "#3f3f46", GRID: "#3f3f46", BEAM: "#fb923c", OTHER: "#71717a",
    };
    const DESIGN_COLORS = {
      room: "#38bdf8", virtualEdge: "#f87171",   // ★ 경고색 — 검수 우선
      core: "#f472b6", zone: "#a78bfa", obstacle: "#fb923c", picked: "#fde047",
      pipeBranch: "#3b82f6", pipeCross: "#2563eb", pipeMain: "#1d4ed8",
      head: "#ef4444", valve: "#facc15",
    };

    // 미분류는 12종이 아니다. "아직 안 봤다" 와 "보고 나서 그 밖으로 판정했다(OTHER)" 를
    // 같은 칸에 넣으면 인식기를 붙였을 때 무엇이 남았는지 셀 수 없다.
    const UNCLASSIFIED = "__none__";
    const UNCLASSIFIED_COLOR = "#94a3b8";

    // 뒤 → 앞. 지시서 §12.4 의 배경(0.20) / 주요(0.45) / 문(0.6) 3단 alpha 를 따르고,
    // 표에 층이 명시되지 않은 BEAM·SHAFT 는 주요 구조로, ROOM_TEXT 는 텍스트라 최상단에 둔다.
    const ARCH_ORDER = [
      // 미분류가 출발 상태다. 여기가 어두우면 분류할 도면이 안 보인다.
      [UNCLASSIFIED, 0.60],
      ["GRID", 0.20], ["DIM", 0.20], ["FURNITURE", 0.20], ["OTHER", 0.20],
      ["WALL", 0.45], ["COLUMN", 0.45], ["STAIR", 0.45], ["WINDOW", 0.45],
      ["BEAM", 0.45], ["SHAFT", 0.45],
      ["DOOR", 0.60],
      ["ROOM_TEXT", 0.85],
    ];
    const ARCH_CATEGORIES = ARCH_ORDER.map(([c]) => c).filter((c) => c !== UNCLASSIFIED).sort();

    const STAGES = [
      ["c1", "C1 건축 인식"], ["gate", "GATE 확정"], ["c2", "C2 규범 조건"],
      ["c3", "C3 밸브·구역"], ["c4", "C4 헤드 배치"], ["c5", "C5 라우팅"],
      ["emit", "솔버 파일 방출"],
    ];

    // 설계 오버레이는 자료가 생기는 PR 에서 하나씩 늘린다. 그릴 것이 없는 항목을
    // 미리 켜 두면 "켰는데 아무것도 없다" 가 인식 실패인지 미구현인지 구분되지 않는다.
    const OVERLAYS = [
      ["rooms", "실 폴리곤", DESIGN_COLORS.room],
      ["roomLabels", "실 이름", DESIGN_COLORS.room],
      ["virtualEdges", "가상 폐합선", DESIGN_COLORS.virtualEdge],
      ["cores", "코어(계단·샤프트)", DESIGN_COLORS.core],
    ];

    // ====== 상태 (지시서 §12.3) ======
    const state = {
      entities: [], bbox: null, layers: [], layerState: {},
      view: { zoom: 1, panX: 0, panY: 0 },
      dpr: window.devicePixelRatio || 1,
      dxfFile: null, dxfToken: null, drag: null, fitZoom: null,

      // 실 편집(§12.6). `tool` 이 "pan" 이 아니면 캔버스 클릭이 편집 조작이 된다.
      tool: "pan", splitDraft: null, editBusy: false,

      // 인식기는 mm 로 답한다. 캔버스 좌표는 도면 단위 그대로라, 받는 자리에서
      // 한 번 되돌리지 않으면 m 단위 도면에서 오버레이가 1000배로 벌어진다.
      unitToMm: 1,
      wallLayers: [],
      sessionId: null,
      stage: "c1",
      gatePassed: false,
      centerlines: [],
      virtualEdges: [],
      rooms: [],
      cores: [],
      overlayVisible: { rooms: true, roomLabels: true, virtualEdges: true, cores: true },

      // 확정 게이트. `values` 는 아직 서버에 보내지 않은 사람의 입력이고, `missing`
      // 은 서버가 판정한 결손이다. 둘을 섞지 않아야 "화면은 다 찼는데 서버는 422"
      // 가 생기지 않는다.
      gate: {
        fields: [], specByField: {}, missing: {}, suggestion: {},
        values: {}, facts: {}, selection: new Set(), order: [], loaded: false,
      },
    };

    const canvas = document.getElementById("dw-canvas");
    const ctx = canvas.getContext("2d", { alpha: false });
    const emptyEl = document.getElementById("dw-empty");
    const overlayInfoEl = document.getElementById("dw-overlay-info");
    const overlayCursorEl = document.getElementById("dw-overlay-cursor");
    const dxfInputEl = document.getElementById("dw-dxf");
    const loadStatusEl = document.getElementById("dw-load-status");
    const layerListEl = document.getElementById("dw-layer-list");
    const unclassifiedNoteEl = document.getElementById("dw-unclassified-note");
    const designListEl = document.getElementById("dw-design-list");
    const wallFieldEl = document.getElementById("dw-wall-field");
    const wallSelectEl = document.getElementById("dw-wall-layers");
    const recognizeBtnEl = document.getElementById("dw-recognize-btn");
    const recognizeStatusEl = document.getElementById("dw-recognize-status");
    const recognizeStagesEl = document.getElementById("dw-recognize-stages");
    const stepperEl = document.getElementById("dw-stepper");
    const sessionEl = document.getElementById("dw-session");
    const gateOpenEl = document.getElementById("dw-gate-open");
    const gateProgressEl = document.getElementById("dw-gate-progress");
    const runC2El = document.getElementById("dw-run-c2");
    const gateStatusEl = document.getElementById("dw-gate-status");
    const gatePanelEl = document.getElementById("dw-gate-panel");
    const gateRemainEl = document.getElementById("dw-gate-remain");
    const gateCloseEl = document.getElementById("dw-gate-close");
    const gateFactsEl = document.getElementById("dw-gate-facts");
    const bulkFieldEl = document.getElementById("dw-bulk-field");
    const bulkValueSlotEl = document.getElementById("dw-bulk-value-slot");
    const bulkScopeEl = document.getElementById("dw-bulk-scope");
    const bulkApplyEl = document.getElementById("dw-bulk-apply");
    const bulkNoteEl = document.getElementById("dw-bulk-note");
    const gateTheadEl = document.getElementById("dw-gate-thead");
    const gateTbodyEl = document.getElementById("dw-gate-tbody");
    const gateOperatorEl = document.getElementById("dw-gate-operator");
    const gateMsgEl = document.getElementById("dw-gate-msg");
    const gateConfirmEl = document.getElementById("dw-gate-confirm");
    const fitBtnEl = document.getElementById("dw-fit-btn");
    const zoomInBtn = document.getElementById("dw-zoom-in");
    const zoomOutBtn = document.getElementById("dw-zoom-out");
    const zoomLevelEl = document.getElementById("dw-zoom-level");
    const toolPanEl = document.getElementById("dw-tool-pan");
    const toolSelectEl = document.getElementById("dw-tool-select");
    const toolSplitEl = document.getElementById("dw-tool-split");
    const mergeBtnEl = document.getElementById("dw-merge-btn");
    const deleteBtnEl = document.getElementById("dw-delete-btn");
    const editNoteEl = document.getElementById("dw-edit-note");

    // ====== Canvas 리사이즈 ======
    function resizeCanvas() {
      const rect = canvas.getBoundingClientRect();
      state.dpr = window.devicePixelRatio || 1;
      canvas.width = Math.round(rect.width * state.dpr);
      canvas.height = Math.round(rect.height * state.dpr);
      render();
    }
    window.addEventListener("resize", resizeCanvas);

    // ====== 뷰 변환 ======
    function fitToBBox() {
      if (!state.bbox) return;
      const rect = canvas.getBoundingClientRect();
      const w = rect.width, h = rect.height;
      const bw = state.bbox.x_max - state.bbox.x_min;
      const bh = state.bbox.y_max - state.bbox.y_min;
      if (bw <= 0 || bh <= 0) return;
      const margin = 30;
      const z = Math.min((w - 2 * margin) / bw, (h - 2 * margin) / bh);
      state.view.zoom = z;
      // DXF 좌표계: Y가 위쪽. Canvas 는 Y가 아래쪽. -z 로 뒤집기.
      state.view.panX = w / 2 - ((state.bbox.x_min + state.bbox.x_max) / 2) * z;
      state.view.panY = h / 2 + ((state.bbox.y_min + state.bbox.y_max) / 2) * z;
      state.fitZoom = z;
      render();
      updateZoomLevelDisplay();
    }
    fitBtnEl.addEventListener("click", fitToBBox);

    function zoomAtPoint(sx, sy, factor) {
      if (!state.entities.length) return;
      const [wx, wy] = screenToWorld(sx, sy);
      state.view.zoom *= factor;
      const [nsx, nsy] = worldToScreen(wx, wy);
      state.view.panX += sx - nsx;
      state.view.panY += sy - nsy;
      render();
      updateZoomLevelDisplay();
    }

    function zoomCenter(factor) {
      const rect = canvas.getBoundingClientRect();
      zoomAtPoint(rect.width / 2, rect.height / 2, factor);
    }

    function updateZoomLevelDisplay() {
      const ratio = state.fitZoom ? (state.view.zoom / state.fitZoom) : state.view.zoom;
      zoomLevelEl.textContent = (ratio * 100).toFixed(0) + "%";
    }

    zoomInBtn.addEventListener("click", () => zoomCenter(1.25));
    zoomOutBtn.addEventListener("click", () => zoomCenter(1 / 1.25));
    zoomLevelEl.addEventListener("click", fitToBBox);

    // 키보드 단축키: + 확대 / - 축소 / 0 화면 맞춤
    window.addEventListener("keydown", (e) => {
      if (e.target && (e.target.tagName === "INPUT" || e.target.tagName === "SELECT" || e.target.tagName === "TEXTAREA" || e.target.isContentEditable)) return;
      if (e.key === "+" || e.key === "=") { e.preventDefault(); zoomCenter(1.25); }
      else if (e.key === "-" || e.key === "_") { e.preventDefault(); zoomCenter(1 / 1.25); }
      else if (e.key === "0") { e.preventDefault(); fitToBBox(); }
    });

    function worldToScreen(x, y) {
      return [x * state.view.zoom + state.view.panX, -y * state.view.zoom + state.view.panY];
    }
    function screenToWorld(sx, sy) {
      return [(sx - state.view.panX) / state.view.zoom, -(sy - state.view.panY) / state.view.zoom];
    }

    // ====== 렌더링 ======
    function render() {
      const w = canvas.width / state.dpr;
      const h = canvas.height / state.dpr;
      ctx.setTransform(state.dpr, 0, 0, state.dpr, 0, 0);
      ctx.fillStyle = "#0f172a";
      ctx.fillRect(0, 0, w, h);
      if (!state.entities.length) return;

      const groups = {};
      for (const ent of state.entities) {
        const ls = state.layerState[ent.l];
        if (!ls || !ls.visible) continue;
        (groups[ls.category] || (groups[ls.category] = [])).push(ent);
      }
      ctx.lineWidth = 1;
      for (const [cat, alpha] of ARCH_ORDER) {
        const arr = groups[cat];
        if (!arr) continue;
        const color = cat === UNCLASSIFIED ? UNCLASSIFIED_COLOR : ARCH_COLORS[cat];
        ctx.strokeStyle = color;
        ctx.fillStyle = color;
        ctx.globalAlpha = alpha;
        ctx.lineWidth = 0.9;
        for (const ent of arr) drawEntity(ent);
      }
      ctx.globalAlpha = 1;

      drawDesignOverlays();
    }

    function drawDesignOverlays() {
      if (state.overlayVisible.rooms && state.rooms.length) {
        for (const room of state.rooms) {
          if (!room.polygon || room.polygon.length < 3) continue;
          const picked = state.gate.selection.has(room.id);
          ctx.strokeStyle = ctx.fillStyle = picked ? DESIGN_COLORS.picked : DESIGN_COLORS.room;
          ctx.lineWidth = picked ? 2.4 : 1.2;
          tracePoly(room.polygon);
          ctx.closePath();
          ctx.globalAlpha = picked ? 0.30 : 0.12;
          ctx.fill();
          ctx.globalAlpha = 0.9;
          ctx.stroke();
        }
        ctx.globalAlpha = 1;
      }

      // 아직 서버에 보내지 않은 자르는 선. 실제 경계가 아니므로 점선으로만 둔다.
      if (state.splitDraft) {
        const [ax, ay] = worldToScreen(state.splitDraft.p1[0], state.splitDraft.p1[1]);
        const [bx, by] = worldToScreen(state.splitDraft.p2[0], state.splitDraft.p2[1]);
        ctx.setLineDash([5, 4]);
        ctx.strokeStyle = DESIGN_COLORS.picked;
        ctx.lineWidth = 1.6;
        ctx.beginPath(); ctx.moveTo(ax, ay); ctx.lineTo(bx, by); ctx.stroke();
        ctx.setLineDash([]);
      }

      // ★ 가상 폐합선 — 실측이 아니라 알고리즘 추정이다. 점선 + 경고색으로
      // 실제 벽과 절대 섞이지 않게 그린다(§12.5). 여기가 오류의 최대 발생원이다.
      if (state.overlayVisible.virtualEdges && state.virtualEdges.length) {
        ctx.setLineDash([6, 4]);
        ctx.strokeStyle = DESIGN_COLORS.virtualEdge;
        ctx.lineWidth = 2.0;
        for (const edge of state.virtualEdges) {
          const [ax, ay] = worldToScreen(edge.p1[0], edge.p1[1]);
          const [bx, by] = worldToScreen(edge.p2[0], edge.p2[1]);
          ctx.beginPath(); ctx.moveTo(ax, ay); ctx.lineTo(bx, by); ctx.stroke();
        }
        ctx.setLineDash([]);
      }

      if (state.overlayVisible.cores && state.cores.length) {
        ctx.strokeStyle = DESIGN_COLORS.core;
        ctx.fillStyle = DESIGN_COLORS.core;
        ctx.lineWidth = 2.4;
        for (const core of state.cores) {
          if (!core.polygon || core.polygon.length < 3) continue;
          tracePoly(core.polygon);
          ctx.closePath();
          ctx.globalAlpha = 0.20;
          ctx.fill();
          ctx.globalAlpha = 1;
          ctx.stroke();
        }
      }

      if (state.overlayVisible.roomLabels && state.rooms.length && state.view.zoom > 0.005) {
        ctx.fillStyle = ARCH_COLORS.ROOM_TEXT;
        ctx.font = "11px ui-monospace, monospace";
        for (const room of state.rooms) {
          if (!room.polygon || !room.polygon.length) continue;
          let cx = 0, cy = 0;
          for (const p of room.polygon) { cx += p[0]; cy += p[1]; }
          const [sx, sy] = worldToScreen(cx / room.polygon.length, cy / room.polygon.length);
          ctx.fillText(room.name || room.id, sx, sy);
        }
      }
    }

    function signedArea2(pts) {
      let s = 0;
      for (let i = 0; i < pts.length; i++) {
        const a = pts[i], b = pts[(i + 1) % pts.length];
        s += a[0] * b[1] - b[0] * a[1];
      }
      return s;
    }

    function tracePoly(pts) {
      ctx.beginPath();
      for (let i = 0; i < pts.length; i++) {
        const [sx, sy] = worldToScreen(pts[i][0], pts[i][1]);
        if (i === 0) ctx.moveTo(sx, sy); else ctx.lineTo(sx, sy);
      }
    }

    function drawEntity(ent) {
      const z = state.view.zoom;
      if (ent.t === "L") {
        const [a, b] = worldToScreen(ent.p[0], ent.p[1]);
        const [c, d] = worldToScreen(ent.p[2], ent.p[3]);
        ctx.beginPath(); ctx.moveTo(a, b); ctx.lineTo(c, d); ctx.stroke();
      } else if (ent.t === "PL") {
        tracePoly(ent.p);
        ctx.stroke();
      } else if (ent.t === "A") {
        const [sx, sy] = worldToScreen(ent.c[0], ent.c[1]);
        const r = ent.r * z;
        if (r < 0.3) return;
        // DXF ARC 는 CCW. Canvas 는 CW 가 양수.
        const sa = ent.a[0] * Math.PI / 180;
        const ea = ent.a[1] * Math.PI / 180;
        ctx.beginPath();
        ctx.arc(sx, sy, r, -ea, -sa, false);
        ctx.stroke();
      } else if (ent.t === "C") {
        const [sx, sy] = worldToScreen(ent.c[0], ent.c[1]);
        const r = ent.r * z;
        if (r < 0.5) {
          ctx.fillRect(sx - 1, sy - 1, 2, 2);
          return;
        }
        ctx.beginPath();
        ctx.arc(sx, sy, r, 0, Math.PI * 2);
        ctx.stroke();
      } else if (ent.t === "H") {
        if (!ent.p || ent.p.length < 3) return;
        tracePoly(ent.p);
        ctx.closePath();
        const prev = ctx.globalAlpha;
        ctx.globalAlpha = prev * 0.25;
        ctx.fill();
        ctx.globalAlpha = prev;
        ctx.stroke();
      } else if (ent.t === "S") {
        if (!ent.p || ent.p.length < 3) return;
        // DXF SOLID/TRACE 는 정점을 0-1-3-2 나비 순서로 저장하고 3DFACE 는 0-1-2-3 이다.
        // 서버가 셋을 같은 "S" 로 방출해 구분이 없으므로, 자기교차하지 않는 순서
        // (= 부호면적 절댓값이 큰 쪽)를 실제 외곽으로 본다. 틀리면 기둥이 모래시계가 된다.
        let pts = ent.p;
        if (pts.length === 4) {
          const alt = [pts[0], pts[1], pts[3], pts[2]];
          if (Math.abs(signedArea2(pts)) < Math.abs(signedArea2(alt))) pts = alt;
        }
        tracePoly(pts);
        ctx.closePath();
        ctx.fill();
      } else if (ent.t === "I") {
        const [sx, sy] = worldToScreen(ent.p[0], ent.p[1]);
        ctx.beginPath();
        ctx.moveTo(sx, sy - 3); ctx.lineTo(sx + 3, sy); ctx.lineTo(sx, sy + 3); ctx.lineTo(sx - 3, sy);
        ctx.closePath();
        ctx.fill();
      } else if (ent.t === "T") {
        if (z < 0.005) return;
        const [sx, sy] = worldToScreen(ent.p[0], ent.p[1]);
        ctx.font = "10px ui-monospace, monospace";
        ctx.fillText(ent.v, sx, sy);
      }
    }

    // ====== 인터랙션: zoom + pan ======
    canvas.addEventListener("wheel", (e) => {
      e.preventDefault();
      const rect = canvas.getBoundingClientRect();
      zoomAtPoint(e.clientX - rect.left, e.clientY - rect.top, e.deltaY < 0 ? 1.15 : 1 / 1.15);
    }, { passive: false });

    canvas.addEventListener("mousedown", (e) => {
      const rect = canvas.getBoundingClientRect();
      if (state.tool === "split" && state.rooms.length) {
        const at = screenToWorld(e.clientX - rect.left, e.clientY - rect.top);
        state.splitDraft = { p1: at, p2: at };
        return;
      }
      // 끈 거리를 재 둔다. 선택 도구에서 화면을 옮기려던 손짓과 실을 고르는
      // 클릭을 나누는 유일한 단서다.
      state.drag = { x: e.clientX, y: e.clientY, panX: state.view.panX,
                     panY: state.view.panY, moved: false };
      canvas.classList.add("is-panning");
    });

    window.addEventListener("mouseup", (e) => {
      if (state.splitDraft) {
        const { p1, p2 } = state.splitDraft;
        state.splitDraft = null;
        render();
        submitSplit(p1, p2);
        return;
      }
      const drag = state.drag;
      state.drag = null;
      canvas.classList.remove("is-panning");
      if (drag && !drag.moved && state.tool === "select") pickRoomAt(e);
    });

    window.addEventListener("mousemove", (e) => {
      const rect = canvas.getBoundingClientRect();
      const inside = e.clientX >= rect.left && e.clientX <= rect.right && e.clientY >= rect.top && e.clientY <= rect.bottom;
      if (state.splitDraft) {
        state.splitDraft.p2 = screenToWorld(e.clientX - rect.left, e.clientY - rect.top);
        render();
      } else if (state.drag) {
        if (Math.abs(e.clientX - state.drag.x) + Math.abs(e.clientY - state.drag.y) > 3) {
          state.drag.moved = true;
        }
        state.view.panX = state.drag.panX + (e.clientX - state.drag.x);
        state.view.panY = state.drag.panY + (e.clientY - state.drag.y);
        render();
        return;
      }
      if (state.entities.length && inside) {
        const [wx, wy] = screenToWorld(e.clientX - rect.left, e.clientY - rect.top);
        overlayCursorEl.style.display = "inline-block";
        overlayCursorEl.textContent = `( ${wx.toFixed(0)} , ${wy.toFixed(0)} )`;
      } else {
        overlayCursorEl.style.display = "none";
      }
    });

    // ====== NDJSON 스트림 ======
    // inspect(§11.1 기존 형식)와 C1 인식이 같은 형식을 쓴다. 청크 경계가 줄
    // 한가운데를 자르므로 버퍼링은 한 곳에만 둔다.
    async function readNdjson(res, onMessage) {
      const reader = res.body.getReader();
      const decoder = new TextDecoder("utf-8");
      let buf = "";
      const feed = (line) => {
        const s = line.trim();
        if (!s) return;
        let msg;
        try { msg = JSON.parse(s); } catch (_) { return; }
        onMessage(msg);
      };
      for (;;) {
        const { done, value } = await reader.read();
        if (done) break;
        buf += decoder.decode(value, { stream: true });
        let nl;
        while ((nl = buf.indexOf("\n")) >= 0) {
          feed(buf.slice(0, nl));
          buf = buf.slice(nl + 1);
        }
      }
      if (buf.trim()) feed(buf);  // 마지막 줄(개행 없을 수 있음)
    }

    // ====== 세션 ======
    async function startSession() {
      try {
        const res = await fetch("/api/design/session", {
          method: "POST", headers: { "Content-Type": "application/json" }, body: "{}",
        });
        const body = await res.json();
        if (!res.ok || !body.session_id) throw new Error(body.message || `HTTP ${res.status}`);
        state.sessionId = body.session_id;
        sessionEl.textContent = "session " + body.session_id;
      } catch (err) {
        sessionEl.textContent = "세션 생성 실패 — " + err.message;
      }
    }

    function renderStepper() {
      const at = STAGES.findIndex(([id]) => id === state.stage);
      stepperEl.innerHTML = STAGES.map(([id, label], i) => {
        const cls = i < at ? "done" : (i === at ? "current" : "todo");
        return `<li class="${cls}"><span class="no">${String(i + 1).padStart(2, "0")}</span>${escHtml(label)}</li>`;
      }).join("");
    }

    // ====== DXF 업로드 + inspect ======
    dxfInputEl.addEventListener("change", async () => {
      const file = dxfInputEl.files[0];
      if (!file) return;
      state.dxfFile = file;
      loadStatusEl.textContent = `업로드 중... (${(file.size / 1024 / 1024).toFixed(1)} MB)`;

      const fd = new FormData();
      fd.append("dxf_file", file);
      try {
        const res = await fetch("/api/remote30/inspect", { method: "POST", body: fd });
        if (!res.ok) {
          let msg = `HTTP ${res.status}`;
          try { const j = await res.json(); if (j && j.message) msg = j.message; } catch (_) {}
          throw new Error(msg);
        }
        // NDJSON 스트림: {"type":"progress",...} 다수 + {"type":"result",...} 1개.
        const entities = [];
        let result = null;
        await readNdjson(res, (msg) => {
          if (msg.type === "progress") {
            if (Array.isArray(msg.entities)) {
              for (const e of msg.entities) entities.push(e);
            }
            if (msg.bbox) {
              state.entities = entities;
              state.bbox = msg.bbox;
              loadStatusEl.textContent = `로딩 중... entity ${entities.length.toLocaleString()}`;
              fitToBBox();
            }
          } else if (msg.type === "result") {
            result = msg;
          } else if (msg.ok === false) {
            throw new Error(msg.message || "DXF 분석 실패");
          }
        });
        if (!result) throw new Error("스트림에 result 메시지가 없습니다");
        if (result.ok === false) throw new Error(result.message || "DXF 분석 실패");

        state.entities = entities;
        state.bbox = result.bbox;
        state.layers = result.layers || [];
        state.dxfToken = result.dxf_token;
        state.layerState = {};
        for (const layer of state.layers) {
          // 12종 분류는 C1 인식기(PR-4)가 지문으로 판정한다. 기존 6종 자동분류는
          // 별개 체계라 그대로 옮기면 근거 없는 주장이 된다 → 미분류로 시작한다.
          // 표시 여부만 기존 판정(EXCLUDE 는 도면 잡음)을 그대로 쓴다.
          state.layerState[layer.name] = {
            visible: layer.auto_category !== "EXCLUDE",
            category: UNCLASSIFIED,
          };
        }
        const counts = result.counts || {};
        const totalEnt = counts.total_entities != null ? counts.total_entities : entities.length;
        const layerN = counts.layers != null ? counts.layers : state.layers.length;
        loadStatusEl.textContent = `${result.dxf_filename} — entity ${totalEnt}, layer ${layerN}`;
        overlayInfoEl.textContent = `${totalEnt} entities | ${layerN} layers`;
        emptyEl.style.display = "none";
        renderLayerList();
        renderWallChoices();
        fitToBBox();
      } catch (err) {
        loadStatusEl.textContent = "오류: " + err.message;
      }
    });

    // ====== C1 인식 (§11.1) ======
    // 인식기 좌표는 mm 다. 도면 단위로 되돌려 담아야 캔버스 오버레이가 겹친다.
    function toDrawingUnits(points) {
      const k = state.unitToMm || 1;
      return k === 1 ? points : points.map((p) => [p[0] / k, p[1] / k]);
    }

    function renderWallChoices(candidates) {
      // 후보를 주면(=인식이 막혔다) 그것만, 아니면 도면의 모든 레이어를 보여준다.
      const names = candidates ? candidates.map((c) => c.name) : state.layers.map((l) => l.name);
      const hint = new Map((candidates || []).map(
        (c) => [c.name, ` (평행쌍 ${(c.parallel_pair_ratio * 100).toFixed(0)}%)`]));
      wallSelectEl.innerHTML = names.map((name) => {
        const keep = state.wallLayers.includes(name) ? " selected" : "";
        return `<option value="${escHtml(name)}"${keep}>${escHtml(name + (hint.get(name) || ""))}</option>`;
      }).join("");
      wallFieldEl.style.display = names.length ? "grid" : "none";
      recognizeBtnEl.disabled = !state.dxfToken || !state.sessionId;
    }

    function renderStages(stages) {
      recognizeStagesEl.innerHTML = (stages || []).map((s) =>
        `<div class="stage-row"><span>${escHtml(s.name)} — ${escHtml(s.summary)}</span>
         <span class="secs">${s.seconds.toFixed(2)}s</span></div>`).join("");
    }

    recognizeBtnEl.addEventListener("click", async () => {
      if (!state.dxfToken || !state.sessionId) return;
      state.wallLayers = Array.from(wallSelectEl.selectedOptions).map((o) => o.value);
      state.rooms = []; state.cores = []; state.virtualEdges = [];
      renderDesignList();
      renderStages([]);
      recognizeBtnEl.disabled = true;
      recognizeStatusEl.classList.remove("warn");
      recognizeStatusEl.textContent = "인식 중…";
      try {
        const res = await fetch("/api/design/c1/recognize", {
          method: "POST", headers: { "Content-Type": "application/json" },
          body: JSON.stringify({
            session_id: state.sessionId, dxf_token: state.dxfToken,
            wall_layers: state.wallLayers,
          }),
        });
        if (!res.ok) {
          const body = await res.json().catch(() => null);
          throw new Error((body && body.message) || `HTTP ${res.status}`);
        }
        let result = null;
        await readNdjson(res, (msg) => {
          if (msg.unit) state.unitToMm = msg.unit.unit_to_mm || 1;
          if (msg.type === "parse") {
            recognizeStatusEl.textContent =
              `도면 읽는 중… ${msg.done.toLocaleString()} / ${msg.total.toLocaleString()}`;
          } else if (msg.type === "phase") {
            recognizeStatusEl.textContent = msg.cached ? "이전 인식 결과 재사용" : "인식 중…";
            if (msg.candidates) renderWallChoices(msg.candidates);
          } else if (msg.type === "virtual_edges") {
            state.virtualEdges = (msg.edges || []).map((e) => {
              const [p1, p2] = toDrawingUnits([e.p1, e.p2]);
              return { ...e, p1, p2 };
            });
          } else if (msg.type === "rooms") {
            state.rooms = (msg.rooms || []).map(
              (r) => ({ ...r, polygon: toDrawingUnits(r.polygon || []) }));
          } else if (msg.type === "cores") {
            state.cores = (msg.cores || []).map(
              (c) => ({ ...c, polygon: toDrawingUnits(c.polygon || []) }));
          } else if (msg.type === "result" || msg.type === "error") {
            result = msg;
          }
        });
        if (!result) throw new Error("스트림에 result 메시지가 없습니다");
        renderStages(result.stages);
        renderDesignList();
        render();
        if (!result.ok) throw new Error(result.message || "인식이 끝나지 않았습니다");
        state.stage = "gate";
        renderStepper();
        state.gate.values = {};
        state.gate.selection = new Set();
        state.gate.order = [];
        state.gate.facts = {};
        await loadGateItems();
        updateEditTools();
        setEditNote("실을 눌러 고른 뒤 합치기·자르기·지우기. 편집은 바로 저장됩니다.");
        const c = result.counts || {};
        recognizeStatusEl.textContent =
          `실 ${c.rooms}개 / 코어 ${c.cores}개 / 가상 폐합선 ${c.virtual_edges}개 — `
          + `${result.seconds.toFixed(1)}s (WALL ${result.wall_layers.join(", ") || "없음"}`
          + `, ${result.wall_source})`;
      } catch (err) {
        recognizeStatusEl.classList.add("warn");
        recognizeStatusEl.textContent = "오류: " + err.message;
      } finally {
        recognizeBtnEl.disabled = false;
      }
    });

    // ====== 실 편집 (§12.6) ======
    // 인식기가 뽑은 face 가 늘 맞지는 않는다. 그 수정은 GATE 안에서만 할 수 있고,
    // 고친 결과가 그대로 C4 의 헤드 배치 면적이 된다. 그래서 폴리곤 수술은 화면이
    // 흉내 내지 않는다 — 서버가 고친 실을 돌려주고 화면은 그것만 그린다.
    function pointInPolygon(x, y, poly) {
      let inside = false;
      for (let i = 0, j = poly.length - 1; i < poly.length; j = i++) {
        const [xi, yi] = poly[i];
        const [xj, yj] = poly[j];
        if ((yi > y) !== (yj > y) && x < ((xj - xi) * (y - yi)) / (yj - yi) + xi) {
          inside = !inside;
        }
      }
      return inside;
    }

    /** 겹친 실은 작은 쪽을 고른다. 큰 실 안에 든 작은 실은 달리 집을 방법이 없다. */
    function roomAt(wx, wy) {
      let best = null;
      for (const room of state.rooms) {
        if (!room.polygon || room.polygon.length < 3) continue;
        if (!pointInPolygon(wx, wy, room.polygon)) continue;
        if (!best || (room.area_m2 || 0) < (best.area_m2 || 0)) best = room;
      }
      return best;
    }

    function toMm(point) {
      const k = state.unitToMm || 1;
      return [point[0] * k, point[1] * k];
    }

    function setEditNote(text, warn) {
      editNoteEl.textContent = text;
      editNoteEl.classList.toggle("warn", !!warn);
    }

    function editableNow() {
      return !!(state.sessionId && state.rooms.length && !state.gatePassed);
    }

    function updateEditTools() {
      const ready = editableNow();
      const n = state.gate.selection.size;
      toolSelectEl.disabled = !ready;
      toolSplitEl.disabled = !ready;
      mergeBtnEl.disabled = !ready || n < 2 || state.editBusy;
      deleteBtnEl.disabled = !ready || n < 1 || state.editBusy;
      if (!ready && state.tool !== "pan") setTool("pan");
    }

    function setTool(tool) {
      state.tool = tool;
      for (const [el, id] of [[toolPanEl, "pan"], [toolSelectEl, "select"], [toolSplitEl, "split"]]) {
        el.classList.toggle("is-on", tool === id);
      }
      canvas.classList.toggle("is-picking", tool !== "pan");
      state.splitDraft = null;
      render();
    }

    toolPanEl.addEventListener("click", () => setTool("pan"));
    toolSelectEl.addEventListener("click", () => setTool("select"));
    toolSplitEl.addEventListener("click", () => setTool("split"));

    function pickRoomAt(ev) {
      const rect = canvas.getBoundingClientRect();
      const [wx, wy] = screenToWorld(ev.clientX - rect.left, ev.clientY - rect.top);
      const room = roomAt(wx, wy);
      if (!room) {
        if (!ev.shiftKey) state.gate.selection = new Set();
      } else if (ev.shiftKey) {
        if (state.gate.selection.has(room.id)) state.gate.selection.delete(room.id);
        else state.gate.selection.add(room.id);
      } else {
        state.gate.selection = new Set([room.id]);
      }
      afterSelectionChange();
      setEditNote(room
        ? `${room.name || room.id} — ${(room.area_m2 || 0).toFixed(1)}㎡ (선택 ${state.gate.selection.size}개)`
        : `선택 ${state.gate.selection.size}개`, false);
    }

    function afterSelectionChange() {
      updateEditTools();
      if (gatePanelEl.classList.contains("open")) {
        renderGateTable();
        gateSyncControls();
      }
      render();
    }

    const EDIT_FIELD_LABEL = (field) => {
      const spec = state.gate.specByField[field];
      return spec ? spec.label : field;
    };

    function editSummary(rec) {
      if (rec.op === "merge") {
        const lost = (rec.cleared || []).map(EDIT_FIELD_LABEL).join(", ");
        return `${rec.rooms.length}개 실을 ${rec.into} 로 합쳤습니다 — ${rec.area_m2}㎡`
          + (lost ? `. 값이 달라 비운 항목: ${lost} (다시 확정해야 합니다)` : "");
      }
      if (rec.op === "split") {
        return `${rec.room} 을 ${rec.into.join(" / ")} 로 잘랐습니다 — `
          + `${rec.area_m2.map((a) => a + "㎡").join(" / ")}`;
      }
      return `${rec.room} 을 지웠습니다 — ${rec.area_m2}㎡`;
    }

    /** 서버가 돌려준 실로 통째로 갈아 끼운다. 사라진 실에 걸려 있던 화면 입력을
     *  남겨 두면 없는 실을 가리키는 값이 확정에 실린다. */
    function adoptRooms(rooms) {
      state.rooms = (rooms || []).map(
        (r) => ({ ...r, polygon: toDrawingUnits(r.polygon || []) }));
      const alive = new Set(state.rooms.map((r) => r.id));
      for (const id of Object.keys(state.gate.values)) {
        if (!alive.has(id)) delete state.gate.values[id];
      }
      state.gate.order = state.gate.order.filter((id) => alive.has(id));
      return alive;
    }

    async function postEdit(edit) {
      if (state.editBusy || !state.sessionId) return false;
      state.editBusy = true;
      updateEditTools();
      setEditNote("반영 중…", false);
      try {
        const res = await fetch("/api/design/gate/edit", {
          method: "POST", headers: { "Content-Type": "application/json" },
          body: JSON.stringify({
            session_id: state.sessionId, operator: gateOperatorEl.value.trim(), edit,
          }),
        });
        const body = await res.json().catch(() => null);
        if (!res.ok || !body || !body.ok) {
          setEditNote((body && body.message) || `HTTP ${res.status}`, true);
          return false;
        }
        const alive = adoptRooms(body.rooms);
        // 고친 실을 그대로 골라 둔다 — 합친 실의 용도를 바로 다시 정해야 한다.
        const rec = body.edit;
        const next = rec.op === "merge" ? [rec.into] : (rec.op === "split" ? rec.into : []);
        state.gate.selection = new Set(next.filter((id) => alive.has(id)));
        renderDesignList();
        await loadGateItems();
        updateEditTools();
        render();
        setEditNote(editSummary(rec), false);
        return true;
      } catch (err) {
        setEditNote("오류: " + err.message, true);
        return false;
      } finally {
        state.editBusy = false;
        updateEditTools();
      }
    }

    function submitSplit(p1, p2) {
      if (!editableNow()) return;
      const px = Math.hypot(p2[0] - p1[0], p2[1] - p1[1]) * state.view.zoom;
      if (px < 8) return;                       // 클릭에 가깝다 — 자를 뜻이 아니다
      const room = roomAt((p1[0] + p2[0]) / 2, (p1[1] + p2[1]) / 2);
      if (!room) {
        setEditNote("자를 실을 가로질러 선을 그으세요. 선의 가운데가 실 안에 있어야 합니다.", true);
        return;
      }
      postEdit({ op: "split", room: room.id, line: [toMm(p1), toMm(p2)] });
    }

    mergeBtnEl.addEventListener("click", () => {
      postEdit({ op: "merge", rooms: [...state.gate.selection] });
    });

    async function deleteSelected() {
      for (const id of [...state.gate.selection]) {
        if (!await postEdit({ op: "delete", room: id })) return;
      }
    }
    deleteBtnEl.addEventListener("click", deleteSelected);

    window.addEventListener("keydown", (e) => {
      if (e.target && (e.target.tagName === "INPUT" || e.target.tagName === "SELECT"
        || e.target.tagName === "TEXTAREA" || e.target.isContentEditable)) return;
      if (e.key !== "Delete" || !state.gate.selection.size || !editableNow()) return;
      e.preventDefault();
      deleteSelected();
    });

    // ====== GATE 인간 확정 (§4 · §12.8) ======
    // 질의를 한 지점에 모은다. 여기서 결손이 0 이 되면 이후는 사람을 부르지 않는다.
    // 규칙(무엇이 무엇에 걸리는지, 무엇이 필수인지)은 전부 서버가 준 `fields` 에서
    // 읽는다 — 화면이 같은 규칙을 다시 구현하면 두 곳이 어긋난다.
    const GATE_BOOL_LABELS = {
      "ceiling.has_finish": ["있음", "없음"],
      "confirmed": ["입상관으로 쓴다", "쓰지 않는다"],
    };
    const GATE_FACT_SCOPES = ["building", "obstacles"];

    function gateSpecs(scope) {
      return state.gate.fields.filter((f) => f.scope === scope);
    }

    function gateRawValue(obj, path) {
      return path.split(".").reduce((o, k) => (o == null ? undefined : o[k]), obj);
    }

    function gateSetRaw(obj, path, value) {
      const keys = path.split(".");
      const last = keys.pop();
      const target = keys.reduce((o, k) => (o[k] || (o[k] = {})), obj);
      target[last] = value;
    }

    function gateLocal(targetId, field) {
      const row = state.gate.values[targetId];
      return row ? row[field] : undefined;
    }

    function gateSetLocal(targetId, field, value) {
      if (value === undefined) {
        const row = state.gate.values[targetId];
        if (!row) return;
        delete row[field];
        if (!Object.keys(row).length) delete state.gate.values[targetId];
        return;
      }
      (state.gate.values[targetId] || (state.gate.values[targetId] = {}))[field] = value;
    }

    function gateMissing(field, targetId) {
      const set = state.gate.missing[field];
      return !!set && set.has(targetId);
    }

    function gateSuggestion(field, targetId) {
      const byId = state.gate.suggestion[field];
      return byId ? byId[targetId] : undefined;
    }

    // 저장값 기준(`useLocal=false`)과 화면 기준(`true`)을 나눠 본다. 반자 유무를
    // 화면에서 방금 "있음" 으로 바꾼 실은 서버가 아직 반자고를 묻지도 않았으므로
    // 결손 목록에 없다 — 그걸 채운 것으로 세면 "0건 남음" 인데 422 가 돌아온다.
    function gateApplies(room, spec, useLocal) {
      const dep = spec.applies_when;
      if (!dep) return true;
      const stored = gateRawValue(room, dep.field);
      const local = gateLocal(room.id, dep.field);
      return (useLocal && local !== undefined ? local : stored) === dep.value;
    }

    function gateNeeds(room, spec) {
      if (!spec.required || !gateApplies(room, spec, true)) return false;
      if (gateLocal(room.id, spec.field) !== undefined) return false;
      if (!gateApplies(room, spec, false)) return true;
      return gateMissing(spec.field, room.id);
    }

    function gateFactNeeds(field) {
      return gateLocal(field, field) === undefined && gateMissing(field, field);
    }

    function gateCoreNeeds(core) {
      return gateLocal(core.id, "confirmed") === undefined
        && gateMissing("confirmed", core.id);
    }

    /** 칸에 보일 값. 결손 칸은 비워 둔다 — 제안값을 미리 골라 두면 사람이 보지
     *  않고 확정을 눌러도 근거 없는 값이 통과한다. 제안은 힌트로만 적는다. */
    function gateDisplay(targetId, spec, stored, missing) {
      const local = gateLocal(targetId, spec.field);
      if (local !== undefined) return local;
      if (missing) return undefined;
      return stored == null ? undefined : stored;
    }

    function gateOptionLabel(field, option) {
      if (option === null) return "없음";
      if (typeof option === "boolean") {
        const labels = GATE_BOOL_LABELS[field] || ["예", "아니오"];
        return option ? labels[0] : labels[1];
      }
      return String(option);
    }

    function gateOptionsHtml(spec) {
      return `<option value="">— 미정 —</option>` + spec.options.map(
        (o, i) => `<option value="${i}">${escHtml(gateOptionLabel(spec.field, o))}</option>`
      ).join("");
    }

    // 같은 항목의 <option> 문자열은 실 수와 무관하게 한 번만 만든다. 값은 DOM 을
    // 세운 뒤 한 번에 넣는다(gateSyncControls) — 실이 600개면 문자열 조립 비용이
    // 그대로 화면 멈춤이 된다.
    const gateOptionCache = new Map();
    function gateControl(spec, targetId, cls) {
      const attrs = `data-gate="1" data-target="${escHtml(targetId)}" `
        + `data-field="${escHtml(spec.field)}"${cls ? ` class="${cls}"` : ""}`;
      if (!spec.options) return `<input type="number" step="any" ${attrs}>`;
      if (!gateOptionCache.has(spec.field)) {
        gateOptionCache.set(spec.field, gateOptionsHtml(spec));
      }
      return `<select ${attrs}>${gateOptionCache.get(spec.field)}</select>`;
    }

    function gateEncode(spec, value) {
      if (!spec.options) return value === undefined || value === null ? "" : String(value);
      if (value === undefined) return "";
      const index = spec.options.findIndex((o) => o === value);
      return index < 0 ? "" : String(index);
    }

    function gateDecode(spec, raw) {
      if (raw === "") return undefined;
      if (spec.options) return spec.options[Number(raw)];
      const num = Number(raw.trim());
      return Number.isFinite(num) ? num : undefined;
    }

    function gateStoredFact(field) {
      return state.gate.facts[field];
    }

    function gateOrderedRooms() {
      if (!state.gate.order.length) return state.rooms;
      const rank = new Map(state.gate.order.map((id, i) => [id, i]));
      return state.rooms.slice().sort(
        (a, b) => (rank.has(a.id) ? rank.get(a.id) : 1e9)
                - (rank.has(b.id) ? rank.get(b.id) : 1e9));
    }

    function renderGateFacts() {
      const cells = [];
      for (const scope of GATE_FACT_SCOPES) {
        for (const spec of gateSpecs(scope)) {
          cells.push(`<label class="cell" title="${escHtml(spec.note)}">
            <span>${escHtml(spec.label)}</span>${gateControl(spec, spec.field, "")}</label>`);
        }
      }
      const coreSpec = gateSpecs("core")[0];
      if (coreSpec) {
        for (const core of state.cores) {
          const sugg = gateSuggestion("confirmed", core.id);
          const hint = sugg ? ` ${(sugg.confidence * 100).toFixed(0)}%` : "";
          cells.push(`<label class="cell" title="${escHtml(coreSpec.note)}">
            <span>${escHtml(`${core.id} ${core.kind}${hint}`)}</span>
            ${gateControl(coreSpec, core.id, "")}</label>`);
        }
      }
      gateFactsEl.innerHTML = cells.join("");
    }

    function renderGateTable() {
      const specs = gateSpecs("room");
      gateTheadEl.innerHTML = `<tr><th><input type="checkbox" id="dw-gate-all"></th>`
        + `<th>실</th><th>층</th>`
        + specs.map((s) => `<th title="${escHtml(s.note)}">${escHtml(s.label)}`
          + `${s.required ? "" : " (선택)"}</th>`).join("") + `</tr>`;

      const rows = gateOrderedRooms().map((room) => {
        const checked = state.gate.selection.has(room.id) ? " checked" : "";
        const cells = specs.map((spec) => {
          if (!gateApplies(room, spec, true)) return `<td class="na">—</td>`;
          const sugg = gateSuggestion(spec.field, room.id);
          const hint = sugg === undefined ? "" : `<span class="sugg">제안 `
            + `${escHtml(gateOptionLabel(spec.field, sugg.value))} `
            + `${(sugg.confidence * 100).toFixed(0)}%</span>`;
          return `<td>${gateControl(spec, room.id, "")}${hint}</td>`;
        }).join("");
        return `<tr><td><input type="checkbox" data-select="${escHtml(room.id)}"${checked}></td>`
          + `<td>${escHtml(room.name || "")}<span class="room-id">${escHtml(room.id)}</span></td>`
          + `<td>${escHtml(room.floor || "")}</td>${cells}</tr>`;
      });
      gateTbodyEl.innerHTML = rows.join("");
    }

    /** 값과 "아직 미확정" 표시를 한 번에 맞춘다. 마크업을 다시 세우지 않으므로
     *  입력 중인 칸이 사라지지 않고, 표를 다시 그리지 않아도 붉은 칸이 풀린다. */
    function gateSyncControls() {
      const byId = new Map(state.rooms.map((r) => [r.id, r]));
      const coreById = new Map(state.cores.map((c) => [c.id, c]));
      for (const el of gatePanelEl.querySelectorAll("[data-gate]")) {
        const { target, field } = el.dataset;
        const spec = state.gate.specByField[field];
        const room = byId.get(target);
        const core = coreById.get(target);
        let stored;
        let missing;
        if (room) {
          stored = gateRawValue(room, field);
          missing = gateNeeds(room, spec);
        } else if (core) {
          stored = core.confirmed;
          missing = gateCoreNeeds(core);
        } else {
          stored = gateStoredFact(field);
          missing = gateFactNeeds(field);
        }
        // 지금 사람이 쓰고 있는 칸에는 값을 다시 넣지 않는다. 프로그램이 value 를
        // 덮으면 브라우저가 그 칸의 change 를 더 이상 내지 않아, 방금 친 숫자가
        // 다른 칸을 건드린 순간 조용히 사라진다.
        if (el !== document.activeElement) {
          el.value = gateEncode(spec, gateDisplay(target, spec, stored, missing));
        }
        el.parentElement.classList.toggle("miss", missing);
      }
    }

    function gateRemaining() {
      let need = 0;
      let total = 0;
      for (const room of state.rooms) {
        for (const spec of gateSpecs("room")) {
          if (!spec.required || !gateApplies(room, spec, true)) continue;
          total += 1;
          if (gateNeeds(room, spec)) need += 1;
        }
      }
      for (const scope of GATE_FACT_SCOPES) {
        for (const spec of gateSpecs(scope)) {
          if (!spec.required) continue;
          total += 1;
          if (gateFactNeeds(spec.field)) need += 1;
        }
      }
      if (gateSpecs("core").length) {
        total += state.cores.length;
        need += state.cores.filter(gateCoreNeeds).length;
      }
      return { need, total };
    }

    function renderGateProgress() {
      const { need, total } = gateRemaining();
      gateRemainEl.textContent = `${total}건 중 ${need}건 남음`;
      gateRemainEl.className = `remain ${need ? "pending" : "done"}`;
      gateProgressEl.textContent = state.gatePassed
        ? "게이트 통과 — 이후 단계는 사람을 부르지 않는다."
        : `${total}건 중 ${need}건 남음`;
      runC2El.disabled = !state.gatePassed;
    }

    function renderGate() {
      renderGateFacts();
      renderGateTable();
      renderBulkBar();
      gateSyncControls();
      renderGateProgress();
    }

    // ── 일괄 적용 (§12.8) — 이게 없으면 실이 200개인 도면에서 사용자가 포기한다.
    /** 용도별 일괄 적용의 기준. 아직 확정 안 된 실은 제안값이 아니라 화면 입력만 본다. */
    function gateEffectiveUse(room) {
      const spec = state.gate.specByField["use"];
      return spec && gateDisplay(room.id, spec, room.use, gateMissing("use", room.id));
    }

    function renderBulkBar() {
      const specs = gateSpecs("room");
      if (bulkFieldEl.options.length !== specs.length) {
        bulkFieldEl.innerHTML = specs.map(
          (s) => `<option value="${escHtml(s.field)}">${escHtml(s.label)}</option>`).join("");
      }
      const spec = state.gate.specByField[bulkFieldEl.value] || specs[0];
      if (!spec) return;
      bulkFieldEl.value = spec.field;

      const hasSuggestion = !!state.gate.suggestion[spec.field];
      bulkValueSlotEl.innerHTML = spec.options
        ? `<select id="dw-bulk-value">${hasSuggestion
            ? `<option value="sugg">제안값 그대로</option>` : ""}${gateOptionsHtml(spec)}</select>`
        : `<input id="dw-bulk-value" type="number" step="any" placeholder="값">`;

      const floors = [...new Set(state.rooms.map((r) => r.floor).filter(Boolean))];
      const uses = [...new Set(state.rooms.map(gateEffectiveUse).filter(Boolean))];
      bulkScopeEl.innerHTML = [`<option value="all">전체 실</option>`,
        `<option value="sel">선택한 실</option>`]
        .concat(spec.bulk_apply.includes("floor")
          ? floors.map((f) => `<option value="floor:${escHtml(f)}">${escHtml(f)} 전체</option>`) : [])
        .concat(spec.bulk_apply.includes("use")
          ? uses.map((u) => `<option value="use:${escHtml(u)}">용도 ${escHtml(u)} 전체</option>`) : [])
        .join("");
      bulkNoteEl.textContent = spec.note || "";
    }

    function bulkTargets(scope) {
      if (scope === "all") return state.rooms;
      if (scope === "sel") return state.rooms.filter((r) => state.gate.selection.has(r.id));
      const [kind, key] = [scope.slice(0, scope.indexOf(":")), scope.slice(scope.indexOf(":") + 1)];
      if (kind === "floor") return state.rooms.filter((r) => r.floor === key);
      return state.rooms.filter((r) => gateEffectiveUse(r) === key);
    }

    bulkFieldEl.addEventListener("change", () => { renderBulkBar(); gateSyncControls(); });

    bulkApplyEl.addEventListener("click", () => {
      const spec = state.gate.specByField[bulkFieldEl.value];
      const valueEl = document.getElementById("dw-bulk-value");
      if (!spec || !valueEl) return;
      const useSuggestion = valueEl.value === "sugg";
      const value = useSuggestion ? undefined : gateDecode(spec, valueEl.value);
      if (!useSuggestion && value === undefined) {
        bulkNoteEl.textContent = "적용할 값을 먼저 고르세요.";
        return;
      }
      let n = 0;
      for (const room of bulkTargets(bulkScopeEl.value)) {
        if (!gateApplies(room, spec, true)) continue;
        const sugg = useSuggestion ? gateSuggestion(spec.field, room.id) : null;
        if (useSuggestion && !sugg) continue;
        gateSetLocal(room.id, spec.field, useSuggestion ? sugg.value : value);
        n += 1;
      }
      renderGate();
      bulkNoteEl.textContent = `${n}개 실에 적용`;
    });

    // ── 입력 반영 ────────────────────────────────────────────────────────
    gatePanelEl.addEventListener("change", (ev) => {
      const el = ev.target;
      if (el.dataset && el.dataset.select !== undefined) {
        if (el.checked) state.gate.selection.add(el.dataset.select);
        else state.gate.selection.delete(el.dataset.select);
        updateEditTools();
        render();
        return;
      }
      if (el.id === "dw-gate-all") {
        state.gate.selection = new Set(el.checked ? state.rooms.map((r) => r.id) : []);
        renderGateTable();
        gateSyncControls();
        updateEditTools();
        render();
        return;
      }
      if (!el.dataset || !el.dataset.gate) return;
      const spec = state.gate.specByField[el.dataset.field];
      gateSetLocal(el.dataset.target, el.dataset.field, gateDecode(spec, el.value));
      // 반자 유무를 바꾸면 반자고 칸이 생기거나 사라진다. 표 전체를 다시 그린다.
      if (state.gate.fields.some((f) => f.applies_when
          && f.applies_when.field === el.dataset.field)) {
        renderGateTable();
      }
      gateSyncControls();
      renderGateProgress();
    });

    // ── 결손 항목 적재 ──────────────────────────────────────────────────
    async function loadGateItems() {
      if (!state.sessionId) return;
      let res;
      let body;
      try {
        res = await fetch(`/api/design/c1/gate_items/${state.sessionId}`);
        body = await res.json();
      } catch (err) {
        body = { message: "결손 항목을 읽지 못했습니다: " + err.message };
      }
      if (!res || !res.ok || !body || !body.ok) {
        state.gate.loaded = false;
        gateOpenEl.disabled = true;
        gateProgressEl.textContent = (body && body.message) || "결손 항목을 읽지 못했습니다.";
        return;
      }
      state.gate.fields = body.fields || [];
      state.gate.specByField = Object.fromEntries(state.gate.fields.map((f) => [f.field, f]));
      state.gate.missing = {};
      state.gate.suggestion = {};
      for (const group of body.groups || []) {
        state.gate.missing[group.field] = new Set(group.targets || []);
        if (group.suggestion) state.gate.suggestion[group.field] = group.suggestion;
      }
      const useGroup = (body.groups || []).find((g) => g.field === "use");
      if (!state.gate.order.length && useGroup) state.gate.order = useGroup.targets.slice();
      state.gate.loaded = true;
      gateOptionCache.clear();
      gateOpenEl.disabled = false;
      renderGate();
    }

    // ── 확정 ─────────────────────────────────────────────────────────────
    function gatePayload() {
      // 화면에서 더 이상 묻지 않는 값(반자를 "없음" 으로 되돌린 실의 반자고)은
      // 보내지 않는다. 서버는 받으면 그대로 쓰므로 아무도 확정하지 않은 값이 남는다.
      const byId = new Map(state.rooms.map((r) => [r.id, r]));
      const out = {};
      for (const [targetId, fields] of Object.entries(state.gate.values)) {
        const room = byId.get(targetId);
        const kept = {};
        for (const [field, value] of Object.entries(fields)) {
          if (room && !gateApplies(room, state.gate.specByField[field], true)) continue;
          kept[field] = value;
        }
        if (Object.keys(kept).length) out[targetId] = kept;
      }
      return out;
    }

    function gateCommit(values, defaults) {
      const byId = new Map(state.rooms.map((r) => [r.id, r]));
      const coreById = new Map(state.cores.map((c) => [c.id, c]));
      for (const [targetId, fields] of Object.entries(values)) {
        const room = byId.get(targetId);
        const core = coreById.get(targetId);
        for (const [field, value] of Object.entries(fields)) {
          if (room) {
            gateSetRaw(room, field, value);
            room.provenance[field] = "GATE";
          } else if (core) {
            core.confirmed = !!value;
          } else {
            state.gate.facts[field] = value;
          }
        }
      }
      // 서버가 근거를 들어 채운 값(천장고=층고). 화면이 모르면 결손이 아닌데
      // 비어 있는 칸이 생긴다.
      for (const applied of defaults || []) {
        const room = byId.get(applied.room);
        if (!room) continue;
        gateSetRaw(room, applied.field, applied.value);
        room.provenance[applied.field] = "default";
      }
      state.gate.values = {};
    }

    gateConfirmEl.addEventListener("click", async () => {
      const operator = gateOperatorEl.value.trim();
      gateMsgEl.classList.remove("warn");
      if (!operator) {
        gateMsgEl.classList.add("warn");
        gateMsgEl.textContent = "확정자 이름을 적어야 합니다 — 감사 기록에 남습니다.";
        gateOperatorEl.focus();
        return;
      }
      const values = gatePayload();
      gateConfirmEl.disabled = true;
      gateMsgEl.textContent = "확정 중…";
      try {
        const res = await fetch("/api/design/gate/confirm", {
          method: "POST", headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ session_id: state.sessionId, operator, values }),
        });
        const body = await res.json().catch(() => null);
        if (res.status === 409) {
          gateMsgEl.classList.add("warn");
          gateMsgEl.textContent = (body && body.message) || "다른 곳에서 먼저 저장했습니다.";
          return;
        }
        if (!body) throw new Error(`HTTP ${res.status}`);
        // 422 여도 서버는 이미 저장했다. 로컬 입력을 안 비우면 다음 확정에서 두 번
        // 보내고, 결손 판정도 옛 응답 기준으로 남는다.
        gateCommit(values, body.defaults);
        await loadGateItems();
        if (res.ok && body.passed) {
          state.gatePassed = true;
          state.stage = "c2";
          updateEditTools();
          setEditNote("게이트를 통과해 실을 더 고칠 수 없습니다.");
          renderStepper();
          renderGateProgress();
          gatePanelEl.classList.remove("open");
          gateMsgEl.textContent = "";
          gateStatusEl.classList.remove("warn");
          gateStatusEl.textContent = `게이트 통과 (${operator}, v${body.version})`;
          return;
        }
        gateMsgEl.classList.add("warn");
        gateMsgEl.textContent = body.code === "GATE_INCOMPLETE"
          ? `아직 ${body.unresolved.length}건이 확정되지 않았습니다.`
          : (body.message || "확정하지 못했습니다.");
      } catch (err) {
        gateMsgEl.classList.add("warn");
        gateMsgEl.textContent = "오류: " + err.message;
      } finally {
        gateConfirmEl.disabled = false;
      }
    });

    gateOpenEl.addEventListener("click", () => {
      if (!state.gate.loaded) return;
      gatePanelEl.classList.add("open");
      renderGate();
    });
    gateCloseEl.addEventListener("click", () => gatePanelEl.classList.remove("open"));
    gatePanelEl.addEventListener("mousedown", (ev) => {
      if (ev.target === gatePanelEl) gatePanelEl.classList.remove("open");
    });

    // C2 는 아직 없다(PR-6). 501 을 그대로 보여준다 — 통과한 척하면 화면이 진행된
    // 것으로 오해한다.
    runC2El.addEventListener("click", async () => {
      runC2El.disabled = true;
      gateStatusEl.classList.remove("warn");
      gateStatusEl.textContent = "C2 실행 중…";
      try {
        const res = await fetch("/api/design/c2/constraints", {
          method: "POST", headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ session_id: state.sessionId }),
        });
        const body = await res.json().catch(() => null);
        gateStatusEl.classList.toggle("warn", !res.ok);
        gateStatusEl.textContent = (body && body.message) || `HTTP ${res.status}`;
      } catch (err) {
        gateStatusEl.classList.add("warn");
        gateStatusEl.textContent = "오류: " + err.message;
      } finally {
        runC2El.disabled = !state.gatePassed;
      }
    });

    // ====== 건축 12종 토글 ======
    function renderLayerList() {
      if (!state.layers.length) return;
      const byCat = {};
      for (const layer of state.layers) {
        const cat = state.layerState[layer.name].category;
        (byCat[cat] || (byCat[cat] = [])).push(layer);
      }
      const opts = (sel) => [[UNCLASSIFIED, "미분류"], ...ARCH_CATEGORIES.map((c) => [c, c])]
        .map(([v, t]) => `<option value="${v}" ${sel === v ? "selected" : ""}>${t}</option>`).join("");

      const blocks = [UNCLASSIFIED, ...ARCH_CATEGORIES].map((cat) => {
        const layers = byCat[cat] || [];
        const entN = layers.reduce((s, l) => s + l.count, 0);
        const allOn = layers.length > 0 && layers.every((l) => state.layerState[l.name].visible);
        const title = cat === UNCLASSIFIED ? "미분류" : cat;
        const color = cat === UNCLASSIFIED ? UNCLASSIFIED_COLOR : ARCH_COLORS[cat];
        const head = `<div class="cat-head ${layers.length ? "" : "empty"}">
          <input type="checkbox" data-cat="${cat}" ${allOn ? "checked" : ""} ${layers.length ? "" : "disabled"}>
          <i class="swatch" style="background:${color}"></i>
          <span>${title}</span>
          <span class="count">${layers.length} layer / ${entN.toLocaleString()} ent</span>
        </div>`;
        const rows = layers.map((l) => {
          const ls = state.layerState[l.name];
          const safeName = escHtml(l.name);
          const key = encodeURIComponent(l.name);
          return `<div class="layer-row">
            <input type="checkbox" data-layer="${key}" ${ls.visible ? "checked" : ""}>
            <div>
              <div class="name" title="${safeName}">${safeName}</div>
              <div class="count">${l.count.toLocaleString()} ent.</div>
            </div>
            <select data-layer-cat="${key}">${opts(ls.category)}</select>
          </div>`;
        }).join("");
        return head + rows;
      }).join("");
      layerListEl.innerHTML = blocks;

      layerListEl.querySelectorAll("input[data-layer]").forEach((cb) => {
        cb.addEventListener("change", () => {
          state.layerState[decodeURIComponent(cb.dataset.layer)].visible = cb.checked;
          renderLayerList();
          render();
        });
      });
      layerListEl.querySelectorAll("input[data-cat]").forEach((cb) => {
        cb.addEventListener("change", () => {
          for (const l of state.layers) {
            if (state.layerState[l.name].category === cb.dataset.cat) {
              state.layerState[l.name].visible = cb.checked;
            }
          }
          renderLayerList();
          render();
        });
      });
      layerListEl.querySelectorAll("select[data-layer-cat]").forEach((sel) => {
        sel.addEventListener("change", () => {
          state.layerState[decodeURIComponent(sel.dataset.layerCat)].category = sel.value;
          renderLayerList();
          render();
        });
      });

      const left = state.layers.filter((l) => state.layerState[l.name].category === UNCLASSIFIED).length;
      unclassifiedNoteEl.style.display = left ? "block" : "none";
      unclassifiedNoteEl.textContent =
        `미분류 ${left}개 — 자동 판정은 C1 인식기(PR-4)가 붙은 뒤에 채워집니다.`;
    }

    // ====== 설계 산출물 토글 ======
    function renderDesignList() {
      designListEl.innerHTML = OVERLAYS.map(([key, label, color]) => {
        const n = key === "roomLabels" ? state.rooms.length : state[key].length;
        return `<label class="overlay-row ${n ? "" : "empty"}">
          <input type="checkbox" data-overlay="${key}" ${state.overlayVisible[key] ? "checked" : ""} ${n ? "" : "disabled"}>
          <i class="swatch" style="background:${color}"></i>
          <span>${label}</span>
          <span class="count">${n}</span>
        </label>`;
      }).join("");
      designListEl.querySelectorAll("input[data-overlay]").forEach((cb) => {
        cb.addEventListener("change", () => {
          state.overlayVisible[cb.dataset.overlay] = cb.checked;
          render();
        });
      });
    }

    renderStepper();
    renderDesignList();
    updateEditTools();
    startSession();
    setTimeout(resizeCanvas, 50);
  