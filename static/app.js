const form = document.getElementById("upload-form");
const reportFileInput = document.getElementById("report-file");
const sdfFileInput = document.getElementById("sdf-file");
const statusEl = document.getElementById("status");
const summaryEl = document.getElementById("summary");
const menuLayoutEl = document.getElementById("menu-layout");
const menuButtons = Array.from(document.querySelectorAll(".menu-btn"));
const menuPanels = Array.from(document.querySelectorAll(".menu-panel"));
const workspacePanelEl = document.getElementById("workspace-panel");
const resultsBody = document.getElementById("results-body");
const insightPanelEl = document.getElementById("insight-panel");
const engineeringListEl = document.getElementById("engineering-list");
const economyListEl = document.getElementById("economy-list");
const tablesPanelEl = document.getElementById("tables-panel");
const tableContainerEl = document.getElementById("table-container");
const statsPanelEl = document.getElementById("stats-panel");
const statsOutputEl = document.getElementById("stats-output");
const reportPanelEl = document.getElementById("report-panel");
const reportOutputEl = document.getElementById("report-output");
const tabButtons = Array.from(document.querySelectorAll(".tab-btn"));
const downloadExcelBtnEl = document.getElementById("download-csv-btn");

const networkSvgEl = document.getElementById("network-svg");
const networkViewportEl = document.getElementById("network-viewport");
const networkEmptyEl = document.getElementById("network-empty");
const zoomInBtnEl = document.getElementById("zoom-in-btn");
const zoomOutBtnEl = document.getElementById("zoom-out-btn");
const zoomResetBtnEl = document.getElementById("zoom-reset-btn");

const logicModalEl = document.getElementById("logic-modal");
const logicModalTitleEl = document.getElementById("logic-modal-title");
const logicModalBodyEl = document.getElementById("logic-modal-body");
const logicModalCloseEl = document.getElementById("logic-modal-close");

const criteriaBtnEl = document.getElementById("criteria-btn");
const criteriaModalEl = document.getElementById("criteria-modal");
const criteriaModalCloseEl = document.getElementById("criteria-modal-close");
const criteriaModalBodyEl = document.getElementById("criteria-modal-body");

let currentTables = null;
let currentTab = "pipes";
let currentMenuPanel = "workspace-panel";
let currentGraph = null;
let graphFilterMode = null; // null | FAIL | PASS
let activeResultButton = null;
let graphSelection = null; // {pipes:Set,nozzles:Set,equipment:Set,valves:Set,nodes:Set}
const MAX_MAIN_COLUMNS = 10;

const viewState = { base: null, zoom: 1, panX: 0, panY: 0, dragging: false, lastX: 0, lastY: 0 };

const statusLabel = { FAIL: "부적합", WARNING: "확인 필요", PASS: "적합" };
const statsLabelMap = {
  report_main_title: "결과서 메인 제목",
  report_zone_title: "결과서 존",
  pipe_count_from_report: "결과서 배관 수",
  nozzle_count_from_report: "결과서 헤드 수",
  equipment_count_from_report: "결과서 특수설비 수",
  valve_count_from_report: "결과서 감압밸브 수",
  min_nozzle_flow_lpm: "최소 헤드 유량 (L/min)",
  min_nozzle_pressure_kgf_cm2: "최소 헤드 압력 (kg/cm²G)",
  max_branch_pipe_velocity_mps: "가지배관 최대 유속 (m/s)",
  max_main_pipe_velocity_mps: "주배관 최대 유속 (m/s)",
  sdf_main_title: "SDF 메인 제목",
  sdf_zone_title: "SDF 존",
  pipe_count: "SDF 배관 수",
  nozzle_count: "SDF 헤드 수",
  equipment_count: "SDF 특수설비 수",
};

const tableConfigs = {
  pipes: {
    title: "배관(FLOW IN PIPES)",
    columns: [["label", "Pipe"], ["input_node", "입력 노드"], ["output_node", "출력 노드"], ["nominal_bore_mm", "구경(mm)"], ["flow_lpm", "유량(L/min)"], ["velocity_mps", "유속(m/s)"], ["inlet_pressure", "입구압"], ["outlet_pressure", "출구압"], ["friction_loss", "마찰손실"], ["special_equipment", "특수설비"]],
  },
  nozzles: {
    title: "헤드(FLOW THROUGH NOZZLES)",
    columns: [["label", "헤드"], ["input_node", "입력 노드"], ["inlet_pressure_kgf_cm2", "압력(kg/cm²G)"], ["required_flow_lpm", "요구 유량"], ["actual_flow_lpm", "실제 유량"], ["deviation_percent", "편차(%)"]],
  },
  equipment: {
    title: "특수설비(SPECIAL EQUIPMENT)",
    columns: [["label", "설비"], ["pipe_label", "배관"], ["description", "구분"], ["equivalent_length_m", "등가길이(m)"]],
  },
  valves: {
    title: "감압밸브(ELASTOMERIC VALVES)",
    columns: [["label", "밸브"], ["inlet_pressure_kgf_cm2", "입구압"], ["outlet_pressure_kgf_cm2", "출구압"], ["pressure_drop_kgf_cm2", "압력강하"], ["flow_lpm", "유량(L/min)"]],
  },
};

function fmt(v) {
  if (typeof v === "number") return Number.isInteger(v) ? String(v) : v.toFixed(3).replace(/\.?0+$/, "");
  if (typeof v === "boolean") return v ? "Y" : "N";
  return v ?? "";
}
function escapeHtml(s) {
  return String(s).replaceAll("&", "&amp;").replaceAll("<", "&lt;").replaceAll(">", "&gt;").replaceAll('"', "&quot;").replaceAll("'", "&#039;");
}
function rowClass(row) {
  if (row.highlight) return "row-fail";
  if (row.warn) return "row-warn";
  if (row.engineering_flag) return "row-eng";
  if (row.economy_flag) return "row-econ";
  return "";
}

function renderSummary(summary, filename, sdfFilename) {
  summaryEl.classList.remove("hidden");
  summaryEl.innerHTML = `<div class="summary-grid">
    <div><strong>결과서</strong><span>${escapeHtml(filename || "-")}</span></div>
    <div><strong>SDF</strong><span>${escapeHtml(sdfFilename || "미업로드")}</span></div>
    <div><strong>적합</strong><span>${fmt(summary.pass)}</span></div>
    <div><strong>부적합</strong><span>${fmt(summary.fail)}</span></div>
    <div><strong>확인 필요</strong><span>${fmt(summary.warning)}</span></div>
  </div>`;
}

function renderResultRows(results) {
  const rows = [];
  for (const key of ["FAIL", "WARNING", "PASS"]) {
    for (const message of results[key] || []) {
      const statusCell =
        key === "FAIL" || key === "PASS"
          ? `<button type="button" class="badge ${key.toLowerCase()} result-status-btn" data-status="${key}">${statusLabel[key]}</button>`
          : `<span class="badge ${key.toLowerCase()}">${statusLabel[key]}</span>`;
      rows.push(`<tr><td>${statusCell}</td><td>${escapeHtml(message)}</td></tr>`);
    }
  }
  resultsBody.innerHTML = rows.length ? rows.join("") : '<tr><td colspan="2" class="empty">검증 결과가 없습니다.</td></tr>';
  workspacePanelEl.classList.remove("hidden");
}

function renderInsights(insights) {
  const eng = insights?.engineering_advice || [];
  const eco = insights?.economy_guide || [];
  engineeringListEl.innerHTML = eng.length ? eng.map((x) => `<li>${escapeHtml(x)}</li>`).join("") : "<li>해당 사항 없음</li>";
  economyListEl.innerHTML = eco.length ? eco.map((x) => `<li>${escapeHtml(x)}</li>`).join("") : "<li>해당 사항 없음</li>";
  insightPanelEl.classList.remove("hidden");
}

function renderStats(stats) {
  const entries = Object.entries(stats || {}).filter(([, v]) => v !== null && v !== undefined && v !== "");
  statsOutputEl.innerHTML = entries.length
    ? entries.map(([k, v]) => `<article class="stat-card"><p class="stat-label">${escapeHtml(statsLabelMap[k] || k)}</p><p class="stat-value">${escapeHtml(fmt(v))}</p></article>`).join("")
    : '<p class="empty">표시할 통계가 없습니다.</p>';
  statsPanelEl.classList.remove("hidden");
}

function renderReport(data) {
  const s = data.summary || { pass: 0, fail: 0, warning: 0 };
  const r = data.results || {};
  const i = data.insights || {};
  const sections = [
    { title: "부적합 항목", cls: "fail", items: r.FAIL || [] },
    { title: "확인 필요 항목", cls: "warning", items: r.WARNING || [] },
    { title: "적합 항목", cls: "pass", items: r.PASS || [] },
    { title: "공학적 개선 제안", cls: "eng", items: i.engineering_advice || [] },
    { title: "경제성 확보 가이드", cls: "econ", items: i.economy_guide || [] },
  ];
  reportOutputEl.innerHTML = `<div class="report-summary-bar">
    <span class="chip pass">적합 ${fmt(s.pass)}</span>
    <span class="chip fail">부적합 ${fmt(s.fail)}</span>
    <span class="chip warning">확인 필요 ${fmt(s.warning)}</span>
  </div>${sections.map((x) => `<article class="report-card ${x.cls}"><h3>${x.title}</h3><ul>${(x.items.length ? x.items : ["해당 사항 없음"]).map((t) => `<li>${escapeHtml(t)}</li>`).join("")}</ul></article>`).join("")}`;
  reportPanelEl.classList.remove("hidden");
}

function buildTableHtml(columns, rows, startCol, wrapClass = "") {
  const head = columns.map(([, l]) => `<th>${l}</th>`).join("");
  const body = rows.length
    ? rows.map((row, idx) => `<tr class="${rowClass(row)}" data-row-index="${idx}">${columns.map(([k], j) => `<td data-col-index="${startCol + j}" data-key="${k}">${escapeHtml(fmt(row[k]))}</td>`).join("")}</tr>`).join("")
    : `<tr><td colspan="${columns.length}" class="empty">데이터가 없습니다.</td></tr>`;
  return `<div class="table-wrap ${wrapClass}"><table><thead><tr>${head}</tr></thead><tbody>${body}</tbody></table></div>`;
}

function renderSingleTable(tabName) {
  const cfg = tableConfigs[tabName];
  const rows = currentTables?.[tabName] || [];
  const mainCols = cfg.columns.slice(0, MAX_MAIN_COLUMNS);
  const extraCols = cfg.columns.slice(MAX_MAIN_COLUMNS);
  const extra = extraCols.length ? `<div class="extra-columns-panel"><h3>추가 열 (${extraCols.length}개)</h3>${buildTableHtml(extraCols, rows, MAX_MAIN_COLUMNS, "extra-wrap")}</div>` : "";
  tableContainerEl.innerHTML = `<div class="table-split">${buildTableHtml(mainCols, rows, 0, "main-wrap")}${extra}<p class="table-note">기본 창은 최대 ${MAX_MAIN_COLUMNS}열만 표시됩니다. 나머지는 추가 열 창에서 스크롤로 확인하세요.</p></div>`;
}
function setActiveTab(tab) {
  currentTab = tab;
  for (const b of tabButtons) b.classList.toggle("active", b.dataset.tab === tab);
  renderSingleTable(tab);
}

function setActiveMenuPanel(panelId) {
  currentMenuPanel = panelId;
  for (const b of menuButtons) b.classList.toggle("active", b.dataset.panel === panelId);
  for (const p of menuPanels) p.classList.toggle("hidden", p.id !== panelId);
}

function buildRowLogicExplanation(tab, row) {
  const lines = [];
  if (row.highlight) {
    if (tab === "pipes") {
      const bore = Number(row.nominal_bore_mm || 0);
      const limit = bore <= 50 ? 6 : 10;
      lines.push(`빨강(기준 위반): 유속 초과, 기준 ${limit} m/s`);
    } else if (tab === "nozzles") lines.push("빨강(기준 위반): 헤드 유량/압력 기준 미충족");
    else if (tab === "equipment") lines.push("빨강(기준 위반): 특수설비 등가길이 기준 미충족");
    else if (tab === "valves") lines.push("빨강(기준 위반): 감압밸브 압력강하 불일치");
  }
  if (row.warn) lines.push("노랑(확인 필요): 보완 검토가 필요합니다.");
  if (row.engineering_flag) lines.push("파랑(공학 후보): 마찰손실 급증 구간입니다.");
  if (row.economy_flag) lines.push("초록(경제성 후보): 과설계/원가 최적화 검토 대상입니다.");
  return lines;
}
function openLogicModal(title, lines) {
  logicModalTitleEl.textContent = title;
  logicModalBodyEl.innerHTML = lines.map((x) => `<li>${escapeHtml(x)}</li>`).join("");
  logicModalEl.classList.remove("hidden");
}

function criteriaGuideHtml() {
  return `
    <section class="criteria-section"><h4>1) 기본 검증 구조</h4><ul>
      <li><strong>입력 파일</strong>: 결과서(docx/pdf) + SDF(선택)</li>
      <li><strong>주요 섹션</strong>: FLOW IN PIPES, FLOW THROUGH NOZZLES, SPECIAL EQUIPMENT, ELASTOMERIC VALVES</li>
      <li><strong>판정 체계</strong>: PASS / FAIL / WARNING</li>
    </ul></section>
    <section class="criteria-section"><h4>2) 배관 유속 기준</h4><ul>
      <li><strong>50A 이하</strong>: 6.0 m/s 이하</li>
      <li><strong>50A 초과</strong>: 10.0 m/s 이하</li>
      <li>기준 초과 시 빨강(기준 위반)으로 표시됩니다.</li>
    </ul></section>
    <section class="criteria-section"><h4>3) 헤드(노즐) 기준</h4><ul>
      <li><strong>최소 유량</strong>: 80 L/min 이상</li>
      <li><strong>압력 범위</strong>: 1.0 ~ 12.0 kg/cm²G</li>
      <li>미충족 항목은 빨강(기준 위반)으로 표시됩니다.</li>
    </ul></section>
    <section class="criteria-section"><h4>4) 특수설비 등가길이 기준</h4><ul>
      <li><strong>FX</strong>: 13 ~ 21 m</li>
      <li><strong>A/V</strong>: 12.9 m (±0.1)</li>
      <li><strong>P/V</strong>: 10.1 m (±0.1)</li>
      <li>범위 이탈 시 빨강(기준 위반), 정보 부족 시 노랑(확인 필요) 처리됩니다.</li>
    </ul></section>
    <section class="criteria-section"><h4>5) 감압밸브 검증 논리</h4><ul>
      <li>계산식: <code>입구압 - 출구압 = 압력강하</code></li>
      <li>허용 오차: ±0.05</li>
      <li>오차 초과 시 빨강(기준 위반)으로 표시됩니다.</li>
    </ul></section>
    <section class="criteria-section"><h4>6) 공학 최적화(파랑) 기준</h4><ul>
      <li>m당 마찰손실 = <code>friction_loss / length</code></li>
      <li>기준: 0.05 kg/cm²/m 초과 시 공학 최적화 후보</li>
      <li>대응: 구경 상향, 피팅 축소, 배관 경로 단순화 검토</li>
    </ul></section>
    <section class="criteria-section"><h4>7) 경제성 검토(초록) 기준</h4><ul>
      <li><strong>배관</strong>: 유속 &lt; 2.0 m/s 이고 구경 &gt; 25A 인 경우 과설계 후보</li>
      <li><strong>밸브 계통</strong>: 연결 배관 구경 &gt; 100A 인 경우 원가 상승 후보</li>
      <li>대응: 법적/기술 조건 유지 범위 내 구경 최적화 검토</li>
    </ul></section>
    <section class="criteria-section"><h4>8) 색상 의미 요약</h4><ul>
      <li><strong>빨강</strong>: 기준 위반(즉시 수정 대상)</li>
      <li><strong>노랑</strong>: 확인 필요(보완 검토 대상)</li>
      <li><strong>파랑</strong>: 공학 최적화 후보(손실 개선 포인트)</li>
      <li><strong>초록</strong>: 경제성 검토 후보(원가 절감 포인트)</li>
    </ul></section>
    <section class="criteria-section"><h4>9) 모형 연동 해석</h4><ul>
      <li>좌측 <strong>부적합/적합</strong> 필터를 누르면 우측 배관망이 해당 상태만 강조됩니다.</li>
      <li>초기 상태는 모든 요소가 비활성(회색)이며, 필터 적용 시 상태별 색상이 활성화됩니다.</li>
    </ul></section>
  `;
}

function getGraphColor(type, label, status) {
  if (graphFilterMode === null) return "#cbd5e1";
  const active = graphFilterMode === "FAIL" ? "#dc2626" : "#16a34a";
  if (!graphSelection) {
    const target = graphFilterMode === "FAIL" ? "fail" : "pass";
    return status === target ? active : "#cbd5e1";
  }
  const set = graphSelection[type];
  return set && set.has(String(label)) ? active : "#cbd5e1";
}

function buildGraphSelectionFromMessage(status, message) {
  const selection = {
    pipes: new Set(),
    nozzles: new Set(),
    equipment: new Set(),
    valves: new Set(),
    nodes: new Set(),
  };
  if (!currentGraph) return selection;

  const targetStatus = status === "FAIL" ? "fail" : "pass";
  const nums = Array.from((message || "").matchAll(/\d+/g)).map((m) => Number(m[0]));
  const msg = String(message || "").toLowerCase();

  const pipesByStatus = (currentGraph.pipes || []).filter((x) => x.status === targetStatus);
  const nozByStatus = (currentGraph.nozzles || []).filter((x) => x.status === targetStatus);
  const eqByStatus = (currentGraph.equipment || []).filter((x) => (targetStatus === "fail" ? x.status === "fail" : x.status === "pass"));
  const valveByStatus = (currentGraph.valves || []).filter((x) => x.status === targetStatus);

  const addPipe = (p) => {
    selection.pipes.add(String(p.label));
    selection.nodes.add(String(p.input_node));
    selection.nodes.add(String(p.output_node));
  };
  const addNozzle = (n) => {
    selection.nozzles.add(String(n.label));
    selection.nodes.add(String(n.input_node));
  };
  const addEq = (e) => selection.equipment.add(String(e.label));
  const addValve = (v) => selection.valves.add(String(v.label));

  let matched = false;

  for (const n of nums) {
    const p = pipesByStatus.find((x) => x.label === n);
    if (p) {
      addPipe(p);
      matched = true;
    }
    const nz = nozByStatus.find((x) => x.label === n);
    if (nz) {
      addNozzle(nz);
      matched = true;
    }
    const e = eqByStatus.find((x) => x.label === n);
    if (e) {
      addEq(e);
      matched = true;
    }
    const v = valveByStatus.find((x) => x.label === n);
    if (v) {
      addValve(v);
      matched = true;
    }
  }

  if (msg.includes("pipe") || msg.includes("배관") || msg.includes("유속")) {
    pipesByStatus.forEach(addPipe);
    matched = true;
  }
  if (msg.includes("노즐") || msg.includes("헤드")) {
    nozByStatus.forEach(addNozzle);
    matched = true;
  }
  if (msg.includes("fx") || msg.includes("a/v") || msg.includes("p/v") || msg.includes("특수설비") || msg.includes("등가길이")) {
    let eqTargets = eqByStatus;
    if (msg.includes("fx")) eqTargets = eqTargets.filter((x) => String(x.description || "").toUpperCase() === "FX");
    if (msg.includes("a/v") || msg.includes("av")) eqTargets = eqTargets.filter((x) => ["A/V", "AV"].includes(String(x.description || "").toUpperCase()));
    if (msg.includes("p/v") || msg.includes("pv")) eqTargets = eqTargets.filter((x) => ["P/V", "PV"].includes(String(x.description || "").toUpperCase()));
    eqTargets.forEach(addEq);
    matched = true;
  }
  if (msg.includes("감압밸브") || msg.includes("밸브")) {
    valveByStatus.forEach(addValve);
    matched = true;
  }

  if (!matched) {
    pipesByStatus.forEach(addPipe);
    nozByStatus.forEach(addNozzle);
    eqByStatus.forEach(addEq);
    valveByStatus.forEach(addValve);
  }

  return selection;
}
function makeSvgEl(tag, attrs = {}) {
  const el = document.createElementNS("http://www.w3.org/2000/svg", tag);
  for (const [k, v] of Object.entries(attrs)) el.setAttribute(k, String(v));
  return el;
}
function updateViewBox() {
  if (!viewState.base) return;
  const b = viewState.base;
  const w = b.width / viewState.zoom;
  const h = b.height / viewState.zoom;
  const cx = b.minX + b.width / 2 + viewState.panX;
  const cy = b.minY + b.height / 2 + viewState.panY;
  networkSvgEl.setAttribute("viewBox", `${cx - w / 2} ${cy - h / 2} ${w} ${h}`);
}
function resetGraphView(bounds) {
  viewState.base = bounds;
  viewState.zoom = 1;
  viewState.panX = 0;
  viewState.panY = 0;
  updateViewBox();
}
function renderNetworkGraph(graph) {
  currentGraph = graph || null;
  while (networkViewportEl.firstChild) networkViewportEl.removeChild(networkViewportEl.firstChild);
  if (!graph || !(graph.pipes || []).length) {
    networkEmptyEl.classList.remove("hidden");
    networkSvgEl.classList.add("hidden");
    return;
  }
  networkEmptyEl.classList.add("hidden");
  networkSvgEl.classList.remove("hidden");
  const fy = (y) => -y; // x축 기준 대칭(상하 반전)

  const pts = [];
  for (const p of graph.pipes || []) for (const q of p.path || []) pts.push([q[0], fy(q[1])]);
  for (const n of graph.nozzles || []) pts.push([n.x, fy(n.y)]);
  for (const e of graph.equipment || []) pts.push([e.x, fy(e.y)]);
  for (const v of graph.valves || []) pts.push([v.x, fy(v.y)]);
  const xs = pts.map((p) => p[0]);
  const ys = pts.map((p) => p[1]);
  const pad = 200;
  resetGraphView({ minX: Math.min(...xs) - pad, minY: Math.min(...ys) - pad, width: Math.max(...xs) - Math.min(...xs) + pad * 2, height: Math.max(...ys) - Math.min(...ys) + pad * 2 });

  for (const p of graph.pipes || []) {
    networkViewportEl.appendChild(
      makeSvgEl("polyline", {
        points: (p.path || []).map((x) => `${x[0]},${fy(x[1])}`).join(" "),
        fill: "none",
        stroke: getGraphColor("pipes", p.label, p.status),
        "stroke-width": 20,
        "stroke-linecap": "round",
        "stroke-linejoin": "round",
      })
    );
  }
  for (const n of graph.nozzles || []) {
    networkViewportEl.appendChild(makeSvgEl("circle", { cx: n.x, cy: fy(n.y), r: 16, fill: getGraphColor("nozzles", n.label, n.status), stroke: "#111827", "stroke-width": 4 }));
  }
  for (const e of graph.equipment || []) {
    networkViewportEl.appendChild(makeSvgEl("rect", { x: e.x - 18, y: fy(e.y) - 18, width: 36, height: 36, fill: getGraphColor("equipment", e.label, e.status), stroke: "#111827", "stroke-width": 4 }));
  }
  for (const v of graph.valves || []) {
    const vy = fy(v.y);
    networkViewportEl.appendChild(makeSvgEl("polygon", { points: `${v.x},${vy - 24} ${v.x + 24},${vy} ${v.x},${vy + 24} ${v.x - 24},${vy}`, fill: getGraphColor("valves", v.label, v.status), stroke: "#111827", "stroke-width": 4 }));
  }

  // 선택된 노드 강조(요청: 노드/엣지 표시)
  if (graphFilterMode !== null) {
    for (const node of graph.nodes || []) {
      const on = graphSelection?.nodes?.has(String(node.id));
      networkViewportEl.appendChild(
        makeSvgEl("circle", {
          cx: node.x,
          cy: fy(node.y),
          r: 8,
          fill: on ? (graphFilterMode === "FAIL" ? "#dc2626" : "#16a34a") : "#cbd5e1",
          stroke: "#334155",
          "stroke-width": 2,
        })
      );
    }
  }
}

function applyGraphFilter(mode) {
  graphFilterMode = mode;
  graphSelection = null;
  renderNetworkGraph(currentGraph);
}

async function downloadTablesAsXlsx() {
  if (!currentTables) return;
  try {
    downloadExcelBtnEl.disabled = true;
    const resp = await fetch("/api/export-xlsx", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ tables: currentTables, report_name: reportFileInput.files[0]?.name || "pipenet_result" }) });
    if (!resp.ok) throw new Error("엑셀 다운로드에 실패했습니다.");
    const blob = await resp.blob();
    const url = URL.createObjectURL(blob);
    const cd = resp.headers.get("Content-Disposition") || "";
    const m = /filename\*=UTF-8''([^;]+)|filename=\"?([^\";]+)\"?/i.exec(cd);
    const filename = decodeURIComponent((m && (m[1] || m[2])) || "pipenet_result.xlsx");
    const a = document.createElement("a");
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
  } catch (e) {
    alert(e.message || "엑셀 다운로드에 실패했습니다.");
  } finally {
    downloadExcelBtnEl.disabled = false;
  }
}

resultsBody.addEventListener("click", (event) => {
  const btn = event.target.closest(".result-status-btn");
  if (!btn) return;
  const status = btn.dataset.status;
  if (status !== "FAIL" && status !== "PASS") return;

  // 동일 버튼 재클릭: 해제
  if (activeResultButton === btn) {
    btn.classList.remove("active-filter");
    activeResultButton = null;
    applyGraphFilter(null);
    return;
  }

  // 다른 버튼 클릭: 기존 1개 해제 후 새 버튼 1개만 활성화
  if (activeResultButton) activeResultButton.classList.remove("active-filter");
  btn.classList.add("active-filter");
  activeResultButton = btn;
  graphFilterMode = status;
  graphSelection = buildGraphSelectionFromMessage(status, btn.closest("tr")?.children?.[1]?.innerText || "");
  renderNetworkGraph(currentGraph);
});

tableContainerEl.addEventListener("mouseup", () => {
  const sel = window.getSelection();
  if (!sel || sel.rangeCount === 0 || sel.isCollapsed) return;
  const r = sel.getRangeAt(0);
  const s = r.startContainer?.parentElement?.closest?.("td");
  const e = r.endContainer?.parentElement?.closest?.("td");
  if (!s || !e) return;
  const si = Number(s.parentElement?.dataset?.rowIndex ?? -1);
  const ei = Number(e.parentElement?.dataset?.rowIndex ?? -1);
  if (si < 0 || ei < 0) return;
  const picked = (currentTables?.[currentTab] || []).slice(Math.min(si, ei), Math.max(si, ei) + 1).find((x) => x.highlight || x.warn || x.engineering_flag || x.economy_flag);
  if (!picked) return;
  openLogicModal(`[${tableConfigs[currentTab]?.columns?.[Number(s.dataset.colIndex || 0)]?.[1] || "선택 항목"}] 필터 판정 로직`, buildRowLogicExplanation(currentTab, picked));
});

tableContainerEl.addEventListener("click", (event) => {
  const c = event.target.closest("td");
  if (!c) return;
  const i = Number(c.parentElement?.dataset?.rowIndex ?? -1);
  if (i < 0) return;
  const row = (currentTables?.[currentTab] || [])[i];
  if (!row || !(row.highlight || row.warn || row.engineering_flag || row.economy_flag)) return;
  openLogicModal(`[${tableConfigs[currentTab]?.columns?.[Number(c.dataset.colIndex || 0)]?.[1] || "선택 항목"}] 필터 판정 로직`, buildRowLogicExplanation(currentTab, row));
});

for (const b of tabButtons) b.addEventListener("click", () => setActiveTab(b.dataset.tab));
for (const b of menuButtons) b.addEventListener("click", () => setActiveMenuPanel(b.dataset.panel));
if (menuButtons.length) setActiveMenuPanel(currentMenuPanel);

networkSvgEl.addEventListener("wheel", (event) => {
  if (!viewState.base) return;
  event.preventDefault();
  viewState.zoom = Math.max(0.2, Math.min(20, viewState.zoom * (event.deltaY < 0 ? 1.15 : 0.87)));
  updateViewBox();
});
networkSvgEl.addEventListener("mousedown", (event) => {
  viewState.dragging = true;
  viewState.lastX = event.clientX;
  viewState.lastY = event.clientY;
});
window.addEventListener("mouseup", () => (viewState.dragging = false));
window.addEventListener("mousemove", (event) => {
  if (!viewState.dragging || !viewState.base) return;
  const dx = event.clientX - viewState.lastX;
  const dy = event.clientY - viewState.lastY;
  viewState.lastX = event.clientX;
  viewState.lastY = event.clientY;
  const w = viewState.base.width / viewState.zoom;
  const h = viewState.base.height / viewState.zoom;
  viewState.panX -= dx * (w / Math.max(1, networkSvgEl.clientWidth));
  viewState.panY -= dy * (h / Math.max(1, networkSvgEl.clientHeight));
  updateViewBox();
});
zoomInBtnEl.addEventListener("click", () => {
  viewState.zoom = Math.min(20, viewState.zoom * 1.2);
  updateViewBox();
});
zoomOutBtnEl.addEventListener("click", () => {
  viewState.zoom = Math.max(0.2, viewState.zoom * 0.83);
  updateViewBox();
});
zoomResetBtnEl.addEventListener("click", () => {
  if (viewState.base) resetGraphView(viewState.base);
});

logicModalCloseEl.addEventListener("click", () => logicModalEl.classList.add("hidden"));
logicModalEl.addEventListener("click", (e) => {
  if (e.target === logicModalEl) logicModalEl.classList.add("hidden");
});
criteriaBtnEl.addEventListener("click", () => {
  criteriaModalBodyEl.innerHTML = criteriaGuideHtml();
  criteriaModalEl.classList.remove("hidden");
});
criteriaModalCloseEl.addEventListener("click", () => criteriaModalEl.classList.add("hidden"));
criteriaModalEl.addEventListener("click", (e) => {
  if (e.target === criteriaModalEl) criteriaModalEl.classList.add("hidden");
});
downloadExcelBtnEl.addEventListener("click", downloadTablesAsXlsx);
window.addEventListener("keydown", (e) => {
  if (e.key === "Escape") {
    logicModalEl.classList.add("hidden");
    criteriaModalEl.classList.add("hidden");
  }
});

form.addEventListener("submit", async (event) => {
  event.preventDefault();
  const reportFile = reportFileInput.files[0];
  if (!reportFile) {
    statusEl.textContent = "결과서 파일을 먼저 선택해 주세요.";
    return;
  }

  statusEl.textContent = "검증 중입니다...";
  downloadExcelBtnEl.disabled = true;
  graphFilterMode = null;
  graphSelection = null;
  if (activeResultButton) {
    activeResultButton.classList.remove("active-filter");
    activeResultButton = null;
  }
  summaryEl.classList.add("hidden");
  if (menuLayoutEl) menuLayoutEl.classList.add("hidden");
  workspacePanelEl.classList.add("hidden");
  insightPanelEl.classList.add("hidden");
  tablesPanelEl.classList.add("hidden");
  statsPanelEl.classList.add("hidden");
  reportPanelEl.classList.add("hidden");

  const formData = new FormData();
  formData.append("report_file", reportFile);
  if (sdfFileInput.files[0]) formData.append("sdf_file", sdfFileInput.files[0]);

  try {
    const resp = await fetch("/api/validate", { method: "POST", body: formData });
    const data = await resp.json();
    if (!resp.ok || !data.ok) throw new Error(data.message || "검증 요청에 실패했습니다.");

    statusEl.textContent = data.message;
    renderSummary(data.summary || {}, data.filename, data.sdf_filename);
    renderResultRows(data.results || {});
    renderNetworkGraph(data.sdf_graph || null);
    renderInsights(data.insights || {});
    currentTables = data.tables || {};
    setActiveTab(currentTab);
    renderStats(data.stats || {});
    renderReport(data);
    if (menuLayoutEl) menuLayoutEl.classList.remove("hidden");
    setActiveMenuPanel(currentMenuPanel || "workspace-panel");
    downloadExcelBtnEl.disabled = false;
  } catch (e) {
    statusEl.textContent = e.message;
    resultsBody.innerHTML = '<tr><td colspan="2" class="empty">결과를 불러오지 못했습니다.</td></tr>';
    if (menuLayoutEl) menuLayoutEl.classList.remove("hidden");
    setActiveMenuPanel("workspace-panel");
    networkEmptyEl.textContent = "배관망을 불러오지 못했습니다.";
    networkEmptyEl.classList.remove("hidden");
    networkSvgEl.classList.add("hidden");
  }
});
