const form = document.getElementById("upload-form");
const appShellEl = document.querySelector(".app-shell");
const reportFileInput = document.getElementById("report-file");
const sdfFileInput = document.getElementById("sdf-file");
const reportFileLabelEl = document.querySelector('label[for="report-file"]');
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
const statsChartLeftEl = document.getElementById("stats-chart-left");
const statsChartRightEl = document.getElementById("stats-chart-right");
const statsOutputLeftEl = document.getElementById("stats-output-left");
const statsOutputRightEl = document.getElementById("stats-output-right");
const reportPanelEl = document.getElementById("report-panel");
const reportOutputEl = document.getElementById("report-output");
const cadComparePanelEl = document.getElementById("cad-compare-panel");
const cadRefFileEl = document.getElementById("cad-ref-file");
const cadIsoFileEl = document.getElementById("cad-iso-file");
const cadUseYoloEl = document.getElementById("cad-use-yolo");
const cadRunBtnEl = document.getElementById("cad-run-btn");
const cadStatusEl = document.getElementById("cad-status");
const cadSummaryEl = document.getElementById("cad-summary");
const cadMessagesEl = document.getElementById("cad-messages");
const cadSvgEl = document.getElementById("cad-svg");
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

const updatesBtnEl = document.getElementById("updates-btn");
const criteriaBtnEl = document.getElementById("criteria-btn");
const criteriaModalEl = document.getElementById("criteria-modal");
const criteriaModalCloseEl = document.getElementById("criteria-modal-close");
const criteriaModalBodyEl = document.getElementById("criteria-modal-body");
const updatesModalEl = document.getElementById("updates-modal");
const updatesModalCloseEl = document.getElementById("updates-modal-close");
const updatesModalBodyEl = document.getElementById("updates-modal-body");
const updatesModalTitleEl = document.getElementById("updates-modal-title");

let currentTables = null;
let currentTab = "pipes";
let currentMenuPanel = null;
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
    columns: [["label", "Pipe"], ["input_node", "입력 노드"], ["output_node", "출력 노드"], ["nominal_bore_mm", "구경(mm)"], ["actual_bore_mm", "실내경(mm)"], ["flow_lpm", "유량(L/min)"], ["velocity_mps", "유속(m/s)"], ["c_factor", "C-Factor"], ["base_length_m", "배관길이(m)"], ["fitting_eq_length_m", "피팅 등가길이(m)"], ["special_eq_length_m", "특수설비 등가길이(m)"], ["total_length_m", "총 등가길이(m)"], ["friction_loss", "결과서 마찰손실"], ["hw_expected_friction_loss", "HW 재계산 마찰손실"], ["hw_abs_diff", "HW 절대오차"], ["hw_rel_diff", "HW 상대오차"], ["hw_formula_ok", "HW 검산 적합"], ["inlet_pressure", "입구압"], ["outlet_pressure", "출구압"], ["special_equipment", "특수설비"], ["hw_fail", "HW 실패"]],
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
function fmtLogicNumber(v, digits = 3) {
  const n = Number(v);
  if (!Number.isFinite(n)) return "-";
  return n.toFixed(digits).replace(/\.?0+$/, "");
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
}

function renderInsights(insights) {
  const eng = insights?.engineering_advice || [];
  const eco = insights?.economy_guide || [];
  engineeringListEl.innerHTML = eng.length ? eng.map((x) => `<li>${escapeHtml(x)}</li>`).join("") : "<li>해당 사항 없음</li>";
  economyListEl.innerHTML = eco.length ? eco.map((x) => `<li>${escapeHtml(x)}</li>`).join("") : "<li>해당 사항 없음</li>";
}

function renderStats(stats, visualizations = []) {
  const entries = Object.entries(stats || {}).filter(([, v]) => v !== null && v !== undefined && v !== "");
  const isSdfKey = (k) => k.startsWith("sdf_") || ["pipe_count", "nozzle_count", "equipment_count"].includes(k);
  const leftEntries = entries.filter(([k]) => !isSdfKey(k)); // report(docx/pdf)
  const rightEntries = entries.filter(([k]) => isSdfKey(k)); // sdf

  const tableHtml = (list, title) =>
    `<article class="stats-table-card">
      <h3>${escapeHtml(title)}</h3>
      ${
        list.length
          ? `<table class="stats-table"><tbody>${list
              .map(
                ([k, v]) =>
                  `<tr><th>${escapeHtml(statsLabelMap[k] || k)}</th><td>${escapeHtml(fmt(v))}</td></tr>`
              )
              .join("")}</tbody></table>`
          : '<p class="empty">표시할 통계가 없습니다.</p>'
      }
    </article>`;

  if (statsOutputLeftEl) statsOutputLeftEl.innerHTML = tableHtml(leftEntries, "결과서 파일 (DOCX/PDF) 통계");
  if (statsOutputRightEl) statsOutputRightEl.innerHTML = tableHtml(rightEntries, "SDF 파일 통계");

  const charts = Array.isArray(visualizations) ? visualizations : [];
  const pipeChart =
    charts.find((c) => String(c?.title || "").toLowerCase().includes("pipe velocity")) || charts[0] || null;
  const nozzleChart =
    charts.find((c) => String(c?.title || "").toLowerCase().includes("nozzle pressure-flow")) ||
    charts.find((c) => c && c !== pipeChart) ||
    null;

  const chartHtml = (chart, fallbackTitle) =>
    chart
      ? `<article class="chart-card">
          <h3>${escapeHtml(chart.title || fallbackTitle)}</h3>
          <p>${escapeHtml(chart.description || "")}</p>
          <img src="${escapeHtml(chart.image_data_url || "")}" alt="${escapeHtml(chart.title || fallbackTitle)}">
        </article>`
      : `<article class="chart-card"><h3>${escapeHtml(fallbackTitle)}</h3><p class="empty">표시할 그래프가 없습니다.</p></article>`;

  if (statsChartLeftEl) statsChartLeftEl.innerHTML = chartHtml(pipeChart, "Pipe Velocity Check");
  if (statsChartRightEl) statsChartRightEl.innerHTML = chartHtml(nozzleChart, "Nozzle Pressure-Flow");
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
}

function fmtInt(v) {
  return Number.isFinite(Number(v)) ? String(Number(v)) : "0";
}

function renderCadSummary(data) {
  const cc = data?.cad_counts || {};
  const sc = data?.sdf_counts || {};
  cadSummaryEl.innerHTML = `
    <div class="cad-summary-grid">
      <div><strong>CAD 파일</strong><br>${escapeHtml(data?.cad_filename || "-")}</div>
      <div><strong>SDF 파일</strong><br>${escapeHtml(data?.sdf_filename || "미업로드")}</div>
      <div><strong>CAD 네트워크 엔티티</strong><br>${fmtInt(cc.entities)}</div>
      <div><strong>CAD 네트워크 레이어 수</strong><br>${fmtInt(cc.network_layers)}</div>
      <div><strong>CAD 탐지 헤드 수</strong><br>${fmtInt(cc.detected_heads)}</div>
      <div><strong>SDF 헤드 수</strong><br>${fmtInt(sc.nozzles)}</div>
    </div>`;
}

function renderCadMessages(messages) {
  const lines = Array.isArray(messages) ? messages : [];
  cadMessagesEl.innerHTML = lines.length ? lines.map((m) => `<li>${escapeHtml(m)}</li>`).join("") : "<li>대조 메시지가 없습니다.</li>";
}

function renderCadSvg(payload, headBoxes) {
  if (!cadSvgEl) return;
  const bounds = payload?.bounds;
  const entities = payload?.entities || [];
  if (!bounds || !entities.length) {
    cadSvgEl.innerHTML = "";
    cadSvgEl.setAttribute("viewBox", "0 0 100 100");
    return;
  }

  const minX = Number(bounds.minX ?? 0);
  const minY = Number(bounds.minY ?? 0);
  const maxX = Number(bounds.maxX ?? 100);
  const maxY = Number(bounds.maxY ?? 100);
  const width = Math.max(1, maxX - minX);
  const height = Math.max(1, maxY - minY);
  cadSvgEl.setAttribute("viewBox", `${minX} ${-maxY} ${width} ${height}`);
  cadSvgEl.innerHTML = "";

  const svgNS = "http://www.w3.org/2000/svg";
  const make = (tag, attrs) => {
    const el = document.createElementNS(svgNS, tag);
    for (const [k, v] of Object.entries(attrs)) el.setAttribute(k, String(v));
    return el;
  };

  for (const e of entities) {
    const color = e.color || "#64748b";
    if (e.type === "LINE") {
      cadSvgEl.appendChild(
        make("line", {
          x1: e.start.x,
          y1: -e.start.y,
          x2: e.end.x,
          y2: -e.end.y,
          stroke: color,
          "stroke-width": 10,
          "stroke-linecap": "round",
        })
      );
    } else if (e.type === "LWPOLYLINE") {
      const pts = (e.points || []).map((p) => `${p.x},${-p.y}`).join(" ");
      cadSvgEl.appendChild(
        make("polyline", {
          points: pts,
          fill: e.closed ? "rgba(100,116,139,0.07)" : "none",
          stroke: color,
          "stroke-width": 8,
        })
      );
    } else if (e.type === "CIRCLE") {
      cadSvgEl.appendChild(
        make("circle", {
          cx: e.center.x,
          cy: -e.center.y,
          r: Math.max(2, e.radius),
          fill: "none",
          stroke: color,
          "stroke-width": 8,
        })
      );
    } else if (e.type === "ARC") {
      const a0 = ((Number(e.startAngle || 0) * Math.PI) / 180.0);
      const a1 = ((Number(e.endAngle || 0) * Math.PI) / 180.0);
      const x1 = e.center.x + e.radius * Math.cos(a0);
      const y1 = e.center.y + e.radius * Math.sin(a0);
      const x2 = e.center.x + e.radius * Math.cos(a1);
      const y2 = e.center.y + e.radius * Math.sin(a1);
      const laf = Math.abs(a1 - a0) > Math.PI ? 1 : 0;
      const d = `M ${x1} ${-y1} A ${e.radius} ${e.radius} 0 ${laf} 0 ${x2} ${-y2}`;
      cadSvgEl.appendChild(make("path", { d, fill: "none", stroke: color, "stroke-width": 8 }));
    }
  }

  const heads = Array.isArray(headBoxes) ? headBoxes : [];
  for (const b of heads) {
    const x = Number(b.minX || 0);
    const y = -Number(b.maxY || 0);
    const w = Math.max(1, Number(b.maxX || 0) - Number(b.minX || 0));
    const h = Math.max(1, Number(b.maxY || 0) - Number(b.minY || 0));
    cadSvgEl.appendChild(
      make("rect", {
        x,
        y,
        width: w,
        height: h,
        fill: "none",
        stroke: "#dc2626",
        "stroke-width": 18,
      })
    );
  }
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
  const extraTitle = tabName === "pipes" ? `HW 검산 상세 열 (${extraCols.length}개)` : `추가 열 (${extraCols.length}개)`;
  const extra = extraCols.length ? `<div class="extra-columns-panel"><h3>${extraTitle}</h3>${buildTableHtml(extraCols, rows, MAX_MAIN_COLUMNS, "extra-wrap")}</div>` : "";
  tableContainerEl.innerHTML = `<div class="table-split">${buildTableHtml(mainCols, rows, 0, "main-wrap")}${extra}<p class="table-note">기본 창은 최대 ${MAX_MAIN_COLUMNS}열만 표시됩니다. 나머지는 추가 열 창에서 스크롤로 확인하세요.</p></div>`;
}
function setActiveTab(tab) {
  currentTab = tab;
  for (const b of tabButtons) b.classList.toggle("active", b.dataset.tab === tab);
  renderSingleTable(tab);
}

function setActiveMenuPanel(panelId) {
  currentMenuPanel = panelId;
  for (const b of menuButtons) b.classList.toggle("active", !!panelId && b.dataset.panel === panelId);
  for (const p of menuPanels) p.classList.toggle("hidden", p.id !== panelId);
}

function setMenuLayoutVisible(visible) {
  if (menuLayoutEl) menuLayoutEl.classList.toggle("hidden", !visible);
  if (appShellEl) appShellEl.classList.toggle("menu-open", visible);
}

function buildRowLogicExplanation(tab, row) {
  const cards = {
    criteria: [],
    formula: [],
    values: [],
    conclusion: [],
  };
  if (tab === "pipes") {
    const bore = Number(row.nominal_bore_mm || 0);
    const velocity = Number(row.velocity_mps || 0);
    const limit = Number(row.velocity_limit_mps ?? (bore <= 50 ? 6 : 10));
    const pipeType = row.pipe_type === "branch" || bore <= 50 ? "가지배관" : "주배관";
    const over = velocity - limit;
    cards.criteria.push(`배관 분류: 구경 ${fmtLogicNumber(bore, 0)}A 이므로 ${bore <= 50 ? "50A 이하" : "50A 초과"} 기준을 적용하여 ${pipeType}으로 판정했습니다.`);
    cards.criteria.push(`허용 유속 기준: ${bore <= 50 ? "6.0" : "10.0"} m/s`);
    cards.formula.push(`허용 유속식: V_limit = ${bore <= 50 ? "6.0" : "10.0"} m/s`);
    cards.formula.push(`판정식: V_actual ${row.highlight ? ">" : "<="} V_limit`);
    cards.values.push(`실제 유속 V_actual = ${fmtLogicNumber(velocity)} m/s`);
    cards.values.push(`허용 유속 V_limit = ${fmtLogicNumber(limit)} m/s`);
    if (row.highlight) {
      cards.conclusion.push(`빨강(기준 위반): ${fmtLogicNumber(velocity)} - ${fmtLogicNumber(limit)} = ${fmtLogicNumber(over)} m/s 초과하여 유속 기준을 만족하지 못했습니다.`);
    } else {
      cards.conclusion.push(`적합 해석: ${fmtLogicNumber(limit)} - ${fmtLogicNumber(velocity)} = ${fmtLogicNumber(limit - velocity)} m/s 여유가 있어 유속 기준을 만족합니다.`);
    }
    if (row.engineering_flag) {
      const frictionLoss = Number(row.friction_loss || 0);
      const length = Number(row.pipe_length_m || 0);
      const unitLoss = length > 0 ? frictionLoss / length : null;
      const spike = Number(row.friction_spike_limit ?? 0.05);
      cards.criteria.push(`공학 최적화 기준: 단위 마찰손실 > ${fmtLogicNumber(spike, 3)} kg/cm²/m`);
      cards.formula.push(`단위 마찰손실 = friction_loss / length`);
      cards.values.push(`단위 마찰손실 = ${fmtLogicNumber(frictionLoss)} / ${fmtLogicNumber(length)} = ${fmtLogicNumber(unitLoss)} kg/cm²/m`);
      cards.conclusion.push(`파랑(공학 최적화 후보): ${fmtLogicNumber(unitLoss)} ${unitLoss > spike ? ">" : "<="} ${fmtLogicNumber(spike, 3)} 이므로 구경 상향, 피팅 축소, 배관 경로 단순화 검토가 필요합니다.`);
    }
    if (row.economy_flag) {
      const econLimit = Number(row.economy_velocity_limit ?? 2.0);
      cards.criteria.push(`경제성 기준: 유속 < ${fmtLogicNumber(econLimit, 1)} m/s AND 구경 > 25A`);
      cards.formula.push(`경제성 판정식: velocity < ${fmtLogicNumber(econLimit, 1)} AND bore > 25`);
      cards.values.push(`실제 값 대입: ${fmtLogicNumber(velocity)} < ${fmtLogicNumber(econLimit, 1)}, ${fmtLogicNumber(bore, 0)}A > 25A`);
      cards.conclusion.push(`초록(경제성 검토 후보): 법적 최소 성능을 유지한다는 전제에서 현재 구간은 저유속 과설계 후보로 분류되며, 자재비 절감을 위한 구경 축소 검토 대상입니다.`);
    }
  } else if (tab === "nozzles") {
    const pressure = Number(row.inlet_pressure_kgf_cm2 || 0);
    const actualFlow = Number(row.actual_flow_lpm || 0);
    const requiredFlow = Number(row.required_flow_lpm || 0);
    const minFlow = Number(row.min_flow_limit_lpm ?? 80);
    const minP = Number(row.min_pressure_limit_kgf_cm2 ?? 1);
    const maxP = Number(row.max_pressure_limit_kgf_cm2 ?? 12);
    cards.criteria.push("헤드 검증은 실제 유량(actual_flow_lpm)과 입구압(inlet_pressure_kgf_cm2)을 기준으로 수행합니다.");
    cards.criteria.push(`유량 기준: ${fmtLogicNumber(minFlow, 1)} L/min 이상`);
    cards.criteria.push(`압력 기준: ${fmtLogicNumber(minP, 1)} ~ ${fmtLogicNumber(maxP, 1)} kg/cm²G`);
    cards.formula.push(`유량 판정식: actual_flow_lpm >= ${fmtLogicNumber(minFlow, 1)}`);
    cards.formula.push(`압력 판정식: ${fmtLogicNumber(minP, 1)} <= inlet_pressure <= ${fmtLogicNumber(maxP, 1)}`);
    cards.values.push(`실제 유량 = ${fmtLogicNumber(actualFlow)} L/min`);
    cards.values.push(`실제 압력 = ${fmtLogicNumber(pressure)} kg/cm²G`);
    cards.values.push(`결과서 요구 유량 = ${fmtLogicNumber(requiredFlow)} L/min`);
    if (row.highlight) {
      if (actualFlow < minFlow) {
        cards.conclusion.push(`빨강(기준 위반) 사유 1: ${fmtLogicNumber(actualFlow)} < ${fmtLogicNumber(minFlow, 1)} 이므로 최소 헤드 유량 기준에 미달합니다.`);
      }
      if (pressure < minP || pressure > maxP) {
        cards.conclusion.push(`빨강(기준 위반) 사유 2: 압력 ${fmtLogicNumber(pressure)} kg/cm²G가 허용 범위 ${fmtLogicNumber(minP, 1)}~${fmtLogicNumber(maxP, 1)} kg/cm²G를 벗어났습니다.`);
      }
    } else {
      cards.conclusion.push("판정 결과: 유량과 압력이 모두 기준 범위 안에 있으므로 해당 헤드는 적합으로 판단됩니다.");
    }
    cards.conclusion.push("참고: 현재 검증 로직은 결과서 요구 유량값보다 법정 최소유량 80 L/min 충족 여부를 우선 기준으로 사용합니다.");
  } else if (tab === "equipment") {
    const desc = String(row.description || "").toUpperCase();
    const eq = Number(row.equivalent_length_m || 0);
    const tol = Number(row.eq_tolerance_m ?? 0.1);
    if (desc === "FX") {
      const min = Number(row.fx_eq_min_m ?? 13);
      const max = Number(row.fx_eq_max_m ?? 21);
      cards.criteria.push(`FX 등가길이 기준: ${fmtLogicNumber(min, 1)} ~ ${fmtLogicNumber(max, 1)} m`);
      cards.formula.push(`판정식: ${fmtLogicNumber(min, 1)} <= equivalent_length <= ${fmtLogicNumber(max, 1)}`);
      cards.values.push(`실제 값 대입: ${fmtLogicNumber(min, 1)} <= ${fmtLogicNumber(eq)} <= ${fmtLogicNumber(max, 1)}`);
      if (row.highlight) {
        cards.conclusion.push(`빨강(기준 위반): FX 등가길이 ${fmtLogicNumber(eq)} m가 허용 범위 ${fmtLogicNumber(min, 1)}~${fmtLogicNumber(max, 1)} m를 벗어나 마찰손실이 과소 또는 과대 평가될 수 있습니다.`);
      } else {
        cards.conclusion.push("판정 결과: FX 등가길이가 권장 범위 안에 있어 적합합니다.");
      }
    } else if (desc === "A/V" || desc === "AV") {
      const ref = Number(row.av_eq_ref_m ?? 12.9);
      const diff = Math.abs(eq - ref);
      cards.criteria.push(`A/V 기준값: ${fmtLogicNumber(ref, 1)} m, 허용 오차 ±${fmtLogicNumber(tol, 1)} m`);
      cards.formula.push(`판정식: |equivalent_length - ${fmtLogicNumber(ref, 1)}| <= ${fmtLogicNumber(tol, 1)}`);
      cards.values.push(`실제 계산: |${fmtLogicNumber(eq)} - ${fmtLogicNumber(ref, 1)}| = ${fmtLogicNumber(diff)}`);
      if (row.highlight) {
        cards.conclusion.push(`빨강(기준 위반): 편차 ${fmtLogicNumber(diff)} m가 허용 오차 ${fmtLogicNumber(tol, 1)} m를 초과하여 A/V 등가길이 기준에 맞지 않습니다.`);
      } else {
        cards.conclusion.push(`판정 결과: 편차 ${fmtLogicNumber(diff)} m로 허용 오차 이내이므로 적합합니다.`);
      }
    } else if (desc === "P/V" || desc === "PV") {
      const ref = Number(row.pv_eq_ref_m ?? 10.1);
      const diff = Math.abs(eq - ref);
      cards.criteria.push(`P/V 기준값: ${fmtLogicNumber(ref, 1)} m, 허용 오차 ±${fmtLogicNumber(tol, 1)} m`);
      cards.formula.push(`판정식: |equivalent_length - ${fmtLogicNumber(ref, 1)}| <= ${fmtLogicNumber(tol, 1)}`);
      cards.values.push(`실제 계산: |${fmtLogicNumber(eq)} - ${fmtLogicNumber(ref, 1)}| = ${fmtLogicNumber(diff)}`);
      if (row.highlight) {
        cards.conclusion.push(`빨강(기준 위반): 편차 ${fmtLogicNumber(diff)} m가 허용 오차 ${fmtLogicNumber(tol, 1)} m를 초과하여 P/V 등가길이 기준에 맞지 않습니다.`);
      } else {
        cards.conclusion.push(`판정 결과: 편차 ${fmtLogicNumber(diff)} m로 허용 오차 이내이므로 적합합니다.`);
      }
    } else {
      cards.criteria.push(`현재 항목은 ${desc || "-"} 설비입니다.`);
      cards.conclusion.push("별도 세부 기준보다 결과서 기재값을 참고용으로 표시합니다.");
    }
    if (row.warn) cards.conclusion.push("노랑(확인 필요): 결과서 내 필수 참조 항목 또는 대응 설비 정보가 부족하여 자동 판정을 보완 검토 대상으로 남겼습니다.");
    if (row.economy_flag) cards.conclusion.push("초록(경제성 검토 후보): 연결 배관 구경이 큰 밸브 계통으로 분류되어 밸브 및 부속류 단가 상승 가능성이 있으므로 100A 이하 대안 검토가 필요합니다.");
  } else if (tab === "valves") {
    const inlet = Number(row.inlet_pressure_kgf_cm2 || 0);
    const outlet = Number(row.outlet_pressure_kgf_cm2 || 0);
    const reportDrop = Number(row.pressure_drop_kgf_cm2 || 0);
    const calcDrop = Number(row.calculated_pressure_drop_kgf_cm2 ?? (inlet - outlet));
    const tol = Number(row.pressure_drop_tolerance_kgf_cm2 ?? 0.05);
    const diff = Math.abs(calcDrop - reportDrop);
    cards.criteria.push(`감압밸브 허용 기준: 편차 <= ${fmtLogicNumber(tol, 2)} kg/cm²G`);
    cards.formula.push("기본식: calculated_drop = inlet_pressure - outlet_pressure");
    cards.formula.push(`비교식: |calculated_drop - reported_drop| <= ${fmtLogicNumber(tol, 2)}`);
    cards.values.push(`실제 계산: ${fmtLogicNumber(inlet)} - ${fmtLogicNumber(outlet)} = ${fmtLogicNumber(calcDrop)} kg/cm²G`);
    cards.values.push(`편차 계산: |${fmtLogicNumber(calcDrop)} - ${fmtLogicNumber(reportDrop)}| = ${fmtLogicNumber(diff)} kg/cm²G`);
    if (row.highlight) {
      cards.conclusion.push(`빨강(기준 위반): 편차 ${fmtLogicNumber(diff)} kg/cm²G가 허용 오차 ${fmtLogicNumber(tol, 2)} kg/cm²G를 초과하여 결과서 압력강하 값과 계산값이 일치하지 않습니다.`);
    } else {
      cards.conclusion.push(`판정 결과: 편차 ${fmtLogicNumber(diff)} kg/cm²G로 허용 오차 이내이므로 적합합니다.`);
    }
  }
  if (!cards.criteria.length && !cards.formula.length && !cards.values.length && !cards.conclusion.length) {
    cards.conclusion.push("이 항목은 현재 자동 필터 규칙에 해당하지 않습니다.");
  }
  return cards;
}
function openLogicModal(title, cards) {
  logicModalTitleEl.textContent = title;
  const sections = [
    ["criteria", "판정 기준", "criteria"],
    ["formula", "계산식", "formula"],
    ["values", "대입값", "values"],
    ["conclusion", "최종 해석", "conclusion"],
  ];
  logicModalBodyEl.innerHTML = sections
    .map(([key, label, cls]) => {
      const items = Array.isArray(cards?.[key]) && cards[key].length ? cards[key] : ["해당 내용 없음"];
      return `<section class="logic-card ${cls}"><h4>${label}</h4><ul>${items.map((item) => `<li>${escapeHtml(item)}</li>`).join("")}</ul></section>`;
    })
    .join("");
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

function updatesHistoryHtml() {
  return `
    <section class="criteria-section"><h4>1) 초기 검증 서버 구성</h4><ul>
      <li>결과서(docx/pdf)와 SDF 파일 업로드 후 PIPENET 검증이 가능하도록 기본 서버/화면 구성</li>
      <li>검증 결과, 통계, 상세리포트, 결과 데이터 테이블 기본 출력 기능 추가</li>
    </ul></section>
    <section class="criteria-section"><h4>2) 검증 로직 고도화</h4><ul>
      <li>배관 유속 기준 반영: 50A 이하 6.0 m/s, 50A 초과 10.0 m/s</li>
      <li>헤드 유량/압력, FX 등가길이, AV/PV 기준값 검증 추가</li>
      <li>헤드 유량편차 과다 평가 항목 제거</li>
    </ul></section>
    <section class="criteria-section"><h4>3) 설계 최적화 가이드 추가</h4><ul>
      <li>시공사 경제성 확보 방안 추가</li>
      <li>공학적 마찰손실 최적화 원칙 추가</li>
      <li>공학 최적화 후보, 경제성 검토 후보를 별도 가이드로 분리 표시</li>
    </ul></section>
    <section class="criteria-section"><h4>4) 결과 데이터 테이블 개선</h4><ul>
      <li>표 형식 정리, 한국어 설명 강화, 색상 범례 추가</li>
      <li>빨강/노랑/파랑/초록 기준으로 필터 로직 사유 확인 기능 추가</li>
      <li>10열까지만 기본 표시하고 나머지는 추가 표 영역으로 분리</li>
      <li>엑셀 다운로드 기능 추가 및 4개 시트 분할, 색상/경계 스타일 반영</li>
    </ul></section>
    <section class="criteria-section"><h4>5) 검진 통계/시각화 개선</h4><ul>
      <li>검진 통계 한글화 및 레이아웃 정리</li>
      <li>Matplotlib 기반 Pipe Velocity, Nozzle Pressure-Flow 차트 추가</li>
      <li>통계 하단 표를 결과서 파일 통계 / SDF 파일 통계로 분리</li>
    </ul></section>
    <section class="criteria-section"><h4>6) SDF 아이소매트릭 배관망 개선</h4><ul>
      <li>좌우 분할 화면, 줌/이동 기능, 범례 추가</li>
      <li>노즐/밸브/특수설비 크기 축소, 배관선 두께 조정</li>
      <li>적합/부적합 필터 시 관련 노드/엣지만 강조하는 로직 보완</li>
      <li>x축 반전 문제와 초기 비활성 표시 문제 보정</li>
    </ul></section>
    <section class="criteria-section"><h4>7) 좌측 메뉴 구조 개편</h4><ul>
      <li>1. 검증결과 / 2. 설계 최적화 가이드 / 3. 결과 데이터 테이블 / 4. 검진 통계 / 5. 상세리포트 / 6. CAD-아이소매트릭 배관망 대조 모듈 구성</li>
      <li>각 메뉴 클릭 시에만 해당 상세 인터페이스가 열리도록 변경</li>
      <li>초기 화면에서도 좌측 메뉴가 보이도록 조정</li>
    </ul></section>
    <section class="criteria-section"><h4>8) CAD-아이소매트릭 배관망 대조 모듈 추가</h4><ul>
      <li>DXF 뷰어와 SDF 뷰어를 나란히 두는 듀얼 뷰어 모듈 추가</li>
      <li>선택 영역 지정 후 핵심 구성요소 대조 분석 기능 추가</li>
      <li>브라우저 DXF 파서 대신 서버 ezdxf 파싱 방식으로 변경하여 호환성 개선</li>
      <li>선택 영역 외 삭제, 캐시 갱신, 화면 반영 문제 디버깅 및 보완</li>
      <li>DXF 텍스트 렌더링 크기 축소</li>
    </ul></section>
    <section class="criteria-section"><h4>9) 운영/배포/문서화</h4><ul>
      <li>다른 PC 접속 가능하도록 서버 개방 방식 안내 및 적용</li>
      <li>설명서 DOCX/HWP 작성 및 인코딩/변환 문제 보정</li>
      <li>GitHub 저장소 업로드 및 업데이트 기록 버튼 요청 반영</li>
    </ul></section>
  `;
}

function renderUpdateHistory(history) {
  const title = history?.title || "업데이트 기록";
  const updatedAt = history?.updated_at || "";
  const items = Array.isArray(history?.items) ? history.items : [];
  if (updatesModalTitleEl) updatesModalTitleEl.textContent = title;
  updatesModalBodyEl.innerHTML = `
    <section class="update-hero">
      <div>
        <p class="update-kicker">PATCH NOTES</p>
        <h4>${escapeHtml(title)}</h4>
      </div>
      <div class="update-meta">${escapeHtml(updatedAt ? `최종 반영: ${updatedAt}` : "최신 기록")}</div>
    </section>
    <div class="update-timeline">
      ${
        items.length
          ? items
              .map(
                (item) => `
                  <article class="update-card">
                    <div class="update-card-head">
                      <span class="update-date">${escapeHtml(item.timestamp || item.date || "")}</span>
                      <h5>${escapeHtml(item.title || "")}</h5>
                    </div>
                    <p class="update-summary">${escapeHtml(item.summary || "")}</p>
                    ${
                      Array.isArray(item.details) && item.details.length
                        ? `<ul class="update-list">${item.details.map((detail) => `<li>${escapeHtml(detail)}</li>`).join("")}</ul>`
                        : ""
                    }
                  </article>
                `
              )
              .join("")
          : `<article class="update-card"><p class="update-summary">표시할 업데이트 기록이 없습니다.</p></article>`
      }
    </div>
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
setActiveMenuPanel(null);
setMenuLayoutVisible(true);
if (updatesBtnEl) updatesBtnEl.textContent = "업데이트 기록";
if (criteriaBtnEl) criteriaBtnEl.textContent = "평가 기준";
if (updatesModalTitleEl) updatesModalTitleEl.textContent = "업데이트 기록";
if (updatesModalCloseEl) updatesModalCloseEl.textContent = "닫기";
if (reportFileLabelEl) reportFileLabelEl.textContent = "결과서 파일 (docx)";

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
if (updatesBtnEl) {
  updatesBtnEl.addEventListener("click", async () => {
    try {
      const resp = await fetch("/api/update-history", { cache: "no-store" });
      const data = await resp.json();
      if (!resp.ok || !data.ok) throw new Error(data.message || "업데이트 기록을 불러오지 못했습니다.");
      renderUpdateHistory(data.history || {});
    } catch (err) {
      updatesModalBodyEl.innerHTML = updatesHistoryHtml();
    }
    updatesModalEl.classList.remove("hidden");
  });
}
criteriaModalCloseEl.addEventListener("click", () => criteriaModalEl.classList.add("hidden"));
if (updatesModalCloseEl) updatesModalCloseEl.addEventListener("click", () => updatesModalEl.classList.add("hidden"));
criteriaModalEl.addEventListener("click", (e) => {
  if (e.target === criteriaModalEl) criteriaModalEl.classList.add("hidden");
});
if (updatesModalEl) {
  updatesModalEl.addEventListener("click", (e) => {
    if (e.target === updatesModalEl) updatesModalEl.classList.add("hidden");
  });
}
downloadExcelBtnEl.addEventListener("click", downloadTablesAsXlsx);
if (cadRunBtnEl) {
  cadRunBtnEl.addEventListener("click", async () => {
    const cadFile = cadRefFileEl?.files?.[0];
    if (!cadFile) {
      cadStatusEl.textContent = "CAD 기준 파일(DXF)을 먼저 선택해 주세요.";
      return;
    }
    cadStatusEl.textContent = "CAD 대조 중입니다...";
    cadRunBtnEl.disabled = true;
    try {
      const fd = new FormData();
      fd.append("cad_file", cadFile);
      if (cadIsoFileEl?.files?.[0]) fd.append("sdf_file", cadIsoFileEl.files[0]);
      if (cadUseYoloEl?.checked) fd.append("use_yolo", "1");
      const resp = await fetch("/api/cad-compare", { method: "POST", body: fd });
      const data = await resp.json();
      if (!resp.ok || !data.ok) throw new Error(data.message || "CAD 대조 요청에 실패했습니다.");
      cadStatusEl.textContent = data.message || "CAD 대조 완료";
      renderCadSummary(data);
      renderCadMessages(data.messages || []);
      renderCadSvg(data.cad_payload || {}, data.detected_heads || []);
    } catch (e) {
      cadStatusEl.textContent = e.message || "CAD 대조 중 오류가 발생했습니다.";
      cadSummaryEl.innerHTML = '<p class="empty">결과가 없습니다.</p>';
      cadMessagesEl.innerHTML = "<li>대조 결과를 불러오지 못했습니다.</li>";
      if (cadSvgEl) cadSvgEl.innerHTML = "";
    } finally {
      cadRunBtnEl.disabled = false;
    }
  });
}
window.addEventListener("keydown", (e) => {
  if (e.key === "Escape") {
    logicModalEl.classList.add("hidden");
    criteriaModalEl.classList.add("hidden");
    if (updatesModalEl) updatesModalEl.classList.add("hidden");
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
  setMenuLayoutVisible(true);
  workspacePanelEl.classList.add("hidden");
  insightPanelEl.classList.add("hidden");
  tablesPanelEl.classList.add("hidden");
  statsPanelEl.classList.add("hidden");
  reportPanelEl.classList.add("hidden");
  if (cadComparePanelEl) cadComparePanelEl.classList.add("hidden");

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
    renderStats(data.stats || {}, data.visualizations || []);
    renderReport(data);
    setMenuLayoutVisible(true);
    setActiveMenuPanel(currentMenuPanel || "workspace-panel");
    downloadExcelBtnEl.disabled = false;
  } catch (e) {
    statusEl.textContent = e.message;
    resultsBody.innerHTML = '<tr><td colspan="2" class="empty">결과를 불러오지 못했습니다.</td></tr>';
    setMenuLayoutVisible(true);
    setActiveMenuPanel("workspace-panel");
    networkEmptyEl.textContent = "배관망을 불러오지 못했습니다.";
    networkEmptyEl.classList.remove("hidden");
    networkSvgEl.classList.add("hidden");
  }
});
