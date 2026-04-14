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
const pipeRulesBtnEl = document.getElementById("pipe-rules-btn");
const insightPanelEl = document.getElementById("insight-panel");
const engineeringListEl = document.getElementById("engineering-list");
const engineeringVisualsEl = document.getElementById("engineering-visuals");
const economyListEl = document.getElementById("economy-list");
const economyVisualsEl = document.getElementById("economy-visuals");
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
const sdfAnalysisPanelEl = document.getElementById("sdf-analysis-panel");
const feedbackPanelEl = document.getElementById("feedback-panel");
const sdfAnalysisFormEl = document.getElementById("sdf-analysis-form");
const sdfAnalysisFileEl = document.getElementById("sdf-analysis-file");
const sdfAnalysisCadFileEl = document.getElementById("sdf-analysis-cad-file");
const sdfAnalysisRunBtnEl = document.getElementById("sdf-analysis-run-btn");
const sdfAnalysisStatusEl = document.getElementById("sdf-analysis-status");
const sdfAnalysisOutputEl = document.getElementById("sdf-analysis-output");
const feedbackFormEl = document.getElementById("feedback-form");
const feedbackAuthorEl = document.getElementById("feedback-author");
const feedbackTitleEl = document.getElementById("feedback-title");
const feedbackBodyEl = document.getElementById("feedback-body");
const feedbackAttachmentEl = document.getElementById("feedback-attachment");
const feedbackSubmitBtnEl = document.getElementById("feedback-submit-btn");
const feedbackRefreshBtnEl = document.getElementById("feedback-refresh-btn");
const feedbackStatusEl = document.getElementById("feedback-status");
const feedbackListEl = document.getElementById("feedback-list");
const feedbackCountEl = document.getElementById("feedback-count");
const feedbackWriteBtnEl = document.getElementById("feedback-write-btn");
const feedbackDetailModalEl = document.getElementById("feedback-detail-modal");
const feedbackDetailCloseEl = document.getElementById("feedback-detail-close");
const feedbackDetailBodyEl = document.getElementById("feedback-detail-body");
const feedbackDetailTitleEl = document.getElementById("feedback-detail-title");
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
const pipeRulesModalEl = document.getElementById("pipe-rules-modal");
const pipeRulesModalCloseEl = document.getElementById("pipe-rules-modal-close");
const pipeRulesModalBodyEl = document.getElementById("pipe-rules-modal-body");

let currentTables = null;
let currentRules = null;
let currentTab = "pipes";
let currentMenuPanel = null;
let currentEngineeringSpikes = [];
let currentFeedbackPosts = [];
let engineeringMapPan = null;
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
  max_other_pipe_velocity_mps: "그 밖의 배관 최대 유속 (m/s)",
  velocity_branch_checked_count: "가지배관 판정 수",
  velocity_other_checked_count: "그 밖의 배관 판정 수",
  velocity_review_count: "유속 판정 검토 수",
  velocity_failed_count: "유속 기준 초과 수",
  worst_velocity_pipe_label: "최대 유속 배관",
  max_pipe_pressure_kgcm2: "최대 관압 (kg/cm²G)",
  high_pressure_pipe_count: "12.0kg/cm²G 이상 배관 수",
  ksd3562_required_pipe_count: "KSD3562 요구 배관 수",
  c_factor_material_fail_count: "재질-C값 불일치 수",
  pipe_review_count: "배관 규칙 검토 수",
  sdf_main_title: "SDF 메인 제목",
  sdf_zone_title: "SDF 존",
  pipe_count: "SDF 배관 수",
  nozzle_count: "SDF 헤드 수",
  equipment_count: "SDF 특수설비 수",
};

const tableConfigs = {
  pipes: {
    title: "배관(FLOW IN PIPES)",
    columns: [["label", "Pipe"], ["input_node", "입력 노드"], ["output_node", "출력 노드"], ["nominal_bore_mm", "구경(mm)"], ["material_name", "재질"], ["pipe_type_id", "Pipe Type"], ["actual_bore_mm", "실내경(mm)"], ["c_factor", "C-Factor"], ["max_pressure_kgcm2", "최대 관압"], ["pipe_role", "배관 역할"], ["downstream_nozzle_count", "하류 헤드 수"], ["subtree_has_cross_split", "하류 교차분기"], ["velocity_limit_mps", "유속 기준(m/s)"], ["velocity_ok", "유속 적합"], ["engineering_reasons", "공학 후보 사유"], ["economy_reasons", "경제성 후보 사유"], ["rule_high_pressure_ok", "고압재질"], ["rule_cfactor_ok", "C값"], ["rule_unit_internal_cpvc_status", "세대내CPVC"], ["rule_unit_inlet_status", "세대유입65A"], ["rule_headcount_size_status", "헤드수-구경정책"], ["flow_lpm", "유량(L/min)"], ["velocity_mps", "유속(m/s)"], ["base_length_m", "배관길이(m)"], ["fitting_eq_length_m", "피팅 등가길이(m)"], ["special_eq_length_m", "특수설비 등가길이(m)"], ["total_length_m", "총 등가길이(m)"], ["friction_loss", "결과서 마찰손실"], ["hw_expected_friction_loss", "HW 재계산 마찰손실"], ["hw_abs_diff", "HW 절대오차"], ["hw_rel_diff", "HW 상대오차"], ["hw_formula_ok", "HW 검산 적합"], ["inlet_pressure", "입구압"], ["outlet_pressure", "출구압"], ["special_equipment", "특수설비"], ["hw_fail", "HW 실패"], ["role_reason", "판정 사유"]],
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
  if (v === null || v === undefined || v === "") return "-";
  return v;
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

function renderPipeRulesButton(rules) {
  currentRules = rules || {};
  const items = Array.isArray(currentRules?.pipe) ? currentRules.pipe.filter((x) => x.subject_type === "meta") : [];
  if (!pipeRulesBtnEl) return;
  pipeRulesBtnEl.classList.toggle("hidden", !items.length);
}

function openPipeRulesModal() {
  const items = Array.isArray(currentRules?.pipe) ? currentRules.pipe.filter((x) => x.subject_type === "meta") : [];
  if (!items.length || !pipeRulesModalEl || !pipeRulesModalBodyEl) return;
  pipeRulesModalBodyEl.innerHTML = items
    .map((item) => {
      const cls = item.status === "FAIL" ? "fail" : item.status === "REVIEW" ? "review" : "pass";
      return `<article class="criteria-section pipe-rule-card ${cls}">
        <h4><span class="pipe-rule-badge ${cls}">${escapeHtml(item.status)}</span> ${escapeHtml(item.rule_id)}</h4>
        <p>${escapeHtml(item.message)}</p>
      </article>`;
    })
    .join("");
  pipeRulesModalEl.classList.remove("hidden");
}

function renderGuideItem(text) {
  const raw = String(text ?? "");
  const markerMatch = raw.match(/(주요\s*(?:구간|배관|헤드|설비|노드)\s*:|二쇱슂\s*援ш컙\s*:)/);
  if (!markerMatch) return `<li>${escapeHtml(raw)}</li>`;

  const markerIndex = markerMatch.index ?? -1;
  const intro = raw.slice(0, markerIndex + markerMatch[0].length).trim();
  const listText = raw.slice(markerIndex + markerMatch[0].length).trim();
  const rows = [];
  const itemPattern = /([A-Za-z가-힣]+)\s*(\d+)\(([^)]*)\)/g;
  let match;
  while ((match = itemPattern.exec(listText)) !== null) {
    rows.push({ type: match[1], label: match[2], detail: match[3] });
  }

  if (!rows.length) return `<li>${escapeHtml(raw)}</li>`;

  const tableRows = rows
    .map(
      (row) => `<tr>
        <td>${escapeHtml(row.type)}</td>
        <td>${escapeHtml(row.label)}</td>
        <td>${escapeHtml(row.detail)}</td>
      </tr>`
    )
    .join("");

  return `<li class="guide-table-item">
    <p>${escapeHtml(intro)}</p>
    <div class="guide-mini-table-wrap">
      <table class="guide-mini-table">
        <thead><tr><th>구분</th><th>번호</th><th>판정값</th></tr></thead>
        <tbody>${tableRows}</tbody>
      </table>
    </div>
  </li>`;
}

function renderInsights(insights) {
  const eng = insights?.engineering_advice || [];
  const eco = insights?.economy_guide || [];
  const visuals = Array.isArray(insights?.engineering_visualizations) ? insights.engineering_visualizations : [];
  currentEngineeringSpikes = [];
  engineeringListEl.innerHTML = eng.length ? eng.map(renderGuideItem).join("") : "<li>해당 사항 없음</li>";
  economyListEl.innerHTML = eco.length ? eco.map(renderGuideItem).join("") : "<li>해당 사항 없음</li>";
  if (engineeringVisualsEl) {
    const chartHtml = visuals.length
      ? visuals
          .map((v) => {
            const spikes = Array.isArray(v.spike_points) ? v.spike_points : [];
            const offset = currentEngineeringSpikes.length;
            currentEngineeringSpikes.push(...spikes);
            const markers = spikes
              .map(
                (spike, i) => `<button type="button" class="friction-spike-marker" data-spike-index="${
                  offset + i
                }" style="left:${Number(spike.left_percent || 50).toFixed(2)}%; top:${Number(
                  spike.top_percent || 50
                ).toFixed(2)}%;" aria-label="Pipe ${escapeHtml(spike.label)} 급증 원인 보기" title="Pipe ${escapeHtml(
                  spike.label
                )} 급증 원인 보기"></button>`
              )
              .join("");
            return `<article class="insight-chart-card">
              <h4>${escapeHtml(v.title || "Friction Loss Ratio")}</h4>
              <p>${escapeHtml(v.description || "")}</p>
              <div class="chart-overlay-wrap">
                <img src="${escapeHtml(v.image_data_url || "")}" alt="${escapeHtml(v.title || "engineering chart")}">
                ${markers}
              </div>
              ${
                spikes.length
                  ? `<p class="spike-help">빨간 삼각형을 클릭하면 해당 배관의 손실 변화율 원인과 조치안을 확인할 수 있습니다.</p>`
                  : ""
              }
            </article>`;
          })
          .join("")
      : "";
    engineeringVisualsEl.innerHTML = renderEngineeringNetworkMap(currentGraph, currentTables?.pipes || []) + chartHtml;
  }
  if (economyVisualsEl) {
    economyVisualsEl.innerHTML = renderEconomyNetworkMap(currentGraph, currentTables?.pipes || []);
  }
}

function openEngineeringSpikeModal(index) {
  const spike = currentEngineeringSpikes[index];
  if (!spike) return;
  openLogicModal(`Pipe ${spike.label} 마찰손실 변화율 급증 원인`, spike.cards || {});
}

function renderEngineeringNetworkMap(graph, pipeRows) {
  const pipes = graph?.pipes || [];
  if (!pipes.length) {
    return `<article class="insight-network-card">
      <h4>Friction Loss Spike Map</h4>
      <p class="empty">SDF 파일을 업로드하면 마찰손실 급증 배관 위치가 표시됩니다.</p>
    </article>`;
  }

  const pipeInfoMap = new Map((pipeRows || []).map((row) => [String(row.label), row]));
  const pipeChangeMap = new Map();
  const ratioRows = (pipeRows || [])
    .map((row) => {
      const label = Number(row.label);
      const frictionLoss = Number(row.friction_loss ?? row.friction_loss_kgcm2);
      const lengthM = Number(row.base_length_m ?? row.length_m ?? row.pipe_length_m);
      const ratio = lengthM > 0 ? frictionLoss / lengthM : Number(row.friction_loss_ratio ?? row.friction_loss_per_m);
      return { label, frictionLoss, lengthM, ratio };
    })
    .filter((row) => Number.isFinite(row.label) && Number.isFinite(row.ratio))
    .sort((a, b) => a.label - b.label);
  ratioRows.forEach((row, index) => {
    if (index === 0) return;
    const prev = ratioRows[index - 1];
    const delta = row.ratio - prev.ratio;
    const changeRatePercent = (delta / Math.max(Math.abs(prev.ratio), 1e-9)) * 100;
    pipeChangeMap.set(String(row.label), {
      previousLabel: prev.label,
      previousRatio: prev.ratio,
      delta,
      changeRatePercent,
    });
  });
  const threshold = 1.0;
  const spikePipeIds = new Set(ratioRows.filter((row) => row.ratio > threshold).map((row) => String(row.label)));
  const fy = (y) => -Number(y);
  const pts = [];
  for (const pipe of pipes) for (const q of pipe.path || []) pts.push([Number(q[0]), fy(q[1])]);
  for (const n of graph.nozzles || []) pts.push([Number(n.x), fy(n.y)]);
  for (const e of graph.equipment || []) pts.push([Number(e.x), fy(e.y)]);
  for (const v of graph.valves || []) pts.push([Number(v.x), fy(v.y)]);
  if (!pts.length) return "";

  const xs = pts.map((p) => p[0]);
  const ys = pts.map((p) => p[1]);
  const pad = 220;
  const minX = Math.min(...xs) - pad;
  const minY = Math.min(...ys) - pad;
  const width = Math.max(...xs) - Math.min(...xs) + pad * 2;
  const height = Math.max(...ys) - Math.min(...ys) + pad * 2;
  const baseViewBox = `${minX} ${minY} ${width} ${height}`;

  const pipeEls = pipes
    .map((pipe) => {
      const isSpike = spikePipeIds.has(String(pipe.label));
      const row = pipeInfoMap.get(String(pipe.label)) || {};
      const points = (pipe.path || []).map((q) => `${Number(q[0])},${fy(q[1])}`).join(" ");
      const frictionLoss = fmtLogicNumber(row.friction_loss ?? row.friction_loss_kgcm2, 4);
      const lengthM = fmtLogicNumber(row.length_m ?? row.base_length_m, 3);
      const ratioSource =
        Number(row.base_length_m ?? row.length_m ?? row.pipe_length_m) > 0
          ? Number(row.friction_loss ?? row.friction_loss_kgcm2) /
            Number(row.base_length_m ?? row.length_m ?? row.pipe_length_m)
          : row.friction_loss_ratio ?? row.friction_loss_per_m;
      const ratioValue = Number(ratioSource);
      const ratio = Number.isFinite(ratioValue) ? fmtLogicNumber(ratioValue, 4) : "-";
      const velocity = fmtLogicNumber(row.velocity_mps ?? row.velocity, 3);
      const change = pipeChangeMap.get(String(pipe.label));
      const changeRate = change ? `${fmtLogicNumber(change.changeRatePercent, 1)}%` : "-";
      const previousPipe = change ? String(change.previousLabel) : "-";
      const deltaRatio = change ? fmtLogicNumber(change.delta, 4) : "-";
      const dataAttrs = isSpike
        ? ` class="spike-network-pipe" data-pipe-label="${escapeHtml(pipe.label)}" data-friction-loss="${escapeHtml(
            frictionLoss
          )}" data-length-m="${escapeHtml(lengthM)}" data-ratio="${escapeHtml(ratio)}" data-change-rate="${escapeHtml(
            changeRate
          )}" data-previous-pipe="${escapeHtml(previousPipe)}" data-delta-ratio="${escapeHtml(
            deltaRatio
          )}" data-velocity="${escapeHtml(velocity)}"`
        : "";
      const visiblePipe = `<polyline${dataAttrs} points="${points}" fill="none" stroke="${isSpike ? "#dc2626" : "#cbd5e1"}" stroke-width="${
        isSpike ? 22 : 10
      }" stroke-linecap="round" stroke-linejoin="round"><title>Pipe ${escapeHtml(pipe.label)}${
        isSpike ? " - 마찰손실 급증 배관" : ""
      }</title></polyline>`;
      if (!isSpike) return visiblePipe;
      return `${visiblePipe}<polyline class="spike-network-hitbox" data-pipe-label="${escapeHtml(
        pipe.label
      )}" data-friction-loss="${escapeHtml(frictionLoss)}" data-length-m="${escapeHtml(
        lengthM
      )}" data-ratio="${escapeHtml(ratio)}" data-change-rate="${escapeHtml(
        changeRate
      )}" data-previous-pipe="${escapeHtml(previousPipe)}" data-delta-ratio="${escapeHtml(
        deltaRatio
      )}" data-velocity="${escapeHtml(velocity)}" points="${points}" fill="none" stroke="transparent" stroke-width="70" stroke-linecap="round" stroke-linejoin="round"></polyline>`;
    })
    .join("");
  const nozzleEls = (graph.nozzles || [])
    .map((n) => `<circle cx="${Number(n.x)}" cy="${fy(n.y)}" r="12" fill="#94a3b8" stroke="#334155" stroke-width="3" />`)
    .join("");
  const equipmentEls = (graph.equipment || [])
    .map((e) => `<rect x="${Number(e.x) - 14}" y="${fy(e.y) - 14}" width="28" height="28" fill="#94a3b8" stroke="#334155" stroke-width="3" />`)
    .join("");
  const valveEls = (graph.valves || [])
    .map((v) => {
      const x = Number(v.x);
      const y = fy(v.y);
      return `<polygon points="${x},${y - 18} ${x + 18},${y} ${x},${y + 18} ${x - 18},${y}" fill="#94a3b8" stroke="#334155" stroke-width="3" />`;
    })
    .join("");

  return `<article class="insight-network-card">
    <h4>Friction Loss Spike Map</h4>
    <p>빨간 배관은 m당 마찰손실 기준(1.00 kg/cm²/m)을 초과한 공학 최적화 후보입니다. 빨간 배관 위에서 드래그하면 간단 정보를 확인할 수 있습니다.</p>
    <div class="insight-network-stage">
      <div class="insight-network-controls" aria-label="Friction Loss Spike Map zoom controls">
        <button type="button" data-network-zoom="in" aria-label="확대">+</button>
        <button type="button" data-network-zoom="out" aria-label="축소">-</button>
      </div>
      <svg class="insight-network-svg" viewBox="${baseViewBox}" data-base-viewbox="${baseViewBox}" preserveAspectRatio="xMidYMid meet">
        <g>${pipeEls}${nozzleEls}${equipmentEls}${valveEls}</g>
      </svg>
    </div>
    <div class="network-legend insight-network-legend">
      <span class="network-legend-item"><i class="shape-line"></i>일반 배관</span>
      <span class="network-legend-item"><i class="shape-line shape-line-red"></i>마찰손실 급증 배관</span>
      <span class="network-legend-item"><i class="shape-circle"></i>헤드</span>
      <span class="network-legend-item"><i class="shape-square"></i>특수설비</span>
      <span class="network-legend-item"><i class="shape-diamond"></i>감압밸브</span>
    </div>
  </article>`;
}

function renderEconomyNetworkMap(graph, pipeRows) {
  const pipes = graph?.pipes || [];
  if (!pipes.length) {
    return `<article class="insight-network-card">
      <h4>Economy Optimization Candidate Map</h4>
      <p class="empty">SDF 파일을 업로드하면 경제성 검토 후보 배관 위치가 표시됩니다.</p>
    </article>`;
  }

  const pipeInfoMap = new Map((pipeRows || []).map((row) => [String(row.label), row]));
  const economyPipeIds = new Set((pipeRows || []).filter((row) => row.economy_flag).map((row) => String(row.label)));
  const fy = (y) => -Number(y);
  const pts = [];
  for (const pipe of pipes) for (const q of pipe.path || []) pts.push([Number(q[0]), fy(q[1])]);
  for (const n of graph.nozzles || []) pts.push([Number(n.x), fy(n.y)]);
  for (const e of graph.equipment || []) pts.push([Number(e.x), fy(e.y)]);
  for (const v of graph.valves || []) pts.push([Number(v.x), fy(v.y)]);
  if (!pts.length) return "";

  const xs = pts.map((p) => p[0]);
  const ys = pts.map((p) => p[1]);
  const pad = 220;
  const minX = Math.min(...xs) - pad;
  const minY = Math.min(...ys) - pad;
  const width = Math.max(...xs) - Math.min(...xs) + pad * 2;
  const height = Math.max(...ys) - Math.min(...ys) + pad * 2;
  const baseViewBox = `${minX} ${minY} ${width} ${height}`;

  const pipeEls = pipes
    .map((pipe) => {
      const isEconomy = economyPipeIds.has(String(pipe.label));
      const row = pipeInfoMap.get(String(pipe.label)) || {};
      const points = (pipe.path || []).map((q) => `${Number(q[0])},${fy(q[1])}`).join(" ");
      const reason = row.economy_reasons || "경제성 검토 후보";
      const velocity = fmtLogicNumber(row.velocity_mps ?? row.velocity, 3);
      const dataAttrs = isEconomy
        ? ` class="spike-network-pipe economy-network-pipe" data-pipe-label="${escapeHtml(pipe.label)}" data-friction-loss="${escapeHtml(
            fmtLogicNumber(row.friction_loss, 4)
          )}" data-length-m="${escapeHtml(fmtLogicNumber(row.base_length_m ?? row.pipe_length_m, 3))}" data-ratio="${escapeHtml(
            reason
          )}" data-change-rate="-" data-previous-pipe="-" data-delta-ratio="${escapeHtml(
            `구경 ${fmtLogicNumber(row.nominal_bore_mm, 0)}A`
          )}" data-velocity="${escapeHtml(velocity)}"`
        : "";
      const visiblePipe = `<polyline${dataAttrs} points="${points}" fill="none" stroke="${isEconomy ? "#16a34a" : "#cbd5e1"}" stroke-width="${
        isEconomy ? 22 : 10
      }" stroke-linecap="round" stroke-linejoin="round"><title>Pipe ${escapeHtml(pipe.label)}${
        isEconomy ? ` - ${escapeHtml(reason)}` : ""
      }</title></polyline>`;
      if (!isEconomy) return visiblePipe;
      return `${visiblePipe}<polyline class="spike-network-hitbox economy-network-hitbox" data-pipe-label="${escapeHtml(
        pipe.label
      )}" data-friction-loss="${escapeHtml(fmtLogicNumber(row.friction_loss, 4))}" data-length-m="${escapeHtml(
        fmtLogicNumber(row.base_length_m ?? row.pipe_length_m, 3)
      )}" data-ratio="${escapeHtml(reason)}" data-change-rate="-" data-previous-pipe="-" data-delta-ratio="${escapeHtml(
        `구경 ${fmtLogicNumber(row.nominal_bore_mm, 0)}A`
      )}" data-velocity="${escapeHtml(velocity)}" points="${points}" fill="none" stroke="transparent" stroke-width="70" stroke-linecap="round" stroke-linejoin="round"></polyline>`;
    })
    .join("");
  const nozzleEls = (graph.nozzles || [])
    .map((n) => `<circle cx="${Number(n.x)}" cy="${fy(n.y)}" r="12" fill="#94a3b8" stroke="#334155" stroke-width="3" />`)
    .join("");
  const equipmentEls = (graph.equipment || [])
    .map((e) => `<rect x="${Number(e.x) - 14}" y="${fy(e.y) - 14}" width="28" height="28" fill="#94a3b8" stroke="#334155" stroke-width="3" />`)
    .join("");
  const valveEls = (graph.valves || [])
    .map((v) => {
      const x = Number(v.x);
      const y = fy(v.y);
      return `<polygon points="${x},${y - 18} ${x + 18},${y} ${x},${y + 18} ${x - 18},${y}" fill="#94a3b8" stroke="#334155" stroke-width="3" />`;
    })
    .join("");

  return `<article class="insight-network-card">
    <h4>Economy Optimization Candidate Map</h4>
    <p>초록 배관은 저유속 과설계, 관경 축소 가능성, CPVC 대구경 등 경제성 검토 후보입니다. 공학 맵의 빨간 급증 구간과 분리해서 비교합니다.</p>
    <div class="insight-network-stage">
      <div class="insight-network-controls" aria-label="Economy Optimization Candidate Map zoom controls">
        <button type="button" data-network-zoom="in" aria-label="확대">+</button>
        <button type="button" data-network-zoom="out" aria-label="축소">-</button>
      </div>
      <svg class="insight-network-svg" viewBox="${baseViewBox}" data-base-viewbox="${baseViewBox}" preserveAspectRatio="xMidYMid meet">
        <g>${pipeEls}${nozzleEls}${equipmentEls}${valveEls}</g>
      </svg>
    </div>
    <div class="network-legend insight-network-legend">
      <span class="network-legend-item"><i class="shape-line"></i>일반 배관</span>
      <span class="network-legend-item"><i class="shape-line shape-line-green"></i>경제성 검토 후보</span>
      <span class="network-legend-item"><i class="shape-circle"></i>헤드</span>
      <span class="network-legend-item"><i class="shape-square"></i>특수설비</span>
      <span class="network-legend-item"><i class="shape-diamond"></i>감압밸브</span>
    </div>
  </article>`;
}

function getEngineeringMapTooltipEl() {
  let el = document.getElementById("engineering-map-tooltip");
  if (!el) {
    el = document.createElement("div");
    el.id = "engineering-map-tooltip";
    el.className = "engineering-map-tooltip hidden";
    document.body.appendChild(el);
  }
  return el;
}

function showEngineeringMapTooltip(event, target) {
  const el = getEngineeringMapTooltipEl();
  el.innerHTML = `
    <strong>Pipe ${escapeHtml(target.dataset.pipeLabel || "-")}</strong>
    <span>마찰손실: ${escapeHtml(target.dataset.frictionLoss || "-")} kg/cm²</span>
    <span>배관길이: ${escapeHtml(target.dataset.lengthM || "-")} m</span>
    <span>m당 마찰손실: ${escapeHtml(target.dataset.ratio || "-")} kg/cm²/m</span>
    <span>변화율: ${escapeHtml(target.dataset.changeRate || "-")} (직전 Pipe ${escapeHtml(
      target.dataset.previousPipe || "-"
    )} 대비)</span>
    <span>증가량: ${escapeHtml(target.dataset.deltaRatio || "-")} kg/cm²/m</span>
    <span>유속: ${escapeHtml(target.dataset.velocity || "-")} m/s</span>
  `;
  const margin = 14;
  el.style.left = `${event.clientX + margin}px`;
  el.style.top = `${event.clientY + margin}px`;
  el.classList.remove("hidden");
}

function hideEngineeringMapTooltip() {
  document.getElementById("engineering-map-tooltip")?.classList.add("hidden");
}

function parseSvgViewBox(svg) {
  const raw = svg?.getAttribute("viewBox") || svg?.dataset.baseViewbox || "";
  const nums = raw.split(/\s+/).map(Number);
  if (nums.length !== 4 || nums.some((n) => !Number.isFinite(n))) return null;
  return { x: nums[0], y: nums[1], width: nums[2], height: nums[3] };
}

function setSvgViewBox(svg, box) {
  if (!svg || !box) return;
  svg.setAttribute("viewBox", `${box.x} ${box.y} ${box.width} ${box.height}`);
}

function zoomEngineeringNetworkMap(svg, direction) {
  const box = parseSvgViewBox(svg);
  if (!box) return;
  const factor = direction === "in" ? 0.8 : 1.25;
  const nextWidth = box.width * factor;
  const nextHeight = box.height * factor;
  const cx = box.x + box.width / 2;
  const cy = box.y + box.height / 2;
  setSvgViewBox(svg, {
    x: cx - nextWidth / 2,
    y: cy - nextHeight / 2,
    width: nextWidth,
    height: nextHeight,
  });
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

function renderMiniTable(headers, rows, emptyText = "데이터가 없습니다.") {
  if (!rows || !rows.length) return `<p class="empty">${escapeHtml(emptyText)}</p>`;
  return `<div class="table-wrap mini-table-wrap"><table><thead><tr>${headers
    .map((h) => `<th>${escapeHtml(h[1])}</th>`)
    .join("")}</tr></thead><tbody>${rows
    .map((row) => `<tr>${headers.map(([key]) => `<td>${escapeHtml(fmt(row[key]))}</td>`).join("")}</tr>`)
    .join("")}</tbody></table></div>`;
}

function renderSdfAnalysisGraph(analysis) {
  const nodes = analysis.nodes || [];
  const pipes = analysis.pipes || [];
  const nozzles = analysis.nozzles || [];
  if (!nodes.length) return '<p class="empty">SDF 좌표 데이터가 없습니다.</p>';
  const xs = nodes.map((n) => Number(n.x || 0));
  const ys = nodes.map((n) => Number(n.y || 0));
  const minX = Math.min(...xs);
  const maxX = Math.max(...xs);
  const minY = Math.min(...ys);
  const maxY = Math.max(...ys);
  const pad = Math.max(maxX - minX, maxY - minY, 1) * 0.08;
  const viewBox = `${minX - pad} ${minY - pad} ${maxX - minX + pad * 2} ${maxY - minY + pad * 2}`;
  const farHeadSet = new Set((analysis.farthest_heads || []).map((h) => String(h.label)));
  const pipeEls = pipes
    .map((p) => {
      const pts = (p.path || []).map((pt) => pt.join(",")).join(" ");
      const cls = p.status === "red" ? "sdf-pipe issue" : p.status === "orange" ? "sdf-pipe warn" : "sdf-pipe";
      return `<polyline class="${cls}" points="${escapeHtml(pts)}"><title>Pipe ${escapeHtml(p.label)} / ${fmt(p.bore_mm)}A / ${fmt(p.length_m)}m / ${escapeHtml(p.material || "-")}</title></polyline>`;
    })
    .join("");
  const nozzleEls = nozzles
    .filter((n) => Number.isFinite(Number(n.x)) && Number.isFinite(Number(n.y)))
    .map((n) => {
      const cls = farHeadSet.has(String(n.label)) ? "sdf-nozzle far" : "sdf-nozzle";
      return `<circle class="${cls}" cx="${Number(n.x)}" cy="${Number(n.y)}" r="${farHeadSet.has(String(n.label)) ? 18 : 12}"><title>Head ${escapeHtml(n.label)} / Node ${escapeHtml(n.input_node)}</title></circle>`;
    })
    .join("");
  return `<div class="sdf-analysis-map">
    <svg viewBox="${viewBox}" preserveAspectRatio="xMidYMid meet">
      <g transform="scale(1,-1) translate(0,${-(minY + maxY)})">${pipeEls}${nozzleEls}</g>
    </svg>
  </div>`;
}

function buildSvgBoxFromPoints(points) {
  const valid = (points || []).filter((p) => Number.isFinite(Number(p.x)) && Number.isFinite(Number(p.y)));
  if (!valid.length) return null;
  const xs = valid.map((p) => Number(p.x));
  const ys = valid.map((p) => Number(p.y));
  const minX = Math.min(...xs);
  const maxX = Math.max(...xs);
  const minY = Math.min(...ys);
  const maxY = Math.max(...ys);
  const pad = Math.max(maxX - minX, maxY - minY, 1) * 0.08;
  return { minX, maxX, minY, maxY, viewBox: `${minX - pad} ${minY - pad} ${maxX - minX + pad * 2} ${maxY - minY + pad * 2}` };
}

function renderCadAnalysisGraph(cad, comparison) {
  if (!cad) {
    return `<div class="sdf-analysis-map cad-analysis-map empty-map"><p class="empty">CAD DXF 파일을 함께 업로드하면 이 영역에 CAD 도면 후보가 표시됩니다.</p></div>`;
  }
  const entities = cad.drawing_entities || [];
  const candidates = cad.candidates || [];
  const matchedCadIds = new Set((comparison?.matches || []).map((m) => String(m.cad_candidate)));
  const points = [];
  for (const ent of entities) {
    for (const pt of ent.points || []) points.push({ x: pt[0], y: pt[1] });
  }
  for (const c of candidates) points.push(c);
  const box = buildSvgBoxFromPoints(points);
  if (!box) return `<div class="sdf-analysis-map cad-analysis-map empty-map"><p class="empty">CAD 좌표 데이터를 표시할 수 없습니다.</p></div>`;

  const lineEls = entities
    .map((ent) => {
      const pts = (ent.points || []).map((pt) => `${Number(pt[0])},${Number(pt[1])}`).join(" ");
      if (!pts) return "";
      return `<polyline class="cad-drawing-entity" points="${escapeHtml(pts)}"><title>${escapeHtml(ent.type || "LINE")} / ${escapeHtml(ent.layer || "-")}</title></polyline>`;
    })
    .join("");
  const candidateEls = candidates
    .map((c) => {
      const matched = matchedCadIds.has(String(c.label));
      const cls = matched ? "cad-head-candidate matched" : "cad-head-candidate";
      return `<circle class="${cls}" cx="${Number(c.x)}" cy="${Number(c.y)}" r="${matched ? 22 : 11}">
        <title>CAD 후보 ${escapeHtml(c.label)} / ${escapeHtml(c.layer || "-")}${matched ? " / SDF 자동 매칭 후보" : ""}</title>
      </circle>`;
    })
    .join("");
  return `<div class="sdf-analysis-map cad-analysis-map">
    <svg viewBox="${box.viewBox}" preserveAspectRatio="xMidYMid meet">
      <g transform="scale(1,-1) translate(0,${-(box.minY + box.maxY)})">${lineEls}${candidateEls}</g>
    </svg>
  </div>`;
}

function renderSdfCadDualGraph(data) {
  const analysis = data.analysis || {};
  const cad = data.cad_analysis || null;
  const comparison = data.comparison || null;
  return `<div class="sdf-cad-dual-map">
    <article class="sdf-map-pane">
      <div class="map-pane-head">
        <h4>SDF 좌표 기반 배관망</h4>
        <span>파랑: A/V 기준 최원단 헤드 후보</span>
      </div>
      ${renderSdfAnalysisGraph(analysis)}
    </article>
    <article class="sdf-map-pane">
      <div class="map-pane-head">
        <h4>CAD 도면 자동 인식 영역</h4>
        <span>빨강: SDF 헤드와 같다고 자동 매칭한 CAD 후보</span>
      </div>
      ${renderCadAnalysisGraph(cad, comparison)}
    </article>
  </div>`;
}

function renderCadSdfHeadComparison(data) {
  const cad = data.cad_analysis;
  const comparison = data.comparison;
  if (!cad || !comparison) {
    return `<article class="sdf-analysis-card">
      <h3>CAD-DXF 헤드 위치 대조</h3>
      <p class="empty">DXF 파일을 함께 업로드하면 CAD 도면의 헤드 후보 위치와 SDF 최원단 헤드 30개를 비교합니다.</p>
    </article>`;
  }
  return `<article class="sdf-analysis-card">
    <h3>CAD-DXF 헤드 위치 대조</h3>
    <div class="sdf-summary-grid">
      <div><strong>CAD 파일</strong><span>${escapeHtml(cad.filename || "-")}</span></div>
      <div><strong>CAD 헤드 후보</strong><span>${fmt(cad.candidate_count)}</span></div>
      <div><strong>CAD 선분 후보</strong><span>${fmt(cad.raw_line_count)}</span></div>
      <div><strong>비교 상태</strong><span>${escapeHtml(comparison.status || "-")}</span></div>
      <div><strong>불일치 후보</strong><span>${fmt(comparison.mismatch_count)}</span></div>
    </div>
    <p class="section-note">${escapeHtml(comparison.message || "")}</p>
    ${renderMiniTable(
      [
        ["sdf_head", "SDF Head"],
        ["sdf_node", "SDF Node"],
        ["cad_candidate", "CAD 후보"],
        ["cad_layer", "CAD Layer"],
        ["normalized_error", "정규화 오차"],
        ["status", "판정"],
        ["reason", "사유"],
      ],
      comparison.matches || [],
      "매칭 결과가 없습니다."
    )}
  </article>`;
}

function renderSdfAnalysis(data) {
  const a = data.analysis || {};
  const s = a.summary || {};
  const fittingRows = a.fitting_summary || [];
  const checklist = a.checklist || [];
  sdfAnalysisOutputEl.classList.remove("hidden");
  sdfAnalysisOutputEl.innerHTML = `
    <div class="sdf-analysis-grid">
      <article class="sdf-analysis-card">
        <h3>분석 요약</h3>
        <p class="section-note">${escapeHtml(a.title || a.filename || "-")}</p>
        <div class="sdf-summary-grid">
          <div><strong>노드</strong><span>${fmt(s.node_count)}</span></div>
          <div><strong>배관</strong><span>${fmt(s.pipe_count)}</span></div>
          <div><strong>헤드</strong><span>${fmt(s.nozzle_count)}</span></div>
          <div><strong>특수설비</strong><span>${fmt(s.equipment_count)}</span></div>
          <div><strong>A/V 추정 노드</strong><span>${fmt(s.av_node)}</span></div>
          <div><strong>A/V 배관</strong><span>${fmt(s.av_pipe_label)}</span></div>
        </div>
      </article>
      <article class="sdf-analysis-card">
        <h3>도면 대조 체크리스트</h3>
        <ul class="guide-list">${checklist.map((x) => `<li>${escapeHtml(x)}</li>`).join("")}</ul>
      </article>
    </div>
    <article class="sdf-analysis-card">
      <h3>SDF 좌표 기반 배관망</h3>
      <p class="section-note">빨강: 길이 대조 주의, 주황: 구경 축소 대조 포인트, 파란 헤드: A/V 기준 최원단 헤드 후보</p>
      ${renderSdfCadDualGraph(data)}
    </article>
    ${renderCadSdfHeadComparison(data)}
    <div class="sdf-analysis-grid">
      <article class="sdf-analysis-card">
        <h3>A/V 기준 최원단 헤드 30개</h3>
        ${renderMiniTable([["label", "Head"], ["input_node", "Node"], ["x", "X"], ["y", "Y"], ["z", "Z"], ["distance_from_av_m", "거리(m)"]], a.farthest_heads || [])}
      </article>
      <article class="sdf-analysis-card">
        <h3>배관 길이 대조 주의 구간</h3>
        ${renderMiniTable([["pipe_label", "Pipe"], ["sdf_length_m", "SDF 길이"], ["xy_length_m", "XY 길이"], ["diff_m", "차이"], ["reason", "사유"]], a.length_checks || [], "길이 대조 주의 구간이 없습니다.")}
      </article>
      <article class="sdf-analysis-card">
        <h3>구경 축소 지점</h3>
        ${renderMiniTable([["node", "Node"], ["from_pipe", "상류 Pipe"], ["from_bore_mm", "상류 구경"], ["to_pipe", "하류 Pipe"], ["to_bore_mm", "하류 구경"]], a.bore_reductions || [], "구경 축소 지점이 없습니다.")}
      </article>
      <article class="sdf-analysis-card">
        <h3>분기티/부속 대조 포인트</h3>
        ${renderMiniTable([["node", "분기 Node"], ["degree", "연결 차수"], ["x", "X"], ["y", "Y"]], a.branch_nodes || [], "분기 노드가 없습니다.")}
        <h4>부속류 집계</h4>
        ${renderMiniTable([["type", "부속"], ["count", "수량"]], fittingRows, "부속류 데이터가 없습니다.")}
      </article>
      <article class="sdf-analysis-card">
        <h3>엘보/티 집중 배관</h3>
        ${renderMiniTable([["pipe_label", "Pipe"], ["fitting_count", "부속 수"], ["fittings", "부속"], ["reason", "사유"]], a.fitting_hotspots || [], "부속 집중 구간이 없습니다.")}
      </article>
      <article class="sdf-analysis-card">
        <h3>수직 배관 확인 구간</h3>
        ${renderMiniTable([["pipe_label", "Pipe"], ["input_node", "입력"], ["output_node", "출력"], ["length_m", "길이"], ["rise_m", "Rise"], ["bore_mm", "구경"]], a.vertical_pipes || [], "수직 배관 확인 구간이 없습니다.")}
      </article>
    </div>
  `;
}

function renderFeedbackPosts(posts) {
  const rows = Array.isArray(posts) ? posts : [];
  currentFeedbackPosts = rows;
  if (feedbackCountEl) feedbackCountEl.textContent = `${rows.length}건`;
  if (!feedbackListEl) return;
  if (!rows.length) {
    feedbackListEl.innerHTML = '<p class="empty">등록된 개선의견이 없습니다.</p>';
    return;
  }
  feedbackListEl.innerHTML = rows
    .map(
      (post) => {
        const attachment = post.attachment || null;
        const attachmentHtml = attachment
          ? `<span class="feedback-list-attachment">첨부 1</span>`
          : "";
        return `<article class="feedback-post feedback-list-post" data-feedback-id="${escapeHtml(post.id || "")}" tabindex="0" role="button">
        <div class="feedback-post-head">
          <h4>${escapeHtml(post.title || "-")}</h4>
          <span>${escapeHtml(post.created_at || "-")}</span>
        </div>
        <div class="feedback-meta">작성자: ${escapeHtml(post.author || "익명")}</div>
        ${attachmentHtml}
      </article>`;
      }
    )
    .join("");
}

function openFeedbackDetail(postId) {
  const post = currentFeedbackPosts.find((item) => String(item.id) === String(postId));
  if (!post || !feedbackDetailModalEl || !feedbackDetailBodyEl) return;
  const attachment = post.attachment || null;
  const attachmentHtml = attachment
    ? `<a class="feedback-attachment" href="${escapeHtml(attachment.download_url || "#")}" download>${escapeHtml(attachment.original_name || "첨부파일")}</a>`
    : '<p class="empty">첨부파일 없음</p>';
  if (feedbackDetailTitleEl) feedbackDetailTitleEl.textContent = post.title || "개선의견 상세";
  feedbackDetailBodyEl.innerHTML = `<article class="feedback-detail-card">
    <div class="feedback-detail-meta">
      <span>작성자: ${escapeHtml(post.author || "익명")}</span>
      <span>작성일: ${escapeHtml(post.created_at || "-")}</span>
    </div>
    <div class="feedback-detail-content">${escapeHtml(post.body || "").replaceAll("\n", "<br>")}</div>
    <h4>첨부파일</h4>
    ${attachmentHtml}
  </article>`;
  feedbackDetailModalEl.classList.remove("hidden");
}

async function loadFeedbackPosts() {
  if (!feedbackListEl) return;
  try {
    if (feedbackStatusEl) feedbackStatusEl.textContent = "개선의견 목록을 불러오는 중입니다...";
    const resp = await fetch("/api/feedback-posts", { cache: "no-store" });
    const data = await resp.json();
    if (!resp.ok || !data.ok) throw new Error(data.message || "개선의견 목록을 불러오지 못했습니다.");
    renderFeedbackPosts(data.posts || []);
    if (feedbackStatusEl) feedbackStatusEl.textContent = "개선의견 목록을 불러왔습니다.";
  } catch (err) {
    if (feedbackStatusEl) feedbackStatusEl.textContent = err.message || "개선의견 목록 조회 중 오류가 발생했습니다.";
    feedbackListEl.innerHTML = '<p class="empty">개선의견 목록을 불러오지 못했습니다.</p>';
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
  const extra = extraCols.length ? `<div class="extra-columns-panel"><h3>${extraTitle}</h3><p class="table-note">우측 패널 테두리를 드래그하면 폭을 조절할 수 있습니다.</p>${buildTableHtml(extraCols, rows, MAX_MAIN_COLUMNS, "extra-wrap")}</div>` : "";
  tableContainerEl.innerHTML = `<div class="table-split">${buildTableHtml(mainCols, rows, 0, "main-wrap")}${extra}</div><p class="table-note">기본 창은 최대 ${MAX_MAIN_COLUMNS}열만 표시됩니다. 나머지는 추가 열 창에서 스크롤로 확인하세요.</p>`;
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
  if (appShellEl) appShellEl.classList.toggle("feedback-mode", panelId === "feedback-panel");
  if (panelId === "feedback-panel") loadFeedbackPosts();
  if (panelId !== "feedback-panel") feedbackPanelEl?.classList.remove("compose-open");
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
    const velocity = Number(row.velocity_mps || 0);
    const limit = row.velocity_limit_mps == null ? null : Number(row.velocity_limit_mps);
    const role = String(row.pipe_role || "review");
    const roleLabel = role === "branch" ? "branch" : role === "other" ? "other" : "review";
    const over = limit == null ? null : velocity - limit;
    cards.criteria.push(`Topology 판정: downstream nozzle 수 = ${fmt(row.downstream_nozzle_count)}, subtree cross split = ${fmt(row.subtree_has_cross_split)}`);
    cards.criteria.push(`최종 분류: ${roleLabel}`);
    cards.criteria.push(`판정 근거: ${row.role_reason || "-"}`);
    cards.formula.push("속도 기준: branch -> 6.0 m/s, other -> 10.0 m/s");
    if (role === "review") cards.formula.push("review -> 자동 기준 미적용, 수동 검토 필요");
    else cards.formula.push(`적용 기준식: V_actual <= ${fmtLogicNumber(limit, 3)} m/s`);
    cards.values.push(`실제 유속 V_actual = ${fmtLogicNumber(velocity)} m/s`);
    cards.values.push(`허용 유속 V_limit = ${limit == null ? "-" : fmtLogicNumber(limit)} m/s`);
    if (row.velocity_ok === false && limit != null) {
      cards.conclusion.push(`빨강(기준 위반): ${fmtLogicNumber(velocity)} - ${fmtLogicNumber(limit)} = ${fmtLogicNumber(over)} m/s 초과하여 유속 기준을 만족하지 못했습니다.`);
    } else if (row.velocity_ok === true && limit != null) {
      cards.conclusion.push(`적합 해석: ${fmtLogicNumber(limit)} - ${fmtLogicNumber(velocity)} = ${fmtLogicNumber(limit - velocity)} m/s 여유가 있어 유속 기준을 만족합니다.`);
    } else {
      cards.conclusion.push("노랑(확인 필요): topology ambiguity로 branch/other 판정이 확정되지 않아 자동 유속 판정을 보류했습니다.");
    }
    if (row.hw_fail) {
      cards.conclusion.push("추가 참고: 이 배관은 HW 마찰손실 검산도 실패 상태입니다.");
    }
    if (row.highlight && row.velocity_ok !== false && limit == null) {
      cards.conclusion.push("현재 highlight는 유속이 아니라 다른 검증 항목(HW 검산 등)에 의해 표시되었을 수 있습니다.");
    }
    if (row.engineering_flag) {
      const frictionLoss = Number(row.friction_loss || 0);
      const length = Number(row.pipe_length_m || 0);
      const unitLoss = length > 0 ? frictionLoss / length : null;
      const spike = Number(row.friction_spike_limit ?? 1.0);
      cards.criteria.push(`공학 후보 사유: ${row.engineering_reasons || "사유 미분류"}`);
      cards.criteria.push(`공학 최적화 기준: 단위 마찰손실 > ${fmtLogicNumber(spike, 3)} kg/cm²/m`);
      cards.formula.push(`단위 마찰손실 = friction_loss / length`);
      cards.values.push(`단위 마찰손실 = ${fmtLogicNumber(frictionLoss)} / ${fmtLogicNumber(length)} = ${fmtLogicNumber(unitLoss)} kg/cm²/m`);
      cards.conclusion.push(`파랑(공학 최적화 후보): ${fmtLogicNumber(unitLoss)} ${unitLoss > spike ? ">" : "<="} ${fmtLogicNumber(spike, 3)} 이므로 구경 상향, 피팅 축소, 배관 경로 단순화 검토가 필요합니다.`);
    }
    if (row.economy_flag) {
      const econLimit = Number(row.economy_velocity_limit ?? 2.0);
      cards.criteria.push(`경제성 후보 사유: ${row.economy_reasons || "사유 미분류"}`);
      cards.criteria.push(`경제성 기준: 유속 < ${fmtLogicNumber(econLimit, 1)} m/s AND 구경 > 25A`);
      cards.formula.push(`경제성 판정식: velocity < ${fmtLogicNumber(econLimit, 1)} AND bore > 25`);
      cards.values.push(`실제 값 대입: ${fmtLogicNumber(velocity)} < ${fmtLogicNumber(econLimit, 1)}, ${fmtLogicNumber(Number(row.nominal_bore_mm || 0), 0)}A > 25A`);
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
      <li><strong>Topology branch</strong>: 50A 이하 배관 중 downstream subtree에 multi-head split node가 없으면 6.0 m/s 이하</li>
      <li><strong>Topology other</strong>: 50A 초과 배관이거나, 50A 이하라도 downstream subtree에 multi-head split node가 있으면 10.0 m/s 이하</li>
      <li>배관 역할은 구경만으로 결정하지 않고, PIPE CONFIGURATION + NOZZLE CONFIGURATION 기반 topology 판정 결과로 확정합니다.</li>
      <li>기준 초과 시 빨강(기준 위반), topology ambiguity 시 노랑(확인 필요)으로 표시됩니다.</li>
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
    <section class="criteria-section"><h4>6) 모형 연동 해석</h4><ul>
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
      <li>배관 유속 기준을 단순 구경 판정에서 topology 기반 branch/other 판정으로 교체</li>
      <li>downstream nozzle count, subtree cross split 판정으로 40A/50A 교차배관을 other로 분류하도록 보강</li>
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

engineeringVisualsEl?.addEventListener("click", (event) => {
  const zoomBtn = event.target.closest("[data-network-zoom]");
  if (zoomBtn) {
    const svg = zoomBtn.closest(".insight-network-card")?.querySelector(".insight-network-svg");
    zoomEngineeringNetworkMap(svg, zoomBtn.dataset.networkZoom);
    return;
  }

  const marker = event.target.closest(".friction-spike-marker");
  if (!marker) return;
  openEngineeringSpikeModal(Number(marker.dataset.spikeIndex));
});

engineeringVisualsEl?.addEventListener("pointermove", (event) => {
  if (engineeringMapPan) {
    const { svg, startX, startY, startBox, tooltipTarget } = engineeringMapPan;
    const rect = svg.getBoundingClientRect();
    const dx = ((event.clientX - startX) * startBox.width) / Math.max(rect.width, 1);
    const dy = ((event.clientY - startY) * startBox.height) / Math.max(rect.height, 1);
    setSvgViewBox(svg, {
      x: startBox.x - dx,
      y: startBox.y - dy,
      width: startBox.width,
      height: startBox.height,
    });
    if (tooltipTarget) showEngineeringMapTooltip(event, tooltipTarget);
    return;
  }

  const target = event.target.closest(".spike-network-pipe, .spike-network-hitbox");
  if (!target) {
    hideEngineeringMapTooltip();
    return;
  }
  showEngineeringMapTooltip(event, target);
});

engineeringVisualsEl?.addEventListener("pointerdown", (event) => {
  const svg = event.target.closest(".insight-network-svg");
  const target = event.target.closest(".spike-network-pipe, .spike-network-hitbox");
  if (svg && event.button === 0) {
    const startBox = parseSvgViewBox(svg);
    if (!startBox) return;
    engineeringMapPan = {
      svg,
      startX: event.clientX,
      startY: event.clientY,
      startBox,
      tooltipTarget: target || null,
    };
    svg.classList.add("is-panning");
    svg.setPointerCapture?.(event.pointerId);
    event.preventDefault();
  }
  if (target) showEngineeringMapTooltip(event, target);
});

engineeringVisualsEl?.addEventListener("pointerup", () => {
  engineeringMapPan?.svg?.classList.remove("is-panning");
  engineeringMapPan = null;
});

engineeringVisualsEl?.addEventListener("pointercancel", () => {
  engineeringMapPan?.svg?.classList.remove("is-panning");
  engineeringMapPan = null;
  hideEngineeringMapTooltip();
});

engineeringVisualsEl?.addEventListener("pointerleave", () => {
  if (!engineeringMapPan) hideEngineeringMapTooltip();
});

economyVisualsEl?.addEventListener("click", (event) => {
  const zoomBtn = event.target.closest("[data-network-zoom]");
  if (!zoomBtn) return;
  const svg = zoomBtn.closest(".insight-network-card")?.querySelector(".insight-network-svg");
  zoomEngineeringNetworkMap(svg, zoomBtn.dataset.networkZoom);
});

economyVisualsEl?.addEventListener("pointermove", (event) => {
  if (engineeringMapPan) {
    const { svg, startX, startY, startBox, tooltipTarget } = engineeringMapPan;
    const rect = svg.getBoundingClientRect();
    const dx = ((event.clientX - startX) * startBox.width) / Math.max(rect.width, 1);
    const dy = ((event.clientY - startY) * startBox.height) / Math.max(rect.height, 1);
    setSvgViewBox(svg, {
      x: startBox.x - dx,
      y: startBox.y - dy,
      width: startBox.width,
      height: startBox.height,
    });
    if (tooltipTarget) showEngineeringMapTooltip(event, tooltipTarget);
    return;
  }

  const target = event.target.closest(".spike-network-pipe, .spike-network-hitbox");
  if (!target) {
    hideEngineeringMapTooltip();
    return;
  }
  showEngineeringMapTooltip(event, target);
});

economyVisualsEl?.addEventListener("pointerdown", (event) => {
  const svg = event.target.closest(".insight-network-svg");
  const target = event.target.closest(".spike-network-pipe, .spike-network-hitbox");
  if (svg && event.button === 0) {
    const startBox = parseSvgViewBox(svg);
    if (!startBox) return;
    engineeringMapPan = {
      svg,
      startX: event.clientX,
      startY: event.clientY,
      startBox,
      tooltipTarget: target || null,
    };
    svg.classList.add("is-panning");
    svg.setPointerCapture?.(event.pointerId);
    event.preventDefault();
  }
  if (target) showEngineeringMapTooltip(event, target);
});

economyVisualsEl?.addEventListener("pointerup", () => {
  engineeringMapPan?.svg?.classList.remove("is-panning");
  engineeringMapPan = null;
});

economyVisualsEl?.addEventListener("pointercancel", () => {
  engineeringMapPan?.svg?.classList.remove("is-panning");
  engineeringMapPan = null;
  hideEngineeringMapTooltip();
});

economyVisualsEl?.addEventListener("pointerleave", () => {
  if (!engineeringMapPan) hideEngineeringMapTooltip();
});

for (const b of tabButtons) b.addEventListener("click", () => setActiveTab(b.dataset.tab));
for (const b of menuButtons) b.addEventListener("click", () => setActiveMenuPanel(b.dataset.panel));
setActiveMenuPanel(null);
setMenuLayoutVisible(true);
if (updatesBtnEl) updatesBtnEl.textContent = "업데이트 기록";
if (criteriaBtnEl) criteriaBtnEl.textContent = "평가 기준";
if (pipeRulesBtnEl) pipeRulesBtnEl.textContent = "배관 규칙 상태 보기";
if (updatesModalTitleEl) updatesModalTitleEl.textContent = "업데이트 기록";
if (updatesModalCloseEl) updatesModalCloseEl.textContent = "닫기";
if (reportFileLabelEl) reportFileLabelEl.textContent = "결과서 파일 (docx)";
if (feedbackWriteBtnEl) feedbackWriteBtnEl.textContent = "의견작성";
if (feedbackRefreshBtnEl) feedbackRefreshBtnEl.textContent = "닫기";

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
pipeRulesBtnEl?.addEventListener("click", openPipeRulesModal);
pipeRulesModalCloseEl?.addEventListener("click", () => pipeRulesModalEl?.classList.add("hidden"));
pipeRulesModalEl?.addEventListener("click", (e) => {
  if (e.target === pipeRulesModalEl) pipeRulesModalEl.classList.add("hidden");
});
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
if (sdfAnalysisFormEl) {
  sdfAnalysisFormEl.addEventListener("submit", async (event) => {
    event.preventDefault();
    const sdfFile = sdfAnalysisFileEl?.files?.[0];
    if (!sdfFile) {
      sdfAnalysisStatusEl.textContent = "SDF 파일을 먼저 선택해 주세요.";
      return;
    }
    sdfAnalysisStatusEl.textContent = "SDF 스프링클러 배관 분석 중입니다...";
    if (sdfAnalysisRunBtnEl) sdfAnalysisRunBtnEl.disabled = true;
    try {
      const fd = new FormData();
      fd.append("sdf_file", sdfFile);
      if (sdfAnalysisCadFileEl?.files?.[0]) fd.append("cad_file", sdfAnalysisCadFileEl.files[0]);
      const resp = await fetch("/api/sdf-sprinkler-analysis", { method: "POST", body: fd });
      const data = await resp.json();
      if (!resp.ok || !data.ok) throw new Error(data.message || "SDF 분석 요청에 실패했습니다.");
      sdfAnalysisStatusEl.textContent = data.message || "SDF 분석이 완료되었습니다.";
      renderSdfAnalysis(data);
    } catch (e) {
      sdfAnalysisStatusEl.textContent = e.message || "SDF 분석 중 오류가 발생했습니다.";
      if (sdfAnalysisOutputEl) {
        sdfAnalysisOutputEl.classList.remove("hidden");
        sdfAnalysisOutputEl.innerHTML = '<p class="empty">SDF 분석 결과를 불러오지 못했습니다.</p>';
      }
    } finally {
      if (sdfAnalysisRunBtnEl) sdfAnalysisRunBtnEl.disabled = false;
    }
  });
}
if (feedbackWriteBtnEl) {
  feedbackWriteBtnEl.addEventListener("click", () => feedbackPanelEl?.classList.add("compose-open"));
}
if (feedbackRefreshBtnEl) {
  feedbackRefreshBtnEl.addEventListener("click", () => feedbackPanelEl?.classList.remove("compose-open"));
}
feedbackListEl?.addEventListener("click", (event) => {
  const postEl = event.target.closest(".feedback-list-post");
  if (postEl) openFeedbackDetail(postEl.dataset.feedbackId);
});
feedbackListEl?.addEventListener("keydown", (event) => {
  if (event.key !== "Enter" && event.key !== " ") return;
  const postEl = event.target.closest(".feedback-list-post");
  if (postEl) {
    event.preventDefault();
    openFeedbackDetail(postEl.dataset.feedbackId);
  }
});
feedbackDetailCloseEl?.addEventListener("click", () => feedbackDetailModalEl?.classList.add("hidden"));
feedbackDetailModalEl?.addEventListener("click", (event) => {
  if (event.target === feedbackDetailModalEl) feedbackDetailModalEl.classList.add("hidden");
});
if (feedbackFormEl) {
  feedbackFormEl.addEventListener("submit", async (event) => {
    event.preventDefault();
    const payload = {
      author: feedbackAuthorEl?.value || "",
      title: feedbackTitleEl?.value || "",
      body: feedbackBodyEl?.value || "",
    };
    if (!payload.title.trim() || !payload.body.trim()) {
      if (feedbackStatusEl) feedbackStatusEl.textContent = "제목과 내용을 모두 입력해주세요.";
      return;
    }
    feedbackSubmitBtnEl.disabled = true;
    if (feedbackStatusEl) feedbackStatusEl.textContent = "개선의견을 등록하는 중입니다...";
    try {
      const fd = new FormData();
      fd.append("author", payload.author);
      fd.append("title", payload.title);
      fd.append("body", payload.body);
      if (feedbackAttachmentEl?.files?.[0]) fd.append("attachment", feedbackAttachmentEl.files[0]);
      const resp = await fetch("/api/feedback-posts", {
        method: "POST",
        body: fd,
      });
      const data = await resp.json();
      if (!resp.ok || !data.ok) throw new Error(data.message || "개선의견 등록에 실패했습니다.");
      if (feedbackTitleEl) feedbackTitleEl.value = "";
      if (feedbackBodyEl) feedbackBodyEl.value = "";
      if (feedbackAttachmentEl) feedbackAttachmentEl.value = "";
      if (feedbackStatusEl) feedbackStatusEl.textContent = data.message || "개선의견이 등록되었습니다.";
      feedbackPanelEl?.classList.remove("compose-open");
      await loadFeedbackPosts();
    } catch (err) {
      if (feedbackStatusEl) feedbackStatusEl.textContent = err.message || "개선의견 등록 중 오류가 발생했습니다.";
    } finally {
      feedbackSubmitBtnEl.disabled = false;
    }
  });
  loadFeedbackPosts();
}
window.addEventListener("keydown", (e) => {
  if (e.key === "Escape") {
    logicModalEl.classList.add("hidden");
    criteriaModalEl.classList.add("hidden");
    if (updatesModalEl) updatesModalEl.classList.add("hidden");
    if (pipeRulesModalEl) pipeRulesModalEl.classList.add("hidden");
    if (feedbackDetailModalEl) feedbackDetailModalEl.classList.add("hidden");
    if (feedbackPanelEl) feedbackPanelEl.classList.remove("compose-open");
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
  if (sdfAnalysisPanelEl) sdfAnalysisPanelEl.classList.add("hidden");

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
    renderPipeRulesButton(data.rules || {});
    currentTables = data.tables || {};
    renderNetworkGraph(data.sdf_graph || null);
    renderInsights(data.insights || {});
    setActiveTab(currentTab);
    renderStats(data.stats || {}, data.visualizations || []);
    renderReport(data);
    setMenuLayoutVisible(true);
    setActiveMenuPanel(currentMenuPanel || "workspace-panel");
    downloadExcelBtnEl.disabled = false;
  } catch (e) {
    statusEl.textContent = e.message;
    resultsBody.innerHTML = '<tr><td colspan="2" class="empty">결과를 불러오지 못했습니다.</td></tr>';
    renderPipeRulesButton({});
    setMenuLayoutVisible(true);
    setActiveMenuPanel("workspace-panel");
    networkEmptyEl.textContent = "배관망을 불러오지 못했습니다.";
    networkEmptyEl.classList.remove("hidden");
    networkSvgEl.classList.add("hidden");
  }
});
