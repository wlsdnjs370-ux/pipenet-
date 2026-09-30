/* H-only progress surface and interaction guard. Loaded before F listeners so
 * keyboard shortcuts cannot edit an in-flight graph. No new requests, worker,
 * progress estimates, global fetch patches, or calculation state writes. */
(() => {
  "use strict";
  const $=id=>document.getElementById(id);
  const text=id=>$(id)?.textContent.trim() || "";
  const put=(id,value)=>{if(text(id)!==value) $(id).textContent=value;};
  let initialized=false, running=false, frame=0, started=0, ended=0, lastReport=0;
  let timer=null, terminalTimer=null, previousBusy="", previousLog="", previousLine="", baseStatus="", baseChip="";
  let logFresh=false, lineFresh=false, statusFresh=false, terminalResolved=false, boundStream=null, serverState=null, generation=0;
  const locked=new Map();
  const controls="button,input,select,textarea,a,[contenteditable=true],[role=button]";
  const safeIds=new Set(["h-more","h-task","h-table","h-all-columns","h-view","h-history","h-drawer-close",
    "h-activity-toggle","h-activity-dismiss","h-other-work","busy-cancel","btn-fit","dg-table",
    "dg-ins-close","mg-ins-close","ne-close","bore-color","dg-bore-color","ed-bg","ed-worst-view",
    "ed-gridline","dg-gridline","dg-plan-view","ha-original","ha-table","ha-show","he-fold","he-search"]);
  const busy=()=>initialized && !$('busy').classList.contains('hidden');
  function control(target) {return target instanceof Element ? target.closest(controls) : null;}
  function safe(node) {
    if(!node) return false;
    if(safeIds.has(node.id)||['hp-close','hp-fold','hp-kind','hp-filter'].includes(node.id)) return true;
    if(node.dataset.hDragHandle==="true" || node.dataset.hResizeHandle==="true") return true;
    if(["dg-plan","dg-under"].includes(node.id) && window.__mf?.edit) return true;
    return node.matches("h2.fold") || node.tagName==="SUMMARY";
  }
  function block(event) {
    event.preventDefault();event.stopImmediatePropagation();
    if($("h-activity-lock-note")) put("h-activity-lock-note","현재 도면은 처리 후 편집할 수 있습니다.");
  }
  // These capture listeners are registered before the native engine scripts.
  function guard(event) {
    if(!busy()) return;
    const target=event.target, node=control(target);
    if(event.type==="keydown") {
      const key=event.key.toLowerCase();
      if(["delete","backspace"].includes(key) || ((event.ctrlKey||event.metaKey)&&["z","x","v"].includes(key))) {block(event);return;}
      if(node && !safe(node) && !["tab","escape"].includes(key) && !((event.ctrlKey||event.metaKey)&&key==="c")) {block(event);return;}
      if(key==="enter" && target?.closest?.("#dg-grid tbody tr")) return;
      return;
    }
    if(target===$("cv")) {
      // Only camera gestures reach the engine; left clicks can edit in several modes.
      if(["pointerdown","mousedown","mouseup","pointerup"].includes(event.type) && [1,2].includes(event.button)) return;
      block(event);return;
    }
    if(target?.closest?.("#dg-grid tbody tr") && !node) return; // existing read-only row inspection
    if(target?.closest?.("summary,h2.fold") && !target?.closest?.("#slots,#steps")) return;
    if(node && !safe(node)) {
      // A form change generated before capture must not leak a changed value.
      const saved=locked.get(node);
      if(saved && ["input","change"].includes(event.type)) {
        if(saved.value!==undefined) node.value=saved.value;
        if(saved.checked!==undefined) node.checked=saved.checked;
      }
      block(event);
    }
  }
  for(const event of ["pointerdown","mousedown","pointerup","mouseup","click","dblclick","contextmenu","keydown","beforeinput","input","change","paste","drop"])
    window.addEventListener(event,guard,true);
  function markLocks() {
    for(const node of document.querySelectorAll(controls)) {
      if(safe(node) || locked.has(node)) continue;
      locked.set(node,{aria:node.getAttribute("aria-disabled"),value:"value" in node ? node.value : undefined,
        checked:"checked" in node ? node.checked : undefined});
      node.dataset.hBusyLocked="true";node.setAttribute("aria-disabled","true");
    }
  }
  function unlock() {
    for(const [node,saved] of locked) {
      delete node.dataset.hBusyLocked;
      if(saved.aria===null) node.removeAttribute("aria-disabled");else node.setAttribute("aria-disabled",saved.aria);
    }
    locked.clear();
  }
  function addEvent(message) {
    if(!message || message==="—" || message==="…") return;
    const list=$("h-activity-events");
    if(list.lastElementChild?.dataset.message===message) return;
    const li=document.createElement("li");li.dataset.message=message;
    li.textContent=((performance.now()-started)/1000).toFixed(1)+"초 · "+message;
    list.append(li);while(list.children.length>60) list.firstElementChild.remove();
  }
  function begin() {
    clearTimeout(terminalTimer);running=true;started=performance.now();ended=0;lastReport=started;
    previousBusy="";previousLog=text("log");previousLine=text("job-line");baseStatus=text("status");baseChip=text("job-chip");
    logFresh=false;lineFresh=false;statusFresh=false;terminalResolved=false;serverState=null;boundStream=null;generation++;
    $("h-activity-events").replaceChildren();
    $("h-activity").hidden=false;$("h-activity-dismiss").hidden=true;$("h-activity-result").hidden=true;
    $("h-activity").dataset.state="run";
    $("h-activity").classList.remove("h-minimized");
    $("h-activity-toggle").setAttribute("aria-expanded","true");put("h-activity-toggle","접기");
    put("h-activity-state","처리 중");put("h-activity-latest","현재 단계의 상세 보고를 기다리고 있습니다.");
    put("h-activity-lock-note","보기 가능 · 현재 도면 편집 대기");
    document.body.classList.add("h-processing");markLocks();
    window.ModuleHProgress?.start(window.__mf?.activityOperation);
    clearInterval(timer);timer=setInterval(()=>{connectStream();tick();},500);
  }
  function connectStream() {
    const es=window.__mf?.es;
    if(!es || es===boundStream) return;
    boundStream=es;
    const current=generation;
    // Read the existing connection. Never open a second stream or polling loop.
    es.addEventListener("state",event=>{
      if(window.__mf?.es!==es && es.readyState!==2) return;
      if(boundStream!==es || current!==generation) return;
      try {serverState=JSON.parse(event.data);lastReport=performance.now();queue();} catch (_) { /* native handler owns errors */ }
    });
  }
  function tick() {
    if(!initialized || $("h-activity").hidden) return;
    const now=ended || performance.now();
    put("h-activity-elapsed",((now-started)/1000).toFixed(1)+"초");
    if(!running) return;
    const transport=window.__mf?.es ? "실시간 수신" : window.__mf?.poll ? "주기 확인" : "처리 현황";
    put("h-activity-channel",transport);
    const age=Math.max(0,Math.floor((performance.now()-lastReport)/1000));
    put("h-activity-freshness",age<2 ? "방금 갱신" : "마지막 보고 "+age+"초 전");
    const bar=$("busy-bar"), percentage=parseFloat(bar.style.width);
    const measured=bar.classList.contains("det") && Number.isFinite(percentage);
    const track=bar.parentElement;
    if(measured) {
      track.setAttribute("aria-valuenow",String(Math.max(0,Math.min(100,percentage))));
      put("h-activity-measure","압축·전송 "+Math.round(percentage)+"%");
    } else {
      track.removeAttribute("aria-valuenow");
      put("h-activity-measure","");
    }
  }
  function terminal() {
    if(busy() || terminalResolved) return;
    const status=text("status"), fresh=statusFresh || status!==baseStatus, chip=text("job-chip");
    const review=window.ModuleHAttributes?.historyConflict();
    const held=review && status===review.conflict && serverState?.state!=="error";
    let state="ended", title="처리 종료";
    if(held) {state="review";title="이전 편집 확인 필요";}
    else if(serverState?.state==="error" || (fresh && $("status").classList.contains("err"))) {state="error";title="처리 실패";}
    else if(serverState?.state==="cancelled" || (chip!==baseChip && chip==="중지됨")) {state="cancelled";title="중지됨";}
    else if(fresh && $("status").classList.contains("ok")) {state="done";title="처리 완료";}
    else if(serverState?.state==="done" || (lineFresh && / · done · /.test(text("job-line")))) {state="done";title="서버 처리 완료";}
    $("h-activity").dataset.state=state;put("h-activity-state",title);
    terminalResolved=["error","cancelled"].includes(state) || (fresh && $("status").classList.contains("ok"));
    $("h-activity-dismiss").hidden=false;$("h-activity-result").hidden=false;
    const error=serverState?.error || (fresh ? status : "");
    $("h-activity-result").classList.toggle("h-actionable-error",state==="error");
    put("h-activity-result",held ? '입력값 생성 후 이전 편집 적용을 보류했습니다. 화면의 「이전 편집 확인」에서 계속할 수 있습니다.' : state==="error" ? error.replace(/관경 표기 충돌[\s\S]*/,"관경이 충돌합니다. 속성표에서 확인하거나 역산 검토안을 선택해 출력하세요.") : "");
    put("h-activity-channel","처리 내역 보존");put("h-activity-freshness","");put("h-activity-measure","");
    put("h-activity-lock-note","편집 가능");
    $("h-activity-details").open=false;
  }
  function update() {
    frame=0;
    if(busy()) {
      if(!running) begin();
      connectStream();markLocks();
      const phase=text("busy-text"), log=text("log"), line=text("job-line");
      if(phase!==previousBusy) {
        // Heartbeat elapsed times are not new processing stages.
        const normalized=phase.replace(/ · [\d.]+s/g,"");
        const before=previousBusy.replace(/ · [\d.]+s/g,"");
        if(normalized!==before) {addEvent(normalized);lastReport=performance.now();}
        previousBusy=phase;
      }
      if(log!==previousLog) {
        previousLog=log;logFresh=true;lastReport=performance.now();
        const rows=log.split(/\r?\n/).filter(s=>s.trim() && !["…","—"].includes(s.trim()));
        const latest=rows.at(-1);
        if(latest) {put("h-activity-latest",latest);addEvent(latest);}
      }
      if(line!==previousLine) {previousLine=line;lineFresh=true;lastReport=performance.now();}
      const queued=serverState?.queued || (lineFresh && line.includes("다른 작업이 끝나기를 기다리는 중"));
      put("h-activity-state",$("busy-cancel").disabled ? "중지 요청 처리 중" : queued ? "순서 대기 중" : "처리 중");
      $("h-activity-log").hidden=!logFresh && !lineFresh;
      tick();
    } else if(running) {
      running=false;ended=performance.now();clearInterval(timer);timer=null;
      document.body.classList.remove("h-processing");unlock();tick();
      window.ModuleHProgress?.stop();
      terminalTimer=setTimeout(terminal,120);
    } else if(started && !$("h-activity").hidden) terminal();
  }
  function queue() {if(!frame) frame=requestAnimationFrame(update);}
  function init() {
    if(initialized) return;
    for(const id of ["h-activity","busy","busy-text","busy-bar","busy-cancel","panel-log","log","job-line"])
      if(!$(id)) throw Error("H 작업 현황 구성 누락: "+id);
    $("h-activity-current").append($("busy"));
    $("h-activity-log").append($("panel-log"));
    $("busy-cancel").textContent="중지";
    $("busy").querySelector(".lbl").textContent="현재 처리 단계";
    const track=$("busy-bar").parentElement;
    track.setAttribute("role","progressbar");track.setAttribute("aria-label","현재 처리 구간 진행");
    track.setAttribute("aria-valuemin","0");track.setAttribute("aria-valuemax","100");
    $("h-activity-toggle").onclick=()=>{
      const collapsed=$("h-activity").classList.toggle("h-minimized");
      $("h-activity-toggle").setAttribute("aria-expanded",String(!collapsed));
      put("h-activity-toggle",collapsed ? "펼치기" : "접기");
    };
    $("h-activity-dismiss").onclick=()=>{if(!busy()) $("h-activity").hidden=true;};
    const observer=new MutationObserver(records=>{
      if(running && records.some(r=>r.target===$("status") || $("status").contains(r.target))) statusFresh=true;
      queue();
    });
    for(const id of ["busy","busy-text","busy-bar","log","job-line","job-chip","status"])
      observer.observe($(id),{childList:true,subtree:true,characterData:true,attributes:true,attributeFilter:["class","style","disabled"]});
    observer.observe($("busy-cancel"),{attributes:true,attributeFilter:["disabled"]});
    // New inspector controls can arrive while work is running.
    for(const id of ["dg-ins-body","mg-ins-body","h-pane-task"])
      if($(id)) observer.observe($(id),{childList:true,subtree:true});
    const sizeObserver=new ResizeObserver(()=>{
      const panel=$("h-activity"), shown=!panel.hidden;
      document.body.classList.toggle("h-activity-visible",shown);
      document.body.style.setProperty("--h-activity-height",shown ? panel.offsetHeight+"px" : "0px");
    });
    sizeObserver.observe($("h-activity"));
    window.addEventListener("pagehide",()=>{clearInterval(timer);clearTimeout(terminalTimer);observer.disconnect();sizeObserver.disconnect();});
    initialized=true;document.body.classList.add("h-activity-ready");queue();
  }
  window.ModuleHActivity={init};
})();
