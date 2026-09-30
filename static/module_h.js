/* H owns only presentation. Engine controls, models and validations stay in F. */
(() => {
  "use strict";
  const $ = id => document.getElementById(id);
  const nativeLabels = {open:"도면 열기",pick:"찍기",edit:"손질",design:"수리계산",conv:"수리계산 입력 변환",sub:"경로 추출",auto:"자동 추출"};
  const labels = {...nativeLabels,open:"도면 업로드",pick:"객체 정의",edit:"배관 네트워크 정의",design:"속성 정의"};
  let frame=0, lastStage=null, lastSlot=null, pane=null, returnFocus=null, fullColumns=false;
  let backScope=null, observedStage=null, mergeReturn=null, returning=false;
  let drawerLayout='standard';const drawerGeometry=new Map();
  const required=["steps","slots","stage","busy","panel-open","panel-pick","panel-edit","panel-design","panel-merge","dg-grid","mg-sizing","h-drawer","h-primary"];
  const nativeBusy=()=>!$("busy").classList.contains("hidden");
  const active=id=>!!$(id) && !$(id).classList.contains("hidden") && !$(id).hidden;
  const txt=id=>$(id)?.textContent.trim() || "";
  const write=(id,value)=>{if(txt(id)!==value) $(id).textContent=value;};
  const field=id=>{if(!$(id)) throw Error("필수 컨트롤 누락: "+id);return $(id).closest("label") || $(id);};
  function move(id,...nodes) {nodes.filter(Boolean).forEach(n=>$(id).append(n));}
  function fold(nodes,title) {
    nodes=nodes.filter(Boolean); if(!nodes.length) return;
    const d=document.createElement("details"), s=document.createElement("summary");
    d.className="h-details"; s.textContent=title; d.append(s);
    nodes[0].before(d); nodes.forEach(n=>d.append(n)); return d;
  }
  function foldSection(panel,heading,title) {
    const start=[...$(panel).children].find(n=>n.tagName==="H2" && n.textContent.trim()===heading);
    if(!start) throw Error("설정 구역 누락: "+heading);
    const nodes=[];let n=start.nextElementSibling;
    while(n && n.tagName!=="H2") {nodes.push(n);n=n.nextElementSibling;}
    fold(nodes,title); start.remove();
  }
  function stageNode(stage) {return [...$("steps").children].find(n=>n.dataset.hStage===stage || n.textContent===nativeLabels[stage]);}
  function invoke(id) {
    const n=$(id);
    if(n && !n.disabled && !nativeBusy()) n.click();
  }
  function navigate(stage) {
    const n=stage==="merge" ? $("btn-merge") : stageNode(stage);
    if(n && !n.disabled && typeof n.onclick==="function" && !nativeBusy()) n.click();
  }
  // Process navigation is not graph undo. Use the native stage entry points,
  // retaining their guards and all current drawing/edit/calculation data.
  function observeProcess() {
    const s=window.__mf, key=JSON.stringify([s.sid,s.slot,s.method]);
    if(key!==backScope) {backScope=key;observedStage=s.stage;mergeReturn=null;return;}
    if(s.stage!==observedStage) {
      if(s.stage==='merge')mergeReturn=observedStage;
      observedStage=s.stage;
    }
  }
  function previousProcess() {
    const current=window.__mf.stage;
    if(current==='open')return null;
    const steps=[...$('steps').children].map(node=>({node,
      stage:node.dataset.hStage || Object.keys(nativeLabels).find(k=>nativeLabels[k]===node.textContent)}));
    const available=step=>step && !step.node.disabled && typeof step.node.onclick==='function';
    if(current==='merge') {
      const entry=steps.find(step=>step.stage===mergeReturn);
      return available(entry)?entry:[...steps].reverse().find(available) || null;
    }
    const index=steps.findIndex(step=>step.stage===current);
    return steps.slice(0,index<0?steps.length:index).reverse().find(available) || null;
  }
  async function returnProcess() {
    observeProcess();
    const target=previousProcess();
    if(returning || nativeBusy() || !target || document.querySelector('dialog[open], #conv-modal:not(.hidden)'))return;
    returning=true;closeDrawer(false);sync();
    try {await target.node.onclick();}
    finally {returning=false;sync();$('h-back').focus();}
  }
  function tableMode(on) {
    document.body.classList.toggle("mf-table-open",on);
    if(on) decorateTable();
    window.dispatchEvent(new Event("module-f-table-visibility"));
  }
  function closeDrawer(restore=true) {
    if($("h-drawer").hidden) return;
    $("h-drawer").hidden=true; pane=null; tableMode(false);
    delete $('h-drawer').dataset.pane;
    document.body.classList.remove("h-drawer-open");
    if(restore && returnFocus?.isConnected && returnFocus.getClientRects().length) returnFocus.focus();
    queue();
  }
  function openDrawer(next,title,trigger) {
    if(trigger && !trigger.closest("#h-drawer")) returnFocus=trigger;
    pane=next;
    $('h-drawer').dataset.pane=next;
    const drawer=$('h-drawer'),layout=next==='task'&&window.__mf.stage==='open'?'upload':'standard';
    if(layout!==drawerLayout){
      drawerGeometry.set(drawerLayout,{style:drawer.getAttribute('style')||'',dragged:drawer.dataset.hDragged});
      const saved=drawerGeometry.get(layout);drawer.setAttribute('style',saved?.style||'');
      if(saved?.dragged){drawer.dataset.hDragged=drawer.dataset.dragged='true';}
      else{delete drawer.dataset.hDragged;delete drawer.dataset.dragged;}
      drawerLayout=layout;
    }
    drawer.classList.toggle('h-upload-compact',layout==='upload');
    document.querySelectorAll(".h-pane").forEach(n=>n.hidden=n.id!=="h-pane-"+next);
    $("h-drawer-title").textContent=title;
    $("h-drawer").classList.toggle("h-wide",next==="table");
    $("h-drawer").hidden=false;
    document.body.classList.add("h-drawer-open");
    tableMode(next==="table");
    $("h-drawer-close").focus(); queue();
  }
  function task(trigger) {openDrawer("task","현재 작업 · "+(labels[window.__mf.stage] || "통합 연결"),trigger);}
  function reveal(id,trigger) {
    task(trigger);
    const n=$(id); if(!n) return;
    for(let p=n.parentElement;p && p!==$("h-drawer");p=p.parentElement) {
      if(p.tagName==="DETAILS") p.open=true;
      if(p.classList.contains("foldbody")) p.classList.remove("hidden");
    }
    n.scrollIntoView({block:"center"});
  }
  // Compare the form with the engine's last returned settings, without writing either.
  // A changed form is not a new calculation and must not look ready for export.
  function unappliedSettings() {
    const s=window.__mf, saved=s.design?.settings;
    if(!saved || !s.design?.tables || s.method==="auto") return false;
    const current={schedule:$("dg-sched").value,fx_profile:$("dg-fx").value,fill_short:$("dg-fill").checked};
    if(saved.selection_mode==='area_all')current.fill_short=false;
    if((s.edit?.network_mode || "tree")!=="tree") {
      current.review_datum_m=Number($("dg-review-datum").value || 0);
    }
    return Object.entries(current).some(([k,v])=>k in saved && saved[k]!==v);
  }
  function notice() {
    const s=window.__mf, st=s.stage;
    if(window.ModuleHAttributes?.historyConflict())return {id:'ha-history-conflict',text:'이전 편집을 보존하고 새 기준망 사용 여부를 확인하세요'};
    let ids=[];
    if(st==="edit") ids=["ed-blocked","ed-rankbad","ed-worst-err"];
    if(st==="design" || st==="conv") ids=["dg-stale"];
    if(st==="merge") ids=["mg-stale","mg-split"];
    const id=ids.find(id=>active(id) && txt(id));
    if(id) return {id,text:txt(id).replace(/\s+/g," ")};
    if(["design","conv"].includes(st) && unappliedSettings())
      return {id:"dg-build-row",text:"입력 조건이 바뀌었습니다 · 표를 다시 확정해야 출력에 반영됩니다"};
    if(["design","conv"].includes(st) && /[1-9]\d*건/.test(txt("dg-issues-n")))
      return {id:"dg-issues-body",text:"검토할 항목 "+txt("dg-issues-n")+" · 위치와 판정 근거 확인"};
    if(["design","conv"].includes(st) && (s.edit?.network_mode || 'tree')!=='tree')
      return {id:'dg-review-notice',text:'미확정'};
    if($("status")?.classList.contains("err") || $("status")?.classList.contains("warn"))
      return {id:null,text:txt("status")};
    return null;
  }
  // Suggestions are read-only. One user gesture calls at most one native action.
  // Settings are shown before executing actions with calculation-affecting inputs.
  function suggestion() {
    const s=window.__mf,e=s.edit || {}, st=s.stage;
    const action=(id,label,status,settings=false)=>({id,label,status,settings});
    if(st==="open") return {pane:"task",label:!s.sid && !window.ModuleHAccess?.selected ? "Access" : "도면 업로드",status:"파일을 선택하면 자동으로 업로드합니다"};
    if(st==="pick") return action("pk-next","배관망 준비","처리 영역과 배관·헤드 분류를 확인하세요",true);
    if(st==="edit") {
      if(!e.sources?.length) return {pane:"task",focus:"panel-edit",label:"연결점 지정하기",status:"알람밸브 연결점을 지정하세요"};
      if(!s.zones?.length) return {pane:"task",focus:"ed-zone-arm",label:"영역 지정",status:"사각형 또는 펜으로 영역을 지정하세요"};
      if(e.worst?.selection_mode!=="area_all" || e.edits_since_worst>0)
        return action("ed-worst","영역 배관망 추출","영역 안의 모든 헤드를 연결합니다",true);
      return action("ed-next","입력값 검토","선정한 배관망의 입력값을 검토할 수 있습니다");
    }
    if(st==="sub") return action("sub-extract","경로 추출","연결점과 표고 조건을 확인하세요",true);
    if(st==="auto") return s.autoDone
      ? action("au-handoff","배관 네트워크 정의로 이동","기존 인식 결과를 이어받습니다",true)
      : {pane:"task",label:"도면 다시 업로드",status:"영역 전체 추출은 도면 업로드에서 시작하세요"};
    if(st==="design") {
      if(window.ModuleHAttributes?.historyConflict())return {recovery:true,label:'편집 기록 확인',status:'이전 편집 적용 보류 · 확인 후 계속'};
      if(active("dg-stale") || !s.design?.tables || s.ovDirty || unappliedSettings())
        return action(active("dg-build-row") ? "dg-build" : "dg-back-auto","입력값 반영","입력 조건을 확인하고 표를 확정하세요",true);
      return action("dg-to-conv","출력 준비","");
    }
    if(st==="conv") {
      if(active("dg-stale") || s.ovDirty || unappliedSettings()) return action("btn-back-design","입력값 다시 확인","이전 표가 표시되고 있습니다 · 변경 내용을 먼저 반영하세요");
      if(!$("btn-download-design").disabled) return action("btn-download-design","SDF 내려받기","생성된 SDF와 SLF를 함께 내려받습니다");
      return action("btn-convert","SDF 생성","출력 범위와 형식을 확인하세요",true);
    }
    if(st==="merge") {
      if(active("mg-stale") || !s.merge?.summary?.merged) return action("mg-build","통합 연결","도면 연결과 급수방식을 확인하세요",true);
      if($('h-output-basis')?.value==='proposal') return !$('sz-download').disabled
        ? action('sz-download','역산 검토안 내려받기','')
        : {...action('mg-emit','역산 검토안 SDF 생성','',true),settingsPane:'export'};
      if(!$("mg-dl-sdf").disabled) return action("mg-dl-sdf","SDF 내려받기","원본 결합망의 생성된 파일을 내려받습니다");
      return {...action("mg-emit","SDF 생성","원본 결합망을 출력합니다 · 역산 검토안은 별도 출력",true),settingsPane:"export"};
    }
    return {pane:"task",label:"현재 작업 확인",status:"현재 작업을 확인하세요"};
  }
  function runPrimary() {
    if(nativeBusy()) return;
    if(window.__mf.stage==="open" && !window.__mf.sid && !window.ModuleHAccess?.selected) {window.ModuleHAccess?.open();return;}
    const a=suggestion();
    if(a.recovery){closeDrawer(false);window.ModuleHAttributes.showHistoryConflict();(window.ModuleHAttributes.historyConflict().basis_stale?$('ha-rebuild-basis'):$('ha-accept-basis')).focus();return;}
    if(window.__mf.stage==="auto" && !window.__mf.autoDone) {navigate("open");task($("h-primary"));return;}
    if(a.pane) {
      task($("h-primary"));
      if(a.focus) reveal(a.focus,$("h-primary"));
      if(a.focus==="ed-zone-arm") $("ed-zone-arm").checked=true;
      return;
    }
    if(a.settings && pane!==(a.settingsPane || "task")) {
      if(a.settingsPane==="export") openDrawer("export","출력 설정 · 다운로드",$("h-primary"));
      else task($("h-primary"));
      return;
    }
    invoke(a.id);
  }
  const basic = {
    pipes:["label","in","out","dia","length","type","dia_src"],
    nodes:["label","elevation","io_node"],
    nozzles:["label","in","out","type","lib","status","flow_lmin","pressure_pa"],
    fittings:["pipe","type","count","eq_len"],
    equipment:["label","pipe","type","count","eq_len","desc"],
  };
  const columns = {label:"번호",in:"시작 노드",out:"끝 노드",dia:"호칭경 (mm)",length:"길이 (m)",type:"종류 / 규격",lib:"헤드 규격",status:"상태",dia_src:"관경 근거",elevation:"표고 (m)",io_node:"입출력",c:"C값",elev:"표고차 (m)",eq_len:"등가길이 (m)",flow_lmin:"유량 (L/min)",pressure_pa:"압력 (Pa)",count:"개수",pipe:"배관",flow_direction:"유향 기준",group:"그룹",desc:"설명"};
  const sources = {text:"도면 근거",nfpc_min:"규약 상향 · 검토",nfpc_fallback:"트리 기준개수",review_default:"검토용 · 미확정",user:"사용자 수정",unresolved:"미지정"};
  const numeric = new Set(["dia","length","c","elev","elevation","eq_len","count","flow_lmin","pressure_pa","flow_m3s","x","y","z"]);
  const numberFormat = new Intl.NumberFormat('ko-KR', {maximumFractionDigits:3});
  function cellText(n,value) {if(n.textContent!==value)n.textContent=value;}
  function decorateTable() {
    const table=$('dg-grid').querySelector('table');if(!table)return;
    const which=$('dg-table').value,rows=window.__mf.design?.tables?.[which] || [],keys=[];
    for(const r of rows)for(const key of Object.keys(r))if(key!=='bore_provenance'&&!keys.includes(key))keys.push(key);
    const headers=[...table.querySelectorAll('thead th')],order=basic[which] || [];
    const indices=[...order.map(k=>keys.indexOf(k)).filter(i=>i>=0),...headers.map((_,i)=>i).filter(i=>!order.includes(keys[i]))];
    const trs=[...table.querySelectorAll('tr')];
    // Read native cells by their stable original index, including repeat decorations.
    for(const tr of trs){const cells=[...tr.children];cells.forEach((cell,i)=>{if(!cell.dataset.hIndex)cell.dataset.hIndex=String(i+1);});
      const at=new Map(cells.map(c=>[Number(c.dataset.hIndex)-1,c]));
      for(const index of indices){const cell=at.get(index);if(!cell)continue;const key=keys[index] || 'override_note';
        cell.dataset.hColumn=key;cell.classList.toggle('h-number',numeric.has(key));cell.hidden=!fullColumns&&!order.includes(key)&&key!=='override_note';
        if(cell.tagName==='TH'){cellText(cell,columns[key] || cell.textContent);cell.scope='col';}
        tr.append(cell);
      }
    }
    table.querySelectorAll('tbody tr').forEach((tr,i)=>{
      const row=rows[i] || {};
      tr.querySelectorAll('td').forEach(cell=>{const key=cell.dataset.hColumn;if(!(key in row))return;let value=row[key];
        if(numeric.has(key)&&typeof value==='number'&&Number.isFinite(value)){
          cell.title=`원래 값: ${value}`;
          if(!fullColumns)value=numberFormat.format(value);
        }
        if(key==='dia_src')value=sources[value] || value;
        if(key==='dia' && !row.dia && row.bore_provenance?.policy==='drawing_first_v1')value='미지정';
        if(key==='flow_direction'&&value==='solver_reference')value='기준 방향 · 유향 미확정';
        if(value===null || value===undefined || value==='')value='—';else if(typeof value==='object')value=JSON.stringify(value);else if(typeof value==='boolean')value=value?'예':'아니오';
        cellText(cell,String(value));
      });
      tr.tabIndex=0;tr.setAttribute('aria-label',`${which==='pipes'?'배관':'항목'} ${row.label ?? row.pipe ?? i+1} 상세`);
      tr.onkeydown=e=>{if(e.key==='Enter'||e.key===' '){e.preventDefault();tr.click();}};
    });
    cellText($('h-row-count'),`${rows.length}행`);cellText($('h-all-columns'),fullColumns?'기본 열 보기':'상세 열 보기');
    $('h-all-columns').setAttribute('aria-pressed',String(fullColumns));
  }
  function sync() {
    const picks=$('sub-picks');
    if(picks){
      let status=$('h-sub-pick-status');
      if(!status){status=document.createElement('div');status.id='h-sub-pick-status';status.setAttribute('role','status');picks.after(status);}
      const s=window.__mf;
      const labels=[$('sub-lab-a').textContent,$('sub-lab-b').textContent];
      const values=labels.map((label,i)=>`${label} · ${s?.sub?.picks?.[i]?'지정됨':'미지정'}`);
      if(s?.sub?.arm!=null)values.push(`${labels[s.sub.arm]} 위치를 도면에서 선택하세요`);
      write('h-sub-pick-status',values.join('　 /　 '));
    }
    frame=0;
    const s=window.__mf, st=s.stage, working=nativeBusy();
    observeProcess();
    const previous=previousProcess();
    $('h-back').disabled=working || returning || !previous;
    $('h-back').title=previous?`이전 공정: ${labels[previous.stage] || previous.stage}`:'이전 공정이 없습니다';
    $("h-region-draw").setAttribute("aria-pressed",String($("ed-zone-arm").checked));
    $("h-region-draw").textContent=$("ed-zone-arm").checked?"영역 그리기 종료":"영역 그리기";
    document.body.classList.toggle("h-table-review",st==="conv");
    if(st!==lastStage || s.slot!==lastSlot) {
      closeDrawer(false);lastStage=st;lastSlot=s.slot;
      document.querySelector(".side").scrollTop=0;
    }
    const a=suggestion(), n=notice();
    write("h-action-title",working ? "처리 중" : st==="design" && s.design?.tables ? "입력표 준비" : labels[st] || "통합 배관망");
    write("h-action-help",working ? "확대·이동·입력표 확인 가능 · 현재 도면 수정은 처리 후 가능합니다."
      : a.settings && pane!==(a.settingsPane || "task") ? "누르면 먼저 필요한 설정이 열립니다." : "도면·아이소에서 선택하여 확인 · 상세 옵션은 더보기");
    write("h-primary",a.label);
    $("h-primary").disabled=working || !!(a.id && $(a.id)?.disabled) || (a.id==='mg-build' && !window.ModuleHAccess?.canMerge);
    $("h-primary").dataset.nativeAction=a.id || "";
    $("h-welcome").hidden=working || !!s.world || st!=="open" || !!window.ModuleHAccess?.engaged;
    const slotLabel=$("slots").querySelector("button.on");
    write("h-drawing-name",slotLabel?.textContent.trim() || (s.sid ? "현재 도면" : "도면 없음"));
    $("h-alert").hidden=!n;
    const legend=$("bore-overlay");
    $("h-alert").style.top=(legend?.getClientRects().length ? legend.offsetTop+legend.offsetHeight+8 : 12)+"px";
    if(n) {
      const caption=n.id?.includes("stale") || n.id==="dg-build-row" ? "변경사항 반영 필요"
        : n.id==="dg-review-notice" || (s.edit?.network_mode!=="tree" && n.id==="mg-summary") ? "검토용 · 미확정"
        : n.id==="dg-issues-body" ? "검토 "+txt("dg-issues-n") : "확인 필요";
      write("h-alert-open",caption+" ›");
    }
    $("h-alert-open").title=n?.text || "";
    for(const el of $("steps").children) {
      const key=el.dataset.hStage || Object.keys(nativeLabels).find(key=>nativeLabels[key]===el.textContent);
      if(key) {el.dataset.hStage=key;if(el.textContent!==labels[key])el.textContent=labels[key];}
      el.setAttribute("role","button");el.tabIndex=el.onclick ? 0 : -1;
      el.setAttribute("aria-disabled",String(!el.onclick || working));
      el.onkeydown=event=>{if(["Enter"," "].includes(event.key) && el.onclick && !nativeBusy()){event.preventDefault();el.click();}};
    }
    $("h-merge").disabled=working || !window.ModuleHAccess?.canMerge;
    $("h-table").disabled=st==="merge" ? !s.merge?.summary?.merged : !["design","conv"].includes(st) || !s.design?.tables;
    $("h-view").disabled=!["edit","design","merge"].includes(st);
    $("h-sizing").disabled=working || st!=="merge";
    $("h-export").disabled=working || !["design","conv","merge"].includes(st);
    $("h-save").disabled=working || st!=="edit" || $("ed-save").disabled;
    $("h-review").disabled=working || !["edit","design","conv","merge"].includes(st);
    $("h-open").disabled=working;
    $("dxf").disabled=working;
    write("h-history-status",txt("status") || "아직 처리 내역이 없습니다.");
    write("h-drawer-status",$("status").classList.contains("err") ? "오류 · 처리 내역 확인" : $("status").classList.contains("warn") ? "확인 필요 · 처리 내역 확인" : "");
    $("h-drawer-status").title=txt("status");
    window.ModuleHIntegration?.sync();
  }
  function queue() {if(!frame) frame=requestAnimationFrame(sync);}
  function layout() {
    const header=document.querySelector("header"), side=document.querySelector(".side");
    const back=document.createElement('button');back.id='h-back';back.type='button';
    back.textContent='돌아가기';
    $('h-history').before(back);
    // Move the native control, preserving its camera handler and busy-safe ID.
    back.before($('btn-fit'));
    $('btn-fit').after($('h-source-view'));
    $('h-source-view').onclick=()=>window.ModuleHSourceView?.restore();
    // Retain engine/status anchors without exposing the coordinate/status strip.
    $('coord').hidden=true;
    $('status').closest('.bar').hidden=true;
    header.querySelector(".back").before($("h-header-tools"));
    move("h-history-steps",$("steps"),$("btn-merge"));
    move("h-slots",$("slots")); move("h-pane-task",side);
    move("stage",$("h-welcome"),$("h-alert"));
    // Keep dormant native IDs for compatibility; they cannot supply any H input.
    for(const id of ["ed-k-preset","ed-k","au-k-preset","au-k","dg-fill"]) {
      field(id).hidden=true;$(id).disabled=true;
    }
    for(const id of ["ed-k-why","au-k-why","ed-sheet-wrap","ed-sheets","adv-auto","au-run"])
      $(id).classList.add("h-obsolete-selection");
    const regionHeading=[...$("panel-edit").querySelectorAll("h2")].find(n=>n.textContent.includes("최불리 선정"));
    if(regionHeading) regionHeading.textContent="영역 지정";
    const areaHeading=[...$("panel-edit").querySelectorAll("h2")].find(n=>n.textContent.includes("② 영역 지정"));
    if(areaHeading) areaHeading.remove();
    field("ed-zone-arm").hidden=true;
    $("ed-zone-arm").checked=false;
    const regionButton=document.createElement("button");
    regionButton.id="h-region-draw";regionButton.type="button";
    regionButton.textContent="영역 그리기";regionButton.setAttribute("aria-pressed","false");
    field("ed-zone-arm").after(regionButton);
    regionButton.onclick=()=>{
      $("ed-zone-arm").checked=!$("ed-zone-arm").checked;
      $("ed-zone-arm").dispatchEvent(new Event("change",{bubbles:true}));
      queue();
    };
    document.querySelector('.emode[data-mode="원클릭"]').addEventListener("click",()=>{
      $("ed-zone-arm").checked=false;queue();
    });
    // The ordinary extraction action already handles changed selections. Keep
    // the shared engine anchors, but remove the redundant re-extraction row in H.
    $("ed-recalc-row").classList.add("h-obsolete-selection");
    $("ed-src").title="급수원 선택";
    $("ed-worst-view").querySelector('[value="only"]').textContent="영역 배관망만 보기";
    if(regionHeading) {
      const group=document.createElement("section");group.className="h-area-selection";
      let node=regionHeading;
      while(node) {
        const last=node===$("ed-recalc-row"),next=node.nextElementSibling;
        group.append(node);if(last)break;node=next;
      }
      field("ed-network-mode").after(group);
    }
    fold([$("ed-flow").parentElement,$("ed-flow-summary")],"전체 배관 연결 진단");
    ["h-header-tools","h-actionbar"].forEach(id=>$(id).hidden=false);
    $('h-actionbar').setAttribute('aria-label','작업 도구');
    // The canvas reaches the bottom edge. Only the compact action cluster floats;
    // default docks clear its actual height even when buttons wrap on small screens.
    new ResizeObserver(()=>document.body.style.setProperty('--h-actions-height',
      $('h-actionbar').getBoundingClientRect().height+'px')).observe($('h-actionbar'));
    foldSection("panel-edit","고른 헤드 종류","선택한 헤드 종류 바꾸기");
    foldSection("panel-edit","자동 이음","끊긴 배관 자동 이음");
    $("adv-body").classList.add("hidden");
    ["ed-worst-view","ed-bg","ed-gridline","ed-gridline-note"].forEach(id=>move("h-pane-edit-view",field(id)));
    ["dg-plan","dg-plan-view-row","dg-under","dg-gridline","dg-gridline-note","dg-iso","dg-iso-note","dg-zscale","dg-canvas","dg-ref","dg-stub"]
      .forEach(id=>move("h-pane-design-view",field(id)));
    fold([field("dg-fx"),field("dg-fill")],"계산 고급 설정 · 신축배관·헤드 채움");
    const planLabel=field("dg-plan").querySelector("span");
    if(planLabel) planLabel.textContent="평면에서 보기 (끄면 아이소)";
    const viewHeading=[...$("panel-design").children].find(n=>n.tagName==="H2" && n.textContent.trim()==="보기");
    if(viewHeading) viewHeading.textContent="계산 설정";
    ["mg-iso","mg-grid","mg-under"].forEach(id=>move("h-pane-merge-view",field(id)));
    move("h-pane-merge-view",document.querySelector(".mg-under-options"),$("mg-under-note"),$("mg-grid-note"));
    $("mg-sizing").open=true;move("h-pane-sizing",$("mg-sizing"));
    fold([field("sz-suction"),field("sz-vhead"),field("sz-vhead").nextElementSibling,field("sz-dn"),field("sz-iterations")],"펌프 상세 · 반복 한계");
    move("h-export-original",$("mg-emit").parentElement);
    const outputBasis=document.createElement('label');outputBasis.className='f';
    outputBasis.innerHTML='<span>출력 대상</span><select id="h-output-basis"><option value="original">정의한 원본 배관망</option><option value="proposal" disabled>역산 검토안</option></select>';
    $('h-export-original').prepend(outputBasis);
    let lastProposal=null;
    window.addEventListener('module-h-sizing-result',event=>{
      const d=event.detail,select=$('h-output-basis');select.options[1].disabled=!d.feasible||d.stale;
      if(d.proposal && d.proposal!==lastProposal && d.feasible && !d.stale)select.value='proposal';
      lastProposal=d.proposal;queue();
    });
    window.addEventListener('module-h-sizing-dirty',()=>{$('h-output-basis').options[1].disabled=true;queue();});
    const planNotice=document.createElement('p');planNotice.className='h-plan-export-notice';
    planNotice.textContent='평면도 검토용 SDF · 미지정 관경과 실제 펌프·수원 조건을 정의하기 전에는 PIPENET 계산 불가';
    $('panel-conv').prepend(planNotice);
    move("h-export-pipenet",field("mg-dl-coord"),$("mg-dl-sdf"),$("mg-files"));
    ["mg-dl-kfp","mg-dl-kfp-edit","mg-dl-has","mg-dl-slf"].forEach(id=>move("h-export-other",$(id)));
    move("h-table-controls",$("dg-table").parentElement);
    move("h-table-data",$("dg-grid"));
    ["dg-bore-row","dg-bore-why","dg-bore-list"].forEach(id=>move("h-table-edit",$(id)));
    const tableHeading=[...$("panel-design").children].find(n=>n.tagName==="H2" && n.textContent.trim()==="표");
    if(tableHeading) tableHeading.remove();
    // H removes static prose, not validation results, units or calculation inputs.
    document.querySelectorAll(".h-pane p.hint:not([id])").forEach(p=>{
      if(!p.querySelector("button,input,select,textarea,a"))p.classList.add("h-prose");
    });
    for(const id of ["ed-network-note","dg-review-notice","open-zoom-window","btn-zoom-window","mg-summary"])
      $(id)?.classList.add("h-removed-copy");
    // H has one evidence-display switch, shared by plan and isometric views.
    field('dg-bore-color').hidden=true;$('dg-bore-legend').hidden=true;
    $("ed-company-box").querySelector("summary").textContent="인식 결과";
    document.querySelector('.emode[data-mode="원클릭"]').textContent="시작노드 정의";
    for(const [id,label] of Object.entries({"dg-build":"배관 속성 확정","ed-save":"배관망 저장","dg-back":"← 배관 네트워크 정의","au-handoff":"배관 네트워크 정의 →"}))write(id,label);
    const editHeading=[...$("panel-edit").querySelectorAll("h2")].find(n=>n.textContent==="손질 모드");
    if(editHeading)editHeading.textContent="배관 네트워크 정의";
    // Rebuild remains explicitly accessible in the drawer even after a successful run.
    for(const id of ["btn-open","pk-next","ed-next","sub-extract","au-run","btn-convert"])
      $(id).classList.add("h-primary-source");
    // H has no separate upload-confirmation step. Keep the shared handler/ID
    // for F compatibility, but remove its entire empty row from the H layout.
    $("btn-open").closest('.row').hidden=true;
    $("dxf").setAttribute('aria-description','파일을 선택하면 자동으로 업로드하고 다음 작업으로 이동합니다.');
    document.body.classList.add("h-ready");window.dispatchEvent(new Event("resize"));
    window.ModuleHActivity.init();
  }
  try {
    required.forEach(id=>{if(!$(id)) throw Error("필수 컨트롤 누락: "+id);});
    if(!window.__mf) throw Error("공용 배관망 엔진을 불러오지 못했습니다.");
    layout();
    $("dxf").addEventListener('change',()=>{
      if(!$("dxf").files?.length || nativeBusy() || window.__mf.stage!=="open")return;
      closeDrawer(false);
      // The native handler captures the File synchronously and retains its
      // upload progress, cancellation, validation and post-read stage routing.
      invoke('btn-open');
      // Permit choosing the same file again after a failure or cancellation.
      // This never retries automatically, nor alters the captured File object.
      $("dxf").value='';queue();
    });
    $('cv-full-kfp').checked=false;$('cv-worst-kfp').checked=false;$('cv-worst-sdf').checked=true;
    $("h-primary").onclick=runPrimary;
    $('h-back').onclick=returnProcess;
    // Process navigation is button-only. Let the shared engine undo the last
    // edit in the current stage (and redo via Ctrl+Shift+Z), not change stages.
    // Keep H's repeat/modal guards without intercepting native text undo.
    window.addEventListener('keydown',e=>{
      if(!(e.ctrlKey||e.metaKey) || e.altKey || e.key.toLowerCase()!=='z')return;
      const focus=document.activeElement;
      if(/^(INPUT|TEXTAREA|SELECT)$/.test(focus?.tagName || '') || focus?.isContentEditable)return;
      if(e.repeat || document.querySelector('dialog[open], #conv-modal:not(.hidden)')) {
        e.preventDefault();e.stopImmediatePropagation();
      }
    },true);
    window.ModuleHUI={task,closeDrawer,queue,unappliedSettings};
    $("h-open").onclick=()=>window.ModuleHAccess?.open();
    $("h-welcome-open").onclick=()=>window.ModuleHAccess?.open();
    $("h-new-drawing").onclick=()=>{navigate("open");requestAnimationFrame(()=>task($("h-open")));};
    $("h-more").onclick=()=>openDrawer("more","더보기",$("h-more"));
    $("h-history").onclick=()=>openDrawer("history","처리 내역",$("h-history"));
    $("h-task").onclick=()=>task();
    $("h-drawings").onclick=()=>window.ModuleHAccess?.open();
    $("h-merge").onclick=()=>window.ModuleHAccess?.merge();
    $("h-save").onclick=()=>invoke("ed-save");
    $("h-view").onclick=()=>openDrawer(window.__mf.stage+"-view","보기 설정");
    $("h-table").onclick=()=>["design","merge"].includes(window.__mf.stage) ? (closeDrawer(false),window.ModuleHProperties?.open()) : openDrawer("table","수리계산 입력표");
    $("h-all-columns").onclick=()=>{fullColumns=!fullColumns;decorateTable();};
    $("h-sizing").onclick=()=>openDrawer("sizing","관경·펌프 역산 검토");
    $("h-export").onclick=()=>{
      if(window.__mf.stage==="merge") openDrawer("export","출력 설정 · 다운로드");
      else if(window.__mf.stage==="design") navigate("conv");
      else task();
    };
    const review=()=>{
      if(window.__mf.stage==="conv") {navigate("design");return;}
      if(window.ModuleHAttributes?.historyConflict()){runPrimary();return;}
      const n=notice();if(n?.id && $(n.id)) reveal(n.id);else task();
    };
    $("h-alert-open").onclick=review;$("h-review").onclick=review;
    $("h-drawer-close").onclick=()=>closeDrawer();
    window.addEventListener("keydown",e=>{if(e.key==="Escape") closeDrawer();});
    document.addEventListener("change",queue);
    document.addEventListener("click",queue);
    const observer=new MutationObserver(queue);
    observer.observe($("steps"),{childList:true});
    new MutationObserver(()=>{if(pane==="table") decorateTable();}).observe($("dg-grid"),{childList:true});
    // Observe native DOM only; H text writes cannot create an observer loop.
    for(const id of ["busy","slots","status","bore-overlay","h-pane-task","h-pane-export","dg-review-notice","dg-stale","mg-stale"])
      observer.observe($(id),{childList:true,subtree:true,characterData:true,attributes:true,attributeFilter:["class","disabled"]});
    sync();
    // Presentation-only add-on: static delivery keeps running sessions intact.
    const integrationStyle=document.createElement('link');
    integrationStyle.rel='stylesheet';integrationStyle.href='/static/module_h_integration.css?v=20260930-1';
    document.head.append(integrationStyle);
    const integrationScript=document.createElement('script');
    integrationScript.src='/static/module_h_integration.js?v=20260930-1';
    document.body.append(integrationScript);
  } catch(error) {
    const note=document.createElement("div");note.id="h-layout-error";note.setAttribute("role","alert");
    note.textContent="H 화면 구성에 문제가 있습니다. 새로고침하거나 초기화면에서 Module F를 이용해주세요. "+error.message;
    document.querySelector("header").after(note);console.error(error);
  }
})();
