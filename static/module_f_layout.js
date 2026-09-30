/* Module F presentation only: native controls/events, five stages and models stay authoritative. */
(() => {
  "use strict";
  if (document.body.classList.contains("module-h")) return;
  const $ = id => document.getElementById(id);
  if (!window.__mf || !$('panel-design')) return;
  const settings = new Map(), notices = [];
  let tableOpen = false, fullColumns = false, pending = 0, previousStage = null, previousSlot = null;
  const labels = {open:"도면 열기",pick:"찍기",edit:"손질",design:"수리계산",conv:"수리계산 입력 변환",merge:"통합 · 결합",sub:"경로 추출",auto:"자동 추출"};
  const basic = {
    pipes:["label","in","out","dia","length","type","dia_src"],
    nodes:["label","elevation","io_node"],
    nozzles:["label","in","out","type","lib","status","flow_lmin","pressure_pa"],
    fittings:["pipe","type","count","eq_len"],
    equipment:["label","pipe","type","count","eq_len","desc"],
  };
  const columns = {label:"번호",in:"시작 노드",out:"끝 노드",dia:"호칭경 (mm)",length:"길이 (m)",type:"종류 / 규격",lib:"헤드 규격",status:"상태",dia_src:"관경 근거",elevation:"표고 (m)",io_node:"입출력",c:"C값",elev:"표고차 (m)",eq_len:"등가길이 (m)",flow_lmin:"유량 (L/min)",pressure_pa:"압력 (Pa)",count:"개수",pipe:"배관",flow_direction:"유향 기준",group:"그룹",desc:"설명"};
  const sources = {text:"도면 표기",nfpc_min:"규약 상향 · 검토",nfpc_fallback:"규약 보완",review_default:"검토용 · 미확정",user:"사용자 수정"};
  const numeric = new Set(["dia","length","c","elev","elevation","eq_len","count","flow_lmin","pressure_pa","flow_m3s","x","y","z"]);
  const numberFormat = new Intl.NumberFormat('ko-KR', {maximumFractionDigits:3});
  function el(tag, className, text) { const n=document.createElement(tag); if(className)n.className=className;if(text)n.textContent=text;return n; }
  function btn(id,text,action) { const b=el("button","",text);b.id=id;b.type="button";if(action)b.onclick=action;return b; }
  function field(id) {const n=$(id);return n?.closest("label") || n;}
  function text(n,value) {if(n && n.textContent!==value)n.textContent=value;}
  function busy() {return !$('busy')?.classList.contains('hidden');}
  function invoke(id) {const n=$(id);if(n && !n.disabled && !busy())n.click();}
  function disclosure(title, id) {const d=el('details','mf-disclosure');if(id)d.id=id;d.append(el('summary','',title));return d;}
  function advanced(panelId) {
    if(settings.has(panelId))return settings.get(panelId);
    const panel=$(panelId), d=disclosure('고급 설정',`mf-advanced-${panelId.replace('panel-','')}`);
    panel.append(d);settings.set(panelId,d);return d;
  }
  function move(panelId, ids) {const d=advanced(panelId);for(const id of ids){const n=field(id);if(n)d.append(n);}}
  function section(panelId, title, destination, newTitle=title) {
    const start=[...$(panelId).children].find(n=>n.tagName==='H2' && n.textContent.trim()===title);
    if(!start)return;
    const nodes=[];let n=start.nextElementSibling;
    while(n && n.tagName!=='H2' && !n.classList.contains('mf-disclosure')){nodes.push(n);n=n.nextElementSibling;}
    const d=disclosure(newTitle);start.before(d);d.append(...nodes);start.remove();if(destination)destination.append(d);return d;
  }
  function foldPair(panelId,id) {
    const n=$(id);if(!n)return;
    const h=document.querySelector(`h2[data-fold="${id}"]`);
    if(h)advanced(panelId).append(h);advanced(panelId).append(n);
  }
  function help(panelId) {
    const panel=$(panelId);if(!panel)return;
    const nodes=[...panel.querySelectorAll('p.hint:not([id]),div.hint:not([id]),p.step-p')]
      .filter(n=>!n.closest('details') && !n.classList.contains('quiet-hidden') && !n.querySelector('input,select,button,a'));
    if(!nodes.length)return;
    const d=disclosure('도움말 · 처리 기준');advanced(panelId).append(d);d.append(...nodes);
  }
  function compactNotice(id,label) {
    const n=$(id);if(!n)return;
    const d=disclosure(label);d.classList.add('mf-notice');n.before(d);d.append(n);
    notices.push({node:n,wrapper:d,label});
  }
  function addSummary(panelId,id) {const n=el('div','mf-config-summary');n.id=id;$(panelId).prepend(n);}
  function layout() {
    const side=document.querySelector('.side');
    // Native stage panels remain in place; only their presentation groups move.
    section('panel-edit','고른 헤드 종류',null,'선택한 헤드 종류');
    section('panel-edit','자동 이음',advanced('panel-edit'),'끊긴 배관 자동 이음');
    move('panel-edit',['ed-network-mode','ed-k-preset','ed-k','ed-k-why','ed-flow-report','ed-worst-view','ed-bg','ed-gridline','ed-gridline-note']);
    foldPair('panel-edit','ed-status-body');
    move('panel-pick',['pk-head-profile']);foldPair('panel-pick','pk-auto-body');foldPair('panel-pick','pk-mats-body');
    // Advanced adoption/layers retain their conditional root panels and all IDs.
    $('adv-body').classList.add('hidden');
    const h=$('panel-advanced').querySelector('h2');h.textContent='추출 설정';
    move('panel-design',['dg-plan-view-row','dg-under','dg-gridline','dg-gridline-note','dg-iso','dg-iso-note','dg-zscale','dg-canvas','dg-fx','dg-ref','dg-stub','dg-fill']);
    foldPair('panel-design','dg-summary-body');foldPair('panel-design','dg-audit-body');foldPair('panel-design','dg-evidence-body');
    $('dg-issues-body').classList.add('hidden');
    // Required loop/grid bore and elevation inputs, errors and stale guards stay on the main panel.
    move('panel-conv',['cv-full-kfp','cv-worst-kfp','btn-download-review','btn-download-worst','btn-download-has','btn-download','btn-download-slf','btn-download-worst-edit','btn-download-edit']);
    foldPair('panel-conv','conv-options-body');
    move('panel-merge',['mg-sizing','mg-grid','mg-under','mg-under-note','mg-grid-note','mg-dl-kfp','mg-dl-kfp-edit','mg-dl-has','mg-dl-slf']);
    const under=document.querySelector('.mg-under-options');if(under)advanced('panel-merge').append(under);
    move('panel-sub',['sub-snap']);foldPair('panel-sub','sub-layers-body');
    // General help is opt-in, not deleted. Live error/progress nodes are never bulk hidden.
    for(const id of ['panel-open','panel-pick','panel-edit','panel-design','panel-conv','panel-sub','panel-auto','panel-merge'])help(id);
    compactNotice('ed-network-note','루프·그리드 · 검토 기준');
    compactNotice('dg-review-notice','검토용 출력 · 미확정 유지');
    compactNotice('dg-stale','입력표 갱신 필요 · 표 확정');
    compactNotice('mg-stale','결합 결과 갱신 필요');
    compactNotice('dg-bore-why','관경 수정 · 적용 방법');
    compactNotice('pk-crop-info','처리 영역 · 안내');
    addSummary('panel-edit','mf-edit-config');addSummary('panel-design','mf-design-config');addSummary('panel-conv','mf-export-config');
    // Shorten labels only, never checked values or selected defaults.
    const plan=$('dg-plan').closest('label').querySelector('span');if(plan)plan.textContent='평면에서 보기 (끄면 아이소)';
    const iso=$('mg-iso').closest('label').querySelector('span');if(iso)iso.textContent='아이소로 보기 · 출력 좌표 연동';
    $('dg-build').textContent='표 확정';$('btn-convert').textContent='파일 생성';
    const head=el('div','mf-workbar');head.id='mf-workbar';
    const title=el('strong');title.id='mf-work-title';head.append(title,el('span','grow'));
    head.append(btn('mf-network','배관망',()=>setTable(false)),btn('mf-table','입력표',()=>setTable(true)));
    head.append(btn('mf-settings','고급 설정',()=>{setTable(false);const d=settings.get(`panel-${window.__mf.stage}`);if(d){d.open=!d.open;if(d.open)d.scrollIntoView({block:'nearest'});}}));
    // The canvas and all hit-test overlays must share the same untouched stage
    // rectangle. Keep the new toolbar outside it, not over a resized canvas.
    const shell=el('div','mf-stage-shell');$('stage').before(shell);shell.append(head,$('stage'));
    const pane=el('section','mf-table-pane');pane.id='mf-table-pane';pane.hidden=true;pane.setAttribute('aria-label','전체 폭 수리계산 입력표');
    const tools=el('div','mf-table-tools');tools.append(el('strong','','수리계산 입력표'));
    tools.append($('dg-table').parentElement);
    const count=el('span','mf-row-count');count.id='mf-row-count';tools.append(count,el('span','grow'));
    tools.append(btn('mf-all-columns','상세 열 보기',()=>{fullColumns=!fullColumns;decorateTable();}));
    tools.append(btn('mf-table-build','표 확정',()=>invoke('dg-build')));
    pane.append(tools);
    const warning=el('div','mf-table-warning');warning.id='mf-table-warning';warning.hidden=true;warning.setAttribute('role','status');pane.append(warning);
    const empty=el('p','hint','아직 확정된 입력표가 없습니다. 수리계산 단계에서 표를 확정하세요.');empty.id='mf-table-empty';pane.append(empty);
    const tableHeading=[...$('panel-design').children].find(n=>n.tagName==='H2'&&n.textContent.trim()==='표');if(tableHeading)tableHeading.remove();
    pane.append($('dg-grid'),$('dg-bore-row'));
    const boreHelp=$('dg-bore-why').closest('.mf-notice');pane.append(boreHelp || $('dg-bore-why'),$('dg-bore-list'));
    const note=el('p','mf-table-caption');note.id='mf-table-caption';pane.append(note);
    $('stage').append(pane);
    // Sidebar entry stays available for discoverability; it never substitutes an engine action.
    const show=btn('mf-open-table','입력표 넓게 보기',()=>setTable(true));$('panel-design').append(show);
    // Empty rows left by moving individual download buttons should not consume space.
    side.querySelectorAll('.row').forEach(n=>{if(!n.children.length)n.classList.add('mf-empty-row');});
    document.body.classList.add('mf-clean');window.dispatchEvent(new Event('resize'));
  }
  function setTable(open) {
    tableOpen=!!open && ['design','conv'].includes(window.__mf.stage);
    $('mf-table-pane').hidden=!tableOpen;document.body.classList.toggle('mf-table-open',tableOpen);
    $('mf-network').setAttribute('aria-pressed',String(!tableOpen));$('mf-table').setAttribute('aria-pressed',String(tableOpen));
    if(tableOpen)decorateTable();queue();window.dispatchEvent(new Event('module-f-table-visibility'));window.dispatchEvent(new Event('resize'));
  }
  function decorateTable() {
    const table=$('dg-grid').querySelector('table');if(!table)return;
    const which=$('dg-table').value,rows=window.__mf.design?.tables?.[which] || [],keys=[];
    for(const r of rows)for(const key of Object.keys(r))if(key!=='bore_provenance'&&!keys.includes(key))keys.push(key);
    const headers=[...table.querySelectorAll('thead th')],order=basic[which] || [];
    const indices=[...order.map(k=>keys.indexOf(k)).filter(i=>i>=0),...headers.map((_,i)=>i).filter(i=>!order.includes(keys[i]))];
    const trs=[...table.querySelectorAll('tr')];
    // Read native cells by their stable original index, including repeat decorations.
    for(const tr of trs){const cells=[...tr.children];cells.forEach((cell,i)=>{if(!cell.dataset.mfIndex)cell.dataset.mfIndex=String(i+1);});
      const at=new Map(cells.map(c=>[Number(c.dataset.mfIndex)-1,c]));
      for(const index of indices){const cell=at.get(index);if(!cell)continue;const key=keys[index] || 'override_note';
        cell.dataset.mfColumn=key;cell.classList.toggle('mf-number',numeric.has(key));cell.hidden=!fullColumns&&!order.includes(key)&&key!=='override_note';
        if(cell.tagName==='TH'){text(cell,columns[key] || cell.textContent);cell.scope='col';}
        tr.append(cell);
      }
    }
    table.querySelectorAll('tbody tr').forEach((tr,i)=>{
      const row=rows[i] || {};
      tr.querySelectorAll('td').forEach(cell=>{const key=cell.dataset.mfColumn;if(!(key in row))return;let value=row[key];
        if(numeric.has(key)&&typeof value==='number'&&Number.isFinite(value)){
          cell.title=`원래 값: ${value}`;
          if(!fullColumns)value=numberFormat.format(value);
        }
        if(key==='dia_src')value=sources[value] || value;
        if(key==='flow_direction'&&value==='solver_reference')value='기준 방향 · 유향 미확정';
        if(value===null || value===undefined || value==='')value='—';else if(typeof value==='object')value=JSON.stringify(value);else if(typeof value==='boolean')value=value?'예':'아니오';
        text(cell,String(value));
      });
      tr.tabIndex=0;tr.setAttribute('aria-label',`${which==='pipes'?'배관':'항목'} ${row.label ?? row.pipe ?? i+1} 상세`);
      tr.onkeydown=e=>{if(e.key==='Enter'||e.key===' '){e.preventDefault();tr.click();}};
    });
    text($('mf-row-count'),`${rows.length}행`);text($('mf-all-columns'),fullColumns?'기본 열 보기':'상세 열 보기');
    $('mf-all-columns').setAttribute('aria-pressed',String(fullColumns));
    text($('mf-table-caption'),`${window.__mf.stage==='conv'?'수정은 04 수리계산에서 / ':''}행 선택: 속성·근거 확인 / ${fullColumns?'원래 정밀도 표시':'표시만 소수 셋째 자리까지 · 원래 값은 상세 열 또는 마우스를 올려 확인'}`);
  }
  function sync() {
    pending=0;const s=window.__mf,stage=s.stage;
    document.body.classList.toggle('mf-table-review',stage==='conv');
    if(stage!==previousStage || s.slot!==previousSlot){previousStage=stage;previousSlot=s.slot;if(tableOpen)setTable(false);document.querySelector('.side').scrollTop=0;}
    text($('mf-work-title'),labels[stage] || stage);
    const tabular=['design','conv'].includes(stage);
    $('mf-table').hidden=!tabular;$('mf-network').hidden=!tabular;
    $('mf-settings').hidden=!settings.has(`panel-${stage}`);
    $('mf-table-empty').hidden=!!$('dg-grid').querySelector('tbody tr');
    $('mf-table-build').hidden=stage!=='design' || $('dg-build-row').classList.contains('hidden');
    $('mf-table-build').disabled=busy() || $('dg-build').disabled;
    $('mf-table').disabled=busy();$('mf-settings').disabled=busy();
    for(const {node,wrapper} of notices)wrapper.hidden=node.classList.contains('hidden') || !node.textContent.trim();
    const stale=$('dg-stale'),review=$('dg-review-notice');
    const messages=[];
    if(!stale.classList.contains('hidden')&&stale.textContent.trim())messages.push('입력표 갱신 필요 — 수리계산에서 표를 다시 확정하세요.');
    if(!review.classList.contains('hidden')&&review.textContent.trim())messages.push('검토용 관경 · 미확정 상태 유지');
    const issue=$('dg-issues-n').textContent.trim();if(parseInt(issue.replaceAll(',',''),10)>0)messages.push(`확인할 항목 ${issue}`);
    $('mf-table-warning').hidden=!messages.length;text($('mf-table-warning'),messages.join(' / '));
    const mode=$('ed-network-mode').value,k=$('ed-k').value;
    text($('mf-edit-config'),`${{tree:'트리',loop:'루프',grid:'그리드'}[mode] || mode} · 기준 헤드 ${k || '—'}개`);
    text($('mf-design-config'),`${$('dg-sched').value} · FX ${$('dg-fx').value || '없음'}${mode!=='tree'?' · 검토용':''}`);
    text($('mf-export-config'),'출력: '+[['cv-worst-sdf','SDF + SLF'],['cv-full-kfp','전체망 KFP'],['cv-worst-kfp','최불리 KFP']].filter(([id])=>$(id).checked).map(([,name])=>name).join(' · '));
    $('status').title=$('status').textContent;
    for(const step of $('steps').children){step.setAttribute('role','button');step.tabIndex=step.onclick?0:-1;step.setAttribute('aria-disabled',String(!step.onclick));step.onkeydown=e=>{if((e.key==='Enter'||e.key===' ')&&step.onclick&&!busy()){e.preventDefault();step.click();}};}
  }
  function queue(){if(!pending)pending=requestAnimationFrame(sync);}
  try {
    layout();
    const observer=new MutationObserver(queue);observer.observe($('steps'),{childList:true});
    for(const id of ['busy','dg-build-row','dg-build','dg-issues-n'])observer.observe($(id),{attributes:true,childList:true,subtree:true,characterData:true,attributeFilter:['class','disabled']});
    observer.observe($('status'),{childList:true,subtree:true,characterData:true});
    for(const {node} of notices)observer.observe(node,{attributes:true,childList:true,subtree:true,characterData:true,attributeFilter:['class']});
    new MutationObserver(()=>{decorateTable();queue();}).observe($('dg-grid'),{childList:true});
    document.addEventListener('change',queue);
    window.addEventListener('keydown',e=>{if(e.key==='Escape'&&tableOpen&&!busy()&&$('dg-ins').classList.contains('hidden')){setTable(false);$('mf-table').focus();}});
    sync();
  } catch(error) {
    const note=el('div','banner',`화면 정리를 완료하지 못했습니다. 새로고침해 주세요. ${error.message}`);note.id='mf-layout-error';note.setAttribute('role','alert');document.querySelector('header').after(note);console.error(error);
  }
})();
