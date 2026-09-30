/* Drawing-first H attributes. Overlay is read-only until an explicit bulk apply. */
(() => {
  'use strict';
  const $=id=>document.getElementById(id), stage=$('stage'), native=$('cv');
  const state=()=>window.__mf, busy=()=>!$('busy').classList.contains('hidden');
  const scope=document.createElement('label');scope.className='f';scope.innerHTML='<span>역산 관경 범위</span><select id="ha-sizing-scope"><option value="unspecified">미지정 관경만 · 도면값 보호</option><option value="enlarge">기존 관경 이상 · 미지정 포함</option><option value="all">전체 관경 재검토</option></select>';
  $('sz-mode').closest('label').after(scope);
  $('ha-sizing-scope').onchange=()=>{$('sz-mode').dispatchEvent(new Event('input'));};
  const bar=document.createElement('div');bar.id='h-attributes-bar';bar.hidden=true;
  bar.innerHTML='<div class="ha-summary"><button id="ha-open">관경 정의</button><span id="ha-count" role="status">도면 근거 확인 중</span><label><input type="checkbox" id="ha-show" checked>근거 표시</label></div><div class="ha-view-options" aria-label="도면 보기"><label><input type="checkbox" id="ha-iso">등각투상</label><label><input type="checkbox" id="ha-original" checked>원본 도면</label><label><input type="checkbox" id="ha-table" aria-controls="h-properties">전체 테이블</label></div>';
  stage.append(bar);
  const recovery=document.createElement('div');recovery.id='ha-history-conflict';recovery.hidden=true;
  recovery.setAttribute('role','region');recovery.setAttribute('aria-label','이전 편집 확인');
  recovery.innerHTML='<header><strong>이전 편집 확인 필요</strong></header><small id="ha-recovery-summary"></small><small>기록은 삭제하지 않습니다. 계속하면 이전 편집을 별도로 보관하고 새 기준망을 사용합니다.</small><div class="ha-row"><button id="ha-review-history">편집 기록 보기</button><button id="ha-accept-basis">이전 편집 보관 후 계속</button><button id="ha-rebuild-basis" hidden>현재 선택으로 기준망 갱신</button></div><small id="ha-recovery-message" role="status"></small>';
  // Recovery must not inherit the toolbar's visibility:hidden while a drawer
  // (especially processing history) is open. It is independent of table loading.
  document.body.append(recovery);
  window.ModuleHWindows.register(recovery,recovery.querySelector('header'));
  let recovering=false;
  function historyConflict(){
    const s=state(),e=window.moduleFNetworkEditor?.getState();
    return s?.stage==='design'&&s.slot==='plan'&&e?.sid===s.sid&&e.scope==='design'&&e.conflict?e:null;
  }
  function showHistoryConflict(){
    const e=historyConflict();if(!e)return false;
    recovery.hidden=false;
    $('ha-recovery-summary').textContent=`배관망 기준이 달라 이전 편집 ${e.undo||0}건의 적용을 보류했습니다.`;
    $('ha-accept-basis').disabled=busy()||recovering||!e.can_accept_basis;
    $('ha-review-history').disabled=busy()||recovering;
    $('ha-rebuild-basis').hidden=!e.basis_stale;
    $('ha-rebuild-basis').disabled=busy()||recovering;
    window.ModuleHUI?.queue();return true;
  }
  $('ha-review-history').onclick=()=>{if(!busy()&&historyConflict())$('ne-history').click();};
  $('ha-rebuild-basis').onclick=()=>{if(!busy()&&!recovering&&historyConflict()?.basis_stale)$('dg-build').click();};
  $('ha-accept-basis').onclick=async()=>{
    if(recovering||busy())return;recovering=true;$('ha-accept-basis').disabled=true;$('ha-recovery-message').textContent='';
    try{if(await window.moduleFNetworkEditor.resolveConflict())await refresh();}
    catch(e){$('ha-recovery-message').textContent=e.message;}
    finally{recovering=false;}
  };
  window.ModuleHView?.init();
  const panel=document.createElement('aside');panel.id='h-attributes-panel';panel.hidden=true;
  panel.setAttribute('aria-label','관경 정의');
  panel.innerHTML=`<header><strong>관경 정의</strong><button id="ha-close" aria-label="관경 정의 닫기">닫기 ×</button></header>
    <section><div id="ha-legend"><span style="--ink:#53dbc6">직접 표기</span><span style="--ink:#5b9df1">연속 관로</span><span style="--ink:#b79af5">반복 구조</span><span style="--ink:#e0ae57">미지정</span></div>
    <div class="ha-row"><button id="ha-analyse">도면 근거 다시 해석</button><button id="ha-undo">되돌리기</button></div><div id="ha-message" role="status"></div></section>
    <section><div class="ha-row" id="ha-modes"><button data-mode="inspect" aria-pressed="true">보기</button><button data-mode="click">클릭</button><button data-mode="lasso">펜 선택</button><button data-mode="rect">사각 선택</button></div>
    <div class="ha-row"><select id="ha-filter" aria-label="선택 대상"><option value="all">배관 + 헤드</option><option value="pipe">배관만</option><option value="head">헤드만</option></select><button id="ha-clear">선택 해제</button><span id="ha-selected">0개 선택</span></div>
    <small>Shift + 클릭 / 드래그로 추가·해제</small><div id="ha-evidence"></div></section>
    <section><strong>도면 표기 연결</strong><div class="ha-row"><select id="ha-annotation" aria-label="도면 관경 표기"></select><button id="ha-link">선택 배관에 연결</button></div>
    <div class="ha-row"><button id="ha-copy">대표 복사 · Ctrl+C</button><button id="ha-paste">대응 확인 · Ctrl+V</button></div>
    <label><input id="ha-overwrite" type="checkbox">기존 도면·사용자 관경도 덮어쓰기</label>
    <div id="ha-copy-preview" hidden><div class="ha-row"><span>상대 구간 길이 기준 · 수동 대응</span><button id="ha-reverse">반대 방향</button></div><div id="ha-copy-rows"></div><button id="ha-paste-apply" class="ha-primary">미리보기 적용</button></div></section>
    <section><details id="ha-manual"><summary>선택 항목 일괄 수정</summary>
    <div class="ha-row"><label>호칭경 <select id="ha-diameter" aria-label="일괄 호칭경"></select></label></div>
    <div class="ha-row"><label>헤드 K <input id="ha-k" type="number" min="0.01" step="any"></label><label>최소 압력(bar) <input id="ha-pressure" type="number" min="0" step="any"></label></div>
    <button id="ha-apply" class="ha-primary">선택 항목에 적용</button></details>
    <details><summary>수직 접속관 선택</summary><div id="ha-generated"></div></details></section>`;
  document.body.append(panel);
  const canvas=document.createElement('canvas');canvas.id='h-attributes-canvas';canvas.hidden=true;stage.append(canvas);
  const ctx=canvas.getContext('2d'), colors={drawing_direct:'#53dbc6',text:'#53dbc6',drawing_run:'#5b9df1',drawing_repeat:'#b79af5',drawing_mixed:'#b79af5',user:'#edce88',nfpc_fallback:'#8aabb0',unresolved:'#e0ae57'};
  let data=null,selected=new Set(),mode='inspect',gesture=null,hover=null,copy=null,reverse=false,preview=null;
  let lastDesign=null,lastSid=null,lastStage=null,attempt=null,fetching=false,generation=0,live=[],paintStamp='',running=false;
  const active=()=>state()?.stage==='design' && state()?.slot==='plan';
  function message(text,error=false){$('ha-message').textContent=text;$('ha-message').dataset.error=String(error);}
  async function request(path,body){
    const operation=globalThis.crypto?.randomUUID?.()||`h-${Date.now()}-${Math.random().toString(36).slice(2)}`;
    const response=await fetch('/api/module-f/h-attributes/'+path+(body?'':`?sid=${encodeURIComponent(state().sid)}`),body?{method:'POST',headers:{'Content-Type':'application/json','X-Module-F-Operation':operation},body:JSON.stringify({sid:state().sid,revision:data?.revision,...body})}:{cache:'no-store'});
    const result=await response.json();if(!response.ok || !result.ok)throw Error(result.message||result.error||'속성 처리 실패');return result;
  }
  async function refresh(){
    if(fetching||!state()?.sid)return;fetching=true;const seq=++generation,sid=state().sid;
    try{const d=await request('state');if(seq!==generation||sid!==state().sid)return;data=d;live=[];
      selected=new Set([...selected].filter(id=>d.records.some(r=>r.id===id)));
      $('ha-count').textContent=`관경 정의 ${d.records.filter(r=>r.kind==='pipe'&&r.nominal_mm).length} · 미지정 ${d.pending}`;
      $('ha-undo').disabled=!d.can_undo;options();selectionChanged();paintStamp='';
      window.dispatchEvent(new CustomEvent('module-h-evidence-state',{detail:d}));
      if(d.pending)message(`미지정 ${d.pending}개 · 근거를 연결하거나 통합에서 역산하세요.`);else message('도면 근거 해석 완료 · 관경을 클릭하면 출처가 표시됩니다.');
    }catch(e){message(e.message,true);}finally{fetching=false;}
  }
  function applyReferencePatch(patch){
    if(!data)return;
    generation++; // An older in-flight GET must not overwrite the committed patch.
    data={...data,revision:patch.revision,pending:patch.pending,can_undo:patch.can_undo,
      records:data.records.map(r=>r.kind==='pipe'&&r.label===patch.record.label?patch.record:r)};
    live=[];lastDesign=state().design;
    $('ha-count').textContent=`관경 정의 ${data.records.filter(r=>r.kind==='pipe'&&r.nominal_mm).length} · 미지정 ${data.pending}`;
    selectionChanged();paintStamp='';
    window.dispatchEvent(new CustomEvent('module-h-evidence-state',{detail:data}));
  }
  function option(value,text){const n=document.createElement('option');n.value=value;n.textContent=text;return n;}
  function options(){
    const list=$('ha-annotation'),old=list.value;list.replaceChildren(option('','원본 관경 표기 선택'));
    data.annotations.forEach(a=>list.append(option(a.identity,`${a.nominal_mm}A · 표기 ${a.identity} (${Math.round(a.x)}, ${Math.round(a.y)})`)));list.value=old;
    const sizes=[...new Set(data.catalog.map(r=>r.dia))].sort((a,b)=>a-b),select=$('ha-diameter'),was=select.value;
    select.replaceChildren(option('','선택'));sizes.forEach(d=>select.append(option(d,`${d}A`)));select.value=was;
    const generated=$('ha-generated');generated.replaceChildren();
    data.records.filter(r=>r.generated).forEach(r=>{const label=document.createElement('label'),input=document.createElement('input');input.type='checkbox';input.checked=selected.has(r.id);input.dataset.recordId=r.id;
      input.onchange=()=>{input.checked?selected.add(r.id):selected.delete(r.id);selectionChanged();};label.append(input,document.createTextNode(` ${r.label} · ${r.nominal_mm?r.nominal_mm+'A':'미지정'}`));generated.append(label);});
  }
  function selectionChanged(){
    const rows=data?.records.filter(r=>selected.has(r.id))||[];
    $('ha-selected').textContent=`${rows.filter(r=>r.kind!=='head').length} 배관 · ${rows.filter(r=>r.kind==='head').length} 헤드`;
    $('ha-copy-preview').hidden=true;preview=null;
    $('ha-generated').querySelectorAll('input').forEach(n=>n.checked=selected.has(n.dataset.recordId));
    const out=$('ha-evidence');out.replaceChildren();
    for(const r of rows.slice(0,3)){
      const div=document.createElement('div'),e=r.evidence||{};
      const origin=e.manual?.note||e.method||({nfpc_fallback:'트리 기준개수',user:'사용자 수정'}[r.source])||r.source||'헤드';
      div.textContent=`${r.label} · ${r.kind==='head'?'헤드':r.nominal_mm?r.nominal_mm+'A'+(r.inner_mm?' / 실제 내경 '+r.inner_mm+' mm':''):'미지정'} · ${origin}`;out.append(div);
      if(e.representative_nodes){const note=document.createElement('div');note.textContent='대표 경로 '+e.representative_nodes.join(' → ');out.append(note);}
      (e.review_reasons||[]).forEach(t=>{const note=document.createElement('div');note.textContent=t;out.append(note);});
    }
    paintStamp='';
    window.dispatchEvent(new CustomEvent('module-h-attributes-selection',{detail:{ids:[...selected]}}));
  }
  window.addEventListener('module-h-evidence-select',e=>{selected=new Set(e.detail.ids);selectionChanged();if(!e.detail.keepView)ensurePlan();});
  function ensurePlan(){if(!$('dg-plan').checked){$('dg-plan').checked=true;$('dg-plan').dispatchEvent(new Event('change'));}}
  function setMode(next){mode=next;if(next!=='inspect')ensurePlan();$('ha-modes').querySelectorAll('button').forEach(n=>n.setAttribute('aria-pressed',String(n.dataset.mode===mode)));canvas.classList.toggle('ha-selecting',mode!=='inspect'&&!busy());paintStamp='';window.dispatchEvent(new CustomEvent('module-h-attributes-mode',{detail:{mode}}));}
  function basisKey(){return `${state().sid}:${state().slot}:${JSON.stringify(state().edit?.worst||{})}:${$('dg-sched').value}:${$('dg-review-datum').value}`;}
  function rebuild(){if(busy()||!active())return;ensurePlan();attempt=basisKey();live=[];message('도면 표기·연속 관로·반복 구조를 해석하고 있습니다.');$('dg-build').click();}
  function selectedRows(){if(!data||!selected.size)throw Error('배관 또는 헤드를 선택하세요.');return data.records.filter(r=>selected.has(r.id));}
  async function mutate(body){
    if(busy())return;const result=await request('apply',body);state().ovDirty=true;copy=null;message(result.skipped?`${result.skipped}개 기존 관경 보호 · 나머지 반영 중`:'속성 반영 중');rebuild();
  }
  function captureCopy(){const rows=selectedRows();if(rows.some(r=>!r.edges?.length||!r.nominal_mm))throw Error('관경이 정의된 연속 평면 배관 경로를 선택하세요.');copy={ids:rows.map(r=>r.id),sid:state().sid};message(`대표 ${rows.length}구간 복사됨 · 대상 경로를 선택하세요.`);}
  async function paste(){if(!copy||copy.sid!==state().sid)throw Error('같은 도면에서 대표 배관을 먼저 복사하세요.');const rows=selectedRows();preview=await request('copy-preview',{source_ids:copy.ids,ids:rows.map(r=>r.id),reverse});
    const body=$('ha-copy-rows');body.replaceChildren();preview.rows.forEach(r=>{const label=document.createElement('label'),select=document.createElement('select');label.append(document.createTextNode(`${r.label} ${r.protected?'· 기존값 보호':''} `));select.dataset.target=r.id;
      if(r.needs_choice)select.append(option('','경계 겹침 · 직접 선택'));
      [...new Set(data.catalog.map(p=>p.dia))].sort((a,b)=>a-b).forEach(v=>select.append(option(v,`${v}A`)));select.value=r.needs_choice?'':String(r.nominal_mm);label.append(select);body.append(label);});
    $('ha-copy-preview').hidden=false;message('경로 방향과 관경 경계를 확인한 뒤 적용하세요. 형상·길이·헤드 위치는 바뀌지 않습니다.');
  }
  function safe(fn){return async()=>{try{await fn();}catch(e){message(e.message,true);}};}
  $('ha-open').onclick=()=>{panel.hidden=!panel.hidden;if(!panel.hidden)ensurePlan();};
  $('ha-close').onclick=()=>{panel.hidden=true;setMode('inspect');};
  $('ha-modes').onclick=e=>{if(e.target.dataset.mode)setMode(e.target.dataset.mode);};
  $('ha-clear').onclick=()=>{window.moduleFNetworkEditor?.clearSelection();selected.clear();hover=null;selectionChanged();};
  $('ha-analyse').onclick=rebuild;$('ha-copy').onclick=safe(captureCopy);$('ha-paste').onclick=safe(paste);
  $('ha-reverse').onclick=safe(async()=>{reverse=!reverse;await paste();});
  $('ha-paste-apply').onclick=safe(async()=>{if(!preview)throw Error('대응을 먼저 확인하세요.');const choices={};for(const n of $('ha-copy-rows').querySelectorAll('select')){if(!n.value)throw Error('경계가 겹친 구간의 관경을 지정하세요.');choices[n.dataset.target]=Number(n.value);}await mutate({action:'paste',ids:[...selected],choices,overwrite:$('ha-overwrite').checked});});
  $('ha-link').onclick=safe(()=>mutate({action:'annotation',ids:selectedRows().map(r=>r.id),annotation_id:$('ha-annotation').value,overwrite:$('ha-overwrite').checked}));
  $('ha-apply').onclick=safe(()=>{selectedRows();return mutate({action:'manual',ids:[...selected],nominal_mm:$('ha-diameter').value||null,head_properties:{k_factor_si:$('ha-k').value?Number($('ha-k').value):null,required_pressure_bar:$('ha-pressure').value?Number($('ha-pressure').value):null}});});
  $('ha-undo').onclick=safe(async()=>{await request('undo',{});state().ovDirty=true;rebuild();});
  $('ha-show').onchange=()=>paintStamp='';$('ha-annotation').onchange=()=>paintStamp='';
  window.addEventListener('keydown',e=>{if(!active()||panel.hidden||busy()||e.target.closest('input,textarea,select,[contenteditable=true]'))return;
    if((e.ctrlKey||e.metaKey)&&['c','v'].includes(e.key.toLowerCase())){e.preventDefault();e.stopImmediatePropagation();safe(e.key.toLowerCase()==='c'?captureCopy:paste)();}
    if(e.key==='Escape'){gesture=null;setMode('inspect');}
  },true);
  const screen=p=>[state().toScreenX(p[0]),state().toScreenY(p[1])];
  function distance(p,a,b){const dx=b[0]-a[0],dy=b[1]-a[1],t=Math.max(0,Math.min(1,((p[0]-a[0])*dx+(p[1]-a[1])*dy)/(dx*dx+dy*dy||1)));return Math.hypot(p[0]-a[0]-t*dx,p[1]-a[1]-t*dy);}
  function inPolygon(p,poly){let yes=false;for(let i=0,j=poly.length-1;i<poly.length;j=i++){const a=poly[i],b=poly[j];if((a[1]>p[1])!==(b[1]>p[1])&&p[0]<(b[0]-a[0])*(p[1]-a[1])/(b[1]-a[1])+a[0])yes=!yes;}return yes;}
  function eligible(r){return $('ha-filter').value==='all'||($('ha-filter').value==='head' ? r.kind==='head' : r.kind!=='head');}
  function hit(p){let best=null,d=10;for(const r of data?.records||[]){if(!eligible(r))continue;const n=r.kind==='head'?Math.hypot(...screen(r.xy[0]).map((x,i)=>x-p[i])):Math.min(...r.xy.map(([a,b])=>distance(p,screen(a),screen(b))));if(n<d){best=r;d=n;}}return best;}
  function pointer(e){const box=canvas.getBoundingClientRect();return[e.clientX-box.left,e.clientY-box.top];}
  canvas.onpointerdown=e=>{if(e.button!==0||busy()||!data)return;e.preventDefault();canvas.setPointerCapture(e.pointerId);gesture={points:[pointer(e)],shift:e.shiftKey};};
  canvas.onpointermove=e=>{if(gesture){const p=pointer(e),last=gesture.points.at(-1);if(Math.hypot(p[0]-last[0],p[1]-last[1])>3)gesture.points.push(p);paintStamp='';}else{const r=hit(pointer(e));if(r?.id!==hover){hover=r?.id;paintStamp='';}}};
  canvas.onpointerup=e=>{if(!gesture)return;const g=gesture;gesture=null;let hits=[];
    if(mode==='click'||g.points.length<3){const r=hit(pointer(e));if(r)hits=[r];}
    else{const p=g.points[0],q=pointer(e),poly=mode==='rect'?[p,[q[0],p[1]],q,[p[0],q[1]]]:g.points;
      hits=(data?.records||[]).filter(r=>eligible(r)&&(r.kind==='head'?inPolygon(screen(r.xy[0]),poly):r.xy.every(([a,b])=>inPolygon(screen(a),poly)&&inPolygon(screen(b),poly))));}
    if(!g.shift)selected.clear();for(const r of hits){if(g.shift&&selected.has(r.id))selected.delete(r.id);else selected.add(r.id);}selectionChanged();};
  canvas.onpointercancel=()=>{gesture=null;paintStamp='';};
  canvas.onwheel=e=>{e.preventDefault();native.dispatchEvent(new WheelEvent('wheel',{deltaY:e.deltaY,deltaX:e.deltaX,clientX:e.clientX,clientY:e.clientY,bubbles:true,cancelable:true}));};
  function paint(){
    running=false;if(!active())return;const plan=window.ModuleHView.isPlan();
    canvas.hidden=!data&&!live.length;canvas.classList.toggle('ha-selecting',plan&&mode!=='inspect'&&!busy()&&!panel.hidden);
    const box=native.getBoundingClientRect(),dpr=devicePixelRatio||1;
    const stamp=[data?.revision,state().view?.scale,state().view?.ox,state().view?.oy,box.width,box.height,hover,selected.size,live.length,$('ha-show').checked,$('ha-annotation').value,gesture?.points.length,window.ModuleHView.projectionKey()].join('|');
    if(stamp!==paintStamp){paintStamp=stamp;canvas.width=Math.round(box.width*dpr);canvas.height=Math.round(box.height*dpr);ctx.setTransform(dpr,0,0,dpr,0,0);
      function line(a,b,color,width=2,dash=[]){a=screen(a);b=screen(b);ctx.strokeStyle=color;ctx.lineWidth=width;ctx.setLineDash(dash);ctx.beginPath();ctx.moveTo(...a);ctx.lineTo(...b);ctx.stroke();ctx.setLineDash([]);}
      const records=live.length?live:(data?.records||[]);
      // Selection is a solid foreground stroke, even for an unspecified bore.
      // Keep every unselected source/network colour unchanged; paint selections
      // last so adjoining evidence strokes cannot cover their active colour.
      const ordered=[...records.filter(r=>!selected.has(r.id)),...records.filter(r=>selected.has(r.id))];
      for(const r of ordered){const picked=selected.has(r.id),emphasis=picked||r.id===hover;if(!emphasis)continue;
        const color=picked?'#ff3b45':emphasis?'#fff1ad':colors[r.source]||'#e0ae57';
        const geometry=window.ModuleHView.geometry(r);
        if(r.kind==='head'){if(!emphasis||!geometry.point)continue;ctx.strokeStyle=color;ctx.lineWidth=picked?3:2;ctx.beginPath();ctx.arc(...screen(geometry.point),picked?10:8,0,Math.PI*2);ctx.stroke();}
        else for(const [a,b] of geometry.segments){if(a[0]===b[0]&&a[1]===b[1]){if(!emphasis)continue;ctx.strokeStyle=color;ctx.lineWidth=picked?3:2;ctx.beginPath();ctx.arc(...screen(a),picked?12:10,0,Math.PI*2);ctx.stroke();}else line(a,b,color,picked?7:emphasis?5:2,picked||r.nominal_mm?[]:[5,5]);}
        // Representative connectors are owned by the evidence overlay. Drawing
        // them here too duplicates links and uses the wrong compacted-pipe end.
      }
      const annotation=data?.annotations.find(a=>a.identity===$('ha-annotation').value);
      const annotationPoint=annotation&&window.ModuleHView.projectPoint([annotation.x,annotation.y]);
      if(annotationPoint){const v=window.ModuleHView,l=v.annotationLayout(annotation,ctx);ctx.strokeStyle='#53dbc6';ctx.setLineDash([4,4]);v.annotationBorder(ctx,l);ctx.setLineDash([]);}
      if(gesture){const p=gesture.points[0],q=gesture.points.at(-1);ctx.strokeStyle='#5b8def';ctx.fillStyle='#5b8def20';ctx.setLineDash([5,4]);ctx.beginPath();if(mode==='rect')ctx.rect(p[0],p[1],q[0]-p[0],q[1]-p[1]);else{ctx.moveTo(...p);gesture.points.forEach(t=>ctx.lineTo(...t));ctx.closePath();}ctx.fill();ctx.stroke();ctx.setLineDash([]);}
    }
    running=true;requestAnimationFrame(paint);
  }
  window.addEventListener('module-h-diameter-progress',e=>{const d=e.detail.data||{};if(d.reset)live=[];for(const r of d.snapshot||[d])if(r.edge&&r.xy)live.push({id:String(r.edge),kind:'pipe',xy:[r.xy],nominal_mm:r.evidence?.text_mm,source:r.evidence?.definition_source,evidence:r.evidence});$('ha-count').textContent=e.detail.phase+(live.length?` · ${live.length}구간`:'');paintStamp='';});
  // Float like other H auxiliary windows, without moving any drawing coordinates.
  const scroll=document.createElement('div');scroll.className='h-window-scroll';
  [...panel.children].filter(n=>n.tagName!=='HEADER').forEach(n=>scroll.append(n));panel.append(scroll);
  window.ModuleHWindows.register(panel,panel.querySelector('header'));
  window.ModuleHAttributes={getState:()=>data,refresh,applyReferencePatch,getMode:()=>mode,historyConflict,showHistoryConflict};
  setInterval(()=>{
    const s=state();if(!s)return;
    window.ModuleHView?.sync();
    if(s.sid!==lastSid){lastSid=s.sid;data=null;lastDesign=null;selected.clear();copy=null;attempt=null;generation++;}
    bar.hidden=!active();if(!active()){recovery.hidden=true;panel.hidden=true;canvas.hidden=true;lastStage=s.stage;return;}
    const bounds=stage.getBoundingClientRect();let top=12;
    for(const id of ['h-alert','bore-overlay']){const n=$(id);if(n?.getClientRects().length){const r=n.getBoundingClientRect();top=Math.max(top,r.bottom-bounds.top+8);}}
    bar.style.top=top+'px';
    if(lastStage!==s.stage){lastStage=s.stage;attempt=null;}
    if(!running){running=true;requestAnimationFrame(paint);}
    const editor=historyConflict();
    recovery.hidden=!editor;
    if(editor)showHistoryConflict();
    $('ha-accept-basis').disabled=busy()||recovering||!editor?.can_accept_basis;
    $('ha-review-history').disabled=busy()||recovering;
    if(!recovery.hidden&&!editor.can_accept_basis)$('ha-recovery-message').textContent=editor.basis_stale
      ? '도면 선정이 다시 바뀌었습니다. 먼저 현재 선택으로 기준망을 갱신하세요. 이전 기록은 보존됩니다.'
      : '복구 기능 적용을 위해 서버를 다시 시작하세요. 기존 기록은 보존됩니다.';
    if(busy())return;
    // A held history needs a user's decision, not another automatic rebuild.
    if(editor)return;
    if(s.design!==lastDesign){lastDesign=s.design;if(s.design?.settings?.diameter_policy==='drawing_first_v1'&&s.design.tables)refresh();}
    const key=basisKey(),stale=!$('dg-stale').classList.contains('hidden');
    if(s.edit?.worst?.heads?.length&&(!s.design?.tables||s.design?.settings?.diameter_policy!=='drawing_first_v1'||s.ovDirty||stale)&&attempt!==key)rebuild();
  },300);
})();
