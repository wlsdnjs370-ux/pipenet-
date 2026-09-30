/* H bottom-docked property grid. All writes use the canonical editor transaction. */
(() => {
  'use strict';
  const $=id=>document.getElementById(id), engine=()=>window.moduleFNetworkEditor;
  const panel=document.createElement('aside');panel.id='h-properties';panel.hidden=true;
  panel.setAttribute('aria-label','배관망 속성표');
  panel.innerHTML=`<header tabindex="0"><strong id="hp-title">배관망 속성</strong><div><button id="hp-fold">접기</button><button id="hp-close">닫기 ×</button></div></header>
    <nav><select id="hp-kind" aria-label="속성 종류"><option value="pipe">배관</option><option value="node">노드 · 헤드</option><option value="fitting">피팅 · 부속</option></select>
    <input id="hp-filter" aria-label="속성 검색" placeholder="번호 검색"><select id="hp-basis" aria-label="속성표 데이터" hidden><option value="original">원본 · 편집</option><option value="proposal" disabled>역산 결과 · 읽기 전용</option></select><button id="hp-undo">되돌리기</button><button id="hp-redo">다시 실행</button></nav>
    <div id="hp-bulk" hidden><strong id="hp-count"></strong><select id="hp-field" aria-label="일괄 수정 속성"></select><span id="hp-value-wrap"></span><button id="hp-apply">선택에 적용</button><button id="hp-clear">선택 해제</button></div>
    <div id="hp-message" role="status"></div><div class="hp-scroll"><table aria-multiselectable="true"><thead id="hp-head"></thead><tbody id="hp-rows"></tbody></table></div>`;
  document.body.append(panel);document.body.classList.add('h-properties-ready');
  let data=null,lastStage=null,lastSid=null,lastSelection='',pending=false,hiddenByUser=false,opened=false;
  let selections=new Map(),notifying=false,drag=null,suppressRowClick=false,lastSlot=null;
  let proposal=null,proposalStale=false,proposalDirty=false,proposalSid=null;
  const showingProposal=()=>$('hp-basis').value==='proposal'&&window.__mf?.stage==='merge'&&proposalSid===window.__mf?.sid&&!!proposal;
  const selectionKey=(kind,label)=>`${kind}:${label}`;
  const rowsFor=kind=>kind==='pipe'?data?.pipes||[]:kind==='node'?data?.nodes||[]:data?.property_fittings||[];
  const selectedRows=kind=>rowsFor(kind).filter(r=>selections.has(selectionKey(kind,r.label)));
  function selectedFittings(){
    return rowsFor('fitting').filter(r=>selections.has(selectionKey('fitting',r.label))||
      r.node!=null&&selections.has(selectionKey('node',r.node)));
  }
  function fittingCommand(r,fields){return {op:'fitting_properties',target:r.pipe,collection:r.collection,index:r.index,expected_type:r.expected_type,...fields};}
  function headCommand(r,fields){return {op:'head',target:r.label,nozzle:r.head_spec.id,flow:r.head_spec.flow_lmin,...fields};}
  const fmt=v=>v==null?'—':Number.isFinite(+v)?Number(v).toLocaleString('ko-KR',{maximumFractionDigits:4}):String(v);
  function cell(row,text){const td=document.createElement('td');td.textContent=text;row.append(td);return td;}
  function input(row,value,key,label){const n=document.createElement('input');n.type='number';n.step='any';n.value=value==null?'':Number.isFinite(+value)?Number(Number(value).toFixed(4)):value;n.dataset.initialValue=n.value;n.dataset.field=key;n.setAttribute('aria-label',label);n.title=value==null?'':String(value);cell(row,'').append(n);return n;}
  const editedValue=(input,original)=>input.value===input.dataset.initialValue?original:input.value;
  function select(row,options,value,key,label){const n=document.createElement('select');n.dataset.field=key;n.setAttribute('aria-label',label);options.forEach(([v,t])=>n.append(new Option(t,v)));n.value=String(value??'');cell(row,'').append(n);return n;}
  async function commit(command){
    if(pending||!$('busy').classList.contains('hidden'))return;
    pending=true;$('hp-message').textContent='저장 중';panel.classList.add('hp-pending');
    try{const ok=await engine().applyCommand(command);if(ok)lastSelection=JSON.stringify(engine().getSelection());$('hp-message').textContent=ok?'저장됨':$('ne-message').textContent||'입력값을 확인하세요.';}
    catch(e){$('hp-message').textContent=e.message;}
    finally{pending=false;panel.classList.remove('hp-pending');data=null;}
  }
  async function commitMany(commands){
    if(!commands.length)return;
    if(data?.property_batch_version!==1){$('hp-message').textContent='일괄 수정을 사용하려면 서버 업데이트 후 도면을 다시 여세요.';return;}
    return commit(commands.length===1?commands[0]:{op:'batch_properties',commands,note:'H 다중 선택 속성 수정'});
  }
  function fieldCommit(kind,r,fields){
    const chosen=selectedRows(kind),targets=chosen.some(x=>x.label===r.label)?chosen:[r];
    if(kind==='pipe')return commitMany(targets.map(p=>fields.c!=null?{op:'pipe_c',target:p.label,...fields}:{op:'pipe',target:p.label,schedule:p.type,dn:p.dia,...fields,note:'H 속성표'}));
    if(kind==='fitting')return commitMany(targets.map(p=>fittingCommand(p,fields)));
    const heads=targets.filter(n=>n.head_spec);
    return commitMany(heads.map(n=>headCommand(n,fields)));
  }
  function bulkFields(){
    const pipes=selectedRows('pipe').length,heads=selectedRows('node').filter(n=>n.head_spec).length,fits=selectedFittings().length;
    return [[pipes,'dia','배관 호칭경'],[pipes,'material','배관 재질'],[pipes,'c','배관 C'],
      [heads,'flow','헤드 방수량'],[heads,'k','헤드 K'],[heads,'pressure','헤드 최소 압력'],[heads,'nozzle','헤드 종류'],
      [fits,'fitting','부속 종류'],[fits,'count','부속 개수']].filter(([n])=>n);
  }
  function bulkInput(){
    const field=$('hp-field').value;let n,options;
    if(field==='dia')options=[...new Set((data?.catalog.pipes||[]).map(r=>r.dia))].sort((a,b)=>a-b).map(d=>[d,`${d} mm`]);
    if(field==='material')options=[...new Set((data?.catalog.pipes||[]).map(r=>r.type))].map(v=>[v,v]);
    if(field==='nozzle')options=(data?.catalog.nozzles||[]).map(r=>[r.id,r.name]);
    if(field==='fitting')options=(data?.catalog.fittings||[]).map(r=>[r.id,r.name]);
    if(options){n=document.createElement('select');n.append(new Option('값 선택',''));options.forEach(([v,t])=>n.append(new Option(t,v)));}
    else {n=document.createElement('input');n.type='number';n.step=field==='count'?'1':'any';n.placeholder=({flow:'L/min',k:'L/min/√bar',pressure:'bar',c:'1~200',count:'1~100'})[field]||'';}
    n.id='hp-value';n.setAttribute('aria-label','일괄 수정 값');$('hp-value-wrap').replaceChildren(n);
  }
  function updateBulk(){
    const options=bulkFields(),prior=$('hp-field').value,signature=JSON.stringify(options);
    $('hp-bulk').hidden=selections.size<2||showingProposal();
    $('hp-count').textContent=`${selections.size}개 선택`;
    if($('hp-field').dataset.signature!==signature){
      $('hp-field').dataset.signature=signature;$('hp-field').replaceChildren(...options.map(([n,key,name])=>new Option(`${name} · ${n}개`,key)));
      if(options.some(o=>o[1]===prior))$('hp-field').value=prior;bulkInput();
    }
    $('hp-apply').disabled=!options.length||pending;
  }
  $('hp-field').onchange=bulkInput;
  $('hp-apply').onclick=()=>{
    const field=$('hp-field').value,value=$('hp-value')?.value;
    if(value==null||value===''){$('hp-message').textContent='적용할 값을 입력하세요.';return;}
    let commands=[];
    if(['dia','material','c'].includes(field))commands=selectedRows('pipe').map(r=>field==='c'?{op:'pipe_c',target:r.label,c:value}:
      {op:'pipe',target:r.label,schedule:field==='material'?value:r.type,dn:field==='dia'?value:r.dia,note:'H 일괄 속성 수정'});
    if(['flow','k','pressure','nozzle'].includes(field))commands=selectedRows('node').filter(r=>r.head_spec).map(r=>headCommand(r,{[({flow:'flow',k:'k_factor_si',pressure:'required_pressure_bar',nozzle:'nozzle'})[field]]:value}));
    if(['fitting','count'].includes(field))commands=selectedFittings().map(r=>fittingCommand(r,{[field]:value}));
    commitMany(commands);
  };
  function render(){
    if(!data)return;
    if(showingProposal()){renderProposal();return;}
    const kind=$('hp-kind').value,filter=$('hp-filter').value.trim().toLowerCase();
    const headers=kind==='pipe'?['배관','시작','끝','재질','호칭경 (mm)','내경 (mm)','길이 (m)','C','근거']:kind==='fitting'?['부속','배관','노드','종류','개수']:['노드','Z (m)','헤드','방수량 (L/min)','K (L/min/√bar)','최소 압력 (bar)','높이 반영'];
    $('hp-head').replaceChildren();const tr=document.createElement('tr');headers.forEach(t=>{const th=document.createElement('th');th.textContent=t;tr.append(th);});$('hp-head').append(tr);
    $('hp-rows').replaceChildren();
    for(const r of rowsFor(kind).filter(r=>!filter||String(r.label).toLowerCase().includes(filter))){
      const row=document.createElement('tr');row.dataset.label=r.label;row.tabIndex=0;
      row.onclick=e=>{if(suppressRowClick){suppressRowClick=false;return;}if(e.target.closest('input,select,button')&&!e.shiftKey){if(!selections.has(selectionKey(kind,r.label)))choose(kind,r.label);return;}choose(kind,r.label,{additive:e.shiftKey});};
      row.onkeydown=e=>{if(e.target===row&&['Enter',' '].includes(e.key)){e.preventDefault();choose(kind,r.label,{additive:e.shiftKey});}};
      cell(row,r.label);
      if(kind==='pipe'){
        cell(row,r.in);cell(row,r.out);
        const mat=select(row,[...new Set(data.catalog.pipes.map(p=>p.type))].map(t=>[t,t]),r.type,'material',`${r.label} 재질`);
        const sizes=()=>data.catalog.pipes.filter(p=>p.type===mat.value);
        const dn=select(row,[['','미지정'],...sizes().map(p=>[p.dia,p.dia])],r.dia||'','dn',`${r.label} 호칭경`);
        const spec=sizes().find(p=>p.dia===r.dia);cell(row,fmt(r.inner_mm??spec?.inner_mm));
        const length=input(row,r.length,'length',`${r.label} 길이`);length.min='0.000001';
        const c=input(row,r.c,'c',`${r.label} C`);c.min='1';c.max='200';
        cell(row,r.bore_provenance?.manual?'수정':r.bore_provenance?.block_export?'확인 필요':r.dia?'정의됨':'미지정');
        mat.onchange=()=>{dn.replaceChildren(new Option('관경 선택',''));sizes().forEach(p=>dn.append(new Option(p.dia,p.dia)));if(sizes().some(p=>p.dia===r.dia)){dn.value=r.dia;dn.onchange();}};
        dn.onchange=()=>{if(dn.value)fieldCommit(kind,r,{schedule:mat.value,dn:dn.value});};
        length.onchange=()=>{if(selections.size>1){$('hp-message').textContent='길이·좌표 이동은 한 항목씩 수정하세요.';length.value=r.length;return;}commit({op:'resize',target:String(r.label),length:length.value,note:'H 하단 속성표 길이'});};
        c.onchange=()=>fieldCommit(kind,r,{c:c.value});
      }else if(kind==='fitting'){
        row.firstChild.textContent=r.name;cell(row,r.pipe);cell(row,r.node??'—');
        const fit=select(row,data.catalog.fittings.map(f=>[f.id,f.name]),r.library,'fitting',`${r.label} 부속 종류`);
        const count=input(row,r.count,'count',`${r.label} 개수`);count.step='1';count.min='1';count.max='100';
        fit.onchange=()=>fieldCommit(kind,r,{fitting:fit.value});count.onchange=()=>fieldCommit(kind,r,{count:count.value});
      }else{
        const z=input(row,r.xyz[2],'z',`${r.label} Z`);
        const nozzle=select(row,[['','없음'],...data.catalog.nozzles.map(n=>[n.id,n.name])],r.head_spec?.id,'nozzle',`${r.label} 헤드`);
        const flow=input(row,r.head_spec?.flow_lmin,'flow',`${r.label} 방수량`);flow.min='0';flow.disabled=!r.head;
        const k=input(row,r.head_spec?.k_factor_si,'k',`${r.label} K`),pressure=input(row,r.head_spec?.min_bar,'pressure',`${r.label} 최소 압력`);
        k.disabled=pressure.disabled=!r.head;k.min='.001';pressure.min='0';
        const b=document.createElement('button');b.textContent='반영';cell(row,'').append(b);
        // X/Y are not table fields. Preserve their full-precision canonical values.
        b.onclick=e=>{e.stopPropagation();if(selections.size>1){$('hp-message').textContent='길이·좌표 이동은 한 항목씩 수정하세요.';return;}commit({op:'move_node',target:String(r.label),x:r.xyz[0],y:r.xyz[1],z:editedValue(z,r.xyz[2]),note:'H 하단 속성표 높이'});};
        nozzle.onchange=()=>{if(selections.size>1){if(!nozzle.value){$('hp-message').textContent='헤드 삭제는 한 항목씩 확인하세요.';nozzle.value=r.head_spec?.id||'';return;}fieldCommit(kind,r,{nozzle:nozzle.value});return;}if(nozzle.value&&!flow.value){flow.disabled=false;flow.focus();$('hp-message').textContent='방수량을 입력하세요.';return;}commit(nozzle.value?{op:'head',target:String(r.label),nozzle:nozzle.value,flow:flow.value}:{op:'node',target:String(r.label)});};
        flow.onchange=()=>fieldCommit(kind,r,{flow:flow.value});
        const headChange=field=>fieldCommit(kind,r,field);
        k.onchange=()=>headChange({k_factor_si:k.value});pressure.onchange=()=>headChange({required_pressure_bar:pressure.value});
      }
      $('hp-rows').append(row);
    }
    $('hp-undo').disabled=!data.undo;$('hp-redo').disabled=!data.redo;paintSelection();
  }
  function renderProposal(){
    const kind=$('hp-kind').value,filter=$('hp-filter').value.trim().toLowerCase();
    const headers=kind==='pipe'?['배관','원본 호칭경','제안 호칭경 (mm)','실제 내경 (mm)','유량 (L/min)','유속 (m/s)','잠금']:kind==='node'?['헤드','노드','유량 (L/min)','압력 (bar)','최소 유량','최소 압력']:['부속 손실은 배관별 역산 등가길이에 포함됩니다.'];
    const head=document.createElement('tr');headers.forEach(h=>{const th=document.createElement('th');th.textContent=h;head.append(th);});$('hp-head').replaceChildren(head);$('hp-rows').replaceChildren();
    for(const r of (kind==='pipe'?proposal.pipes:kind==='node'?proposal.nozzles:[]).filter(r=>!filter||String(r.label).toLowerCase().includes(filter))){
      const row=document.createElement('tr');row.dataset.label=String(kind==='node'?r.node:r.label);row.tabIndex=0;
      const values=kind==='pipe'?[r.label,r.before_dn==null?'미지정':r.before_dn,r.after_dn,r.inner_mm,r.flow_lpm,r.velocity_mps,r.locked?'잠금':'—']:[r.label,r.node,r.flow_lpm,r.pressure_bar,r.min_flow_lpm,r.min_pressure_bar];
      values.forEach(v=>cell(row,typeof v==='number'?fmt(v):v));row.onclick=e=>{if(suppressRowClick){suppressRowClick=false;return;}choose(kind,row.dataset.label,{additive:e.shiftKey});};$('hp-rows').append(row);
    }
    $('hp-message').textContent=proposalStale||proposalDirty?'이전 역산 결과 · 입력 변경됨 · 재계산 필요':proposal.feasible?'조건 충족 검토안 · 이 내경으로 별도 SDF 출력 · 원본은 유지':'조건 미달 시험값 · 내경 산출됨 · SDF 출력 불가';
    $('hp-undo').disabled=$('hp-redo').disabled=true;paintSelection();
  }
  window.addEventListener('module-h-sizing-result',e=>{
    if(!e.detail.report)return;
    const fresh=proposal?.proposal_id!==e.detail.report.proposal_id;proposal=e.detail.report;proposalStale=!!e.detail.stale;proposalDirty=false;proposalSid=window.__mf?.sid;
    $('hp-basis').options[1].disabled=false;
    if(fresh&&proposal.feasible&&!proposalStale)$('hp-basis').value='proposal';render();
  });
  window.addEventListener('module-h-sizing-dirty',()=>{proposalDirty=true;if(showingProposal())render();});
  $('hp-basis').onchange=()=>{$('hp-message').textContent='';render();};
  function syncSelection(){
    const sel=engine()?.getSelection(),key=JSON.stringify(sel);if(key===lastSelection)return;
    lastSelection=key;
    selections=new Map(sel?[[selectionKey(sel.kind,sel.label),sel]]:[]);
    if(!sel){paintSelection();return;}
    if(!panel.hidden&&data){
      if($('hp-kind').value!==sel.kind){$('hp-kind').value=sel.kind;render();lastSelection=key;}
    }
    paintSelection();
  }
  function paintSelection(){
    const kind=$('hp-kind').value;
    for(const row of $('hp-rows').children){const on=selections.has(selectionKey(kind,row.dataset.label));row.classList.toggle('hp-selected',on);row.setAttribute('aria-selected',String(on));}
    if(!drag)updateBulk();
    if(notifying)return;
    notifying=true;
    try{
      const records=window.ModuleHAttributes?.getState()?.records||[];
      const ids=records.filter(r=>selections.has(selectionKey(r.kind==='head'?'node':'pipe',r.kind==='head'?r.node_label:r.label))).map(r=>r.id);
      if(window.__mf?.stage==='design'&&window.ModuleHAttributes?.getMode()==='inspect')window.dispatchEvent(new CustomEvent('module-h-evidence-select',{detail:{ids,keepView:true}}));
      window.dispatchEvent(new CustomEvent('module-h-property-selection',{detail:{data,selection:engine()?.getSelection(),selections:[...selections.values()]}}));
    }finally{notifying=false;}
  }
  function syncContext(){
    const s=window.__mf;if(!s)return;
    if(s.stage!==lastStage||s.sid!==lastSid||s.slot!==lastSlot){lastStage=s.stage;lastSid=s.sid;lastSlot=s.slot;hiddenByUser=false;opened=false;data=null;lastSelection='';selections.clear();drag=null;}
  }
  // Selection never changes window visibility. Keep highlighting even while closed.
  function choose(kind,label,{additive=false}={}){
    syncContext();label=String(label);const key=selectionKey(kind,label);
    if(!additive)selections.clear();
    if(additive&&selections.has(key))selections.delete(key);else selections.set(key,{kind,label});
    setPrimary();if(!panel.hidden&&data&&$('hp-kind').value!==kind){$('hp-kind').value=kind;render();}paintSelection();
  }
  function setPrimary(){
    const primary=[...selections.values()].at(-1);notifying=true;
    try{
      if(primary){const r=primary.kind==='fitting'&&rowsFor('fitting').find(r=>r.label===primary.label);
        engine().select(r?r.node!=null?'node':'pipe':primary.kind,r?r.node??r.pipe:primary.label);
      }else engine()?.clearSelection();
      lastSelection=JSON.stringify(engine()?.getSelection());
    }finally{notifying=false;}
  }
  function open({all=false}={}){syncContext();opened=true;hiddenByUser=false;panel.hidden=false;
    const kinds=new Set([...selections.values()].map(r=>r.kind));if(kinds.size===1){$('hp-kind').value=[...kinds][0];render();}
    if(all){$('hp-filter').value='';if(panel.classList.contains('hp-folded'))$('hp-fold').click();render();}
    engine()?.refresh();}
  function close(){hiddenByUser=true;panel.hidden=true;}
  window.ModuleHProperties={open,close,isVisible:()=>!panel.hidden,select:choose,getData:()=>data,getSelections:()=>[...selections.values()]};
  window.addEventListener('module-h-attributes-mode',e=>{if(e.detail.mode!=='inspect')panel.hidden=true;});
  $('hp-close').onclick=close;
  window.addEventListener('keydown',e=>{if(e.key==='Escape'&&!panel.hidden&&!pending)close();});
  let unfoldedHeight='';
  $('hp-fold').onclick=()=>{const on=panel.classList.toggle('hp-folded');if(on)unfoldedHeight=panel.style.height;
    panel.style.setProperty('height',on?'48px':unfoldedHeight||'350px','important');$('hp-fold').textContent=on?'펼치기':'접기';};
  $('hp-kind').onchange=render;$('hp-filter').oninput=render;
  $('hp-clear').onclick=()=>{selections.clear();setPrimary();paintSelection();};
  // Drag starts on row labels/blank cells, never steals a normal input gesture.
  const scroll=panel.querySelector('.hp-scroll');
  $('hp-rows').addEventListener('pointerdown',e=>{
    if(e.button||pending||e.target.closest('input,select,button'))return;
    const row=e.target.closest('tr[data-label]');if(!row)return;
    const rows=[...$('hp-rows').children],index=rows.indexOf(row);
    drag={id:e.pointerId,start:index,rows,base:e.shiftKey?new Map(selections):new Map(),x:e.clientX,y:e.clientY,moved:false,kind:$('hp-kind').value};
    scroll.setPointerCapture(e.pointerId);e.preventDefault();
  });
  scroll.addEventListener('pointermove',e=>{
    if(!drag||drag.id!==e.pointerId)return;
    if(Math.hypot(e.clientX-drag.x,e.clientY-drag.y)<4&&!drag.moved)return;
    drag.moved=true;const box=scroll.getBoundingClientRect();
    if(e.clientY<box.top+20)scroll.scrollTop-=22;else if(e.clientY>box.bottom-20)scroll.scrollTop+=22;
    let index=drag.rows.findIndex(r=>{const b=r.getBoundingClientRect();return e.clientY>=b.top&&e.clientY<=b.bottom;});
    if(index<0)index=e.clientY<box.top?0:drag.rows.length-1;
    selections=new Map(drag.base);
    for(let i=Math.min(index,drag.start);i<=Math.max(index,drag.start);i++){const label=drag.rows[i].dataset.label;selections.set(selectionKey(drag.kind,label),{kind:drag.kind,label});}
    paintSelection();
  });
  scroll.addEventListener('pointerup',e=>{if(!drag)return;const gesture=drag,moved=drag.moved;drag=null;scroll.releasePointerCapture(e.pointerId);
    if(moved){suppressRowClick=true;setPrimary();paintSelection();setTimeout(()=>suppressRowClick=false,0);}
    else {const row=gesture.rows[gesture.start];choose(gesture.kind,row.dataset.label,{additive:e.shiftKey});suppressRowClick=true;setTimeout(()=>suppressRowClick=false,0);}
  });
  scroll.addEventListener('pointercancel',()=>{drag=null;});
  $('hp-undo').onclick=()=>engine().undo();$('hp-redo').onclick=()=>engine().redo();
  window.ModuleHWindows.register(panel,panel.querySelector('header'),{minWidth:330,minHeight:180});
  window.addEventListener('module-h-attributes-selection',e=>{
    if(notifying||window.__mf?.stage!=='design'||e.detail.ids?.length!==1||window.ModuleHAttributes?.getMode()!=='inspect')return;
    const r=window.ModuleHAttributes?.getState()?.records.find(r=>r.id===e.detail.ids[0]&&(r.kind==='pipe'||r.node_label));
    const kind=r?.kind==='head'?'node':'pipe',label=r?.kind==='head'?r.node_label:r?.label;
    if(r&&(engine()?.getSelection()?.label!==label||engine()?.getSelection()?.kind!==kind))choose(kind,label);
  });
  setInterval(()=>{
    const s=window.__mf;if(!s||!engine())return;
    syncContext();
    const active=(s.stage==='merge'&&s.mergeView)||(s.stage==='design'&&s.slot==='plan'&&s.design?.tables);
    $('hp-basis').hidden=s.stage!=='merge';
    if(proposalSid!==s.sid){proposal=null;$('hp-basis').value='original';$('hp-basis').options[1].disabled=true;}
    document.body.classList.toggle('h-bottom-design',!!active&&s.stage==='design');
    const bulk=s.stage==='design'&&!$('h-attributes-panel').hidden&&window.ModuleHAttributes?.getMode()!=='inspect';
    const wasHidden=panel.hidden;
    panel.hidden=!active||hiddenByUser||bulk||document.body.classList.contains('h-drawer-open')||(!opened&&s.stage!=='merge');
    if(pending||drag)return;
    $('hp-title').textContent=s.stage==='merge'?'통합 배관망 속성':'배관 속성';
    const current=engine().getState();
    if(current&&current.scope===s.stage&&current!==data){const changed=current.revision!==data?.revision;data=current;if(changed){
      selections=new Map([...selections].filter(([,sel])=>rowsFor(sel.kind).some(r=>String(r.label)===sel.label)));render();
      // Initial editor loading must not re-download an already loaded evidence
      // graph. Full design changes are observed by Attributes itself; lightweight
      // patches carry an explicit expected evidence revision when reconciliation
      // is actually needed.
      if(s.stage==='design'&&current.attributesRevision&&current.attributesRevision!==window.ModuleHAttributes?.getState()?.revision)window.ModuleHAttributes?.refresh();}}
    syncSelection();
  },200);
})();
