/* Inspectable representative -> targets, during computation and after commit. */
(() => {
  'use strict';
  const $=id=>document.getElementById(id),s=()=>window.__mf;
  const pane=document.createElement('aside');pane.id='h-evidence';pane.hidden=true;
  pane.innerHTML='<header><strong>대표 표기 → 적용 배관</strong><button id="he-fold">접기</button></header><div id="he-status" role="status"></div><div id="he-body"><input id="he-search" placeholder="배관 번호 · 관경 검색" aria-label="관경 대응 검색"><div class="he-scroll"><table><thead><tr><th>대표</th><th>관경 (mm)</th><th>적용</th><th>근거</th></tr></thead><tbody id="he-rows"></tbody></table></div><div id="he-targets"></div></div>';
  document.body.append(pane);
  const canvas=document.createElement('canvas');canvas.id='h-evidence-canvas';canvas.hidden=true;$('stage').append(canvas);
  const ctx=canvas.getContext('2d');let records=[],groups=[],byEdge=new Map(),selected=null,targetId=null,dirty=false,lastView='',lastSid=null;
  const live=new Map();let livePhase='',timer=null;
  const edgeKey=e=>e?.slice().sort((a,b)=>a-b).join(':');
  const evidence=r=>{
    const root=r.evidence||{};
    const active=e=>{const a=e.reference_annotation;
      if(a)return {...e,text_mm:a.nominal_mm,annotation_id:a.identity,reference_annotation:a,block_export:false,definition_source:'user_reference'};
      return e.manual?null:e;};
    // A current explicit edit owns the whole run. Historical automatic labels
    // remain audit evidence, not additional assignments for this same pipe.
    if(root.reference_annotation||(root.manual&&!root.serial_manual))return [active(root)].filter(Boolean);
    return (root.continuous_sources?.length?root.continuous_sources.map(x=>x.bore_provenance||{}):[root]).map(active).filter(Boolean);
  };
  const groupKey=(e,r)=>e.representative_nodes?'path:'+e.representative_nodes.join(':'):'text:'+(e.annotation_id??edgeKey(e.source_edge||e.owner_edge||r.edges?.[0]));
  function groupRecords(){
    const map=new Map();byEdge=new Map();records.forEach(r=>(r.edges||[]).forEach(e=>byEdge.set(edgeKey(e),r)));
    for(const r of records){if(r.kind!=='pipe')continue;
      const sources=evidence(r).filter(e=>e.text_mm&&!e.block_export&&e.text_mm===r.nominal_mm);
      const owners=[...new Set(sources.map(e=>groupKey(e,r)))].sort();
      // One target belongs to one row, even if a canonical span retains several
      // same-size source annotations. Outside-scenario records are origins only.
      const key=owners.length===1?owners[0]:'combined:'+owners.join('|');
      for(const e of sources){
        if(!map.has(key))map.set(key,{key,origins:new Set(),targets:new Map(),sizes:new Set(),methods:new Set()});
        const g=map.get(key),sourceEdge=e.source_edge||e.owner_edge;
        const origin=sourceEdge?byEdge.get(edgeKey(sourceEdge)):e.definition_source==='drawing_direct'?r:null;
        if(origin)g.origins.add(origin);g.targets.set(r.id,r);g.sizes.add(e.text_mm);g.methods.add(e.definition_source);
      }
    }
    groups=[...map.values()].sort((a,b)=>a.key.localeCompare(b.key,'en',{numeric:true}));
    groups.forEach((g,i)=>g.label='A'+String(i+1).padStart(3,'0'));
    if(!groups.some(g=>g.key===selected))selected=groups.find(g=>g.methods.has('drawing_repeat'))?.key||groups[0]?.key;
    if(!groups.find(g=>g.key===selected)?.targets.has(targetId))targetId=null;
    render();dirty=true;
  }
  function choose(g,id){selected=g.key;targetId=id||[...g.targets.values()].find(r=>!g.origins.has(r))?.id||g.targets.keys().next().value;dirty=true;
    // A group is an inspection choice, never an implicit bulk-edit selection.
    window.dispatchEvent(new CustomEvent('module-h-evidence-select',{detail:{ids:targetId?[targetId]:[],keepView:true}}));render();
  }
  function render(){
    const oldList=$('he-targets').querySelector('.he-target-list');
    const oldScroll=oldList&&oldList.dataset.group===selected?oldList.scrollTop:0;
    const filter=$('he-search').value.toLowerCase();$('he-rows').replaceChildren();
    for(const g of groups){
      const sizes=[...g.sizes].sort((a,b)=>b-a).join(' → '),targetList=[...g.targets.values()];
      if(filter&&!`${g.label} ${sizes} ${targetList.map(r=>r.label).join(' ')}`.toLowerCase().includes(filter))continue;
      const tr=document.createElement('tr');tr.tabIndex=0;tr.dataset.group=g.key;tr.classList.toggle('he-selected',g.key===selected);
      const method=g.methods.has('user_reference')?'직접 참조':g.methods.has('drawing_repeat')?'반복 구조':g.methods.has('drawing_run')?'연속 관로':'직접 표기';
      for(const value of [g.label,sizes,targetList.length,method]){const td=document.createElement('td');td.textContent=value;tr.append(td);}
      tr.onclick=()=>choose(g);tr.onkeydown=e=>{if(e.key==='Enter')choose(g);};$('he-rows').append(tr);
    }
    const g=groups.find(g=>g.key===selected);$('he-targets').replaceChildren();
    if(g){const title=document.createElement('strong');title.textContent=g.label+' 대표 → 대상';$('he-targets').append(title);
      const list=document.createElement('div');list.className='he-target-list';list.dataset.group=g.key;
      [...g.targets.values()].forEach((r,i)=>{const b=document.createElement('button');b.textContent=`B${i+1} · ${r.label} · ${r.nominal_mm||'—'}A`;b.dataset.id=r.id;
        b.setAttribute('aria-pressed',String(r.id===targetId));b.onclick=()=>choose(g,r.id);list.append(b);});$('he-targets').append(list);
      list.scrollTop=oldScroll;
    }
  }
  function setLive(){records=[...live.values()];groupRecords();$('he-status').textContent=livePhase+' · '+records.length+'구간';timer=null;}
  window.addEventListener('module-h-diameter-progress',event=>{
    lastSid=s()?.sid;
    const d=event.detail.data||{};if(d.reset)live.clear();livePhase=event.detail.phase;
    for(const row of d.snapshot||[d])if(row.edge&&row.xy)live.set(edgeKey(row.edge),{id:'live:'+edgeKey(row.edge),label:row.edge.join('–'),kind:'pipe',edges:[row.edge],xy:[row.xy],nominal_mm:row.evidence?.text_mm,evidence:row.evidence});
    if(!timer)timer=setTimeout(setLive,160);
  });
  window.addEventListener('module-h-evidence-state',event=>{
    lastSid=s()?.sid;
    clearTimeout(timer);timer=null;live.clear();records=event.detail.records;groupRecords();
    const total=records.filter(r=>r.kind==='pipe').length;
    const referenced=new Set(groups.flatMap(g=>[...g.targets.keys()])).size;
    $('he-status').textContent=`참조 ${referenced} · 미참조 ${total-referenced} · 전체 ${total}`;
  });
  window.addEventListener('module-h-attributes-selection',event=>{
    const ids=event.detail.ids||[],id=ids.length===1?ids[0]:null;
    const g=groups.find(g=>g.key===selected&&g.targets.has(id))||groups.find(g=>g.targets.has(id));
    if(g)selected=g.key;else if(ids.length===1)selected=null;
    targetId=g?id:null;dirty=true;render();
  });
  let expandedHeight=null;
  $('he-fold').setAttribute('aria-expanded','true');
  $('he-fold').setAttribute('aria-controls','he-body he-status');
  $('he-search').oninput=render;
  $('he-fold').onclick=()=>{
    const folded=!pane.classList.contains('he-folded');
    if(folded)expandedHeight=pane.getBoundingClientRect().height;
    pane.classList.toggle('he-folded',folded);
    $('he-body').hidden=folded;$('he-status').hidden=folded;
    // Window resizing freezes inline !important geometry: hiding contents is
    // insufficient. Restore the expanded height, including after a folded drag.
    pane.style.setProperty('height',(folded?pane.querySelector('header').getBoundingClientRect().height+2:expandedHeight)+'px','important');
    $('he-fold').textContent=folded?'펼치기':'접기';
    $('he-fold').setAttribute('aria-expanded',String(!folded));
    pane.dispatchEvent(new CustomEvent('h-window-resize'));
  };
  window.ModuleHWindows.register(pane,pane.querySelector('header'));
  function paint(){
    const state=s();if(!state)return;
    if(lastSid!==state.sid){lastSid=state.sid;records=[];groups=[];selected=null;targetId=null;live.clear();dirty=true;}
    const active=state.stage==='design'&&state.slot==='plan';pane.hidden=!active||(!records.length&&!live.size);canvas.hidden=!active;
    // An explicit editing/inspection window takes the foreground. The evidence
    // panel remains draggable and returns when the editor closes.
    if(!pane.dataset.dragged){const editor=$('h-attributes-panel');pane.style.right=editor&&!editor.hidden?'460px':'20px';}
    const box=$('cv').getBoundingClientRect(),view=[state.view?.scale,state.view?.ox,state.view?.oy,box.width,box.height,selected,targetId,JSON.stringify(window.moduleFNetworkEditor?.getSelection()),$('ha-show').checked,canvas.hidden,window.ModuleHView.projectionKey()].join('|');
    if(!dirty&&view===lastView)return;lastView=view;dirty=false;
    const dpr=devicePixelRatio||1;canvas.width=box.width*dpr;canvas.height=box.height*dpr;ctx.setTransform(dpr,0,0,dpr,0,0);if(canvas.hidden)return;
    window.ModuleHReferences?.paint(ctx,records,box);
  }
  window.addEventListener('resize',paint);
  setInterval(paint,150);
})();
