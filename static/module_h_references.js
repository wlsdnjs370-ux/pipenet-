/* Pipe -> native DXF text references. All writes use the shared undoable editor. */
(()=>{
  'use strict';
  const $=id=>document.getElementById(id),s=()=>window.__mf,data=()=>window.ModuleHAttributes?.getState();
  const view=()=>window.ModuleHView,busy=()=>!$('busy').classList.contains('hidden');
  const picker=document.createElement('canvas');picker.id='h-reference-picker';picker.hidden=true;$('stage').append(picker);
  const prompt=document.createElement('div');prompt.id='h-reference-prompt';prompt.hidden=true;
  prompt.innerHTML='<span role="status"></span><button type="button">취소 · Esc</button>';$('stage').append(prompt);
  const pc=picker.getContext('2d');let picking=null,hover=null,down=null,stamp='',version=0,selectedIds=null,saving=false,layoutCache=null;
  window.addEventListener('module-h-attributes-selection',e=>{selectedIds=new Set(e.detail.ids||[]);});
  const screen=p=>p&&[s().toScreenX(p[0]),s().toScreenY(p[1])];
  const available=()=>s()?.stage==='design'&&s()?.slot==='plan'&&!!data()&&window.ModuleHAttributes.getMode()==='inspect';
  const selected=()=>window.moduleFNetworkEditor?.getSelection();
  let annotationRows=null,annotationIndex=new Map();
  function byIdentity(id){const rows=data()?.annotations;if(rows!==annotationRows){annotationRows=rows;annotationIndex=new Map((rows||[]).map(a=>[a.identity,a]));}return annotationIndex.get(id);}
  function sources(record){
    const e=record.evidence||{};
    if(e.reference_annotation)return [e.reference_annotation];
    if(e.manual||e.block_export)return [];
    const rows=[],seen=new Set(),pending=[e],found=new Map();
    while(pending.length){const row=pending.pop();if(!row||seen.has(row))continue;seen.add(row);rows.push(row);
      pending.push(...(row.evidence_chain||[]),...(row.continuous_sources||[]).map(r=>r.bore_provenance||{}));}
    for(const row of rows){
      if(row.reference_annotation){const a=row.reference_annotation;found.set(a.identity,a);continue;}
      if(row.manual||row.block_export||!row.text_mm)continue;
      const a=byIdentity(row.annotation_id)||
        (row.text_xy_mm&&row.text_raw?{identity:row.annotation_id,x:row.text_xy_mm[0],y:row.text_xy_mm[1],raw_text:row.text_raw,nominal_mm:row.text_mm}:null);
      if(a?.raw_text&&[a.x,a.y].every(Number.isFinite))found.set(a.identity||JSON.stringify([a.x,a.y,a.raw_text]),a);
    }
    return [...found.values()];
  }
  function anchor(record){
    const segments=view().geometry(record).segments||[];
    const longest=segments.slice().sort((a,b)=>Math.hypot(b[1][0]-b[0][0],b[1][1]-b[0][1])-Math.hypot(a[1][0]-a[0][0],a[1][1]-a[0][1]))[0];
    return longest&&[(longest[0][0]+longest[1][0])/2,(longest[0][1]+longest[1][1])/2];
  }
  function paint(ctx,records,box){
    const sel=selected(),all=$('ha-show').checked;
    const picked=r=>selectedIds?selectedIds.has(r.id):sel?.kind==='pipe'&&String(r.label)===String(sel.label);
    const rows=records.filter(r=>r.kind==='pipe'&&r.editable&&(all||picked(r)));
    rows.sort((a,b)=>Number(picked(a))-Number(picked(b)));
    ctx.save();const labels=new Map();
    for(const r of rows){const a=screen(anchor(r));if(!a)continue;
      const hot=picked(r);
      for(const source of sources(r)){const key=source.identity||JSON.stringify([source.x,source.y,source.raw_text]);
        if(!labels.has(key))labels.set(key,view().annotationLayout(source,ctx));const l=labels.get(key);if(!l)continue;
        const [x,y,w,h]=l.rect,c=[x+w/2,y+h/2],dx=a[0]-c[0],dy=a[1]-c[1];
        const t=1/Math.max(Math.abs(dx)/(w/2),Math.abs(dy)/(h/2),1),b=[c[0]+dx*t,c[1]+dy*t];
        if(Math.max(a[0],b[0])<0||Math.min(a[0],b[0])>box.width||Math.max(a[1],b[1])<0||Math.min(a[1],b[1])>box.height)continue;
        // Screen-pixel widths keep evidence readable without looking like a pipe.
        ctx.strokeStyle=hot?'#ff3b45':'#b4cbd8';ctx.globalAlpha=hot?1:.8;ctx.lineWidth=hot?2.2:1.5;ctx.setLineDash([6,6]);
        ctx.beginPath();ctx.moveTo(...a);ctx.lineTo(...b);ctx.stroke();
        view().annotationBorder(ctx,l);
      }
    }
    ctx.restore();
  }
  function message(text){prompt.querySelector('span').textContent=text;}
  function cancel(){picking=null;hover=null;picker.hidden=prompt.hidden=true;stamp='';version++;}
  prompt.querySelector('button').onclick=cancel;
  function begin(label){
    if(!available()||busy()||saving)return;
    const record=data().records.find(r=>r.kind==='pipe'&&String(r.label)===String(label));
    if(!record)return;
    picking={sid:s().sid,label:String(label),revision:data().revision};hover=null;stamp='';version++;
    $('ha-original').checked=true;$('ha-original').dispatchEvent(new Event('change'));
    picker.hidden=prompt.hidden=false;message(`배관 ${label} · 참조할 원본 관경 텍스트를 클릭하세요`);drawPicker();
  }
  function layouts(){
    const key=[s().sid,data()?.revision,s().view.scale,s().view.ox,s().view.oy,view().projectionKey(),picker.width,picker.height].join('|');
    if(layoutCache?.key===key)return layoutCache.rows;
    const seen=new Set(),rows=(data()?.annotations||[]).flatMap(a=>{
      if(!a.raw_text||![a.x,a.y].every(Number.isFinite))return [];
      const key=JSON.stringify([a.x,a.y,a.raw_text,a.nominal_mm]);if(seen.has(key))return [];seen.add(key);
      const l=view().annotationLayout(a,pc);return l?[l]:[];
    });
    layoutCache={key,rows};return rows;
  }
  function hit(p){
    const matches=layouts().map(l=>{const [x,y,w,h]=l.rect;
      const d=Math.hypot(Math.max(0,x-p[0],p[0]-x-w),Math.max(0,y-p[1],p[1]-y-h));
      return {l,d};}).filter(x=>x.d<8).sort((a,b)=>a.d-b.d);
    // Overlapping different values require an explicit unambiguous click.
    if(matches[1]&&Math.abs(matches[0].d-matches[1].d)<1&&matches[0].l.a.nominal_mm!==matches[1].l.a.nominal_mm)return null;
    return matches[0]?.l.a||null;
  }
  function drawPicker(){
    if(!picking)return;
    if(!available()||picking.sid!==s().sid){cancel();return;}
    const box=$('cv').getBoundingClientRect(),dpr=devicePixelRatio||1;
    const key=[box.width,box.height,dpr,s().view.scale,s().view.ox,s().view.oy,view().projectionKey(),data()?.revision,hover].join('|');
    if(key===stamp)return;stamp=key;picker.width=Math.round(box.width*dpr);picker.height=Math.round(box.height*dpr);pc.setTransform(dpr,0,0,dpr,0,0);
    for(const l of layouts()){
      const [x,y,w,h]=l.rect;if(x+w<0||x>box.width||y+h<0||y>box.height)continue;
      pc.save();
      pc.strokeStyle=l.a.identity===hover?'#ff3b45':'#5bc8df';pc.lineWidth=l.a.identity===hover?2:1;pc.setLineDash([3,3]);
      view().annotationBorder(pc,l);pc.setLineDash([]);
      pc.fillStyle='#f3f4f6';pc.strokeStyle='#05090c';pc.lineWidth=2;
      view().annotationText(pc,l);pc.restore();
    }
  }
  function pointer(e){const b=picker.getBoundingClientRect();return [e.clientX-b.left,e.clientY-b.top];}
  picker.onpointerdown=e=>{down=pointer(e);};
  picker.onpointermove=e=>{const a=hit(pointer(e));if(hover!==a?.identity){hover=a?.identity;stamp='';drawPicker();}};
  picker.onclick=async e=>{
    if(!picking||saving||busy()||e.button||e.shiftKey||down&&Math.hypot(...pointer(e).map((v,i)=>v-down[i]))>4)return;
    e.preventDefault();e.stopImmediatePropagation();const a=hit(pointer(e));
    if(!a){message('원본 관경 텍스트를 정확히 클릭하세요 · 겹치면 확대 후 선택');return;}
    if(data().revision!==picking.revision){message('배관망이 변경되었습니다. 취소 후 배관을 다시 선택하세요.');return;}
    const target=picking.label,operation=picking;saving=true;message(`원문 “${a.raw_text}” 저장 중`);
    try{
      const ok=await window.moduleFNetworkEditor.applyReference({op:'pipe_reference',target,annotation_id:a.identity});
      if(!ok){message($('ne-message').textContent||'참조를 변경하지 못했습니다. 다시 확인하세요.');return;}
      if(picking===operation)cancel();
    }catch(error){message(error.message);}finally{saving=false;}
  };
  picker.onwheel=e=>{e.preventDefault();$('cv').dispatchEvent(new WheelEvent('wheel',{deltaY:e.deltaY,deltaX:e.deltaX,clientX:e.clientX,clientY:e.clientY,bubbles:true,cancelable:true}));};
  picker.oncontextmenu=e=>{e.preventDefault();cancel();};
  window.addEventListener('keydown',e=>{if(picking&&e.key==='Escape'){e.preventDefault();e.stopImmediatePropagation();cancel();}},true);
  window.addEventListener('module-h-attributes-mode',cancel);
  setInterval(drawPicker,100);
  window.ModuleHReferences={available,begin,isPicking:()=>!!picking,paint,sources,version:()=>version};
})();
