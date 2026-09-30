/* Link original CAD geometry to the live bottom grid without a separate viewer. */
(() => {
  'use strict';
  const $=id=>document.getElementById(id),s=()=>window.__mf;
  const records=()=>window.ModuleHAttributes?.getState()?.records||[];
  // Hit original source geometry, not projected IDs. Only inspection mode owns
  // this handler; explicit lasso, topology edits, panning and busy work keep theirs.
  const cv=$('cv');let down=null,moved=false;
  cv.addEventListener('pointerdown',e=>{down=[e.clientX,e.clientY];moved=false;},true);
  cv.addEventListener('pointermove',e=>{if(down&&Math.hypot(e.clientX-down[0],e.clientY-down[1])>4)moved=true;},true);
  function distance(p,a,b){const dx=b[0]-a[0],dy=b[1]-a[1],t=Math.max(0,Math.min(1,((p[0]-a[0])*dx+(p[1]-a[1])*dy)/(dx*dx+dy*dy||1)));return Math.hypot(p[0]-a[0]-t*dx,p[1]-a[1]-t*dy);}
  const active=()=>s()?.stage==='design'&&s()?.slot==='plan'&&window.ModuleHView.isPlan()&&window.ModuleHAttributes?.getMode()==='inspect';
  function inspect(x,y,{additive=false}={}){
    const p=[s().toScreenX(x),s().toScreenY(y)],screen=p=>[s().toScreenX(p[0]),s().toScreenY(p[1])];
    const v=s().stage==='merge'?s().mergeView:window.ModuleHView.planView()||s().design?.view;
    let best=null,d=9;const choose=(kind,label)=>window.ModuleHProperties.select(kind,String(label),{additive});
    const at=new Map((v?.nodes||[]).map(n=>[String(n.label),n]));
    const data=window.ModuleHProperties?.getData()||window.moduleFNetworkEditor?.getState();
    // Plan riser ends overlap in XY. The displayed head must win over its base
    // fitting/node; do not accidentally select the lower, non-head endpoint.
    for(const n of v?.nodes||[])if(n.head&&distance(p,screen([n.x,n.y]),screen([n.x,n.y]))<7){choose('node',n.label);return;}
    // Fittings retain their own property identity even when drawn at a node.
    for(const f of data?.property_fittings||[]){
      const n=f.node!=null&&at.get(f.node),pipe=(v?.pipes||[]).find(p=>String(p.label)===f.pipe);
      const a=pipe&&at.get(String(pipe.a)),b=pipe&&at.get(String(pipe.b)),t=Number(f.position??.5);
      const point=n?[n.x,n.y]:a&&b?[a.x+(b.x-a.x)*t,a.y+(b.y-a.y)*t]:null;
      if(point&&distance(p,screen(point),screen(point))<7){choose('fitting',f.label);return;}
    }
    for(const n of v?.nodes||[]){const a=screen([n.x,n.y]),dist=Math.hypot(a[0]-p[0],a[1]-p[1]);
      if(dist<7){choose('node',n.label);return;}}
    for(const r of v?.pipes||[]){const a=at.get(String(r.a)),b=at.get(String(r.b));if(!a||!b)continue;
      const dist=distance(p,screen([a.x,a.y]),screen([b.x,b.y]));if(dist<d){d=dist;best=r;}}
    if(best)choose('pipe',best.label);
  }
  window.ModuleHSourceSelection={active,inspect};
  cv.addEventListener('click',e=>{
    const allowed=s()?.stage==='merge'||s()?.stage==='design'&&s()?.slot==='plan'&&window.ModuleHAttributes?.getMode()==='inspect';
    if(e.button||moved||!allowed||cv.classList.contains('zoom-window-armed')||
      !$('busy').classList.contains('hidden')||s().opArm||
      $('ne-panel').getClientRects().length)return;
    if(window.ModuleHReferences?.isPicking())return;
    if(window.moduleFNetworkEditor?.isPicking()){
      e.preventDefault();e.stopImmediatePropagation();
      const b=cv.getBoundingClientRect(),x=(e.clientX-b.left)/s().view.scale+s().view.ox,y=(b.height-e.clientY+b.top)/s().view.scale+s().view.oy;
      window.moduleFNetworkEditor.canvasClick(x,y,9/s().view.scale);return;
    }
    const box=cv.getBoundingClientRect(),p=[e.clientX-box.left,e.clientY-box.top];
    // A legacy join/delete mode can survive the preceding process stage. It
    // must not turn a property inspection click into a topology mutation.
    e.preventDefault();e.stopImmediatePropagation();
    inspect(p[0]/s().view.scale+s().view.ox,(box.height-p[1])/s().view.scale+s().view.oy,{additive:e.shiftKey});
  },true);
  const overlay=document.createElement('canvas');overlay.id='h-multiselect-canvas';overlay.setAttribute('aria-hidden','true');$('stage').append(overlay);
  let stamp='';
  setInterval(()=>{
    const selections=window.ModuleHProperties?.getSelections()||[],state=s();
    const active=state?.stage==='merge'||state?.stage==='design'&&state?.slot==='plan'&&window.ModuleHAttributes?.getMode()==='inspect';
    overlay.hidden=!active||!selections.length;if(overlay.hidden){stamp='';return;}
    const v=state.stage==='merge'?state.mergeView:window.ModuleHView.planView()||state.design?.view;
    const b=cv.getBoundingClientRect(),data=window.moduleFNetworkEditor?.getState(),dpr=devicePixelRatio||1;
    const key=JSON.stringify([selections,state.view,b.width,b.height,dpr,data?.revision,window.ModuleHView.isPlan()]);if(stamp===key)return;stamp=key;
    overlay.width=Math.round(b.width*dpr);overlay.height=Math.round(b.height*dpr);const ctx=overlay.getContext('2d');ctx.setTransform(dpr,0,0,dpr,0,0);
    const nodes=new Map((v?.nodes||[]).map(n=>[String(n.label),n])),pipes=new Map((v?.pipes||[]).map(p=>[String(p.label),p]));
    ctx.strokeStyle='#62b5ff';ctx.lineWidth=4;ctx.shadowColor='#399dff';ctx.shadowBlur=6;
    const circle=n=>{if(!n)return;ctx.beginPath();ctx.arc(state.toScreenX(n.x),state.toScreenY(n.y),10,0,Math.PI*2);ctx.stroke();};
    for(const sel of selections){
      if(sel.kind==='node')circle(nodes.get(sel.label));
      else if(sel.kind==='fitting'){
        const f=data?.property_fittings?.find(r=>r.label===sel.label),p=f&&pipes.get(f.pipe),a=p&&nodes.get(String(p.a)),b=p&&nodes.get(String(p.b));
        circle(f?.node!=null?nodes.get(f.node):a&&b?{x:a.x+(b.x-a.x)*f.position,y:a.y+(b.y-a.y)*f.position}:null);
      }else{const p=pipes.get(sel.label),a=p&&nodes.get(String(p.a)),b=p&&nodes.get(String(p.b));if(!a||!b)continue;
        ctx.beginPath();ctx.moveTo(state.toScreenX(a.x),state.toScreenY(a.y));ctx.lineTo(state.toScreenX(b.x),state.toScreenY(b.y));ctx.stroke();}
    }
  },100);
})();
