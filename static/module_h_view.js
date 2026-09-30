/* H view controls: original DXF underlay, never inferred network geometry. */
(() => {
  'use strict';
  const $=id=>document.getElementById(id), state=()=>window.__mf;
  const paths=new WeakMap(), frames=new WeakMap(), views=new WeakMap();
  const planViews=new WeakMap();
  const markIds=['dg-mk-dry','dg-mk-unatt','dg-mk-unpicked','dg-mk-swap'];
  let switching=false,viewVersion=0;
  let sourceTexts=[],sourceSid=null,textVersion=0,textStamp='';
  const textCanvas=document.createElement('canvas');textCanvas.id='h-source-text-canvas';
  textCanvas.hidden=true;textCanvas.setAttribute('aria-hidden','true');
  $('stage').append(textCanvas);const textContext=textCanvas.getContext('2d');
  function originalTexts(annotations) {
    sourceSid=state()?.sid;const seen=new Set();
    // Raw native DXF only: never fabricate a suffix, inner bore, or missing text.
    sourceTexts=(annotations||[]).filter(a=>{
      if(!a.raw_text||![a.x,a.y].every(Number.isFinite))return false;
      const key=JSON.stringify([a.x,a.y,a.raw_text,a.rotation_deg,a.layer]);
      if(seen.has(key))return false;seen.add(key);return true;
    });textVersion++;textStamp='';
  }
  window.addEventListener('module-h-diameter-progress',e=>{
    if(e.detail.data?.reset)originalTexts(e.detail.data.annotations||[]);
  });
  window.addEventListener('module-h-evidence-state',e=>originalTexts(e.detail.annotations));
  const isPlan=()=>$('dg-plan')?.checked||markIds.some(id=>$(id)?.checked);
  const ownsOriginal=()=>state()?.stage==='design'&&state()?.slot==='plan'&&!!$('ha-original');
  const plainNetwork=ownsOriginal;
  function viewIndex(s=state(),view=s?.design?.view) {
    if(!view)return null;
    let index=views.get(view);
    if(!index){index={version:++viewVersion,nodes:new Map((view.nodes||[]).map(n=>[String(n.label),n])),
      pipes:new Map((view.pipes||[]).map(p=>[String(p.label),p]))};views.set(view,index);}
    return index;
  }
  function matrix(s) {
    if(isPlan())return [1,0,0,1,0,0];
    // This is the server's board -> design display transform. Do not estimate
    // an alignment from bounding boxes or confuse display units with lengths.
    const u=s.design?.view?.underlay;
    if(!u)return null;
    const dz=(u.e-u.e_ref)*u.lift;
    const m=u.iso?[u.k*u.cos30,u.k*u.sin30,-u.k*u.cos30,u.k*u.sin30,
      (u.tx-u.ty)*u.cos30,(u.tx+u.ty)*u.sin30+dz]:[u.k,0,0,u.k,u.tx,u.ty];
    return m.every(Number.isFinite)?m:null;
  }
  function projectPoint(point,s=state()) {
    const m=matrix(s);if(!point||!m)return null;
    const [x,y]=point;return [m[0]*x+m[2]*y+m[4],m[1]*x+m[3]*y+m[5]];
  }
  function projectSegments(segments,s=state()) {
    const m=matrix(s);if(!m)return [];
    return (segments||[]).map(segment=>segment.map(([x,y])=>[m[0]*x+m[2]*y+m[4],m[1]*x+m[3]*y+m[5]]));
  }
  function geometry(record,s=state()) {
    if(isPlan()){
      const index=viewIndex(s,planView(s)),at=index?.nodes||new Map();
      if(record.kind==='head'){const n=at.get(String(record.node_label));return {point:n?[n.x,n.y]:record.xy?.[0],segments:[]};}
      const p=record.kind==='pipe'&&index?.pipes.get(String(record.label));
      if(p){const a=at.get(String(p.a)),b=at.get(String(p.b));return {segments:a&&b?[[[a.x,a.y],[b.x,b.y]]]:[]};}
      return {segments:record.xy||[]};
    }
    const index=viewIndex(s);
    if(record.kind==='head'){
      const n=index?.nodes.get(String(record.node_label));
      return {point:n?[n.x,n.y]:projectPoint(record.xy?.[0],s),segments:[]};
    }
    // Calculation pipes/risers follow the exact preview node positions, including
    // generated vertical links. Original references use the CAD underlay plane.
    const p=record.kind==='pipe'&&index?.pipes.get(String(record.label));
    if(p){const a=index.nodes.get(String(p.a)),b=index.nodes.get(String(p.b));
      return {segments:a&&b?[[[a.x,a.y],[b.x,b.y]]]:[]};}
    return {segments:projectSegments(record.xy,s)};
  }
  function projectionKey(){return isPlan()?'plan':`${viewIndex()?.version}:${matrix(state())?.join(',')}`;}
  // Shared screen-space layout. CAD position/rotation remain unmodified.
  function annotationLayout(a,ctx){
    const s=state(),p=projectPoint([a.x,a.y],s);if(!p||!a.raw_text)return null;
    const q=projectPoint([a.x+1,a.y],s),xy=[s.toScreenX(p[0]),s.toScreenY(p[1])];
    const scale=Math.hypot(s.toScreenX(q[0])-xy[0],s.toScreenY(q[1])-xy[1]);
    const size=Math.max(11,Math.min(28,(a.height_mm||100)*scale)),lines=String(a.raw_text).split(/\r?\n/);
    const font=`${size}px "Noto Sans KR", sans-serif`;ctx.save();ctx.font=font;ctx.textBaseline='alphabetic';ctx.textAlign='left';
    const metrics=lines.map(t=>ctx.measureText(t));ctx.restore();
    const left=Math.min(0,...metrics.map(m=>-m.actualBoundingBoxLeft||0));
    const right=Math.max(...metrics.map(m=>Math.max(m.width,m.actualBoundingBoxRight||0)));
    const top=-Math.max(size*.8,...metrics.map(m=>m.actualBoundingBoxAscent||0));
    const bottom=(lines.length-1)*size*1.25+Math.max(size*.2,...metrics.map(m=>m.actualBoundingBoxDescent||0));
    return {a,p:xy,size,font,lines,rect:[xy[0]+left-6,xy[1]+top-6,right-left+12,bottom-top+12]};
  }
  function annotationBorder(ctx,l){ctx.beginPath();ctx.roundRect(...l.rect,6);ctx.stroke();}
  function annotationText(ctx,l){
    ctx.save();ctx.translate(...l.p);ctx.font=l.font;ctx.textBaseline='alphabetic';ctx.textAlign='left';
    l.lines.forEach((line,i)=>{ctx.strokeText(line,0,i*l.size*1.25);ctx.fillText(line,0,i*l.size*1.25);});ctx.restore();
  }
  function planView(s=state()){
    const v=s?.design?.view;
    if(!ownsOriginal()||!isPlan()||!v?.nodes?.length||v.nodes.some(n=>!n.plan_xy))return null;
    if(!planViews.has(v))planViews.set(v,{...v,nodes:v.nodes.map(n=>({...n,x:n.plan_xy[0],y:n.plan_xy[1]}))});
    return planViews.get(v);
  }
  function drawPlanNetwork(ctx,s,sx,sy){
    const v=planView(s);if(!v)return false;
    const at=new Map(v.nodes.map(n=>[String(n.label),n]));ctx.save();ctx.lineCap='round';
    ctx.strokeStyle='#ffffff';ctx.lineWidth=4;ctx.setLineDash([]);
    for(const p of v.pipes){const a=at.get(String(p.a)),b=at.get(String(p.b));if(!a||!b)continue;
      ctx.beginPath();ctx.moveTo(sx(a.x),sy(a.y));ctx.lineTo(sx(b.x),sy(b.y));ctx.stroke();}
    ctx.strokeStyle='#ff3b3b';ctx.lineWidth=2;ctx.setLineDash([9,5]);ctx.beginPath();
    (v.worst_path||[]).forEach((id,i)=>{const n=at.get(String(id));if(n)ctx[i?'lineTo':'moveTo'](sx(n.x),sy(n.y));});ctx.stroke();ctx.setLineDash([]);
    for(const n of v.nodes){if(!n.head&&!n.input&&!n.valve)continue;
      ctx.strokeStyle=n.worst_head?'#ff3b3b':'#ffffff';ctx.lineWidth=2;ctx.beginPath();
      ctx.arc(sx(n.x),sy(n.y),n.head?7:4,0,Math.PI*2);ctx.stroke();}
    ctx.restore();return true;
  }
  function sync() {
    if($('h-source-view'))$('h-source-view').hidden=!(state()?.world?.source_view_bounds&&state()?.stage==='sub');
    if(!$('ha-iso'))return;
    const s=state(),busy=!$('busy').classList.contains('hidden');
    $('ha-iso').checked=!isPlan()&&!!$('dg-iso').checked;
    $('ha-iso').disabled=switching||busy||!s?.design?.view?.nodes?.length;
    $('ha-iso').title=$('ha-iso').disabled?'배관 속성 처리가 끝나면 사용할 수 있습니다.':'선택한 배관망을 등각투상으로 보기';
    $('ha-original').disabled=!s?.world||!matrix(s);
    $('ha-original').title=$('ha-original').disabled?'원본 도면 또는 좌표 변환이 없습니다.':'원본 DXF와 흰 점선 도면 테두리 표시';
    $('ha-table').checked=!!window.ModuleHProperties?.isVisible();
    $('ha-table').disabled=!s?.design?.tables||!window.ModuleHProperties||window.ModuleHAttributes?.getMode()!=='inspect';
  }
  function init() {
    $('ha-original').onchange=()=>window.dispatchEvent(new Event('resize'));
    $('ha-table').onchange=()=>{
      if($('ha-table').checked)window.ModuleHProperties?.open({all:true});
      else window.ModuleHProperties?.close();
      sync();
    };
    $('ha-iso').onchange=async()=>{
      const on=$('ha-iso').checked,s=state(),sel=s?.design?.sel;
      const previous=Object.fromEntries(['dg-plan','dg-iso',...markIds].map(id=>[id,$(id).checked]));
      switching=true;
      try{
        if(on){
          markIds.forEach(id=>$(id).checked=false);
          if(!$('dg-iso').checked){
            $('dg-iso').checked=true;
            if(await $('dg-iso').onchange()===false)throw Error('View preview failed');
          }
        }
        $('dg-plan').checked=!on;
        await $('dg-plan').onchange();
        // Preview may replace the view object. A view toggle is not deselection.
        if(sel&&s.design)window.moduleFNetworkEditor?.select(sel.kind,sel.label);
      }catch(error){
        // The native preview reports its error. Keep the previous view instead
        // of labelling an old flat preview as an isometric one after failure.
        Object.entries(previous).forEach(([id,value])=>$(id).checked=value);
        await $('dg-plan').onchange();
      }finally{switching=false;sync();}
    };
    sync();
  }
  function bundlePath(bundle) {
    let cached=paths.get(bundle);if(cached)return cached;
    const path=new Path2D(),segs=bundle.segs||[],circles=bundle.circles||[],arcs=bundle.arcs||[];
    for(let i=0;i+3<segs.length;i+=4){path.moveTo(segs[i],segs[i+1]);path.lineTo(segs[i+2],segs[i+3]);}
    for(let i=0;i+2<circles.length;i+=3){const [x,y,r]=circles.slice(i,i+3);if(r<=0)continue;
      path.moveTo(x+r,y);path.arc(x,y,r,0,Math.PI*2);}
    for(let i=0;i+4<arcs.length;i+=5){const [x,y,r,start,sweep]=arcs.slice(i,i+5);if(r<=0)continue;
      const a=start*Math.PI/180,b=(start+sweep)*Math.PI/180;
      path.moveTo(x+r*Math.cos(a),y+r*Math.sin(a));path.arc(x,y,r,a,b,sweep<0);}
    cached={path};paths.set(bundle,cached);return cached;
  }
  function project(cached,m,signature) {
    if(cached.signature!==signature){cached.projected=new Path2D();cached.projected.addPath(cached.path,new DOMMatrix(m));cached.signature=signature;}
    return cached.projected;
  }
  function drawOriginal(ctx,s,sx,sy) {
    if(!ownsOriginal()||!$('ha-original').checked||!s.world)return;
    const m=matrix(s),scale=s.view.scale;if(!m||!(scale>0))return;
    const signature=m.join(',');
    ctx.save();ctx.translate(sx(0),sy(0));ctx.scale(scale,-scale);
    ctx.lineWidth=1/scale;ctx.globalAlpha=.22;ctx.setLineDash([]);
    for(const bundle of s.world.bundles||[]){
      if(s.hidden?.has(bundle.id))continue;
      ctx.strokeStyle=bundle.css||'#7aa2ff';ctx.stroke(project(bundlePath(bundle),m,signature));
    }
    const b=s.world.bounds;
    if(b&&[b.minx,b.miny,b.maxx,b.maxy].every(Number.isFinite)){
      let frame=frames.get(b);
      if(!frame){const path=new Path2D();path.rect(b.minx,b.miny,b.maxx-b.minx,b.maxy-b.miny);frame={path};frames.set(b,frame);}
      ctx.globalAlpha=.75;ctx.strokeStyle='#ffffff';ctx.setLineDash([6/scale,4/scale]);
      ctx.stroke(project(frame,m,signature));
    }
    ctx.restore();
  }
  function drawOriginalTexts() {
    const s=state();if(!s)return;
    if(sourceSid!==s.sid){sourceTexts=[];sourceSid=s.sid;textVersion++;}
    const active=ownsOriginal()&&$('ha-original').checked&&!!matrix(s)&&sourceTexts.length;
    textCanvas.hidden=!active;if(!active){textStamp='';return;}
    const b=$('cv').getBoundingClientRect(),dpr=devicePixelRatio||1;
    const stamp=[textVersion,projectionKey(),s.view.scale,s.view.ox,s.view.oy,b.width,b.height,dpr].join('|');
    if(stamp===textStamp)return;textStamp=stamp;
    textCanvas.width=Math.round(b.width*dpr);textCanvas.height=Math.round(b.height*dpr);
    const ctx=textContext;ctx.setTransform(dpr,0,0,dpr,0,0);
    ctx.fillStyle='#f3f4f6';ctx.strokeStyle='#05090c';ctx.lineWidth=2;ctx.lineJoin='round';
    for(const a of sourceTexts){
      const l=annotationLayout(a,ctx);if(!l)continue;const [x,y,w,h]=l.rect;
      if(x+w<0||y+h<0||x>b.width||y>b.height)continue;
      annotationText(ctx,l);
    }
  }
  window.addEventListener('resize',drawOriginalTexts);
  setInterval(drawOriginalTexts,100);
  window.ModuleHView={init,sync,ownsOriginal,drawOriginal,plainNetwork,isPlan,projectPoint,projectSegments,geometry,projectionKey,planView,drawPlanNetwork,annotationLayout,annotationBorder,annotationText};
})();
