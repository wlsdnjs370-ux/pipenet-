/* Live H canvas observations. Preview geometry never enters __mf or export data. */
(() => {
  "use strict";
  const stage=document.getElementById("stage"), native=document.getElementById("cv");
  const canvas=document.createElement("canvas"), label=document.createElement("div");
  canvas.id="h-progress-canvas";canvas.hidden=true;canvas.setAttribute("aria-hidden","true");
  label.id="h-progress-label";label.hidden=true;label.setAttribute("role","status");
  stage.append(canvas,label);
  const ctx=canvas.getContext("2d"), LIMIT=50000;
  let token=null, cursor=0, timer=null, controller=null, mode=null, phase="", rows=[], roots=[], tip=null;
  let generation=0, frame=0, layers=new Set(), partial=false, width=0, height=0, camera=null, manualCamera=false;
  let bounds=null, drag=null, paintView="", lastCount=0;
  let originalWorld=null, worldReceived=false;
  function clear() {
    rows=[];roots=[];tip=null;layers=new Set();partial=false;bounds=null;camera=null;manualCamera=false;lastCount=0;drag=null;
    canvas.hidden=true;label.hidden=true;document.body.classList.remove("h-progress-world");
  }
  function grow(x,y,r=0) {
    if(!bounds) bounds={minx:x-r,maxx:x+r,miny:y-r,maxy:y+r};
    else {bounds.minx=Math.min(bounds.minx,x-r);bounds.maxx=Math.max(bounds.maxx,x+r);bounds.miny=Math.min(bounds.miny,y-r);bounds.maxy=Math.max(bounds.maxy,y+r);}
  }
  function accept(event) {
    if(event.kind==="diameter") {window.dispatchEvent(new CustomEvent('module-h-diameter-progress',{detail:event}));return;}
    const d=event.data || {};
    if(worldReceived && (event.kind==="world" || (event.kind==="reset" && d.mode==="world")))return;
    if(event.kind==="reset") {clear();mode=d.mode;phase=event.phase;roots=d.roots || [];roots.forEach(p=>grow(...p));return;}
    if(!mode) {mode=event.kind==="graph"?"graph":"world";partial=true;}
    phase=event.phase;
    const items=event.kind==="graph" ? (d.segments || []).map(xy=>({kind:"segs",xy})) : d.items || [];
    for(const item of items) {
      const p=item.xy;if(!Array.isArray(p) || !p.every(Number.isFinite)) continue;
      if(rows.length>=LIMIT) {partial=true;break;}
      rows.push(item);if(item.layer!==undefined)layers.add(item.layer);
      if(item.kind==="segs") {grow(p[0],p[1]);grow(p[2],p[3]);}else grow(p[0],p[1],p[2]);
    }
    if(d.cursor) tip=d.cursor;
    lastCount=d.count || rows.length;
    canvas.dataset.items=String(rows.length);canvas.dataset.mode=mode;
    canvas.hidden=!rows.length;label.hidden=!rows.length;
    document.body.classList.toggle("h-progress-world",mode==="world" && rows.length>0);
    label.textContent=phase+" · "+(mode==="graph" ? lastCount.toLocaleString()+"구간 확인 중" : layers.size+"개 레이어 표시 중")
      +(partial?" · 간략 미리보기":"")+" · 미확정";
    paintView="";queue();
  }
  function view() {
    if(mode==="graph") return window.__mf?.view;
    if(!bounds) return null;
    if(!manualCamera || !camera) {
      const dx=Math.max(bounds.maxx-bounds.minx,1),dy=Math.max(bounds.maxy-bounds.miny,1);
      const scale=Math.min(width*.85/dx,height*.85/dy);
      camera={scale,ox:(bounds.minx+bounds.maxx)/2-width/(2*scale),oy:(bounds.miny+bounds.maxy)/2-height/(2*scale)};
    }
    return camera;
  }
  function paint() {
    frame=0;
    if(canvas.hidden || !token) return;
    if(mode==="world" && window.__mf?.world && window.__mf.world!==originalWorld) {
      worldReceived=true;clear();return; // Native drawing is ready; show the full authoritative view.
    }
    const rect=native.getBoundingClientRect(),dpr=window.devicePixelRatio || 1;
    width=rect.width;height=rect.height;
    const w=Math.round(width*dpr),h=Math.round(height*dpr);
    if(canvas.width!==w || canvas.height!==h) {canvas.width=w;canvas.height=h;paintView="";}
    const v=view();if(!v) return;
    const stamp=[v.scale,v.ox,v.oy,rows.length,mode].join("|");
    if(stamp!==paintView) {
      paintView=stamp;ctx.setTransform(dpr,0,0,dpr,0,0);ctx.clearRect(0,0,width,height);
      const x=n=>(n-v.ox)*v.scale,y=n=>height-(n-v.oy)*v.scale;
      for(const row of rows) {
        const p=row.xy;ctx.strokeStyle=mode==="graph"?"#D9A441":row.css || "#5B8DEF";ctx.lineWidth=mode==="graph"?2:1;
        ctx.beginPath();
        if(row.kind==="segs") {ctx.moveTo(x(p[0]),y(p[1]));ctx.lineTo(x(p[2]),y(p[3]));}
        else if(row.kind==="arcs") {const sa=-p[3]*Math.PI/180;ctx.arc(x(p[0]),y(p[1]),Math.abs(p[2]*v.scale),sa,sa-p[4]*Math.PI/180,true);}
        else ctx.arc(x(p[0]),y(p[1]),Math.abs(p[2]*v.scale),0,Math.PI*2);
        ctx.stroke();
      }
      roots.forEach(p=>{ctx.fillStyle="#5B8DEF";ctx.beginPath();ctx.arc(x(p[0]),y(p[1]),5,0,Math.PI*2);ctx.fill();});
      if(tip) {ctx.strokeStyle="#E8EBEF";ctx.lineWidth=2;ctx.beginPath();ctx.arc(x(tip[0]),y(tip[1]),7,0,Math.PI*2);ctx.stroke();}
    }
    // Camera can move through native handlers without changing preview data.
    frame=requestAnimationFrame(paint);
  }
  function queue() {if(!frame)frame=requestAnimationFrame(paint);}
  async function poll(id,seq) {
    if(id!==token || seq!==generation) return;
    try {
      controller=new AbortController();
      const response=await fetch(`/api/module-h/progress?operation=${encodeURIComponent(id)}&after=${cursor}`,{signal:controller.signal,cache:"no-store"});
      if(!response.ok) throw Error("progress unavailable");
      const data=await response.json();
      if(id!==token || seq!==generation) return;
      if(data.gap) {clear();partial=true;}
      for(const event of data.events || []) accept(event);
      cursor=data.cursor ?? cursor;
      if(data.diameter_snapshot?.length||Array.isArray(data.annotation_snapshot)){
        accept({kind:'diameter',phase:'관경 대응 수신',data:{reset:true,snapshot:data.diameter_snapshot||[],annotations:data.annotation_snapshot||[]}});
        cursor=data.snapshot_cursor;
      }
    } catch(error) {
      if(error.name!=="AbortError" && token===id && seq===generation) {label.textContent="화면 진행 수신 지연";label.hidden=false;}
    } finally {
      if(token===id && seq===generation)timer=setTimeout(()=>poll(id,seq),160);
    }
  }
  function stop() {generation++;token=null;controller?.abort();clearTimeout(timer);cancelAnimationFrame(frame);frame=0;clear();}
  function start(id) {stop();if(!id)return;token=id;cursor=0;mode=null;originalWorld=window.__mf?.world;worldReceived=false;poll(id,generation);}
  // During initial decoding the native canvas has no new world/camera yet.
  native.addEventListener("wheel",event=>{
    if(mode!=="world" || canvas.hidden || !camera) return;
    event.preventDefault();event.stopImmediatePropagation();manualCamera=true;
    const rect=native.getBoundingClientRect(),px=event.clientX-rect.left,py=event.clientY-rect.top;
    const wx=camera.ox+px/camera.scale,wy=camera.oy+(height-py)/camera.scale;
    camera.scale*=Math.exp(-event.deltaY*.001);camera.ox=wx-px/camera.scale;camera.oy=wy-(height-py)/camera.scale;
  },{capture:true,passive:false});
  native.addEventListener("mousedown",event=>{
    if(mode!=="world" || canvas.hidden || !camera || ![1,2].includes(event.button))return;
    event.preventDefault();event.stopImmediatePropagation();manualCamera=true;drag={x:event.clientX,y:event.clientY,ox:camera.ox,oy:camera.oy};
  },true);
  window.addEventListener("mousemove",event=>{if(drag && camera){camera.ox=drag.ox-(event.clientX-drag.x)/camera.scale;camera.oy=drag.oy+(event.clientY-drag.y)/camera.scale;}});
  window.addEventListener("mouseup",()=>{drag=null;});
  window.addEventListener("pagehide",stop);
  window.ModuleHProgress={start,stop};
})();
