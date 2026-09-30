/* Read-only property cards and always-on fitting glyphs. Types/losses are server-owned. */
window.createModuleFInspection = function (h) {
  "use strict";
  const {state:S,ctx,sx,sy,esc,kv,grp,insLink} = h;
  const view = () => h.scope === 'merge' ? S.mergeView : S.design?.view;
  const selection = () => h.scope === 'merge' ? S.mergeSel : S.design?.sel;
  const records = () => (view()?.inspection?.fittings || []).filter(f=>!['tee-run','cross-run'].includes(f.kind));
  // Valves/equipment keep their own symbols and losses; fitting overlay is
  // exclusively branch tees and elbows, not guessed schematic junctions.
  const glyphs = () => records().filter(f=>['tee','elbow'].includes(f.shape));
  const number = value => value == null ? "미확정" : Number(value).toLocaleString("ko-KR",{maximumFractionDigits:4});
  const physical = f => ({tee:"티",cross:"크로스",elbow:"엘보",valve:"밸브",other:"기타 부속·기기"})[f.shape] || "부속";
  const selected = f => selection()?.kind === "node" ? String(f.node)===String(selection().label)
    : selection()?.kind === "pipe" && String(f.pipe)===String(selection().label);
  function segments(f) {
    // Geometry lives in the same coordinate system as the network. Zoom is
    // applied exactly once by sx/sy; an incomplete tee is not a stock T icon.
    const r=Number(f.glyph_radius) || 200*Math.abs(view()?.underlay?.k || 1);
    const line=(a,b,c,d)=>[sx(f.x+a),sy(f.y+b),sx(f.x+c),sy(f.y+d)];
    if(f.arms?.length) return f.arms.map(([dx,dy])=>line(0,0,dx*r,dy*r));
    // A small location diamond means geometry is unknown, not an invented elbow.
    const q=r*.3;
    return [line(-q,0,0,q),line(0,q,q,0),line(q,0,0,-q),line(0,-q,-q,0)];
  }
  function draw() {
    ctx.save();ctx.lineCap="round";ctx.lineJoin="round";ctx.globalAlpha=1;
    const painted=new Set();
    for(const f of glyphs()) {
      if(!Number.isFinite(f.x)||!Number.isFinite(f.y))continue;
      const key=[f.x,f.y,f.shape].join('|');if(painted.has(key))continue;painted.add(key);
      const lines=segments(f);
      // Local dark casing separates a fitting from the red worst-path overlay.
      ctx.strokeStyle="#000000";ctx.lineWidth=5.6;ctx.setLineDash([]);ctx.beginPath();
      for(const [x,y,a,b] of lines){ctx.moveTo(x,y);ctx.lineTo(a,b);}ctx.stroke();
      ctx.strokeStyle="#ff616c";ctx.lineWidth=selected(f)?3.3:2.3;ctx.setLineDash([4,3]);
      ctx.beginPath();for(const [x,y,a,b] of lines){ctx.moveTo(x,y);ctx.lineTo(a,b);}ctx.stroke();
    }
    ctx.restore();
  }
  function distance(px,py,line) {
    const [ax,ay,bx,by]=line,dx=bx-ax,dy=by-ay;
    const t=Math.max(0,Math.min(1,((px-ax)*dx+(py-ay)*dy)/(dx*dx+dy*dy||1)));
    return Math.hypot(px-ax-t*dx,py-ay-t*dy);
  }
  function hit(x,y) {
    let picked=null,best=8;
    for(const f of glyphs()) {
      if(!Number.isFinite(f.x)||!Number.isFinite(f.y))continue;
      const d=Math.min(...segments(f).map(s=>distance(sx(x),sy(y),s)));
      if(d<best){best=d;picked=f;}
    }
    return picked ? {kind:picked.node!=null?'node':'pipe',label:String(picked.node ?? picked.pipe)} : null;
  }
  function card(kind,label) {
    const rows=records().filter(f=>kind==='node' ? f.node!=null && String(f.node)===String(label) : String(f.pipe)===String(label));
    if(!view()?.inspection) return grp("부속 속성")+'<div class="ins-none">부속 상세 정보를 아직 받지 못했습니다. 서버 업데이트 후 미리보기를 다시 불러오세요.</div>';
    if(!rows.length)return grp("부속 속성")+'<div class="ins-none">이 위치에 등록된 부속이 없습니다.</div>';
    let html=grp(`부속 속성 · ${rows.length}개 항목`);
    for(const f of rows) {
      html+=`<section class="fit-property" data-fitting-id="${esc(f.id)}"><div class="fit-heading">${esc(f.name)} <span>× ${esc(f.count)}</span></div>`;
      html+=kv("물리 부속",esc(physical(f)))+kv("계산 통과 / 종류",esc(f.name));
      if(f.flow_path)html+=kv(f.flow_label||"계산 경로",esc(f.flow_path.join(f.loss_status?' — ':' → ')));
      html+=kv("재질 / 호칭경",`${esc(f.material || "미지정")} / ${esc(f.dia ?? "미지정")}A`);
      if(f.original_degree!=null)html+=kv("원본 → 계산 연결",`${esc(f.original_degree)} → ${esc(f.current_degree)} 포트`);
      html+=kv("소속 배관",insLink("pipe",f.pipe));
      html+=kv("등가길이 / 개",`${number(f.eq_m)}${f.eq_m==null?'':' m'}`);
      html+=kv("이 항목 합계",`${number(f.eq_total_m)}${f.eq_total_m==null?'':' m'}`);
      html+=`<div class="fit-source">${esc(f.eq_source)}</div>`;
      html+=`<details class="fit-evidence"><summary>판정·위치 근거</summary><div>${esc(f.origin)} · ${esc(f.note)}</div><div>${esc(f.geometry_source)}</div>`;
      if(f.source_node!=null)html+=`<div>원본 절점 ${esc(f.source_node)}</div>`;
      if(f.symbolic)html+='<div class="fit-warning">확인된 연결 방향만 표시합니다. 생략된 방향은 추측하지 않으며, 형상 미확인은 작은 마름모로 표시합니다.</div>';
      if(f.x==null || f.y==null)html+='<div class="fit-warning">설치 위치 미확인: 배관 위 임의의 위치에 표시하지 않았습니다.</div>';
      html+='</details></section>';
    }
    if(kind==='pipe') {
      const p=(h.scope==='merge' ? view()?.pipes || [] : S.design?.tables?.pipes || []).find(p=>String(p.label)===String(label));
      if(p)html+=kv("현재 배관 부속 합계",`${number(p.eq_len)}${p.eq_len==null?'':' m'} (입력표 저장값)`);
      const native=rows.filter(f=>f.origin==='부속 입력표');
      if(native.some(f=>f.eq_total_m==null))html+='<div class="fit-warning">미확정 등가길이가 있습니다. 저장 합계가 0이어도 손실 없음이 확인된 것은 아닙니다.</div>';
      else if(p?.eq_len!=null && Math.abs(native.reduce((sum,f)=>sum+Number(f.eq_total_m),0)-Number(p.eq_len))>.001)
        html+='<div class="fit-warning">종류별 근거 합계와 저장된 배관 합계가 다릅니다. 직접 수정·관경 변경 이력을 확인하세요. 이 카드는 계산값을 자동 변경하지 않습니다.</div>';
      html+='<div class="ins-none">실제 배관 길이와 등가길이는 별도입니다. 위 부속값은 종류·관경별 근거 값이며, 기기 손실은 별도 집계됩니다.</div>';
    }
    return html;
  }
  function role(label) {
    return [...new Set(records().filter(f=>f.node!=null && String(f.node)===String(label)).map(physical))].join(' · ');
  }
  return {draw,hit,card,role};
};
