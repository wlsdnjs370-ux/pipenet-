/* Shared graph inspector. All hydraulic rules and library values live on the server. */
window.createModuleFEditor = function (h) {
  "use strict";
  const $ = id => document.getElementById(id), S = h.state;
  let data = null, selection = null, stamp = "", sequence = 0, preview = null, pending = false;
  let pick = null;
  let inflight = null, clipboard = null, diagonal = false, popupAt = [20,20], menuSequence = 0;
  const panel=$("ne-panel"), stage=$("stage"), context=$("ne-context");
  stage.append(panel); panel.classList.add("ne-popover");
  const view = () => scope()==="merge" ? S.mergeView : S.design?.view;
  const selectedNode = () => data?.nodes.find(n=>String(n.label)===String(selection?.label));
  const attached = () => data?.pipes.filter(p=>[String(p.in),String(p.out)].includes(String(selection?.label))) || [];
  const selectedPipe = () => data?.pipes.find(p=>String(p.label)===String(selection?.kind==='node' ? $("ne-attached").value : selection?.label));
  const scope = () => S.stage === "merge" ? "merge" : "design";
  const current = () => S.stage === "merge" ? S.mergeSel : S.design?.sel;
  const active = () => ["design", "merge"].includes(S.stage);
  const show = (id, on) => $(id).classList.toggle("hidden", !on);
  function options(id, rows, value) {
    const el = $(id); el.replaceChildren();
    for (const [v, label, disabled] of rows) {
      const op = new Option(label, v); op.disabled = !!disabled; el.add(op);
    }
    if ([...el.options].some(o => o.value === String(value) && !o.disabled)) el.value = value;
    else el.selectedIndex = [...el.options].findIndex(o => !o.disabled);
  }
  function clearPreview() {
    preview = null; $("ne-apply").disabled = false;
    show("ne-preview-canvas", false); $("ne-message").textContent = "";
  }
  function message(text, error = false) {
    $("ne-message").textContent = text; $("ne-message").className = error ? "ne-error" : "";
  }
  function load(force = false, history = false) {
    if (!active() || !S.sid) return;
    const sel = current(), key = `${S.sid}|${S.stage}|${sel?.kind}|${sel?.label}`;
    if (!force && key === stamp) return inflight;
    stamp = key; const seq = ++sequence;
    selection = sel ? {...sel} : null;
    clearPreview();
    inflight=(async()=>{
    try {
      const d = await h.api(`/api/module-f/network-editor?sid=${encodeURIComponent(S.sid)}&scope=${scope()}${history?'&history=1':''}`);
      if (seq !== sequence || !active()) return;
      data = d; selection = sel ? {...sel} : null;
      if (d.nodes && d.pipes) $("ne-count-note").textContent = `노드 ${d.nodes.length} · 배관 ${d.pipes.length} · 헤드 ${d.nodes.filter(n=>n.head).length}`;
      $("ne-undo").disabled = !d.undo; $("ne-redo").disabled = !d.redo;
      $("ne-source").textContent = (d.catalog?.sources || []).join(" · ");
      if (selection) form();
      if (d.conflict) { message(d.conflict,true); h.say(d.conflict,"warn"); }
    } catch (err) { if (seq === sequence && selection) message(err.message,true); }
    finally { overlay(); }
    })();
    return inflight;
  }
  function render(force = false) {
    show("ne-toolbar", active());
    stage.classList.toggle('ne-direct',active());
    stage.classList.toggle('ne-merge',S.stage==='merge');
    const sel = current();
    if (!active() || String(sel?.label)!==String(selection?.label) || sel?.kind!==selection?.kind) close();
    if (active()) {
      // The immediate editor replaces the deferred topology buttons.
      show("dg-ins-ops",false); show("dg-ins-noops",false);
      load(force);
    } else { sequence++; stamp = ""; pick=null; clearPreview(); }
    overlay();
  }
  function materialSizes() {
    if (!data?.catalog) return;
    const rows = data.catalog.pipes.filter(p => p.type === $("ne-material").value);
    options("ne-dn",rows.map(p => [p.dia,`${p.dia}A`]),$("ne-dn").value);
    inner();
  }
  function inner() {
    const p = data?.catalog.pipes.find(p => p.type === $("ne-material").value && p.dia === +$("ne-dn").value);
    $("ne-inner").textContent = p ? `내경 ${p.inner_mm} mm · C=${p.c} · 조도 ${p.roughness_mm} mm` : "사용 가능한 관경이 없습니다.";
  }
  function form() {
    if (!selection || !data?.catalog) return;
    const pipe = data.pipes.find(p => String(p.label) === String(selection.label));
    const node = data.nodes.find(n => String(n.label) === String(selection.label));
    const isPipe = selection.kind === "pipe";
    $("ne-selected").textContent = isPipe ? `배관 ${selection.label} · ${pipe?.length ?? "?"} m`
      : `노드 ${selection.label} · XYZ ${(node?.xyz || []).map(v => Number(v).toFixed(4)).join(", ")} m`;
    const rows = isPipe ? [["resize","배관 늘리기 / 줄이기"],["split","배관 분할 → 분기 노드"],
      ["pipe","재질 / 관경 변경"],["fitting","부속 / 밸브 추가"],
      ["remove_fitting","추가한 부속 제거"],["delete","배관 삭제"]]
      : [["extend","배관 + 끝 노드 추가"],["connect","기존 노드 연결"],
         ["head","헤드 설치 / 방수량 변경"],["node","일반 노드로 변경"],
         ["fitting","노드에 부속 / 밸브 추가"],["remove_fitting","추가한 부속 제거"],
         ["move_node","노드 세부속성 변경"],["delete_node","노드 삭제"],
         ["merge_node","노드 삭제 (배관 합치기)"],["paste","복사한 노드 + 접속 배관 붙여넣기"]];
    options("ne-op",rows,isPipe ? "resize" : "extend");
    const incident=attached();
    options('ne-attached',incident.map(p=>[p.label,`${p.label} · ${p.dia}A · ${p.length}m`]),incident.find(p=>String(p.out)===String(selection.label))?.label);
    if(node) ['x','y','z'].forEach((key,i)=>$("ne-"+key).value=node.xyz[i]);
    const materials = [...new Set(data.catalog.pipes.map(p=>p.type))];
    const basePipe=isPipe?pipe:selectedPipe();
    options("ne-material",materials.map(m=>[m,m]),basePipe?.type || $("ne-material").value);
    materialSizes(); if (basePipe) { $("ne-dn").value = basePipe.dia; inner(); }
    options("ne-end",data.nodes.filter(n=>String(n.label)!==String(selection.label))
      .map(n=>[n.label,`${n.label} · Z=${n.xyz[2]}m${n.head ? " · 헤드" : ""}`]));
    options("ne-fitting",data.catalog.fittings.map(f=>[f.id,f.name]));
    removalOptions();
    options("ne-nozzle",data.catalog.nozzles.map(n=>[n.id,n.name]),node?.head_spec?.id);
    $("ne-flow").value = node?.head_spec?.flow_lmin ?? '';
    if (isPipe && pipe) $("ne-length").value = pipe.length;
    fields();
  }
  function fields() {
    clearPreview(); pick=null;
    const op = $("ne-op").value;
    show("ne-pipe-fields",["extend","connect","pipe","paste"].includes(op));
    show("ne-axis-wrap",["extend","paste"].includes(op));
    show("ne-length-wrap",["extend","resize","split","paste"].includes(op));
    show("ne-attached-wrap",selection?.kind==='node' && ['fitting','remove_fitting'].includes(op));
    show("ne-coordinates",op==='move_node');
    show("ne-end-wrap",op==="connect"); show("ne-fitting-fields",op==="fitting");
    show("ne-pick-end",op==="connect"); show("ne-pick-split",op==="split");
    show("ne-remove-wrap",op==="remove_fitting"); show("ne-head-fields",op==="head");
    $("ne-length-label").textContent = op==="split" ? "시작 노드에서 거리 (m)" : "길이 (m)";
    if (op==="split") {
      const p=data.pipes.find(p=>String(p.label)===String(selection.label));
      $("ne-length").value = +(p.length/2).toFixed(6);
    }
    $("ne-note-op").textContent = ({extend:"적용 후 새 끝 노드가 선택됩니다. 방향·길이를 바꾸며 이어 그릴 수 있습니다.",
      resize:"끝점 쪽에 연결된 배관망을 함께 이동합니다. 다른 배관의 길이는 유지됩니다.",
      split:"분할 후 새 노드에서 배관을 추가할 수 있습니다. 기존 부속은 시작 조각에 남습니다.",
      fitting:"라이브러리 등가길이를 기기로 추가해 SDF·KFP에 같은 손실을 반영합니다.",
      connect:"두 노드는 X·Y·Z 중 한 축과 평행해야 합니다. 중간 교차는 먼저 분할하세요.",
      head:"헤드 종류는 라이브러리, 상·하향은 접속관의 +Z/−Z 방향으로 정합니다.",
      delete:"관말 배관과 끝 노드를 삭제합니다.",delete_node:"관말 노드와 연결된 배관 1개를 삭제합니다.",
      merge_node:"중간 노드를 제거하고 두 배관의 길이·손실을 보존해 하나로 합칩니다.",
      move_node:"실제 좌표(m)를 바꾸고 접속 배관 길이도 갱신합니다. 축 정렬과 연결을 검사합니다.",
      paste:"복사한 노드와 접속 배관 1개를 새 가지로 추가합니다. 헤드 제원과 배관의 부속·밸브 손실도 복사합니다. 자르기는 확정할 때 원래 관말을 제거합니다."})[op] || "확정하면 배관망과 출력 데이터에 반영됩니다.";
    loss(); headSpec();
  }
  function loss() {
    const p = selectedPipe();
    const f = data?.catalog.fittings.find(f=>f.id===$("ne-fitting").value);
    const length=f?.lengths[String(p?.dia)];
    $("ne-loss").textContent = length == null ? "이 관경의 등가길이가 없어 적용할 수 없습니다."
      : `${p.dia}A · ${length}m × ${$("ne-count").value}개 = ${+(length*+$("ne-count").value).toFixed(6)}m`;
  }
  function headSpec() {
    const n=data?.catalog.nozzles.find(n=>n.id===$("ne-nozzle").value);
    $("ne-head-spec").textContent = n ? `K=${n.k_factor_si.toFixed(3)} L/min/√bar · 라이브러리 압력 ${n.min_bar.toFixed(3)}~${n.max_bar.toFixed(3)} bar` : "";
  }
  function command() {
    const op=$("ne-op").value, atNode=selection.kind==='node' && ['fitting','remove_fitting'].includes(op);
    return {op,target:atNode ? $("ne-attached").value : String(selection.label),axis:$("ne-axis").value,
      length:$("ne-length").value,distance:$("ne-length").value,end:$("ne-end").value,
      schedule:$("ne-material").value,dn:$("ne-dn").value,fitting:$("ne-fitting").value,
      count:$("ne-count").value,position:+$("ne-position").value/100,equipment:$("ne-remove").value,
      nozzle:$("ne-nozzle").value,flow:$("ne-flow").value,note:$("ne-note").value,
      node:atNode ? String(selection.label) : '',x:$("ne-x").value,y:$("ne-y").value,z:$("ne-z").value,
      ...(op==='paste' && clipboard ? {source:clipboard.source,source_pipe:clipboard.pipe,
        copied_revision:clipboard.revision,cut:clipboard.cut} : {})};
  }
  function drawPreview(view, result) {
    const canvas=$("ne-preview-canvas"), ctx=canvas.getContext("2d");
    const nodes=view.nodes, xs=nodes.map(n=>n.x),ys=nodes.map(n=>n.y);
    const minx=Math.min(...xs),maxx=Math.max(...xs),miny=Math.min(...ys),maxy=Math.max(...ys);
    const scale=Math.min((canvas.width-40)/Math.max(maxx-minx,1),(canvas.height-40)/Math.max(maxy-miny,1));
    const at=Object.fromEntries(nodes.map(n=>[n.label,[20+(n.x-minx)*scale,canvas.height-20-(n.y-miny)*scale]]));
    ctx.fillStyle="#0a111b";ctx.fillRect(0,0,canvas.width,canvas.height);
    for(const p of view.pipes){const a=at[p.a],b=at[p.b];if(!a||!b)continue;
      ctx.strokeStyle=(p.label===result.pipe || p.label===result.label)?"#ffd36e":"#72aaa5";
      ctx.lineWidth=2;ctx.beginPath();ctx.moveTo(...a);ctx.lineTo(...b);ctx.stroke();}
    if(result.kind==="node" && at[result.label]){ctx.fillStyle="#ffd36e";ctx.beginPath();ctx.arc(...at[result.label],6,0,Math.PI*2);ctx.fill();}
    show("ne-preview-canvas",true);
  }
  async function run(action, explicit=null) {
    if(pending || !active() || !$("busy").classList.contains("hidden")) return false;
    if(!data){await load(true);if(!data)return false;}
    if(action==="undo" && !data.undo || action==="redo" && !data.redo) return false;
    if(action==="reset" && !confirm("이 배관망의 직접 편집을 모두 되돌리고 추출 직후 상태로 복원할까요?"))return false;
    if(action==="apply" && !preview)return false;
    const cmd=action==="apply" ? preview.command : explicit || (selection ? command() : {});
    pending=true;h.busy(true,action==="preview" ? "배관 편집 미리보기" : "배관망 편집 반영");
    try {
      const result=await h.post('/api/module-f/network-editor',{sid:S.sid,scope:scope(),action,
        revision:data.revision,command:cmd,iso:!!$("mg-iso")?.checked,iso_z_scale:1});
      if(action==="preview"){
        preview={command:cmd};drawPreview(result.preview,result.selection);
        message(`적용 후 노드 ${result.counts.nodes} · 배관 ${result.counts.pipes} · 헤드 ${result.counts.heads}`);
        $("ne-apply").disabled=false;
      }else{
        const camera={...S.view},stage=S.stage;
        // [오너 2026-09-21] 세로관을 붙이면(예) 평면도 쪽 표고가 바뀐다. 급수원(기계실)이
        //   화면에서 제자리에 있게 맞춘다 — 다른 편집에서는 급수원이 안 움직여 그대로다.
        const pin=stage==="merge" ? S.mergeView?.nodes?.find(n=>n.input) : null;
        clearPreview();stamp="";
        if(stage==="merge")await h.reloadMerge();else await h.reloadDesign();
        Object.assign(S.view,camera);
        const moved=pin && S.mergeView?.nodes?.find(n=>String(n.label)===String(pin.label));
        if(moved){S.view.ox+=moved.x-pin.x;S.view.oy+=moved.y-pin.y;}
        if(result.selection)h.select(stage==="merge" ? "mg"+result.selection.kind : result.selection.kind,result.selection.label);
        if(action==='apply' && ['extend','paste','split'].includes(cmd.op))revealSelection();
        close();h.draw();await load(true);
        h.say(action==="undo"?"한 단계 되돌렸습니다.":action==="redo"?"다시 적용했습니다.":"배관망·표·출력 데이터에 반영했습니다.");
      }
      return true;
    }catch(err){clearPreview();$("ne-apply").disabled=true;message(err.message,true);h.say(err.message,"err");return ['undo','redo'].includes(action);}
    finally{pending=false;h.busy(false);}
  }
  $("ne-op").onchange=fields;
  $("ne-material").onchange=()=>{materialSizes();clearPreview();};
  $("ne-dn").onchange=()=>{inner();clearPreview();};
  for(const input of $("ne-panel").querySelectorAll("input,select"))input.addEventListener("input",()=>{clearPreview();loss();headSpec();});
  $("ne-preview").onclick=()=>run("preview");$("ne-apply").onclick=async()=>{
    if(!preview && !await run('preview'))return;
    await run('apply');
  };
  $("ne-apply").textContent='확정';
  $("ne-focus").onclick=()=>{h.editView();render(true);h.say("계산망의 노드나 배관을 클릭해 편집하세요.");};
  $("ne-cancel").onclick=close;$("ne-close").onclick=close;
  $("ne-join-yes").onclick=()=>riserDelete(true);$("ne-join-no").onclick=()=>riserDelete(false);
  $("ne-undo").onclick=()=>run("undo");$("ne-redo").onclick=()=>run("redo");
  $("ne-pick-end").onclick=()=>beginPick('end');
  $("ne-pick-split").onclick=()=>beginPick('split');
  $("ne-attached").onchange=()=>{removalOptions();loss();clearPreview();};
  $("ne-reset").onclick=()=>run("reset");
  $("ne-history").onclick=async()=>{await load(true,true);$("ne-history-text").textContent=data?.conflict || data?.notice || "";
    $("ne-history-text").textContent+="\n"+(data?.history || []).map((c,i)=>`${i+1}. ${c.op} · ${c.target} ${c.note || ""}`).join("\n") || "편집 기록 없음";
    for (const row of data?.archived || []) {
      $("ne-history-text").textContent+=`\n\n[이전 기준망 · ${row.cursor}건 보관]\n${row.file}\n`
        +(row.commands || []).map((c,i)=>`${i+1}. ${c.op} · ${c.target} ${c.note || ''}`).join('\n');
    }
    if (data?.archived?.length) $("ne-history-text").textContent+='\n\n같은 기준 배관망으로 표를 확정하면 해당 편집 기록을 자동 복원합니다.';
    $("ne-history-dialog").showModal();};
  window.addEventListener("keydown",e=>{
    if(!active())return;
    if(e.key==="Escape"){pick=null;clearPreview();return;}
    if((e.ctrlKey||e.metaKey) && e.shiftKey && e.key.toLowerCase()==="z"
      && !/INPUT|TEXTAREA|SELECT/.test(document.activeElement?.tagName || "") && !document.activeElement?.isContentEditable){e.preventDefault();run("redo");}
  });
  function canvasClick(x,y,maxD) {
    if(!pick || !active() || !selection)return false;
    const v=scope()==="merge" ? S.mergeView : S.design?.view;
    if(!v)return true;
    if(pick==="end"){
      let best=maxD,hit=null;
      for(const n of v.nodes){const d=Math.hypot(n.x-x,n.y-y);if(d<=best && String(n.label)!==selection.label){best=d;hit=n;}}
      if(hit){$("ne-end").value=hit.label;pick=null;show('ne-pick-note',false);show('ne-panel',true);clearPreview();message(`연결 대상: 노드 ${hit.label}. 확정하면 연결됩니다.`);}
    }else{
      const p=v.pipes.find(p=>String(p.label)===selection.label),a=v.nodes.find(n=>String(n.label)===String(p?.a)),b=v.nodes.find(n=>String(n.label)===String(p?.b));
      if(a&&b){const dx=b.x-a.x,dy=b.y-a.y,dd=dx*dx+dy*dy,t=((x-a.x)*dx+(y-a.y)*dy)/dd;
        if(dd>0 && t>0 && t<1 && Math.hypot(x-a.x-t*dx,y-a.y-t*dy)<=maxD){
          $("ne-length").value=+(t*data.pipes.find(p=>String(p.label)===selection.label).length).toFixed(6);
          pick=null;show('ne-pick-note',false);show('ne-panel',true);clearPreview();message("분할 위치를 지정했습니다. 확정하면 분할됩니다.");}}
    }
    return true;
  }
  async function deletePreview() {
    if(!selection)return;
    await load();
    if(!selection)return;
    if(riserPipe()){openJoin();return;}
    openAction(selection.kind==='pipe' ? 'delete' : 'delete_node');
  }
  // [오너 2026-09-21] 통합망 계통도의 세로관(높이차가 있는 배관) — 지울 때 위·아래를 붙일지 묻는다.
  function riserPipe(){
    if(scope()!=='merge'||selection?.kind!=='pipe'||!data)return null;
    const label=String(selection.label);
    const shown=S.mergeView?.pipes?.find(p=>String(p.label)===label);
    const pipe=data.pipes.find(p=>String(p.label)===label);
    const at=l=>data.nodes.find(n=>String(n.label)===String(l));
    const a=at(pipe?.in),b=at(pipe?.out);
    return shown?.part==='system' && a && b && Math.abs(a.xyz[2]-b.xyz[2])>1e-6 ? {pipe,a,b} : null;
  }
  async function openJoin(){
    await load();if(!active()||!selection||!data)return;
    const r=riserPipe();
    if(!r){openAction('delete');return;}
    show('ne-context',false);show('ne-panel',true);clearPreview();
    panel.classList.add('ne-join-mode');
    $("ne-title").textContent='배관 삭제';
    const [hi,lo]=r.a.xyz[2]>=r.b.xyz[2]?[r.a,r.b]:[r.b,r.a];
    const m=z=>`${(+z).toFixed(2)} m`;
    $("ne-selected").textContent=`배관 ${selection.label} · ${r.pipe.length} m · 계통도 세로관 · 위 ${hi.label} (${m(hi.xyz[2])}) → 아래 ${lo.label} (${m(lo.xyz[2])})`;
    const v=view(),ends=[r.pipe.in,r.pipe.out].map(l=>v?.nodes?.find(n=>String(n.label)===String(l)));
    if(ends.every(Boolean))popupAt=h.screen((ends[0].x+ends[1].x)/2,(ends[0].y+ends[1].y)/2);
    const px=popupAt[0]+110+panel.offsetWidth < stage.clientWidth ? popupAt[0]+110 : popupAt[0]-panel.offsetWidth-110;
    place(panel,px,popupAt[1]+12);
    $("ne-join-yes").focus();
  }
  // 예 → 지우고 위·아래를 붙인다(기계실 표고 고정, 평면도 쪽이 옮겨진다).
  // 아니오 → 배관만 지운다. 통합망이 끊겨 산출할 수 없다는 문구가 하단에 남는다.
  async function riserDelete(join){
    if(pending||!riserPipe())return;
    const cmd={op:join?'delete_join':'delete_cut',target:String(selection.label),note:$("ne-note").value};
    if(!await run('preview',cmd))return;
    if(!await run('apply'))return;
    if(!join)h.say(S.mergeView?.split || '통합 배관망이 완전히 연결되지 않아 아직 산출할 수 없습니다.','warn');
  }
  function removalOptions(){
    options('ne-remove',(data?.equipment || []).filter(e=>e.editor_library && String(e.pipe)===String(selectedPipe()?.label))
      .map(e=>[e.label,`${e.desc} ×${e.count} · ${e.eq_len}m`]));
  }
  function close(){
    menuSequence++;pick=null;show('ne-context',false);show('ne-panel',false);show('ne-pick-note',false);clearPreview();
    panel.classList.remove('ne-join-mode');
  }
  function place(el,x,y){
    el.style.left=Math.max(8,Math.min(x,stage.clientWidth-el.offsetWidth-8))+'px';
    el.style.top=Math.max(8,Math.min(y,stage.clientHeight-el.offsetHeight-72))+'px';
  }
  function revealSelection(){
    const n=view()?.nodes?.find(n=>String(n.label)===String(current()?.label));
    if(!n)return;
    const [x,y]=h.screen(n.x,n.y),margin=110;
    const nx=Math.max(margin,Math.min(x,stage.clientWidth-margin));
    const ny=Math.max(margin,Math.min(y,stage.clientHeight-margin-45));
    S.view.ox+=(x-nx)/S.view.scale;S.view.oy-=(y-ny)/S.view.scale;
  }
  async function openAction(op,axisValue){
    await load();if(!active()||!selection||!data)return;
    show('ne-context',false);show('ne-panel',true);panel.classList.remove('ne-join-mode');
    $("ne-op").value=op;fields();
    if(axisValue)$("ne-axis").value=axisValue;
    $("ne-title").textContent=$('ne-op').selectedOptions[0]?.textContent || '속성 변경';
    if(op==='paste'){
      if(!clipboard || clipboard.sid!==S.sid || clipboard.scope!==scope() || clipboard.revision!==data.revision){
        close();h.say('배관망이 바뀌었습니다. 원본 노드를 다시 복사하세요.','warn');return;
      }
      $('ne-material').value=clipboard.schedule;materialSizes();$('ne-dn').value=clipboard.dn;inner();
      $('ne-length').value=clipboard.length;
    }
    const px=popupAt[0]+110+panel.offsetWidth < stage.clientWidth ? popupAt[0]+110 : popupAt[0]-panel.offsetWidth-110;
    place(panel,px,popupAt[1]+12);
    if(['extend','resize','split','paste'].includes(op)){$('ne-length').focus();$('ne-length').select();}
  }
  function beginPick(mode){
    pick=mode;show('ne-panel',false);show('ne-context',false);show('ne-pick-note',true);
    $('ne-pick-note').textContent=mode==='end'?'연결할 다른 노드를 클릭하세요 · Esc 취소':'선택 배관의 분할 위치를 클릭하세요 · Esc 취소';
  }
  function menuButton(text,fn,disabled=false,title=''){
    const b=document.createElement('button');b.type='button';b.textContent=text;b.disabled=disabled;b.title=title;
    b.setAttribute('role','menuitem');b.onclick=fn;context.append(b);return b;
  }
  function copyNode(cut){
    const ps=attached(),p=ps.find(p=>String(p.out)===String(selection.label))||ps[0];
    if(!p)return;
    clipboard={sid:S.sid,scope:scope(),revision:data.revision,source:String(selection.label),pipe:String(p.label),
      schedule:p.type,dn:p.dia,length:p.length,cut};
    show('ne-context',false);h.say(`노드 ${selection.label} + 접속 배관 ${p.label} ${cut?'자르기 대기':'복사'}. 다른 시작 노드에서 우클릭 → 붙여넣기를 선택하세요.`);
  }
  function buildMenu(submenu){
    context.replaceChildren();
    const heading=document.createElement('div');heading.className='ne-menu-heading';
    heading.textContent=`${selection.kind==='pipe'?'배관':'노드'} ${selection.label}`;context.append(heading);
    const line=()=>context.append(document.createElement('hr'));
    if(submenu){
      menuButton('‹ 돌아가기',()=>buildMenu());
      menuButton('일반 노드로 변경',()=>openAction('node'));
      menuButton('헤드 설치 / 방수량 변경',()=>openAction('head'));
      menuButton('피팅 / 밸브 추가',()=>openAction('fitting'),!attached().length);
      menuButton('추가한 부속 제거',()=>openAction('remove_fitting'),!attached().some(p=>data.equipment.some(e=>e.editor_library&&String(e.pipe)===String(p.label))));
    }else if(selection.kind==='pipe'){
      const p=selectedPipe(),end=data.nodes.find(n=>String(n.label)===String(p?.out));
      menuButton('배관속성 변경',()=>openAction('pipe'));
      menuButton('배관길이 변경',()=>openAction('resize'));
      menuButton('배관 분할',()=>openAction('split'));line();
      menuButton('부속 / 밸브 추가',()=>openAction('fitting'));
      menuButton('추가한 부속 제거',()=>openAction('remove_fitting'),!data.equipment.some(e=>e.editor_library&&String(e.pipe)===String(p?.label)));line();
      const riser=riserPipe();
      menuButton('배관 삭제',()=>riser?openJoin():openAction('delete'),!riser && String(end?.terminal_pipe)!==String(selection.label),
        riser?'지운 뒤 위·아래 배관을 붙일지 묻습니다.':'연결을 보존할 수 있는 관말 배관에서 사용합니다.');
    }else{
      const n=selectedNode();
      menuButton('노드속성 변경 ›',()=>buildMenu(true));
      menuButton('세부속성 변경',()=>openAction('move_node'),n?.protected);line();
      menuButton('노드 → 노드 연결',async()=>{await openAction('connect');beginPick('end');},n?.head);line();
      menuButton('복사',()=>copyNode(false),!attached().length);
      menuButton('자르기',()=>copyNode(true),!n?.terminal_pipe);
      menuButton('삭제',()=>openAction('delete_node'),!n?.terminal_pipe);
      menuButton('노드 삭제 (배관 합치기)',()=>openAction('merge_node'),!!n?.merge_reason,n?.merge_reason||'');line();
      menuButton('붙여넣기',()=>openAction('paste'),!clipboard || clipboard.sid!==S.sid || clipboard.scope!==scope() || clipboard.revision!==data.revision || clipboard.source===String(selection.label) || n?.head);line();
      const toggle=menuButton(`${diagonal?'✓ ':''}45° 대각 화살표`,()=>{diagonal=!diagonal;buildMenu();overlay();},false,'화살표 표시각만 변경합니다. 실제 배관은 X·Y·Z 축 방향으로 생성합니다.');
      toggle.setAttribute('role','menuitemcheckbox');toggle.setAttribute('aria-checked',String(diagonal));
    }
    show('ne-context',true);place(context,...popupAt);
  }
  async function contextMenu(x,y){
    if(!active()||pending||!$('busy').classList.contains('hidden')||!h.calculationVisible())return;
    close();h.inspect(...h.world(x,y));
    const token=++menuSequence;popupAt=[x,y];await load();
    if(token!==menuSequence||!current()||!data)return;
    buildMenu();
  }
  // Six native buttons stay attached to the selected projected node on every paint.
  const axes=['+X','-X','+Y','-Y','+Z','-Z'];
  for(const axis of axes){const b=document.createElement('button');b.type='button';b.className='ne-arrow';b.dataset.axis=axis;
    b.textContent=axis;b.setAttribute('aria-label',axis+' 방향 배관 추가');
    b.onclick=()=>openAction('extend',axis);$('ne-arrows').append(b);}
  function overlay(){
    const sel=current(),v=view(),n=v?.nodes?.find(n=>String(n.label)===String(sel?.label));
    // Keep the read-only property card beside, not over, the XYZ handles.
    const inspector=$('dg-ins');
    if(scope()==='design' && !inspector.classList.contains('hidden')){
      const pipe=v?.pipes?.find(p=>String(p.label)===String(sel?.label));
      const ends=pipe?[pipe.a,pipe.b].map(l=>v.nodes.find(n=>String(n.label)===String(l))):[];
      const point=sel?.kind==='node'?n:ends.every(Boolean)&&ends.length===2
        ? {x:(ends[0].x+ends[1].x)/2,y:(ends[0].y+ends[1].y)/2}:null;
      const w=inspector.offsetWidth,ht=inspector.offsetHeight;
      const points=(v?.nodes || []).map(n=>h.screen(n.x,n.y));
      const focus=point?h.screen(point.x,point.y):null;
      const corners=[[stage.clientWidth-w-12,12],[stage.clientWidth-w-12,stage.clientHeight-ht-35],
        [12,12],[12,stage.clientHeight-ht-35]];
      const score=([a,b])=>points.filter(([x,y])=>x>a-24&&x<a+w+24&&y>b-24&&y<b+ht+24).length
        +(focus&&focus[0]>a-110&&focus[0]<a+w+110&&focus[1]>b-110&&focus[1]<b+ht+110?100:0);
      corners.sort((a,b)=>score(a)-score(b));
      inspector.style.left=Math.max(12,corners[0][0])+'px';inspector.style.right='auto';
      inspector.style.top=Math.max(12,corners[0][1])+'px';inspector.style.bottom='auto';
    }else{inspector.style.left='';inspector.style.right='';inspector.style.top='';inspector.style.bottom='';}
    const on=active()&&sel?.kind==='node'&&!!n&&h.calculationVisible();
    show('ne-gizmo',on);if(!on)return;
    const [x,y]=h.screen(n.x,n.y);
    if(x<0||y<0||x>stage.clientWidth||y>stage.clientHeight-32){show('ne-gizmo',false);return;}
    Object.assign($('ne-gizmo').style,{left:x+'px',top:y+'px'});
    if(panel.classList.contains('hidden')&&context.classList.contains('hidden'))popupAt=[x,y];
    const iso=scope()==='merge'?$('mg-iso').checked:$('dg-iso').checked;
    const c=diagonal?Math.SQRT1_2:Math.sqrt(3)/2,s=diagonal?Math.SQRT1_2:.5;
    const basis=iso?[[c,-s],[-c,s],[-c,-s],[c,s],[0,-1],[0,1]]:
      [[1,0],[-1,0],[0,-1],[0,1],[-Math.SQRT1_2,-Math.SQRT1_2],[Math.SQRT1_2,Math.SQRT1_2]];
    const colors=['#ff9991','#ff9991','#7fd9b3','#7fd9b3','#91bbff','#91bbff'];
    const obstacles=v.nodes.filter(o=>String(o.label)!==String(sel.label)).map(o=>h.screen(o.x,o.y));
    const used=[];
    let lines='';
    [...$('ne-arrows').children].forEach((b,i)=>{
      const [dx,dy]=basis[i];
      const candidates=[76,108,140,52,172].map(r=>[
        Math.max(24-x,Math.min(dx*r,stage.clientWidth-x-24)),
        Math.max(20-y,Math.min(dy*r,stage.clientHeight-y-100))]);
      const [bx,by]=candidates.find(([bx,by])=>
        obstacles.every(([nx,ny])=>Math.abs(nx-x-bx)>32||Math.abs(ny-y-by)>28) &&
        used.every(([px,py])=>Math.abs(px-bx)>44||Math.abs(py-by)>36)) || candidates[0];
      used.push([bx,by]);
      Object.assign(b.style,{left:bx+'px',top:by+'px',color:colors[i]});
      b.disabled=pending || selectedNode()?.head || !data;
      b.title=selectedNode()?.head?'먼저 우클릭 → 노드속성 변경 → 일반 노드로 변경하세요.':`${axes[i]} 방향으로 배관과 노드를 추가`;
      const len=Math.hypot(bx,by)||1,ux=bx/len,uy=by/len,ex=bx-ux*23,ey=by-uy*23;
      lines+=`<path d="M ${110+ux*28} ${110+uy*28} L ${110+ex} ${110+ey} m ${-ux*7-uy*4} ${-uy*7+ux*4} l ${ux*7+uy*4} ${uy*7-ux*4} l ${-ux*7+uy*4} ${-uy*7-ux*4}" fill="none" stroke="${colors[i]}" stroke-width="2"/>`;
    });
    $('ne-gizmo-lines').innerHTML=lines;
  }
  document.addEventListener('pointerdown',e=>{
    if(!e.target.closest('#ne-context'))show('ne-context',false);
  });
  window.addEventListener('keydown',e=>{
    if(!active())return;
    if(e.key==='Escape' && (!panel.classList.contains('hidden')||!context.classList.contains('hidden')||pick)){
      e.preventDefault();e.stopImmediatePropagation();close();return;
    }
    if(e.key==='Enter' && !panel.classList.contains('hidden') && e.target.tagName==='INPUT'){
      e.preventDefault();$('ne-apply').click();
    }
    if((e.ctrlKey||e.metaKey)&&!e.altKey&&!/INPUT|TEXTAREA|SELECT/.test(document.activeElement?.tagName||'') && !document.activeElement?.isContentEditable){
      const key=e.key.toLowerCase();
      if(selection?.kind==='node' && ['c','x','v'].includes(key)){
        e.preventDefault();if(key==='v')openAction('paste');else if(key==='c'||selectedNode()?.terminal_pipe)copyNode(key==='x');
      }
    }
  },true);
  return {render,overlay,contextMenu,openAction,canvasClick,deletePreview,undo:()=>run("undo"),refresh:()=>load(true)};
};
