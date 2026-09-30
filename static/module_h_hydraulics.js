/* H-only hydraulic diagnostics: presentation never mutates original geometry. */
(() => {
  'use strict';
  const $=id=>document.getElementById(id),state=()=>window.__mf;
  const fmt=(v,n=2)=>typeof v==='number'&&Number.isFinite(v)?v.toLocaleString('ko-KR',{maximumFractionDigits:n}):'—';
  let report=null,sid=null,stale=false,dirty=false,defaults=null,ready=false,stamp='',lastView=null;
  const valid=()=>!!report&&!stale&&!dirty&&sid===state()?.sid&&
    (!state()?.mergeFingerprint||state().mergeFingerprint===report.fingerprint);
  const make=(tag,text='',cls='')=>{const n=document.createElement(tag);n.textContent=text;n.className=cls;return n;};
  function issues(){
    const heads=new Map((report?.nozzles||[]).map(n=>[String(n.label),n]));
    return (report?.violations||[]).map(v=>{
      const h=heads.get(String(v.label));let kind='node',label=String(v.node??h?.node??v.label??''),text=v.message||v.kind;
      if(v.kind==='velocity'){kind='pipe';label=String(v.label);text=`배관 ${label} · 유속 ${fmt(v.velocity_mps)} / ${fmt(v.limit)} m/s 이하`;}
      else if(v.kind==='head_min')text=`헤드 ${v.label} · ${fmt(v.flow_lpm)} / ${fmt(v.min_flow_lpm??h?.min_flow_lpm)} L/min 이상 · ${fmt(v.pressure_bar,3)} / ${fmt(v.min_pressure_bar??h?.min_pressure_bar,3)} bar 이상`;
      else if(v.kind==='head_max')text=`헤드 ${v.label} · 압력 ${fmt(v.pressure_bar,3)} / ${fmt(v.limit,3)} bar 이하`;
      else if(v.kind==='negative_pressure')text=`노드 ${label} · 음압 ${fmt(v.pressure_bar,3)} bar`;
      else kind=null;
      return {kind,label,text};
    });
  }
  function focus(kind,label){
    if(!valid()||state()?.stage!=='merge')return;
    const s=state(),v=s.mergeView,cv=$('cv'),nodes=new Map((v?.nodes||[]).map(n=>[String(n.label),n]));
    const p=kind==='pipe'&&v.pipes.find(p=>String(p.label)===label);
    const points=(p?[nodes.get(String(p.a)),nodes.get(String(p.b))]:[nodes.get(label)]).filter(Boolean);if(!points.length)return;
    const xs=points.map(p=>p.x),ys=points.map(p=>p.y),all=v.nodes;
    const extent=Math.max(...all.map(n=>n.x))-Math.min(...all.map(n=>n.x));
    const pad=Math.max(100,extent/80),dx=Math.max(...xs)-Math.min(...xs)+2*pad,dy=Math.max(...ys)-Math.min(...ys)+2*pad;
    s.view.scale=Math.min(cv.clientWidth*.48/dx,cv.clientHeight*.5/dy);
    s.view.ox=(Math.min(...xs)+Math.max(...xs))/2-cv.clientWidth*.42/s.view.scale;
    s.view.oy=(Math.min(...ys)+Math.max(...ys))/2-cv.clientHeight*.55/s.view.scale;
    window.ModuleHProperties.select(kind,label);stamp='';
  }
  function render(){
    if(!ready)return;
    const box=$('hh-result');box.replaceChildren();
    if(!report){box.append(make('p','아직 계산하지 않았습니다.'));return;}
    if(!valid()){box.append(make('p','이전 결과 · 입력 또는 배관망이 변경되어 재계산이 필요합니다.','hh-warning'));return;}
    box.append(make('h3',report.feasible?'조건 충족 · 역산 검토안':`조건 미달 ${report.violations.length}건`));
    if(report.feasible){
      const metrics=make('dl','','hh-specs'),water=report.water_supply;
      for(const [name,value] of [['펌프 필요 유량',`${fmt(report.pump.flow_lpm)} L/min`],['차압 양정',`${fmt(report.pump.differential_head_m)} m`],
        ['토출 압력',`${fmt(report.pump.discharge_pressure_bar,3)} bar`],['수원 검토량',water?.usable?`${fmt(water.required_volume_m3)} m³ · ${fmt(water.duration_minutes)}분`:'시간 반영 후 재계산']]){
        const cell=make('div');cell.append(make('dt',name),make('dd',value));metrics.append(cell);
      }
      box.append(metrics,make('p','선택한 작동망의 필요 성능입니다. 실제 펌프 선정·전체 수조 용량 확정은 별도 검토가 필요합니다.','hm-review-note'));
    }else{
      const locked=report.pipes.filter(p=>p.locked&&(report.violations||[]).some(v=>v.kind==='velocity'&&String(v.label)===String(p.label)));
      box.append(make('p',locked.length?`잠긴 배관 ${locked.map(p=>p.label).join(', ')}에서 유속 초과. 펌프 압력만 올려서는 해결되지 않습니다.`:'현재 관경·압력 상한·배관 잠금 및 아래 미달 위치를 확인하세요.','hh-warning'));
      const trial=make('details');trial.append(make('summary','마지막 시험값 · 설계값 아님'),make('p',`${fmt(report.pump.flow_lpm)} L/min · ${fmt(report.pump.differential_head_m)} m · ${fmt(report.pump.discharge_pressure_bar,3)} bar`));box.append(trial);
    }
    const table=make('button','역산 내경·유량 표 보기');table.type='button';table.onclick=()=>{
      $('hm-table').click();$('hp-basis').value='proposal';$('hp-kind').value='pipe';$('hp-basis').dispatchEvent(new Event('change'));
    };box.append(table);
    if(report.violations.length){
      const list=make('ul','','hh-issues');
      for(const issue of issues()){
        const item=make('li'),button=make('button',issue.text);button.type='button';button.dataset.kind=issue.kind||'';button.dataset.label=issue.label;
        button.disabled=!issue.kind;button.onclick=()=>focus(issue.kind,issue.label);item.append(button);list.append(item);
      }box.append(list);
    }
    const evidence=make('details');evidence.append(make('summary',`산정 근거 · 확인 사항 ${report.warnings.length}건`));
    evidence.append(make('p',`질량 잔차 ${fmt(report.solution.mass_error_lps,9)} L/s · 에너지 잔차 ${fmt(report.solution.energy_error_m,9)} m · 수동력 ${fmt(report.pump.hydraulic_power_kw)} kW (모터 정격 아님)`));
    const list=make('ul');report.warnings.forEach(w=>list.append(make('li',w)));evidence.append(list);box.append(evidence);
  }
  window.addEventListener('module-h-sizing-basis',e=>{
    defaults=e.detail.review_defaults;if(ready&&defaults&&!$('hh-duration').value)$('hh-duration').value=defaults.water_duration_minutes;
  });
  window.addEventListener('module-h-sizing-result',e=>{
    report=e.detail.report;sid=state()?.sid;stale=!!e.detail.stale;dirty=false;stamp='';render();
  });
  window.addEventListener('module-h-sizing-dirty',()=>{dirty=true;stamp='';render();});
  window.addEventListener('module-h-sizing-error',e=>{
    dirty=true;stamp='';if(ready)$('hh-result').replaceChildren(make('p',`역산을 시작하지 못했습니다. ${e.detail.message}`,'hh-warning'));
  });
  function init(){
    if(!window.ModuleHIntegration)return false;
    const duration=make('label','','f');duration.innerHTML='<span>방수 지속시간 (분)</span><input id="hh-duration" type="number" min="1" step="1" placeholder="기준망 불러오기 후 기본값">';
    $('sz-flow').closest('.hm-field-grid').append(duration);
    const note=make('details');note.id='hh-water-basis';note.innerHTML='<summary>수원 기준 · 적용 범위</summary><p>일반 검토 기본값 20분. NFTC 103 2.1.1의 수리계산 수원 기준을 사용합니다. 고층·특수용도·기존 건축물은 적용 기준을 확인하세요.</p><div class="hh-duration-presets"><button type="button" data-duration="20">일반 20분</button><button type="button" data-duration="40">40분 검토</button><button type="button" data-duration="60">60분 검토</button></div><p>고층건축물은 NFPC 604 제6조의 별도 수원 기준을 확인해야 합니다. 50층 이상은 더 강화된 기준이 적용됩니다. 현재 선택망 × 지속시간 결과는 옥상 추가수원·겸용설비·최불리 작동구역·유효/사수량 등을 포함한 최종 수조 용량이 아닙니다.</p><a href="https://www.law.go.kr/LSW/admRulInfoP.do?admRulSeq=2100000281674&chrClsCd=010201" target="_blank" rel="noopener">NFTC 103 2.1.1</a> · <a href="https://law.go.kr/LSW/admRulLsInfoP.do?admRulId=42497&efYd=0" target="_blank" rel="noopener">NFPC 604 제6조</a>';
    duration.closest('section').after(note);
    $('hh-duration').oninput=()=>{$('sz-mode').dispatchEvent(new Event('input'));};
    note.querySelectorAll('[data-duration]').forEach(b=>b.onclick=()=>{$('hh-duration').value=b.dataset.duration;$('hh-duration').dispatchEvent(new Event('input'));});
    if(defaults)$('hh-duration').value=defaults.water_duration_minutes;
    const joint=make('button','관경·펌프 함께 재검토');joint.id='hh-joint';joint.type='button';
    joint.title='기존 관경 이상으로 확대 검토합니다. 명시한 배관 잠금·기기 손실 잠금은 유지하고 원본은 덮어쓰지 않습니다.';
    $('sz-run').before(joint);
    joint.onclick=()=>{
      if(!$('busy').classList.contains('hidden'))return;
      $('sz-mode').value='joint';$('ha-sizing-scope').value='enlarge';$('sz-mode').dispatchEvent(new Event('input'));window.ModuleHUI.queue();
      $('sz-run').click();
    };
    const result=make('div');result.id='hh-result';result.setAttribute('aria-live','polite');$('sz-result').after(result);
    const label=make('label');label.id='hh-fail-label';label.innerHTML='<input type="checkbox" id="hh-show-fail" checked><span id="hh-fail-count">미달 위치</span>';$('bore-overlay').prepend(label);
    $('hh-show-fail').onchange=()=>{stamp='';};
    const overlay=make('canvas');overlay.id='hh-fail-canvas';overlay.setAttribute('aria-hidden','true');$('stage').append(overlay);
    ready=true;render();return true;
  }
  const boot=setInterval(()=>{if(init())clearInterval(boot);},100);
  setInterval(()=>{
    if(!ready)return;
    const s=state(),active=s?.stage==='merge',ok=valid(),bad=issues(),overlay=$('hh-fail-canvas');
    $('hh-fail-label').hidden=!active;$('hh-show-fail').disabled=!ok||!bad.length;
    const caption=ok?`미달 위치 ${bad.length}건`:report?'미달 위치 · 재계산 필요':'미달 위치';
    if($('hh-fail-count').textContent!==caption)$('hh-fail-count').textContent=caption;
    $('hh-joint').disabled=!$('busy').classList.contains('hidden');
    overlay.hidden=!active||!ok||!bad.length||!$('hh-show-fail').checked;
    if(overlay.hidden){stamp='';return;}
    const v=s.mergeView,b=$('cv').getBoundingClientRect(),dpr=devicePixelRatio||1;
    const key=JSON.stringify([report.proposal_id,s.view,b.width,b.height,dpr]);if(key===stamp&&lastView===v)return;stamp=key;lastView=v;
    overlay.width=Math.round(b.width*dpr);overlay.height=Math.round(b.height*dpr);const ctx=overlay.getContext('2d');ctx.setTransform(dpr,0,0,dpr,0,0);
    const nodes=new Map((v?.nodes||[]).map(n=>[String(n.label),n])),pipes=new Map((v?.pipes||[]).map(p=>[String(p.label),p]));
    ctx.strokeStyle='#ff3838';ctx.lineWidth=5;ctx.shadowColor='#ff3838';ctx.shadowBlur=7;
    const drawn=new Set();
    for(const issue of bad){
      const key=issue.kind+':'+issue.label;if(drawn.has(key))continue;drawn.add(key);
      if(issue.kind==='pipe'){
        const p=pipes.get(issue.label),a=p&&nodes.get(String(p.a)),b=p&&nodes.get(String(p.b));if(!a||!b)continue;
        ctx.beginPath();ctx.moveTo(s.toScreenX(a.x),s.toScreenY(a.y));ctx.lineTo(s.toScreenX(b.x),s.toScreenY(b.y));ctx.stroke();
      }else if(issue.kind==='node'){
        const n=nodes.get(issue.label);if(!n)continue;const x=s.toScreenX(n.x),y=s.toScreenY(n.y);
        ctx.beginPath();ctx.arc(x,y,11,0,Math.PI*2);ctx.moveTo(x-5,y-5);ctx.lineTo(x+5,y+5);ctx.moveTo(x+5,y-5);ctx.lineTo(x-5,y+5);ctx.stroke();
      }
    }
  },150);
  window.ModuleHHydraulics={getReport:()=>report,issues,focus};
})();
