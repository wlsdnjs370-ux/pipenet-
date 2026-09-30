/* Bore provenance presentation only. No sizing rules or second edit store. */
window.ModuleFBore = (() => {
  "use strict";
  const styles = {
    drawing:{color:'#42ffe0',dash:[],label:'도면 표기 참조'},
    rule:{color:'#ffb347',dash:[7,4],label:'규약 보완'},
    review:{color:'#e7a1ff',dash:[10,3,2,3],label:'보정 · 검토 필요'},
    unknown:{color:'#8795a5',dash:[2,5],label:'근거 미확인'}
  };
  const esc = s => String(s ?? '').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
  const num = n => n == null ? '미기록' : Number(n).toLocaleString('ko-KR',{maximumFractionDigits:2});
  function info(p) {
    if(p?.bore_info)return p.bore_info;
    const source=p?.src || p?.dia_src || p?.dia_source;
    return {source,category:({text:'drawing',nfpc_min:'review',nfpc_fallback:'rule',review_default:'review'})[source] || 'unknown',
      edited:['user','사람','직접 편집 · 라이브러리','user_fix'].includes(source),current_mm:p?.dia};
  }
  function style(p){return styles[info(p).category] || styles.unknown;}
  function paint(ctx,p,x1,y1,x2,y2,width=3,stroke=null) {
    const s=style(p),i=info(p);
    ctx.save();ctx.strokeStyle=s.color;ctx.lineWidth=Math.max(2.5,width);
    ctx.lineCap='round';ctx.setLineDash(s.dash);ctx.shadowColor=s.color;ctx.shadowBlur=5;
    if(stroke)stroke();else{ctx.beginPath();ctx.moveTo(x1,y1);ctx.lineTo(x2,y2);ctx.stroke();}
    ctx.shadowBlur=0;ctx.setLineDash([]);
    const x=(x1+x2)/2,y=(y1+y2)/2;
    if(i.edited){ctx.strokeStyle='#f0f6ff';ctx.fillStyle='#08111c';ctx.lineWidth=1.5;
      ctx.beginPath();ctx.moveTo(x,y-4);ctx.lineTo(x+4,y);ctx.lineTo(x,y+4);ctx.lineTo(x-4,y);ctx.closePath();ctx.fill();ctx.stroke();}
    if(i.review || i.category==='review'){ctx.font='bold 12px sans-serif';ctx.fillStyle=s.color;ctx.fillText('!',x+5,y-6);}
    ctx.restore();
  }
  function card(p) {
    const i=info(p),s=style(p),kv=(k,v)=>`<div class="bore-kv"><span>${esc(k)}</span><b>${esc(v)}</b></div>`;
    let h=`<section class="bore-card" data-bore-category="${esc(i.category)}"><div class="bore-card-head"><span style="color:${s.color}">━━ ${s.label}</span>${i.edited?'<span class="bore-user">◇ 사용자 수정</span>':''}</div>`;
    h+=kv('현재 호칭경',`${num(p.dia)} mm`);
    if(i.continuous_sources?.length){
      const oldLength=i.continuous_sources.reduce((sum,row)=>sum+Number(row.length || 0),0);
      const currentLength=Number(p.len_m ?? p.length);
      const changed=Number.isFinite(currentLength)&&Math.abs(oldLength-currentLength)>1e-6;
      h+=kv('연속관 통합',changed?`통합 당시 원본 ${i.continuous_sources.length}개 구간 · 이후 길이/분할 변경`:`원본 ${i.continuous_sources.length}개 구간 → 배관 1개`);
      h+='<div class="bore-note">클릭·관경 수정은 이 연속 구간 전체에 적용됩니다. 길이와 손실은 합산하며 원본 구간별 관경 근거는 아래에 보존합니다.</div>';
      h+='<details class="bore-segments"><summary>통합 전 구간·관경 근거</summary>';
      for(const row of i.continuous_sources){
        const e=row.bore_provenance || {},category=e.serial_category || ({text:'drawing',nfpc_fallback:'rule',review_default:'review',nfpc_min:'review'})[e.source] || 'unknown';
        h+=kv(`${row.label} · ${row.in} → ${row.out}`,`${num(row.length)} m · ${num(row.dia)}A`);
        h+=kv('원본 관경 근거',`${styles[category]?.label || '근거 미확인'}${e.manual?' · 사용자 수정':''}`);
        if(e.text_mm!=null)h+=kv('원본 문자값',`${num(e.text_mm)} mm`);
        if(e.text_xy_mm)h+=kv('원본 문자 위치 (CAD mm)',e.text_xy_mm.map(num).join(', '));
        if(e.manual?.note)h+=kv('수정 메모',e.manual.note);
        for(const reason of e.review_reasons||[])h+=`<div class="bore-warning">${esc(reason)}</div>`;
      }
      h+='</details>';
      if(i.review_only || i.source==='review_default')h+='<div class="bore-warning">검토용 미확정 상태 유지 · 연속관 통합은 관경 확정이나 수리계산 승인이 아닙니다.</div>';
    }else if(i.source==='text'||i.source==='nfpc_min'){
      h+=kv('도면 문자값',i.text_mm==null?'기존 표에 근거값 미기록':`${num(i.text_mm)} mm`);
      h+=kv('매칭 방식',i.method || '기존 문자 매칭');
      if(i.text_xy_mm)h+=kv('원본 문자 위치 (CAD mm)',i.text_xy_mm.map(num).join(', '));
      if(i.distance_mm!=null)h+=kv('문자–배관 거리',`${num(i.distance_mm)} mm`);
      h+='<div class="bore-note">문자 참조는 정답 보증이 아닙니다. 같은 관로의 전파값은 추론이며, 반복 배관 대표 표기·리더선·임의 범위 표기는 자동 확정하지 않습니다.</div>';
    }else if(i.source==='review_default'){
      h+=kv('관경 근거',`사용자 지정 검토용 기본값 ${num(i.review_default_mm)} mm`);
      h+='<div class="bore-warning">미확정 · 루프/그리드에는 규약 관경을 적용하지 않았습니다. PIPENET 계산·도면 대조가 필요합니다.</div>';
    }else if(i.source==='nfpc_fallback'){
      h+=kv('보완 사유',i.reason==='no_source_edge'?'원본 선분 대응 없음 (생성 배관)':i.reason==='no_coordinates'?'원본 선분 좌표 없음 · 문자 매칭 미실행':i.reason==='no_match'?'매칭 범위 안에서 관경 표기 미확인':'기존 표의 규약 보완 기록');
    }else h+='<div class="bore-note">배관별 자동 산정 근거가 기록되지 않았습니다. 이를 표기 없음이나 규약 적용으로 단정하지 않습니다.</div>';
    if(i.version>=2){
      if(i.source!=='text'&&i.source!=='nfpc_min'&&i.method)h+=kv('위상 판정',i.method);
      if(i.text_rotation_deg!=null)h+=kv('원문자 방향',`${num(i.text_rotation_deg)}° (CAD 기준)`);
      if(i.full_head_count!=null)h+=kv('원본 경로 전체 담당 헤드',`${num(i.full_head_count)}개`);
      if(i.excluded_nearest)h+=kv('제외한 가까운 문자',`${num(i.excluded_nearest.text_mm)} mm · ${i.excluded_nearest.reason}`);
      if(i.candidate_mm?.length)h+=kv('충돌하는 후보 호칭경',i.candidate_mm.map(num).join(' / ')+' mm');
      if(i.partial_text_mm?.length)h+=kv('일부 구간에서만 확인한 표기',i.partial_text_mm.map(num).join(' / ')+' mm');
      for(const reason of i.review_reasons||[])h+=`<div class="bore-warning">${esc(reason)}</div>`;
      if(i.block_export)h+='<div class="bore-warning">확인 전 파일 산출 보류 · 아래 값은 임시 보완값입니다.</div>';
    }
    if(i.head_count!=null)h+=kv('규약 산정 담당 헤드',`${num(i.head_count)}개`);
    if(i.rule_mm!=null)h+=kv('코드에 설정된 규약값',`${num(i.rule_mm)} mm`);
    if(i.source==='nfpc_min')h+='<div class="bore-warning">도면 표기는 찾았지만 코드의 규약값으로 상향되었습니다. 표기 없음과 다른 검토 대상입니다.</div>';
    if(i.adjustment)h+=kv(i.adjustment.reason,`${num(i.adjustment.previous_mm)} → ${num(i.adjustment.value_mm)} mm`);
    if(i.untracked_change)h+='<div class="bore-warning">기록된 산정값과 현재 관경이 다릅니다. 후속 변경 근거를 확인하세요.</div>';
    if(i.manual)h+=kv('최초 자동값',`${num(i.auto_mm)} mm`)+kv('최근 사용자 변경',`${num(i.manual.previous_mm)} → ${num(i.manual.value_mm)} mm`)+kv('수정 메모',i.manual.note || '미기록');
    h+='<button type="button" class="bore-edit">관경·재질 바로 수정</button><div class="bore-note">확정 시 망·표·산출 데이터 갱신 · 기존 편집 기록으로 되돌리기 가능</div></section>';
    return h;
  }
  function legend(pipes) {
    const counts={drawing:0,rule:0,review:0,unknown:0};let edited=0;
    const reviewOnly=(pipes||[]).some(p=>p.review_only || p.bore_provenance?.review_only || p.bore_info?.review_only || info(p).source==='review_default');
    for(const p of pipes || []){const i=info(p);counts[i.category in counts?i.category:'unknown']++;if(i.edited)edited++;}
    return Object.entries(styles).filter(([k])=>counts[k] || k==='drawing' || (k==='rule'&&!reviewOnly)).map(([k,s])=>
      `<span><i style="color:${s.color}">${s.dash.length?'┈┈':'━━'}</i> ${s.label} ${counts[k]}</span>`).join('')+(reviewOnly?'<span>규약 미적용 · 검토용 미확정</span>':'')+(edited?`<span>◇ 사용자 수정 ${edited}</span>`:'');
  }
  return {info,style,paint,card,legend,styles};
})();
