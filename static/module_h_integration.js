/* H integration workspace. Reuses native controls; never builds or edits a model. */
(() => {
  'use strict';
  function init() {
    const $=id=>document.getElementById(id), ui=window.ModuleHUI;
    if(!ui || !$('panel-merge') || window.ModuleHIntegration)return;
    const field=id=>$(id)?.closest('label') || $(id);
    const text=(node,value)=>{value=String(value);if(node.textContent!==value)node.textContent=value;};
    const move=(target,...nodes)=>nodes.filter(Boolean).forEach(n=>target.append(n));
    const make=(tag,cls,content)=>{const n=document.createElement(tag);n.className=cls;if(content)n.textContent=content;return n;};
    const section=(title,kicker)=>{
      const n=make('section','hm-section'), h=make('div','hm-section-heading');
      h.append(make('span','hm-kicker',kicker),make('h3','',title));n.append(h);return n;
    };
    const details=(title,...nodes)=>{
      const d=make('details','hm-details');d.append(make('summary','',title));move(d,...nodes);return d;
    };
    const bar=make('section','');bar.id='hm-workspace';bar.hidden=true;bar.setAttribute('aria-label','통합 배관망 작업 도구');
    bar.innerHTML=`<div class="hm-workspace-title"><span class="hm-kicker">INTEGRATION</span><strong>통합 배관망</strong><span id="hm-state" role="status"></span></div>
      <nav aria-label="통합 작업"><button id="hm-connect" type="button">연결 설정</button><button id="hm-table" type="button">입력표</button><button id="hm-sizing" type="button">관경·펌프 역산</button><button id="hm-output" type="button">SDF 출력 <span aria-hidden="true">↗</span></button><button id="hm-view" type="button" aria-label="통합 보기 설정">보기</button></nav>`;
    document.body.append(bar);
    const panel=$('panel-merge');
    // Remove only now-empty captions. Every native control and handler survives.
    [...panel.children].filter(n=>n.tagName==='H2').forEach(n=>n.hidden=true);
    const sources=section('연결할 도면','01');
    const cards=make('ol','hm-sources');cards.id='hm-sources';sources.append(cards);
    const sourceAction=make('button','hm-text-button','도면 관리 ↗');sourceAction.type='button';sourceAction.onclick=()=>$('h-open').click();sources.firstElementChild.append(sourceAction);
    const supply=section('급수방식','02');move(supply,$('mg-modes'),$('mg-drop-row'));
    const build=make('div','hm-build');$('mg-build').textContent='배관망 통합';move(build,$('mg-build'),$('mg-why'));
    const result=section('연결 결과','03');
    const metrics=make('dl','hm-metrics');
    for(const [key,label] of [['nodes','노드'],['pipes','배관'],['nozzles','노즐'],['valves','밸브']]){
      const cell=make('div','');cell.append(make('dt','',label));const value=make('dd','','—');value.id='hm-'+key;cell.append(value);metrics.append(cell);
    }
    const note=make('p','hm-review-note','결합 완료는 수리계산 검증 완료를 뜻하지 않습니다.');
    const attachment=make('p','hm-attachment');attachment.id='hm-attachment';
    result.append(metrics,attachment,note);
    $('mg-summary').classList.remove('h-removed-copy');
    const evidence=details('연결 검사 · 좌표 근거',$('mg-ready'),$('mg-summary'),$('mg-legend'));
    evidence.id='hm-evidence';result.append(evidence);
    const stale=$('mg-stale');stale.classList.add('hm-validation');
    panel.append(stale,sources,supply,build,result);
    [...panel.children].filter(n=>n.classList.contains('row')&&!n.children.length&&!n.textContent.trim()).forEach(n=>n.hidden=true);

    // Group existing sizing fields, retaining exact IDs, units, values and events.
    const sizing=$('mg-sizing');sizing.querySelector(':scope > summary').hidden=true;
    const prepare=section('계산 대상','01');move(prepare,$('sz-load'),$('sz-basis'));sizing.prepend(prepare);
    const scope=section('역산 방식','02');move(scope,field('sz-mode'),field('ha-sizing-scope'));prepare.after(scope);
    const criteria=section('목표 조건','03'),grid=make('div','hm-field-grid');
    move(grid,...['sz-flow','sz-min','sz-max','sz-branch','sz-other','sz-pressure','sz-ceiling'].map(field));criteria.append(grid);scope.after(criteria);
    const boundary=details('공급 절점 · 작동 노즐',field('sz-source'),field('sz-heads'));criteria.after(boundary);
    const confirm=field('sz-confirm');confirm.classList.add('hm-confirm');
    const reviewFooter=make('p','hm-review-note','실제 펌프 곡선·흡입조건 및 PIPENET 대조 전에는 설계 확정값이 아닙니다.');sizing.append(reviewFooter);
    // Validation output remains visible; it is not part of collapsed help text.
    $('sz-result').classList.add('hm-result');

    const output=$('h-pane-export');
    const outputIntro=make('div','hm-intro');outputIntro.append(make('span','hm-kicker','PIPENET / SDF + SLF'),make('h2','','수리계산 프로그램으로 내보내기'));
    const outputNote=make('p','hm-review-note');outputNote.id='hm-output-note';outputNote.setAttribute('role','status');outputIntro.append(outputNote);output.prepend(outputIntro);
    const coordinate=field('mg-dl-coord');coordinate.className='f';
    [...coordinate.childNodes].filter(n=>n.nodeType===Node.TEXT_NODE).forEach(n=>n.remove());
    coordinate.prepend(make('span','','표시 좌표'));
    const outputSettings=section('출력 대상과 좌표','01');move(outputSettings,field('h-output-basis'),field('mg-dl-coord'));outputIntro.after(outputSettings);
    const outputActions=section('파일 생성 · 다운로드','02');move(outputActions,$('mg-emit'),$('mg-dl-sdf'));outputSettings.after(outputActions);
    $('mg-emit').textContent='SDF 파일 생성';$('mg-dl-sdf').textContent='SDF + SLF 내려받기';
    const files=details('생성 파일 목록',$('mg-files'));outputActions.append(files);
    const outputFoot=make('p','hm-review-note','SDF와 SLF는 같은 폴더에 보관하세요. 미확정 관경·부속은 출력 후에도 확인이 필요합니다.');outputActions.append(outputFoot);

    function click(id){const n=$(id);if(n && !n.disabled)n.click();}
    $('hm-connect').onclick=()=>ui.task($('hm-connect'));
    for(const [button,target] of [['hm-table','h-table'],['hm-sizing','h-sizing'],['hm-output','h-export'],['hm-view','h-view']])$(button).onclick=()=>click(target);
    const underControls=make('div','');underControls.id='hm-under-controls';
    const underLabel=make('label','');underLabel.id='hm-under-label';
    underLabel.innerHTML='<input type="checkbox" id="hm-under">원본 도면';
    const underKinds=make('fieldset','');underKinds.id='hm-under-kinds';underKinds.setAttribute('aria-label','원본 도면 종류');
    // Move the native controls, including dynamic system slots, without cloning
    // their state or handlers. Visibility never changes the hydraulic network.
    move(underKinds,document.querySelector('.mg-under-options'));
    underControls.append(underLabel,underKinds);$('bore-overlay').prepend(underControls);
    $('hm-under').onchange=()=>{
      $('mg-under').checked=$('hm-under').checked;underKinds.disabled=!$('hm-under').checked;
      $('mg-under').dispatchEvent(new Event('change',{bubbles:true}));ui.queue();
    };
    let sourceStamp='';
    function sync() {
      const s=window.__mf, active=s?.stage==='merge';
      bar.hidden=!active;document.body.classList.toggle('h-integration',active);
      underControls.hidden=!active;
      if(!active)return;
      $('hm-under').checked=$('mg-under').checked;
      underKinds.disabled=!$('mg-under').checked;
      underLabel.title=$('mg-under-note').textContent||'원본 도면을 반투명하게 겹쳐 봅니다. 계산·출력 좌표는 바뀌지 않습니다.';
      const busy=!$('busy').classList.contains('hidden'), d=s.merge || {}, summary=d.summary || {};
      const stale=!$('mg-stale').classList.contains('hidden') && !!$('mg-stale').textContent.trim();
      const split=!$('mg-split').classList.contains('hidden') && !!$('mg-split').textContent.trim();
      const state=busy?'busy':stale?'stale':split?'check':summary.merged?'review':'waiting';
      bar.dataset.state=state;
      text($('hm-state'),busy?'처리 중':stale?'변경사항 반영 필요':split?'연결 확인 필요':summary.merged?'결합 완료 · 검토 필요':'결합 전');
      for(const key of ['nodes','pipes','nozzles','valves'])text($('hm-'+key),summary.merged&&Number.isFinite(summary[key])?summary[key].toLocaleString('ko-KR'):'—');
      text(attachment,!summary.merged?'아직 결합하지 않았습니다.':summary.attached?'기계실 접속됨':'기계실 미접속 · 연결 조건을 확인하세요.');
      attachment.dataset.attached=String(!!summary.attached);
      // Read native material readiness; never promote an uploaded file to ready.
      const order=d.order || ['plan','system','machineroom'];
      const stamp=JSON.stringify([s.sid,order,d.ready,d.labels]);
      if(stamp!==sourceStamp){
        sourceStamp=stamp;cards.replaceChildren();
        order.forEach((key,i)=>{
          const ready=!!d.ready?.[key], row=make('li','hm-source');row.dataset.ready=String(ready);
          row.append(make('span','hm-source-index',String(i+1).padStart(2,'0')),make('strong','',d.labels?.[key]||({plan:'평면도',system:'계통도',machineroom:'기계실'})[key]||key),make('span','hm-source-state',ready?'준비됨':'미준비'));cards.append(row);
        });
      }
      const drawer=$('h-drawer'),pane=drawer.hidden?'':drawer.dataset.pane;
      for(const [button,p] of [['hm-connect','task'],['hm-sizing','sizing'],['hm-output','export'],['hm-view','merge-view']])$(button).setAttribute('aria-pressed',String(p===pane));
      $('hm-table').setAttribute('aria-pressed',String(!$('h-properties')?.hidden));
      for(const [button,target] of [['hm-table','h-table'],['hm-sizing','h-sizing'],['hm-output','h-export'],['hm-view','h-view']])$(button).disabled=$(target).disabled;
      // Busy protections are owned by the engine and Access; do not enable them.
      $('hm-connect').disabled=busy;sourceAction.disabled=busy;
      const proposal=$('h-output-basis').value==='proposal';
      text(outputNote,stale?'배관망이 변경되었습니다. 연결을 갱신한 뒤 출력하세요.':proposal
        ? $('h-output-basis').selectedOptions[0]?.disabled || $('sz-emit').disabled?'역산 조건이 변경되었거나 유효한 검토안이 없습니다. 다시 계산하세요.':'역산 검토안 출력 · 화면의 원본 관경은 변경하지 않습니다.'
        : '정의한 원본 배관망 출력 · 역산 제안값은 반영하지 않습니다.');
      // Hide irrelevant inputs, never reset their values or change solver defaults.
      field('sz-pressure').hidden=$('sz-mode').value!=='fixed_supply';
      field('sz-ceiling').hidden=$('sz-mode').value==='fixed_supply';
    }
    window.ModuleHIntegration={sync};
    const observer=new MutationObserver(()=>ui.queue());
    for(const id of ['sz-result','sz-emit','sz-download','h-properties','mg-modes','mg-ready','mg-split'])if($(id))observer.observe($(id),{childList:true,subtree:id==='sz-result',attributes:true,attributeFilter:['disabled','hidden','class']});
    window.addEventListener('module-h-sizing-result',()=>ui.queue());
    new ResizeObserver(()=>{document.body.style.setProperty('--hm-toolbar-height',bar.getBoundingClientRect().height+'px');}).observe(bar);
    sync();
  }
  if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',init,{once:true});else init();
})();
