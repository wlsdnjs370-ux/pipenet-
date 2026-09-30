/* H-only Access navigation. Completion is server-derived, never click-derived. */
(() => {
  'use strict';
  const $=id=>document.getElementById(id), hub=$('h-access'), ui=window.ModuleHUI, native=window.ModuleHNative;
  if(!hub || !ui || !native)return;
  const roles=['plan','system','machineroom'], reduced=matchMedia('(prefers-reduced-motion: reduce)');
  let engaged=false, selected=false, moving=false, checking=false, state=null, stateSid=null, error='', timer=0, generation=0;
  let lastSid=null, lastBusy=false, pending=null;
  const busy=()=>!$('busy').classList.contains('hidden');
  const text=(node,value)=>{if(node.textContent!==value)node.textContent=value;};
  const dirty=()=>__mf.slot==='plan' && ['design','conv'].includes(__mf.stage) && (__mf.ovDirty || ui.unappliedSettings());
  const ready=role=>stateSid===__mf.sid && !!state?.groups?.[role] && !(role==='plan' && dirty());
  const canMerge=()=>!!__mf.sid && !!state?.can_merge && !error && !checking && !busy() && roles.every(ready);
  function render() {
    const locked=moving || busy(), count=roles.filter(ready).length;
    for(const role of roles) {
      const card=hub.querySelector(`.h-access-card[data-role="${role}"]`), item=state?.slots?.find(s=>s.kind===role);
      const active=selected && (role==='system' ? /^system\d*$/.test(__mf.slot) : __mf.slot===role) && __mf.stage!=='merge';
      card.classList.toggle('is-ready',ready(role));card.classList.toggle('is-active',active);
      const button=card.querySelector('.h-access-select');button.disabled=locked;button.setAttribute('aria-current',String(active));
      text(card.querySelector('.h-access-status'),ready(role)?'✓ 준비 완료':active?'● 작업 중':'대기');
      const file=card.querySelector('.h-access-file');if(file){text(file,item?.key || 'DXF 도면 선택');file.title=item?.key || '';}
      button.title=item?.reason || '도면을 선택하고 작업을 시작합니다';
      hub.querySelector(`.h-access-wires [data-role="${role}"]`).classList.toggle('is-ready',ready(role));
    }
    $('h-access-add').disabled=locked || !__mf.sid || !state?.can_add_system;
    $('h-access-add').title=!__mf.sid?'먼저 도면을 한 장 업로드하세요':state?.can_add_system?'추가 계통도도 경로 추출 후 통합합니다':'계통도는 최대 10장입니다';
    $('h-access-size').disabled=locked;
    $('h-access-upload').hidden=!selected || !__mf.world;
    $('h-access-upload').disabled=locked;
    $('h-access-merge').disabled=locked || !canMerge();
    text($('h-access-merge').querySelector('span'),`준비 ${count} / 3`);
    text($('h-access-note'),error || (checking?'준비 상태 확인 중…':count===3?'세 공정 준비 완료 · 통합 가능':'세 공정이 준비되면 통합할 수 있습니다.'));
    $('h-access-note').classList.toggle('is-error',!!error);
    for(const id of ['h-merge','btn-merge'])if($(id))$(id).disabled=locked || !canMerge();
    // Do not overwrite the native build button's own disabled state while a
    // read is in flight. The capture guard below enforces H's additional gate.
    $('mg-build').classList.toggle('h-access-blocked',!canMerge());
    const signature=JSON.stringify([state?.slots?.filter(s=>s.role==='system'),__mf.slot,locked]);
    const list=$('h-access-systems');
    if(list.dataset.signature!==signature) {
      list.dataset.signature=signature;list.replaceChildren();
      const systems=state?.slots?.filter(s=>s.role==='system') || [];
      list.hidden=systems.length<2;
      for(const item of systems) {
        const row=document.createElement('div');row.className='h-access-system-item';
        const choose=document.createElement('button');choose.type='button';choose.textContent=`${item.ready?'✓ ':''}${item.label}`;
        choose.dataset.slot=item.kind;choose.disabled=locked;choose.title=item.key || item.reason;
        choose.classList.toggle('is-ready',item.ready);choose.setAttribute('aria-current',String(__mf.slot===item.kind));row.append(choose);
        if(item.removable) {const remove=document.createElement('button');remove.type='button';remove.textContent='×';remove.dataset.remove=item.kind;remove.disabled=locked;remove.setAttribute('aria-label',`${item.label} 삭제`);row.append(remove);}
        list.append(row);
      }
    }
    ui.queue();
  }
  async function refresh() {
    const sid=__mf.sid;
    if(!sid){state=null;stateSid=null;error='';render();return;}
    if(busy())return;
    if(pending && stateSid===sid)return pending;
    const token=++generation;stateSid=sid;checking=true;render();
    pending=(async()=>{
      try {
        const response=await fetch(`/api/module-h/access-state?sid=${encodeURIComponent(sid)}`,{cache:'no-store'});
        if(!response.ok)throw Error(response.status===404?'Access 적용을 위해 서버를 재시작해 주세요.':'도면 준비 상태를 확인할 수 없습니다. Access를 다시 열어 재시도하세요.');
        const data=await response.json();if(!data.ok || !Array.isArray(data.slots))throw Error('도면 준비 상태 응답을 확인할 수 없습니다.');
        if(token===generation && sid===__mf.sid){state=data;error='';}
      } catch(e) {if(token===generation && sid===__mf.sid){state=null;error=e.message;}}
      finally {if(token===generation){checking=false;pending=null;render();}}
    })();
    return pending;
  }
  const frame=()=>new Promise(resolve=>requestAnimationFrame(resolve));
  async function animate(element,frames,options) {
    if(reduced.matches)return;
    const animation=element.animate(frames,options);
    try{await animation.finished;}catch(_){}finally{animation.cancel();}
  }
  async function resize(mode) {
    if(hub.dataset.mode===mode)return;
    const from=hub.getBoundingClientRect();hub.dataset.mode=mode;
    text($('h-access-size'),mode==='docked'?'확대 ↗':'접기 ↘');
    $('h-access-size').setAttribute('aria-label',mode==='docked'?'입력 도면 회로 확대':'입력 도면 회로 접기');
    const to=hub.getBoundingClientRect(), base=mode==='docked'?'none':'translate(-50%,-50%)';
    await animate(hub,[{transformOrigin:'0 0',transform:`${base==='none'?'':base} translate(${from.x-to.x}px,${from.y-to.y}px) scale(${from.width/to.width},${from.height/to.height})`},{transformOrigin:'0 0',transform:base}],{duration:760,easing:'cubic-bezier(.22,1,.36,1)'});
  }
  async function open() {
    if(moving || busy())return;
    ui.closeDrawer(false);engaged=true;moving=true;render();
    if(hub.hidden) {
      hub.hidden=false;hub.dataset.mode='expanded';
      const cards=[...hub.querySelectorAll('.h-access-card')];
      if(!reduced.matches)cards.forEach(c=>c.style.opacity='0');
      await animate(hub,[{clipPath:'inset(0 calc(50% - 1px))'},{clipPath:'inset(0 0)'}],{duration:650,easing:'cubic-bezier(.76,0,.24,1)'});
      await Promise.all(cards.map((card,index)=>{
        card.style.opacity='';return animate(card,[{opacity:0,transform:'translateY(12px)'},{opacity:1,transform:'none'}],{duration:280,delay:index*110,fill:'backwards',easing:'ease-out'});
      }));
    } else await resize('expanded');
    moving=false;render();$('h-access-plan').focus({preventScroll:true});await refresh();
  }
  async function select(kind) {
    if(moving || busy())return;
    moving=true;engaged=true;error='';render();ui.closeDrawer(false);
    try {
      await resize('docked');
      if(!await native.selectSlot(kind))throw Error('도면 전환을 완료하지 못했습니다. 처리 내역을 확인하세요.');
      selected=true;
      // H closes drawers on a stage/slot change. Open after that sync frame.
      ui.queue();await frame();await frame();ui.task();
    } catch(e){error=e.message;}
    finally{moving=false;await refresh();render();}
  }
  async function systemAction(removeKind) {
    if(moving || busy() || !__mf.sid)return;
    moving=true;render();ui.closeDrawer(false);
    try {
      if(removeKind) {
        const item=state?.slots?.find(s=>s.kind===removeKind);if(item)await native.removeSystem({...item,active:__mf.slot===item.kind});
      } else {
        await resize('docked');await native.addSystem();selected=true;
        ui.queue();await frame();await frame();ui.task();
      }
    } finally {moving=false;await refresh();render();}
  }
  async function merge() {
    if(moving || busy())return;
    moving=true;render();
    try {await refresh();if(!canMerge())return;ui.closeDrawer(false);await resize('docked');await native.navigate('merge');ui.queue();await frame();await frame();ui.task();}
    finally {moving=false;render();}
  }
  hub.addEventListener('click',event=>{
    const button=event.target.closest('button');if(!button || button.disabled)return;
    if(button.dataset.slot)select(button.dataset.slot);
    if(button.dataset.remove)systemAction(button.dataset.remove);
  });
  $('h-access-add').onclick=()=>systemAction();$('h-access-merge').onclick=merge;
  $('h-access-upload').onclick=async()=>{
    if(moving || busy())return;
    moving=true;render();ui.closeDrawer(false);
    try {await resize('docked');await native.navigate('open');ui.queue();await frame();await frame();ui.task();}
    finally {moving=false;render();}
  };
  $('h-access-size').onclick=async()=>{
    if(moving || busy())return;
    if(hub.dataset.mode==='docked'){await open();return;}
    if(!selected){hub.hidden=true;engaged=false;render();$('h-welcome-open').focus();return;}
    moving=true;render();await resize('docked');moving=false;render();
  };
  // Other H entry points may expose the native merge tab. Apply the same gate.
  document.addEventListener('click',event=>{
    const button=event.target.closest('#btn-merge,#mg-build');
    if(!button)return;
    if(button.id==='btn-merge'){event.preventDefault();event.stopImmediatePropagation();merge();}
    else if(!canMerge()){event.preventDefault();event.stopImmediatePropagation();}
  },true);
  function observe() {
    const nowBusy=busy(), sid=__mf.sid;
    if(nowBusy && !lastBusy){generation++;pending=null;checking=false;}
    if(sid!==lastSid){lastSid=sid;state=null;stateSid=null;generation++;pending=null;checking=false;}
    if(sid && !engaged){engaged=true;selected=true;hub.hidden=false;hub.dataset.mode='docked';text($('h-access-size'),'확대 ↗');}
    const finished=lastBusy && !nowBusy;lastBusy=nowBusy;render();
    if(nowBusy)return;
    clearTimeout(timer);timer=setTimeout(refresh,finished?80:140);
  }
  const observer=new MutationObserver(observe);
  observer.observe($('busy'),{attributes:true,attributeFilter:['class']});
  for(const id of ['slots','steps','dg-stale','status'])observer.observe($(id),{childList:true,subtree:true,characterData:true});
  document.addEventListener('change',()=>{render();clearTimeout(timer);timer=setTimeout(refresh,160);});
  window.addEventListener('keydown',event=>{if(event.key==='Escape' && !hub.hidden && hub.dataset.mode==='expanded' && !moving){$('h-access-size').click();event.stopPropagation();}},true);
  window.ModuleHAccess={open,merge,refresh,get engaged(){return engaged;},get selected(){return selected;},get canMerge(){return canMerge();}};
  render();
})();
