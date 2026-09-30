/* H-only window geometry. Moving/resizing never changes network coordinates. */
(() => {
  'use strict';
  const registered=new WeakSet(), margin=8;
  function register(panel,head,{minWidth=280,minHeight=140}={}) {
    if(!panel||!head||registered.has(panel))return;
    registered.add(panel);panel.classList.add('h-window');
    const saved=panel.getAttribute('style')||'';
    let gesture=null;
    const set=(key,value)=>panel.style.setProperty(key,value,'important');
    function freeze(){
      const r=panel.getBoundingClientRect();
      panel.dataset.hDragged='true';panel.dataset.dragged='true';
      for(const [key,value] of Object.entries({position:'fixed',left:r.left+'px',top:r.top+'px',
        width:r.width+'px',height:r.height+'px',right:'auto',bottom:'auto',margin:'0',
        'box-sizing':'border-box','max-width':'calc(100vw - 16px)','max-height':'calc(100dvh - 16px)'}))set(key,value);
    }
    function clamp(x,y){
      const r=panel.getBoundingClientRect();
      set('left',Math.max(margin,Math.min(x,innerWidth-r.width-margin))+'px');
      set('top',Math.max(margin,Math.min(y,innerHeight-r.height-margin))+'px');
    }
    function start(e,edge=''){
      if(e.button!==0||(!edge&&e.target.closest('button,a,input,select,summary')))return;
      freeze();const r=panel.getBoundingClientRect();
      gesture={edge,x:e.clientX,y:e.clientY,r};e.currentTarget.setPointerCapture(e.pointerId);
      document.body.classList.add(edge?'h-window-resizing':'h-window-dragging');e.preventDefault();e.stopPropagation();
    }
    function move(e){
      if(!gesture)return;
      const {edge,x,y,r}=gesture,dx=e.clientX-x,dy=e.clientY-y;
      if(!edge){clamp(r.left+dx,r.top+dy);return;}
      const mw=Math.min(minWidth,innerWidth-16),mh=Math.min(minHeight,innerHeight-16);
      let l=r.left,t=r.top,rr=r.right,b=r.bottom;
      if(edge.includes('w'))l=Math.max(margin,Math.min(rr-mw,l+dx));
      if(edge.includes('e'))rr=Math.min(innerWidth-margin,Math.max(l+mw,rr+dx));
      if(edge.includes('n'))t=Math.max(margin,Math.min(b-mh,t+dy));
      if(edge.includes('s'))b=Math.min(innerHeight-margin,Math.max(t+mh,b+dy));
      set('left',l+'px');set('top',t+'px');set('width',(rr-l)+'px');set('height',(b-t)+'px');
      panel.dispatchEvent(new CustomEvent('h-window-resize'));
    }
    function end(){gesture=null;document.body.classList.remove('h-window-dragging','h-window-resizing');}
    function bind(handle,edge){
      handle.addEventListener('pointerdown',e=>start(e,edge));handle.addEventListener('pointermove',move);
      for(const type of ['pointerup','pointercancel','lostpointercapture'])handle.addEventListener(type,end);
    }
    head.dataset.hDragHandle='true';head.tabIndex=0;
    head.title='드래그하여 이동 · 방향키로 이동 · 더블클릭으로 위치·크기 초기화';bind(head,'');
    head.addEventListener('dblclick',e=>{if(e.target.closest('button,a,input,select'))return;
      panel.setAttribute('style',saved);delete panel.dataset.hDragged;delete panel.dataset.dragged;
      panel.dispatchEvent(new CustomEvent('h-window-resize'));
    });
    head.addEventListener('keydown',e=>{
      if(e.target!==head||!['ArrowLeft','ArrowRight','ArrowUp','ArrowDown'].includes(e.key))return;
      e.preventDefault();e.stopPropagation();freeze();const r=panel.getBoundingClientRect(),d=e.shiftKey?40:10;
      clamp(r.left+(e.key==='ArrowLeft'?-d:e.key==='ArrowRight'?d:0),r.top+(e.key==='ArrowUp'?-d:e.key==='ArrowDown'?d:0));
    });
    for(const edge of ['n','e','s','w','ne','nw','se','sw']){
      const h=document.createElement('span');h.className='h-resize-edge h-resize-'+edge;
      h.dataset.hResizeHandle='true';h.dataset.edge=edge;h.setAttribute('aria-hidden','true');
      panel.append(h);bind(h,edge);
    }
    new ResizeObserver(()=>{if(panel.dataset.hDragged&&panel.getClientRects().length&&!gesture){const r=panel.getBoundingClientRect();clamp(r.left,r.top);}}).observe(panel);
    window.addEventListener('resize',()=>{if(panel.dataset.hDragged){const r=panel.getBoundingClientRect();clamp(r.left,r.top);}});
  }
  window.ModuleHWindows={register};
  for(const [id,selector,inner] of [['h-drawer','.h-drawer-head'],['h-activity','.h-activity-head'],
    ['dg-ins',':scope > header'],['ne-panel','.ne-pop-head'],['ne-history-dialog',':scope > h3'],['conv-modal',':scope > header','.sheet']]){
    const container=document.getElementById(id),panel=inner?container?.querySelector(inner):container;
    register(panel,panel?.querySelector(selector));
  }
})();
