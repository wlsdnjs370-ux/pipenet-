"use strict";
// Pure geometry shared by drawing, candidate count and freehand capture. CAD mm.
(function (root) {
  const clone = zones => JSON.parse(JSON.stringify(zones || []));
  const points = z => Array.isArray(z)
    ? [[z[0],z[1]],[z[2],z[1]],[z[2],z[3]],[z[0],z[3]]] : z.points;
  const enclosedCache=new WeakMap(), EPS=1e-7;
  const cross=(a,b)=>a[0]*b[1]-a[1]*b[0];
  // Same planar-face rule as graph/freehand.py. Used only for completed pen
  // crops, not on every pointermove. No hull/box expansion or even-odd holes.
  function enclosed(z) {
    if(enclosedCache.has(z))return enclosedCache.get(z);
    const ps=points(z),[ox,oy]=ps[0],local=ps.map(([x,y])=>[x-ox,y-oy]);
    const sides=local.map((p,i)=>[p,local[(i+1)%local.length]])
      .filter(([a,b])=>Math.hypot(b[0]-a[0],b[1]-a[1])>EPS);
    const splits=sides.map(([a,b])=>[[0,a],[1,b]]);
    for(let i=0;i<sides.length;i++){
      const [a,b]=sides[i],v=[b[0]-a[0],b[1]-a[1]],lv=Math.hypot(...v);
      for(let j=0;j<i;j++){
        const [c,d]=sides[j];
        if(Math.max(a[0],b[0])+EPS<Math.min(c[0],d[0])||Math.max(c[0],d[0])+EPS<Math.min(a[0],b[0])||
           Math.max(a[1],b[1])+EPS<Math.min(c[1],d[1])||Math.max(c[1],d[1])+EPS<Math.min(a[1],b[1]))continue;
        const w=[d[0]-c[0],d[1]-c[1]],q=[c[0]-a[0],c[1]-a[1]],lw=Math.hypot(...w),det=cross(v,w);
        if(Math.abs(det)>1e-12*lv*lw){
          let t=cross(q,w)/det,u=cross(q,v)/det;
          if(t>=-EPS/lv&&t<=1+EPS/lv&&u>=-EPS/lw&&u<=1+EPS/lw){
            t=Math.max(0,Math.min(1,t));u=Math.max(0,Math.min(1,u));
            const p=[a[0]+t*v[0],a[1]+t*v[1]];splits[i].push([t,p]);splits[j].push([u,p]);
          }
        }else if(Math.abs(cross(q,v))<=EPS*lv){
          for(const [index,start,vector,length,ends] of [[i,a,v,lv,[c,d]],[j,c,w,lw,[a,b]]]){
            for(const p of ends){const t=((p[0]-start[0])*vector[0]+(p[1]-start[1])*vector[1])/length**2;
              if(t>=-EPS/length&&t<=1+EPS/length)splits[index].push([Math.max(0,Math.min(1,t)),p]);}
          }
        }
      }
    }
    const vertices=[],ids=new Map(),neighbors=new Map();
    function vertex(p){const key=p.map(x=>Math.round(x*1e6)).join(',');
      if(!ids.has(key)){ids.set(key,vertices.length);vertices.push(p);}return ids.get(key);}
    function connect(a,b){if(!neighbors.has(a))neighbors.set(a,new Set());neighbors.get(a).add(b);}
    for(const split of splits){const order=split.sort((a,b)=>a[0]-b[0]).map(([,p])=>vertex(p));
      for(let i=1;i<order.length;i++)if(order[i]!==order[i-1]){connect(order[i-1],order[i]);connect(order[i],order[i-1]);}}
    const next=new Map(),edgeKey=(a,b)=>`${a},${b}`;
    for(const [b,ns] of neighbors){const around=[...ns].sort((a,c)=>
      Math.atan2(vertices[a][1]-vertices[b][1],vertices[a][0]-vertices[b][0])-
      Math.atan2(vertices[c][1]-vertices[b][1],vertices[c][0]-vertices[b][0]));
      around.forEach((a,i)=>next.set(edgeKey(a,b),[b,around[(i+around.length-1)%around.length]]));}
    const visited=new Set(),faces=[],counts=new Map();let total=0;
    for(const start of next.keys()){
      if(visited.has(start))continue;
      let edge=start;const walk=[];
      while(!visited.has(edge)){visited.add(edge);const [a,b]=edge.split(',').map(Number);
        walk.push([a,b]);edge=edgeKey(...next.get(edge));}
      if(edge!==start)throw Error('펜 영역의 둘레를 해석하지 못했습니다. 다시 그려주세요.');
      const ring=walk.map(([a])=>vertices[a]),[x0,y0]=ring[0];let sum=0;
      ring.forEach((a,i)=>{const b=ring[(i+1)%ring.length];sum+=(a[0]-x0)*(b[1]-y0)-(b[0]-x0)*(a[1]-y0);});
      if(sum>2e-8){faces.push(ring.map(([x,y])=>[x+ox,y+oy]));total+=sum/2;
        for(const [a,b] of walk){const key=edgeKey(Math.min(a,b),Math.max(a,b));counts.set(key,(counts.get(key)||0)+1);}}
    }
    const boundary=[...counts].filter(([,n])=>n===1).map(([key])=>key.split(',').map(i=>[vertices[i][0]+ox,vertices[i][1]+oy]));
    const result={faces,boundary,area:total};enclosedCache.set(z,result);return result;
  }
  function distance(p,a,b) {
    const dx=b[0]-a[0],dy=b[1]-a[1],len=dx*dx+dy*dy;
    const t=len?Math.max(0,Math.min(1,((p[0]-a[0])*dx+(p[1]-a[1])*dy)/len)):0;
    return Math.hypot(p[0]-a[0]-t*dx,p[1]-a[1]-t*dy);
  }
  function contains(z,p) {
    if (Array.isArray(z)) return z[0]<=p[0] && p[0]<=z[2] && z[1]<=p[1] && p[1]<=z[3];
    if(z.fill_rule==='enclosed')return enclosed(z).faces.some(ring=>contains({points:ring},p));
    const ps=points(z); let inside=false;
    for(let i=0;i<ps.length;i++) {
      const a=ps[i],b=ps[(i+1)%ps.length];
      if(distance(p,a,b)<=1e-7) return true;
      if((a[1]>p[1])!==(b[1]>p[1]) && p[0]<(b[0]-a[0])*(p[1]-a[1])/(b[1]-a[1])+a[0]) inside=!inside;
    }
    return inside;
  }
  function area(z) {
    if(z.fill_rule==='enclosed')return enclosed(z).area;
    const ps=points(z),[ox,oy]=ps[0]; let sum=0;
    for(let i=0;i<ps.length;i++) {
      const a=ps[i],b=ps[(i+1)%ps.length];
      sum+=(a[0]-ox)*(b[1]-oy)-(b[0]-ox)*(a[1]-oy);
    }
    return Math.abs(sum)/2;
  }
  function simplify(ps,epsilon) {
    if(ps.length<3) return ps.map(p=>p.slice());
    const keep=new Set([0,ps.length-1]), stack=[[0,ps.length-1]];
    while(stack.length) {
      const [a,b]=stack.pop(); let far=epsilon,index=-1;
      for(let i=a+1;i<b;i++) { const d=distance(ps[i],ps[a],ps[b]); if(d>far){far=d;index=i;} }
      if(index!==-1){keep.add(index);stack.push([a,index],[index,b]);}
    }
    return [...keep].sort((a,b)=>a-b).map(i=>ps[i].slice());
  }
  function finish(ps,scale,{enclose=false}={}) {
    if(!Number.isFinite(scale)||scale<=0) throw new Error("영역 배율이 올바르지 않습니다.");
    const cleaned=simplify(ps,1.2/scale);
    if(cleaned.length>1 && distance(cleaned[0],cleaned.at(-1),cleaned.at(-1))<1e-7) cleaned.pop();
    if(cleaned.length>512) throw new Error("영역이 너무 복잡합니다. 둘레를 더 간단히 그려주세요.");
    const z={type:"polygon",points:cleaned,...(enclose?{fill_rule:'enclosed'}:{})};
    if(cleaned.length<3 || area(z)*scale*scale<64) throw new Error("둘레를 충분히 넓게 그려주세요.");
    return z;
  }
  root.ModuleFRegions={clone,points,contains,area,finish,enclosed};
})(typeof window!=="undefined"?window:globalThis);
