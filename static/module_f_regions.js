"use strict";
// Pure geometry shared by drawing, candidate count and freehand capture. CAD mm.
(function (root) {
  const clone = zones => JSON.parse(JSON.stringify(zones || []));
  const points = z => Array.isArray(z)
    ? [[z[0],z[1]],[z[2],z[1]],[z[2],z[3]],[z[0],z[3]]] : z.points;
  function distance(p,a,b) {
    const dx=b[0]-a[0],dy=b[1]-a[1],len=dx*dx+dy*dy;
    const t=len?Math.max(0,Math.min(1,((p[0]-a[0])*dx+(p[1]-a[1])*dy)/len)):0;
    return Math.hypot(p[0]-a[0]-t*dx,p[1]-a[1]-t*dy);
  }
  function contains(z,p) {
    if (Array.isArray(z)) return z[0]<=p[0] && p[0]<=z[2] && z[1]<=p[1] && p[1]<=z[3];
    const ps=points(z); let inside=false;
    for(let i=0;i<ps.length;i++) {
      const a=ps[i],b=ps[(i+1)%ps.length];
      if(distance(p,a,b)<=1e-7) return true;
      if((a[1]>p[1])!==(b[1]>p[1]) && p[0]<(b[0]-a[0])*(p[1]-a[1])/(b[1]-a[1])+a[0]) inside=!inside;
    }
    return inside;
  }
  function area(z) {
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
  function finish(ps,scale) {
    if(!Number.isFinite(scale)||scale<=0) throw new Error("영역 배율이 올바르지 않습니다.");
    const cleaned=simplify(ps,1.2/scale);
    if(cleaned.length>1 && distance(cleaned[0],cleaned.at(-1),cleaned.at(-1))<1e-7) cleaned.pop();
    const z={type:"polygon",points:cleaned};
    if(cleaned.length<3 || area(z)*scale*scale<64) throw new Error("둘레를 충분히 넓게 그려주세요.");
    if(cleaned.length>512) throw new Error("영역이 너무 복잡합니다. 둘레를 더 간단히 그려주세요.");
    return z;
  }
  root.ModuleFRegions={clone,points,contains,area,finish};
})(typeof window!=="undefined"?window:globalThis);
