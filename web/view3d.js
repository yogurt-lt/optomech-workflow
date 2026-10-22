/* Offline, data-driven solid-of-revolution preview. No external dependencies. */
(() => {
  const canvas=document.getElementById('assembly3d'),ctx=canvas.getContext('2d');
  let meshes=[],bounds,drag,last,queued=false,yaw=.48,pitch=-.42,zoom=1;
  const colors={lens:[64,158,223],barrel:[71,87,107],spacer:[89,174,117],retainer:[235,147,62],adapter:[161,120,201]};
  const kind=name=>/lens/i.test(name)?'lens':/retainer/i.test(name)?'retainer':/spacer/i.test(name)?'spacer':/barrel/i.test(name)?'barrel':'adapter';
  function simplify(points,tolerance){
    if(points.length<5)return points;
    const keep=new Set([0,points.length-1]),stack=[[0,points.length-1]];
    while(stack.length){const [a,b]=stack.pop(),p=points[a],q=points[b],dx=q[0]-p[0],dy=q[1]-p[1],den=dx*dx+dy*dy;let best=tolerance*tolerance,j=-1;
      for(let i=a+1;i<b;i++){const v=points[i],t=den?Math.max(0,Math.min(1,((v[0]-p[0])*dx+(v[1]-p[1])*dy)/den)):0,d=(v[0]-p[0]-t*dx)**2+(v[1]-p[1]-t*dy)**2;if(d>best){best=d;j=i;}}
      if(j>=0){keep.add(j);stack.push([a,j],[j,b]);}
    }return [...keep].sort((a,b)=>a-b).map(i=>points[i]);
  }
  function rotate([x,y,z]){const a=x*Math.cos(yaw)+z*Math.sin(yaw),b=-x*Math.sin(yaw)+z*Math.cos(yaw);return [a,y*Math.cos(pitch)-b*Math.sin(pitch),y*Math.sin(pitch)+b*Math.cos(pitch)];}
  function point(z,r,a){return [z-bounds.center,r*Math.cos(a),r*Math.sin(a)];}
  function build(plan){
    const pts=plan.parts.flatMap(p=>p.profile_zr_mm),all=pts.concat((plan.lenses||[]).flatMap(l=>l.polygon_zr_mm||[]));
    let zmin=Infinity,zmax=-Infinity,rmax=0;for(const [z,r] of all){zmin=Math.min(zmin,z);zmax=Math.max(zmax,z);rmax=Math.max(rmax,r);}
    bounds={center:(zmin+zmax)/2,size:Math.max(zmax-zmin,2*rmax),rmax};
    meshes=plan.parts.map(p=>({name:p.name,kind:kind(p.name),profile:p.profile_zr_mm,color:p.color_rgb?p.color_rgb.map(v=>Math.round(v*255)):colors[kind(p.name)],role:p.role||''}));
    for(const [i,l] of (plan.lenses||[]).entries())if(l.polygon_zr_mm?.length)meshes.push({name:'L'+(l.id||i+1),kind:'lens',profile:simplify(l.polygon_zr_mm,bounds.size*.00004)});
    document.getElementById('meshStats').textContent=window.i18n?.language==='en'?`${meshes.length} revolved bodies · follows profile edits`:`${meshes.length} 个旋转体 · 随截面同步`;
    const legend=document.getElementById('legend3d'),groups=new Map();
    for(const mesh of meshes){
      const col=mesh.color||colors[mesh.kind],key=col.join(','),name=mesh.name.toLowerCase();
      let label=mesh.kind==='lens'?'镜片':mesh.kind==='retainer'?'压环':mesh.kind==='spacer'?'隔圈':/cell/.test(name)?'内部镜框':mesh.kind==='barrel'?(/front/.test(name)?'前筒':/middle/.test(name)?'中筒':/rear/.test(name)?'后筒':'镜筒'):'定位套';
      if(mesh.kind==='barrel'&&/stop/i.test(mesh.role))label+=' / 光阑';
      if(!groups.has(key))groups.set(key,{col,labels:new Set()});groups.get(key).labels.add(label);
    }
    legend.replaceChildren(...[...groups.values()].map(({col,labels})=>{const span=document.createElement('span'),i=document.createElement('i');i.style.background=`rgb(${col.join(',')})`;span.append(i,document.createTextNode([...labels].join('、')));return span}));
    schedule();
  }
  function draw(){
    queued=false;if(!bounds)return;
    const w=canvas.clientWidth,h=canvas.clientHeight,dpr=Math.min(devicePixelRatio||1,2);canvas.width=w*dpr;canvas.height=h*dpr;ctx.setTransform(dpr,0,0,dpr,0,0);
    const bg=ctx.createLinearGradient(0,0,w,h);bg.addColorStop(0,'#f7fbfd');bg.addColorStop(1,'#dfe8f0');ctx.fillStyle=bg;ctx.fillRect(0,0,w,h);
    const shadow=ctx.createRadialGradient(w*.51,h*.78,0,w*.51,h*.78,w*.37);shadow.addColorStop(0,'#30435924');shadow.addColorStop(1,'#30435900');ctx.save();ctx.translate(0,h*.54);ctx.scale(1,.32);ctx.fillStyle=shadow;ctx.fillRect(0,0,w,h*2);ctx.restore();
    const scale=Math.min(w*.78,h*1.12)/bounds.size*zoom,cut=document.getElementById('section3d').checked,hide=document.getElementById('hideBarrel3d').checked,faces=[];
    const project=v=>{const f=1/(1-v[2]/(bounds.size*6)),x=v[0]*f,y=v[1]*f,tilt=-.26;return [w*.5+(x*Math.cos(tilt)-y*Math.sin(tilt))*scale,h*.5-(x*Math.sin(tilt)+y*Math.cos(tilt))*scale];};
    function face(vertices,color,cap=false,lens=false){
      const v=vertices.map(rotate),depth=v.reduce((a,p)=>a+p[2],0)/v.length;
      const a=v[0],b=v[1],c=v[2],u=b.map((n,i)=>n-a[i]),t=c.map((n,i)=>n-a[i]),n=[u[1]*t[2]-u[2]*t[1],u[2]*t[0]-u[0]*t[2],u[0]*t[1]-u[1]*t[0]],len=Math.hypot(...n);
      if(!cap&&len<1e-10)return;
      const light=len?Math.abs((n[0]*-.25+n[1]*.65+n[2]*.72)/len):.6,shade=cap?.95:.48+.50*light;
      faces.push({xy:v.map(project),depth,fill:`rgb(${color.map(c=>Math.round(c*shade)).join(',')})`,cap,lens});
    }
    for(const mesh of meshes){if(hide&&mesh.kind==='barrel')continue;const p=mesh.profile,col=mesh.color||colors[mesh.kind],start=cut?Math.PI:0,arc=cut?Math.PI:Math.PI*2,steps=cut?48:80;
      for(let i=0;i<p.length;i++){const a=p[i],b=p[(i+1)%p.length];for(let j=0;j<steps;j++){const t=start+j*arc/steps,q=start+(j+1)*arc/steps;face([point(...a,t),point(...b,t),point(...b,q),point(...a,q)],col,false,mesh.kind==='lens');}}
      if(cut){face(p.map(([z,r])=>point(z,r,start)),col,true,mesh.kind==='lens');face(p.map(([z,r])=>point(z,r,start+arc)).reverse(),col,true,mesh.kind==='lens');}
    }
    faces.sort((a,b)=>a.depth-b.depth);
    for(const f of faces){ctx.beginPath();f.xy.forEach(([x,y],i)=>i?ctx.lineTo(x,y):ctx.moveTo(x,y));ctx.closePath();ctx.fillStyle=f.fill;ctx.globalAlpha=f.lens&&!cut?.68:1;ctx.fill();if(f.cap){ctx.strokeStyle='#23394b99';ctx.lineWidth=.75;ctx.stroke();}}
    ctx.globalAlpha=1;ctx.fillStyle='#526b7d';ctx.font='12px Microsoft YaHei, sans-serif';ctx.fillText(window.i18n?.t('Z  光轴')||'Z  光轴',22,h-22);
  }
  function schedule(){if(!queued){queued=true;requestAnimationFrame(draw);}}
  canvas.addEventListener('pointerdown',e=>{drag=true;last=[e.clientX,e.clientY];canvas.setPointerCapture(e.pointerId);});
  canvas.addEventListener('pointermove',e=>{if(!drag)return;yaw+=(e.clientX-last[0])*.008;pitch=Math.max(-1.5,Math.min(1.5,pitch+(e.clientY-last[1])*.008));last=[e.clientX,e.clientY];schedule();});
  canvas.addEventListener('pointerup',()=>drag=false);canvas.addEventListener('pointercancel',()=>drag=false);
  canvas.addEventListener('wheel',e=>{e.preventDefault();zoom=Math.max(.45,Math.min(3,zoom*Math.exp(-e.deltaY*.001)));schedule();},{passive:false});
  document.getElementById('section3d').onchange=schedule;document.getElementById('hideBarrel3d').onchange=schedule;
  document.getElementById('reset3d').onclick=()=>{yaw=.48;pitch=-.42;zoom=1;schedule();};
  new ResizeObserver(schedule).observe(canvas);window.optomech3d={update:build};
})();
