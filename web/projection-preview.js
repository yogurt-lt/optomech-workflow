(()=>{'use strict';const $=id=>document.getElementById(id),c=$('scene'),gl=c.getContext('webgl2',{antialias:true,alpha:true});let en=new URLSearchParams(location.search).get('lang')==='en',playing=false,angle=45,dir=1,yaw=.32,pitch=-.34,scale=85,last=0,meshes=[],law,viewMode="all",focus=0,focusPlaying=false,focusDir=1;
if(!gl){$('error').textContent='WebGL2 unavailable';return}const shader=(type,src)=>{const s=gl.createShader(type);gl.shaderSource(s,src);gl.compileShader(s);if(!gl.getShaderParameter(s,gl.COMPILE_STATUS))throw Error(gl.getShaderInfoLog(s));return s};
const program=gl.createProgram();gl.attachShader(program,shader(gl.VERTEX_SHADER,`#version 300 es
in vec3 position;in vec3 normal;uniform float rot,dz,pivot,yaw,pitch,scale,aspect,lineMode;out float light;
vec3 rz(vec3 p){float a=cos(rot),b=sin(rot);return vec3(a*p.x-b*p.y,b*p.x+a*p.y,p.z);}
vec3 view(vec3 p){p=p.zyx;p.x+=62.;float a=cos(yaw),b=sin(yaw);p=vec3(a*p.x+b*p.z,p.y,-b*p.x+a*p.z);a=cos(pitch);b=sin(pitch);return vec3(p.x,a*p.y-b*p.z,b*p.y+a*p.z);}
void main(){vec3 p=position-vec3(pivot,0,0);p=rz(p)+vec3(pivot,0,dz);vec3 v=view(p);gl_Position=vec4(v.x/scale/aspect,v.y/scale,-v.z/1000.,1);vec3 n=rz(normal);light=lineMode>.5?1.:.55+.4*abs(dot(normalize(n),normalize(vec3(-.4,.6,.8))));}`));gl.attachShader(program,shader(gl.FRAGMENT_SHADER,`#version 300 es
precision highp float;uniform vec3 color;uniform float alpha;in float light;out vec4 outColor;void main(){outColor=vec4(color*light,alpha);}`));gl.linkProgram(program);if(!gl.getProgramParameter(program,gl.LINK_STATUS))throw Error(gl.getProgramInfoLog(program));gl.useProgram(program);
const u=Object.fromEntries(['rot','dz','pivot','yaw','pitch','scale','aspect','color','alpha','lineMode'].map(k=>[k,gl.getUniformLocation(program,k)]));
function mesh(m){const arr=[];for(const tri of m.triangles){const [a,b,c]=tri.map(i=>m.vertices[i]),v=b.map((x,i)=>x-a[i]),w=c.map((x,i)=>x-a[i]),n=[v[1]*w[2]-v[2]*w[1],v[2]*w[0]-v[0]*w[2],v[0]*w[1]-v[1]*w[0]];let l=Math.hypot(...n)||1;for(const p of [a,b,c])arr.push(...p,...n.map(x=>x/l))}m.vao=gl.createVertexArray();gl.bindVertexArray(m.vao);const b=gl.createBuffer();gl.bindBuffer(gl.ARRAY_BUFFER,b);gl.bufferData(gl.ARRAY_BUFFER,new Float32Array(arr),gl.STATIC_DRAW);for(const [key,offset]of[['position',0],['normal',12]]){let loc=gl.getAttribLocation(program,key);gl.enableVertexAttribArray(loc);gl.vertexAttribPointer(loc,3,gl.FLOAT,false,24,offset)}m.count=arr.length/6;delete m.vertices;delete m.triangles;return m}

function travel(g,a){const q=law.coefficients[g].find(q=>a<=q.hi)||law.coefficients[g].at(-1),t=a-q.lo;return q.coefficients.reduce((v,c,i)=>v+c*t**i,0)}
const cols={a:[.25,.70,1],b:[.30,.76,1],c:[.20,.65,1]},trackCols={a:[.96,.10,.04],b:[.65,.08,.89],c:[.04,.55,.4]};
// Favor the optical train while retaining the sleeve and drive context.
function opacityFor(m){
 if(m.glass)return +$('lens-opacity').value;
 if(m.overlay)return .98;
 if(['cam','guide'].includes(m.group))return +$('opacity').value;
 if(m.group==='focus')return .20;
 if(/Cell$/.test(m.name))return .14;
 if(/Spacer/.test(m.name))return .24;
 if(/Retain|Collar|Shim/.test(m.name))return .28;
 if(/BlackStopFoil/.test(m.name))return .65;
 if(/Axle|Roller|Dowel|Button/.test(m.name))return .65;
 return .42;
}
function tracks(g,offset){let vertices=[],triangles=[];for(let a=0;a<=90;a++){const q=law.coefficients[g].find(q=>a<=q.hi)||law.coefficients[g].at(-1),t=a-q.lo,der=q.coefficients.slice(1).reduce((v,c,i)=>v+c*(i+1)*t**i,0)*180/Math.PI,beta=-a*Math.PI/180,z=law.groups[g]+travel(g,a),den=Math.hypot(29.7,der);for(const v of [offset-.16,offset+.16])vertices.push([31.4*Math.cos(beta)-v*der/den*Math.sin(beta),31.4*Math.sin(beta)+v*der/den*Math.cos(beta),z+v*29.7/den]);if(a){let k=2*a;triangles.push([k-2,k-1,k],[k-1,k+1,k])}}return mesh({name:'Overlay'+g+offset,group:'cam',overlay:true,color:trackCols[g],vertices,triangles})}
function pose(m){const g=m.group;return {rot:(g==='cam'?angle:g==='focus'?focus:g==='zp'?-8*angle:g==='fp'?-8*focus:0)*Math.PI/180,dz:(['a','b','c','guide','cam'].includes(g)?focus/75:0)+(['a','b','c'].includes(g)?travel(g,angle):0),pivot:['zp','fp'].includes(g)?36:0}}
function place(p){gl.uniform1f(u.rot,p.rot);gl.uniform1f(u.dz,p.dz);gl.uniform1f(u.pivot,p.pivot)}
function shown(m){return !(viewMode==='optics'&&!m.glass||viewMode==='mechanical'&&m.glass||m.overlay&&!$('drive').checked)}
function edgeOpacity(m){
 if(m.glass)return .82;
 if(viewMode==='mechanical')return .72;
 if(['cam','guide','focus'].includes(m.group))return .72;
 if(/Cell$/.test(m.name))return .54;
 return .32;
}
function edgeMesh(m,data){
 m.hard=new Float32Array(data.hard);m.smooth=new Float32Array(data.silhouettes);m.lineArray=new Float32Array(m.hard.length+m.smooth.length/2);m.lineArray.set(m.hard);
 m.lineVao=gl.createVertexArray();gl.bindVertexArray(m.lineVao);m.lineBuffer=gl.createBuffer();gl.bindBuffer(gl.ARRAY_BUFFER,m.lineBuffer);gl.bufferData(gl.ARRAY_BUFFER,m.lineArray.byteLength,gl.DYNAMIC_DRAW);
 const pos=gl.getAttribLocation(program,'position'),normal=gl.getAttribLocation(program,'normal');gl.enableVertexAttribArray(pos);gl.vertexAttribPointer(pos,3,gl.FLOAT,false,12,0);gl.disableVertexAttribArray(normal);gl.vertexAttrib3f(normal,0,0,1);
}
function updateLines(m,rot){
 const key=[yaw,pitch,rot].join('/');if(key===m.lineKey)return;m.lineKey=key;
 const nx=-Math.cos(pitch)*Math.sin(yaw),ny=Math.sin(pitch),nz=Math.cos(pitch)*Math.cos(yaw),co=Math.cos(rot),si=Math.sin(rot),vx=co*nx+si*ny,vy=-si*nx+co*ny;
 let n=m.hard.length;const a=m.smooth;for(let i=0;i<a.length;i+=12){const d1=a[i+6]*vx+a[i+7]*vy+a[i+8]*nz,d2=a[i+9]*vx+a[i+10]*vy+a[i+11]*nz;if(d1*d2<0){for(let j=0;j<6;j++)m.lineArray[n++]=a[i+j]}}
 m.lineCount=n/3;gl.bindBuffer(gl.ARRAY_BUFFER,m.lineBuffer);gl.bufferSubData(gl.ARRAY_BUFFER,0,m.lineArray.subarray(0,n));
}
function draw(){
 if(!law)return;const w=c.clientWidth,h=c.clientHeight,d=Math.min(devicePixelRatio,1.5);c.width=Math.round(w*d);c.height=Math.round(h*d);gl.viewport(0,0,c.width,c.height);gl.clearColor(.94,.96,.97,1);gl.clear(gl.COLOR_BUFFER_BIT|gl.DEPTH_BUFFER_BIT);gl.enable(gl.DEPTH_TEST);gl.depthFunc(gl.LEQUAL);gl.enable(gl.BLEND);gl.blendFuncSeparate(gl.SRC_ALPHA,gl.ONE_MINUS_SRC_ALPHA,gl.ONE,gl.ONE_MINUS_SRC_ALPHA);gl.uniform1f(u.yaw,yaw);gl.uniform1f(u.pitch,pitch);gl.uniform1f(u.scale,scale*Math.max(1,1.35/(w/h)));gl.uniform1f(u.aspect,w/h);gl.uniform1f(u.lineMode,0);
 gl.enable(gl.POLYGON_OFFSET_FILL);gl.polygonOffset(1,1);
 for(const m of meshes){if(!shown(m))continue;const alpha=opacityFor(m);gl.uniform1f(u.alpha,alpha);gl.depthMask(alpha>=.85);gl.uniform3fv(u.color,m.glass?cols[m.group]:m.color);place(pose(m));gl.bindVertexArray(m.vao);gl.drawArrays(gl.TRIANGLES,0,m.count)}
 gl.disable(gl.POLYGON_OFFSET_FILL);gl.depthMask(false);
 if($('edges').checked){gl.uniform1f(u.lineMode,1);gl.lineWidth(1);for(const m of meshes){if(!shown(m)||!m.lineVao)continue;const p=pose(m);updateLines(m,p.rot);place(p);gl.uniform3fv(u.color,m.glass?[.12,.32,.48]:m.group==='cam'?[.30,.27,.13]:[.17,.23,.28]);gl.uniform1f(u.alpha,edgeOpacity(m));gl.bindVertexArray(m.lineVao);gl.drawArrays(gl.LINES,0,m.lineCount)}}
 gl.depthMask(true);gl.uniform1f(u.lineMode,0);
 $('angle').value=angle;$('focus').value=focus;$('readout').textContent=`${angle.toFixed(1)}° · G1 ${travel('a',angle).toFixed(2)} / G2 ${travel('b',angle).toFixed(2)} / G3 ${travel('c',angle).toFixed(2)} mm`;$('focus-readout').textContent=`${focus.toFixed(1)}° · ${focus>=0?'+':''}${(focus/75).toFixed(3)} mm`;
}
function labels(){for(const e of document.querySelectorAll('[data-en]')){e.dataset.zh??=e.textContent;e.textContent=en?e.dataset.en:e.dataset.zh}$('lang').textContent=en?'中文':'English';$('play').textContent=playing?(en?'Pause zoom':'暂停变焦'):(en?'Play zoom':'播放变焦');$('focus-play').textContent=focusPlaying?(en?'Pause focus':'暂停调焦'):(en?'Play focus':'播放调焦');document.documentElement.lang=en?'en':'zh-CN'}
for(const mode of ['all','mechanical','optics'])$('mode-'+mode).onclick=()=>{viewMode=mode;for(const k of ['all','mechanical','optics']){$('mode-'+k).classList.toggle('active',k===mode);$('mode-'+k).setAttribute('aria-pressed',String(k===mode))}draw()};
$('lang').onclick=()=>{en=!en;labels()};$('play').onclick=()=>{playing=!playing;focusPlaying=false;labels()};$('focus-play').onclick=()=>{focusPlaying=!focusPlaying;playing=false;labels()};$('angle').oninput=()=>{playing=false;angle=+$('angle').value;labels();draw()};$('focus').oninput=()=>{focusPlaying=false;focus=+$('focus').value;labels();draw()};for(const id of ['opacity','lens-opacity'])$(id).oninput=draw;$('drive').onchange=draw;$('edges').onchange=draw;$('reset').onclick=()=>{yaw=.32;pitch=-.34;scale=85;draw()};for(const a of [0,45,90])$('pose'+a).onclick=()=>{angle=a;playing=false;focus=0;focusPlaying=false;labels();draw()};$('focus-zero').onclick=()=>{focus=0;focusPlaying=false;labels();draw()};let drag=null;c.onpointerdown=e=>{drag=[e.clientX,e.clientY];c.setPointerCapture(e.pointerId)};c.onpointermove=e=>{if(!drag)return;yaw+=(e.clientX-drag[0])*.006;pitch=Math.max(-1.3,Math.min(1.3,pitch+(e.clientY-drag[1])*.006));drag=[e.clientX,e.clientY];draw()};c.onpointerup=()=>drag=null;c.addEventListener('wheel',e=>{e.preventDefault();scale=Math.max(50,Math.min(250,scale*Math.exp(e.deltaY*.001)));draw()},{passive:false});new ResizeObserver(draw).observe(c);document.addEventListener('visibilitychange',()=>{if(document.hidden){playing=false;focusPlaying=false;labels()}});
if(new URLSearchParams(location.search).has('embed'))new ResizeObserver(()=>parent.postMessage({type:'focus-height',height:document.querySelector('main').getBoundingClientRect().height+45},location.origin)).observe(document.querySelector('main'));
Promise.all(['meshes','edges'].map(name=>fetch('projection-assets/'+name+'.json').then(r=>{if(!r.ok)throw Error(r.status);return r.json()}))).then(([data,edges])=>{law=data.motion;meshes=data.meshes.map(m=>{const result=mesh(m);if(edges[m.name])edgeMesh(result,edges[m.name]);return result}).sort((a,b)=>Number(!a.glass)-Number(!b.glass));for(const g of ['a','b','c'])for(const o of [-2.1,2.1])meshes.push(tracks(g,o));$('error').removeAttribute('data-en');$('error').textContent='';labels();draw();requestAnimationFrame(function frame(t){const dt=Math.min(.05,(t-last)/1000);last=t;if(playing){angle+=dir*dt*10;if(angle>=90){angle=90;dir=-1}if(angle<=0){angle=0;dir=1}}if(focusPlaying){focus+=focusDir*dt*22;if(focus>=45){focus=45;focusDir=-1}if(focus<=-45){focus=-45;focusDir=1}}if(playing||focusPlaying)draw();requestAnimationFrame(frame)})}).catch(e=>{$('error').textContent=String(e)});labels();
})();
