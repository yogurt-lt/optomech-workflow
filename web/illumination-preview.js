(() => {
  const $ = id => document.getElementById(id);
  const params = new URLSearchParams(location.search);
  let en = params.get('lang') === 'en';
  let angle = 0, explode = false, playing = false, detail = false, dir = 1;
  let yaw = -.32, pitch = .72, zoom = 1;
  const canvas = $('scene'), gl = canvas.getContext('webgl2', {antialias:true});
  const words = (zh, eng) => en ? eng : zh;
  function labels() {
    document.documentElement.lang = en ? 'en' : 'zh-CN';
    document.querySelectorAll('[data-en]').forEach(el => {
      el.dataset.zh ??= el.textContent;
      el.textContent = words(el.dataset.zh, el.dataset.en);
    });
    $('play').textContent = playing ? words('暂停调焦','Pause focus') : words('播放调焦','Play focus');
    $('language').textContent = words('English','中文');
  }
  $('language').onclick = () => { en = !en; labels(); };
  if (params.has('embed')) {
    document.body.classList.add('embed');
    const main = document.querySelector('main');
    new ResizeObserver(() => parent.postMessage({type:'focus-height',height:main.offsetHeight+8}, location.origin)).observe(main);
  }
  labels();
  const fail = () => {
    $('loading').removeAttribute('data-en');
    $('loading').textContent = words('三维预览无法载入。可查看上方整机图及下方爆炸图。','3D preview unavailable. The assembly image and exploded view remain available.');
  };
  if (!gl) { fail(); return; }
  function shader(type, source) {
    const handle = gl.createShader(type);
    gl.shaderSource(handle, source);
    gl.compileShader(handle);
    if (!gl.getShaderParameter(handle, gl.COMPILE_STATUS)) throw new Error(gl.getShaderInfoLog(handle));
    return handle;
  }
  const program = gl.createProgram();
  gl.attachShader(program, shader(gl.VERTEX_SHADER, `#version 300 es
    in vec3 pos;
    in vec3 normal;
    uniform float turn,yaw,pitch,scale,aspect,mid,lift;
    uniform vec3 offset;
    uniform vec2 axis;
    out float shade;
    vec3 xy(vec3 p,float a) {
      float c=cos(a),s=sin(a);
      return vec3(c*p.x-s*p.y,s*p.x+c*p.y,p.z);
    }
    vec3 view(vec3 p) {
      p=p.zyx;
      float c=cos(yaw),s=sin(yaw);
      p=vec3(c*p.x+s*p.z,p.y,-s*p.x+c*p.z);
      c=cos(pitch);s=sin(pitch);
      return vec3(p.x,c*p.y-s*p.z,s*p.y+c*p.z);
    }
    void main() {
      vec3 pivot=vec3(axis,0.);
      vec3 p=xy(pos-pivot,turn)+pivot+offset;
      vec3 v=view(p-vec3(0.,0.,mid));
      gl_Position=vec4(v.x/scale/aspect,(v.y-lift)/scale,-v.z/350.,1.);
      shade=.63+.37*abs(dot(normalize(view(xy(normal,turn))),normalize(vec3(-.25,.65,.72))));
    }`));
  gl.attachShader(program, shader(gl.FRAGMENT_SHADER, `#version 300 es
    precision highp float;
    uniform vec3 color;
    uniform float opacity;
    in float shade;
    out vec4 pixel;
    void main() {pixel=vec4(color*shade,opacity);}
  `));
  gl.linkProgram(program);
  if (!gl.getProgramParameter(program,gl.LINK_STATUS)) {fail();return;}
  gl.useProgram(program);
  const u=Object.fromEntries(['turn','yaw','pitch','scale','aspect','mid','lift','axis','offset','color','opacity'].map(k=>[k,gl.getUniformLocation(program,k)]));
  let meshes=[];
  function upload(m) {
    const points=new Float32Array(m.triangles.length*18);
    let at=0;
    for (const tri of m.triangles) {
      const [a,b,c]=tri.map(i=>m.vertices[i]), x=b.map((v,i)=>v-a[i]),y=c.map((v,i)=>v-a[i]);
      const n=[x[1]*y[2]-x[2]*y[1],x[2]*y[0]-x[0]*y[2],x[0]*y[1]-x[1]*y[0]];
      const length=Math.hypot(...n)||1;
      for(const p of [a,b,c]) {points.set(p,at);points.set(n.map(v=>v/length),at+3);at+=6;}
    }
    const vao=gl.createVertexArray();
    gl.bindVertexArray(vao);
    gl.bindBuffer(gl.ARRAY_BUFFER,gl.createBuffer());
    gl.bufferData(gl.ARRAY_BUFFER,points,gl.STATIC_DRAW);
    for(const [name,start] of [['pos',0],['normal',12]]) {
      const loc=gl.getAttribLocation(program,name);
      gl.enableVertexAttribArray(loc);
      gl.vertexAttribPointer(loc,3,gl.FLOAT,false,24,start);
    }
    let lineVao=null;
    if(m.edges?.length){
      lineVao=gl.createVertexArray();
      gl.bindVertexArray(lineVao);
      gl.bindBuffer(gl.ARRAY_BUFFER,gl.createBuffer());
      gl.bufferData(gl.ARRAY_BUFFER,new Float32Array(m.edges),gl.STATIC_DRAW);
      gl.enableVertexAttribArray(gl.getAttribLocation(program,'pos'));
      gl.vertexAttribPointer(gl.getAttribLocation(program,'pos'),3,gl.FLOAT,false,12,0);
      gl.vertexAttrib3f(gl.getAttribLocation(program,'normal'),0,0,1);
    }
    const lineCount=(m.edges?.length||0)/3;
    delete m.edges;
    delete m.vertices;
    delete m.triangles;
    return {...m,vao,count:at/6,lineVao,lineCount};
  }
  function setPose(m) {
    const turn=explode ? 0 : m.pinion ? -4*angle : m.cam ? angle : 0;
    gl.uniform1f(u.turn,turn*Math.PI/180);
    gl.uniform2fv(u.axis,m.pivot);
    gl.uniform3fv(u.offset,explode?m.explode:[0,0,m.move?angle*.02:0]);
  }
  function draw() {
    if (!meshes.length) return;
    const d=Math.min(devicePixelRatio,1.5), w=Math.round(canvas.clientWidth*d), h=Math.round(canvas.clientHeight*d);
    if(!w||!h) return;
    if(w!==canvas.width||h!==canvas.height){canvas.width=w;canvas.height=h;}
    gl.viewport(0,0,w,h);
    gl.clearColor(.982,.988,.992,1);
    gl.clear(gl.COLOR_BUFFER_BIT|gl.DEPTH_BUFFER_BIT);
    gl.enable(gl.DEPTH_TEST);
    gl.enable(gl.BLEND);
    gl.blendFunc(gl.SRC_ALPHA,gl.ONE_MINUS_SRC_ALPHA);
    gl.uniform1f(u.yaw,yaw);
    gl.uniform1f(u.pitch,pitch);
    gl.uniform1f(u.aspect,w/h);
    gl.uniform1f(u.mid,explode?-46:detail?-32:-14);
    gl.uniform1f(u.lift,detail&&!explode?8:0);
    gl.uniform1f(u.scale,(explode?87:detail?25:64)*Math.max(1,1.85*h/w)/zoom);
    const translucent=m=>!explode&&$('transparent').checked&&(m.shell||m.sleeve);
    gl.enable(gl.POLYGON_OFFSET_FILL);
    gl.polygonOffset(1,1);
    for(const transparent of [false,true]) {
      gl.depthMask(!transparent);
      for (const m of meshes) {
        if(translucent(m)!==transparent)continue;
        setPose(m);
        gl.uniform3fv(u.color,m.color);
        gl.uniform1f(u.opacity,transparent?(m.shell ? .22 : .27):1);
        gl.bindVertexArray(m.vao);
        gl.drawArrays(gl.TRIANGLES,0,m.count);
      }
    }
    gl.disable(gl.POLYGON_OFFSET_FILL);
    gl.depthMask(false);
    gl.depthFunc(gl.LEQUAL);
    for(const m of meshes){
      if(!m.lineVao)continue;
      setPose(m);
      gl.uniform3fv(u.color,[.14,.22,.29]);
      gl.uniform1f(u.opacity,translucent(m) ? .63 : .46);
      gl.bindVertexArray(m.lineVao);
      gl.drawArrays(gl.LINES,0,m.lineCount);
    }
    gl.depthMask(true);
    $('readout').textContent=`${angle.toFixed(1)}° · ${(angle*.02).toFixed(2)} mm · ${words("小齿轮","Pinion")} ${(-4*angle).toFixed(1)}°`;
    $('angle').value=angle;
  }
  $('angle').oninput=()=>{playing=false;angle=+$('angle').value;labels();draw();};
  $('play').onclick=()=>{playing=!playing;labels();};
  $('explode').onclick=()=>{
    explode=!explode;playing=false;
    $('driveDetail').disabled=explode;
    $('explode').setAttribute('aria-pressed',String(explode));
    $('angle').disabled=explode;$('play').disabled=explode;
    labels();draw();
  };
  $('transparent').onchange=draw;
  $('driveDetail').onclick=()=>{
    if(explode)return;
    detail=!detail;
    $('driveDetail').setAttribute('aria-pressed',String(detail));
    draw();
  };
  $('reset').onclick=()=>{yaw=-.32;pitch=.72;zoom=1;detail=false;$('driveDetail').setAttribute('aria-pressed','false');draw();};
  let drag=null;
  canvas.onpointerdown=e=>{drag=[e.clientX,e.clientY];canvas.setPointerCapture(e.pointerId);};
  canvas.onpointermove=e=>{if(!drag)return;yaw+=(e.clientX-drag[0])*.008;pitch=Math.max(-1.5,Math.min(1.5,pitch+(e.clientY-drag[1])*.008));drag=[e.clientX,e.clientY];draw();};
  canvas.onpointerup=canvas.onpointercancel=()=>{drag=null;};
  canvas.addEventListener('wheel',e=>{e.preventDefault();zoom=Math.max(.4,Math.min(5,zoom*Math.exp(-e.deltaY*.001)));draw();},{passive:false});
  new ResizeObserver(draw).observe(canvas);
  fetch('illumination-assets/meshes.json').then(r=>{if(!r.ok)throw new Error(r.status);return r.json();}).then(data=>{
    meshes=data.map(upload);$('loading').hidden=true;$('controls').hidden=false;canvas.hidden=false;$('poster').hidden=true;draw();
    let last=0;
    requestAnimationFrame(function frame(now){
      if(playing){
        angle+=dir*Math.min(50,now-last)*.024;
        if(angle>=30){angle=30;dir=-1;}
        if(angle<=-30){angle=-30;dir=1;}
        draw();
      }
      last=now;requestAnimationFrame(frame);
    });
  }).catch(fail);
})();
