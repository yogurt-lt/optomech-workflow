/* Actual native CAD section frames; no external libraries or background autoplay. */
(() => {
  const $=id=>document.getElementById(id),range=$('focus-angle'),play=$('focus-play'),frame=$('focus-frame');
  let mode="3d";const query=new URLSearchParams(location.search);if(query.has("lang"))window.i18n.setLanguage(query.get("lang"));
  let motion=null,index=0,direction=1,timer=null,loading=0;
  const cache=new Map();
  const copy={
    'focus-3d':['三维剖面','3D cutaway'], 'focus-2d':['二维截面','2D section'],
    'focus-back':['← 案例工作台','← Demo workbench'],
    'focus-kicker':['对焦机构实验 · 双高斯 50 mm F1.8','Focus mechanism experiment · Double Gauss 50 mm F1.8'],
    'focus-title':['转动调焦环，看整组镜片怎样移动','Turn the focus ring and watch the optics move'],
    'focus-intro':['青色外壳固定，金色调焦环旋转并前移，蓝色内筒带着镜片沿光轴移动。可切换三维半剖与二维截面，均来自不同姿态下的原生 CAD 实体。','The teal housing stays fixed. The gold focus ring rotates and translates, while the blue optical barrel moves the lens elements along the axis. Switch between 3D cutaway and 2D section views, both derived from native CAD solids at each pose.'],
    'focus-reference':['参考位置','Reference'], 'focus-middle':['行程中点','Mid travel'], 'focus-near':['近端位置','Near end'],
    'focus-gold':['调焦环：旋转＋平移','Focus ring: rotation + translation'],
    'focus-blue':['光学内筒：只平移','Optical barrel: translation only'],
    'focus-teal':['固定外壳：不动','Fixed housing: stationary'],
    'focus-caption':['6 mm 导程 · 270° 转角 · 4.5 mm 行程；每帧相隔 15°。圆盘指针表示调焦环转角；二维模式的固定虚线为安装基准。三维模式为固定视角的 CAD 半剖动画。','6 mm lead · 270° rotation · 4.5 mm travel; 15° per frame. The dial shows focus-ring rotation. Dashed reference lines appear in 2D; 3D uses a fixed camera and a stationary half-section plane.'],
    'focus-step1-title':['01 / 螺旋副产生行程','01 / The helicoid creates axial travel'],
    'focus-step1':['调焦环的内螺旋牙与固定壳的外螺旋牙啮合，把手动转角转为轴向位移。','The ring’s female helicoid engages the fixed housing’s male helicoid, converting manual rotation into axial travel.'],
    'focus-step2-title':['02 / 止推连接带动内筒','02 / The thrust joint moves the optical barrel'],
    'focus-step2':['调焦环前端凸缘由两片垫片、保持环和内筒凸缘捕获，允许相对旋转，同时向前推、向后拉。','Two washers, a keeper and the optical barrel’s flange capture the ring’s front lip. The joint allows relative rotation while transmitting axial motion in both directions.'],
    'focus-step3-title':['03 / 轴向直槽止转','03 / Straight guides prevent rotation'],
    'focus-step3':['固定壳上的导向件伸入内筒直槽，使镜片保持朝向。直槽不负责把旋转转换为行程。','Guides fixed to the housing engage axial slots in the optical barrel and keep the optics from rotating. The straight slots do not generate axial travel.'],
    'focus-scope':['这是独立的手动对焦机构 demo，不是已完成相机适配的量产镜头。0° 尚未标定为无穷远；反光镜避让、通光、预紧、公差和回差仍需校核。动画表现名义几何运动，不是光学成像或应力仿真。','This is a separate manual-focus mechanism demo, not a production lens validated for a camera body. The 0° position is not calibrated infinity. Mirror clearance, clear aperture, preload, tolerances and backlash remain to be checked. The animation shows nominal geometry, not optical imaging or stress simulation.'],
    'focus-gif':['下载 GIF 动图','Download GIF'], 'focus-section-link':['查看结构截面','View structural section'], 'focus-detail-link':['查看调焦连接放大图','View focus-joint detail'],
    'focus-footer':['保持内部光学间距，移动整套光学组件；网页动画不操作本地 CAD 文件。','The complete optical unit moves with its internal spacing preserved. This animation does not control local CAD files.']
  };
  const en=()=>window.i18n?.language==='en';
  function labels(){for(const [id,values] of Object.entries(copy))$(id).textContent=values[en()?1:0];play.textContent=timer?(en()?'Pause':'暂停'):(en()?'Play':'播放');range.setAttribute('aria-label',en()?'Focus rotation':'调焦转角');}
  function image(i){if(!cache.has(i)){const im=new Image();im.src='./focus-assets/'+(mode==='3d'?motion.frames_3d:motion.frames)[i];cache.set(i,im)}return cache.get(i)}
  async function show(i){
    if(!motion)return;
    index=Math.max(0,Math.min(motion.frames.length-1,i));range.value=index;
    const im=image(index),ticket=++loading;
    try{await im.decode();if(ticket!==loading)return;frame.src=im.src;$('focus-error').textContent=''}
    catch{if(ticket===loading)$('focus-error').textContent=en()?'Could not load this frame. Try moving the slider again.':'这一帧未能加载，请重新拖动进度。';return;}
    $('focus-readout').textContent=`${motion.angles[index]}° · ${Number(motion.travel_mm[index]).toFixed(2)} mm`;
    for(const button of document.querySelectorAll('[data-focus-pose]'))button.setAttribute('aria-pressed',String(+button.dataset.focusPose===index));
    image(Math.min(index+1,motion.frames.length-1));image(Math.max(index-1,0));
  }
  function pause(){if(timer)clearTimeout(timer);timer=null;labels();play.setAttribute('aria-pressed','false')}
  async function tick(){if(!timer)return;let next=index+direction;if(next>=motion.frames.length){direction=-1;next=index-1}if(next<0){direction=1;next=1}await show(next);if(timer)timer=setTimeout(tick,(index===0||index===motion.frames.length-1)?600:130)}
  for(const dimension of ['3d','2d'])$('focus-'+dimension).onclick=()=>{if(!motion)return;pause();mode=dimension;cache.clear();for(const k of ['3d','2d'])$('focus-'+k).setAttribute('aria-pressed',String(k===mode));$('focus-gif').href='./focus-assets/'+(mode==='3d'?motion.gif_3d:motion.gif);show(index)};
  play.onclick=()=>{if(timer){pause();return;}timer=setTimeout(tick,130);labels();play.setAttribute('aria-pressed','true')};
  range.oninput=()=>{pause();show(+range.value)};
  for(const button of document.querySelectorAll('[data-focus-pose]'))button.onclick=()=>{pause();show(+button.dataset.focusPose)};
  document.addEventListener('visibilitychange',()=>{if(document.hidden)pause()});
  document.addEventListener('languagechange',labels);labels();
  if(query.has('embed'))new ResizeObserver(()=>parent.postMessage({type:'focus-height',height:document.querySelector('.card').getBoundingClientRect().height+12},location.origin)).observe(document.querySelector('.card'));
  range.disabled=true;
  for(const button of document.querySelectorAll('[data-focus-pose]'))button.disabled=true;
  fetch('./focus-assets/motion.json').then(r=>{if(!r.ok)throw Error(r.status);return r.json()}).then(data=>{motion=data;$('focus-gif').href='./focus-assets/'+data.gif_3d;range.max=data.frames.length-1;range.disabled=false;play.disabled=false;for(const button of document.querySelectorAll('[data-focus-pose]'))button.disabled=false;show(0)}).catch(()=>{$('focus-error').textContent=en()?'Animation data is unavailable. You can download the GIF below.':'动画数据暂时无法加载，可下载下方 GIF 查看。'});
})();
