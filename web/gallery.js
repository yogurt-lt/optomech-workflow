/* Static gallery entry point: works under any GitHub Pages project subpath. */
(() => {
  const $=id=>document.getElementById(id);
  let demos={},plan=null,key=null;
  const english=()=>window.i18n.language==='en';
  function meta(){const p=demos[key];return p?(english()?{...p,...p.en,status:'User-run and reviewed in SolidWorks 2023',section:'section.en.svg',design:'DESIGN.en.md'}:p):null}
  async function json(path){const r=await fetch(path);if(!r.ok)throw new Error(`${r.status}: ${path}`);return r.json()}
  function render(){
    if(!plan)return;const m=meta();
    $('status').textContent=m.status;$('caseSummary').textContent=m.summary;$('banner').textContent=m.title+' · '+m.status;
    $('caseTitle').textContent=m.title;$('methodTag').textContent=m.summary;$('storyIntro').textContent=m.intro;
    $('fixing').textContent=m.fixing;$('loadPath').textContent=m.load_path;$('storyNote').textContent=m.note;
    $('stageList').replaceChildren(...m.stages.map(([title,body],i)=>{const card=document.createElement('div'),h=document.createElement('h3'),p=document.createElement('p');card.className='stage';h.textContent=`${String(i+1).padStart(2,'0')} / ${title}`;p.textContent=body;card.append(h,p);return card}));
    for(const [id,file] of [['fullMacro',m.full_macro],['mechMacro',m.mechanics_macro],['designLink',m.design]])$(id).href=`examples/${key}/${file}`;
    const image=`examples/${key}/${m.section}`;$('sectionImage').src=image;$('sectionImage').alt=m.title+(english()?' reviewed section drawing':'已检查结构截面图');$('imageLink').href=image;$('imageDownload').href=image;
    $('imageStats').textContent=english()?`${plan.lenses?.length||0} glass solids / ${plan.parts.length} mechanical parts`:`${plan.lenses?.length||0} 个玻璃实体 / ${plan.parts.length} 个机械件`;
    $('skillGuide').href=english()?'docs/SKILL_USAGE.en.md':'docs/SKILL_USAGE.md';
    window.optomech3d.update(plan);window.i18n.apply();
  }
  function options(){const value=key;$('demo').replaceChildren(...Object.entries(demos).map(([k,m])=>new Option(english()?m.en.title:m.title,k)));if(value)$('demo').value=value}
  async function load(value){
    $('demo').disabled=true;$('language-toggle').disabled=true;
    try{const data=await json(`examples/${value}/plan.json`);key=value;plan=data;render()}
    catch(e){$('banner').textContent=(english()?'Could not load this demo: ':'案例加载失败：')+e.message}
    finally{$('demo').disabled=false;$('language-toggle').disabled=false}
  }
  $('demo').onchange=()=>load($('demo').value);
  document.addEventListener('languagechange',()=>{options();render()});
  (async()=>{try{demos=await json('data/catalog.json');options();await load('fixed-projection')}catch(e){$('banner').textContent=e.message}})();
})();
