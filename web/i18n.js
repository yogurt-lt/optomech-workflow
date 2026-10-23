/* Local translations. No external translation service or network dependency. */
(() => {
  const pairs={
    "下载 Skill":"Download Skill",
    "下载设计 Skill":"Download design Skill",
    "开源 AI 光机设计技能 · SolidWorks / FreeCAD":"Open-source AI optomechanical design Skill · SolidWorks / FreeCAD",
    "下载 Skill，把光学设计变成可编辑机械初稿":"Download the Skill. Turn optical layouts into editable mechanical concepts.",
    "先核对光学文件，再推理装配顺序、定位与固定方式，最后生成 CAD 宏和截面图。下方案例展示这套方法的实际结果。":"Check optical inputs, reason through assembly, location and retention, then generate CAD macros and section drawings. Explore the resulting designs below.",
    "下载 optomech-design Skill":"Download optomech-design Skill",
    "安装与使用说明":"Setup & usage",
    "三维剖面调焦演示":"3D cutaway focus animation",
    "本页展示七个结构案例，含整组手动调焦。各案例的检查范围见说明。":"Seven structural demos, including whole-unit manual focus. See each case for its verification scope.",
    "开源 Skill · 七个结构案例 · SolidWorks / FreeCAD":"Open-source Skill · seven structural demos · SolidWorks / FreeCAD",

    '手动调焦动画':'Manual focus animation',
    '光学模型 → 装配推理 → 截面检查 → 可编辑 CAD 模型':'Optical model → assembly reasoning → section review → editable CAD models',
    '本地运行 · v0.7.0':'Local · v0.7.0',
    '下载 FreeCAD 宏':'Download FreeCAD macro',
    '下载可编辑 FCStd 模型':'Download editable FCStd model',
    'FreeCAD 运行与检查说明':'FreeCAD run & review guide',
    'SolidWorks：导入 .bas，保存为 .swp 后运行 main。FreeCAD：在宏菜单执行 .FCMacro，或直接打开 .FCStd。三维预览省略专用孔槽。':'SolidWorks: import .bas, save as .swp and run main. FreeCAD: execute .FCMacro from the Macro menu or open .FCStd directly. Dedicated cuts are omitted from the web preview.',
    'CAD 软件':'CAD application',
    '宏已下载。在 FreeCAD 宏菜单执行 .FCMacro，并检查生成结果。':'Macro downloaded. Execute the .FCMacro inside FreeCAD and inspect the generated model.',
    '光学模型 → 装配推理 → 截面检查 → 可编辑 SolidWorks 模型':'Optical model → assembly reasoning → section review → editable SolidWorks model',
    '本地运行 · v0.4.0':'Local · v0.5.0','本地运行 · v0.5.0':'Local · v0.5.0',
    '01 / 选择案例':'01 / Choose a demo','光学系统':'Optical system','查看与下载':'View & download',
    '02 / 查看截面图':'02 / View section','下载已验证完整宏':'Download reviewed full macro','仅下载机械宏':'Mechanics-only macro',
    '导入 SolidWorks 2023 VBA 编辑器，保存为 .swp 后运行 main。三维预览用于理解结构，孔槽与螺纹细节以 CAD 为准。':'Import into the SolidWorks 2023 VBA editor, save as .swp, then run main. The preview explains the structure; CAD defines the detailed cuts and interfaces.',
    '装配依据与完整说明 ↗':'Design rationale & instructions ↗',
    '本页展示六个经用户运行检查的初稿案例。通光、公差与预紧需按实际系统继续校核。':'Six concept demos have been run and reviewed by the user. Aperture clearance, tolerances and preload need application-specific verification.',
    '三维装配预览':'3D assembly preview','可拖动旋转的镜头三维装配剖切预览':'Interactive cutaway lens assembly preview','拖动旋转 · 滚轮缩放':'Drag to orbit · scroll to zoom','半剖视图':'Half section','隐藏镜筒':'Hide barrels','复位视角':'Reset view',
    '截面旋转预览 · 点胶孔与螺纹细节以 CAD 为准':'Revolved-profile preview · detailed ports and threads belong in CAD',
    '装配过程与固定方式':'Assembly sequence & retention','如何定位与固定':'Location & retention','压紧力传到哪里':'Preload path',
    '02 / 截面图':'02 / Section drawing','下载原尺寸截面图':'Download full-size section','点击图片可在新标签页放大':'Click the drawing to open it in a new tab',
    '这是已检查的基准截面图；修改后的几何请查看上方三维预览与下方编辑画布。':'This is the reviewed baseline drawing. See the 3D preview and editor for modified geometry.',
    '高级工具：修改截面、检查方案与分析新文件':'Advanced: edit profiles, check plans & analyze files',
    '这里的编辑会同步三维预览；上方案例截面图保留为基准。修改版生成器输出旋转主体，专用孔槽请以已验证案例宏为基础补充。':'Edits update the 3D preview; the section image remains the baseline. The general generator emits revolved bodies. Preserve dedicated cuts using the reviewed case macro.',
    '机械零件':'Mechanical part','半径 / mm':'Radius / mm','应用截面':'Apply profile','恢复案例':'Reset demo','包含镜片参考体':'Include optical reference blanks',
    '检查方案':'Check plan','生成修改版主体宏':'Generate modified bodies','保存方案 JSON':'Save plan JSON','导入已有方案 JSON':'Import plan JSON',
    '截面编辑':'Profile editor','原始方案与边界':'Source plan & scope','检查结果':'Check results','先检查方案，再下载修改版宏。':'Check the plan before downloading a modified macro.',
    '尝试新光学系统':'Try another optical system','支持文件初检与装配候选分析；新系统仍需人工 / AI 编制具体截面方案。':'File inspection and loading candidates are supported. A new system still needs a reviewed mechanical plan.',
    'Zemax（可选）':'Zemax (optional)','分析新系统':'Analyze files','上传文件后在这里查看分析。':'File analysis will appear here.','导出分析 JSON':'Download analysis JSON',
    '原生可编辑旋转特征 · 本地离线预览 · 六个案例已完成用户运行检查':'Native editable revolved features · offline preview · six user-reviewed demos',
    '正在读取案例…':'Loading demo…','方案已载入；修改后请重新检查。':'Plan loaded. Check again after editing.',
    '自定义方案：修改后请重新检查':'Custom plan: check after editing','导入方案':'Imported plan','自定义方案':'Custom plan',
    '请先明确装配顺序与固定方式。':'Specify the assembly sequence and retention method first.','请查看原始方案。':'See the source plan.','尚未提供。':'Not specified.',
    '导入文件不代表已经完成机械和光学校核。':'Importing a plan does not verify its mechanical or optical behavior.',
    '截面已修改；基准光线检查报告不再适用于修改后的模型。':'Profile changed. Baseline ray checks no longer describe the modified model.',
    '正在检查截面，请稍候…':'Checking profiles…','检查发现问题：':'Issues found:',
    '机械 / 玻璃截面采样检查未发现交叠。':'No overlap found in the sampled mechanical/glass sections.',
    '正在检查方案并生成宏，请稍候…':'Checking the plan and generating the macro…',
    '宏已下载。导入 SolidWorks VBA 编辑器，运行 Optomech_Web.main。':'Macro downloaded. Import it into the SolidWorks VBA editor and run Optomech_Web.main.',
    '请选择 STEP 文件。':'Select a STEP file.','STEP 限 40 MB，Zemax 限 2 MB。':'Limits: STEP 40 MB; Zemax 2 MB.',
    '镜片':'Glass','压环':'Retainers','隔圈':'Spacers','内部镜框':'Internal cells','定位套':'Adapter','镜筒':'Barrel','前筒':'Front barrel','中筒':'Middle barrel','后筒':'Rear barrel',
    '前筒 / 光阑':'Front barrel / stop','中筒 / 光阑':'Middle barrel / stop','后筒 / 光阑':'Rear barrel / stop','镜筒 / 光阑':'Barrel / stop',
    'Z  光轴':'Z  optical axis','下载工作流 Skill':'Download workflow Skill','Skill 使用说明':'Skill setup & usage',
    '把文件核对、装配推理、宏生成和反馈修订流程带到自己的 AI 助手中。':'Bring the file checks, assembly reasoning, macro generation and review loop to your own AI assistant.'
  };
  const reverse=Object.fromEntries(Object.entries(pairs).map(([a,b])=>[b,a]));
  let lang='zh';try{lang=localStorage.getItem('optomech-language')==='en'?'en':'zh'}catch{}
  function t(s){return lang==='en'?s.split('\n').map(x=>pairs[x]||x).join('\n'):s}
  function apply(){
    document.documentElement.lang=lang==='en'?'en':'zh-CN';document.title=document.body.dataset.page==='focus'?(lang==='en'?'Optomech Design Skill · Manual focus':'Optomech Design Skill · 手动调焦动画'):(lang==='en'?'Optomech Design Skill · optical-to-mechanical demos':'Optomech Design Skill · 光机结构工作台');
    const walk=document.createTreeWalker(document.body,NodeFilter.SHOW_TEXT);let node;
    while(node=walk.nextNode()){
      if(['SCRIPT','STYLE','PRE','TEXTAREA','OPTION'].includes(node.parentElement?.tagName))continue;
      const s=node.nodeValue.trim(),translated=lang==='en'?pairs[s]:reverse[s];
      if(translated&&translated!==s)node.nodeValue=node.nodeValue.replace(s,translated);
    }
    for(const el of document.querySelectorAll('[aria-label],[alt]'))for(const attr of ['aria-label','alt']){const s=el.getAttribute(attr);if(s){const value=lang==='en'?pairs[s]:reverse[s];if(value&&s!==value)el.setAttribute(attr,value)}}
    const button=document.getElementById('language-toggle');if(button){const label=lang==='en'?'中文':'English';if(button.textContent!==label)button.textContent=label;button.setAttribute('aria-label',lang==='en'?'切换到中文':'Switch to English');button.setAttribute('aria-pressed',String(lang==='en'))}
  }
  function setLanguage(value){lang=value==='en'?'en':'zh';try{localStorage.setItem('optomech-language',lang)}catch{}document.dispatchEvent(new Event('languagechange'));apply()}
  window.i18n={t,apply,setLanguage,get language(){return lang}};
  document.getElementById('language-toggle').onclick=()=>setLanguage(lang==='en'?'zh':'en');
  new MutationObserver(()=>apply()).observe(document.body,{childList:true,subtree:true,characterData:true});apply();
})();
