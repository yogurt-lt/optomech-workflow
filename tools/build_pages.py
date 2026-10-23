"""Export the reviewed bilingual gallery as static files for GitHub Pages."""
from pathlib import Path
import argparse,json,re,shutil

ROOT=Path(__file__).resolve().parents[1]

def build(output):
    output=output.resolve()
    if output==ROOT or output in ROOT.parents:raise ValueError('Choose a dedicated output directory.')
    output.mkdir(parents=True,exist_ok=True)
    page=(ROOT/'web/index.html').read_text(encoding='utf-8')
    # The local workbench keeps its editor/backend. The public gallery has static data.
    page=re.sub(r'<details class="card advanced-panel">.*?</details>','',page,flags=re.S)
    page=re.sub(r'<script>.*?</script>','<script src="./gallery.js"></script>',page,flags=re.S)
    page=page.replace('src="/i18n.js"','src="./i18n.js"').replace('src="/view3d.js"','src="./view3d.js"')
    page=page.replace('href="/downloads/','href="./downloads/').replace('href="/skill-guide"','href="./docs/SKILL_USAGE.md"')
    page=page.replace('href="/focus.html"','href="./focus.html"')
    page=page.replace('本地运行 · v0.7.0','开源展示 · v0.7.0')
    page=page.replace('原生可编辑旋转特征 · 本地离线预览 · 六个案例已完成用户运行检查','六个已检查案例 · 中英双语 · 开源代码与 Skill')
    page=page.replace('</header>','<a class="repo-link" href="https://github.com/yogurt-lt/optomech-workflow" target="_blank" rel="noopener">GitHub ↗</a></header>')
    page=page.replace('</style>', '.repo-link{color:#c9e8ee;text-decoration:none;font-size:13px}</style>')
    (output/'index.html').write_text(page,encoding='utf-8')
    for file in ['view3d.js','gallery.js']:shutil.copy2(ROOT/'web'/file,output/file)
    for file in ['focus.html','focus-motion.js']:shutil.copy2(ROOT/'web'/file,output/file)
    shutil.copytree(ROOT/'web/focus-assets',output/'focus-assets',dirs_exist_ok=True)
    i18n=(ROOT/'web/i18n.js').read_text(encoding='utf-8').replace('const pairs={',"const pairs={\n    '开源展示 · v0.7.0':'Showcase · v0.7.0',\n    '六个已检查案例 · 中英双语 · 开源代码与 Skill':'Six reviewed demos · Chinese / English · open-source code & Skill',")
    (output/'i18n.js').write_text(i18n,encoding='utf-8')
    catalog=json.loads((ROOT/'examples/catalog.json').read_text(encoding='utf-8'));english=json.loads((ROOT/'examples/catalog.en.json').read_text(encoding='utf-8'))
    for key,item in catalog.items():
        item['en']=english[key];dest=output/'examples'/key;dest.mkdir(parents=True,exist_ok=True)
        files={'plan.json','section.png','section.en.svg','DESIGN.md','DESIGN.en.md',item['full_macro'],item['mechanics_macro'],item['freecad_macro'],item['freecad_model'],'FREECAD.md','FREECAD.en.md','freecad_report.json'}
        for file in files:shutil.copy2(ROOT/'examples'/key/file,dest/file)
    (output/'data').mkdir(exist_ok=True);(output/'data/catalog.json').write_text(json.dumps(catalog,ensure_ascii=False,separators=(',',':')),encoding='utf-8')
    (output/'downloads').mkdir(exist_ok=True);shutil.copy2(ROOT/'downloads/optomech-design-skill.zip',output/'downloads/optomech-design-skill.zip')
    (output/'docs').mkdir(exist_ok=True)
    for file in ['SKILL_USAGE.md','SKILL_USAGE.en.md']:shutil.copy2(ROOT/'docs'/file,output/'docs'/file)
    (output/'.nojekyll').touch()
    print('Static GitHub Pages gallery built:',output)
    return output

if __name__=='__main__':
    p=argparse.ArgumentParser(description=__doc__);p.add_argument('--output',type=Path,default=ROOT/'dist');args=p.parse_args();build(args.output)
