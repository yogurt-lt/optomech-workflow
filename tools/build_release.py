"""Package the approved six-case project, standalone Skill and static gallery."""
from pathlib import Path
import argparse,hashlib,json,shutil,sys,zipfile
ROOT=Path(__file__).resolve().parents[1]
sys.path.insert(0,str(ROOT/'tools'))
from build_pages import build as build_pages
VERSION='0.7.0'

def archive(path,files):
    with zipfile.ZipFile(path,'w',zipfile.ZIP_DEFLATED,compresslevel=6) as z:
        for source,name in files:z.write(source,name)

def build(output):
    output=output.resolve();output.mkdir(parents=True,exist_ok=True)
    skill=ROOT/'skills/optomech-design'
    for folder in ['optomech','templates']:
        for src in (ROOT/folder).glob('*'):
            if src.is_file() and src.suffix in ('.py','.in'):
                dest=skill/'assets/engine'/folder/src.name;dest.parent.mkdir(parents=True,exist_ok=True);shutil.copy2(src,dest)
    downloads=ROOT/'downloads';downloads.mkdir(exist_ok=True)
    skillzip=downloads/'optomech-design-skill.zip'
    archive(skillzip,[(p,Path('optomech-design')/p.relative_to(skill)) for p in skill.rglob('*') if p.is_file() and '__pycache__' not in p.parts and p.suffix!='.pyc'])
    shutil.copy2(skillzip,output/f'optomech-design-skill-v{VERSION}.zip')
    site=build_pages(ROOT/'dist')
    archive(output/f'optomech-gallery-v{VERSION}.zip',[(p,Path('optomech-workflow')/p.relative_to(site)) for p in site.rglob('*') if p.is_file()])
    catalog=json.loads((ROOT/'examples/catalog.json').read_text(encoding='utf-8'))
    files=[]
    for key in catalog:
        folder=ROOT/'examples'/key
        for name in ['FreeCAD.FCMacro','FreeCAD_Model.FCStd','plan.json','FREECAD.md','FREECAD.en.md','freecad_report.json','section.png','section.en.svg']:
            src=folder/name;assert src.is_file();files.append((src,Path('FreeCAD-demos')/key/name))
    archive(output/f'FreeCAD-demos-v{VERSION}.zip',files)
    excluded={'.git','__pycache__','node_modules','dist','exports','generated','release'}
    files=[]
    for p in ROOT.rglob('*'):
        rel=p.relative_to(ROOT)
        if not p.is_file() or any(part in excluded for part in rel.parts) or p.suffix.lower() in {'.pyc','.swp','.sldprt','.sldasm','.zda','.fcstd1'}:continue
        if p.resolve().is_relative_to(output):continue
        files.append((p,Path('optomech-workflow')/rel))
    archive(output/f'optomech-workflow-v{VERSION}.zip',files)
    manifest=dict(version=VERSION,date='2026-10-22',cases=6,cad_review='All six cases user-confirmed in SolidWorks 2023 and FreeCAD; native FreeCAD checks used 1.1.4',publication='Local release files only; no GitHub push or remote deployment performed',files=[])
    for p in sorted(output.glob('*.zip')):
        manifest['files'].append(dict(name=p.name,bytes=p.stat().st_size,sha256=hashlib.sha256(p.read_bytes()).hexdigest()))
    (output/'release_manifest.json').write_text(json.dumps(manifest,indent=2),encoding='utf-8')
    print(json.dumps(manifest,indent=2))

if __name__=='__main__':
    parser=argparse.ArgumentParser(description=__doc__);parser.add_argument('--output',type=Path,default=ROOT/'release');args=parser.parse_args();build(args.output)
