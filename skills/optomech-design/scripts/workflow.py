"""Standalone numerical checks and baseline macro generation; no CAD automation."""
from pathlib import Path
import argparse,json,sys,math
SKILL=Path(__file__).resolve().parents[1];sys.path.insert(0,str(SKILL/'assets/engine'))
from optomech.__main__ import analyze
from optomech.geometry import validate_plan
from optomech.generator import generate_macro
from optomech.freecad_generator import generate_freecad_macro,capability_report

def distance(p,a,b):
    dx,dy=b[0]-a[0],b[1]-a[1];den=dx*dx+dy*dy
    t=max(0,min(1,((p[0]-a[0])*dx+(p[1]-a[1])*dy)/den)) if den else 0
    return math.hypot(p[0]-a[0]-t*dx,p[1]-a[1]-t*dy)

def check(plan):
    result=validate_plan(plan);result['wall_distances']=[]
    for part in plan.get('parts',[]):
        rgb=part.get('color_rgb')
        if rgb is not None and (not isinstance(rgb,list) or len(rgb)!=3 or any(isinstance(v,bool) or not isinstance(v,(int,float)) or not math.isfinite(v) or not 0<=v<=1 for v in rgb)):
            result['errors'].append(part.get('name','Part')+': color_rgb must contain three finite numbers in [0,1].')
        n=part.get('outer_chain_vertex_count');poly=part.get('profile_zr_mm',[])
        if not isinstance(n,int) or not 2<=n<=len(poly)-2:continue
        outer,inner=poly[:n],poly[n:]
        w=min(min(distance(a,x,y),distance(b,x,y),distance(x,a,b),distance(y,a,b)) for a,b in zip(outer,outer[1:]) for x,y in zip(inner,inner[1:]))
        result['wall_distances'].append(dict(part=part['name'],minimum_inner_outer_chain_distance_mm=w))
    result['scope']='Inner/outer meridional chain distance includes step roots. Check ring axial thickness and local holes/slots separately. Insertion, stress, tolerance and optical aperture checks are separate.'
    return result

def main():
    parser=argparse.ArgumentParser(description=__doc__);subs=parser.add_subparsers(dest='command',required=True)
    a=subs.add_parser('analyze');a.add_argument('--step',type=Path,required=True);a.add_argument('--zmx',type=Path)
    a=subs.add_parser('check');a.add_argument('plan',type=Path);a.add_argument('--report',type=Path)
    a=subs.add_parser('macro');a.add_argument('plan',type=Path);a.add_argument('--output',type=Path,required=True);a.add_argument('--include-glass',action='store_true');a.add_argument('--cad',choices=['solidworks','freecad'],default='solidworks');a.add_argument('--allow-base-only',action='store_true');a.add_argument('--report',type=Path)
    args=parser.parse_args()
    if args.command=='analyze':
        result=analyze(args.step.read_text(encoding='utf-8'),args.zmx.read_bytes() if args.zmx else None)
    else:
        plan=json.loads(args.plan.read_text(encoding='utf-8-sig'));result=check(plan)
        if args.command=='macro':
            if result['errors']:print(json.dumps(result,indent=2));return 1
            if args.cad=='freecad':
                result['capabilities']=capability_report(plan,args.include_glass)
                if result['capabilities']['extra_features_not_implemented'] and not args.allow_base_only:
                    result['errors'].append('Some declared cuts lack supported data. Review capabilities; use --allow-base-only only if the listed omissions are acceptable.')
                    print(json.dumps(result,indent=2));return 1
                text,capabilities=generate_freecad_macro(plan,args.include_glass,args.allow_base_only)
                args.output.parent.mkdir(parents=True,exist_ok=True);args.output.write_text(text,encoding='utf-8')
                result['macro']=str(args.output)
                report=args.report or args.output.with_suffix('.capabilities.json')
                report.parent.mkdir(parents=True,exist_ok=True);report.write_text(json.dumps(result,indent=2),encoding='utf-8')
                print(json.dumps(result,indent=2));return 0
            text=generate_macro(plan,'Optomech_Concept',include_blanks=args.include_glass)
            palette=[]
            for p in plan['parts']:
                if 'color_rgb' not in p:continue
                r,g,b=p['color_rgb'];palette.append(f'    If InStr(key, "{p["name"].lower()}") > 0 Then\n        values(0) = {r}: values(1) = {g}: values(2) = {b}\n        ConceptPalette = values\n        Exit Function\n    End If')
            text=text.replace('    If InStr(key, "lens") > 0 Then','\n'.join(palette)+'\n    If InStr(key, "lens") > 0 Then')
            if plan.get('radial_features') or any(p.get('cap_slot') or p.get('radial_windows') for p in plan['parts']):result['warnings'].append('Extra radial_features/cap_slot/radial_windows metadata requires dedicated VBA cuts; the general SolidWorks baseline generator does not implement it. Use the reviewed case macro for complete details.')
            args.output.parent.mkdir(parents=True,exist_ok=True);args.output.write_bytes(text.replace('\r\n','\n').replace('\n','\r\n').encode('ascii'));result['macro']=str(args.output)
        elif args.report:
            args.report.parent.mkdir(parents=True,exist_ok=True);args.report.write_text(json.dumps(result,indent=2),encoding='utf-8')
    print(json.dumps(result,ensure_ascii=True,indent=2));return bool(result.get('errors'))

if __name__=='__main__':sys.exit(main())
