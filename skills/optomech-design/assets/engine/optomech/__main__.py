import argparse,json
from pathlib import Path
from .prescription import parse_zmx,step_metadata,optical_groups,assembly_candidates
from .generator import generate_macro
from .geometry import validate_plan

def analyze(step,zmx=None):
    meta=step_metadata(step);result={'step':meta}
    if zmx:
        surfaces=parse_zmx(zmx);groups=optical_groups(surfaces,meta['surface_pairs'])
        result['groups']=[dict(surfaces=g,radius_mm=max(surfaces[n]['mechanical_radius_mm'] for n in g)) for g in groups]
        result['assembly_candidates']=assembly_candidates([g['radius_mm'] for g in result['groups']])
        result['surfaces']=list(surfaces.values())
    result['next_step']='Review coordinates, physical edges, contacts and ray clearance; create an explicit design plan before generating CAD.'
    return result

def main():
    p=argparse.ArgumentParser(description='Reviewable optical-to-mechanical workflow')
    sub=p.add_subparsers(dest='command',required=True)
    a=sub.add_parser('analyze');a.add_argument('--step',type=Path,required=True);a.add_argument('--zmx',type=Path)
    g=sub.add_parser('generate');g.add_argument('plan',type=Path);g.add_argument('-o','--output',type=Path,required=True);g.add_argument('--blank-references',action='store_true')
    v=sub.add_parser('check');v.add_argument('plan',type=Path)
    args=p.parse_args()
    if args.command=='analyze':print(json.dumps(analyze(args.step.read_text(),args.zmx.read_bytes() if args.zmx else None),indent=2,ensure_ascii=False))
    else:
        plan=json.loads(args.plan.read_text(encoding='utf-8'))
        if args.command=='check':
            result=validate_plan(plan);print(json.dumps(result,indent=2,ensure_ascii=False));raise SystemExit(bool(result['errors']))
        args.output.parent.mkdir(parents=True,exist_ok=True);args.output.write_text(generate_macro(plan,include_blanks=args.blank_references),encoding='ascii');print(str(args.output))

if __name__=='__main__':main()
