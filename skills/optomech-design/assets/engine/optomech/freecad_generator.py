"""Emit an editable FreeCAD baseline from the shared mechanical plan, not from VBA."""
from pathlib import Path
import ast,json,math,re
from .geometry import validate_plan
from .generator import blank_references

ROOT=Path(__file__).resolve().parent.parent

def capability_report(plan,include_blanks=False):
    extra=[]
    if plan.get('glue_ports'):extra.append('glue_ports: dedicated adhesive access cuts')
    for feature in plan.get('radial_features',[]):
        if not feature.get('angles_degrees'):extra.append('radial_features: missing explicit angles_degrees')
    for part in plan.get('parts',[]):
        if part.get('cap_slot') and not part.get('wrench_slots'):extra.append(part['name']+': cap_slot wrench cuts')
    glass_count=len(blank_references(plan)) if include_blanks else 0
    return dict(target='freecad',adapter_version='0.3',baseline_only=bool(extra),native_features=['App::Part coordinate container','PartDesign::Body','Sketcher::SketchObject','PartDesign::Revolution','PartDesign::Pocket for explicit radial holes, radial windows and wrench slots'],
                requested_glass=bool(include_blanks),included_glass=bool(glass_count),included_glass_count=glass_count,extra_features_not_implemented=extra,
                assembly='Shared-coordinate bodies; no Assembly workbench joints or mates',
                units='millimeters, source optical axis +Z',runtime_validation='pending in FreeCAD',feature_audit='Only declared metadata is detected; audit legacy per-case macros separately for cuts absent from plan JSON.')

def generate_freecad_macro(plan,include_blanks=False,allow_base_only=False):
    result=validate_plan(plan)
    if result['errors']:raise ValueError('\n'.join(result['errors']))
    offset=plan.get('source_origin_z_mm',0.)
    if isinstance(offset,bool) or not isinstance(offset,(int,float)) or not math.isfinite(offset) or abs(offset)>100000:raise ValueError('Invalid finite source Z offset.')
    report=capability_report(plan,include_blanks)
    if report['extra_features_not_implemented'] and not allow_base_only:
        raise ValueError('FreeCAD baseline does not implement dedicated cuts. Use allow_base_only explicitly after reviewing the capability report: '+', '.join(report['extra_features_not_implemented']))
    payload=[]
    for part in plan['parts']+(blank_references(plan) if include_blanks else []):
        name=part['name'];is_glass=name.startswith('Lens')
        color=part.get('color_rgb',[.25,.62,.87] if is_glass else [.35,.65,.36] if 'Spacer' in name else [.30,.32,.35] if 'Retainer' in name else [.32,.36,.42])
        if not isinstance(color,list) or len(color)!=3 or any(isinstance(x,bool) or not isinstance(x,(int,float)) or not math.isfinite(x) or not 0<=x<=1 for x in color):raise ValueError('Invalid color_rgb: '+name)
        profile=part['profile_zr_mm']
        if len(profile)<3 or any(len(p)!=2 or any(isinstance(x,bool) or not isinstance(x,(int,float)) or not math.isfinite(x) for x in p) or p[1]<0 for p in profile):raise ValueError('Invalid native profile: '+name)
        # Keep native profiles on one side of the axis; optical boundaries may touch it.
        cuts=[]
        for feature in plan.get('radial_features',[]):
            if feature['part']!=name or not feature.get('angles_degrees'):continue
            for field in ('diameter_mm','zs_mm','angles_degrees'):
                values=feature[field] if isinstance(feature[field],list) else [feature[field]]
                if not values or any(isinstance(v,bool) or not isinstance(v,(int,float)) or not math.isfinite(v) for v in values):raise ValueError('Invalid radial '+field)
            if feature['diameter_mm']<=0:raise ValueError('Invalid radial diameter')
            cuts.append(dict(kind='radial_holes',diameter_mm=feature['diameter_mm'],zs_mm=feature['zs_mm'],angles_degrees=feature['angles_degrees']))
        slots=part.get('wrench_slots')
        for window in part.get('radial_windows',[]):
            for field in ('width_mm','z_start_mm','z_end_mm'):
                v=window.get(field)
                if isinstance(v,bool) or not isinstance(v,(int,float)) or not math.isfinite(v):raise ValueError('Invalid radial window '+field)
            angles=window.get('angles_degrees',[])
            if not angles or any(isinstance(v,bool) or not isinstance(v,(int,float)) or not math.isfinite(v) for v in angles):raise ValueError('Invalid radial window angles')
            if window['width_mm']<=0 or window['z_end_mm']<=window['z_start_mm']:raise ValueError('Invalid radial window dimensions')
            cuts.append(dict(kind='radial_windows',**window))
        if slots:
            if part.get('cap_slot') not in ('front','rear'):raise ValueError('Wrench slot requires front/rear cap_slot')
            for field in ('width_mm','radial_depth_mm','axial_depth_mm'):
                v=slots.get(field)
                if isinstance(v,bool) or not isinstance(v,(int,float)) or not math.isfinite(v) or v<=0:raise ValueError('Invalid wrench slot '+field)
            cuts.append(dict(kind='wrench_slots',side=part['cap_slot'],**slots))
        payload.append(dict(name=name,role=part.get('role',''),profile_zr_mm=profile,color_rgb=color,is_glass=is_glass,cuts=cuts))
    data=dict(title=plan.get('title','Optomech'),source_origin_z_mm=float(offset),parts=payload,capabilities=report)
    template=(ROOT/'templates/freecad.FCMacro.in').read_text(encoding='utf-8')
    text=template.replace('__EMBEDDED_PLAN_JSON__',repr(json.dumps(data,ensure_ascii=True,separators=(',',':'))))
    if '__EMBEDDED_PLAN_JSON__' in text:raise ValueError('Unresolved FreeCAD template token.')
    ast.parse(text)
    return text,report
