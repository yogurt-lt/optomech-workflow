"""Emit native SolidWorks VBA from an explicit reviewed plan."""
from pathlib import Path
import re,math
from .geometry import validate_plan
from .native_glue import glue_macro_code

ROOT=Path(__file__).resolve().parent.parent

def blank_references(plan):
    if not plan.get('reference_lens_ids') and not plan.get('blank_revisions'):return []
    result=[]
    ids=plan.get('reference_lens_ids') or [r['lens'] for r in plan['blank_revisions']]
    revised={r['lens'] for r in plan.get('blank_revisions',[])}
    for lens in plan['lenses']:
        if lens['id'] not in ids:continue
        poly=lens['polygon_zr_mm']
        # Adaptive display-only simplification, within 0.0005 mm of the sampled outline.
        indices={0,len(poly)-1};stack=[(0,len(poly)-1)]
        while stack:
            lo,hi=stack.pop();a,b=poly[lo],poly[hi];dx=b[0]-a[0];dy=b[1]-a[1];den=dx*dx+dy*dy
            best=0.;index=lo
            for j in range(lo+1,hi):
                p=poly[j];t=max(0,min(1,((p[0]-a[0])*dx+(p[1]-a[1])*dy)/den)) if den else 0.
                dist=((p[0]-a[0]-t*dx)**2+(p[1]-a[1]-t*dy)**2)**.5
                if dist>best:best=dist;index=j
            if best>.0005:indices.add(index);stack.extend([(lo,index),(index,hi)])
        pts=[poly[i] for i in sorted(indices)]
        # Front and rear both end at the axis. Do not duplicate the closing axis point.
        suffix='RevisedBlank_Reference' if lens['id'] in revised else 'Optical_Reference'
        result.append(dict(name=f"Lens{lens['id']}_{suffix}",profile_zr_mm=pts,role='Polyline reference only; optical surfaces remain defined by original prescription'))
    return result

def generate_macro(plan,module='Optomech_Demo',include_blanks=False):
    if not re.fullmatch(r'[A-Za-z][A-Za-z0-9_]{0,60}',module):raise ValueError('Invalid VBA module name.')
    checks=validate_plan(plan)
    if checks['errors']:raise ValueError('\n'.join(checks['errors']))
    parts=plan['parts']+(blank_references(plan) if include_blanks else [])
    offset=float(plan.get('source_origin_z_mm',0.))
    if not math.isfinite(offset) or abs(offset)>100000:raise ValueError('Coordinate offset must be finite and within 100000 mm.')
    zs=[p[0]+offset for part in parts for p in part['profile_zr_mm']]
    text=(ROOT/'templates'/'solidworks.bas.in').read_text(encoding='ascii')
    tokens={'__MODULE__':module,'__LAST_INDEX__':str(len(parts)-1),'__NAMES__':'Array('+', '.join('"'+p['name']+'"' for p in parts)+')',
        '__AXIS_MIN__':f'{min(zs)-1:.12f}','__AXIS_MAX__':f'{max(zs)+1:.12f}'}
    for a,b in tokens.items():text=text.replace(a,b)
    if plan.get('glue_ports'):
        name=plan['parts'][0]['name']
        text=text.replace('    currentStage = label & ": check geometry"',f'    If label = "{name}" Then AddAdhesivePorts model, plane, axis\n    currentStage = label & ": check geometry"')
        text+=glue_macro_code(plan['glue_ports'],offset)
    # Cosmetic colors do not affect model validity; keep API errors out of the build path.
    text=text.replace('    currentStage = label & ": save"','    ApplyConceptColor model, label\n    currentStage = label & ": save"')
    text=text.replace('        log.WriteLine "Inserted " & names(i)', '        log.WriteLine "Inserted " & names(i)\n        ApplyComponentColor component, CStr(names(i))')
    text+=(ROOT/'templates'/'colors.bas.in').read_text(encoding='ascii')
    text+='''
Private Sub ApplyConceptColor(ByVal model As SldWorks.ModelDoc2, ByVal label As String)
    On Error Resume Next
    model.MaterialPropertyValues = ConceptPalette(label)
    On Error GoTo 0
End Sub
'''
    rows=[]
    for i,part in enumerate(parts):
        rows.append(f'        Case {i}: ProfileFor = Profile_{i}()')
    text+='Private Function ProfileFor(ByVal index As Long) As Variant\n    Select Case index\n'+'\n'.join(rows)+'\n        Case Else: Err.Raise vbObjectError + 418, , "Unknown profile index."\n    End Select\nEnd Function\n'
    for i,part in enumerate(parts):
        points=part['profile_zr_mm']
        text+=f'\nPrivate Function Profile_{i}() As Variant\n    Dim pts({len(points)-1}) As Variant\n'
        for j,(z,r) in enumerate(points):text+=f'    pts({j}) = Array({z+offset:.12f}, {r:.12f})\n'
        text+=f'    Profile_{i} = pts\nEnd Function\n'
    if '__' in text:raise ValueError('Unresolved template token.')
    if max(map(len,text.splitlines()))>1000:raise ValueError('VBA line length limit.')
    if text.count('0, SAVE_SILENT)')!=2:raise ValueError('Save option invariant failed.')
    return text
