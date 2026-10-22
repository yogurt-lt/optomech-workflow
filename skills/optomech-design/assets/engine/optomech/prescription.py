"""Read a limited, explicit subset of sequential Zemax text prescriptions.

No OpticStudio evaluation, solve execution, glass catalog lookup, or ray trace.
"""
import math
import re

def parse_zmx(raw):
    if isinstance(raw, bytes):
        raw=raw.decode('utf-16' if raw[:2] in (b'\xff\xfe',b'\xfe\xff') else 'utf-8-sig')
    raw=raw.replace('\r','')
    if not re.search(r'^UNIT MM\b',raw,re.M):raise ValueError('Only millimeter prescriptions are supported.')
    if not re.search(r'^MODE SEQ\b',raw,re.M):raise ValueError('Only sequential mode is supported.')
    surfaces={};z=0.0
    for m in re.finditer(r'^SURF (\d+)\s*\n(.*?)(?=^SURF |^CONF |\Z)',raw,re.M|re.S):
        n=int(m[1]);block=m[2]
        def field(k,default='0'):
            q=re.search(r'^\s+'+k+r'\s+([^\n]+)',block,re.M)
            return q[1].split()[0] if q else default
        kind=field('TYPE','STANDARD')
        if kind not in ('STANDARD','EVENASPH'):raise ValueError(f'Unsupported surface type {kind} at S{n}.')
        p={int(k):float(v) for k,v in re.findall(r'^\s+PARM\s+(\d+)\s+(\S+)',block,re.M)}
        surfaces[n]=dict(number=n,z_mm=z if n else None,curvature=float(field('CURV')),conic=float(field('CONI')),
            parameters=p,type=kind,optical_radius_mm=float(field('DIAM')),mechanical_radius_mm=float(field('MEMA')),
            thickness_mm=float(field('DISZ')),glass=field('GLAS',''),stop=bool(re.search(r'^\s+STOP\s*$',block,re.M)))
        if n:z+=surfaces[n]['thickness_mm']
    if not surfaces:raise ValueError('No sequential surfaces found.')
    return surfaces

def face(surface,r):
    """Peripheral flat outside DIAM follows the demonstrated STEP export convention."""
    r=min(r,surface['optical_radius_mm']);c=surface['curvature'];q=1-(1+surface['conic'])*c*c*r*r
    if q<0:raise ValueError('Sag outside real conic domain.')
    z=c*r*r/(1+math.sqrt(q))
    if surface['type']=='EVENASPH':z+=sum(float(v)*r**(2*int(k)) for k,v in surface['parameters'].items() if int(k)>=1)
    return surface['z_mm']+z

def step_metadata(text):
    if len(text)>40_000_000:raise ValueError('STEP exceeds 40 MB analysis limit.')
    if 'ISO-10303-21;' not in text:raise ValueError('Not a STEP text file.')
    pairs=[(int(a),int(b)) for a,b in re.findall(r"ADVANCED_BREP_SHAPE_REPRESENTATION\('surfaces (\d+),(\d+)'",text)]
    return dict(solid_definitions=len(re.findall(r'=MANIFOLD_SOLID_BREP\(',text)),surface_pairs=pairs,
        millimeters='SI_UNIT(.MILLI.,.METRE.)' in text,ray_wireframe="GEOMETRICALLY_BOUNDED_WIREFRAME_SHAPE_REPRESENTATION('rays'" in text)

def optical_groups(surfaces,pairs):
    """Use exported media only, joining consecutive glass intervals at shared faces."""
    groups=[]
    for a,b in sorted(pairs):
        if b!=a+1 or a not in surfaces or b not in surfaces:raise ValueError('Surface pair is not supported.')
        if not surfaces[a]['glass']:raise ValueError(f'No glass after S{a}.')
        if groups and groups[-1][-1]==a:groups[-1].append(b)
        else:groups.append([a,b])
    return groups

def assembly_candidates(radii):
    """Suggest radius-compatible loading only; no claim of optimal design."""
    if not radii:return []
    eps=1e-5;choices=[]
    if all(a>=b-eps for a,b in zip(radii,radii[1:])):choices.append(dict(kind='front',reason='Radii decrease from front to rear.'))
    if all(a<=b+eps for a,b in zip(radii,radii[1:])):choices.append(dict(kind='rear',reason='Radii increase from front to rear.'))
    for i in range(1,len(radii)):
        left=radii[:i];right=radii[i:]
        if all(a>=b-eps for a,b in zip(left,left[1:])) and all(a<=b+eps for a,b in zip(right,right[1:])):
            choices.append(dict(kind='two-ended',split_after_group=i,reason='Each half narrows toward the split.'))
    return choices or [dict(kind='review',reason='Consider separate barrels or centering adapters; a simple radius ordering is insufficient.')]
