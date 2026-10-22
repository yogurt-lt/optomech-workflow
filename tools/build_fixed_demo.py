from pathlib import Path
import sys,json,math
root=Path(__file__).resolve().parents[1];repo=root;sys.path.insert(0,str(repo))
from optomech.prescription import parse_zmx,face,optical_groups,step_metadata
from optomech.geometry import validate_plan,radial_ranges
out=root/'generated'/'legacy-fixed-projection';out.mkdir(parents=True,exist_ok=True)
s=parse_zmx((root/'examples'/'fixed-projection'/'source.zmx').read_bytes());offset=-s[2]['z_mm']
groups=optical_groups(s,step_metadata((root/'examples'/'fixed-projection'/'source.STP').read_text())['surface_pairs'])
def f(n,r):return face(s[n],r)+offset
G=[];lenses=[]
for i,seq in enumerate(groups):
    R=22.3 if i==2 else max(s[n]['mechanical_radius_mm'] for n in seq)
    g=dict(id=i+1,surfaces=seq,radius_mm=R)
    for side,n in [('front',seq[0]),('rear',seq[-1])]:
        opt=s[n]['optical_radius_mm'];r=max(opt+.05,R-.8) if R-opt>.35 else R-.3
        rr=[r+(R-r)*j/1000 for j in range(1001)];zz=[f(n,x) for x in rr]
        g[side]=dict(r=r,z=min(zz) if side=='front' else max(zz),flat=r>=opt)
    G.append(g)
    for a,b in zip(seq,seq[1:]):
        rr=sorted(set([R*j/600 for j in range(601)]+[s[a]['optical_radius_mm'],s[b]['optical_radius_mm']]))
        assert all(f(b,r)>f(a,r)-1e-6 for r in rr)
        lenses.append(dict(id=len(lenses)+1,name=f'surfaces {a},{b}',radius_mm=R,front=[[r,f(a,r)] for r in rr],rear=[[r,f(b,r)] for r in rr],
            polygon_zr_mm=[[f(a,r),r] for r in rr]+[[f(b,r),r] for r in rr[::-1]]))
seats=[dict(lens=1,z_mm=f(3,G[0]['radius_mm']),inner_radius_mm=26.5,outer_radius_mm=G[0]['radius_mm']),
       dict(lens=2,z_mm=f(5,G[1]['radius_mm']),inner_radius_mm=22.5,outer_radius_mm=G[1]['radius_mm']),
       dict(lens=3,z_mm=f(7,22.3),inner_radius_mm=21.8,outer_radius_mm=22.3)]
start=G[0]['front']['z']-3;end=G[-1]['rear']['z']+3
stop=next(v for v in s.values() if v['stop']);stopz=stop['z_mm']+offset;stopr=stop['optical_radius_mm']
z1,z2,z3=[p['z_mm'] for p in seats]
inner=[[start,G[0]['radius_mm']+.08],[z1,G[0]['radius_mm']+.08],[z1,26.5],[z1+.3,26.5],[z1+1.3,G[1]['radius_mm']+.08],
       [z2,G[1]['radius_mm']+.08],[z2,22.5],[z2+.3,22.5],[z2+1.,22.38],
       [z3,22.38],[z3,21.8],[z3+.5,21.8],[z3+1.5,19.05],
       [G[3]['front']['z']+.025,19.05],[G[3]['front']['z']+.3,G[3]['radius_mm']+.08],
       [G[4]['front']['z']+.05,G[3]['radius_mm']+.08],[G[4]['front']['z']+.8,G[4]['radius_mm']+.08],
       [G[4]['rear']['z'],G[4]['radius_mm']+.08],[G[4]['rear']['z'],G[4]['rear']['r']],
       [stopz-1.2,G[4]['rear']['r']],[stopz-.2,stopr+.05],[stopz-.15,stopr],[stopz+.15,stopr],[stopz+.2,stopr+.05],
       [stopz+1.2,G[5]['front']['r']],[G[5]['front']['z'],G[5]['front']['r']],[G[5]['front']['z'],G[5]['radius_mm']+.08]]
for a,b in zip(G[5:7],G[6:8]):inner += [[a['rear']['z']-.8,a['radius_mm']+.08],[a['rear']['z']-.05,b['radius_mm']+.08]]
inner += [[G[-1]['rear']['z']-.8,G[-1]['radius_mm']+.08],[G[-1]['rear']['z']-.1,G[-1]['radius_mm']+.55],[end,G[-1]['radius_mm']+.55]]
outer=[[start,30.9],[14.8,30.9],[15.8,26.3],[23.,26.3],[24.,25.5],[39.2,25.5],[40.2,21.6],[end,21.6]]
parts=[]
def part(name,poly,role):parts.append(dict(name=name,profile_zr_mm=poly,role=role))
part('01_BondedSeatBarrel',outer+inner[::-1],'L1-L3 separate rear shoulders and nine adhesive access ports; G4-G5 separate preload circuit')
zc=G[3]['front']['z']
part('02_G4G5_Retainer',[[zc-1.6,18.98],[zc,18.98],[zc,G[3]['front']['r']],[zc-1.6,G[3]['front']['r']]],'Independent front retainer for G4-G5; nominal M38 envelope only; no preload through bonded lenses')
a,b=G[3],G[4]
part('03_Spacer_G4_G5',[[a['rear']['z'],a['radius_mm']+.05],[b['front']['z'],a['radius_mm']+.05],[b['front']['z'],b['front']['r']],[a['rear']['z'],a['rear']['r']]],'Long downstream spacer; no L1-L2 or L2-L3 spacers')
for i,(a,b) in enumerate(zip(G[5:7],G[6:8])):
    part(f'{i+4:02d}_RearSpacer_{i+1}',[[a['rear']['z'],b['radius_mm']+.05],[b['front']['z'],b['radius_mm']+.05],[b['front']['z'],b['front']['r']],[a['rear']['z'],a['rear']['r']]],'Rear compression stack spacer')
rear=G[-1]['rear'];R=G[-1]['radius_mm']
part('06_RearRetainer',[[rear['z'],R+.48],[end-.3,R+.48],[end-.3,rear['r']],[rear['z'],rear['r']]],'Rear stack retainer; nominal thread envelope only')
ports=[dict(lens=i+1,z_mm=z,diameter_mm=2.,count=3,azimuth_spacing_deg=120,radial_gap_mm=.08) for i,z in enumerate([9.5,16.,20.1])]
plan=dict(title='PF-2 / Bonded seats v0.2',version='0.2.0',source_origin_z_mm=0.,parts=parts,lenses=lenses,groups=G,seats=seats,glue_ports=ports,
    reference_lens_ids=list(range(1,11)),blank_revisions=[dict(lens=3,change='Radius increased from 21.711578596 to 22.3 mm; peripheral flats extended outside original clear apertures only.')],
    assembly_scheme='L1-L3 individually seated and bonded; G4-G5 and G6-G8 each retain an independent compression path',
    assembly_order=dict(front=['G5','long G4-G5 spacer','G4','G4-G5 retainer','L3 on shoulder / bond','L2 on shoulder / bond','L1 on shoulder / bond'],rear=['G6','spacer','G7','spacer','G8','rear retainer']),
    stop=dict(z_mm=stopz,radius_mm=stopr,flat_width_mm=.3),
    warnings=['New macro includes nine native adhesive access holes; SolidWorks execution pending.','Only L3 blank is enlarged; L1 and L2 return to original STEP geometry.','No adhesive solid is modeled; port dimensions and adhesive process are provisional.','Optical clipping at remaining curved contacts, tolerances and bond stress remain to be verified.'])
result=validate_plan(plan);print('CHECKS',result);assert not result['errors']
# Verify insertion envelope conservatively for downstream rings and all lens groups.
paths=[(g['radius_mm'],g['front']['z'] if g['id']<=5 else g['rear']['z'],'front' if g['id']<=5 else 'rear',f"G{g['id']}") for g in G]
for p in parts[1:]:
    paths.append((max(q[1] for q in p['profile_zr_mm']),min(q[0] for q in p['profile_zr_mm']) if p['name'] in ['02_G4G5_Retainer','03_Spacer_G4_G5'] else max(q[0] for q in p['profile_zr_mm']), 'front' if p['name'] in ['02_G4G5_Retainer','03_Spacer_G4_G5'] else 'rear',p['name']))
for radius,z,side,name in paths:
    lo,hi=(start+.0001,z-.0001) if side=='front' else (z+.0001,end-.0001)
    for j in range(1201):
        zz=lo+(hi-lo)*j/1200;bore=radial_ranges(parts[0]['profile_zr_mm'],zz)[0][0]
        assert radius<bore+1e-6,(name,'blocked insertion',zz,radius,bore)
print('Insertion envelopes passed; seats:',seats)
(out/'plan.json').write_text(json.dumps(plan,indent=2))
