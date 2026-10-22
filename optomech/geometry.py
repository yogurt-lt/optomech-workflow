"""Meridional concept checks, not a CAD kernel or manufacturing certification."""
import math

def radial_ranges(poly,z):
    xs=[]
    for (a,b),(c,d) in zip(poly,poly[1:]+poly[:1]):
        if min(a,c)<=z<max(a,c):xs.append(b+(d-b)*(z-a)/(c-a))
    xs.sort()
    return list(zip(xs[::2],xs[1::2]))

def point_inside(poly,z,r,eps=1e-7):
    return any(a+eps<r<b-eps for a,b in radial_ranges(poly,z))

def polygon_area(poly):return sum(a*d-c*b for (a,b),(c,d) in zip(poly,poly[1:]+poly[:1]))/2

def _cross(a,b,c):return (b[0]-a[0])*(c[1]-a[1])-(b[1]-a[1])*(c[0]-a[0])

def self_intersects(poly):
    n=len(poly)
    for i in range(n):
        a,b=poly[i],poly[(i+1)%n]
        if math.dist(a,b)<1e-9:return True
        for j in range(i+1,n):
            if j==i+1 or (i==0 and j==n-1):continue
            c,d=poly[j],poly[(j+1)%n]
            if max(min(a[0],b[0]),min(c[0],d[0]))>min(max(a[0],b[0]),max(c[0],d[0]))+1e-9:continue
            if max(min(a[1],b[1]),min(c[1],d[1]))>min(max(a[1],b[1]),max(c[1],d[1]))+1e-9:continue
            x,y=_cross(a,b,c),_cross(a,b,d);u,v=_cross(c,d,a),_cross(c,d,b)
            if x*y<=0 and u*v<=0:return True
    return False

def validate_plan(plan,full=True):
    errors=[];warnings=[];parts=plan.get('parts',[]);names=set()
    if not 1<=len(parts)<=60:return dict(errors=['Plan requires 1–60 parts.'],warnings=[])
    for p in parts:
        name=p.get('name','');poly=p.get('profile_zr_mm',[])
        if name in names or not name or len(name)>80 or not all(c.isascii() and (c.isalnum() or c in '_-') for c in name):errors.append('Part names must be unique ASCII identifiers.')
        names.add(name)
        if not 3<=len(poly)<=300:errors.append(name+': requires 3–300 vertices.');continue
        if any(not isinstance(q,list) or len(q)!=2 or any(isinstance(v,bool) or not isinstance(v,(int,float)) or not math.isfinite(v) or abs(v)>100000 for v in q) or q[1]<=0 for q in poly):errors.append(name+': invalid finite (Z, radius) vertices.');continue
        if abs(polygon_area(poly))<1e-8 or self_intersects(poly):errors.append(name+': degenerate or self-intersecting profile.')
    if errors:return dict(errors=errors,warnings=warnings)
    for i,a in enumerate(parts):
        for b in parts[i+1:]:
            zs=sorted({q[0] for p in (a,b) for q in p['profile_zr_mm']})
            hit=False
            for x,y in zip(zs,zs[1:]):
                for t in (.01,.25,.5,.75,.99):
                    z=x+(y-x)*t
                    for al,ah in radial_ranges(a['profile_zr_mm'],z):
                        for bl,bh in radial_ranges(b['profile_zr_mm'],z):
                            if min(ah,bh)-max(al,bl)>1e-6:hit=True
            if hit:errors.append(f"Mechanical overlap: {a['name']} / {b['name']}")
    optical_hits=[]
    if full:
        for p in parts:
            for lens in plan.get('lenses',[]):
                poly=lens.get('polygon_zr_mm',[])
                if not poly:continue
                # Radial scan of glass, including the flat peripheral parts of optical faces.
                zs=sorted({q[0] for q in poly})
                hit=False
                for z in zs[::max(1,len(zs)//250)]:
                    for al,ah in radial_ranges(poly,z):
                        for bl,bh in radial_ranges(p['profile_zr_mm'],z):
                            if min(ah,bh)-max(al,bl)>1e-5:hit=True;break
                    if hit:break
                if hit:optical_hits.append(f"{p['name']} / L{lens.get('id','?')}")
        errors.extend('Sampled glass overlap: '+s for s in optical_hits)
    warnings.append('Sampled meridional checks only; CAD interference, fits and contact stress require review.')
    if not plan.get('rays'):warnings.append('No embedded rays for rechecking edited profiles. Existing ray reports apply only to the supplied baseline.')
    return dict(errors=errors,warnings=warnings)
