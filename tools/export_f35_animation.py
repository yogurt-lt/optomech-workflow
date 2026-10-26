"""Run with FreeCAD's Python. Export F35 native poses without saving the model."""
from pathlib import Path
import argparse,json
import FreeCAD as A,Part

def export(model,design,output):
 data=json.loads(design.read_text('utf8'));plan=data.get('plan',data)
 items=plan['parts']+data.get('accessories',[])+plan['cad_reference_parts']
 groups={p['name']:p.get('group','fixed') for p in items}
 d=A.openDocument(str(model.resolve()));control=d.getObject('FocusControl')
 bodies=[b for b in d.Objects if b.TypeId=='PartDesign::Body'];assert len(bodies)==24
 assert all(b.Label in groups for b in bodies)
 half=Part.makeBox(220,110,220,A.Vector(-110,0,-60))
 plane=Part.makePlane(220,220,A.Vector(-110,0,-60),A.Vector(0,1,0))
 def geometry(b):
  shape=b.Tip.Shape.copy();shape.transformShape(b.getGlobalPlacement().toMatrix(),True)
  cut=shape.common(half);verts,triangles=cut.tessellate(.18)
  loops=[]
  for edges in Part.sortEdges(shape.section(plane).Edges):
   try: wire=Part.Wire(edges)
   except Part.OCCError: continue
   if wire.isClosed():loops.append([[p.z,p.x] for p in wire.discretize(Deflection=.025)])
  return dict(name=b.Label,group=groups[b.Label],vertices=[[p.x,p.y,p.z] for p in verts],triangles=triangles,loops=loops)
 control.AngleDeg=0;d.recompute();base=[geometry(b) for b in bodies]
 frames=[]
 for angle in range(-40,41,4):
  control.AngleDeg=angle;d.recompute()
  frames.append(dict(angle=angle,travel=control.Travel.Value,pinion_angle=-5*angle,rotating=[geometry(b) for b in bodies if groups[b.Label] in ('cam','pinion')]))
  print('Exported native F35 pose',angle,flush=True)
 output.write_text(json.dumps(dict(base=base,frames=frames),separators=(',',':')),'utf8')
 # Deliberately do not save: all changes above are transient animation poses.
 print('Saved animation mesh data',output,flush=True)

if __name__=='__main__':
 p=argparse.ArgumentParser(description=__doc__);p.add_argument('--model',type=Path,required=True);p.add_argument('--design',type=Path,required=True);p.add_argument('--output',type=Path,required=True);a=p.parse_args();export(a.model,a.design,a.output)

