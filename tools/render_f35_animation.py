"""Render the exported native F35 meshes; requires numpy and Pillow."""
from pathlib import Path
import argparse,json,math,shutil
import numpy as np
from PIL import Image,ImageDraw,ImageFont

def render(mesh_path,design_path,dest):
 dest.mkdir(parents=True,exist_ok=True)
 data=json.loads(mesh_path.read_text('utf8'));design=json.loads(design_path.read_text('utf8'));plan=design['plan']
 colors={p['name']:p.get('color_rgb',[.43,.75,.89]) for p in plan['parts']+design['accessories']+plan['cad_reference_parts']}
 norm=lambda v:v/np.linalg.norm(v)
 view=norm(np.array([.7,-1,-.6]));right=norm(np.cross(view,[1,0,0]));up=np.cross(right,view);basis=np.array([right,up,view]).T;light=norm(np.array([1,-1,-.3]))
 W,H=1280,800
 ps=np.concatenate([np.array(m['vertices'])[np.unique(np.array(m['triangles']).ravel())]@basis for m in data['base'] if m['triangles']]);low=ps.min(axis=0);high=ps.max(axis=0)
 scale=min(1090/(high[0]-low[0]+7),570/(high[1]-low[1]+7));ox=W/2-(low[0]+high[0])/2*scale;oy=425+(low[1]+high[1])/2*scale
 font=lambda n:ImageFont.truetype('C:/Windows/Fonts/arial.ttf',n)
 def paint(canvas,depth,m,travel=0):
  if not m['vertices']:return
  verts=np.array(m['vertices'],dtype=float);verts[:,2]+=travel;coords=verts@basis;col=colors[m['name']]
  for tri in m['triangles']:
   a,b,c=verts[tri];n=np.cross(b-a,c-a);length=np.linalg.norm(n)
   if length<1e-10:continue
   shade=.57+.38*abs(np.dot(n/length,light));fill=np.minimum(255,np.round(255*np.array(col)*shade)).astype(np.uint8)
   pts=coords[tri];(x0,y0,z0),(x1,y1,z1),(x2,y2,z2)=[(ox+p[0]*scale,oy-p[1]*scale,p[2]) for p in pts]
   den=(y1-y2)*(x0-x2)+(x2-x1)*(y0-y2)
   if abs(den)<1e-9:continue
   xmin=max(0,math.floor(min(x0,x1,x2)));xmax=min(W-1,math.ceil(max(x0,x1,x2)));ymin=max(0,math.floor(min(y0,y1,y2)));ymax=min(H-1,math.ceil(max(y0,y1,y2)))
   if xmin>xmax or ymin>ymax:continue
   yy,xx=np.ogrid[ymin:ymax+1,xmin:xmax+1];xx=xx+.5;yy=yy+.5
   u=((y1-y2)*(xx-x2)+(x2-x1)*(yy-y2))/den;v=((y2-y0)*(xx-x2)+(x0-x2)*(yy-y2))/den;w=1-u-v;zz=u*z0+v*z1+w*z2
   sub=depth[ymin:ymax+1,xmin:xmax+1];mask=(u>=-1e-8)&(v>=-1e-8)&(w>=-1e-8)&(zz>sub+1e-6);sub[mask]=zz[mask];canvas[ymin:ymax+1,xmin:xmax+1][mask]=fill
 fixed=np.full((H,W,3),247,dtype=np.uint8);fixed_depth=np.full((H,W),-np.inf,np.float32)
 for m in data['base']:
  if m['group']=='fixed':paint(fixed,fixed_depth,m)
 all3=[];all2=[]
 for i,pose in enumerate(data['frames']):
  canvas=fixed.copy();depth=fixed_depth.copy()
  for m in data['base']:
   if m['group']=='moving':paint(canvas,depth,m,pose['travel'])
  for m in pose['rotating']:paint(canvas,depth,m)
  im=Image.fromarray(canvas)
  section=Image.new('RGB',(W,H),'#f7f7f7');ds=ImageDraw.Draw(section)
  # Native closed wires with even/odd fills preserve holes and multiple islands.
  for m in [x for x in data['base'] if x['group'] not in ('cam','pinion')]+pose['rotating']:
   mask=Image.new('1',(W,H),0)
   for loop in m['loops']:
    if len(loop)<3:continue
    one=Image.new('1',(W,H),0);d1=ImageDraw.Draw(one);shift=pose['travel'] if m['group']=='moving' else 0
    d1.polygon([(115+(z+3+shift)*14,435-r*7.1) for z,r in loop],fill=1)
    from PIL import ImageChops
    mask=ImageChops.logical_xor(mask,one)
   color=tuple(round(c*255) for c in colors[m['name']]);section.paste(color,(0,0),mask.convert('L'))
  ds=ImageDraw.Draw(section);ds.line((90,435,1160,435),fill='#a8b8c4',width=1)
  for result in (im,section):
   d=ImageDraw.Draw(result);d.text((35,23),'F35 F2.8 v0.2 | Native CAD '+('cutaway' if result is im else 'section'),font=font(27),fill='#263f51')
   d.text((35,65),f"Cam {pose['angle']:+d} deg  |  Input {pose['pinion_angle']:+d} deg  |  L2 {pose['travel']:+.2f} mm",font=font(22),fill='#187c83')
   cx,cy=1180,95;r=38;rad=math.radians(pose['angle']-90);d.ellipse((cx-r,cy-r,cx+r,cy+r),outline='#9caeba',width=2);d.line((cx,cy,cx+r*.85*math.cos(rad),cy+r*.85*math.sin(rad)),fill='#c58a32',width=5)
   d.text((35,743),'Fixed housing + other optics | Rotating gold cam + green pinion | Translating blue L2 cell',font=font(19),fill='#536d7f')
   d.text((35,772),'4 mm geometric travel; focus distances and the lens-surface aperture require optical validation.',font=font(16),fill='#6d8090')
  im.save(dest/f'3d-{i:02d}.png',optimize=True);section.save(dest/f'frame-{i:02d}.png',optimize=True);all3.append(im);all2.append(section);print('Rendered F35',pose['angle'],flush=True)
 for frames,name in [(all3,'electric-focus-3d.gif'),(all2,'electric-focus.gif')]:
  ping=frames+frames[-2:0:-1];ping[0].save(dest/name,save_all=True,append_images=ping[1:],duration=[600 if i in (0,len(frames)-1) else 130 for i in range(len(ping))],loop=0,optimize=True)
 meta=dict(angles=[p['angle'] for p in data['frames']],travel_mm=[p['travel'] for p in data['frames']],pinion_angles=[p['pinion_angle'] for p in data['frames']],frames=[f'frame-{i:02d}.png' for i in range(21)],frames_3d=[f'3d-{i:02d}.png' for i in range(21)],gif='electric-focus.gif',gif_3d='electric-focus-3d.gif',initial_index=10,preset_indices=[0,10,20],provenance='Native FreeCAD Body Tip shapes transformed once to global coordinates; rotating solids re-cut at every pose. Fixed camera and section plane.')
 if (dest/'motion.json').is_file():
  previous=json.loads((dest/'motion.json').read_text('utf8'))
  if 'copy' in previous:meta['copy']=previous['copy']
 (dest/'motion.json').write_text(json.dumps(meta,ensure_ascii=False,indent=2),'utf8')
 for src,target in [('F35_section.png','section.png'),('F35_rim_contact_detail.png','detail.png')]:shutil.copy2(design_path.parent/src,dest/target)

if __name__=='__main__':
 p=argparse.ArgumentParser(description=__doc__);p.add_argument('--meshes',type=Path,required=True);p.add_argument('--design',type=Path,required=True);p.add_argument('--output',type=Path,required=True);a=p.parse_args();render(a.meshes,a.design,a.output)
