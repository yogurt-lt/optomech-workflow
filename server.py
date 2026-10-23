"""Local-only UI. No cloud service, AI key, CAD installation or third-party Python packages needed for preview."""
import argparse,base64,json,webbrowser
from http.server import BaseHTTPRequestHandler,ThreadingHTTPServer
from pathlib import Path
from optomech.__main__ import analyze
from optomech.generator import generate_macro
from optomech.freecad_generator import generate_freecad_macro
from optomech.geometry import validate_plan

ROOT=Path(__file__).resolve().parent
DEMOS=json.loads((ROOT/'examples/catalog.json').read_text(encoding='utf-8'))
ENGLISH=json.loads((ROOT/'examples/catalog.en.json').read_text(encoding='utf-8'))
for key,item in DEMOS.items():item['en']=ENGLISH[key]

class Handler(BaseHTTPRequestHandler):
    def send(self,status,content,mime='application/json; charset=utf-8'):
        if not isinstance(content,bytes):content=json.dumps(content,ensure_ascii=False).encode('utf-8')
        self.send_response(status);self.send_header('Content-Type',mime);self.send_header('Content-Length',str(len(content)))
        self.send_header('X-Content-Type-Options','nosniff');self.send_header('Cache-Control','no-store');self.end_headers();self.wfile.write(content)
    def do_GET(self):
        route=self.path.split('?')[0]
        if route=='/':return self.send(200,(ROOT/'web'/'index.html').read_bytes(),'text/html; charset=utf-8')
        if route=='/focus.html':return self.send(200,(ROOT/'web/focus.html').read_bytes(),'text/html; charset=utf-8')
        if route=='/focus-motion.js':return self.send(200,(ROOT/'web/focus-motion.js').read_bytes(),'text/javascript; charset=utf-8')
        if route.startswith('/focus-assets/'):
            name=route.removeprefix('/focus-assets/')
            if '/' not in name and '\\' not in name and name not in ('.','..'):
                path=ROOT/'web/focus-assets'/name
                mime={'.png':'image/png','.gif':'image/gif','.json':'application/json; charset=utf-8'}.get(path.suffix)
                if mime and path.is_file():return self.send(200,path.read_bytes(),mime)
        if route=='/view3d.js':return self.send(200,(ROOT/'web'/'view3d.js').read_bytes(),'text/javascript; charset=utf-8')
        if route=='/i18n.js':return self.send(200,(ROOT/'web/i18n.js').read_bytes(),'text/javascript; charset=utf-8')
        if route=='/downloads/optomech-design-skill.zip':return self.send(200,(ROOT/'downloads/optomech-design-skill.zip').read_bytes(),'application/zip')
        if route=='/skill-guide':
            name='SKILL_USAGE.en.md' if 'lang=en' in self.path else 'SKILL_USAGE.md'
            return self.send(200,(ROOT/'docs'/name).read_bytes(),'text/markdown; charset=utf-8')
        if route=='/api/demos':return self.send(200,DEMOS)
        if route.startswith('/api/plan/'):
            key=route.removeprefix('/api/plan/')
            if key in DEMOS:return self.send(200,(ROOT/'examples'/key/'plan.json').read_bytes())
        if route.startswith('/examples/'):
            segments=route.split('/')
            if len(segments)==4:
                _,_,key,name=segments
                if key in DEMOS and name in {'section.png','section.en.svg','DESIGN.md','DESIGN.en.md','plan.json','checks.json','source.STP','source.zmx','FreeCAD.FCMacro','FreeCAD_Model.FCStd','FREECAD.md','FREECAD.en.md','freecad_report.json',DEMOS[key]['full_macro'],DEMOS[key]['mechanics_macro']}:
                    path=ROOT/'examples'/key/name
                    if path.is_file():
                        mime={'.png':'image/png','.svg':'image/svg+xml; charset=utf-8','.md':'text/markdown; charset=utf-8','.bas':'text/plain; charset=us-ascii','.FCMacro':'text/plain; charset=utf-8','.json':'application/json; charset=utf-8'}.get(path.suffix,'application/octet-stream')
                        return self.send(200,path.read_bytes(),mime)
        return self.send(404,{'error':'Not found'})
    def do_POST(self):
        # Local browser only; block cross-origin sites posting to this loopback service.
        try:
            size=int(self.headers.get('Content-Length','0'))
            if not 0<size<45_000_000:raise ValueError('Request size limit is 45 MB.')
            if 'application/json' not in self.headers.get('Content-Type',''):raise ValueError('JSON is required.')
            payload=json.loads(self.rfile.read(size))
            origin=self.headers.get('Origin')
            if origin and origin not in (f'http://127.0.0.1:{self.server.server_port}',f'http://localhost:{self.server.server_port}'):
                return self.send(403,{'error':'Cross-origin requests are disabled.'})
            if self.path=='/api/analyze':
                raw=base64.b64decode(payload.get('zmx_base64',''),validate=True) or None
                return self.send(200,analyze(payload['step'],raw))
            if self.path in ('/api/check','/api/generate'):
                plan=payload['plan']
                if len(plan.get('lenses',[]))>30 or sum(len(l.get('polygon_zr_mm',[])) for l in plan.get('lenses',[]))>100000:raise ValueError('Lens sample limit exceeded.')
                result=validate_plan(plan)
                if self.path=='/api/check':return self.send(200,result)
                if result['errors']:return self.send(422,result)
                cad=payload.get('cad','solidworks')
                if cad not in ('solidworks','freecad'):raise ValueError('Choose solidworks or freecad.')
                if cad=='freecad':
                    macro,capabilities=generate_freecad_macro(plan,include_blanks=bool(payload.get('include_blanks')))
                    return self.send(200,{'macro':macro,'filename':'Optomech_Web.FCMacro','cad':cad,'capabilities':capabilities,'checks':result})
                if plan.get('radial_features') or any(p.get('cap_slot') or p.get('radial_windows') for p in plan['parts']):
                    result['warnings'].append('General SolidWorks export contains revolved bodies only for dedicated metadata; use the reviewed case macro for complete holes/slots.')
                return self.send(200,{'macro':generate_macro(plan,module='Optomech_Web',include_blanks=bool(payload.get('include_blanks'))),'filename':'Optomech_Web.bas','cad':cad,'checks':result})
            self.send(404,{'error':'Not found'})
        except (ValueError,KeyError,TypeError,IndexError,OverflowError) as e:self.send(400,{'error':str(e)})

def main():
    parser=argparse.ArgumentParser();parser.add_argument('--port',type=int,default=8765);parser.add_argument('--no-browser',action='store_true');args=parser.parse_args()
    server=ThreadingHTTPServer(('127.0.0.1',args.port),Handler)
    url=f'http://127.0.0.1:{server.server_port}';print('Optomech workflow: '+url,flush=True)
    if not args.no_browser:webbrowser.open(url)
    try:server.serve_forever()
    except KeyboardInterrupt:pass
    finally:server.server_close()

if __name__=='__main__':main()
