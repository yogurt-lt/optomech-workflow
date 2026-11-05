"""Restore exact large demo assets from small, hash-checked repository chunks."""
from pathlib import Path
import hashlib,json,os

def restore_large_assets(root):
    root=Path(root).resolve()
    manifest=root/'assets/large-files.json'
    if not manifest.exists():return
    def local(name):
        path=(root/name).resolve()
        if not path.is_relative_to(root):raise ValueError('Asset path leaves repository')
        return path
    def digest(path):
        h=hashlib.sha256()
        with path.open('rb') as f:
            for block in iter(lambda:f.read(1024*1024),b''):h.update(block)
        return h.hexdigest()
    for name,spec in json.loads(manifest.read_text('utf8')).items():
        target=local(name)
        if target.exists():
            if target.stat().st_size!=spec['bytes'] or digest(target)!=spec['sha256']:
                raise ValueError('Existing asset differs from reviewed chunks: '+name)
            continue
        target.parent.mkdir(parents=True,exist_ok=True)
        temp=target.with_name(target.name+'.assembling')
        try:
            with temp.open('wb') as out:
                for chunk in spec['parts']:
                    raw=local(chunk['path']).read_bytes()
                    if hashlib.sha256(raw).hexdigest()!=chunk['sha256']:raise ValueError('Asset chunk hash mismatch')
                    out.write(raw)
            if temp.stat().st_size!=spec['bytes'] or digest(temp)!=spec['sha256']:raise ValueError('Restored asset hash mismatch')
            os.replace(temp,target)
            print('Restored reviewed asset:',name)
        finally:
            if temp.exists():temp.unlink()
