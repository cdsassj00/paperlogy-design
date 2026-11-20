import sys,zipfile,posixpath,json
from lxml import etree as E
from pathlib import Path
R='http://schemas.openxmlformats.org/officeDocument/2006/relationships';A='http://schemas.openxmlformats.org/drawingml/2006/main';P='http://schemas.openxmlformats.org/presentationml/2006/main'
def target(p,t):return t.lstrip('/') if t.startswith('/') or not p else posixpath.normpath(posixpath.dirname(p)+'/'+t)
def check(path):
 errors=[];texts=0;shapes=0;pics=0
 with zipfile.ZipFile(path) as z:
  names=set(z.namelist())
  if len(names)!=len(z.namelist()):errors.append('duplicate ZIP entries')
  for p in names:
   if p.endswith('.rels'):
    owner=p.replace('/_rels/','/')[:-5] if p!='_rels/.rels' else '';rr=E.fromstring(z.read(p));ids={r.get('Id') for r in rr}
    for r in rr:
     if r.get('TargetMode')!='External' and target(owner,r.get('Target')) not in names:errors.append('missing target '+p+' '+r.get('Target'))
    if owner in names and owner.endswith('.xml'):
     x=E.fromstring(z.read(owner))
     for e in x.iter():
      for k,v in e.attrib.items():
       if k.startswith('{'+R+'}') and v not in ids:errors.append('missing rId '+owner+' '+v)
   if 'replicaSlides/' in p and p.endswith('.xml'):
    x=E.fromstring(z.read(p));texts+=len(x.findall('.//{'+A+'}t'));shapes+=len(x.findall('.//{'+P+'}sp'));pics+=len(x.findall('.//{'+P+'}pic'))
  video=[p for p in names if Path(p).suffix.lower() in ['.mp4','.mov','.wav','.mp3','.avi','.wmv']]
  pres=E.fromstring(z.read('ppt/presentation.xml'));count=len(pres.find('{'+P+'}sldIdLst'))
 return {'file':str(path),'bytes':Path(path).stat().st_size,'slides':count,'text_runs':texts,'native_shapes':shapes,'pictures':pics,'media_binaries':video,'errors':errors}
if __name__=='__main__':
 result=check(sys.argv[1]);print(json.dumps(result,ensure_ascii=False,indent=2));sys.exit(bool(result['errors']))

