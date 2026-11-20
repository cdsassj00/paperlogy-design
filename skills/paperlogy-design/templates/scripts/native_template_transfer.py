"""Native OOXML template transfer. Never replaces a page with a screenshot."""
import argparse,copy,hashlib,json,posixpath,zipfile,io
from PIL import Image
from pathlib import Path
from lxml import etree as E
P='http://schemas.openxmlformats.org/presentationml/2006/main'
A='http://schemas.openxmlformats.org/drawingml/2006/main'
R='http://schemas.openxmlformats.org/officeDocument/2006/relationships'
PK='http://schemas.openxmlformats.org/package/2006/relationships'
CT='http://schemas.openxmlformats.org/package/2006/content-types'
NS={'p':P,'a':A,'r':R}
def xml(b):return E.fromstring(b)
def dump(e):return E.tostring(e,encoding='UTF-8',xml_declaration=True,standalone=True)
def relpath(p):return posixpath.join(posixpath.dirname(p),'_rels',posixpath.basename(p)+'.rels')
def target(part,t):return t.lstrip('/') if t.startswith('/') else posixpath.normpath(posixpath.join(posixpath.dirname(part),t))
def text_replace(shape,value):
 body=shape.find('p:txBody',NS);old=body.findall('a:p',NS); lines=value.split('\n')
 if not old:raise ValueError('No paragraph template')
 for p in old:body.remove(p)
 for i,line in enumerate(lines):
  base=old[min(i,len(old)-1)];p=E.Element('{'+A+'}p')
  pp=base.find('a:pPr',NS)
  if pp is not None:p.append(copy.deepcopy(pp))
  r=E.SubElement(p,'{'+A+'}r')
  rp=base.find('a:r/a:rPr',NS)
  if rp is None:rp=base.find('a:fld/a:rPr',NS)
  if rp is None:rp=base.find('a:endParaRPr',NS)
  if rp is not None:
   rp=copy.deepcopy(rp);rp.tag='{'+A+'}rPr';r.append(rp)
  E.SubElement(r,'{'+A+'}t').text=line
  body.append(p)
def main():
 ap=argparse.ArgumentParser();ap.add_argument('source');ap.add_argument('spec');ap.add_argument('--out',required=True);ap.add_argument('--reference',action='store_true');args=ap.parse_args()
 src=Path(args.source);out=Path(args.out);spec=json.loads(Path(args.spec).read_text(encoding='utf-8-sig'))
 if out.exists():raise FileExistsError(out)
 out.parent.mkdir(parents=True,exist_ok=True)
 report={'mode':'native-template-transfer','reference_mode':args.reference,'source':str(src),'source_bytes':src.stat().st_size,'source_mtime_ns':src.stat().st_mtime_ns,'spec_sha256':hashlib.sha256(Path(args.spec).read_bytes()).hexdigest(),'script_sha256':hashlib.sha256(Path(__file__).read_bytes()).hexdigest(),'slides':[],'warnings':['Notes and animations removed for a static editable sample. Fonts retain source declarations; render substitution must be checked.']}
 with zipfile.ZipFile(src) as z:
  names=set(z.namelist());pres=xml(z.read('ppt/presentation.xml'));pr=xml(z.read('ppt/_rels/presentation.xml.rels'));prmap={e.get('Id'):e for e in pr}
  oldids=list(pres.find('p:sldIdLst',NS));parts={};queue=[]
  for idx,item in enumerate(spec['slides'],1):
   old=target('ppt/presentation.xml',prmap[oldids[item['source_page']-1].get('{'+R+'}id')].get('Target'))
   root=xml(z.read(old));before=copy.deepcopy(root)
   shapes={e.find('p:nvSpPr/p:cNvPr',NS).get('id'):e for e in root.findall('.//p:sp',NS)}
   actual={k:''.join(e.xpath('.//a:t/text()',namespaces=NS)) for k,e in shapes.items() if ''.join(e.xpath('.//a:t/text()',namespaces=NS)).strip()}
   if set(actual)!=set(item['text']):raise ValueError(f'Every text shape must be assigned: page {item["source_page"]}, missing {set(actual)-set(item["text"])} extra {set(item["text"])-set(actual)}')
   if not args.reference:
    for sid,value in item['text'].items():text_replace(shapes[sid],value)
   if spec.get('font_fallback'):
    for body in root.findall('.//p:txBody',NS):
     for rp in body.xpath('.//a:rPr | .//a:defRPr | .//a:endParaRPr',namespaces=NS):
      oldfonts=rp.xpath('./a:latin/@typeface | ./a:ea/@typeface',namespaces=NS)
      if any('bold' in f.lower() for f in oldfonts):rp.set('b','1')
      for tag in ['latin','ea','cs']:
       font=rp.find('{'+A+'}'+tag)
       if font is None:font=E.SubElement(rp,'{'+A+'}'+tag)
       font.set('typeface',spec['font_fallback'])
   for tag in ['timing','transition','extLst']:
    for e in root.findall('p:'+tag,NS):root.remove(e)
   new=f'ppt/replicaSlides/slide{idx}.xml';parts[new]=dump(root)
   rels=xml(z.read(relpath(old))) if relpath(old) in names else E.Element('{'+PK+'}Relationships')
   for e in list(rels):
    if e.get('TargetMode')=='External' or e.get('Type').endswith(('/notesSlide','/slide','/comments','/tags')):
     rels.remove(e);continue
    t=target(old,e.get('Target'));e.set('Target','/'+t);queue.append(t)
   parts[relpath(new)]=dump(rels)
   # Geometry/style preservation checks exclude text bodies only; no arbitrary coordinates are supplied.
   a=copy.deepcopy(before);b=copy.deepcopy(root)
   for r in [a,b]:
    for e in r.xpath('.//p:txBody | ./p:timing | ./p:transition | ./p:extLst',namespaces=NS):e.getparent().remove(e)
   geom=E.tostring(a)==E.tostring(b)
   if not geom:raise ValueError('Non-text geometry changed')
   measured=[]
   for sid,s in shapes.items():
    if sid not in actual:continue
    x=s.find('p:spPr/a:xfrm',NS)
    measured.append({'shape_id':sid,'source_text':actual[sid],'new_text':item['text'][sid],'xfrm':[{E.QName(v).localname:dict(v.attrib)} for v in x] if x is not None else None,'font_sizes':sorted(set(s.xpath('.//a:rPr/@sz',namespaces=NS))),'fonts':sorted(set(s.xpath('.//a:latin/@typeface | .//a:ea/@typeface',namespaces=NS)))})
   report['slides'].append({'output_page':idx,'source_page':item['source_page'],'source_part':old,'geometry_preserved':geom,'md_sections':item.get('md_sections'),'slots':measured})
  # Dependency closure: copy only assets reachable from selected slides.
  while queue:
   p=queue.pop()
   if p in parts:continue
   if Path(p).suffix.lower() in ['.mp4','.mov','.avi','.wmv','.wav','.mp3','.m4v']:continue
   if p not in names:raise ValueError('Missing relationship '+p)
   parts[p]=z.read(p)
   rp=relpath(p)
   if rp in names:
    rels=xml(z.read(rp))
    for e in list(rels):
     if e.get('TargetMode')=='External':continue
     t=target(p,e.get('Target'));queue.append(t)
    parts[rp]=dump(rels)
  # Static portable copy: media binaries omitted before reads.
  # Rebuild slide and presentation relationships; preserve only masters actually used.
  sidlst=pres.find('p:sldIdLst',NS)
  for e in list(sidlst):sidlst.remove(e)
  newrels=E.Element('{'+PK+'}Relationships',nsmap={None:PK})
  for i in range(1,len(spec['slides'])+1):
   rid='rIdReplica'+str(i)
   E.SubElement(sidlst,'{'+P+'}sldId',id=str(255+i),attrib={'{'+R+'}id':rid})
   E.SubElement(newrels,'{'+PK+'}Relationship',Id=rid,Type=R+'/slide',Target=f'replicaSlides/slide{i}.xml')
  ml=pres.find('p:sldMasterIdLst',NS)
  if ml is not None:
   for e in list(ml):
    rid=e.get('{'+R+'}id');rr=prmap[rid];p=target('ppt/presentation.xml',rr.get('Target'))
    if p not in parts:ml.remove(e)
    else:newrels.append(copy.deepcopy(rr))
  for tag in ['notesMasterIdLst','handoutMasterIdLst','custShowLst','extLst']:
   for e in pres.findall('p:'+tag,NS):pres.remove(e)
  parts['ppt/presentation.xml']=dump(pres);parts['ppt/_rels/presentation.xml.rels']=dump(newrels)
  rootrels=E.Element('{'+PK+'}Relationships',nsmap={None:PK})
  E.SubElement(rootrels,'{'+PK+'}Relationship',Id='rId1',Type=R+'/officeDocument',Target='ppt/presentation.xml')
  parts['_rels/.rels']=dump(rootrels)
  ct=xml(z.read('[Content_Types].xml'))
  for e in list(ct):
   if E.QName(e).localname=='Override' and e.get('PartName').lstrip('/') not in parts:ct.remove(e)
  for i in range(1,len(spec['slides'])+1):E.SubElement(ct,'{'+CT+'}Override',PartName=f'/ppt/replicaSlides/slide{i}.xml',ContentType='application/vnd.openxmlformats-officedocument.presentationml.slide+xml')
  parts['[Content_Types].xml']=dump(ct)
  # Remove missing relationships and media playback actions. Keep poster images and native geometry.
  removed=[];optimized=[]
  for pn in list(parts):
   if pn.endswith('.rels'):
    rr=xml(parts[pn]); owner=pn.replace('/_rels/','/')[:-5] if pn!='_rels/.rels' else ''
    bad=[]
    for e in list(rr):
     if e.get('TargetMode')=='External' or target(owner,e.get('Target')) not in parts:
      bad.append(e.get('Id'));removed.append({'part':pn,'type':e.get('Type'),'target':e.get('Target')});rr.remove(e)
    parts[pn]=dump(rr)
    if owner in parts and owner.endswith('.xml'):
     root=xml(parts[owner]);valid={e.get('Id') for e in rr}
     for e in list(root.iter()):
      if e.getparent() is None:continue
      refs=[v for k,v in e.attrib.items() if k.startswith('{'+R+'}')]
      if any(v not in valid for v in refs) or E.QName(e).localname in ['videoFile','audioFile','media'] or e.get('action')=='ppaction://media':e.getparent().remove(e)
     parts[owner]=dump(root)
  for pn,b in list(parts.items()):
   if pn.lower().endswith(('.png','.jpg','.jpeg')):
    try:
     im=Image.open(io.BytesIO(b));original=im.size;im.thumbnail((1920,1920),Image.Resampling.LANCZOS);buf=io.BytesIO()
     if pn.lower().endswith('.png'):im.save(buf,format='PNG',optimize=True)
     else:im.convert('RGB').save(buf,format='JPEG',quality=88,optimize=True)
     if len(buf.getvalue())<len(b):parts[pn]=buf.getvalue();optimized.append({'part':pn,'old_bytes':len(b),'new_bytes':len(parts[pn]),'old_px':original,'new_px':im.size})
    except Exception:pass
  report['static_removed_relationships']=removed;report['optimized_images']=optimized
  report['warnings'].append('Static pack: audio/video/animations/notes/external links omitted; poster retained; raster images limited to 1920 px and JPEG quality 88; original font declarations retained, font files not bundled.')
  with zipfile.ZipFile(out,'w',zipfile.ZIP_DEFLATED) as w:
   for p,data in parts.items():w.writestr(p,data)
 report['output_bytes']=out.stat().st_size;report['parts']=len(parts);report['explicit_font_fallback']=spec.get('font_fallback')
 out.with_suffix('.evidence.json').write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8')
 print(str(out))
if __name__=='__main__':main()
