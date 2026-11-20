import argparse,json,tempfile,subprocess,sys
from pathlib import Path
p=argparse.ArgumentParser();p.add_argument('--id',required=True);p.add_argument('--replace');p.add_argument('--out',required=True);p.add_argument('--font');a=p.parse_args()
root=Path(__file__).resolve().parents[1];catalog=json.loads((root/'references/catalog.json').read_text(encoding='utf8'));item=next((x for x in catalog if x['id']==a.id),None)
if item is None:raise ValueError('Unknown template ID')
texts=dict(item['text_slots']);changes=json.loads(Path(a.replace).read_text(encoding='utf-8-sig')) if a.replace else {}
if set(changes)-set(texts):raise ValueError('Unknown text slot IDs')
texts.update(changes);spec={'slides':[{'source_page':item['pack_page'],'text':texts,'md_sections':{'template_id':a.id,'original_page':item['original_page']}}]}
if a.font:spec['font_fallback']=a.font
with tempfile.TemporaryDirectory() as t:
 f=Path(t)/'spec.json';f.write_text(json.dumps(spec,ensure_ascii=False),encoding='utf8');cmd=[sys.executable,str(root/'scripts/native_template_transfer.py'),str(root/item['pack']),str(f),'--out',a.out]
 if not changes:cmd.append('--reference')
 subprocess.run(cmd,check=True)
