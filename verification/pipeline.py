"""Application adapter only. Independent check modules do not import this."""
from pathlib import Path
import json
import core
import suggest
import layouts
from .rawxlsx import digest
from .check import check


def certificate(source, plan):
    tables=[]
    for section in plan.sections:
        if section.table is None:continue
        marker=core.MARKER.fullmatch(section.table.typed)
        if marker is None or not marker.group(2).startswith('layout:'):
            raise ValueError('CPU source/output checker currently supports explicit structured layouts only')
        key=marker.group(2)
        ws,view,start,stop,lang,_=layouts.resolve(source,key)
        tables.append({'key':key,'sheet':ws.title,'view':view.id,'lang':lang,'start':start,'stop':stop,
                       'headers':[h[0 if lang=='ko' else 1] for h in view.headers],
                       'rows':[[t.certificate() for t in row] for row in view.rows[start:stop]]})
    return {'version':1,'workbook_sha256':digest(source.path),'tables':tables,'charts':[]}


def generate(path, output, lang='en'):
    folder=Path(output)
    if folder.exists():raise ValueError('Output directory already exists; choose a new directory')
    source=core.ExcelSource(str(path))
    plan=suggest.plan_report(source,lang)
    if plan is None:raise ValueError('No supported report layout')
    proof=certificate(source,plan)
    folder.mkdir(parents=True)
    template=folder/'template.pptx';report=folder/'report.pptx'
    suggest.make_report(str(template),suggest.title_of(str(path),plan.sections),plan.sections)
    result=core.run(str(template),str(path),str(report),source=source)
    (folder/'certificate.json').write_text(json.dumps(proof,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
    checked=check(path,report,proof)
    checked['application_errors']=result.count('error')
    if result.count('error'):checked['result']='FAIL';checked['errors'].append('application reported errors')
    checked['source_unchanged']=digest(path)==proof['workbook_sha256']
    if not checked['source_unchanged']:checked['result']='FAIL';checked['errors'].append('source changed during run')
    (folder/'checks.json').write_text(json.dumps(checked,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
    return checked
