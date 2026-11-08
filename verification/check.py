"""Read-only verification entry point, safe to use independently of the app."""
from zipfile import BadZipFile
import xml.etree.ElementTree as ET
from .rawxlsx import workbook, digest
from .oracle import coverage, findings
from .output import tables
from . import formulas


def check(workbook_path, report_path, certificate):
    if not isinstance(certificate,dict) or not isinstance(certificate.get('tables'),list) or not isinstance(certificate.get('charts',[]),list):
        return {'result':'FAIL','errors':['invalid certificate structure'],'source_findings':[],'tables':0,'charts':0}
    errors=[];source_findings=[];formula_checked=0
    try:
        if certificate.get('version')!=1:raise ValueError('unknown certificate version')
        if certificate['workbook_sha256']!=digest(workbook_path):errors.append('source SHA256 differs')
        book=workbook(workbook_path)
        refs=formulas.references(certificate,book=book)
        # Coverage discovers schemas across every source sheet, including cells
        # outside declared table rows. Their caches can change that discovery.
        for name,sheet in book.items():
            refs.update((name,ref) for ref in sheet.formulas or {})
            refs.update((name,span.split(':',1)[0].replace('$','')) for span in sheet.formula_regions or ())
        audit=formulas.inspect(book,refs)
        formula_checked=audit['checked']
        errors.extend(audit['failures'])
        errors.extend('source formula cannot be independently checked: '+reason for reason in audit['partial'])
        source_findings=findings(book)
        errors.extend(coverage(book,certificate))
        if not certificate['tables']:errors.append('no verified structured tables')
        if certificate.get('charts'):errors.append('source-linked chart certificates are outside this checker contract')
        errors.extend(tables(report_path,book,certificate))
    except (ValueError,KeyError,TypeError,IndexError,AttributeError,OSError,BadZipFile,ET.ParseError) as exc:
        errors.append(f'unsupported or invalid verification evidence: {exc}')
    return {'result':'FAIL' if errors else 'PASS','errors':errors,
            'source_findings':source_findings,
            'scope':'Declared bounded rate/organization schemas, exact source values, required source warnings and saved tables. All formula caches in the parsed workbook must pass the bounded independent formula checker because schema coverage scans every sheet. Generic charts and other schemas are outside this report checker. Not a whole-app or business-intent proof.',
            'formula_checked':formula_checked,
            'tables':len(certificate.get('tables',[])),'charts':len(certificate.get('charts',[]))}
