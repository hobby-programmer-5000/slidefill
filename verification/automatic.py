"""In-process automatic guard. Never imports core, Tk, a model or a network client."""
import io
import os
from pathlib import Path
import tempfile
import hashlib
import posixpath
import xml.etree.ElementTree as ET
from zipfile import ZipFile,BadZipFile
from fractions import Fraction
from .rawxlsx import digest
from . import oracle,charts,recipes,sources
from .kernel import same_ints

A='{http://schemas.openxmlformats.org/drawingml/2006/main}'
P='{http://schemas.openxmlformats.org/presentationml/2006/main}'
R='{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'


class Failure(Exception):
    pass


def serialize(prs):
    stream=io.BytesIO();prs.save(stream);return stream.getvalue()


def cell_text(cell):
    return '\n'.join(''.join((n.text or '') if n.tag==A+'t' else '\v' if n.tag==A+'br' else '' for n in p.iter())
                     for p in cell.findall('.//'+A+'p'))


class Session:
    def __init__(self,source):
        self.source=source;self.bindings=[];self.books={};self.files={}
        for workbook_source in getattr(source,'sources',[source]):
            if hasattr(workbook_source,'path'):
                self.files[workbook_source.path]=getattr(workbook_source,'_origin_sha256',None)

    def capture(self,shape,kind,key,grid,recipe,source,item,mode='category',date_axis=False):
        if source is not None:
            self.files[source.path]=getattr(source,'_origin_sha256',None)
        self.bindings.append({'shape':shape.shape_id,'part':str(shape.part.partname).lstrip('/'),
                              'chart_part':str(shape.chart.part.partname).lstrip('/') if kind=='chart' else None,
                              'kind':kind,'key':key,'grid':[[tuple((raw,str(text or ''),fmt)) for raw,text,fmt in line] for line in grid],
                              'recipe':recipe,'source':source,'item':item,'mode':mode,'date_axis':date_axis})

    def unchanged(self):
        for path,wanted in self.files.items():
            if wanted is None or digest(path)!=wanted:raise Failure('source changed or could not be fingerprinted; reopen the input')

    def _book(self,source):
        if source is None:raise recipes.Unsupported('source is unavailable')
        if source.path not in self.books:self.books[source.path]=sources.read(source.path)
        return self.books[source.path]

    def _layout(self,binding,words):
        recipe=binding['recipe'];book=self._book(binding['source']);sheet=book[recipe['sheet']]
        rows=oracle.expected(sheet).get(recipe['view'])
        if rows is None:raise Failure('structured source schema no longer matches')
        expected=rows[recipe['start']:recipe['stop']]
        if expected!=recipe['rows']:raise Failure('structured source roles/coverage differ')
        grid=binding['grid'];lang=recipe['lang']
        if len(expected)+1!=len(grid):raise Failure('structured row count differs')
        if [cell[1] for cell in grid[0]]!=oracle.schema_headers(sheet,recipe['view'],lang):raise Failure('structured header meaning differs')
        for row,line in zip(expected,grid[1:]):
            if len(row)!=len(line):raise Failure('structured source column coverage differs')
            for token,actual in zip(row,line):
                wanted=oracle.expected_text(sheet,token,lang)
                if token['op']=='cell':
                    ref=token['refs'][0]
                    if wanted!='' and (sheet.types or {}).get(ref)=='n':
                        if not recipes.near(Fraction(wanted),recipes.figure(actual[0])):raise Failure('structured source value differs')
                    elif str(actual[0] if actual[0] is not None else '')!=wanted:raise Failure('structured source notation differs')
                elif actual[1]!=wanted:raise Failure('structured source label differs')
        for label in oracle.context_labels(sheet,recipe['view'],expected):
            if not any(oracle.clean(label) in oracle.clean(p) for p in words):raise Failure('source period/group label missing')
        if recipe['view'].startswith('outline-') and not oracle.grouping_caveat_visible(words):raise Failure('required grouping interpretation caveat missing')
        refs={ref for row in expected for token in row for ref in token['refs']}
        for finding in oracle.findings({recipe['sheet']:sheet}):
            if finding['cell'] in refs and not oracle.warning_visible(finding,words):raise Failure('required source warning missing or incorrect')

    def verify(self,blob):
        result={'state':'passed','checked':len(self.bindings),'source_checked':0,'reasons':[],'failures':[]}
        try:self.unchanged()
        except (Failure,OSError) as exc:
            result.update(state='blocked');result['failures'].append((None,str(exc)));return result
        try:
            with ZipFile(io.BytesIO(blob)) as z:
                for binding in self.bindings:
                    try:
                        root=ET.fromstring(z.read(binding['part']))
                        words=[''.join(t.text or '' for t in p.findall('.//'+A+'t')) for p in root.findall('.//'+A+'p')]
                        frames=[f for f in root.findall('.//'+P+'graphicFrame') if f.find('.//'+P+'cNvPr') is not None and int(f.find('.//'+P+'cNvPr').get('id'))==binding['shape']]
                        if len(frames)!=1:raise Failure('verified shape missing or duplicated')
                        frame=frames[0]
                        if binding['kind']=='chart':
                            node=frame.find('.//'+charts.C+'chart')
                            if node is None:raise Failure('bound chart shape changed')
                            folder=posixpath.dirname(binding['part'])
                            rels=ET.fromstring(z.read(folder+'/_rels/'+posixpath.basename(binding['part'])+'.rels'))
                            relation=next((n for n in rels if n.get('Id')==node.get(R+'id')),None)
                            if relation is None or relation.get('TargetMode')=='External':raise Failure('bound chart relationship changed')
                            target=relation.get('Target')
                            member=target.lstrip('/') if target.startswith('/') else posixpath.normpath(folder+'/'+target)
                            if member!=binding['chart_part']:raise Failure('bound chart part changed')
                            problems,partial=charts.check(z,binding['chart_part'],binding['grid'],binding['mode'],binding['date_axis'])
                            result['reasons'].extend(partial)
                            if problems:raise Failure('; '.join(dict.fromkeys(problems)))
                        else:
                            table=frame.find('.//'+A+'tbl')
                            actual=[[cell_text(cell) for cell in row.findall(A+'tc')] for row in table.findall(A+'tr')]
                            # Existing template header can be intentionally retained by core.
                            wanted=[[cell[1] for cell in row] for row in binding['grid']]
                            if binding.get('skip'):wanted=[binding['header']]+wanted[1:]
                            normalized=lambda text:text.replace('\v','\n').replace('\r\n','\n')
                            flat=lambda rows:[len(rows)]+[v for row in rows for v in [len(row),*[n for cell in row for text in [normalized(cell)] for n in [len(text),*map(ord,text)]]]]
                            if not same_ints(flat(actual),flat(wanted)):raise Failure('saved table values/count/order differ')
                        recipe=binding['recipe']
                        if isinstance(recipe,dict) and recipe.get('kind')=='layout':self._layout(binding,words)
                        else:
                            problems=recipes.audit(self._book(binding['source']),recipe,binding['grid'])
                            if problems:raise Failure('; '.join(problems[:4]))
                        result['source_checked']+=1
                    except recipes.Unsupported as exc:
                        result['reasons'].append(str(exc))
                    except (Failure,ValueError,TypeError,KeyError,IndexError,AttributeError,ET.ParseError,BadZipFile) as exc:
                        result['failures'].append((binding['item'],str(exc)))
        except (BadZipFile,OSError) as exc:result['failures'].append((None,str(exc)))
        if result['failures']:result['state']='blocked'
        elif result['reasons']:result['state']='partial'
        result['reasons']=list(dict.fromkeys(result['reasons']))
        return result

    def publish(self,blob,output):
        folder=os.path.dirname(os.path.abspath(output));fd,path=tempfile.mkstemp(prefix='.SlideFill-checked-',suffix='.pptx',dir=folder)
        try:
            with os.fdopen(fd,'wb') as f:f.write(blob)
            self._verify_candidate(path)
            if hashlib.sha256(Path(path).read_bytes()).digest()!=hashlib.sha256(blob).digest():raise Failure('staged report bytes changed after verification')
            self.unchanged()
            if os.name=='nt':os.rename(path,output)  # Windows refuses a destination that already exists.
            else:os.link(path,output);os.unlink(path)
        finally:
            if os.path.exists(path):os.unlink(path)
        return output

    def _verify_candidate(self,path):
        result=self.verify(Path(path).read_bytes())
        if result['state']=='blocked':raise Failure('; '.join(reason for _,reason in result['failures'][:3]))
