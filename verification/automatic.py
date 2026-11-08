"""In-process automatic guard. Never imports core, Tk, a model or a network client."""
import io
import os
from pathlib import Path
import tempfile
import hashlib
import datetime as dt
import posixpath
import xml.etree.ElementTree as ET
from zipfile import ZipFile,BadZipFile
from fractions import Fraction
from .rawxlsx import digest
from . import oracle,charts,recipes,sources,formulas,text_values
from .kernel import same_ints, round_rational

A='{http://schemas.openxmlformats.org/drawingml/2006/main}'
P='{http://schemas.openxmlformats.org/presentationml/2006/main}'
R='{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'


class Failure(Exception):
    pass


class Incomplete(Failure):
    pass


def serialize(prs):
    stream=io.BytesIO();prs.save(stream);return stream.getvalue()


def paragraph_text(paragraph):
    return ''.join((n.text or '') if n.tag==A+'t' else '\v' if n.tag==A+'br' else '' for n in paragraph.iter())


def cell_text(cell):
    return '\n'.join(paragraph_text(p) for p in cell.findall('.//'+A+'p'))


class Session:
    def __init__(self,source):
        self.source=source;self.bindings=[];self.books={};self.files={}
        self.paragraphs=[];self.text_inputs=[]
        for workbook_source in getattr(source,'sources',[source]):
            if hasattr(workbook_source,'path'):
                self.files[workbook_source.path]=getattr(workbook_source,'_origin_sha256',None)
        self.files.update(getattr(source,'document_fingerprints',{}))

    def capture_paragraph(self,part,index,text):
        self.paragraphs.append((str(part.partname).lstrip('/'),index,text))

    def capture_text_input(self,source,key,item,evidence,actual):
        self.files[source.path]=getattr(source,'_origin_sha256',None)
        self.text_inputs.append({'source':source,'key':key,'item':item,'actual':actual,**evidence})

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

    def _document(self,binding):
        recipe=binding['recipe'];rows=recipe['rows'];grid=binding['grid'];ko=recipe['lang']=='ko'
        if len(grid)!=len(rows)+1:raise Failure('document comparison row coverage differs')
        windows={row['window'] for row in rows};window=next(iter(windows)) if len(windows)==1 else 4
        head=('시기','항목','문서 기록',f'전 {window}개 평균',f'후 {window}개 평균','변화') if ko else ('When','Figure','In the document',f'Mean of {window} before',f'Mean of {window} after','Change')
        if len(windows)>1:
            head=('시기','항목','문서 기록','이전 평균 (기간 수)','이후 평균 (기간 수)','변화') if ko else ('When','Figure','In the document','Mean before (periods)','Mean after (periods)','Change')
        if [c[1] for c in grid[0]]!=list(head):raise Failure('document comparison header/window differs')
        document_words={}
        for row,actual in zip(rows,grid[1:]):
            book=self._book(row['source'])
            errors=recipes.audit(book,row['recipe'],row['grid'])
            if errors:raise Failure('; '.join(errors[:3]))
            points=[]
            for expressions in row['recipe'][1:]:
                date=recipes.evaluate(book,expressions[0],decoded=True)
                if isinstance(date,str):date=recipes.calendar_text(date)
                if isinstance(date,dt.date) and not isinstance(date,dt.datetime):date=dt.datetime.combine(date,dt.time())
                value=recipes.figure(recipes.evaluate(book,expressions[1],decoded=True))
                if not isinstance(date,dt.datetime) or value is None:
                    raise recipes.Unsupported('document comparison date/value cannot be independently reconstructed')
                points.append((date,value))
            if len(points)<3 or any(a[0]>=b[0] for a,b in zip(points,points[1:])):
                raise Failure('document comparison period order/coverage differs')
            gaps=sorted((b[0]-a[0]).days for a,b in zip(points,points[1:]))
            middle=len(gaps)//2
            gap=Fraction(gaps[middle]) if len(gaps)%2 else Fraction(gaps[middle-1]+gaps[middle],2)
            expected_window=7 if gap<=1 else 4 if gap<=8 else 12 if gap<=35 else 4 if gap<=100 else 3
            if row['window']!=expected_window:raise Failure('document comparison uses the wrong period window')
            event=row['event'];at=next((i for i,(date,_value) in enumerate(points) if (date.year>=event.year if gap>100 else date>=event)),None)
            if at is None or at<expected_window or at+expected_window>=len(points):
                raise Failure('document comparison has insufficient periods')
            before=points[at-expected_window:at];after=points[at+1:at+1+expected_window]
            for side in (before,after):
                if (side[-1][0]-side[0][0]).days>(expected_window-1)*gap*Fraction(3,2)+1:
                    raise Failure('document comparison hides a gap in the periods')
            left=sum((v for _d,v in before),Fraction())/expected_window
            right=sum((v for _d,v in after),Fraction())/expected_window
            suffix=f' ({expected_window})' if len(windows)>1 else ''
            change=f'{round_rational((right/left-1)*100,1):+.1f}%' if left>0 else f'{round_rational(left,0):,.0f} → {round_rational(right,0):,.0f}'
            month=('Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec')[event.month-1]
            when=f'{event.year}년 {event.month}월' if ko else f'{month} {event.year}'
            expected=[when,row['name'],row['text'],f'{round_rational(left,0):,.0f}'+suffix,f'{round_rational(right,0):,.0f}'+suffix,change]
            if [cell[1] for cell in actual]!=expected:raise Failure('document source means/change/labels differ')
            path=row['document']
            if path not in document_words:
                with ZipFile(path) as package:root=ET.fromstring(package.read('word/document.xml'))
                w='{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
                paragraphs=[''.join(n.text or '' for n in p.iter(w+'t')) for p in root.iter(w+'p')]
                document_words[path]=' '.join(' '.join(paragraphs).split())
            if ' '.join(row['text'].split()) not in document_words[path]:
                raise Failure('document comparison quotation is not present in the source')

    def verify(self,blob):
        result={'state':'passed','checked':len(self.bindings),'source_checked':0,'text_source_checked':0,'reasons':[],'failures':[]}
        formula_sources = {}
        for binding in self.bindings+self.text_inputs:
            recipe=binding.get('recipe')
            if isinstance(recipe,dict) and recipe.get('kind')=='document':
                for row in recipe['rows']:
                    entry=formula_sources.setdefault(row['source'].path,(row['source'],[]))
                    entry[1].append(row['recipe'])
                continue
            source = binding['source']
            if source is not None:
                entry = formula_sources.setdefault(source.path, (source, []))
                entry[1].append(binding.get('recipe'))
        try:self.unchanged()
        except (Failure,OSError) as exc:
            result.update(state='blocked');result['failures'].append((None,str(exc)));return result
        try:
            with ZipFile(io.BytesIO(blob)) as z:
                if len(z.namelist())!=len(set(z.namelist())):raise BadZipFile('duplicate output ZIP member')
                roots={}
                result['text_checked']=0
                for member,index,expected in self.paragraphs:
                    if member not in roots:roots[member]=ET.fromstring(z.read(member))
                    paragraphs=list(roots[member].iter(A+'p'))
                    if index>=len(paragraphs) or paragraph_text(paragraphs[index])!=expected:
                        result['failures'].append((None,'saved text value/order differs'))
                    else:result['text_checked']+=1
                for binding in self.text_inputs:
                    try:
                        if binding.get('unavailable'):raise recipes.Unsupported(binding['unavailable'])
                        if binding.get('expected') is not None and binding['actual']!=binding['expected']:
                            raise Failure('resolved text differs from its source value')
                        if isinstance(binding['recipe'],dict) and binding['recipe'].get('kind')=='layout':
                            self._layout(binding,[paragraph_text(p) for root in roots.values() for p in root.iter(A+'p')])
                        else:
                            problems=recipes.audit(self._book(binding['source']),binding['recipe'],binding['grid'])
                            if problems:raise Failure('; '.join(problems[:4]))
                        if binding.get('calculation') is not None:
                            expected=text_values.expected(self._book(binding['source']),binding)
                            if binding['actual']!=expected:raise Failure('calculated text differs from independent source arithmetic')
                        result['text_source_checked']+=1
                    except recipes.Unsupported as exc:result['reasons'].append(str(exc))
                    except (Failure,ValueError,TypeError,KeyError,IndexError) as exc:
                        result['failures'].append((binding['item'],str(exc)))
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
                        elif isinstance(recipe,dict) and recipe.get('kind')=='document':self._document(binding)
                        else:
                            problems=recipes.audit(self._book(binding['source']),recipe,binding['grid'])
                            if problems:raise Failure('; '.join(problems[:4]))
                        result['source_checked']+=1
                    except recipes.Unsupported as exc:
                        result['reasons'].append(str(exc))
                    except (Failure,ValueError,TypeError,KeyError,IndexError,AttributeError,ET.ParseError,BadZipFile) as exc:
                        result['failures'].append((binding['item'],str(exc)))
        except (BadZipFile,OSError,KeyError,ET.ParseError) as exc:result['failures'].append((None,str(exc)))
        result['formula_checked'] = 0
        for source, source_recipes in formula_sources.values():
            try:
                book=self._book(source);refs=set()
                for recipe in source_recipes:refs.update(formulas.references(recipe,book=book))
                evidence = formulas.inspect(book, refs)
                result['formula_checked'] += evidence['checked']
                result['reasons'].extend(evidence['partial'])
                result['failures'].extend((None, reason) for reason in evidence['failures'])
            except recipes.Unsupported as exc:
                result['reasons'].append(str(exc))
            except (ValueError,TypeError,KeyError,IndexError,OSError,BadZipFile,ET.ParseError) as exc:
                result['failures'].append((None,str(exc)))
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
        if result['state']=='partial':
            raise Incomplete('; '.join(result['reasons'][:3]))
