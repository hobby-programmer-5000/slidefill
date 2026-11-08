"""Independent table tree and grids for bounded HTML exports named .xls."""
from html.parser import HTMLParser
from pathlib import Path
from .recipes import Unsupported
from .sources import MAX_CELLS,_text_table,_title


class Tables(HTMLParser):
    def __init__(self):
        super().__init__(convert_charrefs=True);self.tables=[];self.frames=[]
    def finish(self,row=False):
        if not self.frames:return
        node=self.frames[-1]
        if node['cell'] is not None:
            node['row'].append(node['cell']);node['cell']=None
        if row and node['row'] is not None:
            if node['row']:node['rows'].append(node['row'])
            node['row']=None
    def handle_starttag(self,tag,attrs):
        if tag=='table':
            if len(self.frames)>=32 or len(self.tables)>=256:raise Unsupported('HTML table tree exceeds budget')
            if self.frames:
                parent=self.frames[-1]['cell']
                if parent is None:raise Unsupported('nested HTML table outside its parent cell')
                parent[0].append(len(self.tables))
            node={'rows':[],'row':None,'cell':None};self.tables.append(node['rows']);self.frames.append(node)
        elif tag=='tr' and self.frames:self.finish(True);self.frames[-1]['row']=[]
        elif tag in ('td','th') and self.frames and self.frames[-1]['row'] is not None:
            self.finish()
            attrs=dict(attrs)
            def span(name):
                try:n=int(attrs.get(name,'1'))
                except ValueError:n=1
                if not 1<=n<=1000:raise Unsupported('HTML span exceeds independent reader contract')
                return n
            self.frames[-1]['cell']=([],span('rowspan'),span('colspan'))
    def handle_data(self,text):
        if self.frames and self.frames[-1]['cell'] is not None:self.frames[-1]['cell'][0].append(text)
    def handle_endtag(self,tag):
        if tag in ('td','th'):
            self.finish()
        elif tag=='tr':
            self.finish(True)
        elif tag=='table' and self.frames:self.finish(True);self.frames.pop()
    def close(self):
        super().close()
        if self.frames:raise Unsupported('unclosed HTML table tree')


def read_html(path,text):
    parser=Tables();parser.feed(text);parser.close();result={};names=[];budget=0
    stem=Path(path).stem
    for i in range(len(parser.tables)):names.append(_title(stem if i==0 else f'{stem} {i+1}',names))
    for name,table in zip(names,parser.tables):
        if not table:continue
        occupied=set();values={};merges=[];width={}
        for r,row in enumerate(table,1):
            c=1
            for parts,down,across in row:
                text=' '.join(''.join(f' [{names[p]}] ' if isinstance(p,int) else p for p in parts).split())
                while (r,c) in occupied:c+=1
                budget+=down*across
                if budget>MAX_CELLS:raise Unsupported('HTML grid exceeds independent cell budget')
                values[r,c]=text
                for rr in range(r,r+down):
                    width[rr]=max(width.get(rr,0),c+across-1)
                    for cc in range(c,c+across):
                        if (rr,cc) in occupied:raise Unsupported('overlapping HTML spans')
                        occupied.add((rr,cc))
                if down>1 or across>1:merges.append((r,c,r+down-1,c+across-1))
                c+=across
        covered_rows={r for top,left,bottom,right in merges for r in range(top,bottom+1)}
        keep=[r for r in range(1,len(table)+1) if any(values.get((r,c),'') for c in range(1,width.get(r,0)+1)) or r in covered_rows]
        moved={old:new for new,old in enumerate(keep,1)}
        grid=[[values.get((r,c),'') for c in range(1,width.get(r,0)+1)] for r in keep]
        adjusted=[(moved[top],left,moved.get(bottom,moved[top]),right) for top,left,bottom,right in merges if top in moved]
        result.update(_text_table(grid,name,adjusted))
    if not result:raise Unsupported('HTML worksheet tables missing')
    return result
