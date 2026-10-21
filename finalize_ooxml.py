"""Finish native XLSX features unavailable in the authoring API.
Only stdlib ZIP/XML used. Cells/formulas/charts are authored by artifact-tool.
Restore original INDIRECT formulas while retaining verified starting caches.
"""
from pathlib import Path
import os, copy, json, re, zipfile, posixpath, xml.etree.ElementTree as E
BASE=Path(os.environ.get('PROTOTYPE_OUTPUT_DIR','/workspace/scratch/25bdfc598688/outputs/prototype-v02'))
path=BASE/'prototype.xlsx'
NS='http://schemas.openxmlformats.org/spreadsheetml/2006/main'
REL='http://schemas.openxmlformats.org/officeDocument/2006/relationships'
PKG='http://schemas.openxmlformats.org/package/2006/relationships'
CT='http://schemas.openxmlformats.org/package/2006/content-types'
E.register_namespace('',NS);E.register_namespace('r',REL)
def t(x):return '{'+NS+'}'+x
def xml(x):
    # OPC readers require the package roots in the default namespace.
    ns=x.tag.split('}')[0][1:] if x.tag.startswith('{') else NS
    E.register_namespace('',ns)
    result=E.tostring(x,encoding='utf-8',xml_declaration=True)
    E.register_namespace('',NS)
    return result
with zipfile.ZipFile(path) as z: data={n:z.read(n) for n in z.namelist()}
book=E.fromstring(data['xl/workbook.xml'])
sheets=book.find(t('sheets'))
names=[x.get('name') for x in sheets]
inputs=json.loads((BASE/'input_ranges.json').read_text())
styles=E.fromstring(data['xl/styles.xml']); xfs=styles.find(t('cellXfs')); unlock_map={}
for font in styles.findall('.//'+t('fonts')+'/'+t('font')):
    nm=font.find(t('name'))
    if nm is not None and nm.get('val')=='Noto Sans CJK JP':nm.set('val','Yu Gothic')
def unlocked(idx):
    if idx not in unlock_map:
        x=copy.deepcopy(xfs[idx]);x.set('applyProtection','1')
        pr=x.find(t('protection'))
        if pr is None: pr=E.SubElement(x,t('protection'))
        pr.set('locked','0');pr.set('hidden','0');unlock_map[idx]=len(xfs);xfs.append(x)
    return unlock_map[idx]
def xy(a):
    m=re.fullmatch(r'([A-Z]+)(\d+)',a); col=0
    for c in m[1]:col=col*26+ord(c)-64
    return col,int(m[2])
def inside(cell,area):
    start,*tail=area.split(':');end=tail[0] if tail else start
    x,y=xy(cell);a,b=xy(start);c,d=xy(end)
    return a<=x<=c and b<=y<=d
for i,name in enumerate(names,1):
    part=f'xl/worksheets/sheet{i}.xml';root=E.fromstring(data[part]);sd=root.find(t('sheetData'))
    for cell in root.findall('.//'+t('sheetData')+'/'+t('row')+'/'+t('c')):
        if cell.find(t('f')) is None and any(inside(cell.get('r'),a) for a in inputs.get(name,[])):
            cell.set('s',str(unlocked(int(cell.get('s','0')))))
    for old in root.findall(t('sheetProtection')):root.remove(old)
    for dv in root.findall('.//'+t('dataValidation')):
        dv.set('showErrorMessage','1');dv.set('errorStyle','stop');dv.set('errorTitle','入力値を確認');dv.set('error','指定された範囲・選択肢で入力してください。')
    sv=root.find(t('sheetViews'))
    if sv is not None:
        for v in sv:
            v.set('zoomScale','80');v.set('zoomScaleNormal','100')
            for pane in v.findall(t('pane')):v.remove(pane)
    if name=='計算':
        originals=json.loads((BASE/'indirect_formulas.json').read_text())
        for cell in root.findall('.//'+t('c')):
            if cell.get('r') in originals:cell.find(t('f')).text=originals[cell.get('r')].lstrip('=')
    if name=='見積書':
        for old in root.findall(t('pageSetup')):root.remove(old)
        margins=root.find(t('pageMargins'))
        if margins is not None:margins.attrib.update({'left':'0.25','right':'0.25','top':'0.35','bottom':'0.35','header':'0.15','footer':'0.15'})
        setup=E.Element(t('pageSetup'),{'paperSize':'9','orientation':'landscape','fitToWidth':'1','fitToHeight':'1'})
        root.insert(list(root).index(margins)+1,setup)
        pr=root.find(t('sheetPr'))
        if pr is None:pr=E.Element(t('sheetPr'));root.insert(0,pr)
        page=pr.find(t('pageSetUpPr'))
        if page is None:page=E.SubElement(pr,t('pageSetUpPr'))
        page.set('fitToPage','1')
    if i==1:
        # Legacy VML form-control scrollbar. No macro, ActiveX, or external link.
        existing_legacy=root.find(t('legacyDrawing'))
        existing_legacy_id=existing_legacy.get('{'+REL+'}id') if existing_legacy is not None else None
        if existing_legacy is None:
            legacy=E.Element(t('legacyDrawing'),{'{'+REL+'}id':'rIdProfitSlider'})
            root.append(legacy)
    data[part]=xml(root)
for xf in xfs:
    pr=xf.find(t('protection'))
    if pr is None:pr=E.SubElement(xf,t('protection'))
    xf.set('applyProtection','1');pr.set('locked','0');pr.set('hidden','0')
xfs.set('count',str(len(xfs)));data['xl/styles.xml']=xml(styles)
calc=book.find(t('calcPr'))
if calc is None:calc=E.SubElement(book,t('calcPr'))
calc.attrib.update({'calcMode':'auto','fullCalcOnLoad':'1','forceFullCalc':'1'})
dn=book.find(t('definedNames'))
if dn is None:dn=E.Element(t('definedNames'));book.insert(list(book).index(calc),dn)
for node in list(dn):
    if node.get('name')=='_xlnm.Print_Area':dn.remove(node)
E.SubElement(dn,t('definedName'),{'name':'_xlnm.Print_Area','localSheetId':'1'}).text="'見積書'!$A$1:$J$34"
data['xl/workbook.xml']=xml(book)
rels_part='xl/worksheets/_rels/sheet1.xml.rels';rels=E.fromstring(data[rels_part])
if not existing_legacy_id:E.SubElement(rels,'{'+PKG+'}Relationship',{'Id':'rIdProfitSlider','Type':REL+'/vmlDrawing','Target':'../drawings/profitSlider.vml'})
data[rels_part]=xml(rels)
vml='''<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">
<o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout>
<v:shapetype id="_x0000_t201" coordsize="21600,21600" o:spt="201" path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/><v:path shadowok="f" o:extrusionok="f" strokeok="f" fillok="f" o:connecttype="rect"/><o:lock v:ext="edit" shapetype="t"/></v:shapetype>
<v:shape id="_x0000_s1025" type="#_x0000_t201" style="position:absolute;margin-left:0;margin-top:234pt;width:365pt;height:18pt;z-index:1" o:button="t" fillcolor="buttonFace [67]" strokecolor="windowText [64]" o:insetmode="auto">
<v:fill color2="buttonFace [67]"/><o:lock v:ext="edit" rotation="t"/>
<x:ClientData ObjectType="Scroll"><x:MoveWithCells/><x:SizeWithCells/><x:Anchor>0, 2, 9, 2, 3, 0, 10, 0</x:Anchor><x:PrintObject>False</x:PrintObject><x:AutoFill>False</x:AutoFill><x:FmlaLink>'採算シミュレーション'!$B$9</x:FmlaLink><x:Val>0</x:Val><x:Min>0</x:Min><x:Max>100</x:Max><x:Inc>1</x:Inc><x:Page>5</x:Page><x:Horiz>True</x:Horiz><x:Dx>16</x:Dx></x:ClientData>
</v:shape></xml>'''
if existing_legacy_id:
    target=next(r.get('Target') for r in rels if r.get('Id')==existing_legacy_id)
    target=posixpath.normpath(posixpath.join('xl/worksheets',target)) if not target.startswith('/') else target[1:]
    vmroot=E.fromstring(data[target]);slider=E.fromstring(vml)
    for el in slider:
        if el.tag in ['{urn:schemas-microsoft-com:vml}shapetype','{urn:schemas-microsoft-com:vml}shape']:
            if el.tag.endswith('}shape'):el.set('id','_x0000_s9999')
            vmroot.append(el)
    data[target]=xml(vmroot)
else:data['xl/drawings/profitSlider.vml']=vml.encode()
# Legacy cell Notes stay collapsed (mouse hover / Show Note opens them).
for part in list(data):
    if part.endswith('.vml'):
        vr=E.fromstring(data[part])
        for shape in vr.findall('.//{urn:schemas-microsoft-com:vml}shape'):
            cd=shape.find('{urn:schemas-microsoft-com:office:excel}ClientData')
            if cd is not None and cd.get('ObjectType')=='Note':
                style=re.sub(r'visibility\s*:\s*[^;]+;?','',shape.get('style',''))
                shape.set('style',style+';visibility:hidden')
                for vis in cd.findall('{urn:schemas-microsoft-com:office:excel}Visible'):cd.remove(vis)
        data[part]=xml(vr)
types=E.fromstring(data['[Content_Types].xml'])
if not any(n.get('Extension')=='vml' for n in types):E.SubElement(types,'{'+CT+'}Default',{'Extension':'vml','ContentType':'application/vnd.openxmlformats-officedocument.vmlDrawing'})
data['[Content_Types].xml']=xml(types)
# Use the installed Windows Japanese font in the delivered file, retaining Noto for previews.
for part in list(data):
    if '/charts/' in part and part.endswith('.xml'):
        chartns='http://schemas.openxmlformats.org/drawingml/2006/chart'
        drawns='http://schemas.openxmlformats.org/drawingml/2006/main'
        chartroot=E.fromstring(data[part])
        donut=chartroot.find('.//{'+chartns+'}doughnutChart')
        if donut is not None:
            series=donut.find('{'+chartns+'}ser')
            # Per-slice palette is not exposed by the documented authoring API.
            for idx,color in enumerate(['07988A','A9CD55']):
                point=E.Element('{'+chartns+'}dPt')
                E.SubElement(point,'{'+chartns+'}idx',{'val':str(idx)})
                sp=E.SubElement(point,'{'+chartns+'}spPr')
                solid=E.SubElement(sp,'{'+drawns+'}solidFill')
                E.SubElement(solid,'{'+drawns+'}srgbClr',{'val':color})
                cats=series.find('{'+chartns+'}cat')
                series.insert(list(series).index(cats),point)
            data[part]=xml(chartroot)
        data[part]=data[part].replace(b'Noto Sans CJK JP',b'Yu Gothic')
with zipfile.ZipFile(path,'w',zipfile.ZIP_DEFLATED) as z:
    for n,b in data.items():z.writestr(n,b)
checks={}
checks['all_xml_well_formed']=all(E.fromstring(b) is not None for n,b in data.items() if n.endswith(('.xml','.rels','.vml')))
checks['no_vba']=not any('vbaProject' in n for n in data)
checks['no_external_links']=not any(n.startswith('xl/externalLinks') for n in data)
checks['indirect_formula_count']=sum((x.text or '').count('INDIRECT(') for x in E.fromstring(data['xl/worksheets/sheet8.xml']).iter(t('f')))
checks['no_optional_sheet_direct_references']=not any(re.search(r"'?原価\((木工所|農家)\)'?!",x.text or '') for n,b in data.items() if n.startswith('xl/worksheets/') and n.endswith('.xml') for x in E.fromstring(b).iter(t('f')))
checks['no_cached_formula_errors']=not any(x.get('t')=='e' for n,b in data.items() if n.startswith('xl/worksheets/') and n.endswith('.xml') for x in E.fromstring(b).iter(t('c')))
checks['scrollbar_link']='採算シミュレーション!B9 (0–100; increment 1)'
checks['native_excel_verified']=False
checks['notes_count']=sum(len(E.fromstring(b).findall('.//'+t('comment'))) for n,b in data.items() if 'comments' in n.lower() and n.endswith('.xml'))
checks['all_requested_notes_present']=checks['notes_count']==74
checks['no_sheet_protection_or_panes']=not any(E.fromstring(b).find('.//'+t(tag)) is not None for n,b in data.items() if n.startswith('xl/worksheets/sheet') and n.endswith('.xml') for tag in ['sheetProtection','pane'])
checks['all_cell_styles_unlocked']=all(x.find(t('protection')).get('locked')=='0' for x in xfs)
checks['single_legacy_drawing_per_sheet']=all(len(E.fromstring(b).findall(t('legacyDrawing')))<=1 for n,b in data.items() if n.startswith('xl/worksheets/sheet') and n.endswith('.xml'))
(BASE/'package_checks.json').write_text(json.dumps(checks,ensure_ascii=False,indent=2))
print(json.dumps(checks,ensure_ascii=False))
