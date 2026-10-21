"""Windows + Excel + pywin32 development-only checks. Uses a disposable copy.
Run: python verify_v02_windows.py prototype.xlsx
"""
import sys, json, tempfile, shutil
from pathlib import Path
import win32com.client
source=Path(sys.argv[1]).resolve()
results=[]
def check(name,expected,actual):
    results.append(dict(name=name,expected=expected,actual=actual,pass_=expected==actual))
with tempfile.TemporaryDirectory() as td:
    dest=Path(td)/'test.xlsx';shutil.copy2(source,dest)
    app=win32com.client.DispatchEx('Excel.Application');app.Visible=False;app.DisplayAlerts=False;app.AutomationSecurity=3
    book=None
    try:
        book=app.Workbooks.Open(str(dest),UpdateLinks=0)
        check('notes count',74,sum(s.Comments.Count for s in book.Worksheets))
        for s in book.Worksheets:
            s.Activate()
            check(s.Name+' protection',False,bool(s.ProtectContents))
            check(s.Name+' freeze panes',False,bool(app.ActiveWindow.FreezePanes))
            for c in s.Comments:check(s.Name+' note hidden '+c.Parent.Address,False,bool(c.Visible))
        q=book.Worksheets('見積入力');a=book.Worksheets('売上分析')
        app.CalculateFullRebuild();check('D12 initial','入力OK',q.Range('D12').Value)
        green=q.Range('D12').DisplayFormat.Interior.Color
        q.Range('F16').Value=-1;app.CalculateFullRebuild()
        check('D12 invalid','要確認',q.Range('D12').Value)
        check('D12 color changed',True,green!=q.Range('D12').DisplayFormat.Interior.Color)
        q.Range('F16').Value=50000
        fixture=json.loads(source.with_name('sales_fixture.json').read_text(encoding='utf-8'))
        for industry,year,start,end in [('全業種',2026,1,12),('農家',2025,4,9)]:
            for cell,v in [('B5',industry),('E5',year),('H5',start),('K5',end)]:a.Range(cell).Value=v
            app.CalculateFullRebuild()
            rows=[r for r in fixture if r['year']==year and start<=r['month']<=end and (industry=='全業種' or r['industry']==industry)]
            check('filtered revenue '+industry,sum(r['sales'] for r in rows),a.Range('A9').Value)
        book.Save();book.Close(False);book=None
        book=app.Workbooks.Open(str(dest),UpdateLinks=0);app.CalculateFullRebuild()
        check('reopen notes',74,sum(s.Comments.Count for s in book.Worksheets))
    finally:
        if book is not None:book.Close(False)
        app.Quit()
Path('windows_v02_results.json').write_text(json.dumps(results,ensure_ascii=False,indent=2),encoding='utf-8')
print('PASS',sum(x['pass_'] for x in results),'/',len(results))
