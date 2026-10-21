"""Optional developer verification. Requires Windows Excel + pywin32.
Never changes the source workbook. Tests temporary copies in a dedicated Excel instance.
Run: python verify_excel_windows.py prototype.xlsx
"""
import json, pathlib, platform, shutil, sys, tempfile, time
import win32com.client
source=pathlib.Path(sys.argv[1]).resolve()
assert source.is_file()
results=[]; excel=None; books=[]
def result(name,expected,actual):
    passed=abs(actual-expected)<1e-6 if isinstance(expected,(float,int)) and isinstance(actual,(float,int)) else expected==actual
    results.append(dict(name=name,expected=expected,actual=actual,passed=passed))
def calc():
    excel.CalculateFullRebuild()
    limit=time.monotonic()+30
    while excel.CalculationState!=0:
        if time.monotonic()>limit:raise TimeoutError('Excel calculation timed out')
        time.sleep(.1)
def open_copy(folder,name):
    p=pathlib.Path(folder)/name;shutil.copy2(source,p)
    b=excel.Workbooks.Open(str(p),UpdateLinks=0,ReadOnly=False);books.append(b)
    return b,p
try:
    excel=win32com.client.DispatchEx('Excel.Application')
    excel.Visible=False;excel.DisplayAlerts=False;excel.AutomationSecurity=3
    environment={'os':platform.platform(),'excel_version':excel.Version,'excel_build':str(excel.Build)}
    with tempfile.TemporaryDirectory(prefix='estimate_excel_test_') as tmp:
        b,p=open_copy(tmp,'normal.xlsx');calc();sim=b.Worksheets('採算シミュレーション');c=b.Worksheets('計算')
        result('wood status','OK',c.Range('B10').Value2);result('wood profit',72000,c.Range('B32').Value2)
        sim.Range('B9').Value2=10;calc();result('10% discount profit',56000,c.Range('B32').Value2)
        result('quote unaffected by simulation',176000,c.Range('B39').Value2)
        bars=sim.ScrollBars();result('native scrollbar exists',True,bars.Count>=1)
        if bars.Count:
            bars.Item(1).Value=20;calc();result('scrollbar linked cell',20,sim.Range('B9').Value2);result('scrollbar changes profit',40000,c.Range('B32').Value2)
        sim.Range('B9').Value2=0;b.Save();b.Close(False);books.remove(b)
        b=excel.Workbooks.Open(str(p),UpdateLinks=0);books.append(b);calc();result('reopen profit',72000,b.Worksheets('計算').Range('B32').Value2)
        b.Close(False);books.remove(b)
        b,p=open_copy(tmp,'wood_only.xlsx')
        b.Worksheets('設定').Unprotect();b.Worksheets('設定').Rows(7).Delete();b.Worksheets('原価(農家)').Delete();calc()
        result('delete farm sheet',72000,b.Worksheets('計算').Range('B32').Value2)
        b.Save();b.Close(False);books.remove(b)
        b=excel.Workbooks.Open(str(p),UpdateLinks=0);books.append(b);calc();result('wood-only reopen',72000,b.Worksheets('計算').Range('B32').Value2)
        b.Close(False);books.remove(b)
        b,p=open_copy(tmp,'farm_only.xlsx');b.Worksheets('採算シミュレーション').Range('B5').Value2='農家'
        q=b.Worksheets('見積入力');q.Range('A16:G25').ClearContents();q.Range('A16:G16').Value2=(('M001','野菜5kg箱','同一作物・同一品質',100,'箱',2000,.08),)
        b.Worksheets('設定').Unprotect();b.Worksheets('設定').Rows(6).Delete();b.Worksheets('原価(木工所)').Delete();calc()
        result('delete wood sheet',60000,b.Worksheets('計算').Range('B32').Value2)
        result('farm quote',216000,b.Worksheets('計算').Range('B39').Value2)
        b.Save();b.Close(False);books.remove(b)
        b=excel.Workbooks.Open(str(p),UpdateLinks=0);books.append(b);calc();result('farm-only reopen',60000,b.Worksheets('計算').Range('B32').Value2)
except Exception as exc:
    results.append({'name':'execution error','passed':False,'error':repr(exc)})
finally:
    for b in books:
        try:b.Close(False)
        except Exception:pass
    if excel is not None:excel.Quit()
    report={'environment':locals().get('environment',{}),'results':results,'manual_checks':['印刷/PDFの文字切れ','スクロールバーをマウスで操作','シート削除後のプルダウン候補','長い品名・住所の表示']}
    dest=source.parent/'windows_test_results.json'
    dest.write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8')
    print(dest)
