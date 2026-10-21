import fs from 'node:fs/promises';
import {Workbook,SpreadsheetFile} from '@oai/artifact-tool';
import {enhance,verifyAnalysis} from './enhance.mjs';
const out=process.env.PROTOTYPE_OUTPUT_DIR || '/workspace/scratch/25bdfc598688/outputs/prototype-v02';
await fs.mkdir(out,{recursive:true});
const w=Workbook.create();
const names=['採算シミュレーション','見積書','見積入力','原価(木工所)','原価(農家)','設定','サンプル・入力例','計算','使い方・検証'];
const sh=Object.fromEntries(names.map(n=>[n,w.worksheets.add(n)]));
const navy='#164D45',teal='#07988A',blue='#E5F3EE',ink='#233E39',gray='#F1F7F4',red='#FBE9E7';
const inputs={};
function val(s,a,v){sh[s].getRange(a).values=Array.isArray(v)?v:[[v]];}
function f(s,a,v){sh[s].getRange(a).formulas=[[v]];}
function merge(s,a,t){sh[s].mergeCells(a);val(s,a.split(':')[0],t);}
function input(s,a){sh[s].getRange(a).format.fill=blue;sh[s].getRange(a).format.font.color='#1354A0';(inputs[s]??=[]).push(a);}
function header(s,a,values){val(s,a, [values]);sh[s].getRange(a).format={fill:navy,font:{color:'#FFFFFF',bold:true},rowHeight:30,wrapText:true};}
function note(s,row,text,end='J'){merge(s,`A${row}:${end}${row}`,text);sh[s].getRange(`A${row}:${end}${row}`).format.font={color:'#536477',size:10};sh[s].getRange(`A${row}:${end}${row}`).format.wrapText=true;sh[s].getRange(`A${row}:${end}${row}`).format.rowHeight=32;}
function dv(s,a,values){sh[s].getRange(a).dataValidation={rule:{type:'list',values}};}
function numeric(s,a,min,max){sh[s].getRange(a).dataValidation={rule:{type:'decimal',operator:'between',formula1:min,formula2:max}};}
for(const n of names){const s=sh[n];s.showGridLines=false;s.getRange('A1:K65').format={font:{name:'Noto Sans CJK JP',size:11,color:ink},rowHeight:24,verticalAlignment:'center'};s.getRange('A1:K65').format.columnWidth=15;merge(n,'A2:J2',n==='採算シミュレーション'?'見積・原価・利益シミュレーター':n);s.getRange('A2:J2').format.font={size:17,bold:true,color:navy};s.getRange('A2:J2').format.rowHeight=38;s.tabColor=['採算シミュレーション','見積書'].includes(n)?navy:teal;}
// Registry is deliberately text, not a direct reference to optional industry sheets.
header('設定','A5:B5',['業種','原価シート名']);val('設定','A6:B7',[['木工所','原価(木工所)'],['農家','原価(農家)']]);input('設定','A6:B7');
w.names.add('IndustryList',"'設定'!$A$6:$A$7");
val('設定','A11:A14',[['自社名'],['住所'],['連絡先'],['担当者']]);val('設定','B11:B14',[['サンプル事業者'],['住所を入力してください'],['連絡先を入力してください'],['担当者名']]);input('設定','B11:F14');
note('設定',17,'業種を減らす場合：残す業種を選択 → この表の不要行を行ごと削除 → 不要な原価シートを削除。');
note('設定',19,'表の行削除はシート保護を解除して実施（パスワードなし）。原価シートの名前と出力領域は変更しないでください。');
// Common sales input: no industry-specific assumptions.
const q='見積入力';
val(q,'A5:A12',[['件名'],['宛先'],['見積番号'],['確定値引率'],['発行日'],['有効期限'],['納期・支払条件'],['備考']]);
val(q,'B5:B12',[['家具製作・納品'],['サンプルのお客様'],['DEMO-001'],[0],[new Date('2026-10-20T00:00:00Z')],['発行日から30日'],['納期・支払条件を入力'],['試作品の架空データです']]);input(q,'B5:F12');sh[q].getRange('B8').setNumberFormat('0.0%');sh[q].getRange('B9').setNumberFormat('yyyy/mm/dd');numeric(q,'B8',0,1);
header(q,'A15:I15',['明細ID','品名','規格・仕様','数量','単位','税抜単価','税率','税抜金額','入力確認']);
const woodSales=[['M001','テーブル','標準仕様',2,'台',50000,.10],['M002','椅子','標準仕様',4,'脚',15000,.10]];
const farmSales=[['M001','野菜5kg箱','同一作物・同一品質',100,'箱',2000,.08]];
function setSales(rows){sh[q].getRange('A16:G25').clear({applyTo:'contents'});if(rows.length)val(q,`A16:G${15+rows.length}`,rows);}
setSales(woodSales);input(q,'A16:G25');numeric(q,'D16:D25',.000001,1e9);numeric(q,'F16:F25',0,1e12);dv(q,'G16:G25',['0','0.08','0.1']);
for(let r=16;r<=25;r++){
f(q,`I${r}`,`=IF(COUNTA(A${r}:G${r})=0,"",IF(OR(A${r}="",B${r}="",E${r}=""),"必須欄不足",IF(COUNTIFS($A$16:$A$25,A${r})<>1,"ID重複",IF(OR(NOT(ISNUMBER(D${r})),NOT(ISNUMBER(F${r})),NOT(ISNUMBER(G${r}))),"数値を入力",IF(OR(D${r}<=0,F${r}<0,AND(G${r}<>0,G${r}<>0.08,G${r}<>0.1)),"数量・価格・税率確認","OK")))))`);
f(q,`H${r}`,`=IF(I${r}="OK",ROUND(D${r}*F${r},0),"")`);
}
f(q,'I28','=IF(COUNTIFS(I16:I25,"OK")=0,"見積明細なし",IF(SUMPRODUCT((I16:I25<>"")*(I16:I25<>"OK"))>0,"見積明細を確認",IF(OR(NOT(ISNUMBER(B8)),B8<0,B8>1),"確定値引率を確認","OK")))');
val(q,'G28','入力状態');f(q,'H29','=IF(I28="OK",SUM(H16:H25),"")');val(q,'G29','税抜小計');
note(q,31,'明細IDは半角英数字で固定。原価も同じIDに紐付けます。数量変更で一式費用は増減しません。');
note(q,33,'税率は取引に応じて0・8%・10%を選択。値引き後明細金額を円単位四捨五入、税額は税率別に切り捨て。');
sh[q].getRange('F16:H29').setNumberFormat('#,##0');sh[q].getRange('G16:G25').setNumberFormat('0%');
// Optional cost sheets have a shared output contract in L:N, all own formulas are local.
const woodCosts=[['M001','材料','テーブル材料',1,15000,'台当たり','数量連動'],['M001','作業','テーブル作業',10,1000,'時間/台','数量連動'],['M002','材料','椅子材料',1,5000,'脚当たり','数量連動'],['M002','作業','椅子作業',2.5,1000,'時間/脚','数量連動'],['案件共通','配送','一括配送',1,8000,'納品1回','一式']];
const farmCosts=[['M001','生産原価','対象出荷分の生産原価',5,null,'kg/箱・単価自動','数量連動'],['M001','材料','箱・包装材',1,150,'箱当たり','数量連動'],['M001','作業','選別・梱包',.1,1500,'時間/箱','数量連動'],['案件共通','配送','今回の配送',1,10000,'出荷1回','一式']];
for(const [n,rows] of [['原価(木工所)',woodCosts],['原価(農家)',farmCosts]]){
const s=sh[n];val(n,'A4','選択中の業種');f(n,'B4',"='採算シミュレーション'!B5"); merge(n,'D4:J4',''); f(n,'D4',`=IF(B4="${n==='原価(木工所)'?'木工所':'農家'}","この原価を採算計算に使用中","この原価は現在の採算計算に使われていません")`);
val(n,'A11','共通経費の扱い');val(n,'B11','計上済み');input(n,'B11');dv(n,'B11',['計上済み','単価に含む','今回は未算入']);
if(n==='原価(農家)'){
val(n,'A5:A9',[['作物・栽培期間'],['生産費用合計'],['販売可能収穫量'],['生産量の単位'],['生産原価単価']]);val(n,'B5','サンプル野菜・1作期');f(n,'B6','=SUM(F5:F8)');val(n,'B7',5000);val(n,'B8','kg');f(n,'B9','=IF(AND(COUNT(F5:F8)=4,MIN(F5:F8)>=0,ISNUMBER(B7),B7>0),B6/B7,"")');input(n,'B5');input(n,'B7:B8');numeric(n,'B7',.000001,1e12);
val(n,'E5:F8',[['種苗・肥料・資材',300000],['栽培・収穫作業',400000],['その他の直接費',100000],['共通経費の割当',200000]]);input(n,'F5:F8');numeric(n,'F5:F8',0,1e12);
}else {note(n,6,'材料・加工・外注・配送を、明細IDまたは「案件共通」に紐付けます。');note(n,8,'数量連動：使用量/販売単位 × 原価単価 × 見積数量。一式：使用量 × 原価単価。');}
note(n,13,'青色が入力欄。未算入の共通経費は利益に含まれません。既に単価に含む費用は二重計上しないでください。');
header(n,'A15:J15',['対象明細ID','区分','内容','使用量・時間','原価単価','原価金額','根拠・単位','入力確認','数量の扱い','見積数量']);
input(n,'A16:E55');input(n,'G16:G55');input(n,'I16:I55');numeric(n,'D16:E55',0,1e12);dv(n,'I16:I55',['数量連動','一式']);dv(n,'B16:B55',['材料','作業','外注','配送','共通経費','生産原価','その他']);
for(let i=0;i<rows.length;i++){let [id,cat,desc,qty,price,note,mode]=rows[i];let r=16+i;val(n,`A${r}:E${r}`,[[id,cat,desc,qty,price]]);val(n,`G${r}`,note);val(n,`I${r}`,mode);}
val(n,'K15','適用単価'); if(n==='原価(木工所)')val(n,'B11','今回は未算入');
for(let r=16;r<=55;r++){
f(n,`K${r}`,n==='原価(農家)'?`=IF(A${r}="","",IF(B${r}="生産原価",$B$9,E${r}))`:`=IF(A${r}="","",E${r})`);
f(n,`J${r}`,`=IF(A${r}="","",IF(I${r}="一式",1,IFERROR(INDEX('見積入力'!$D$16:$D$25,MATCH(A${r},'見積入力'!$A$16:$A$25,0)),"ID未登録")))`);
f(n,`H${r}`,`=IF(COUNTA(A${r}:E${r})+COUNTA(G${r}:I${r})-COUNTA(H${r})=0,"",IF(OR(A${r}="",B${r}="",C${r}="",I${r}=""),"必須欄不足",IF(AND(A${r}<>"案件共通",COUNTIFS('見積入力'!$A$16:$A$25,A${r})<>1),"ID未登録・重複",IF(OR(AND(I${r}<>"数量連動",I${r}<>"一式"),AND(A${r}="案件共通",I${r}<>"一式")),"数量の扱い確認",IF(OR(NOT(ISNUMBER(D${r})),NOT(ISNUMBER(K${r})),AND(B${r}<>"生産原価",E${r}=""),D${r}="",NOT(ISNUMBER(J${r}))),"数値を入力",IF(OR(D${r}<0,K${r}<0,J${r}<=0),"負数・数量確認","OK"))))))`);
// Avoid self-reference in the blank-row test.
f(n,`H${r}`,sh[n].getRange(`H${r}`).formulas[0][0].replace(`COUNTA(G${r}:I${r})-COUNTA(H${r})`,`COUNTA(G${r})+COUNTA(I${r})`));
f(n,`F${r}`,`=IF(H${r}="OK",ROUND(D${r}*K${r}*J${r},0),"")`);
}
val(n,'L4','COST-V1');val(n,'L5','原価入力状態');
let extra=n==='原価(農家)'?',IF(OR(COUNT(F5:F8)<>4,MIN(F5:F8)<0,NOT(ISNUMBER(B7)),B7<=0,B8=""),"生産費用・収穫量確認","OK")':',"OK"';
f(n,'M5',`=IF(COUNTIFS(H16:H55,"OK")=0,"原価明細なし",IF(SUMPRODUCT((H16:H55<>"")*(H16:H55<>"OK"))>0,"原価明細を確認",IF(AND(B11<>"計上済み",B11<>"単価に含む",B11<>"今回は未算入"),"共通経費の扱い確認"${extra})))`);
val(n,'L6','案件共通費');f(n,'M6','=SUMIFS(F16:F55,A16:A55,"案件共通")');val(n,'L7','共通経費表示');f(n,'M7','=B11');
header(n,'L15:N15',['明細ID','紐付く原価','原価行数']);
for(let r=16;r<=25;r++){f(n,`L${r}`,`=IF('見積入力'!A${r}="","",'見積入力'!A${r})`);f(n,`M${r}`,`=IF(L${r}="","",SUMIFS($F$16:$F$55,$A$16:$A$55,L${r}))`);f(n,`N${r}`,`=IF(L${r}="","",COUNTIFS($A$16:$A$55,L${r}))`);}
s.getRange('L4:N25').format={fill:gray,font:{size:10,color:'#536477'},columnWidth:18};s.getRange('D16:F55').setNumberFormat('#,##0.00');s.getRange('F16:F55').setNumberFormat('#,##0');s.getRange('B6:B9').setNumberFormat('#,##0.00');s.freezePanes.freezeRows(15);
}
// Main user controls.
const d='採算シミュレーション';
val(d,'A5','対象業種');val(d,'B5','木工所');input(d,'B5');sh[d].getRange('B5').dataValidation={rule:{type:'list',formula1:'IndustryList'}};
val(d,'A7','目標利益率');val(d,'B7',.30);input(d,'B7');numeric(d,'B7',0,.9999);sh[d].getRange('B7').setNumberFormat('0.0%');
val(d,'A9','試算値引率(%)');val(d,'B9',0);input(d,'B9');sh[d].getRange('B9').dataValidation={rule:{type:'whole',operator:'between',formula1:0,formula2:100}};
val(d,'A11','原価増減率');val(d,'B11',0);input(d,'B11');numeric(d,'B11',-1,10);sh[d].getRange('B11').setNumberFormat('0.0%');
note(d,13,'原価増減率は原価全体への一律試算です。個別条件・数量変更は見積入力と業種別原価シートで行います。');
note(d,15,'試算値引率は見積書に反映されません。採用する率を「見積入力」の確定値引率へ入力してください。');
// Calculation owns the business model and standardized import.
const c='計算';val(c,'A4:A12',[['参照シート名'],['出力形式'],['原価入力状態'],['見積入力状態'],['試算条件'],['案件共通費'],['総合状態'],['共通経費'],['税抜小計']]);
f(c,'B4',`=IFERROR(INDEX('設定'!$B$6:$B$7,MATCH('${d}'!B5,'設定'!$A$6:$A$7,0)),"業種未登録")`);
const ind=(addr)=>`INDIRECT("'"&$B$4&"'!${addr}")`;
f(c,'B5',`=IFERROR(${ind('L4')},"原価シートなし")`);f(c,'B6',`=IFERROR(${ind('M5')},"原価シートなし")`);f(c,'B7',`='見積入力'!I28`);
f(c,'B8',`=IF(OR(NOT(ISNUMBER('${d}'!B7)),NOT(ISNUMBER('${d}'!B9)),NOT(ISNUMBER('${d}'!B11))),"試算条件は数値",IF(OR('${d}'!B7<0,'${d}'!B7>=1,'${d}'!B9<0,'${d}'!B9>100,MOD('${d}'!B9,1)<>0,'${d}'!B11< -1,'${d}'!B11>10),"試算条件の範囲確認","OK"))`);
f(c,'B9',`=IFERROR(${ind('M6')},"取得不可")`);f(c,'B11',`=IFERROR(${ind('M7')},"取得不可")`);f(c,'B12',`=IF(B7="OK",'見積入力'!H29,"")`);
header(c,'A15:J15',['明細ID','品名','基準売上','原価','原価件数','試算売上','試算原価','確定売上','税率','紐付け']);
for(let r=16;r<=25;r++){
f(c,`A${r}`,`=IF('見積入力'!A${r}="","",'見積入力'!A${r})`);f(c,`B${r}`,`=IF(A${r}="","",'見積入力'!B${r})`);f(c,`C${r}`,`=IF(A${r}="","",'見積入力'!H${r})`);
f(c,`D${r}`,`=IF(A${r}="","",IFERROR(${ind('M'+r)},"取得不可"))`);f(c,`E${r}`,`=IF(A${r}="","",IFERROR(${ind('N'+r)},"取得不可"))`);
f(c,`J${r}`,`=IF(A${r}="","",IF(AND(ISNUMBER(D${r}),ISNUMBER(E${r}),E${r}>0),"OK","原価の紐付け不足"))`);
f(c,`F${r}`,`=IF(OR(A${r}="",$B$10<>"OK"),"",ROUND(C${r}*(1-'${d}'!$B$9/100),0))`);f(c,`G${r}`,`=IF(OR(A${r}="",$B$10<>"OK"),"",ROUND(D${r}*(1+'${d}'!$B$11),0))`);
f(c,`H${r}`,`=IF(OR(A${r}="",$B$10<>"OK"),"",ROUND(C${r}*(1-'見積入力'!$B$8),0))`);f(c,`I${r}`,`=IF(A${r}="","",'見積入力'!G${r})`);
}
f(c,'B10','=IF(B5<>"COST-V1",B5,IF(B6<>"OK",B6,IF(B7<>"OK",B7,IF(B8<>"OK",B8,IF(COUNTIFS(J16:J25,"原価の紐付け不足")>0,"商品に原価未登録","OK")))))');
val(c,'A29:A41',[['基準原価'],['試算売上'],['試算原価'],['予定案件利益'],['予定案件利益率'],['目標必要売価'],['目標維持値引上限'],['確定売上'],['10%税額'],['8%税額'],['税込見積額'],['基準利益'],['試算共通費']]);
const cf={B29:'SUM(D16:D25)+B9',B30:'SUM(F16:F25)',B31:'SUM(G16:G25)+B41',B32:'B30-B31',B33:'IF(B30=0,"算出不可",B32/B30)',B34:`ROUNDUP(B31/(1-'${d}'!B7),0)`,B35:'IF(B12=0,"算出不可",IF(B34>B12,"値上げが必要",1-B34/B12))',B36:'SUM(H16:H25)',B37:'ROUNDDOWN(SUMIFS(H16:H25,I16:I25,0.1)*0.1,0)',B38:'ROUNDDOWN(SUMIFS(H16:H25,I16:I25,0.08)*0.08,0)',B39:'B36+B37+B38',B40:'B12-B29',B41:`ROUND(B9*(1+'${d}'!B11),0)`};
for(const [cell,ex] of Object.entries(cf))f(c,cell,`=IF(B10="OK",${ex},"")`);
sh[c].getRange('B29:B41').setNumberFormat('#,##0');sh[c].getRange('B33').setNumberFormat('0.0%');sh[c].getRange('B35').setNumberFormat('0.0%');
// KPI cards and actionable status, all linked to calculation.
for(const [area,label,formula,fmt] of [['D5:F5','試算売上',"='計算'!B30",'#,##0" 円"'],['G5:J5','試算原価',"='計算'!B31",'#,##0" 円"'],['D9:F9','予定案件利益',"='計算'!B32",'#,##0" 円"'],['G9:J9','予定案件利益率',"='計算'!B33",'0.0%']]){
merge(d,area,label);const row=Number(area.match(/\d+/)[0])+1,cols=area.split(':').map(x=>x.replace(/\d+/,''));const a=`${cols[0]}${row}:${cols[1]}${row+1}`;merge(d,a,'');f(d,a.split(':')[0],formula);sh[d].getRange(a).format={fill:gray,font:{size:20,bold:true,color:teal},numberFormat:fmt};}
merge(d,'D12:J12','');f(d,'D12',`=IF('計算'!B10<>"OK",'計算'!B10,IF('計算'!B30=0,"売価0円・利益率算出不可",IF('計算'!B32<0,"赤字",IF('計算'!B33<B7,"黒字・目標未達","目標達成"))))`);
val(d,'A17','目標必要売価');f(d,'B17',"='計算'!B34");sh[d].getRange('B17').setNumberFormat('#,##0');val(d,'D17','目標維持値引上限');f(d,'F17',"='計算'!B35");sh[d].getRange('F17').setNumberFormat('0.0%');val(d,'G17','共通経費');merge(d,'H17:J17','');f(d,'H17',"='計算'!B11");
header(d,'A35:F35',['明細ID','品名','試算売上','紐付く原価','共通費控除前差額','同利益率']);
for(let r=36;r<=45;r++){let x=r-20;f(d,`A${r}`,`='計算'!A${x}`);f(d,`B${r}`,`='計算'!B${x}`);f(d,`C${r}`,`='計算'!F${x}`);f(d,`D${r}`,`='計算'!G${x}`);f(d,`E${r}`,`=IF(OR(A${r}="",'計算'!B10<>"OK"),"",C${r}-D${r})`);f(d,`F${r}`,`=IF(OR(A${r}="",'計算'!B10<>"OK"),"",IF(C${r}=0,"算出不可",E${r}/C${r}))`);}
val(d,'A47','案件共通費');f(d,'D47',"='計算'!B41");note(d,49,'商品別の差額は案件共通費の控除前です。案件全体の予定利益は共通費を差し引いた結果です。');
sh[d].getRange('C36:E47').setNumberFormat('#,##0');sh[d].getRange('F36:F45').setNumberFormat('0.0%');
header(c,'L29:N29',['項目','基準','試算']);val(c,'L30:L32',[['売上'],['原価'],['利益']]);for(const [r,b,t] of [[30,'B12','B30'],[31,'B29','B31'],[32,'B40','B32']]){f(c,`M${r}`,`=${b}`);f(c,`N${r}`,`=${t}`);}
header(c,'L35:M35',['値引率','利益(円)']);
for(let i=0;i<=10;i++){let r=36+i;val(c,`L${r}`,`${i*5}%`);f(c,`M${r}`,`=IF(B10="OK",B12*(1-${i*.05})-B31,"")`);}
const ch=sh[d].charts.add('bar',sh[c].getRange('L29:N32'));ch.title='基準と試算の比較（円）';ch.setPosition('A19','F33');ch.yAxis={numberFormatCode:'#,##0',numberFormatSourceLinked:false};ch.series.items[0].fill='#9DAEC0';ch.series.items[1].fill=teal;
const line=sh[d].charts.add('line',sh[c].getRange('L35:M46'));line.title='値引率と利益（概算・円）';line.setPosition('G19','K33');line.hasLegend=false;line.series.items[0].line={fill:teal,style:'solid',width:2};line.yAxis={numberFormatCode:'#,##0',numberFormatSourceLinked:false};
for(const chart of [ch,line]){chart.titleTextStyle.typeface='Noto Sans CJK JP';chart.titleTextStyle.fontSize=13;}
// Customer-facing quote. No cost/profit links.
const z='見積書';merge(z,'A4:F4','');f(z,'A4',`='見積入力'!B6&" 御中"`);merge(z,'G4:J4','');f(z,'G4',`="見積番号："&'見積入力'!B7`);merge(z,'G5:J5','');f(z,'G5',`='設定'!B11`);merge(z,'G6:J6','');f(z,'G6',`='設定'!B12`);merge(z,'G7:J7','');f(z,'G7',`='設定'!B13`);merge(z,'A6:F6','');f(z,'A6',`="件名："&'見積入力'!B5`);
val(z,'A8','税込見積金額');merge(z,'C8:F9','');f(z,'C8',`=IF('計算'!B10="OK",'計算'!B39,"入力を確認")`);sh[z].getRange('C8:F9').format={font:{size:20,bold:true,color:teal},fill:gray,numberFormat:'#,##0" 円"'};
merge(z,'G9:J9','');f(z,'G9',`="有効期限："&'見積入力'!B10`);val(z,'A11','発行日');f(z,'B11',"='見積入力'!B9");sh[z].getRange('B11').setNumberFormat('yyyy/mm/dd');
header(z,'A13:J13',['No.','品名','規格・仕様','数量','単位','税抜単価','値引前金額','値引額','税抜金額','税率']);
for(let r=14;r<=23;r++){let x=r+2;f(z,`A${r}`,`=IF('計算'!A${x}="","",${r-13})`);for(const [col,src] of [['B','B'],['C','C'],['D','D'],['E','E'],['F','F']])f(z,`${col}${r}`,`=IF(A${r}="","",'見積入力'!${src}${x})`);f(z,`G${r}`,`=IF(A${r}="","",'見積入力'!H${x})`);f(z,`I${r}`,`=IF(A${r}="","",'計算'!H${x})`);f(z,`H${r}`,`=IF(OR(A${r}="",'計算'!B10<>"OK"),"",G${r}-I${r})`);f(z,`J${r}`,`=IF(A${r}="","",'見積入力'!G${x})`);}
val(z,'G25:G30',[['税抜合計'],['10%対象額'],['10%消費税'],['8%対象額'],['8%消費税'],['税込合計']]);
for(const [r,ex] of [[25,"'計算'!B36"],[26,"SUMIFS('計算'!H16:H25,'計算'!I16:I25,0.1)"],[27,"'計算'!B37"],[28,"SUMIFS('計算'!H16:H25,'計算'!I16:I25,0.08)"],[29,"'計算'!B38"],[30,"'計算'!B39"]]){merge(z,`I${r}:J${r}`,'');f(z,`I${r}`,`=IF('計算'!B10="OK",${ex},"")`);}
merge(z,'A25:E26','');f(z,'A25',`="納期・支払条件："&'見積入力'!B11`);merge(z,'A28:E30','');f(z,'A28',`="備考："&'見積入力'!B12`);sh[z].getRange('A25:E30').format.wrapText=true;
merge(z,'A32:J32','');f(z,'A32',`=IF('計算'!B10<>"OK","発行不可："&'計算'!B10,IF(OR('見積入力'!B5="",'見積入力'!B6="",'見積入力'!B7="",'設定'!B11=""),"宛先・件名・番号・自社名を確認",""))`);
note(z,34,'8%対象・10%対象の税額はそれぞれ円未満切捨て。0%入力分は税額0円として集計。');sh[z].getRange('F14:I30').setNumberFormat('#,##0');sh[z].getRange('J14:J23').setNumberFormat('0%');
// Permanent examples, separate from live inputs.
const e='サンプル・入力例';note(e,4,'架空の記入例です。相場ではありません。このシートの数値は実際の見積に集計されません。');
header(e,'A6:G6',['木工所の販売例','品名','仕様','数量','単位','税抜単価','税率']);val(e,'A7:G8',woodSales);header(e,'A11:G11',['対象ID','区分','内容','使用量','原価単価','根拠','数量の扱い']);val(e,'A12:G16',woodCosts);note(e,18,'木工所：売上160,000円、原価88,000円、利益72,000円、利益率45.0%。');
header(e,'A21:G21',['農家の販売例','品名','仕様','数量','単位','税抜単価','税率']);val(e,'A22:G22',farmSales);header(e,'A25:G25',['対象ID','区分','内容','使用量','原価単価','根拠','数量の扱い']);val(e,'A26:G29',farmCosts.map(x=>x.map(v=>v===null?200:v)));note(e,31,'農家：生産費用1,000,000円 ÷ 販売可能収穫量5,000kg ＝ 200円/kg。');note(e,33,'100箱×5kg。売上200,000円、原価140,000円、利益60,000円、利益率30.0%。');note(e,35,'農家を試す：業種を農家に変更し、見積入力A16:G25を消去してA22:G22の値をA16:G16に貼り付け。');note(e,37,'原価側は両業種とも記入例を入力済み。実案件では入力欄を置き換え、不要行は入力セルだけ消去してください。');
const g='使い方・検証';const guide=[
'v0.1 操作評価用プロトタイプ。Windows版Excel上の動作・印刷・スライダーは未検証です。',
'1. 採算シミュレーションで業種を選び、見積入力に販売明細を入力します。',
'2. 対象の原価シートに原価を入力します。商品ごとに少なくとも1行の原価が必要です。',
'3. 値引率・原価増減率を変更し、利益を確認します。見積書は確定値引率だけを使用します。',
'4. 見積書の宛先・発行日・税区分を確認し、Excelの印刷／PDF保存で出力します。',
'青色セルが入力欄、灰色・白色の数式欄は計算結果。シート保護はパスワードなしで解除可能です。',
'上限：販売10明細、原価40明細/業種。行追加・削除ではなく、用意した入力欄を使用してください。',
'案件共通費は商品へ自動配賦しません。商品別差額は案件共通費控除前、案件利益は控除後です。',
'数量連動は一定の使用量・単価を仮定します。数量変更後は歩留まり・段取り・配送条件を再確認してください。',
'農業原価は1作物・1栽培単位の平均原価。等級別原価・複数作物・仕掛品・在庫管理は対象外です。',
'生産費用には栽培・収穫まで、出荷費用には選別・梱包・配送を計上する例です。重複計上に注意。',
'予定案件利益は入力費用に基づく管理用試算。会計上の売上総利益・営業利益とは一致しない場合があります。',
'共通経費を含めない場合は「今回は未算入」を選択します。原価0円は実際に0円と判断したときだけ入力。',
'業種削除：残す業種を選択 → 設定の保護解除 → 不要業種の行を削除 → 不要原価シートを削除。',
'共通シート・計算シートは削除不可。原価シートL:Nは共通出力領域で編集不可。名前変更も対象外です。',
'見積書には原価・利益は出しません。ただしExcel本体には原価があるため、顧客へ渡すのはPDFです。',
'試算値引率は整数0～100%。目標利益率は0%以上100%未満。原価増減率は-100%～1000%。',
'商品別原価・売価は円単位四捨五入。税額は税率ごとの値引き後合計に対して円未満切捨て。',
'見積書は10%・8%・0%の入力に対応。税区分の適用判断・適格請求書の発行機能は対象外です。',
'下記互換性表は仕様の確認範囲。最古の動作保証バージョンは実機検証後に確定します。'];
guide.forEach((t,i)=>note(g,4+i*2,t));
header(g,'A46:D46',['機能','仕様確認範囲','実機確認','最古の保証版']);
val(g,'A47:D51',[['IF/SUMIFS/INDEX/MATCH等','2016以降の公式資料を参照','未実施','未確定'],['INDIRECTによる切替','2016以降の公式資料を参照','未実施','未確定'],['標準グラフ・入力規則','対象Excelごとに要確認','未実施','未確定'],['スクロールバー','VML形式で組込予定','未実施','未確定'],['印刷・PDF保存','対象Excelごとに要確認','未実施','未確定']]);
note(g,54,'参考：https://support.microsoft.com/en-us/excel/functions/indirect-function');note(g,56,'参考：https://support.microsoft.com/en-us/excel/add-a-scroll-bar-or-spin-button-to-a-worksheet');
// Final coherent styling and live error highlights.
for(const n of names){let s=sh[n];s.getRange('A1:A65').format.columnWidth=22;s.getRange('B1:B65').format.columnWidth=21;s.getRange('C1:C65').format.columnWidth=24;s.getRange('D1:J65').format.columnWidth=16;}
for(const n of [q,'原価(木工所)','原価(農家)']){const col=n===q?'I':'H',last=n===q?25:55;sh[n].getRange(`${col}16:${col}${last}`).conditionalFormats.addCustom(`AND(${col}16<>"",${col}16<>"OK")`,{fill:red,font:{color:'#9D2828',bold:true}});}
sh[d].getRange('D10:F11').conditionalFormats.add('cellIs',{operator:'lessThan',formula:0,format:{fill:red,font:{color:'#9D2828'}}});
sh[q].freezePanes.freezeRows(15);sh[d].getRange('B17').format.font={bold:true,color:teal};
const indirectFormulas={}; for(const cell of ["B5","B6","B9","B11",...Array.from({length:10},(_,i)=>["D"+(16+i),"E"+(16+i)]).flat()]) indirectFormulas[cell]=sh[c].getRange(cell).formulas[0][0];
function recalc(){ const source=sh[d].getRange("B5").values[0][0]==="農家"?"原価(農家)":"原価(木工所)"; for(const [cell,expr] of Object.entries(indirectFormulas)){const address=expr.match(/!([A-Z]+[0-9]+)/)[1]; const direct=expr.replace(/INDIRECT\([^)]*\)/,String.fromCharCode(39)+source+String.fromCharCode(39)+"!"+address); f(c,cell,direct.startsWith("=")?direct:"="+direct);} w.recalculate();}
await fs.writeFile(`${out}/indirect_formulas.json`,JSON.stringify(indirectFormulas,null,2));
// Layout and safeguards added after the shared structure.
for(let r=16;r<=25;r++){
 const old=sh[q].getRange(`I${r}`).formulas[0][0];
 f(q,`I${r}`,`=IF(OR(A${r}="案件共通",ISNUMBER(SEARCH("~*",A${r})),ISNUMBER(SEARCH("~?",A${r})),ISNUMBER(SEARCH("~~",A${r}))),"IDに予約語・特殊文字不可",${old.replace(/^=/,'')})`);
}
for(const n of ['原価(木工所)','原価(農家)']){
 sh[n].getRange('G16:G55').format.wrapText=true;sh[n].getRange('H16:H55').format.wrapText=true;
 sh[n].getRange('A16:J55').format.rowHeight=30;
 sh[n].getRange('K15:K55').format.columnWidth=15;sh[n].getRange('K16:K55').setNumberFormat('#,##0.00');
}
sh['原価(農家)'].getRange('F5:F8').setNumberFormat('#,##0');
sh['原価(農家)'].getRange('E5:E8').format.wrapText=true;sh['原価(農家)'].getRange('E5:F8').format.rowHeight=33;
for(let r=5;r<=12;r++) if(r!==8&&r!==9)sh[q].mergeCells(`B${r}:F${r}`);
for(let r=11;r<=14;r++)sh['設定'].mergeCells(`B${r}:F${r}`);
for(let r=25;r<=30;r++)sh[z].mergeCells(`G${r}:H${r}`);
sh[z].getRange('A1:A34').format.columnWidth=10;sh[z].getRange('B1:C34').format.columnWidth=24;
sh[z].getRange('D1:E34').format.columnWidth=10;sh[z].getRange('F1:I34').format.columnWidth=16;sh[z].getRange('J1:J34').format.columnWidth=9;
sh[z].getRange('A13:J13').format.rowHeight=35;sh[z].getRange('B14:C23').format.wrapText=true;sh[z].getRange('A14:J23').format.rowHeight=29;
sh[z].mergeCells('A8:B9');sh[z].getRange('A8:B9').format.wrapText=true;
sh[q].getRange('A16:I25').format.rowHeight=29;sh[q].getRange('I16:I25').format.wrapText=true;
sh[e].getRange('F7:F8').setNumberFormat('#,##0');sh[e].getRange('G7:G8').setNumberFormat('0%');sh[e].getRange('F22').setNumberFormat('#,##0');sh[e].getRange('G22').setNumberFormat('0%');
val(g,'B50','VML組込・操作は未検証');
// Engine checks, mutating then restoring the actual workbook.
await enhance({w,sh,names,val,f,merge,header,input,note,dv,numeric,out});
const tests=[];
const read=(s,a)=>sh[s].getRange(a).values[0][0];
function test(name,expected,actual,tol=0.000001){let pass=typeof expected==='number'?Math.abs(actual-expected)<tol:actual===expected;tests.push({name,expected,actual,pass});if(!pass)console.log('TEST_FAIL',name,expected,actual);}
recalc();
test('木工 売上',160000,read(c,'B30'));test('木工 原価',88000,read(c,'B31'));test('木工 利益',72000,read(c,'B32'));test('木工 利益率',.45,read(c,'B33'));
val(d,'B9',10);recalc();test('木工 10%値引利益',56000,read(c,'B32'));test('試算変更で見積書は不変',176000,read(c,'B39'));val(d,'B9',0);
val(d,'B11',.1);recalc();test('原価10%増',96800,read(c,'B31'));val(d,'B11',0);
val(q,'A17','M001');recalc();test('重複IDを遮断','原価明細を確認',read(c,'B10'));val(q,'A17','M002');
val(d,'B9',100);recalc();test('売価0円の率','算出不可',read(c,'B33'));val(d,'B9',0);
val(d,'B7',1);recalc();test('目標100%を遮断','試算条件の範囲確認',read(c,'B10'));val(d,'B7',.3);
val(d,'B5','農家');setSales(farmSales);recalc();test('農家 生産単価',200,read('原価(農家)','B9'));test('農家 原価',140000,read(c,'B31'));test('農家 利益',60000,read(c,'B32'));test('農家 税込',216000,read(c,'B39'));
val('原価(農家)','B7',4000);recalc();test('収穫量減少 原価',165000,read(c,'B31'));val('原価(農家)','B7',5000);
val('原価(農家)','B7',0);recalc();test('収穫量0時は利益非表示','',read(c,'B32'));val('原価(農家)','B7',5000);
val(q,'B8',.1);recalc();test('確定10%値引税込',194400,read(c,'B39'));val(q,'B8',0);
val(d,'B5','木工所');setSales(woodSales);recalc();
val(q,'D16',3);recalc();test('数量変更で一式配送は固定',113000,read(c,'B31'));val(q,'D16',2);
val('原価(木工所)','E16',-1);recalc();test('負の原価単価を遮断','',read(c,'B32'));val('原価(木工所)','E16',15000);
val('原価(木工所)','A16','M999');recalc();test('未登録原価IDを遮断','',read(c,'B32'));val('原価(木工所)','A16','M001');
val(q,'F16','');recalc();test('売価欠損を遮断','',read(c,'B32'));val(q,'F16',50000);
val(q,'A16','*');recalc();test('ワイルドカードIDを遮断','',read(c,'B32'));val(q,'A16','M001');
val(q,'G17',.08);recalc();test('税率混在の税額',174800,read(c,'B39'));val(q,'G17',.1);
val(d,'B11',-1);recalc();test('原価100%減の試算',0,read(c,'B31'));val(d,'B11',0);recalc();
await verifyAnalysis({w,sh,val,recalc,test});
await fs.writeFile(`${out}/engine_tests.json`,JSON.stringify(tests,null,2));
await fs.writeFile(`${out}/input_ranges.json`,JSON.stringify(inputs,null,2));
console.log((await w.inspect({kind:'table',range:"'計算'!A29:B41",include:'values,formulas',tableMaxRows:13,tableMaxCols:2,maxChars:2500})).ndjson);
console.log((await w.inspect({kind:'match',searchTerm:'#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!|#NULL!',options:{useRegex:true,maxResults:15},summary:'formula error scan',maxChars:1500})).ndjson);
const file=await SpreadsheetFile.exportXlsx(w);await file.save(`${out}/prototype.xlsx`);
for(const [n,range] of [['売上分析','A1:N38'],['売上データ','A1:I15']]){const b=await w.render({sheetName:n,range,scale:1.3,format:'png'});await fs.writeFile(`${out}/${n}.png`,new Uint8Array(await b.arrayBuffer()));}
for(const [n,range] of [[d,'A1:K18'],[z,'A1:J34'],[q,'A1:J33'],['原価(木工所)','A1:J23'],['原価(農家)','A1:J23'],['設定','A1:J20'],[e,'A1:J37'],[c,'A1:J41'],[g,'A1:J24'],[d,'A19:K33']]){try{const b=await w.render({sheetName:n,range,scale:1.3,format:'png'});await fs.writeFile(`${out}/preview-${names.indexOf(n)}-${range.includes('19')?'charts':'sheet'}.png`,new Uint8Array(await b.arrayBuffer()));}catch(err){console.log('RENDER_ERROR',n,String(err));}}
console.log('DONE',tests.filter(t=>t.pass).length,tests.length);
