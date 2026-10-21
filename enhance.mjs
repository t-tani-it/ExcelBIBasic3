import fs from 'node:fs/promises';
const products=['テーブル','椅子','収納棚','木製小物','野菜5kg箱','トマト箱','米5kg','季節の果物'];
const source=[];
export async function enhance({w,sh,names,val,f,merge,header,input,note,dv,numeric,out}){
  // User-authorized broad visual redesign; preserve existing calculation addresses.
  for(const n of names){sh[n].freezePanes.unfreeze();sh[n].getRange('A2:J2').format={fill:'#164D45',font:{color:'#FFFFFF',bold:true,size:17},rowHeight:42};}
  for(const a of ['D6:F7','G6:J7','D10:F11','G10:J11'])sh['採算シミュレーション'].getRange(a).format.fill='#E3F1EB';
  sh['採算シミュレーション'].charts.items[0].series.items[0].fill='#B2CF69';
  sh['見積入力'].unmergeCells('B12:F12');sh['見積入力'].mergeCells('B12:C12');
  sh['見積入力'].getRange('B12:C12').format.wrapText=true;
  merge('見積入力','D12:F12','');f('見積入力','D12','=IF(I28="OK","入力OK",IF(I28="見積明細なし","未入力","要確認"))');
  sh['見積入力'].getRange('A12:F12').format.rowHeight=42;
  for(const [v,bg,fg] of [['入力OK','#DFF0E5','#165440'],['未入力','#F0F3F2','#5E6E68'],['要確認','#FBE3DE','#9B3127']])sh['見積入力'].getRange('D12:F12').conditionalFormats.addCustom(`$D$12="${v}"`,{fill:bg,font:{color:fg,bold:true}});
  val('設定','A19','業種を減らす場合は不要な登録行を行ごと削除。シート保護・数式ロックは解除済みです。');
  val('使い方・検証','A4','v0.2 デザイン・売上分析評価用。Windows版Excelでの操作・保存・印刷は未検証です。');
  val('使い方・検証','A14','淡い緑色が入力欄。全シートの固定・保護・数式ロックを解除済み。数式の上書きにご注意ください。');
  val('使い方・検証','A30','業種削除：残す業種を選択 → 設定の不要業種行を削除 → 不要原価シートを削除。');
  val('使い方・検証','A32','共通シート・計算・原価L:Nは計算に必要です。ロックはありませんが、数式を変更しないでください。');
  // Each requested note has a definition, numerical/example case and practical purpose.
  const notes=[];
  const add=(n,a,definition,example,usage)=>{let text=`【意味】${definition}\n【例】${example}\n【実務】${usage}`;notes.push({sheet:n,cell:a,text});w.notes.add({id:`${n}:${a}`,target:{cell:{sheetName:n,sheetId:sh[n].sheetId,address:a}},authorId:'',createdAt:'',body:{plainText:text}});};
  const sales=[
    ['販売明細を識別する固定ID。原価の対象明細IDと一致させます。','M001をテーブルに設定。','行番号ではなくIDで原価を結び付けます。重複・案件共通・*?~は使用不可。'],
    ['顧客へ販売する商品・役務の名称。','テーブル、野菜5kg箱。','見積書に表示されるため顧客が理解できる名称にします。'],
    ['サイズ・材質・品質等の見積条件。','幅120cm・無垢材、野菜5kg。','価格の前提を明記し、仕様変更時の追加費用判断に使います。'],
    ['今回販売する数量。0より大きい値。','テーブル2台、野菜100箱。','数量連動の原価も変わります。一式原価は変わりません。'],
    ['販売数量の単位。','台、脚、箱、kg、式。','原価の使用量と単位が一致しているか確認します。'],
    ['確定値引き前の税抜販売単価。','50,000円×2台＝100,000円。','交渉前の提示価格を入力します。原価単価と混同しません。'],
    ['当該取引に適用する消費税率。','0.1は10%、0.08は8%。','適用区分を確認して選択。税額は値引き後の税率別合計から計算します。'],
    ['数量×税抜単価を円単位四捨五入した値引き前金額。','2×50,000＝100,000円。','明細の計算確認に使います。確定値引き後金額は見積書を確認します。'],
    ['必須項目・ID重複・数量・単価・税率の検査結果。','ID重複、数値を入力、OK。','要確認行を修正してから見積を発行します。OKは採算の良さを意味しません。']
  ];
  const costs=[
    ['原価を結び付ける販売明細ID。共通費は案件共通。','M001はテーブル用、一括配送は案件共通。','共通費は案件利益で一度だけ控除し商品へ自動配賦しません。'],
    ['原価の種類。','材料、作業、外注、配送、生産原価。','費用の漏れ・重複を確認。農家の生産原価は上部の計算単価を使います。'],
    ['費用の具体的な内容。','テーブル材料、選別・梱包。','見積根拠を後から説明できる名称を残します。'],
    ['数量連動は販売1単位あたり、一式は案件合計の使用量・時間。','1台10時間、1箱5kg、一括配送1回。','単位を根拠欄に書き、販売数量を二重に掛けないようにします。'],
    ['材料1単位・作業1時間等あたりの税抜費用。','時給換算1,000円。農家の生産原価行は空欄で自動単価を利用。','仕入単価や負担作業費を入力。税込売価と比較しません。'],
    ['使用量×適用単価×見積数量。一式は数量1。円単位四捨五入。','10時間×1,000円×2台＝20,000円。','商品別原価と案件共通費の積上げを確認します。'],
    ['使用量・原価単価の単位や見積根拠。','時間/台、kg/箱、仕入先見積。','将来の単価改定や担当者変更時に再現できるよう記録します。'],
    ['ID・必須欄・数値・数量の扱いの入力検査結果。','ID未登録・重複、OK。','OK以外の行を修正。未入力を原価0円として扱わない設計です。'],
    ['数量連動は販売数量に比例。一式は案件全体で固定。','加工は数量連動、一括配送は一式。','数量変更の試算時に固定費用が誤って増えないよう選択します。'],
    ['数量連動なら対象販売明細から取得、一式なら1。','M001の販売数量2、一式配送は1。','参照結果の確認用。変更は見積入力側で行います。']
  ];
  sales.forEach((x,i)=>add('見積入力',`${String.fromCharCode(65+i)}15`,...x));
  for(const n of ['原価(木工所)','原価(農家)'])costs.forEach((x,i)=>add(n,`${String.fromCharCode(65+i)}15`,...x));
  for(const r of [6,21])sales.slice(0,7).forEach((x,i)=>add('サンプル・入力例',`${String.fromCharCode(65+i)}${r}`,...x));
  for(const r of [11,25])[0,1,2,3,4,6,8].forEach((k,i)=>add('サンプル・入力例',`${String.fromCharCode(65+i)}${r}`,...costs[k]));
  const kpis=[
    ['A7','売上を分母にした目標の予定案件利益率。0%以上100%未満。','原価88,000円、目標30%なら必要売価は125,715円。','最低受注価格の検討に使います。原価への上乗せ率とは異なります。'],
    ['A9','基準売上に対する仮の値引率。整数0〜100を入力。','10なら10%引き。160,000円→144,000円。','利益への影響を比較。見積書へは別途「確定値引率」に反映します。'],
    ['A11','商品原価と案件共通費全体への一律増減割合。','10%で88,000円→96,800円。','仕入価格上昇への耐性を見る概算。個別変更は原価明細で行います。'],
    ['D5','各明細の基準売上に試算値引率を適用した税抜合計。','160,000円から10%引き＝144,000円。','値引き交渉後に残る売上を比較します。消費税は含みません。'],
    ['G5','選択業種の原価と案件共通費に原価増減率を適用した合計。','原価88,000円から10%増＝96,800円。','未算入の間接費がないか確認。木工と農業の原価は合算しません。'],
    ['D9','試算売上−試算原価。入力済みの費用に基づく管理用利益。','160,000−88,000＝72,000円。','受注判断に使います。未計上費用があると実利益を過大評価します。会計上の営業利益とは限りません。'],
    ['G9','予定案件利益÷試算売上。売上0円なら算出不可。','72,000÷160,000＝45%。','案件を比較し目標と照合します。商品別差額率とは共通費の範囲が異なります。']
  ];
  kpis.forEach(([a,...x])=>add('採算シミュレーション',a,...x));
  const calc=[sales[0],sales[1],['値引き前の税抜売上。','テーブル2台×50,000円＝100,000円。','入力側の数量・単価との一致を確認します。'],['選択業種で当該IDに紐付く原価。共通費は別集計。','テーブル材料30,000＋作業20,000＝50,000円。','明細別採算を確認。案件共通費を重ねて入力しません。'],['対象IDに紐付いた原価行の件数。','材料と作業の2行なら2。','0件は原価登録漏れを示します。金額0円と区別します。'],kpis[3].slice(1),['明細原価に原価増減率を適用した金額。','50,000円×1.1＝55,000円。','案件共通費は別途B41で増減し合計します。'],['基準売上に確定値引率を適用した税抜金額。','100,000円を10%引きなら90,000円。','実際に見積書へ表示する金額。試算値引率とは独立です。'],sales[6],['原価取得・行数の整合状態。','原価件数が0なら原価の紐付け不足。','ID一致を確認して見積原価の漏れを防ぎます。']];
  calc[5]=['当該明細の基準売上に試算値引率を適用した税抜金額。円単位四捨五入。','テーブル100,000円を10%引きなら90,000円。','商品単位の値引き後売上を確認し、案件合計はB30で確認します。'];
  calc.forEach((x,i)=>add('計算',`${String.fromCharCode(65+i)}15`,...x));
  await fs.writeFile(`${out}/notes_manifest.json`,JSON.stringify(notes,null,2));
  // Historical demonstration data: separate from the current quote.
  for(const n of ['売上分析','売上データ']){sh[n]=w.worksheets.add(n);names.push(n);sh[n].showGridLines=false;sh[n].tabColor='#07988A';sh[n].getRange('A1:M55').format={font:{name:'Noto Sans CJK JP',size:11,color:'#233E39'},rowHeight:25,columnWidth:12,verticalAlignment:'center'};}
  const a='売上分析',raw='売上データ';
  merge(a,'A2:M2','売上分析  |  商品と季節の動きをつかむ');sh[a].getRange('A2:M2').format={fill:'#164D45',font:{color:'#FFFFFF',bold:true,size:18},rowHeight:44};
  merge(a,'A3:M3','すべて架空の販売履歴です。現在の見積・予定案件利益には加算されません。');
  val(a,'A5','対象業種');val(a,'B5','全業種');dv(a,'B5',['全業種','木工所','農家']);input(a,'B5');
  val(a,'D5','対象年');val(a,'E5',2026);numeric(a,'E5',2025,2026);input(a,'E5');
  val(a,'G5','開始月');val(a,'H5',1);val(a,'J5','終了月');val(a,'K5',12);for(const cell of ['H5','K5']){input(a,cell);sh[a].getRange(cell).dataValidation={rule:{type:'whole',operator:'between',formula1:1,formula2:12}};}
  merge(a,'A6:M6','');f(a,'A6','=IF(AND(OR(B5="全業種",B5="木工所",B5="農家"),OR(E5=2025,E5=2026),ISNUMBER(H5),ISNUMBER(K5),H5>=1,K5<=12,H5<=K5,MOD(H5,1)=0,MOD(K5,1)=0),"共通フィルター：KPI・3グラフ・ヒートマップが連動","期間・業種の入力を確認してください")');
  header(raw,'A4:I4',['日付','業種','商品','数量','税抜単価','税抜売上','計上原価','計上原価控除後差額','データ区分']);
  merge(raw,'A2:I2','架空の売上履歴  |  24か月 × 8商品 × 3件');
  const prices=[50000,15000,35000,2500,2000,1800,3200,4000];
  for(let y=2025;y<=2026;y++)for(let m=1;m<=12;m++)for(let p=0;p<8;p++)for(let k=0;k<3;k++){
    const qty=(p<4?2:20)+((m*7+p*3+k*5+y)%17),price=prices[p],sales=qty*price,cost=Math.round(sales*(.48+((m+p)%5)*.06));
    source.push({year:y,month:m,industry:p<4?'木工所':'農家',product:products[p],qty,price,sales,cost});
    let r=source.length+4;val(raw,`A${r}:I${r}`,[[new Date(Date.UTC(y,m-1,5+k*10)),p<4?'木工所':'農家',products[p],qty,price,null,cost,null,'架空サンプル']]);f(raw,`F${r}`,`=ROUND(D${r}*E${r},0)`);f(raw,`H${r}`,`=F${r}-G${r}`);
  }
  sh[raw].getRange('A5:I580').format={font:{name:'Noto Sans CJK JP',size:11},rowHeight:25};sh[raw].getRange('A5:A580').setNumberFormat('yyyy/mm/dd');sh[raw].getRange('D5:H580').setNumberFormat('#,##0');sh[raw].getRange('A1:C580').format.columnWidth=19;sh[raw].getRange('D1:I580').format.columnWidth=19;
  sh[raw].tables.add('A4:I580',true,'SalesHistory');
  note(raw,583,'入力範囲は5〜580行（576件）。試作版では行数固定。実データの移行・追記範囲拡張は次版で対応。','I');
  header(a,'A27:M27',['商品／月',...Array.from({length:12},(_,i)=>`${i+1}月`)]);
  for(let p=0;p<8;p++){let r=28+p;val(a,`A${r}`,products[p]);for(let m=1;m<=12;m++){let col=String.fromCharCode(65+m);f(a,`${col}${r}`,`=IF(OR($A$6<>"共通フィルター：KPI・3グラフ・ヒートマップが連動",${m}<$H$5,${m}>$K$5,AND($B$5<>"全業種",$B$5<>"${p<4?'木工所':'農家'}")),"",SUMIFS('売上データ'!$F$5:$F$580,'売上データ'!$C$5:$C$580,$A${r},'売上データ'!$A$5:$A$580,">="&DATE($E$5,${m},1),'売上データ'!$A$5:$A$580,"<"&DATE($E$5,${m+1},1)))`);}}
  sh[a].getRange('A1:A38').format.columnWidth=22;sh[a].getRange('B28:M35').setNumberFormat('#,##0,');sh[a].getRange('A28:M35').format.rowHeight=32;
  sh[a].getRange('B28:M35').conditionalFormats.add('colorScale',{colors:['#E3F2EE','#70C8B9','#077F74'],thresholds:['min',{type:'percentile',value:50},'max']});
  sh[a].getRange('B28:M35').conditionalFormats.addCustom('AND(ISNUMBER(B28),B28>MAX($B$28:$M$35)*0.55)',{font:{color:'#FFFFFF',bold:true}});
  merge(a,'A37:M37','ヒートマップ単位：千円（表示丸め）。濃い色ほど売上が大きく、空欄は選択範囲外です。');
  merge(a,'A38:M38','開始月・終了月は1〜12。対象年は2025／2026。データは売上データシートで確認できます。');
  // Visible supporting summary below dashboard. All charts bind these live cells.
  header(a,'A42:C42',['月','売上（円）','原価控除後差額（円）']);
  for(let m=1;m<=12;m++){let r=42+m,col=String.fromCharCode(65+m);val(a,`A${r}`,`${m}月`);f(a,`B${r}`,`=SUM(${col}28:${col}35)`);f(a,`C${r}`,`=IF(OR($A$6<>"共通フィルター：KPI・3グラフ・ヒートマップが連動",${m}<$H$5,${m}>$K$5),0,SUMIFS('売上データ'!$H$5:$H$580,'売上データ'!$A$5:$A$580,">="&DATE($E$5,${m},1),'売上データ'!$A$5:$A$580,"<"&DATE($E$5,${m+1},1),'売上データ'!$B$5:$B$580,IF($B$5="全業種","*",$B$5)))`);}
  header(a,'E42:F42',['商品','売上（円）']);for(let p=0;p<8;p++){val(a,`E${43+p}`,products[p]);f(a,`F${43+p}`,`=SUM(B${28+p}:M${28+p})`);}
  header(a,'H42:I42',['業種','売上（円）']);val(a,'H43:H44',[['木工所'],['農家']]);f(a,'I43','=SUM(F43:F46)');f(a,'I44','=SUM(F47:F50)');
  for(const [label,area,anchor,formula,fmt] of [['税抜売上','A8:D8','A9','=SUM(B43:B54)','#,##0" 円"'],['計上原価控除後差額','F8:I8','F9','=SUM(C43:C54)','#,##0" 円"'],['同差額率','K8:M8','K9','=IF(A9=0,"算出不可",F9/A9)','0.0%']]){merge(a,area,label);let col=anchor[0],end=area.split(':')[1][0];merge(a,`${col}9:${end}10`,'');f(a,anchor,formula);sh[a].getRange(`${col}9:${end}10`).format={fill:'#E3F1EB',font:{size:22,bold:true,color:'#077F74'}};sh[a].getRange(anchor).setNumberFormat(fmt);}
  merge(a,'A11:M11','差額は売上履歴の計上原価のみを控除。会計上の粗利・営業利益や現在の見積利益とは区別します。');
  for(const [type,range,title,start,end] of [['doughnut','H42:I44','業種別売上構成','A13','D25'],['line','A42:B54','月別売上（千円）','E13','I25'],['bar','E42:F50','商品別売上（千円）','J13','N25']]){const chart=sh[a].charts.add(type,sh[a].getRange(range));chart.title=title;chart.setPosition(start,end);chart.titleTextStyle.typeface='Noto Sans CJK JP';chart.titleTextStyle.fontSize=13;chart.hasLegend=type==='doughnut';chart.legend={position:'bottom',textStyle:{typeface:'Noto Sans CJK JP',fontSize:10}};if(type!=='doughnut'){chart.yAxis={numberFormatCode:'#,##0,',numberFormatSourceLinked:false,textStyle:{typeface:'Noto Sans CJK JP',fontSize:10}};chart.xAxis={textStyle:{typeface:'Noto Sans CJK JP',fontSize:10}};}chart.series.items[0].fill='#07988A';if(type==='line')chart.series.items[0].line={fill:'#07988A',width:3};}
  sh[a].getRange('B43:C54').setNumberFormat('#,##0');sh[a].getRange('F43:F50').setNumberFormat('#,##0');
  await fs.writeFile(`${out}/sales_fixture.json`,JSON.stringify(source,null,2));
}
export async function verifyAnalysis({w,sh,val,recalc,test}){
  const n='売上分析',read=a=>sh[n].getRange(a).values[0][0];
  for(const [industry,year,start,end] of [['全業種',2026,1,12],['木工所',2025,3,6],['農家',2026,12,12],['全業種',2025,1,1]]){
    for(const [a,v] of [['B5',industry],['E5',year],['H5',start],['K5',end]])val(n,a,v);recalc();
    const rows=source.filter(r=>r.year===year&&r.month>=start&&r.month<=end&&(industry==='全業種'||r.industry===industry));
    test(`分析売上 ${industry}/${year}/${start}-${end}`,rows.reduce((s,r)=>s+r.sales,0),read('A9'));
    test(`分析差額 ${industry}/${year}/${start}-${end}`,rows.reduce((s,r)=>s+r.sales-r.cost,0),read('F9'));
    test(`ヒートマップ照合 ${industry}`,rows.filter(r=>r.product===products[0]&&r.month===start).reduce((s,r)=>s+r.sales,0),Number(read(`${String.fromCharCode(65+start)}28`)));
  }
  val(n,'H5',12);val(n,'K5',1);recalc();test('逆転期間を警告','期間・業種の入力を確認してください',read('A6'));test('逆転期間集計除外',0,read('A9'));
  val(n,'B5','全業種');val(n,'E5',2026);val(n,'H5',1);val(n,'K5',12);recalc();
  test('D12自動状態','入力OK',sh['見積入力'].getRange('D12').values[0][0]);val('見積入力','F16',-1);recalc();test('D12異常時','要確認',sh['見積入力'].getRange('D12').values[0][0]);val('見積入力','F16',50000);recalc();
}
