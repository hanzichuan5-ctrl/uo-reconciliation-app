import fs from 'node:fs/promises';
import path from 'node:path';
import assert from 'node:assert/strict';
import { Workbook, SpreadsheetFile, FileBlob } from '@oai/artifact-tool';

const input=JSON.parse(await fs.readFile(process.argv[2],'utf8'));
const output=process.argv[3];
const mode=process.argv[4]||'income';
const fmt='#,##0.00;(#,##0.00);"—"';
const status={ready:'待在UO操作',review:'待核实',done:'系统已同步',harmony:'鸿蒙仅核开票',scope:'范围待核实'};
let wb;
if(mode==='preview-template'){
  wb=await SpreadsheetFile.importXlsx(await FileBlob.load(input.path));
  const img=await wb.render({sheetName:'回款',range:'A1:D18',scale:1.5,format:'png'});
  await fs.writeFile(output,new Uint8Array(await img.arrayBuffer()));
  process.exit(0);
}
wb=Workbook.create();
if(mode==='income'){
  const cases=input.income.cases;
  const s=wb.worksheets.add('收入操作清单'),e=wb.worksheets.add('发票依据');
  s.showGridLines=false;e.showGridLines=false;
  const headers=['主体','购买方公司','渠道','收入账期','游戏','含税发票金额（元）','系统明细金额（元）','差额（元）','核对状态','下一步','实际开票月份','核对组ID'];
  s.getRange('A2').values=[[`${input.month} 收入操作清单`]];
  s.getRange('A3').values=[['从实际发票出发；待核实事项保留原始依据。本地完成记录与系统状态分别保存。']];
  s.getRange('A5:L5').values=[headers];
  const rows=cases.map(c=>{const candidate=c.candidates?.length===1?c.candidates[0]:null;const candidateTotal=candidate?.systems.reduce((a,s)=>a+s.amountCents,0);return [c.entity,c.company,c.channel,c.months.join('、')||(candidate?candidate.month+'（候选）':'待核实'),c.games.join('、')||(candidate?candidate.games.join('、')+'（候选）':'待对应'),null,c.systemTotalCents==null?(candidateTotal==null?null:candidateTotal/100):c.systemTotalCents/100,null,
    c.localDone?'本地已记录完成，待系统回读':c.candidates?.length?'候选待确认':c.reviewType==='missing'?'缺少系统资料':c.reviewType==='difference'?'金额差异待核实':status[c.status],c.reason||c.action,input.month,c.id]});
  const evidence=cases.flatMap(c=>c.invoices.map(i=>[c.id,i.no,i.entity,i.company,new Date(`${i.date}T00:00:00Z`),i.amountCents/100,i.invoiceState,i.remark,`${i.source.file} / ${i.source.sheet} / 第${i.source.row}行`]));
  e.getRange('A2').values=[['实际发票依据']];
  e.getRange('A4:I4').values=[['核对组ID','发票号码','主体','购买方名称','开票日期','价税合计（元）','发票状态','发票备注','原始来源']];
  if(rows.length){
    s.getRange(`A6:L${5+rows.length}`).values=rows;
    e.getRange(`A5:I${4+evidence.length}`).values=evidence;
    for(let i=0;i<rows.length;i++){
      const rn=i+6;
      s.getRange(`F${rn}`).formulas=[[`=SUMIFS('发票依据'!$F$5:$F$${evidence.length+4},'发票依据'!$A$5:$A$${evidence.length+4},L${rn})`]];
      if(rows[i][6]!=null)s.getRange(`H${rn}`).formulas=[[`=F${rn}-G${rn}`]];
    }
    s.getRange(`E${rows.length+7}`).values=[['渠道发票合计（含税）']];
    s.getRange(`F${rows.length+7}`).formulas=[[`=SUMIFS(F6:F${rows.length+5},I6:I${rows.length+5},"<>范围待核实")`]];
    e.getRange(`E${evidence.length+6}`).values=[['合计']];
    e.getRange(`F${evidence.length+6}`).formulas=[[`=SUM(F5:F${evidence.length+4})`]];
    s.tables.add(`A5:L${rows.length+5}`,true,'IncomeCases').style='TableStyleLight1';
    e.tables.add(`A4:I${evidence.length+4}`,true,'InvoiceEvidence').style='TableStyleLight1';
  }
  for(const [sheet,end,cols,header] of [[s,rows.length+7,12,5],[e,evidence.length+6,9,4]]){
    sheet.getRangeByIndexes(0,0,end,cols).format.font={name:'Arial',size:10,color:'#24313A'};
    sheet.getRangeByIndexes(0,0,end,cols).format.rowHeight=26;
    sheet.getRangeByIndexes(0,0,end,cols).format.verticalAlignment='center';
    sheet.getRangeByIndexes(header-1,0,1,cols).format={fill:'#34495C',font:{bold:true,color:'#FFFFFF'},rowHeight:30};
    sheet.getRangeByIndexes(header-1,0,1,cols).format.horizontalAlignment='center';
    sheet.getRange('A2').format.font={size:15,bold:true};
    sheet.freezePanes.freezeRows(header);
  }
  [17,38,17,24,40,23,23,21,31,76,19,28].forEach((v,i)=>s.getRangeByIndexes(0,i,rows.length+8,1).format.columnWidth=v);
  [28,29,17,38,18,23,18,66,96].forEach((v,i)=>e.getRangeByIndexes(0,i,evidence.length+7,1).format.columnWidth=v);
  s.getRange(`F6:H${rows.length+7}`).setNumberFormat(fmt);
  e.getRange(`F5:F${evidence.length+6}`).setNumberFormat(fmt);
  e.getRange(`E5:E${evidence.length+4}`).setNumberFormat('yyyy-mm-dd');
  e.getRange(`B5:B${evidence.length+4}`).setNumberFormat('@');
  e.getRange(`G5:G${evidence.length+4}`).format.horizontalAlignment='center';
  s.getRange(`E6:E${rows.length+5}`).format.wrapText=true;
  s.getRange(`J6:J${rows.length+5}`).format.wrapText=true;
  s.getRange(`A6:L${rows.length+5}`).format.rowHeight=52;
  s.getRange(`I6:I${rows.length+5}`).conditionalFormats.add('containsText',{text:'待核实',format:{fill:'#FFF2D9',font:{color:'#8A5618'}}});
  wb.recalculate();
  assert.equal(Math.round(s.getRange(`F${rows.length+7}`).values[0][0]*100),input.income.inScopeAmountCents);
  console.log((await wb.inspect({kind:'match',searchTerm:'#REF!|#DIV/0!|#VALUE!|#NAME\\?|#NUM!|#SPILL!',options:{useRegex:true,maxResults:20},maxChars:1500})).ndjson);
  if(input.preview){
    for(const [sheet,range,name] of [['收入操作清单','A1:J10','income'],['发票依据','B1:H10','invoices']]){
      const img=await wb.render({sheetName:sheet,range,scale:1.2,format:'png'});
      await fs.writeFile(path.join(path.dirname(output),`${name}.png`),new Uint8Array(await img.arrayBuffer()));
    }
  }
}else if(mode==='receipt-payload'){
  const s=wb.worksheets.add('回款');
  for(const b of input.blocks){
    const rows=[];
    for(let i=0;i<b.capacity;i++){
      const r=b.rows[i]; rows.push([String(i+1),r?.payer??null,r?r.amountCents/100:null,r?.date??null]);
    }
    s.getRange(`A${b.start}:D${b.total-1}`).values=rows;
    s.getRange(`C${b.total}`).formulas=[[`=SUM(C${b.start}:C${b.total-1})`]];
  }
  wb.recalculate();
  for(const b of input.blocks)assert.equal(Math.round(s.getRange(`C${b.total}`).values[0][0]*100),b.rows.reduce((a,r)=>a+r.amountCents,0));
}
await(await SpreadsheetFile.exportXlsx(wb)).save(output);
console.log(output);
