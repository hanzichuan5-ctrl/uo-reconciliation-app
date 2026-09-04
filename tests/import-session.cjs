// Run with: NODE_PATH=/path/to/node_modules node tests/import-session.cjs
// Uses an isolated browser and synthetic workbooks only; never reads user Excel files.
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const http = require('node:http');
const { chromium } = require('playwright');
const ExcelJS = require('../assets/exceljs.min.js');
const root = path.resolve(__dirname, '..');
const company = '北京世界星辉科技有限责任公司';
const hd = '上海幻电', wh = '芜湖享游';
const bankHeaders = ['收入金额','对方账号名称','本方账号名称','交易日期','货币','交易流水号'];
const invoiceHeaders = ['销方名称','购买方名称','价税合计','发票状态','是否正数发票','开票日期','发票号码'];
const mappingHeaders = ['渠道','公司名称','开票金额','账期'];

async function workbook(name, sheets) {
  const book = new ExcelJS.Workbook();
  for (const [title, rows] of sheets) book.addWorksheet(title).addRows(rows);
  return { name, mimeType:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', buffer:Buffer.from(await book.xlsx.writeBuffer()) };
}
const bankRow = (entity, amount=entity===hd?100:200, id=entity+'-bank', month='2026-07') => [amount,company,entity,month+'-15','CNY',id];
const invoiceRow = (entity, amount=entity===hd?100:200, month='2026-07') => [entity,company,amount,'正常','是',month+'-16',entity+'-invoice'];
const bankFile = (entity, rows=[bankRow(entity)], name=entity+'银行流水.xlsx') => workbook(name,[['原始流水对账单明细',[bankHeaders,...rows]]]);
const invoiceFile = (entity, rows=[invoiceRow(entity)], name=entity+'全量发票.xlsx') => workbook(name,[['发票基础信息',[invoiceHeaders,...rows]]]);
const mappingFile = (wuhu=true) => workbook('开票申请.xlsx',[
  ['2026年（幻电）',[mappingHeaders,['360',company,100,'2026-06']]],
  ...(wuhu?[['2026年（芜湖）',[mappingHeaders,['360',company,200,'2026-06'],['360',company,12,'']]]]:[])
]);

async function upload(page, files, viaButton=false) {
  if (viaButton) {
    const chooserPromise=page.waitForEvent('filechooser');
    await page.locator('#appendFilesBtn').click();
    await (await chooserPromise).setFiles(files);
  } else await page.locator('#fileInput').setInputFiles(files);
  await page.waitForFunction(()=>!document.querySelector('#fileInput').disabled&&document.querySelector('#fileInput').value==='');
}
const snapshot = page => page.evaluate(()=>{
  const s=UOReconcile.state;
  return {files:s.files.map(f=>f.name),checks:s.checks,mapping:s.mapping,results:s.results,banks:s.banks,invoices:s.invoices,month:s.month,overrides:s.overrides,groups:s.matchGroups,entries:s.entryStatus,checksState:s.checkStatus,dirty:s.dirty,exportDisabled:document.querySelector('#exportBtn').disabled,exported:UOReconcile.exportRows()};
});

async function main() {
  const server=http.createServer((req,res)=>{
    const file=req.url==='/'?path.join(root,'index.html'):req.url.startsWith('/assets/')?path.resolve(root,'.'+decodeURIComponent(req.url)):null;
    if(!file||!file.startsWith(root+path.sep)||!fs.existsSync(file)){res.writeHead(404);res.end();return;}
    res.setHeader('Content-Type',file.endsWith('.js')?'application/javascript':file.endsWith('.png')?'image/png':'text/html; charset=utf-8');
    res.end(fs.readFileSync(file));
  });
  await new Promise(resolve=>server.listen(0,'127.0.0.1',resolve));
  const url=`http://127.0.0.1:${server.address().port}/`;
  const browser=await chromium.launch({headless:true});
  const errors=[];
  let passed=0;
  const pass=label=>{passed++;console.log(`PASS ${passed}: ${label}`)};
  const fresh=async(rememberedOperator='')=>{
    const context=await browser.newContext({viewport:{width:1640,height:1000},acceptDownloads:true});
    if(rememberedOperator)await context.addInitScript(name=>localStorage.setItem('uo-operator-v1',name),rememberedOperator);
    await context.route('**/*',route=>route.request().url().startsWith(url)?route.continue():route.abort());
    const page=await context.newPage();page.on('pageerror',error=>errors.push(error.message));
    await page.goto(url);await page.waitForFunction(()=>window.UOReconcile);return page;
  };
  try {
    const mapping=await mappingFile(), hb=await bankFile(hd), hi=await invoiceFile(hd), wb=await bankFile(wh), wi=await invoiceFile(wh);
    const page=await fresh();
    await upload(page,[mapping,hi,hb]);
    let s=await snapshot(page);
    assert.equal(s.files.length,3);assert.equal(s.results.length,1);assert.equal(s.exportDisabled,false);
    assert(s.mapping.every(r=>r.entity===hd));assert(!s.checks.some(c=>c.msg.includes('芜湖')));assert(!s.checks.some(c=>c.level==='error'));
    assert.equal(await page.locator('#appendFilesBtn').isEnabled(),true);
    pass('幻电单主体三文件成功，无芜湖提醒或五文件数量提醒');

    await page.locator('select[data-action="status"][data-side="bank"]').first().selectOption('待复核');
    await page.evaluate(()=>{
      const s=UOReconcile.state,d=s.mapping[0],b=s.banks[0];
      const fp='d|'+[d.entity,d.channelKey,d.companyKey,d.periodKey,String(d.amount||''),d.requestBatch||'',d.sheet||'',d.row||''].join('|');
      s.matchGroups.push({id:'test-hd',type:'bank',sourceKeys:[b.stableKey],detailFps:[fp],status:'confirmed',confirmer:'测试员',confirmTime:'2026-07-31 12:00:00',_orphan:true});
      s.checkStatus['bank:'+fp]={status:'已勾选',confirmer:'测试员',time:'2026-07-31 12:00:00'};
      s.entryStatus[b.stableKey]=1;
      localStorage.setItem('uo-match-groups-v1',JSON.stringify(s.matchGroups));
      localStorage.setItem('uo-check-status-v1',JSON.stringify(s.checkStatus));
      localStorage.setItem('uo-entry-status-v1',JSON.stringify(s.entryStatus));
    });
    const original=await snapshot(page);
    await upload(page,[{...hb,name:'重命名的同一文件.xlsx'}],true);
    s=await snapshot(page);assert.equal(s.files.length,3);assert.equal(s.banks.length,1);assert(s.checks.some(c=>c.msg.includes('重复文件已跳过')));
    assert.deepEqual(s.overrides,original.overrides);
    pass('完全相同内容即使改名也跳过，且复核状态不变');

    await upload(page,[wb],true);
    s=await snapshot(page);assert.equal(s.files.length,4);assert.equal(s.results.length,1);assert.equal(s.exportDisabled,false);
    assert(s.checks.some(c=>c.msg==='芜湖享游待补充：缺少全量发票'));
    assert(!s.checks.some(c=>c.msg.includes('芜湖')&&c.msg.includes('行未导入')));
    assert.deepEqual(s.overrides,original.overrides);assert.deepEqual(s.checksState,original.checksState);assert.deepEqual(s.entries,original.entries);
    assert(!s.groups[0]._orphan);assert(s.dirty);
    pass('补一份芜湖文件只提示具体缺项，不生成误导结论；幻电人工状态保留');

    await upload(page,[wi],true);
    s=await snapshot(page);assert.equal(s.files.length,5);assert.equal(s.results.length,2);assert.equal(s.month,'2026-07');assert.equal(s.exportDisabled,false);
    assert(!s.checks.some(c=>c.msg.includes('待补充：')));assert(s.checks.some(c=>c.msg.includes('芜湖')&&c.msg.includes('行未导入')));
    const hResult=s.results.find(r=>r.entity===hd);assert.equal(s.overrides[hResult.id].bank.status,'待复核');
    assert.deepEqual(s.checksState,original.checksState);assert.deepEqual(s.entries,original.entries);
    assert.equal(s.groups[0].confirmTime,original.groups[0].confirmTime);
    pass('芜湖补齐后加入核对并显示真实行异常，幻电勾选和操作记录不变');

    if(process.env.UO_TEST_SCREENSHOT) await page.screenshot({path:process.env.UO_TEST_SCREENSHOT,fullPage:true,animations:'disabled'});
    const downloadPromise=page.waitForEvent('download');await page.locator('#exportBtn').click();
    const download=await downloadPromise,stream=await download.createReadStream(),chunks=[];for await(const chunk of stream)chunks.push(chunk);
    const exported=new ExcelJS.Workbook();await exported.xlsx.load(Buffer.concat(chunks));
    assert(exported.getWorksheet('02_核对清单'));assert(exported.getWorksheet('99_导入日志'));assert.equal((await snapshot(page)).dirty,false);
    const logRows=[];exported.getWorksheet('99_导入日志').eachRow(row=>logRows.push(row.values.join(' ')));
    assert(logRows.some(row=>row.includes('芜湖')));assert(!logRows.some(row=>row.includes('银行流水缺失')));
    pass('补传后的真实 Excel 导出成功、日志正确，导出后保存人工操作检查点');

    const corrected=await bankFile(hd,[bankRow(hd,150)],'幻电修订流水.xlsx');
    page.once('dialog',dialog=>dialog.dismiss());await upload(page,[corrected],true);
    s=await snapshot(page);assert.equal(s.banks.find(r=>r.entity===hd).amount,100);assert.equal(s.files.length,5);
    pass('取消同主体同类型替换，原文件、金额和记录保留');

    page.once('dialog',dialog=>dialog.accept());await upload(page,[corrected],true);
    s=await snapshot(page);assert.equal(s.banks.find(r=>r.entity===hd).amount,150);assert.equal(s.files.length,5);
    assert.equal(s.banks.filter(r=>r.entity===hd).length,1);assert(s.groups[0]._sourceChanged);assert.equal(s.entries[original.banks[0].stableKey],undefined);
    assert(s.checks.some(c=>c.msg.includes('人工复核需重新确认')));assert.deepEqual(s.checksState,original.checksState);
    assert.equal(s.overrides[s.results.find(r=>r.entity===hd).id].bank.status,'');
    pass('确认替换不会重复累计；内容变化撤下旧人工覆盖，并标记旧匹配需复核');

    const beforeBad=await snapshot(page);
    await upload(page,[{name:'坏文件.xlsx',mimeType:hb.mimeType,buffer:Buffer.from('broken workbook')}],true);
    s=await snapshot(page);assert.deepEqual(s.files,beforeBad.files);assert.deepEqual(s.banks,beforeBad.banks);assert.deepEqual(s.overrides,beforeBad.overrides);
    assert.equal(s.exportDisabled,false);assert(s.checks.some(c=>c.msg.includes('未导入：坏文件')));
    pass('坏文件不破坏已有数据，也不阻断现有主体导出');

    const changedKey=await bankFile(hd,[bankRow(hd,100,'另一个流水号')],'另一流水.xlsx');
    page.once('dialog',dialog=>dialog.accept());await upload(page,[changedKey],true);
    assert((await snapshot(page)).groups[0]._orphan);
    page.once('dialog',dialog=>dialog.accept());await upload(page,[hb],true);
    s=await snapshot(page);assert(!s.groups[0]._orphan);assert(!s.groups[0]._sourceChanged);
    assert(!s.checks.some(c=>c.msg.includes('悬空')||c.msg.includes('引用了当前导入')));
    pass('恢复原文件后解除悬空和来源变更提醒');

    const laterBank=await bankFile(wh,[bankRow(wh,200,wh+'-bank','2026-08')],'芜湖八月流水.xlsx');
    page.once('dialog',dialog=>dialog.accept());await upload(page,[laterBank],true);
    assert.equal((await snapshot(page)).month,'2026-07');
    pass('补传含其他月份的数据不会切换当前核对月份');

    await page.evaluate(()=>{
      const s=UOReconcile.state;s.matchGroups.push({...s.matchGroups[0],id:'not-exported'});s.dirty=true;
      localStorage.setItem('uo-match-groups-v1',JSON.stringify(s.matchGroups));
    });
    page.once('dialog',dialog=>dialog.dismiss());await page.locator('#resetBtn').click();assert((await snapshot(page)).groups.some(g=>g.id==='not-exported'));
    page.once('dialog',dialog=>dialog.accept());await page.locator('#resetBtn').click();await page.waitForFunction(()=>window.UOReconcile&&UOReconcile.state.files.length===0);
    s=await snapshot(page);assert.equal(s.groups.length,1);assert.equal(s.groups[0].id,'test-hd');assert.deepEqual(s.entries,original.entries);
    pass('重新导入可取消；确认后清空文件并回退未导出操作，保留已导出记录');

    const wp=await fresh();await upload(wp,[wi,mapping,wb]);s=await snapshot(wp);
    assert(s.mapping.every(r=>r.entity===wh));assert.equal(s.results.length,1);assert(!s.checks.some(c=>c.msg.includes('上海幻电')));assert.equal(s.exportDisabled,false);
    pass('芜湖单主体与不同文件选择顺序同样成立');

    const partial=await fresh();await upload(partial,[mapping,wb]);s=await snapshot(partial);
    assert.equal(s.results.length,0);assert(s.checks.some(c=>c.msg.includes('缺少全量发票')));assert(!s.checks.some(c=>c.msg.includes('行未导入')));
    await upload(partial,[wi],true);assert.equal((await snapshot(partial)).results.length,1);
    pass('单主体分两次上传完整闭环');

    const empty=await fresh();await upload(empty,[mapping,await bankFile(hd,[]),hi]);s=await snapshot(empty);
    assert(!s.checks.some(c=>c.msg.includes('缺少银行流水')));assert.equal(s.exportDisabled,false);assert.equal(s.banks.length,0);
    pass('已上传但没有收入行的银行文件不被误判缺失');

    const combinedBank=await bankFile(hd,[bankRow(hd),bankRow(wh)],'两主体银行.xlsx');
    const combinedInvoice=await invoiceFile(hd,[invoiceRow(hd),invoiceRow(wh)],'两主体发票.xlsx');
    const mixed=await fresh();await upload(mixed,[mapping,combinedBank,combinedInvoice]);
    mixed.once('dialog',dialog=>dialog.accept());await upload(mixed,[corrected],true);s=await snapshot(mixed);
    assert.equal(s.banks.length,2);assert.equal(s.banks.find(r=>r.entity===wh).amount,200);assert.equal(s.banks.find(r=>r.entity===hd).amount,150);
    assert.equal(s.files.length,4);assert.equal(s.exportDisabled,false);
    mixed.once('dialog',dialog=>dialog.accept());await upload(mixed,[combinedBank],true);s=await snapshot(mixed);
    assert.equal(s.banks.length,2);assert.equal(s.files.length,3);assert.equal(s.banks.find(r=>r.entity===hd).amount,100);
    pass('合并主体文件只替换重叠主体，另一主体保留；恢复合并文件也不重复');

    const missingMap=await mappingFile(false);mixed.once('dialog',dialog=>dialog.accept());await upload(mixed,[missingMap],true);s=await snapshot(mixed);
    assert(s.checks.some(c=>c.level==='error'&&c.entity===wh));assert.equal(s.exportDisabled,false);assert(s.results.some(r=>r.entity===hd));
    pass('某主体申请明细错误不阻断另一完整主体导出');

    const isolated=await fresh();await isolated.evaluate(()=>{
      UOReconcile.state.matchGroups.push({id:'old-wuhu',type:'bank',sourceKeys:['芜湖享游|old'],detailFps:['d|芜湖享游|old'],status:'confirmed'});
    });await upload(isolated,[mapping,hi,hb]);s=await snapshot(isolated);
    assert(!s.checks.some(c=>c.msg.includes('芜湖')||c.msg.includes('悬空')));assert.equal(s.exported.matchGroupRows.length,0);
    pass('未上传主体的旧匹配记录静默保留、不产生悬空提醒或进入本次导出');

    const selectedPage=await fresh();await upload(selectedPage,[mapping,hb,hi]);
    await selectedPage.evaluate(()=>{
      const s=UOReconcile.state,r=s.results[0];s.overrides[r.id].bank={status:'已回款',selected:[s.banks[0].id]};s.expanded.add(r.id);s.dirty=true;
    });
    const reorderedMap=await workbook('调整工作表顺序的开票申请.xlsx',[
      ['2026年（芜湖）',[mappingHeaders,['360',company,200,'2026-06']]],
      ['2026年（幻电）',[mappingHeaders,['360',company,100,'2026-06']]]
    ]);
    selectedPage.once('dialog',dialog=>dialog.accept());await upload(selectedPage,[reorderedMap,wb,wi],true);
    s=await snapshot(selectedPage);const hdAfterReorder=s.results.find(r=>r.entity===hd);
    assert.equal(hdAfterReorder.id,'g2');assert.equal(s.overrides.g2.bank.status,'已回款');assert.equal(s.overrides.g1.bank.status,'');
    const previousId=s.banks.find(r=>r.entity===hd).id;
    const sameRowsNewBook=await workbook('幻电同数据新版本.xlsx',[
      ['原始流水对账单明细',[bankHeaders,bankRow(hd)]],['版本说明',[['第二版']]]
    ]);
    selectedPage.once('dialog',dialog=>dialog.accept());await upload(selectedPage,[sameRowsNewBook],true);
    s=await snapshot(selectedPage);const nextId=s.banks.find(r=>r.entity===hd).id;assert.notEqual(nextId,previousId);
    assert.deepEqual(s.overrides.g2.bank.selected,[nextId]);assert.equal(s.overrides.g2.bank.status,'已回款');
    pass('主体增量导致组编号变化、同数据文件替换导致行 ID 变化时，人工选择仍指向正确单据');

    await selectedPage.locator('#toggleUploadDetails').click();
    if(!(await selectedPage.locator('#uploadPanel').getAttribute('class')).includes('is-collapsed'))await selectedPage.locator('#toggleUploadDetails').click();
    await selectedPage.locator('select[data-action="status"][data-side="bank"]').first().selectOption('待复核');
    assert((await selectedPage.locator('#uploadPanel').getAttribute('class')).includes('is-collapsed'));
    pass('人工复核不会反复展开用户已收起的上传详情');

    const dropPage=await fresh();await upload(dropPage,[mapping,hb,hi]);
    await dropPage.evaluate(({name,mimeType,base64})=>{
      const data=new DataTransfer(),bytes=Uint8Array.from(atob(base64),c=>c.charCodeAt(0));
      data.items.add(new File([bytes],name,{type:mimeType}));
      document.querySelector('#dropzone').dispatchEvent(new DragEvent('drop',{bubbles:true,dataTransfer:data}));
    },{name:wb.name,mimeType:wb.mimeType,base64:wb.buffer.toString('base64')});
    await dropPage.waitForFunction(()=>UOReconcile.state.files.length===4&&!document.querySelector('#fileInput').disabled);
    s=await snapshot(dropPage);assert.equal(s.banks.length,2);assert.equal(s.results.length,1);assert.equal(s.exportDisabled,false);
    pass('拖拽补传与补充文件按钮遵循相同追加规则');

    const noOperator=await fresh('以前填写的姓名');await upload(noOperator,[mapping,hb,hi]);
    assert.equal(await noOperator.locator('#operatorInput').count(),0);
    assert(!(await noOperator.locator('header.topbar').innerText()).includes('本地确认人'));
    await noOperator.evaluate(()=>{
      const s=UOReconcile.state,d=s.mapping[0],b=s.banks[0];
      const fp='d|'+[d.entity,d.channelKey,d.companyKey,d.periodKey,String(d.amount||''),d.requestBatch||'',d.sheet||'',d.row||''].join('|');
      s.matchGroups.push({id:'historical-group',type:'bank',sourceKeys:[b.stableKey],detailFps:[fp],status:'confirmed',confirmer:'历史确认人',confirmTime:'2026-07-31 12:00:00'});
    });
    await noOperator.locator('[data-module="checklist"]').click();
    await noOperator.locator('[data-checklist-tab="bank-pending"]').click();
    await noOperator.locator('[data-chk-side="bank"][data-chk-set="已勾选"]').click();
    s=await snapshot(noOperator);
    assert.equal(s.groups[0].confirmer,'历史确认人');
    assert.equal(Object.values(s.checksState)[0].confirmer,'');assert(Object.values(s.checksState)[0].time);
    assert.equal(s.exported.matchGroupRows[0]['确认人'],'历史确认人');
    assert.equal(s.exported.checkStatusRows[0]['回款确认人'],'');
    if(process.env.UO_TEST_NO_OPERATOR_SCREENSHOT){
      await noOperator.setViewportSize({width:902,height:897});
      await noOperator.locator('[data-module="reconcile"]').click();
      await noOperator.screenshot({path:process.env.UO_TEST_NO_OPERATOR_SCREENSHOT,animations:'disabled'});
    }
    pass('移除本地确认人输入框，历史署名保留，新勾选不沿用旧姓名且仍记录时间');

    const narrow=await fresh();await narrow.setViewportSize({width:390,height:844});await upload(narrow,[mapping,hi,hb]);
    assert(await narrow.locator('#appendFilesBtn').isVisible());
    const box=await narrow.locator('#appendFilesBtn').boundingBox();assert(box.x>=0&&box.x+box.width<=390);
    if(process.env.UO_TEST_MOBILE_SCREENSHOT)await narrow.screenshot({path:process.env.UO_TEST_MOBILE_SCREENSHOT,fullPage:true,animations:'disabled'});
    pass('窄屏补充文件入口可见且不溢出');
    assert.deepEqual(errors,[]);console.log(`All ${passed} scenarios passed; no browser JavaScript errors.`);
  } finally {await browser.close();await new Promise(resolve=>server.close(resolve));}
}
main().catch(error=>{console.error(error);process.exitCode=1});
