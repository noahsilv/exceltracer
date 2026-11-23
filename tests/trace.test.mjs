import assert from 'node:assert/strict';
import test from 'node:test';
import { inspect } from '../public/build/excel.js';
import { Navigator } from '../public/build/navigation.js';

function mockExcel() {
  let active;
  let nextError;
  const inputs = new Map(), dependents = new Map(), sheets = new Map(), cells = new Map();
  function cell(sheetId, address, value, formula = '', rows = 1, columns = 1) {
    if (!sheets.has(sheetId)) sheets.set(sheetId, {
      id: sheetId, name: sheetId, load() {}, activate() {},
      getRange(a) { const r = cells.get(sheetId + '!' + a); if (!r) throw Error('Missing cell'); return r; }
    });
    const key = sheetId + '!' + address;
    const result = {
      address: "'" + sheetId.replaceAll("'", "''") + "'!" + address,
      worksheet: sheets.get(sheetId), rowCount: rows, columnCount: columns,
      text: [[String(value)]], formulas: [[formula || value]], load() {},
      select() { active = this; }, getCell() { return cells.get(sheetId + '!' + address.split(':')[0]); },
      getDirectPrecedents() { return related(inputs.get(key)); },
      getDirectDependents() { return related(dependents.get(key)); }
    };
    cells.set(key, result);
    return result;
  }
  function related(items) {
    if (!items) nextError = Object.assign(Error('No links'), {code:'ItemNotFound'});
    return {ranges:{items:items || [], load() {}}};
  }
  active = cell('Summary', 'B2', '1.50x', "='Cash Flow'!B4/Debt!B4");
  const cash = cell('Cash Flow','B4',150,'=B2-B3');
  const debt = cell('Debt','B4',100,'=B2+B3');
  const interest = cell('Debt','B2',25);
  const principal = cell('Debt','B3',75);
  inputs.set('Summary!B2',[cash,debt]);
  inputs.set('Debt!B4',[interest,principal]);
  dependents.set('Debt!B2',[debt]);
  globalThis.Excel = {run: fn => fn({
    workbook:{getSelectedRange:() => active, worksheets:{getItem:id => sheets.get(id)}},
    sync:async () => { if (nextError) { const error = nextError; nextError = undefined; throw error; } }
  })};
  return {cell, inputs, sheets, active:() => active};
}

test('walk Summary → Debt → interest, then restore both cells and the original mode',async () => {
  const mock = mockExcel();
  const nav = new Navigator(inspect);
  await nav.refresh('inputs');
  assert.equal(nav.current.value,'1.50x');
  assert.deepEqual(nav.current.related.map(row => [row.sheet,row.label,row.value]),[
    ['Cash Flow','B4','150'],['Debt','B4','100']
  ]);
  await nav.jump(nav.current.related[1]);
  await nav.jump(nav.current.related[0]);
  assert.equal(mock.active().address,"'Debt'!B2");
  assert.equal(nav.current.total,0);
  await nav.refresh('dependents');
  assert.equal(nav.current.related[0].label,'B4');
  await nav.back();
  assert.equal(nav.mode,'inputs');
  assert.equal(nav.current.value,'100');
  await nav.back();
  assert.equal(mock.active().address,"'Summary'!B2");
  assert.equal(nav.canGoBack,false);
});

test('groups a large range and opens only its first cell on a sheet with punctuation',async () => {
  const mock = mockExcel();
  const name = "Bank's! Inputs";
  mock.cell(name,'A1',987);
  mock.inputs.set('Summary!B2',[mock.cell(name,'A1:A50000',987,'',50000)]);
  const trace = await inspect('inputs');
  assert.deepEqual(trace.related[0],{
    sheetId:name,sheet:name,address:'A1',label:'A1:A50000',value:'First cell A1 · 987'
  });
  assert.equal((await inspect('inputs',trace.related[0])).value,'987');
});

test('manual selection starts a fresh trail; failed navigation preserves history',async () => {
  const mock = mockExcel();
  const nav = new Navigator(inspect);
  await nav.refresh('inputs');
  await nav.jump(nav.current.related[1]);
  await assert.rejects(nav.jump({sheetId:'Debt',address:'Z99'}), /Missing cell/);
  assert.equal(nav.canGoBack,true);
  assert.equal(nav.current.value,'100');
  mock.sheets.get('Cash Flow').getRange('B4').select();
  await nav.refresh('inputs');
  assert.equal(nav.canGoBack,false);
});

test('limits previews and propagates permission errors instead of reporting zero links',async () => {
  const mock = mockExcel();
  mock.inputs.set('Summary!B2',Array.from({length:101}, (_,i)=> mock.cell('Inputs','A'+(i+1),i)));
  const result = await inspect('inputs');
  assert.equal(result.related.length,101);
  assert.deepEqual(result.related[100],{sheetId:'Inputs',sheet:'Inputs',address:'A101',label:'A101',value:''});
  assert.equal(result.total,101);
  mock.active().getDirectPrecedents = () => {throw Object.assign(Error('Access denied'),{code:'AccessDenied'});};
  await assert.rejects(inspect('inputs'), /Access denied/);
});

test('preview activates the reference sheet, selects the whole range and scrolls to its first cell',async()=>{
  const calls=[];
  const first={load(){},text:[['25']],formulas:[['=A1+1']],left:340,top:710,width:60,height:16};
  const sheet={activate(){calls.push('activate');},getRange(address){
    assert.equal(address,'H40:H80');
    return {getCell(r,c){assert.deepEqual([r,c],[0,0]);return first;},select(){calls.push('select');}};
  }};
  globalThis.Office={context:{requirements:{isSetSupported:()=>true}}};
  globalThis.Excel={run:fn=>fn({workbook:{
    worksheets:{getItem(id){assert.equal(id,'debt-id');return sheet;}},
    application:{activeWindow:{scrollIntoView(...args){calls.push(args);}}}
  },sync:async()=>{calls.push('sync');}})};
  const {selectReference}=await import('../public/build/excel.js');
  assert.deepEqual(await selectReference({sheetId:'debt-id',address:'H40:H80'}),{value:'25',formula:'=A1+1'});
  assert.deepEqual(calls,['activate','select','sync',[340,710,60,16],'sync']);
});

test('returning a level remembers which cross-sheet sibling was followed',async()=>{
  const mock=mockExcel();
  const nav=new Navigator(inspect);
  await nav.refresh('inputs');
  await nav.jump(nav.current.related[1]);
  const via=await nav.back(false);
  assert.deepEqual(via,{sheetId:'Debt',address:'B4'});
  assert.equal(nav.current.sheet,'Summary');
  // The pane resumes its preview on Debt without briefly selecting Summary.
  assert.equal(mock.active().worksheet.id,'Debt');
});
