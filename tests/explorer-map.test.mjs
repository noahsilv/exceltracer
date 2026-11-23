import assert from 'node:assert/strict';
import test from 'node:test';
import {mapKind, palette, toggleMap} from '../public/build/map.js';
import {expression} from '../public/build/expression.js';
import {traceTree} from '../public/build/rows.js';

test('map classifications use relative formula patterns and quoted strings',()=>{
  const f='=RC[-1]*R2C';
  assert.equal(mapKind(f,f,'','Double'),'across');
  assert.equal(mapKind(f,'',f,'Double'),'down');
  assert.equal(mapKind(f,f,f,'Double'),'both');
  assert.equal(mapKind(f,'','','Double'),'unique');
  assert.equal(mapKind("='[Prior.xlsx]Budget'!RC",'','','Double'),'external');
  assert.equal(mapKind("='Debt'!RC",'','','Double'),'sheet');
  assert.equal(mapKind("=R[-1]C-'Debt'!RC",'','','Double'),'sheet');
  assert.equal(mapKind("=Table1[Sales]-Debt!RC",'','','Double'),'sheet');
  assert.equal(mapKind("=RC-[Prior.xlsx]Debt!RC",'','','Double'),'external');
  assert.equal(mapKind('="Book!X"','','','String'),'constant');
  assert.equal(mapKind('','','','Empty'),undefined);
});

test('tree retains formula spans, nested IF argument roles and quoted punctuation',()=>{
  const formula='=IF(B5>0,SUM(\'O\'\'Brien, Inc.\'!$B$2,C3)/B5,"a,b")';
  const part=expression(formula);
  assert.equal(part.text,'IF');
  assert.equal(part.children[1].role,'When true');
  assert.equal(part.children[1].text,'/');
  const ref=part.children[1].children[0].children[0];
  assert.equal(formula.slice(ref.start,ref.end),"'O''Brien, Inc.'!$B$2");
  assert.equal(part.children[2].text,'"a,b"');
  assert.equal(expression('=IF(B2,,3)').children[1].text,'(omitted)');
});

test('references inside a merged dependency area point to each actual cell',()=>{
  const t={sheetId:'s',sheet:'Debt',address:'D2',value:'100',formula:'=B2+B3',total:1,
    related:[{sheet:'Debt',sheetId:'s',address:'B2',label:'B2:B3',value:'block'}]};
  const root=traceTree(t,'inputs');
  assert.equal(root.children.length,1);
  assert.deepEqual(root.children[0].children.map(r=>r.ref.address),['B2','B3']);
});

test('two toggles remove only our rules and skip disabled colors', async()=>{
  const prior={type:'Custom',custom:{rule:{formula:'=B2>0',load(){}}},delete(){assert.fail('User rule removed');}};
  const all=[prior], applied=[];
  const mockSheet={
    name:'Map example',load(){},getRange(){return {conditionalFormats:{items:all,load(){}}};},
    getUsedRangeOrNullObject(){return {rowIndex:2,columnIndex:1,rowCount:2,columnCount:4,load(){}};},
    getRangeByIndexes(){return {load(){},formulasR1C1:[['=RC[-1]','=RC[-1]',42,'=RC[-1]'],['=RC[-1]','=RC[-1]','','=RC[-1]']],
      valueTypes:[['Double','Double','Double','Double'],['Double','Double','Empty','Double']],
      getSpecialCellsOrNullObject(){return {areas:{load(){},items:[{rowIndex:2,columnIndex:1,rowCount:2,columnCount:2},{rowIndex:2,columnIndex:4,rowCount:2,columnCount:1}]}};}
    };},
    getRanges(addr){return {conditionalFormats:{add(){const f={type:'Custom',custom:{rule:{load(){}},format:{fill:{}}},delete(){all.splice(all.indexOf(f),1);}};all.push(f);applied.push([addr,f]);return f;}}};}
  };
  globalThis.Excel={run:fn=>fn({workbook:{worksheets:{getActiveWorksheet:()=>mockSheet}},sync:async()=>{}})};
  assert.match(await toggleMap({number:''}),/Mapped 6 cells/);
  assert.equal(applied.find(([addr])=>addr==='C4')[1].custom.format.fill.color,palette.both[1]);
  assert.match(await toggleMap(),/removed/);
  assert.deepEqual(all,[prior]);
});

test('a partially resolved formula still exposes the entire Excel-reported dependency block',async()=>{
  const {referenceRows}=await import('../public/build/rows.js');
  const root=traceTree({sheetId:'s',sheet:'Debt',address:'D2',value:'100',formula:'=B2+OtherInputs',total:1,
    related:[{sheet:'Debt',sheetId:'s',address:'B2',label:'B2:B7',value:'25'}]},'inputs');
  assert.deepEqual(referenceRows(root).map(row=>row.ref.address),['B2','B2:B7']);
});
