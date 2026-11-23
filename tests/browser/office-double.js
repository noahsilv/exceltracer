// Test-only Office host. Selecting a cell requires its sheet to be active and blurs the pane.
window.commands = {};
window.active = 'Summary!B2';
window.activeSheet = 'Summary';
window.visible = true;
const sample = {
  'Summary!B2': ["='Cash Flow'!B4/Debt!B4", '1.50x', ['Cash Flow!B4', 'Debt!B4']],
  'Cash Flow!B4': ['=B2-B3', '150', ['Cash Flow!B2', 'Cash Flow!B3']],
  'Debt!B4': ['=B2+B3', '100', ['Debt!B2', 'Debt!B3']],
  'Debt!B2': [25, '25', []],
  'Debt!B3': [75, '75', []],
  'Cash Flow!B2': [200, '200', []],
  'Cash Flow!B3': [50, '50', []],
  'Cash Flow!B7': ['=Summary!B2-1', '0.50x', ['Summary!B2']],
  'Debt!B7': ['=IF(Summary!B2>=1,"Pass","Fail")', 'Pass', ['Summary!B2']]
};
let empty = false;

function sheet(id) {
  return {
    id, name: id, load() {},
    activate() { window.activeSheet = id; },
    getRange: address => range(id + '!' + address),
    getRanges: address => group(address.split(',').map(a => range(id + '!' + a)))
  };
}

function group(items) {
  return {
    worksheet: items[0].worksheet,
    areas: {items, load() {}},
    select() { items[0].select(); }
  };
}

function related(keys) {
  if (!keys.length) empty = true;
  return {ranges: {items: keys.map(range), load() {}}};
}

function range(key) {
  const [name, address] = key.split('!');
  const [formula, value, precedents] = sample[key];
  return {
    worksheet: sheet(name), address: "'" + name + "'!" + address,
    text: [[value]], formulas: [[formula]], rowCount: 1, columnCount: 1,
    left: 50, top: 35, width: 45, height: 18, load() {}, getCell: () => range(key),
    select() {
      if (window.activeSheet !== name) throw Error('Wrong sheet active');
      window.active = key;
      document.activeElement?.blur();
    },
    getDirectPrecedents: () => related(precedents),
    getDirectDependents: () => related(Object.keys(sample).filter(k => sample[k][2].includes(key)))
  };
}

window.Office = {
  HostType: {Excel: 'Excel'},
  context: {requirements: {isSetSupported: () => true}},
  actions: {associate(id, fn) { window.commands[id] = fn; }},
  addin: {
    async showAsTaskpane() { window.visible = true; },
    async hide() { window.visible = false; }
  },
  onReady: fn => fn({host: 'Excel'})
};
window.Excel = {
  run: fn => fn({
    workbook: {
      getSelectedRanges: () => group([range(window.active)]),
      worksheets: {getItem: sheet},
      application: {activeWindow: {scrollIntoView() {}}}
    },
    async sync() {
      await new Promise(resolve => setTimeout(resolve, 2));
      if (empty) {
        empty = false;
        throw Object.assign(Error('No results'), {code: 'ItemNotFound'});
      }
    }
  })
};
