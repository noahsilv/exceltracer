export interface Part { text: string; role?: string; children?: Part[]; start: number; end: number }
const roles: Record<string,string[]> = {
  IF:['Condition','When true','When false'], SUMIF:['Range','Criteria','Sum range'],
  INDEX:['Reference','Row','Column','Area'], OFFSET:['Reference','Rows','Columns','Height','Width'],
  XLOOKUP:['Lookup value','Lookup array','Return array','If missing','Match mode','Search mode'],
  VLOOKUP:['Lookup value','Table','Column','Match mode']
};

// Split only at operators outside strings, sheet names, tables, arrays and parentheses.
// Unsupported syntax remains selectable as the original source text.
export function expression(formula: string): Part {
  const body = formula.startsWith('=') ? 1 : 0;
  function read(lo: number, hi: number, depth: number): Part {
    while (lo < hi && /\s/.test(formula[lo])) lo++;
    while (hi > lo && /\s/.test(formula[hi-1])) hi--;
    const raw = formula.slice(lo,hi);
    const leaf: Part = {text:raw || '(omitted)', start:lo, end:hi};
    if (depth > 45 || !raw) return leaf;
    let quote = '', bracket = 0, braces = 0, paren = 0, open = -1, close = -1;
    let split = -1, size = 0, priority = 99;
    const commas: number[] = [];
    for (let i = lo; i < hi; i++) {
      const ch = formula[i];
      if (quote) {
        if (ch === quote) { if (formula[i+1] === ch) i++; else quote = ''; }
        continue;
      }
      if (ch === '"' || ch === "'") { quote = ch; continue; }
      if (ch === '[') bracket++;
      if (ch === ']') bracket--;
      if (ch === '{') braces++;
      if (ch === '}') braces--;
      if (bracket || braces) continue;
      if (ch === '(') { if (paren++ === 0 && open < 0) open = i; continue; }
      if (ch === ')') { if (--paren === 0 && close < 0) close = i; continue; }
      if (paren === 1 && ch === ',') commas.push(i);
      if (paren !== 0) continue;
      const op = /^(<>|<=|>=|[=<>&+*/^-])/.exec(formula.slice(i,i+2))?.[0];
      if (!op) continue;
      if ((op === '+' || op === '-') && (i === lo || /[=(,+*/^<>&-]\s*$/.test(formula.slice(lo,i)) ||
          /(?:^|[^\w.])\d+(?:\.\d*)?E$/i.test(formula.slice(lo,i)))) continue;
      const p = /^[<>=]/.test(op) ? 1 : op === '&' ? 2 : /[+-]/.test(op) ? 3 : /[*\/]/.test(op) ? 4 : 5;
      if (p <= priority) { split = i; size = op.length; priority = p; }
      i += op.length-1;
    }
    if (quote || paren || bracket || braces) return leaf;
    if (split >= 0) return {...leaf, text:formula.slice(split,split+size), children:[
      read(lo,split,depth+1),read(split+size,hi,depth+1)
    ]};
    const name = formula.slice(lo,open);
    if (open >= lo && close === hi-1 && (!name || /^[\w.]+$/.test(name))) {
      if (!name) return {...leaf, children:[read(open+1,close,depth+1)]};
      const ends = [...commas,close];
      const children = close === open+1 ? [] : ends.map((end,i) => ({
        ...read(i ? ends[i-1]+1 : open+1,end,depth+1), role:roles[name.toUpperCase()]?.[i] || 'Argument ' + (i+1)
      }));
      return {...leaf,text:name,children};
    }
    return leaf;
  }
  return read(body,formula.length,0);
}
