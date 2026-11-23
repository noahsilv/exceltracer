export type Mode = 'inputs' | 'dependents';
export interface CellRef { sheetId: string; address: string }
interface Related extends CellRef { sheet: string; label: string; value: string }
export interface Trace extends CellRef {
  sheet: string; formula: string; value: string; related: Related[]; total: number;
}

function localAddress(range: Excel.Range) {
  return range.address.slice(range.address.lastIndexOf('!') + 1);
}

function isMissingRelationship(error: unknown) {
  return typeof error === 'object' && error !== null && 'code' in error && error.code === 'ItemNotFound';
}

export async function inspect(mode: Mode, target?: CellRef, selectTarget = true): Promise<Trace> {
  return Excel.run(async context => {
    const workbook = context.workbook;
    const selection = target
      ? workbook.worksheets.getItem(target.sheetId).getRanges(target.address)
      : workbook.getSelectedRanges();
    selection.worksheet.load('id,name');
    selection.areas.load('items/address,items/rowCount,items/columnCount');
    await context.sync();

    const sources = selection.areas.items;
    const first = sources[0];
    const single = sources.length === 1 && first.rowCount === 1 && first.columnCount === 1;
    if (single) {
      first.load('text,formulas');
      await context.sync();
    }
    const expression = single ? first.formulas[0][0] : undefined;
    const trace: Trace = {
      sheetId: selection.worksheet.id, sheet: selection.worksheet.name,
      address: sources.map(localAddress).join(','),
      formula: typeof expression === 'string' && expression.startsWith('=') ? expression : '',
      value: single ? first.text[0][0] : sources.length === 1
        ? first.rowCount + ' × ' + first.columnCount : sources.length + ' selected areas',
      related: [], total: 0
    };
    const unique = new Map<string,Excel.Range>();
    for (const source of sources) {
      // Excel signals an empty result with ItemNotFound at sync. Isolate each query so
      // an unrelated selected cell cannot discard the other cells' dependencies.
      try {
        const links = mode === 'inputs' ? source.getDirectPrecedents() : source.getDirectDependents();
        links.ranges.load('items/address,items/rowCount,items/columnCount,items/worksheet/id,items/worksheet/name');
        await context.sync();
        for (const area of links.ranges.items) unique.set(area.address, area);
      } catch (error) {
        if (!isMissingRelationship(error)) throw error;
      }
    }
    const areas = [...unique.values()];
    trace.total = areas.length;
    // Keep large blocks intact. Fetch display text only for the first cell of each visible area.
    const visible = areas.slice(0, 100);
    const previews = visible.map(area => {
      const firstCell = area.getCell(0, 0);
      firstCell.load('address,text');
      return firstCell;
    });
    if (previews.length) await context.sync();
    trace.related = areas.map((area, index) => {
      const preview = previews[index];
      const isBlock = area.rowCount > 1 || area.columnCount > 1;
      const label = localAddress(area);
      const leading = label.split(':')[0];
      const address = preview ? localAddress(preview) : /^\$?[A-Z]+$/i.test(leading) ? leading+'1'
        : /^\$?\d+$/.test(leading) ? 'A'+leading : leading;
      const value = preview?.text[0][0] || '';
      return {
        sheetId:area.worksheet.id, sheet:area.worksheet.name, address, label,
        value:preview && isBlock ? 'First cell '+address+' · '+value : value
      };
    });
    if (target && selectTarget) {
      selection.worksheet.activate();
      selection.select();
      await context.sync();
    }
    return trace;
  });
}

export async function selectReference(ref: CellRef) {
  return Excel.run(async ctx => {
    const sheet = ctx.workbook.worksheets.getItem(ref.sheetId);
    const [leading, second] = ref.address.split(',', 2);
    const leadingRange = sheet.getRange(leading);
    const range = second ? sheet.getRanges(ref.address) : leadingRange;
    const first = leadingRange.getCell(0, 0);
    const canScroll = Office.context.requirements.isSetSupported('ExcelApiDesktop', '1.1');
    first.load(canScroll ? 'text,formulas,left,top,width,height' : 'text,formulas');
    sheet.activate();
    range.select();
    await ctx.sync();
    if (canScroll) {
      // Keep the full block selected, but scroll to its leading cell even for whole-column references.
      try {
        ctx.workbook.application.activeWindow.scrollIntoView(first.left, first.top, first.width, first.height);
        await ctx.sync();
      } catch {
        // The selection already succeeded. Some desktop windows cannot scroll (for example, a frozen pane).
      }
    }
    const expression = first.formulas[0][0];
    return { value:first.text[0][0], formula:typeof expression === 'string' && expression.startsWith('=') ? expression : '' };
  });
}
