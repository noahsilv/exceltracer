export const palette = {
    across: ['Copied across', '#B7CDE7'], down: ['Copied down', '#D9E5B3'],
    both: ['Copied across and down', '#CDC5AA'], unique: ['Unique formula', '#FFC75C'],
    sheet: ['Other worksheet', '#FFAB70'], external: ['Other workbook', '#F09BB1'],
    constant: ['Formula with only constants', '#E3CAF4'], number: ['Number or logical value', '#FFF98A'],
    text: ['Text label', '#E3EFF2']
};
export function readColors() {
    try {
        const stored = JSON.parse(localStorage.getItem('formulaTrace.mapColors') || '{}');
        return Object.fromEntries(Object.entries(stored).filter(([k, v]) => Object.prototype.hasOwnProperty.call(palette, k) &&
            (v === '' || (typeof v === 'string' && /^#[0-9a-f]{6}$/i.test(v)))));
    }
    catch {
        return {};
    }
}
// A constant, identifiable rule survives save/reopen without changing base fills.
const marker = '="FormulaTrace.Map.030c881e"="FormulaTrace.Map.030c881e"';
export function mapKind(formula, left, above, valueType) {
    if (!formula) {
        if (valueType === 'Empty')
            return;
        return valueType === 'String' ? 'text' : 'number';
    }
    if (formula === left)
        return formula === above ? 'both' : 'across';
    if (formula === above)
        return 'down';
    const syntax = formula.replace(/"(?:[^"]|"")*"/g, '""');
    if (/\[[^\]]+\][^!+*/(),;]*!/.test(syntax))
        return 'external';
    if (syntax.includes('!'))
        return 'sheet';
    // Only positively identify expressions made up of literals and operators.
    if (/^(?:[=+\-*/^%&<>().,\s\d]|""|TRUE\b|FALSE\b|\dE[+-]?\d)+$/i.test(syntax))
        return 'constant';
    return 'unique';
}
function column(index) {
    let name = '';
    for (let n = index + 1; n; n = Math.floor((n - 1) / 26))
        name = String.fromCharCode(65 + (n - 1) % 26) + name;
    return name;
}
async function removeMap(sheet, ctx) {
    const formats = sheet.getRange().conditionalFormats;
    formats.load('items/type');
    await ctx.sync();
    const custom = formats.items.filter(item => item.type === 'Custom');
    for (const format of custom)
        format.custom.rule.load('formula');
    await ctx.sync();
    const ours = custom.filter(format => format.custom.rule.formula.replace(/^=?/, '=') === marker);
    for (const format of ours)
        format.delete();
    if (ours.length)
        await ctx.sync();
    return ours.length;
}
export async function toggleMap(colors = {}) {
    return Excel.run(async (ctx) => {
        const sheet = ctx.workbook.worksheets.getActiveWorksheet();
        sheet.load('name');
        if (await removeMap(sheet, ctx))
            return 'Map removed from ' + sheet.name + '.';
        const used = sheet.getUsedRangeOrNullObject(true);
        used.load('rowIndex,columnIndex,rowCount,columnCount');
        await ctx.sync();
        if (used.isNullObject)
            return sheet.name + ' is empty.';
        const width = used.columnCount;
        const stride = Math.max(1, Math.floor(12000 / width));
        let above = [];
        let colored = 0;
        try {
            for (let offset = 0; offset < used.rowCount; offset += stride) {
                const start = used.rowIndex + offset;
                const height = Math.min(stride, used.rowCount - offset);
                const tile = sheet.getRangeByIndexes(start, used.columnIndex, height, width);
                tile.load('formulasR1C1,valueTypes');
                const formulas = tile.getSpecialCellsOrNullObject('Formulas');
                await ctx.sync();
                const mask = Array.from({ length: height }, () => new Uint8Array(width));
                if (!formulas.isNullObject) {
                    formulas.areas.load('items/rowIndex,items/columnIndex,items/rowCount,items/columnCount');
                    await ctx.sync();
                    for (const area of formulas.areas.items) {
                        const first = Math.max(0, area.columnIndex - used.columnIndex);
                        const end = Math.min(width, area.columnIndex + area.columnCount - used.columnIndex);
                        for (let r = Math.max(0, area.rowIndex - start); r < Math.min(height, area.rowIndex + area.rowCount - start); r++)
                            mask[r].fill(1, first, end);
                    }
                }
                const groups = new Map();
                for (let r = 0; r < height; r++) {
                    const row = tile.formulasR1C1[r].map((v, c) => mask[r][c] ? String(v) : '');
                    const kinds = row.map((v, c) => mapKind(v, row[c - 1], above[c], tile.valueTypes[r][c]));
                    for (let c = 0; c < width;) {
                        const kind = kinds[c];
                        const begin = c++;
                        while (c < width && kinds[c] === kind)
                            c++;
                        if (!kind || colors[kind] === '')
                            continue;
                        const num = String(start + r + 1);
                        const from = column(used.columnIndex + begin) + num;
                        const to = column(used.columnIndex + c - 1) + num;
                        if (!groups.has(kind))
                            groups.set(kind, []);
                        groups.get(kind).push(from === to ? from : from + ':' + to);
                        colored += c - begin;
                    }
                    above = row;
                }
                for (const [kind, runs] of groups) {
                    // Bound each multi-area address below Excel's address-length limit.
                    for (let i = 0; i < runs.length; i += 128) {
                        const rule = sheet.getRanges(runs.slice(i, i + 128).join(',')).conditionalFormats.add('Custom');
                        rule.custom.rule.formula = marker;
                        rule.stopIfTrue = false;
                        rule.custom.format.fill.color = colors[kind] || palette[kind][1];
                    }
                }
                await ctx.sync();
            }
        }
        catch (error) {
            await removeMap(sheet, ctx);
            throw error;
        }
        if (!colored)
            return 'No cells match the enabled map colors.';
        return 'Mapped ' + colored.toLocaleString() + ' cells on ' + sheet.name + '.';
    });
}
