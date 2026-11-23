import { expression } from './expression.js';
const plain = (s) => s.replace(/\$/g, '').toUpperCase();
function coords(a) {
    const m = /^([A-Z]+)(\d+)$/.exec(plain(a));
    if (!m)
        return;
    return [Array.from(m[1]).reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0), Number(m[2])];
}
function contains(rect, cell) {
    const [a, b = a] = rect.split(':');
    const p = coords(a), q = coords(b), t = coords(cell);
    return !!(p && q && t && p.every((v, i) => t[i] >= v && t[i] <= q[i]));
}
export function traceTree(trace, mode) {
    const used = new Set();
    const individual = new Map();
    function related(i) {
        const r = trace.related[i];
        return { title: r.label, location: r.sheet + '!' + r.label, value: r.value,
            ref: { sheetId: r.sheetId, address: r.label } };
    }
    function describe(part) {
        const row = { title: part.text, info: part.role, part };
        if (part.children) {
            row.children = part.children.map(describe);
            return row;
        }
        const word = part.text;
        const bang = word.lastIndexOf('!');
        const sheet = bang < 0 ? trace.sheet : word.slice(0, bang).replace(/^'|'$/g, '').replace(/''/g, "'");
        const addr = plain(word.slice(bang + 1));
        const [start, end = start] = addr.split(':');
        const i = trace.related.findIndex(r => r.sheet.toUpperCase() === sheet.toUpperCase() &&
            (plain(r.label) === addr || (contains(r.label, start) && contains(r.label, end))));
        if (i >= 0) {
            const r = trace.related[i];
            if (plain(r.label) === addr)
                used.add(i);
            else if (start === end) {
                const members = individual.get(i) || new Set();
                members.add(addr);
                individual.set(i, members);
            }
            row.location = r.sheet + '!' + addr;
            row.ref = { sheetId: r.sheetId, address: addr };
            if (plain(r.label) === addr || (start === end && plain(r.address) === addr && !r.label.includes(':')))
                row.value = r.value;
        }
        else if (/^(?:\d+(?:\.\d*)?(?:E[+-]?\d+)?|TRUE|FALSE|"(?:[^"]|"")*")$/i.test(word))
            row.value = word;
        return row;
    }
    const root = { title: trace.address, location: trace.sheet + '!' + trace.address, value: trace.value,
        ref: { sheetId: trace.sheetId, address: trace.address }, children: [] };
    if (mode === 'inputs' && trace.formula)
        root.children.push(describe(expression(trace.formula)));
    trace.related.forEach((area, i) => {
        if (used.has(i))
            return;
        // If a name or a structured reference accounts for the rest of a merged area,
        // retain Excel's full block as a navigable result.
        const [a, b = a] = plain(area.label).split(':');
        const p = coords(a), q = coords(b);
        const size = p && q ? (q[0] - p[0] + 1) * (q[1] - p[1] + 1) : undefined;
        if (!individual.has(i) || individual.get(i).size !== size)
            root.children.push(related(i));
    });
    return root;
}
export function referenceRows(tree) {
    const rows = [];
    const seen = new Set();
    function visit(row) {
        if (row.ref) {
            const key = JSON.stringify([row.ref.sheetId, plain(row.ref.address)]);
            if (!seen.has(key)) {
                seen.add(key);
                rows.push(row);
            }
        }
        row.children?.forEach(visit);
    }
    tree.children?.forEach(visit);
    return rows;
}
