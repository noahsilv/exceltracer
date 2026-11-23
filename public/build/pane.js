import { inspect, selectReference } from './excel.js';
import { Navigator } from './navigation.js';
import { toggleMap, palette, readColors } from './map.js';
import { traceTree, referenceRows } from './rows.js';
const nav = new Navigator(inspect);
const byId = (id) => document.getElementById(id);
const button = (id) => byId(id);
let tree;
let selected;
let collapsed = new Set();
let flat = [];
let closedRoot;
const colors = readColors();
let pending = Promise.resolve();
let waiting = 0;
const pageSize = 100;
let pageStart = 0;
const treeSwitch = () => byId('expression-view');
const inTree = () => nav.mode === 'inputs' && treeSwitch().checked;
function focusRow(row) {
    if (!document.hasFocus())
        window.focus();
    row.focus({ preventScroll: true });
    row.scrollIntoView({ block: 'nearest', inline: 'nearest' });
}
function queue(task) {
    waiting++;
    pending = pending.then(task).catch(error => {
        byId('status').textContent = error instanceof Error ? error.message : String(error);
    }).finally(() => { waiting--; });
    return pending;
}
function drawFormula() {
    const formula = nav.current?.formula || '';
    const node = byId('formula');
    node.replaceChildren();
    const p = selected?.part;
    if (!p) {
        node.textContent = formula || nav.current?.value || '';
        return;
    }
    const mark = document.createElement('mark');
    mark.textContent = formula.slice(p.start, p.end);
    node.append(formula.slice(0, p.start), mark, formula.slice(p.end));
}
function render(focus) {
    if (!tree)
        return;
    flat = [];
    function visit(row, depth, parent) {
        flat.push({ row, depth, parent });
        if (!collapsed.has(row))
            row.children?.forEach(child => visit(child, depth + 1, row));
    }
    if (inTree())
        visit(tree, 0);
    else
        flat = [{ row: tree, depth: 0 }, ...referenceRows(tree).map(row => ({ row, depth: 1, parent: tree }))];
    treeSwitch().disabled = nav.mode !== 'inputs';
    byId('scroll').dataset.direction = nav.mode;
    const body = byId('rows');
    body.replaceChildren();
    let target;
    const index = Math.max(0, flat.findIndex(entry => entry.row === selected));
    pageStart = Math.floor(index / pageSize) * pageSize;
    byId('pages').hidden = flat.length <= pageSize;
    button('prev-page').disabled = pageStart === 0;
    button('next-page').disabled = pageStart + pageSize >= flat.length;
    for (const { row, depth } of flat.slice(pageStart, pageStart + pageSize)) {
        const tr = document.createElement('tr');
        tr.tabIndex = row === selected ? 0 : -1;
        tr.setAttribute('aria-current', String(row === selected));
        if (row === selected)
            target = tr;
        const cell = tr.insertCell();
        cell.style.paddingLeft = (4 + depth * 12) + 'px';
        const twist = document.createElement('span');
        twist.className = 'twisty';
        if (inTree() && row.children?.length) {
            twist.textContent = collapsed.has(row) ? '▸' : '▾';
            twist.onclick = e => { e.stopPropagation(); void queue(async () => { await preview(row); toggle(row); }); };
        }
        cell.append(twist, row.title);
        for (const field of [row.info, row.value, row.location])
            tr.insertCell().textContent = field || '';
        tr.onclick = () => void queue(() => preview(row));
        tr.ondblclick = () => void queue(() => drill(row, nav.mode));
        body.append(tr);
    }
    if (focus && target)
        focusRow(target);
    button('back').disabled = !nav.canGoBack;
    button('inputs').setAttribute('aria-pressed', String(nav.mode === 'inputs'));
    button('dependents').setAttribute('aria-pressed', String(nav.mode === 'dependents'));
    describeStep();
    drawFormula();
}
function setTrace(focus = true) {
    if (!nav.current)
        return;
    tree = traceTree(nav.current, nav.mode);
    selected = tree;
    collapsed = new Set();
    showTraceCount();
    render(focus);
}
function showTraceCount() {
    const count = nav.current?.total || 0;
    byId('status').textContent = count ? count + ' direct reference ' + (count === 1 ? 'area.' : 'areas.')
        : 'Excel found no direct ' + nav.mode + ' in this workbook.';
}
function describeStep() {
    const i = flat.findIndex(item => item.row === selected);
    const label = nav.mode === 'inputs' ? 'Precedent' : 'Dependent';
    byId('step').textContent = selected === tree
        ? 'Exploring ' + (tree?.location || '')
        : (inTree() ? 'Component' : label + ' ' + i + ' of ' + (flat.length - 1)) +
            ' · ' + (selected?.location || selected?.title || '');
    if (selected === tree || !selected?.ref)
        byId('target-formula').textContent = '';
}
async function preview(row) {
    byId('target-formula').textContent = '';
    showTraceCount();
    if (row.ref) {
        try {
            const result = await selectReference(row.ref);
            if (row !== tree || !/[,:]/.test(row.ref.address))
                row.value = result.value;
            if (row !== tree && result.formula) {
                byId('target-formula').textContent = (row.ref.address.includes(':') ? 'First cell: ' : '') + result.formula;
            }
        }
        catch (error) {
            // Keep the list navigable when Excel cannot select a hidden or protected reference.
            const detail = error instanceof Error ? error.message : String(error);
            byId('status').textContent = 'Excel could not select ' + row.location + '. ' + detail;
        }
    }
    selected = row;
    const index = flat.findIndex(item => item.row === row);
    if (index < pageStart || index >= pageStart + pageSize) {
        render(true);
        return;
    }
    const elements = byId('rows').querySelectorAll('tr');
    flat.slice(pageStart, pageStart + pageSize).forEach(({ row: item }, i) => {
        elements[i].tabIndex = item === row ? 0 : -1;
        elements[i].setAttribute('aria-current', String(item === row));
        if (item === row) {
            elements[i].cells[2].textContent = row.value || '';
            focusRow(elements[i]);
        }
    });
    describeStep();
    drawFormula();
}
function toggle(row) {
    if (collapsed.has(row))
        collapsed.delete(row);
    else
        collapsed.add(row);
    render(true);
}
async function drill(row, mode) {
    if (!row?.ref) {
        if (row?.children?.length)
            toggle(row);
        return;
    }
    const ref = { sheetId: row.ref.sheetId, address: row.ref.address };
    await nav.jump(ref, mode);
    setTrace();
}
async function fromSheet(mode) {
    await nav.refresh(mode);
    await Office.addin.showAsTaskpane();
    closedRoot = undefined;
    setTrace();
}
async function back(resumeStep = false) {
    if (closedRoot) {
        await selectReference(closedRoot);
        closedRoot = undefined;
        return;
    }
    if (nav.canGoBack) {
        const via = await nav.back(!resumeStep);
        setTrace(!resumeStep);
        if (resumeStep && tree) {
            const prior = flat.find(({ row }) => row !== tree && row.ref?.sheetId === via?.sheetId &&
                row.ref?.address === via?.address);
            await preview(prior?.row || tree);
        }
    }
    else if (tree)
        await preview(tree);
}
async function close(restore) {
    if (tree?.ref && restore)
        await selectReference(tree.ref);
    closedRoot = tree?.ref;
    await Office.addin.hide();
}
async function mapSheet() {
    byId('status').textContent = 'Checking formula map…';
    try {
        byId('status').textContent = await toggleMap(colors);
    }
    catch (error) {
        await Office.addin.showAsTaskpane();
        throw error;
    }
}
function move(key) {
    const i = flat.findIndex(x => x.row === selected);
    if (i < 0)
        return;
    const { row, parent } = flat[i];
    if (!inTree()) {
        if (key === 'ArrowLeft')
            return nav.canGoBack ? back(true) : tree && preview(tree);
        if (key === 'ArrowRight') {
            if (row === tree)
                return flat[1] && preview(flat[1].row);
            return drill(row, nav.mode);
        }
    }
    if (key === 'ArrowLeft' && row.children?.length && !collapsed.has(row)) {
        toggle(row);
        return;
    }
    if (key === 'ArrowRight' && row.children?.length && collapsed.has(row)) {
        toggle(row);
        return;
    }
    const next = key === 'Home' ? tree : key === 'ArrowUp' ? flat[i - 1]?.row :
        key === 'ArrowDown' ? flat[i + 1]?.row : key === 'ArrowLeft' ? parent : row.children?.[0];
    if (next)
        return preview(next);
}
document.addEventListener('keydown', e => {
    if (e.metaKey || e.altKey || (e.target instanceof Element && e.target.closest('input,textarea,select')))
        return;
    const ctrl = e.ctrlKey;
    const key = e.key.toLowerCase();
    let action;
    if (ctrl && key === 'q')
        action = () => drill(selected, e.shiftKey ? 'dependents' : 'inputs');
    else if (ctrl && e.shiftKey && key === 'm')
        action = mapSheet;
    else if (ctrl && key === 'backspace')
        action = back;
    else if (!ctrl && key === 'escape')
        action = () => close(true);
    else if (!ctrl && e.target instanceof Element && (e.target.closest('#rows') || e.target === document.body)) {
        if (key === 'enter')
            action = () => close(false);
        else if (key === 'home' || e.key.startsWith('Arrow'))
            action = () => move(e.key);
    }
    if (action) {
        e.preventDefault();
        e.stopPropagation();
        if (e.repeat && waiting && e.key.startsWith('Arrow'))
            return;
        void queue(async () => { await action(); });
    }
});
Office.onReady(async (info) => {
    if (info.host !== Office.HostType.Excel) {
        byId('status').textContent = 'Open Formula Trace in Excel.';
        return;
    }
    const commands = {
        TraceInputs: () => fromSheet('inputs'), TraceUsers: () => fromSheet('dependents'),
        ToggleMap: mapSheet, ReturnCell: back
    };
    for (const [name, action] of Object.entries(commands))
        Office.actions.associate(name, () => queue(action));
    Office.actions.associate('MapFromRibbon', (event) => {
        void queue(mapSheet).finally(() => event.completed());
    });
    for (const mode of ['inputs', 'dependents'])
        button(mode).onclick = () => void queue(() => fromSheet(mode));
    button('back').onclick = () => void queue(back);
    button('close').onclick = () => void queue(() => close(false));
    button('prev-page').onclick = () => void queue(() => preview(flat[Math.max(0, pageStart - pageSize)].row));
    button('next-page').onclick = () => void queue(() => preview(flat[Math.min(flat.length - 1, pageStart + pageSize)].row));
    treeSwitch().onchange = () => void queue(async () => {
        if (!tree)
            return;
        const keep = selected?.ref;
        selected = tree;
        render(false);
        const match = keep && flat.find(({ row }) => row.ref?.sheetId === keep.sheetId && row.ref?.address === keep.address);
        await preview(match?.row || tree);
    });
    button('map').onclick = () => void queue(mapSheet);
    for (const kind of Object.keys(palette)) {
        const [label, defaultColor] = palette[kind];
        const enabled = document.createElement('input');
        enabled.type = 'checkbox';
        enabled.checked = colors[kind] !== '';
        enabled.setAttribute('aria-label', 'Map ' + label);
        const color = document.createElement('input');
        color.type = 'color';
        color.value = colors[kind] || defaultColor;
        color.setAttribute('aria-label', label + ' color');
        const save = () => {
            colors[kind] = enabled.checked ? color.value : '';
            try {
                localStorage.setItem('formulaTrace.mapColors', JSON.stringify(colors));
            }
            catch {
                byId('status').textContent = 'Colors will be used for this session.';
            }
        };
        color.onchange = save;
        enabled.onchange = save;
        const entry = document.createElement('label');
        entry.append(enabled, color, label);
        byId('colors').append(entry);
    }
    if (Office.context.requirements.isSetSupported('KeyboardShortcuts', '1.1')) {
        const form = byId('bindings');
        const details = byId('keys');
        details.ontoggle = () => {
            if (!details.open || !form.hidden)
                return;
            void queue(async () => {
                const hotkeys = await Office.actions.getShortcuts();
                const labels = { TraceInputs: 'Precedents', TraceUsers: 'Dependents', ToggleMap: 'Formula map', ReturnCell: 'Return' };
                form.replaceChildren();
                for (const [id, label] of Object.entries(labels)) {
                    const field = document.createElement('input');
                    field.name = id;
                    field.value = hotkeys[id] || '';
                    field.placeholder = 'Unassigned';
                    const item = document.createElement('label');
                    item.append(label, field);
                    form.append(item);
                }
                const save = document.createElement('button');
                save.type = 'submit';
                save.textContent = 'Save worksheet shortcuts';
                form.append(save);
                form.hidden = false;
            });
        };
        form.onsubmit = event => {
            event.preventDefault();
            void queue(async () => {
                const keys = {};
                form.querySelectorAll('input').forEach(input => { if (input.value.trim())
                    keys[input.name] = input.value.trim(); });
                await Office.actions.replaceShortcuts(keys);
                byId('status').textContent = 'Worksheet shortcuts saved in Excel.';
            });
        };
    }
    await queue(async () => { await nav.refresh('inputs'); setTrace(false); });
});
