// The MCP tools: what the model can do with workbooks and databases, plus the few the page alone may call.
// `_meta.ui.visibility: ['app']` hides a tool from the model; `readOnlyHint` lets Zod run it without asking.
import { existsSync, readFileSync, writeFileSync } from 'node:fs'
import { homedir } from 'node:os'
import { extname, resolve } from 'node:path'
import { Book, FUNCTION_NAMES, XlErr, addr, colIndex, mapRefs, parseAddr, parseInput, parseRange, rangeText, rawText } from './index.js'
import { cellsFromRows, parseCSV, readXlsx, rowsOf, toCSV, writeXlsx } from './io.js'
import * as db from './db.js'
import * as store from './store.js'

export const UI_URI = 'ui://sheets/app'
const MAX_CELLS = 100000
const text = (t, structured) => ({ content: [{ type: 'text', text: t }], ...(structured ? { structuredContent: structured } : {}) })
const S = (props, required = []) => ({ type: 'object', properties: props, required })
const str = (description) => ({ type: 'string', description })
const WB = str('Workbook id or name. Leave out for the one used last.')
const SH = str('Sheet name or id. Leave out for the first sheet.')
const RANGE = str('A1 notation: A1, A1:C10, B:B, 3:5. May start with a sheet name, like Budget!A1:B9.')
const own = (o, k) => Object.hasOwn(o || {}, k)
const noFiles = () => {
  if (process.env.SHEETS_NO_FILES) throw new Error('this server does not read or write files on its own disk')
}
const path = (p) => resolve(String(p).replace(/^~(?=\/|$)/, homedir()))

// ── helpers ──────────────────────────────────────────────────────────────────────────────────────────────────────────────
function where(a, { needRange = false } = {}) {
  const book = store.get(a.workbook)
  let sheet = a.sheet
  let range = a.range ?? null
  const m = typeof range === 'string' && range.match(/^(?:'((?:[^']|'')+)'|([^'!]+))!(.+)$/)
  if (m) {
    sheet = (m[1] ?? m[2]).replace(/''/g, "'")
    range = m[3]
  }
  const sh = book.sheet(sheet === undefined || sheet === null || sheet === '' ? 0 : sheet)
  if (needRange && !range) throw new Error('a range is needed, like A1:C10')
  let g = null
  if (range) {
    g = parseRange(range)
    if (!g) throw new Error(`"${range}" is not a range. Use A1 notation like A1:C10.`)
  }
  return { book, sh, g }
}
const trimmed = (book, sh, g) => {
  // a whole-column or whole-row range is cut to the used part
  const used = book.usedRange(sh.id)
  if (!used) return { r1: g.r1, c1: g.c1, r2: g.r1, c2: g.c1 }
  return { r1: g.r1, c1: g.c1, r2: Math.min(g.r2, used.r2), c2: Math.min(g.c2, used.c2) }
}
const commit = (book, ops) => {
  const inv = book.applyAll(ops)
  book.recalc()
  store.save(book)
  return inv
}
const show = (v) => (v instanceof XlErr ? v.code : v === null || v === undefined ? '' : v)
const grid = (rows) => rows.map((r) => r.map((v) => String(show(v)).replace(/\t|\n/g, ' ')).join('\t')).join('\n')
function errorsOf(book, sh, g) {
  const out = []
  const i = book.sheetIndex(sh.id)
  for (let r = g.r1; r <= g.r2 && out.length < 20; r++)
    for (let c = g.c1; c <= g.c2 && out.length < 20; c++) {
      const v = book.get(i, r, c)
      if (v instanceof XlErr) out.push(`${addr(r, c)} ${v.code}`)
    }
  return out
}
function setCells(book, sh, entries) {
  // entries: { A1: input }, input a number, text, boolean, null (clear) or a formula text. Formats stay.
  const cells = {}
  let n = 0
  for (const [a1, input] of Object.entries(entries)) {
    const p = parseAddr(a1)
    if (!p) throw new Error(`"${a1}" is not a cell address`)
    if (++n > MAX_CELLS) throw new Error(`that is more than ${MAX_CELLS} cells at once`)
    const old = sh.cells.get(p.r * 16384 + p.c)
    const parsed = parseInput(input)
    const s = { ...(old?.s || {}) }
    if (parsed.nf && !s.nf) s.nf = parsed.nf
    cells[addr(p.r, p.c)] = parsed.cell ? { ...parsed.cell, ...(Object.keys(s).length ? { s } : {}) } : Object.keys(s).length ? { s } : null
  }
  return { op: 'set', sheet: sh.id, cells }
}
const styleOf = (s) => {
  const out = {}
  const keys = ['b', 'i', 'u', 'st', 'fg', 'bg', 'al', 'va', 'wrap', 'nf', 'fs', 'ff', 'bd', 'bc', 'ind']
  for (const k of keys) if (own(s, k) && s[k] !== null && s[k] !== '') out[k] = s[k]
  const alias = { bold: 'b', italic: 'i', underline: 'u', strike: 'st', color: 'fg', fill: 'bg', background: 'bg', align: 'al', number_format: 'nf', font_size: 'fs', font: 'ff', wrap_text: 'wrap', vertical: 'va' }
  for (const [k, to] of Object.entries(alias)) if (own(s, k) && s[k] !== null && s[k] !== '') out[to] = s[k]
  if (own(s, 'borders')) {
    const b = s.borders
    out.bd = b === 'all' ? { l: 'thin', r: 'thin', t: 'thin', b: 'thin' } : b === 'none' ? null : b
  }
  return out
}
const describeBook = (book) => ({
  id: book.id,
  name: book.name,
  rev: book.rev,
  sheets: book.sheets.map((s, i) => {
    const used = book.usedRange(s.id)
    let formulas = 0
    for (const c of s.cells.values()) if (c.f != null) formulas++
    return { name: s.name, used: used ? rangeText(used) : '(empty)', cells: s.cells.size, formulas, charts: s.charts.length, ...(s.cf.length ? { conditional_formats: s.cf.length } : {}), ...(s.dv.length ? { validations: s.dv.length } : {}), index: i }
  }),
  names: book.names,
})
const fileKind = (p, hint) => (hint || extname(p).slice(1) || 'csv').toLowerCase()

// ── tools ────────────────────────────────────────────────────────────────────────────────────────────────────────────────
export const TOOLS = [
  {
    name: 'open_sheets',
    description: 'Open the Sheets app: a full spreadsheet (formulas, formatting, charts, many sheets) with a database explorer and SQL editor. Shows the workbook on screen.',
    inputSchema: S({ workbook: WB, view: { type: 'string', enum: ['sheet', 'database'], description: 'Start on the spreadsheet or the database explorer' } }),
    annotations: { readOnlyHint: true },
    _meta: { ui: { resourceUri: UI_URI } },
    run: (a) => {
      const b = store.get(a.workbook)
      return text(`Opened "${b.name}" (${b.sheets.map((s) => s.name).join(', ')}).`, { workbook: b.toJSON(), workbooks: store.list(), connections: db.listConnections(), view: a.view || 'sheet' })
    },
  },
  {
    name: 'list_workbooks',
    description: 'List the workbooks, newest first.',
    inputSchema: S({}),
    annotations: { readOnlyHint: true },
    run: () => {
      const l = store.list()
      return text(l.length ? l.map((w) => `${w.name} [${w.id}] sheets: ${w.sheets.join(', ')}`).join('\n') : 'No workbooks yet. Use create_workbook.', { workbooks: l })
    },
  },
  {
    name: 'create_workbook',
    description: 'Create a workbook. Optionally fill sheets with rows of values (strings starting with = are formulas).',
    inputSchema: S(
      {
        name: str('Workbook name'),
        sheets: { type: 'array', description: 'Sheets to create', items: S({ name: str('Sheet name'), rows: { type: 'array', description: 'Rows of values from A1', items: { type: 'array', items: {} } } }, ['name']) },
      },
      ['name'],
    ),
    run: (a) => {
      const b = store.create(a.name, { sheets: a.sheets?.length ? a.sheets.map((s) => ({ name: s.name })) : undefined })
      for (const [i, s] of (a.sheets || []).entries())
        if (s.rows?.length) {
          const entries = {}
          s.rows.forEach((row, r) => row.forEach((v, c) => v !== null && v !== '' && (entries[addr(r, c)] = v)))
          b.apply(setCells(b, b.sheets[i], entries))
        }
      b.recalc()
      store.save(b)
      return text(`Created "${b.name}" [${b.id}] with ${b.sheets.map((s) => s.name).join(', ')}.`, describeBook(b))
    },
  },
  {
    name: 'manage_workbook',
    description: 'Rename, duplicate or delete a workbook. Deleting cannot be undone.',
    inputSchema: S({ workbook: WB, action: { type: 'string', enum: ['rename', 'duplicate', 'delete'] }, name: str('New name (rename, duplicate)') }, ['workbook', 'action']),
    run: (a) => {
      const b = store.get(a.workbook)
      if (a.action === 'delete') return text(`Deleted "${store.remove(b.id)}".`)
      if (a.action === 'duplicate') {
        const d = store.duplicate(b.id, a.name)
        return text(`Duplicated as "${d.name}" [${d.id}].`)
      }
      if (!a.name) throw new Error('give the new name')
      commit(b, [{ op: 'rename', name: a.name }])
      return text(`Renamed to "${b.name}".`)
    },
  },
  {
    name: 'workbook_info',
    description: 'Describe a workbook: its sheets, the used range of each, formula counts, charts and defined names.',
    inputSchema: S({ workbook: WB }),
    annotations: { readOnlyHint: true },
    run: (a) => {
      const d = describeBook(store.get(a.workbook))
      return text(`${d.name}: ${d.sheets.map((s) => `${s.name} (${s.used}, ${s.formulas} formulas)`).join('; ')}`, d)
    },
  },
  {
    name: 'read_range',
    description: 'Read cells. mode "values" gives results, "formulas" gives what was typed (formulas as =...), "display" gives them as formatted on screen. Whole columns are cut to the used part.',
    inputSchema: S({ workbook: WB, sheet: SH, range: RANGE, mode: { type: 'string', enum: ['values', 'formulas', 'display'] }, max_rows: { type: 'number', description: 'Default 500' } }, ['range']),
    annotations: { readOnlyHint: true },
    run: (a) => {
      const { book, sh, g } = where(a, { needRange: true })
      const t = trimmed(book, sh, g)
      const cap = Math.min(Number(a.max_rows) || 500, 5000)
      const cut = { ...t, r2: Math.min(t.r2, t.r1 + cap - 1) }
      const mode = a.mode || 'values'
      const i = book.sheetIndex(sh.id)
      const rows =
        mode === 'formulas'
          ? Array.from({ length: cut.r2 - cut.r1 + 1 }, (_, r) => Array.from({ length: cut.c2 - cut.c1 + 1 }, (_, c) => rawText(sh.cells.get((cut.r1 + r) * 16384 + cut.c1 + c))))
          : rowsOf(book, sh.id, cut, { raw: mode === 'values' })
      void i
      const clipped = t.r2 > cut.r2
      return text(`${sh.name}!${rangeText(cut)}${clipped ? ` (first ${cap} rows of ${t.r2 - t.r1 + 1})` : ''}\n${grid(rows)}`, { sheet: sh.name, range: rangeText(cut), mode, values: rows.map((r) => r.map(show)), truncated: clipped })
    },
  },
  {
    name: 'write_cells',
    description: 'Write values or formulas. Give `cells` as {"A1": 5, "B1": "=A1*2"} or `range` (the top-left cell is enough) with `values` as rows. Text starting with = is a formula; numbers, percents, dates and money typed as text are understood. Formats are kept. Returns any errors the formulas produced.',
    inputSchema: S({ workbook: WB, sheet: SH, cells: { type: 'object', description: 'Map of cell address to value', additionalProperties: true }, range: RANGE, values: { type: 'array', description: 'Rows of values starting at range', items: { type: 'array', items: {} } } }),
    run: (a) => {
      const { book, sh, g } = where(a)
      const entries = { ...(a.cells || {}) }
      if (a.values) {
        if (!g) throw new Error('say where the values go: range, like A1')
        a.values.forEach((row, r) => row.forEach((v, c) => (entries[addr(g.r1 + r, g.c1 + c)] = v)))
      }
      if (!Object.keys(entries).length) throw new Error('nothing to write: pass cells, or range with values')
      commit(book, [setCells(book, sh, entries)])
      const touched = Object.keys(entries).map(parseAddr)
      const box = { r1: Math.min(...touched.map((p) => p.r)), c1: Math.min(...touched.map((p) => p.c)), r2: Math.max(...touched.map((p) => p.r)), c2: Math.max(...touched.map((p) => p.c)) }
      const errs = errorsOf(book, sh, box)
      return text(`Wrote ${Object.keys(entries).length} cells to ${sh.name}.${errs.length ? ` Errors: ${errs.join(', ')}` : ''}`, { written: Object.keys(entries).length, errors: errs, rev: book.rev })
    },
  },
  {
    name: 'clear_range',
    description: 'Clear cells: contents, formats, or both.',
    inputSchema: S({ workbook: WB, sheet: SH, range: RANGE, what: { type: 'string', enum: ['all', 'contents', 'formats'] } }, ['range']),
    run: (a) => {
      const { book, sh, g } = where(a, { needRange: true })
      const t = trimmed(book, sh, g)
      const what = a.what || 'all'
      const cells = {}
      for (const [k, c] of sh.cells) {
        const r = Math.floor(k / 16384)
        const col = k % 16384
        if (r < t.r1 || r > t.r2 || col < t.c1 || col > t.c2) continue
        cells[addr(r, col)] = what === 'all' ? null : what === 'contents' ? (c.s ? { s: c.s } : null) : c.f != null ? { f: c.f } : { v: c.v }
      }
      commit(book, [{ op: 'set', sheet: sh.id, cells }])
      return text(`Cleared ${Object.keys(cells).length} cells (${what}).`)
    },
  },
  {
    name: 'format_range',
    description: 'Style a range. style keys: bold, italic, underline, strike, color (text, #rrggbb), fill (#rrggbb), align (left|center|right), vertical (top|middle|bottom), wrap_text, number_format (like "#,##0.00", "0%", "$#,##0", "yyyy-mm-dd"), font_size, borders ("all"|"none"|{t,b,l,r:"thin"|"thick"|"dashed"}). Also: column_width and row_height in pixels for the whole range, merge (true/false).',
    inputSchema: S({ workbook: WB, sheet: SH, range: RANGE, style: { type: 'object', additionalProperties: true }, column_width: { type: 'number' }, row_height: { type: 'number' }, merge: { type: 'boolean' } }, ['range']),
    run: (a) => {
      const { book, sh, g } = where(a, { needRange: true })
      const ops = []
      const st = styleOf(a.style || {})
      if (Object.keys(st).length || own(a.style, 'borders')) {
        const t = trimmedForStyle(g)
        const cells = {}
        for (let r = t.r1; r <= t.r2; r++)
          for (let c = t.c1; c <= t.c2; c++) {
            const old = sh.cells.get(r * 16384 + c)
            const s = { ...(old?.s || {}), ...st }
            for (const k of Object.keys(s)) if (s[k] === null || s[k] === false) delete s[k]
            cells[addr(r, c)] = { ...(old ? (old.f != null ? { f: old.f } : old.v !== undefined ? { v: old.v } : {}) : {}), ...(Object.keys(s).length ? { s } : {}) }
          }
        ops.push({ op: 'set', sheet: sh.id, cells })
      }
      if (a.column_width) {
        const cols = { ...sh.cols }
        for (let c = g.c1; c <= Math.min(g.c2, g.c1 + 200); c++) cols[addr(0, c).replace(/\d+/, '')] = Number(a.column_width)
        ops.push({ op: 'prop', sheet: sh.id, key: 'cols', value: cols })
      }
      if (a.row_height) {
        const rows = { ...sh.rows }
        for (let r = g.r1; r <= Math.min(g.r2, g.r1 + 2000); r++) rows[String(r + 1)] = Number(a.row_height)
        ops.push({ op: 'prop', sheet: sh.id, key: 'rows', value: rows })
      }
      if (own(a, 'merge')) {
        const txt = rangeText(g)
        const merges = sh.merges.filter((m) => m !== txt)
        if (a.merge && txt.includes(':')) merges.push(txt)
        ops.push({ op: 'prop', sheet: sh.id, key: 'merges', value: merges })
      }
      if (!ops.length) throw new Error('nothing to change: pass style, column_width, row_height or merge')
      commit(book, ops)
      return text(`Formatted ${rangeText(g)}.`)
    },
  },
  {
    name: 'edit_structure',
    description: 'Insert or delete rows or columns (formulas, charts and other sheets that point there are updated), or freeze panes, hide rows or columns.',
    inputSchema: S({ workbook: WB, sheet: SH, action: { type: 'string', enum: ['insert_rows', 'delete_rows', 'insert_columns', 'delete_columns', 'freeze', 'hide_rows', 'hide_columns', 'show_all'] }, at: { type: 'string', description: 'Row number (1-based) or column letter where it happens. For freeze: not used.' }, count: { type: 'number' }, freeze_rows: { type: 'number' }, freeze_columns: { type: 'number' } }, ['action']),
    run: (a) => {
      const { book, sh } = where(a)
      const n = Math.max(1, Number(a.count) || 1)
      if (a.action === 'freeze') {
        commit(book, [{ op: 'prop', sheet: sh.id, key: 'freeze', value: { rows: Number(a.freeze_rows) || 0, cols: Number(a.freeze_columns) || 0 } }])
        return text(`Froze ${Number(a.freeze_rows) || 0} rows and ${Number(a.freeze_columns) || 0} columns.`)
      }
      if (a.action === 'show_all') {
        commit(book, [{ op: 'prop', sheet: sh.id, key: 'hiddenRows', value: [] }, { op: 'prop', sheet: sh.id, key: 'hiddenCols', value: [] }])
        return text('Showed all rows and columns.')
      }
      const row = a.action.endsWith('rows')
      const at = row ? Number(a.at) - 1 : colIndex(String(a.at))
      if (!Number.isInteger(at) || at < 0) throw new Error(row ? 'at is a row number, like 3' : 'at is a column letter, like C')
      if (a.action.startsWith('hide')) {
        const key = row ? 'hiddenRows' : 'hiddenCols'
        commit(book, [{ op: 'prop', sheet: sh.id, key, value: [...new Set([...sh[key], ...Array.from({ length: n }, (_, i) => at + i)])] }])
        return text(`Hid ${n} ${row ? 'rows' : 'columns'}.`)
      }
      commit(book, [{ op: a.action.startsWith('insert') ? 'insert' : 'delete', sheet: sh.id, dim: row ? 'row' : 'col', at, n }])
      return text(`${a.action.startsWith('insert') ? 'Inserted' : 'Deleted'} ${n} ${row ? 'row' : 'column'}${n > 1 ? 's' : ''} at ${a.at}.`)
    },
  },
  {
    name: 'manage_sheets',
    description: 'Add, rename, duplicate, move, recolour or delete a sheet (tab) of a workbook.',
    inputSchema: S({ workbook: WB, action: { type: 'string', enum: ['add', 'rename', 'duplicate', 'move', 'delete', 'color'] }, sheet: SH, name: str('New name (add, rename, duplicate)'), position: { type: 'number', description: 'Zero-based position (move, add)' }, color: str('#rrggbb') }, ['action']),
    run: (a) => {
      const book = store.get(a.workbook)
      if (a.action === 'add') {
        commit(book, [{ op: 'addSheet', index: a.position, sheet: { name: a.name } }])
        return text(`Added a sheet. Sheets: ${book.sheets.map((s) => s.name).join(', ')}.`)
      }
      const sh = book.sheet(a.sheet ?? 0)
      if (a.action === 'rename') commit(book, [{ op: 'renameSheet', id: sh.id, name: a.name }])
      else if (a.action === 'delete') commit(book, [{ op: 'removeSheet', id: sh.id }])
      else if (a.action === 'move') commit(book, [{ op: 'moveSheet', id: sh.id, index: Number(a.position) || 0 }])
      else if (a.action === 'color') commit(book, [{ op: 'prop', sheet: sh.id, key: 'color', value: a.color || '' }])
      else {
        const j = book.sheetJSON(sh)
        commit(book, [{ op: 'addSheet', index: book.sheets.indexOf(sh) + 1, sheet: { ...j, id: undefined, name: a.name || `${sh.name} copy` } }])
      }
      return text(`Sheets: ${book.sheets.map((s) => s.name).join(', ')}.`)
    },
  },
  {
    name: 'sort_range',
    description: 'Sort rows of a range by one or more columns. Formulas move with their rows.',
    inputSchema: S({ workbook: WB, sheet: SH, range: RANGE, keys: { type: 'array', description: 'Sort keys in order', items: S({ column: str('Column letter, like B'), ascending: { type: 'boolean' } }, ['column']) }, has_header: { type: 'boolean', description: 'First row is a header and stays put' } }, ['range', 'keys']),
    run: (a) => {
      const { book, sh, g } = where(a, { needRange: true })
      const keys = a.keys.map((k) => ({ col: colIndex(k.column) - g.c1, asc: k.ascending !== false }))
      if (keys.some((k) => k.col < 0 || k.col > g.c2 - g.c1)) throw new Error('a sort column is outside the range')
      commit(book, [book.planSort(sh.id, g, keys, { header: !!a.has_header })])
      return text(`Sorted ${rangeText(g)}.`)
    },
  },
  {
    name: 'find_replace',
    description: 'Find text in a workbook (all sheets, or one) and optionally replace it. Without `replace` it only lists matches.',
    inputSchema: S({ workbook: WB, sheet: SH, find: str('Text to find'), replace: str('Replacement; leave out to only search'), match_case: { type: 'boolean' }, whole_cell: { type: 'boolean' }, in_formulas: { type: 'boolean' } }, ['find']),
    run: (a) => {
      const book = store.get(a.workbook)
      book.recalc()
      const plan = book.planFind(a.sheet ?? null, a.find, a.replace ?? null, { matchCase: !!a.match_case, whole: !!a.whole_cell, formulas: !!a.in_formulas })
      if (a.replace !== undefined && plan.ops.length) commit(book, plan.ops)
      return text(`${plan.matches.length} match${plan.matches.length === 1 ? '' : 'es'}${a.replace !== undefined ? ' replaced' : ''}.\n${plan.matches.slice(0, 50).map((m) => `${m.sheet}!${m.cell}: ${m.text}`).join('\n')}`, { matches: plan.matches.slice(0, 200) })
    },
  },
  {
    name: 'fill_range',
    description: 'Fill a target range from a source range, like dragging the fill handle: copies formulas (references move) and continues number, date-name and text-number series.',
    inputSchema: S({ workbook: WB, sheet: SH, source: RANGE, target: RANGE, series: { type: 'boolean', description: 'Continue series (default true); false copies' } }, ['source', 'target']),
    run: (a) => {
      const book = store.get(a.workbook)
      const sh = book.sheet(a.sheet ?? 0)
      const src = parseRange(a.source)
      const dst = parseRange(a.target)
      if (!src || !dst) throw new Error('source and target must be A1 ranges')
      commit(book, [book.planFill(sh.id, src, dst, { series: a.series !== false })])
      return text(`Filled ${a.target} from ${a.source}.`)
    },
  },
  {
    name: 'define_name',
    description: 'Create or remove a named range or constant (used in formulas by name). refers_to is a range like Sheet1!$B$1:$B$20 or a formula; leave it out to remove the name.',
    inputSchema: S({ workbook: WB, name: str('Name, letters, digits, underscores'), refers_to: str('Range or formula') }, ['name']),
    run: (a) => {
      const book = store.get(a.workbook)
      if (!/^[A-Za-z_][\w.]*$/.test(a.name) || /^[A-Za-z]{1,3}\d+$/.test(a.name)) throw new Error('that is not a valid name')
      const names = { ...book.names }
      if (a.refers_to) names[a.name] = String(a.refers_to).replace(/^=/, '')
      else delete names[a.name]
      commit(book, [{ op: 'names', value: names }])
      return text(a.refers_to ? `${a.name} refers to ${names[a.name]}.` : `Removed ${a.name}.`)
    },
  },
  {
    name: 'add_chart',
    description: 'Add a chart on a sheet from a data range (first row/column are labels). type: column, bar, line, area, pie, scatter. Give remove_id to delete a chart instead.',
    inputSchema: S({ workbook: WB, sheet: SH, type: { type: 'string', enum: ['column', 'bar', 'line', 'area', 'pie', 'scatter'] }, range: RANGE, title: str('Title'), anchor: str('Top-left cell, like F2'), stacked: { type: 'boolean' }, remove_id: str('Chart id to remove') }),
    run: (a) => {
      const book = store.get(a.workbook)
      const sh = book.sheet(a.sheet ?? 0)
      if (a.remove_id) {
        commit(book, [{ op: 'prop', sheet: sh.id, key: 'charts', value: sh.charts.filter((c) => c.id !== a.remove_id) }])
        return text('Removed the chart.')
      }
      const w = where({ ...a, sheet: sh.id })
      if (!w.g) throw new Error('a data range is needed')
      const ch = { id: Math.random().toString(36).slice(2, 8), type: a.type || 'column', title: a.title || '', range: `${/^[A-Za-z_]\w*$/.test(sh.name) ? sh.name : `'${sh.name}'`}!${rangeText(w.g)}`, anchor: a.anchor || addr(w.g.r1, w.g.c2 + 2), w: 480, h: 300, stacked: !!a.stacked }
      commit(book, [{ op: 'prop', sheet: sh.id, key: 'charts', value: [...sh.charts, ch] }])
      return text(`Added a ${ch.type} chart "${ch.title}" at ${ch.anchor} [${ch.id}].`, { chart: ch })
    },
  },
  {
    name: 'conditional_format',
    description: 'Add a conditional format to a range. rule examples: {"type":"cell","op":">","a":100,"style":{"bg":"#fde2e2","fg":"#9b1c1c"}}, {"type":"text","op":"contains","text":"late","style":{...}}, {"type":"scale","colors":["#f8696b","#ffeb84","#63be7b"]}, {"type":"bar","color":"#5b9bd5"}, {"type":"formula","formula":"=$B2>$C2","style":{...}}, {"type":"dup"}, {"type":"top","n":3}, {"type":"blank"}, {"type":"error"}. Pass clear:true to remove all rules from the range.',
    inputSchema: S({ workbook: WB, sheet: SH, range: RANGE, rule: { type: 'object', additionalProperties: true }, clear: { type: 'boolean' } }, ['range']),
    run: (a) => {
      const { book, sh, g } = where(a, { needRange: true })
      const t = rangeText(g)
      const keep = sh.cf.filter((r) => r.range !== t)
      if (a.clear) {
        commit(book, [{ op: 'prop', sheet: sh.id, key: 'cf', value: keep }])
        return text(`Cleared conditional formats on ${t}.`)
      }
      if (!a.rule?.type) throw new Error('rule.type is needed')
      commit(book, [{ op: 'prop', sheet: sh.id, key: 'cf', value: [...sh.cf, { id: Math.random().toString(36).slice(2, 8), ...a.rule, range: t, ...(a.rule.style ? { style: styleOf(a.rule.style) } : {}) }] }])
      return text(`Added a ${a.rule.type} rule on ${t}.`)
    },
  },
  {
    name: 'data_validation',
    description: 'Restrict what a range accepts. rule: {"type":"list","items":["Open","Closed"]} or {"type":"list","source":"Sheet2!A1:A9"}, {"type":"number"|"whole"|"date","op":"between","a":0,"b":100}, {"type":"text","op":"<=","a":20}, {"type":"custom","formula":"=ISNUMBER(A1)"}; add "message" for the error text. clear:true removes the rules on the range.',
    inputSchema: S({ workbook: WB, sheet: SH, range: RANGE, rule: { type: 'object', additionalProperties: true }, clear: { type: 'boolean' } }, ['range']),
    run: (a) => {
      const { book, sh, g } = where(a, { needRange: true })
      const t = rangeText(g)
      const keep = sh.dv.filter((r) => r.range !== t)
      if (!a.clear && !a.rule?.type) throw new Error('rule.type is needed')
      commit(book, [{ op: 'prop', sheet: sh.id, key: 'dv', value: a.clear ? keep : [...keep, { ...a.rule, range: t }] }])
      return text(a.clear ? `Removed validation on ${t}.` : `${t} now accepts only ${a.rule.type} values.`)
    },
  },
  {
    name: 'evaluate',
    description: 'Evaluate a formula without writing it anywhere, optionally in the context of a workbook (so it can refer to cells and names). Example: =SUMIFS(Sales!C:C,Sales!A:A,"North").',
    inputSchema: S({ formula: str('A formula, with or without the leading ='), workbook: WB, sheet: SH }, ['formula']),
    annotations: { readOnlyHint: true },
    run: (a) => {
      const book = a.workbook ? store.get(a.workbook) : new Book()
      const v = book.evaluate(a.formula, a.sheet ?? 0)
      const out = Array.isArray(v) ? v.map((r) => r.map(show)) : show(v)
      return text(Array.isArray(out) ? grid(out) : String(out), { value: out })
    },
  },
  {
    name: 'function_list',
    description: 'List the spreadsheet functions this engine supports.',
    inputSchema: S({ contains: str('Only names containing this') }),
    annotations: { readOnlyHint: true },
    run: (a) => {
      const l = FUNCTION_NAMES().filter((n) => !a.contains || n.includes(String(a.contains).toUpperCase()))
      return text(`${l.length} functions: ${l.join(', ')}`)
    },
  },
  {
    name: 'import_data',
    description: 'Put data into a workbook from CSV text, or from a .csv, .tsv or .xlsx file on the computer running Sheets. A .xlsx file with no `workbook` becomes a new workbook; CSV goes to the sheet at the cell `at`.',
    inputSchema: S({ workbook: WB, sheet: SH, at: str('Top-left cell, default A1'), csv: str('CSV or TSV text'), path: str('File path (.csv, .tsv, .xlsx)'), name: str('Name for a new workbook from an .xlsx file'), delimiter: str('Delimiter, guessed when left out') }),
    run: (a) => {
      if (a.path) noFiles()
      if (a.path && fileKind(a.path) === 'xlsx') {
        const p = path(a.path)
        if (!existsSync(p)) throw new Error(`${p} does not exist`)
        const imp = readXlsx(readFileSync(p), a.name || p.split('/').pop().replace(/\.xlsx$/i, ''))
        const b = store.create(imp.name, imp.toJSON())
        b.recalc()
        return text(`Imported "${b.name}" [${b.id}]: ${b.sheets.map((s) => s.name).join(', ')}.`, describeBook(b))
      }
      let raw = a.csv
      if (!raw && a.path) {
        const p = path(a.path)
        if (!existsSync(p)) throw new Error(`${p} does not exist`)
        raw = readFileSync(p, 'utf8')
      }
      if (!raw) throw new Error('pass csv text or a path')
      const { book, sh } = where(a)
      const top = parseAddr(a.at || 'A1')
      if (!top) throw new Error('at must be a cell like B2')
      const rows = parseCSV(raw, a.delimiter)
      if (rows.length * Math.max(...rows.map((r) => r.length)) > MAX_CELLS) throw new Error(`that is more than ${MAX_CELLS} cells; import it in pieces`)
      commit(book, [{ op: 'set', sheet: sh.id, cells: cellsFromRows(rows, top) }])
      return text(`Imported ${rows.length} rows into ${sh.name} from ${addr(top.r, top.c)}.`)
    },
  },
  {
    name: 'export_data',
    description: 'Export a sheet or range as CSV (returned as text, or written to `path`) or the whole workbook as .xlsx (written to `path`, default in the Sheets data folder).',
    inputSchema: S({ workbook: WB, sheet: SH, range: RANGE, format: { type: 'string', enum: ['csv', 'tsv', 'xlsx'] }, path: str('Where to write the file') }),
    annotations: { readOnlyHint: false },
    run: (a) => {
      const { book, sh, g } = where(a)
      const fmt = a.format || (a.path ? fileKind(a.path) : 'csv')
      if (fmt === 'xlsx') {
        noFiles()
        const p = path(a.path || `${db.dataDir()}/${book.name.replace(/[^\w.-]+/g, '_')}.xlsx`)
        writeFileSync(p, writeXlsx(book))
        return text(`Wrote ${p} (${book.sheets.length} sheets).`, { path: p })
      }
      const used = book.usedRange(sh.id)
      const box = g ? trimmed(book, sh, g) : used
      if (!box) return text('')
      const out = toCSV(rowsOf(book, sh.id, box, { raw: false }), fmt === 'tsv' ? '\t' : ',')
      if (a.path) {
        noFiles()
        writeFileSync(path(a.path), out)
        return text(`Wrote ${path(a.path)}.`, { path: path(a.path) })
      }
      return text(out)
    },
  },

  // ── databases ──────────────────────────────────────────────────────────────────────────────────────────────────────────
  {
    name: 'db_connections',
    description: 'List the saved database connections (SQLite files, PostgreSQL, MySQL). Passwords are never shown. New remote connections are added by the person in the app, so their passwords stay out of the chat.',
    inputSchema: S({}),
    annotations: { readOnlyHint: true },
    run: () => {
      const l = db.listConnections()
      return text(l.length ? l.map((c) => `${c.name} (${db.TYPES[c.type]}) ${c.label}${c.readOnly ? ' [read-only]' : ''}`).join('\n') : 'No connections yet. Add one in the Database tab, or db_add_sqlite for a SQLite file.', { connections: l })
    },
  },
  {
    name: 'db_add_sqlite',
    description: 'Save a SQLite database file as a connection (the file is created if it does not exist).',
    inputSchema: S({ name: str('Connection name'), file: str('Path to the .db / .sqlite file, or :memory:'), read_only: { type: 'boolean' } }, ['name', 'file']),
    run: (a) => {
      noFiles()
      const c = db.saveConnection({ name: a.name, type: 'sqlite', file: a.file, readOnly: !!a.read_only })
      return text(`Saved connection "${c.name}" → ${c.label}.`)
    },
  },
  {
    name: 'db_schema',
    description: 'Show the tables, views and columns (types, primary keys, foreign keys) of a connection.',
    inputSchema: S({ connection: str('Connection name or id') }, ['connection']),
    annotations: { readOnlyHint: true },
    run: async (a) => {
      const s = await db.schema(a.connection)
      const lines = s.schemas.flatMap((sc) => sc.tables.map((t) => `${s.type === 'sqlite' ? '' : `${sc.name}.`}${t.name} (${t.kind}): ${t.columns.map((c) => `${c.name} ${c.type}${c.pk ? ' PK' : ''}${c.fk ? ` → ${c.fk}` : ''}`).join(', ')}`))
      return text(lines.join('\n') || '(no tables)', s)
    },
  },
  {
    name: 'db_query',
    description: 'Run a read-only SQL query (SELECT, WITH, EXPLAIN, SHOW) and return rows. It refuses anything that changes data; use db_execute for that. Rows are capped at `limit` (default 200, max 5000).',
    inputSchema: S({ connection: str('Connection name or id'), sql: str('The query'), params: { type: 'array', description: 'Values for ? (or $1) placeholders', items: {} }, limit: { type: 'number' } }, ['connection', 'sql']),
    annotations: { readOnlyHint: true },
    run: async (a) => {
      const r = await db.runQuery(a.connection, a.sql, { params: a.params || [], limit: Math.min(Number(a.limit) || 200, 5000) })
      return text(`${r.rowCount} row${r.rowCount === 1 ? '' : 's'}${r.truncated ? ' (more exist)' : ''} in ${r.ms} ms\n${grid([r.columns, ...r.rows])}`, r)
    },
  },
  {
    name: 'db_execute',
    description: 'Run SQL that changes data or the schema (INSERT, UPDATE, DELETE, CREATE, ALTER, DROP). The person is asked to allow it first. Not available on a read-only connection.',
    inputSchema: S({ connection: str('Connection name or id'), sql: str('The statement'), params: { type: 'array', items: {} } }, ['connection', 'sql']),
    run: async (a) => {
      const r = await db.runQuery(a.connection, a.sql, { params: a.params || [], write: true, limit: 200 })
      return text(r.readOnly ? `${r.rowCount} rows\n${grid([r.columns, ...r.rows])}` : `Done in ${r.ms} ms${r.affected ? `, ${r.affected} rows affected` : ''}.`, r)
    },
  },
  {
    name: 'db_query_to_sheet',
    description: 'Run a read-only query and put the result on a sheet (header row in bold), starting at a cell. The sheet is then ordinary cells: formulas, charts and formatting work on it.',
    inputSchema: S({ connection: str('Connection name or id'), sql: str('A SELECT query'), workbook: WB, sheet: SH, at: str('Top-left cell, default A1'), header: { type: 'boolean' }, limit: { type: 'number', description: 'Max rows, default 5000' } }, ['connection', 'sql']),
    run: async (a) => {
      const { book, sh } = where(a)
      const top = parseAddr(a.at || 'A1')
      if (!top) throw new Error('at must be a cell like B2')
      const r = await db.runQuery(a.connection, a.sql, { limit: Math.min(Number(a.limit) || 5000, 20000) })
      commit(book, [{ op: 'set', sheet: sh.id, cells: db.cellsFromResult(r, top, { header: a.header !== false }) }])
      return text(`Put ${r.rowCount} rows × ${r.columns.length} columns on ${sh.name} at ${addr(top.r, top.c)}${r.truncated ? ' (the query had more rows; raise limit)' : ''}.`)
    },
  },
  {
    name: 'sheet_to_table',
    description: 'Write a range of a sheet (first row = column names) into a database table, creating it. The person is asked to allow it. Values written are the results of formulas.',
    inputSchema: S({ connection: str('Connection name or id'), table: str('Table name'), schema: str('Schema (PostgreSQL), default public'), workbook: WB, sheet: SH, range: RANGE, replace: { type: 'boolean', description: 'Drop and recreate the table if it exists' } }, ['connection', 'table', 'range']),
    run: async (a) => {
      const { book, sh, g } = where(a, { needRange: true })
      const t = trimmed(book, sh, g)
      const rows = rowsOf(book, sh.id, t, { raw: true }).map((r) => r.map((v) => (v === undefined ? null : v)))
      const c = db.getConnection(a.connection)
      const r = await db.writeTable(a.connection, c.type === 'postgres' ? a.schema || 'public' : a.schema, a.table, rows, { replace: !!a.replace })
      return text(`Wrote ${r.inserted} rows × ${r.columns} columns to ${a.table}.`)
    },
  },

  // ── only the page may call these ──────────────────────────────────────────────────────────────────────────────────────────
  {
    name: 'app_state',
    description: 'Page only: load a workbook, the workbook list and the connections.',
    inputSchema: S({ workbook: WB }),
    annotations: { readOnlyHint: true },
    _meta: { ui: { visibility: ['app'] } },
    run: (a) => {
      const b = store.get(a.workbook)
      return text('', { workbook: b.toJSON(), workbooks: store.list(), connections: db.listConnections(), functions: FUNCTION_NAMES() })
    },
  },
  {
    name: 'app_poll',
    description: 'Page only: has this workbook changed since rev?',
    inputSchema: S({ id: str('Workbook id'), rev: { type: 'number' } }, ['id']),
    annotations: { readOnlyHint: true },
    _meta: { ui: { visibility: ['app'] } },
    run: (a) => {
      const b = store.get(a.id)
      return text('', b.rev === a.rev ? { changed: false, rev: b.rev } : { changed: true, rev: b.rev, workbook: b.toJSON() })
    },
  },
  {
    name: 'app_apply',
    description: 'Page only: apply the edits a person made on screen.',
    inputSchema: S({ id: str('Workbook id'), ops: { type: 'array', items: { type: 'object', additionalProperties: true } } }, ['id', 'ops']),
    _meta: { ui: { visibility: ['app'] } },
    run: (a) => {
      const b = store.get(a.id)
      if (a.ops.length > 5000) throw new Error('too many edits at once')
      const before = b.rev
      b.applyAll(a.ops)
      b.recalc()
      store.save(b)
      return text('', { rev: b.rev, before })
    },
  },
  {
    name: 'app_workbook',
    description: 'Page only: create, rename, duplicate or delete a workbook.',
    inputSchema: S({ action: { type: 'string', enum: ['create', 'rename', 'duplicate', 'delete'] }, id: str('Workbook id'), name: str('Name') }, ['action']),
    _meta: { ui: { visibility: ['app'] } },
    run: (a) => {
      let b
      if (a.action === 'create') b = store.create(a.name || 'Untitled workbook')
      else if (a.action === 'duplicate') b = store.duplicate(a.id, a.name)
      else if (a.action === 'rename') {
        b = store.get(a.id)
        commit(b, [{ op: 'rename', name: a.name }])
      } else {
        store.remove(a.id)
        b = store.get()
      }
      return text('', { workbook: b.toJSON(), workbooks: store.list() })
    },
  },
  {
    name: 'app_import',
    description: 'Page only: import CSV text into a sheet.',
    inputSchema: S({ id: str('Workbook id'), sheet: str('Sheet id'), at: str('Cell'), csv: str('CSV text'), name: str('File name') }, ['id', 'csv']),
    _meta: { ui: { visibility: ['app'] } },
    run: (a) => {
      const b = store.get(a.id)
      const sh = b.sheet(a.sheet ?? 0)
      const top = parseAddr(a.at || 'A1') || { r: 0, c: 0 }
      const rows = parseCSV(a.csv)
      const op = { op: 'set', sheet: sh.id, cells: cellsFromRows(rows, top) }
      b.apply(op)
      b.recalc()
      store.save(b)
      return text('', { workbook: b.toJSON() })
    },
  },
  {
    name: 'app_db_save',
    description: 'Page only: save a database connection (the password is typed on screen and never reaches the model).',
    inputSchema: S({ connection: { type: 'object', additionalProperties: true } }, ['connection']),
    _meta: { ui: { visibility: ['app'] } },
    run: (a) => text('', { connection: db.saveConnection(a.connection), connections: db.listConnections() }),
  },
  {
    name: 'app_db_delete',
    description: 'Page only: remove a saved connection.',
    inputSchema: S({ connection: str('Connection id') }, ['connection']),
    _meta: { ui: { visibility: ['app'] } },
    run: (a) => {
      db.deleteConnection(a.connection)
      return text('', { connections: db.listConnections() })
    },
  },
  {
    name: 'app_db_test',
    description: 'Page only: try a connection before saving it.',
    inputSchema: S({ connection: { type: 'object', additionalProperties: true } }, ['connection']),
    annotations: { readOnlyHint: true },
    _meta: { ui: { visibility: ['app'] } },
    run: async (a) => text('', await db.testConnection(a.connection)),
  },
  {
    name: 'app_db_schema',
    description: 'Page only: tables and columns of a connection.',
    inputSchema: S({ connection: str('Connection id') }, ['connection']),
    annotations: { readOnlyHint: true },
    _meta: { ui: { visibility: ['app'] } },
    run: async (a) => text('', await db.schema(a.connection)),
  },
  {
    name: 'app_db_rows',
    description: 'Page only: rows of a table for the data grid.',
    inputSchema: S({ connection: str('Connection id'), schema: str('Schema'), table: str('Table'), limit: { type: 'number' }, offset: { type: 'number' }, sort: { type: 'object', additionalProperties: true }, where: str('Filter condition') }, ['connection', 'table']),
    annotations: { readOnlyHint: true },
    _meta: { ui: { visibility: ['app'] } },
    run: async (a) => text('', await db.tableData(a.connection, a.schema, a.table, a)),
  },
  {
    name: 'app_db_query',
    description: 'Page only: run a read-only query from the SQL editor.',
    inputSchema: S({ connection: str('Connection id'), sql: str('Query'), limit: { type: 'number' } }, ['connection', 'sql']),
    annotations: { readOnlyHint: true },
    _meta: { ui: { visibility: ['app'] } },
    run: async (a) => text('', await db.runQuery(a.connection, a.sql, { limit: Math.min(Number(a.limit) || 1000, 20000) })),
  },
  {
    name: 'app_db_exec',
    description: 'Page only: run a statement that changes data, or save table edits.',
    inputSchema: S({ connection: str('Connection id'), sql: str('Statement'), schema: str('Schema'), table: str('Table'), changes: { type: 'array', items: { type: 'object', additionalProperties: true } } }, ['connection']),
    _meta: { ui: { visibility: ['app'] } },
    run: async (a) => {
      if (a.changes) return text('', await db.applyChanges(a.connection, a.schema, a.table, a.changes))
      return text('', await db.runQuery(a.connection, a.sql, { write: true, limit: 1000 }))
    },
  },
  {
    name: 'app_db_to_sheet',
    description: 'Page only: run a query and place the result on a sheet.',
    inputSchema: S({ connection: str('Connection id'), sql: str('Query'), id: str('Workbook id'), sheet: str('Sheet id'), at: str('Cell') }, ['connection', 'sql', 'id']),
    _meta: { ui: { visibility: ['app'] } },
    run: async (a) => {
      const b = store.get(a.id)
      const sh = b.sheet(a.sheet ?? 0)
      const r = await db.runQuery(a.connection, a.sql, { limit: 20000 })
      const top = parseAddr(a.at || 'A1') || { r: 0, c: 0 }
      b.apply({ op: 'set', sheet: sh.id, cells: db.cellsFromResult(r, top) })
      b.recalc()
      store.save(b)
      return text('', { workbook: b.toJSON(), rows: r.rowCount })
    },
  },
  {
    name: 'app_sheet_to_table',
    description: 'Page only: write a range to a database table.',
    inputSchema: S({ connection: str('Connection id'), table: str('Table'), schema: str('Schema'), id: str('Workbook id'), sheet: str('Sheet id'), range: RANGE, replace: { type: 'boolean' } }, ['connection', 'table', 'id', 'range']),
    _meta: { ui: { visibility: ['app'] } },
    run: async (a) => {
      const { book, sh, g } = where({ workbook: a.id, sheet: a.sheet, range: a.range }, { needRange: true })
      const t = trimmed(book, sh, g)
      const rows = rowsOf(book, sh.id, t, { raw: true }).map((r) => r.map((v) => (v === undefined ? null : v)))
      const c = db.getConnection(a.connection)
      return text('', await db.writeTable(a.connection, c.type === 'postgres' ? a.schema || 'public' : a.schema, a.table, rows, { replace: !!a.replace }))
    },
  },
  {
    name: 'app_queries_list',
    description: 'Page only: saved SQL queries.',
    inputSchema: S({}),
    annotations: { readOnlyHint: true },
    _meta: { ui: { visibility: ['app'] } },
    run: () => text('', { queries: db.listQueries() }),
  },
  {
    name: 'app_queries_save',
    description: 'Page only: save or remove a SQL query.',
    inputSchema: S({ name: str('Name'), sql: str('SQL'), connection: str('Connection id'), remove: str('Name to remove') }),
    _meta: { ui: { visibility: ['app'] } },
    run: (a) => text('', { queries: a.remove ? db.removeQuery(a.remove) : db.saveQuery(a.name, a.sql, a.connection) }),
  },
  {
    name: 'app_export',
    description: 'Page only: the sheet as CSV text.',
    inputSchema: S({ id: str('Workbook id'), sheet: str('Sheet id'), range: RANGE, format: str('csv or tsv') }, ['id']),
    annotations: { readOnlyHint: true },
    _meta: { ui: { visibility: ['app'] } },
    run: (a) => {
      const { book, sh, g } = where({ workbook: a.id, sheet: a.sheet, range: a.range })
      const box = g ? trimmed(book, sh, g) : book.usedRange(sh.id)
      return text('', { text: box ? toCSV(rowsOf(book, sh.id, box), a.format === 'tsv' ? '\t' : ',') : '' })
    },
  },
]
void mapRefs
const trimmedForStyle = (g) => ({ r1: g.r1, c1: g.c1, r2: Math.min(g.r2, g.r1 + 9999), c2: Math.min(g.c2, g.c1 + 200) })
