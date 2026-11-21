// ── spreadsheet chrome: toolbar, formula bar, tabs, status bar, menus, dialogs, charts, saving ─────────────────────────────────────────
const App = {}

// ── saving and live sync ─────────────────────────────────────────────────────────────────────────────────────────────────
const Save = {
  timer: null,
  inflight: false,
  blocked: false,
  soon() {
    if (this.blocked) return setStatus('Not saved. Click to retry.', 'warn')
    clearTimeout(this.timer)
    setStatus('Saving…')
    this.timer = setTimeout(() => this.flush(), 600)
  },
  async flush() {
    if (this.inflight || !A.pending.length) return
    this.inflight = true
    const ops = A.pending.splice(0)
    const prev = A.rev
    try {
      const r = await Host.tool('app_apply', { id: A.book.id, ops })
      A.rev = r.rev
      this.blocked = false
      setStatus('All changes saved')
      if (r.before !== prev && !A.pending.length) await reloadWorkbook(true)
    } catch (e) {
      A.pending.unshift(...ops)
      this.blocked = true
      setStatus('Not saved. Click to retry.', 'warn')
      if (/said no/i.test(e.message)) toast('Saving needs your permission. In Connectors, set Sheets to trusted so edits save without asking each time.', 'warn')
      else toast(`Could not save: ${e.message}`, 'err')
    } finally {
      this.inflight = false
      if (A.pending.length && !this.blocked) this.soon()
    }
  },
  retry() {
    this.blocked = false
    this.flush()
  },
}
function setStatus(text, kind = '') {
  const el = $('#save-status')
  if (!el) return
  el.textContent = text
  el.className = `save-status ${kind}`
}
function loadWorkbook(json, { keepView = false } = {}) {
  A.book = new XL.Book(json)
  A.rev = json.rev || 0
  A.undo = []
  A.redo = []
  A.pending = []
  const ids = A.book.sheets.map((s) => s.id)
  if (!keepView) A.si = 0
  else A.si = clamp(A.si, 0, ids.length - 1)
  A.book.recalc()
  buildAxes()
  const v = view()
  A.sel = { ...v.sel }
  scroller.scrollLeft = v.sx
  scroller.scrollTop = v.sy
  A.edit = null
  editor.style.display = 'none'
  App.refreshAll()
  schedule()
}
async function reloadWorkbook(quiet) {
  const r = await Host.tool('app_state', { workbook: A.book.id })
  loadWorkbook(r.workbook, { keepView: true })
  A.wbs = r.workbooks
  A.conns = r.connections
  if (!quiet) toast('Updated')
  else toast('Updated from another change', '')
}
setInterval(async () => {
  if (document.hidden || !A.book || Save.inflight || A.pending.length || A.edit || Save.blocked || !Host.ready) return
  try {
    const r = await Host.tool('app_poll', { id: A.book.id, rev: A.rev })
    if (r.changed && !A.pending.length && !A.edit) {
      loadWorkbook(r.workbook, { keepView: true })
      toast('This workbook was changed outside the app. Updated.')
    }
  } catch {
    /* the host may be busy */
  }
}, 3000)

// ── function help ────────────────────────────────────────────────────────────────────────────────────────────────────────
const HELP = `SUM(number1, [number2], …)|Adds numbers
AVERAGE(number1, [number2], …)|Mean of numbers
COUNT(value1, [value2], …)|Counts numbers
COUNTA(value1, [value2], …)|Counts non-empty cells
MIN(number1, [number2], …)|Smallest number
MAX(number1, [number2], …)|Largest number
ROUND(number, digits)|Rounds to a number of digits
IF(condition, if_true, [if_false])|Chooses a result by a test
IFS(test1, value1, [test2, value2], …)|First test that is true
IFERROR(value, if_error)|Value, or a fallback on error
AND(logical1, [logical2], …)|True if all are true
OR(logical1, [logical2], …)|True if any is true
NOT(logical)|Reverses true and false
SWITCH(expression, value1, result1, [default])|Matches a value
SUMIF(range, criteria, [sum_range])|Adds cells that match
SUMIFS(sum_range, criteria_range1, criteria1, …)|Adds cells that match all
COUNTIF(range, criteria)|Counts cells that match
COUNTIFS(criteria_range1, criteria1, …)|Counts cells matching all
AVERAGEIF(range, criteria, [average_range])|Mean of matching cells
AVERAGEIFS(average_range, criteria_range1, criteria1, …)|Mean of cells matching all
MAXIFS(max_range, criteria_range1, criteria1, …)|Largest matching value
MINIFS(min_range, criteria_range1, criteria1, …)|Smallest matching value
SUMPRODUCT(array1, [array2], …)|Sum of products
VLOOKUP(lookup_value, table, col_index, [exact_match])|Looks down the first column
HLOOKUP(lookup_value, table, row_index, [exact_match])|Looks across the first row
XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found], [match_mode], [search_mode])|Looks up and returns a match
INDEX(array, row, [column])|Value at a position
MATCH(lookup_value, lookup_array, [match_type])|Position of a value
OFFSET(reference, rows, cols, [height], [width])|A shifted reference
INDIRECT(ref_text)|A reference from text
FILTER(array, include, [if_empty])|Rows that pass a test
SORT(array, [sort_index], [order], [by_col])|Sorted array
SORTBY(array, by_array1, [order1], …)|Sorted by other ranges
UNIQUE(array, [by_col], [exactly_once])|Distinct rows
SEQUENCE(rows, [columns], [start], [step])|A series of numbers
TRANSPOSE(array)|Flips rows and columns
LET(name1, value1, …, calculation)|Names values inside a formula
CONCAT(text1, [text2], …)|Joins text
TEXTJOIN(delimiter, ignore_empty, text1, …)|Joins text with a separator
LEFT(text, [num_chars])|First characters
RIGHT(text, [num_chars])|Last characters
MID(text, start, num_chars)|Characters from the middle
LEN(text)|Length of text
TRIM(text)|Removes extra spaces
UPPER(text)|Upper case
LOWER(text)|Lower case
PROPER(text)|Capitalises each word
SUBSTITUTE(text, old_text, new_text, [instance])|Replaces text
FIND(find_text, within_text, [start])|Position of text (case-sensitive)
SEARCH(find_text, within_text, [start])|Position of text
TEXT(value, format)|Formats a number as text
VALUE(text)|Text to number
TODAY()|Today's date
NOW()|Date and time now
DATE(year, month, day)|Builds a date
YEAR(date)|Year of a date
MONTH(date)|Month of a date
DAY(date)|Day of a date
EDATE(start_date, months)|Date months away
EOMONTH(start_date, months)|End of a month
DATEDIF(start_date, end_date, unit)|Difference between dates
NETWORKDAYS(start_date, end_date, [holidays])|Working days
WEEKDAY(date, [type])|Day of the week
PMT(rate, nper, pv, [fv], [type])|Loan payment
FV(rate, nper, pmt, [pv], [type])|Future value
PV(rate, nper, pmt, [fv], [type])|Present value
NPV(rate, value1, [value2], …)|Net present value
IRR(values, [guess])|Internal rate of return
RANK(number, ref, [order])|Rank in a list
LARGE(array, k)|k-th largest
SMALL(array, k)|k-th smallest
MEDIAN(number1, [number2], …)|Middle value
STDEV(number1, [number2], …)|Sample standard deviation
ABS(number)|Absolute value
MOD(number, divisor)|Remainder
POWER(number, power)|Number to a power
SQRT(number)|Square root
RAND()|Random number from 0 to 1
RANDBETWEEN(bottom, top)|Random whole number
ISBLANK(value)|True for an empty cell
ISNUMBER(value)|True for a number
ISERROR(value)|True for any error
CHOOSE(index, value1, [value2], …)|Picks by position`
const FUNC_HELP = Object.fromEntries(HELP.split('\n').map((l) => [l.split('(')[0], { sig: l.split('|')[0], desc: l.split('|')[1] }]))

// ── toolbar ──────────────────────────────────────────────────────────────────────────────────────────────────────────────
const refreshers = []
function btn(ic, title, onclick, { active, label, wide } = {}) {
  const b = h('button.tb', { title, onmousedown: (e) => e.preventDefault(), onclick, class: wide ? 'wide' : '' }, ic ? icon(ic, 16) : '', label ? h('span', label) : '')
  if (active) refreshers.push(() => b.classList.toggle('on', !!active()))
  return b
}
const sep = () => h('span.tb-sep')
const PALETTE_ROWS = [
  ['#000000', '#434343', '#666666', '#999999', '#cccccc', '#efefef', '#ffffff'],
  ['#b91c1c', '#ea580c', '#ca8a04', '#15803d', '#0e7490', '#1d4ed8', '#7e22ce'],
  ['#fecaca', '#fed7aa', '#fef08a', '#bbf7d0', '#a5f3fc', '#bfdbfe', '#e9d5ff'],
  ['#fca5a5', '#fdba74', '#fde047', '#86efac', '#67e8f9', '#93c5fd', '#d8b4fe'],
  ['#dc2626', '#f97316', '#eab308', '#22c55e', '#06b6d4', '#3b82f6', '#a855f7'],
  ['#7f1d1d', '#7c2d12', '#713f12', '#14532d', '#164e63', '#1e3a8a', '#581c87'],
]
function colorMenu(anchor, current, onPick, allowNone = true) {
  const r = anchor.getBoundingClientRect()
  closeMenu()
  const input = h('input.input', { value: current || '', placeholder: '#rrggbb', style: { width: '100%' } })
  const m = h(
    '.menu.color-menu',
    { onmousedown: (e) => e.stopPropagation() },
    allowNone ? h('button.menu-item', { onclick: () => (closeMenu(), onPick(null)) }, h('span.mi-label', 'No color')) : '',
    PALETTE_ROWS.map((row) => h('.sw-row', row.map((c) => h('button.swatch', { style: { background: c }, title: c, onclick: () => (closeMenu(), onPick(c)) })))),
    h('.sw-custom', input, h('button.btn', { onclick: () => /^#[0-9a-f]{6}$/i.test(input.value) && (closeMenu(), onPick(input.value.toLowerCase())) }, 'Apply')),
  )
  document.body.append(m)
  m.style.left = `${clamp(r.left, 4, innerWidth - 230)}px`
  m.style.top = `${r.bottom + 2}px`
  openMenu = m
}
const curStyle = () => sheet().cells.get(key(A.sel.ar, A.sel.ac))?.s || {}
App.toggleStyle = (k) => {
  const on = !curStyle()[k]
  styleRange(norm(), { [k]: on ? true : null })
}
const NUMFMTS = [
  ['General', 'General'],
  ['0.00', 'Number'],
  ['$#,##0.00', 'Currency'],
  ['#,##0.00;(#,##0.00)', 'Accounting'],
  ['0.00%', 'Percent'],
  ['m/d/yyyy', 'Short date'],
  ['dddd, mmmm d, yyyy', 'Long date'],
  ['h:mm AM/PM', 'Time'],
  ['0.00E+00', 'Scientific'],
  ['@', 'Text'],
]
function setFormat(nf) {
  styleRange(norm(), { nf: nf === 'General' ? null : nf })
}
function stepDecimals(delta) {
  const nf = curStyle().nf || 'General'
  const v = A.book.get(A.si, A.sel.ar, A.sel.ac)
  if (nf === 'General') {
    const d = typeof v === 'number' ? Math.min(10, ((String(+v.toFixed(10)).split('.')[1] || '').length)) : 0
    return setFormat(`0${d + delta > 0 ? `.${'0'.repeat(d + delta)}` : ''}`)
  }
  const m = nf.match(/^([^0#?]*[0#?,]*)(?:\.([0#?]+))?(.*)$/)
  if (!m) return
  const dec = (m[2] || '').length + delta
  if (dec < 0) return
  setFormat(`${m[1]}${dec ? `.${'0'.repeat(dec)}` : ''}${m[3]}`)
}
function selectionRegion() {
  // a lone cell stands for the block of data around it
  const g = norm()
  if (g.r1 !== g.r2 || g.c1 !== g.c2) return g
  const book = A.book
  const has = (r, c) => r >= 0 && c >= 0 && (() => { const v = book.get(A.si, r, c); return v !== null && v !== undefined && v !== '' })()
  if (!has(g.r1, g.c1)) return g
  let { r1, c1, r2, c2 } = g
  let grew = true
  while (grew) {
    grew = false
    const colHas = (c, a, b) => Array.from({ length: b - a + 1 }, (_, i) => has(a + i, c)).some(Boolean)
    const rowHas = (r, a, b) => Array.from({ length: b - a + 1 }, (_, i) => has(r, a + i)).some(Boolean)
    if (c1 > 0 && colHas(c1 - 1, r1 - 1 < 0 ? 0 : r1 - 1, r2 + 1)) (c1--, (grew = true))
    if (colHas(c2 + 1, Math.max(0, r1 - 1), r2 + 1)) (c2++, (grew = true))
    if (r1 > 0 && rowHas(r1 - 1, Math.max(0, c1 - 1), c2 + 1)) (r1--, (grew = true))
    if (rowHas(r2 + 1, Math.max(0, c1 - 1), c2 + 1)) (r2++, (grew = true))
    if (r2 - r1 > 100000 || c2 - c1 > 2000) break
  }
  return { r1, c1, r2, c2 }
}
function sortBy(asc) {
  const g = selectionRegion()
  const col = clamp(A.sel.ac, g.c1, g.c2) - g.c1
  const first = Array.from({ length: g.c2 - g.c1 + 1 }, (_, i) => A.book.get(A.si, g.r1, g.c1 + i))
  const second = Array.from({ length: g.c2 - g.c1 + 1 }, (_, i) => A.book.get(A.si, g.r1 + 1, g.c1 + i))
  const header = first.every((v) => typeof v === 'string') && second.some((v) => typeof v !== 'string')
  applyOps([A.book.planSort(sheet().id, g, [{ col, asc }], { header })], { label: 'Sort' })
  setSel(g.r1, g.c1, g.r2, g.c2)
}
function toggleFilter() {
  const sh = sheet()
  if (sh.filter) return applyOps([{ op: 'prop', sheet: sh.id, key: 'filter', value: null }], { label: 'Remove filter' })
  const g = selectionRegion()
  if (g.r1 === g.r2 && g.c1 === g.c2) return toast('Select a cell inside your data first.', 'warn')
  applyOps([{ op: 'prop', sheet: sh.id, key: 'filter', value: { range: XL.rangeText(g), cols: {} } }], { label: 'Filter' })
}
function freezeMenu(el) {
  const sh = sheet()
  menuBelow(el, [
    { label: 'Freeze top row', onclick: () => setFreeze(1, sh.freeze.cols) },
    { label: 'Freeze first column', onclick: () => setFreeze(sh.freeze.rows, 1) },
    { label: 'Freeze up to current cell', onclick: () => setFreeze(A.sel.ar, A.sel.ac) },
    { sep: true },
    { label: 'Unfreeze panes', disabled: !sh.freeze.rows && !sh.freeze.cols, onclick: () => setFreeze(0, 0) },
  ])
}
const setFreeze = (rows, cols) => applyOps([{ op: 'prop', sheet: sheet().id, key: 'freeze', value: { rows, cols } }], { label: 'Freeze' })
function toggleMerge() {
  const g = norm()
  const sh = sheet()
  const t = XL.rangeText(g)
  const inside = merges.filter((m) => m.r1 >= g.r1 && m.r2 <= g.r2 && m.c1 >= g.c1 && m.c2 <= g.c2)
  let list = sh.merges.filter((m) => !inside.some((x) => XL.rangeText(x) === m))
  if (!inside.length && t.includes(':')) {
    list.push(t)
    // only the top-left value is kept
    const cells = {}
    for (let r = g.r1; r <= g.r2; r++) for (let c = g.c1; c <= g.c2; c++) if ((r !== g.r1 || c !== g.c1) && sh.cells.has(key(r, c))) cells[A1(r, c)] = sh.cells.get(key(r, c)).s ? { s: sh.cells.get(key(r, c)).s } : null
    applyOps([{ op: 'set', sheet: sh.id, cells }, { op: 'prop', sheet: sh.id, key: 'merges', value: list }], { label: 'Merge' })
  } else applyOps([{ op: 'prop', sheet: sh.id, key: 'merges', value: list }], { label: 'Unmerge' })
}
function bordersMenu(el) {
  const g = norm()
  const set = (spec) => {
    const sh = sheet()
    const cells = {}
    for (let r = g.r1; r <= Math.min(g.r2, g.r1 + 2000); r++)
      for (let c = g.c1; c <= Math.min(g.c2, g.c1 + 200); c++) {
        const old = sh.cells.get(key(r, c))
        const s = { ...(old?.s || {}) }
        let bd = spec === 'none' ? null : { ...(s.bd || {}) }
        if (bd) {
          const edge = { l: c === g.c1, r: c === g.c2, t: r === g.r1, b: r === g.r2 }
          if (spec === 'all') bd = { l: 'thin', r: 'thin', t: 'thin', b: 'thin' }
          if (spec === 'outer') for (const k of Object.keys(edge)) if (edge[k]) bd[k] = 'thin'
          if (spec === 'thick') for (const k of Object.keys(edge)) if (edge[k]) bd[k] = 'thick'
          if (spec === 'top' && edge.t) bd.t = 'thin'
          if (spec === 'bottom' && edge.b) bd.b = 'thin'
          if (spec === 'left' && edge.l) bd.l = 'thin'
          if (spec === 'right' && edge.r) bd.r = 'thin'
          if (spec === 'topbottom') {
            if (edge.t) bd.t = 'thin'
            if (edge.b) bd.b = 'thick'
          }
        }
        if (bd && Object.keys(bd).length) s.bd = bd
        else delete s.bd
        const base = old ? cellJSON(old) : {}
        delete base.s
        cells[A1(r, c)] = Object.keys(s).length ? { ...base, s } : Object.keys(base).length ? base : null
      }
    applyOps([{ op: 'set', sheet: sh.id, cells }], { label: 'Borders' })
  }
  menuBelow(el, [
    { label: 'All borders', onclick: () => set('all') },
    { label: 'Outside borders', onclick: () => set('outer') },
    { label: 'Thick outside border', onclick: () => set('thick') },
    { label: 'Top border', onclick: () => set('top') },
    { label: 'Bottom border', onclick: () => set('bottom') },
    { label: 'Left border', onclick: () => set('left') },
    { label: 'Right border', onclick: () => set('right') },
    { label: 'Top and double-weight bottom', onclick: () => set('topbottom') },
    { sep: true },
    { label: 'No border', onclick: () => set('none') },
  ])
}

function buildToolbar() {
  const nf = sel(NUMFMTS, 'General', { class: 'tb-select', title: 'Number format', style: { width: '104px' }, onchange: (e) => (e.target.value === 'custom' ? App.openFormatCells('number') : setFormat(e.target.value)) })
  nf.append(h('option', { value: 'custom' }, 'More formats…'))
  refreshers.push(() => {
    const f = curStyle().nf || 'General'
    nf.value = NUMFMTS.some((x) => x[0] === f) ? f : 'General'
  })
  const mk = (...kids) => h('.tb-group', kids)
  let fgBtn
  let bgBtn
  let moreBtn
  const more = () =>
    menuBelow(moreBtn, [
      { label: 'Conditional formatting…', icon: 'rules', onclick: () => App.openCF() },
      { label: 'Data validation…', icon: 'validate', onclick: () => App.openDV() },
      { label: 'Named ranges…', icon: 'tag', onclick: () => App.openNames() },
      { label: 'Insert function…', icon: 'sigma', onclick: () => App.openFunctions() },
      { label: 'Find and replace…', icon: 'search', key: 'Ctrl+H', onclick: () => App.openFind(true) },
      { sep: true },
      { label: 'Number format…', key: 'Ctrl+1', onclick: () => App.openFormatCells() },
      { label: 'Strikethrough', onclick: () => App.toggleStyle('st'), checked: !!curStyle().st },
      { label: 'Fewer decimals', onclick: () => stepDecimals(-1) },
      { label: 'More decimals', onclick: () => stepDecimals(1) },
      { label: 'Show formulas', key: 'Ctrl+`', checked: A.showFormulas, onclick: () => ((A.showFormulas = !A.showFormulas), schedule()) },
      { sep: true },
      { label: 'Data from a database…', icon: 'database', onclick: () => App.openDbImport() },
      { label: 'Send selection to a table…', icon: 'upload', onclick: () => (A.conns.length ? App.openSheetToTable() : toast('Add a database connection first.', 'warn')) },
      { label: 'Import pasted CSV…', icon: 'upload', onclick: () => App.openImport() },
      { label: 'Copy sheet as CSV…', icon: 'download', onclick: () => App.openExport() },
    ], { minWidth: 230 })
  const bar = h(
    '.toolbar',
    mk(btn('undo', 'Undo (Ctrl+Z)', undo), btn('redo', 'Redo (Ctrl+Y)', redo)),
    sep(),
    mk(nf, btn('dollar', 'Currency', () => setFormat('$#,##0.00')), btn('percent', 'Percent', () => setFormat('0%'))),
    sep(),
    mk(btn('bold', 'Bold (Ctrl+B)', () => App.toggleStyle('b'), { active: () => curStyle().b }), btn('italic', 'Italic (Ctrl+I)', () => App.toggleStyle('i'), { active: () => curStyle().i }), btn('underline', 'Underline (Ctrl+U)', () => App.toggleStyle('u'), { active: () => curStyle().u })),
    sep(),
    mk(
      (fgBtn = btn('textcolor', 'Text colour', () => colorMenu(fgBtn, curStyle().fg, (c) => styleRange(norm(), { fg: c })))),
      (bgBtn = btn('fill', 'Fill colour', () => colorMenu(bgBtn, curStyle().bg, (c) => styleRange(norm(), { bg: c })))),
      btn('borders', 'Borders', (e) => bordersMenu(e.currentTarget)),
      btn('merge', 'Merge cells', toggleMerge, { active: () => !!mergeOf(A.sel.ar, A.sel.ac) }),
    ),
    sep(),
    mk(btn('left', 'Align left', () => styleRange(norm(), { al: curStyle().al === 'left' ? null : 'left' }), { active: () => curStyle().al === 'left' }), btn('center', 'Centre', () => styleRange(norm(), { al: curStyle().al === 'center' ? null : 'center' }), { active: () => curStyle().al === 'center' }), btn('right', 'Align right', () => styleRange(norm(), { al: curStyle().al === 'right' ? null : 'right' }), { active: () => curStyle().al === 'right' }), btn('wrap', 'Wrap text', () => App.toggleStyle('wrap'), { active: () => curStyle().wrap })),
    sep(),
    mk(btn('sortasc', 'Sort…', () => App.openSort()), btn('filter', 'Filter', toggleFilter, { active: () => !!sheet().filter }), btn('freeze', 'Freeze panes', (e) => freezeMenu(e.currentTarget), { active: () => !!(sheet().freeze.rows || sheet().freeze.cols) }), btn('chart', 'Insert chart', () => App.openChart())),
    h('span.grow'),
    (moreBtn = btn('more', 'More', more)),
  )
  return bar
}

// ── formula bar ──────────────────────────────────────────────────────────────────────────────────────────────────────────
let nameBox
let formulaInput
let hint
function buildFormulaBar() {
  nameBox = h('input.name-box', {
    spellcheck: false,
    onkeydown: (e) => {
      if (e.key === 'Enter') {
        goTo(nameBox.value)
        scroller.focus()
      }
      if (e.key === 'Escape') {
        App.onSelect()
        scroller.focus()
      }
    },
    onfocus: (e) => e.target.select(),
  })
  formulaInput = h('textarea.formula-input', {
    rows: 1,
    spellcheck: false,
    placeholder: '',
    onfocus: () => {
      if (!A.edit) {
        startEdit(XL.rawText(sheet().cells.get(key(A.sel.ar, A.sel.ac))), 'edit')
        formulaInput.focus()
      }
      A.edit && (A.edit.fromBar = true)
    },
    oninput: () => {
      if (!A.edit) return
      editor.value = formulaInput.value
      onEditInput()
      growBar()
    },
    onkeydown: (e) => {
      if (App.autocompleteKey(e)) return
      if (e.key === 'Enter' && !e.altKey) {
        e.preventDefault()
        commitEdit([e.shiftKey ? -1 : 1, 0])
      } else if (e.key === 'Escape') cancelEdit()
      else if (e.key === 'Tab') {
        e.preventDefault()
        commitEdit([0, 1])
      } else if (e.key === 'Enter' && e.altKey) {
        e.preventDefault()
        formulaInput.setRangeText('\n', formulaInput.selectionStart, formulaInput.selectionEnd, 'end')
        formulaInput.dispatchEvent(new Event('input'))
      }
    },
    onkeyup: updateHint,
    onclick: updateHint,
  })
  hint = h('.fn-hint', { style: { display: 'none' } })
  return h('.formulabar', nameBox, h('button.fx', { title: 'Insert function', onmousedown: (e) => e.preventDefault(), onclick: () => App.openFunctions() }, 'fx'), h('.fbar-input', formulaInput, hint))
}
const growBar = () => {
  formulaInput.style.height = 'auto'
  formulaInput.style.height = `${clamp(formulaInput.scrollHeight, 26, 120)}px`
}
function goTo(text) {
  const t = text.trim()
  const m = t.match(/^(?:(?:'((?:[^']|'')+)'|([^'!]+))!)?(.+)$/)
  let name = t
  if (m && m[3] !== t) {
    const si = A.book.sheetIndex((m[1] ?? m[2]).replace(/''/g, "'"))
    if (si >= 0) switchSheet(si)
    name = m[3]
  }
  const g = XL.parseRange(name)
  if (g) return setSel(g.r1, g.c1, g.r2, g.c2)
  const def = Object.keys(A.book.names).find((k) => k.toLowerCase() === name.toLowerCase())
  if (def) {
    const ref = String(A.book.names[def]).replace(/^=/, '')
    const mm = ref.match(/^(?:(?:'((?:[^']|'')+)'|([^'!]+))!)?(.+)$/)
    const si = mm && (mm[1] || mm[2]) ? A.book.sheetIndex((mm[1] ?? mm[2]).replace(/''/g, "'")) : A.si
    if (si >= 0 && si !== A.si) switchSheet(si)
    const gg = XL.parseRange(mm ? mm[3] : ref)
    if (gg) return setSel(gg.r1, gg.c1, gg.r2, gg.c2)
  }
  toast(`"${t}" is not a cell, range or name.`, 'warn')
}
// signature hint for the function under the caret
function updateHint() {
  const text = A.edit ? editor.value : ''
  const el = A.edit?.fromBar ? formulaInput : editor
  if (!text.startsWith('=')) return void (hint.style.display = 'none')
  const pos = el.selectionStart ?? text.length
  const before = text.slice(0, pos)
  let depth = 0
  let name = ''
  let argIndex = 0
  const stack = []
  for (let i = 0; i < before.length; i++) {
    const ch = before[i]
    if (ch === '"') {
      const j = before.indexOf('"', i + 1)
      if (j < 0) break
      i = j
    } else if (ch === '(') {
      const m = before.slice(0, i).match(/([A-Za-z][\w.]*)$/)
      stack.push({ name: m ? m[1].toUpperCase() : '', arg: 0 })
      depth++
    } else if (ch === ')') {
      stack.pop()
      depth--
    } else if (ch === ',' && stack.length) stack[stack.length - 1].arg++
  }
  void depth
  const top = stack[stack.length - 1]
  if (!top?.name) return void (hint.style.display = 'none')
  name = top.name
  argIndex = top.arg
  const info = FUNC_HELP[name]
  if (!info) return void (hint.style.display = 'none')
  const inner = info.sig.slice(name.length + 1, -1).split(', ')
  const parts = inner.map((a, i) => h(i === Math.min(argIndex, inner.length - 1) ? 'b' : 'span', a))
  hint.style.display = 'block'
  hint.replaceChildren(name, '(', ...parts.flatMap((p, i) => (i ? [', ', p] : [p])), ')', h('span.fn-desc', ` ${info.desc}`))
}

// autocomplete popup for function names
let acEl = null
let acItems = []
let acIndex = 0
let acTouched2 = false
function closeAc() {
  acEl?.remove()
  acEl = null
  acItems = []
}
function updateAutocomplete(text) {
  closeAc()
  if (!A.edit || !text?.startsWith('=')) return
  const el = A.edit.fromBar ? formulaInput : editor
  const pos = el.selectionStart ?? text.length
  const m = text.slice(0, pos).match(/(?:^|[=+\-*/^&,(<>;:\s])([A-Za-z_][\w.]*)$/)
  if (!m || m[1].length < 1) return
  const q = m[1].toUpperCase()
  const names = Object.keys(XL.FN).filter((n) => !n.startsWith('__') && n.startsWith(q)).sort((a, b) => (a.length - b.length) || a.localeCompare(b)).slice(0, 8)
  const defs = Object.keys(A.book.names).filter((n) => n.toUpperCase().startsWith(q))
  acItems = [...names.map((n) => ({ n, fn: true })), ...defs.map((n) => ({ n, fn: false }))].slice(0, 9)
  if (!acItems.length || (acItems.length === 1 && acItems[0].n.toUpperCase() === q && !text.slice(pos).startsWith('('))) {
    if (!acItems.length || acItems[0].n.toUpperCase() === q) return
  }
  acIndex = 0
  acTouched2 = false
  const r = el.getBoundingClientRect()
  acEl = h(
    '.ac',
    { style: { left: `${r.left}px`, top: `${r.bottom + 2}px` }, onmousedown: (e) => e.preventDefault() },
    acItems.map((it, i) => h('.ac-item', { class: i === 0 ? 'on' : '', onclick: () => acceptAc(i) }, h('span.ac-name', it.n), h('span.ac-desc', it.fn ? FUNC_HELP[it.n]?.desc || 'Function' : 'Name'))),
  )
  document.body.append(acEl)
}
function acceptAc(i) {
  const it = acItems[i]
  if (!it) return
  const el = A.edit.fromBar ? formulaInput : editor
  const pos = el.selectionStart
  const m = el.value.slice(0, pos).match(/([A-Za-z_][\w.]*)$/)
  const start = pos - m[1].length
  const ins = it.fn ? `${it.n}(` : it.n
  el.value = el.value.slice(0, start) + ins + el.value.slice(pos)
  el.setSelectionRange(start + ins.length, start + ins.length)
  editor.value = el.value
  formulaInput.value = el.value
  closeAc()
  onEditInput()
  el.focus()
}
App.autocompleteKey = (e) => {
  if (!acEl) return false
  if (e.key === 'ArrowDown' || e.key === 'ArrowUp') {
    e.preventDefault()
    acTouched2 = true
    acIndex = (acIndex + (e.key === 'ArrowDown' ? 1 : -1) + acItems.length) % acItems.length
    $$('.ac-item', acEl).forEach((x, i) => x.classList.toggle('on', i === acIndex))
    return true
  }
  if (e.key === 'Tab' || (e.key === 'Enter' && acTouched2)) {
    e.preventDefault()
    acceptAc(acIndex)
    return true
  }
  if (e.key === 'Escape') {
    e.preventDefault()
    closeAc()
    return true
  }
  return false
}
App.onEditInput = (text) => {
  if (text === null || text === undefined) {
    formulaInput.value = XL.rawText(sheet().cells.get(key(A.sel.ar, A.sel.ac)))
    closeAc()
    hint.style.display = 'none'
    setModeText('Ready')
    growBar()
    return
  }
  if (formulaInput.value !== text) formulaInput.value = text
  growBar()
  updateAutocomplete(text)
  updateHint()
  setModeText(A.edit?.mode === 'enter' ? 'Enter' : 'Edit')
}
const setModeText = (t) => {
  const el = $('#mode-text')
  if (el) el.textContent = t
}

// ── sheet tabs, status bar ───────────────────────────────────────────────────────────────────────────────────────────────
let tabsEl
let statsEl
function buildTabs() {
  tabsEl = h('.tabs')
  statsEl = h('.stats')
  const zoom = h('input.zoom', { type: 'range', min: 50, max: 200, step: 10, value: 100, title: 'Zoom', oninput: (e) => App.setZoom(Number(e.target.value) / 100) })
  App.zoomInput = zoom
  return h('.statusbar', h('button.tb.add-sheet', { title: 'Add sheet', onclick: addSheet }, icon('plus', 16)), tabsEl, h('span.grow'), h('span#mode-text.mode', 'Ready'), statsEl, h('.zoomctl', h('span#zoom-pct.zoom-pct', '100%'), zoom))
}
App.setZoom = (z) => {
  A.zoom = Math.round(z * 10) / 10
  $('#zoom-pct').textContent = `${Math.round(A.zoom * 100)}%`
  App.zoomInput.value = Math.round(A.zoom * 100)
  buildAxes()
  schedule()
}
function renderTabs() {
  tabsEl.replaceChildren(
    ...A.book.sheets.map((s, i) =>
      h(
        '.tab',
        {
          class: i === A.si ? 'on' : '',
          style: s.color ? { boxShadow: `inset 0 -3px 0 ${s.color}` } : {},
          onclick: () => switchSheet(i),
          ondblclick: () => renameSheet(i),
          oncontextmenu: (e) => {
            e.preventDefault()
            sheetMenu(i, e.clientX, e.clientY)
          },
        },
        s.name,
      ),
    ),
  )
}
function sheetMenu(i, x, y) {
  const s = A.book.sheets[i]
  menu(
    [
      { label: 'Insert sheet', icon: 'plus', onclick: addSheet },
      { label: 'Rename', icon: 'edit', onclick: () => renameSheet(i) },
      { label: 'Duplicate', icon: 'copy', onclick: () => applyOps([{ op: 'addSheet', index: i + 1, sheet: { ...A.book.sheetJSON(s), id: undefined, name: `${s.name} copy` } }], { label: 'Duplicate sheet' }) },
      { label: 'Tab colour…', icon: 'fill', onclick: () => colorMenu(tabsEl.children[i], s.color, (c) => applyOps([{ op: 'prop', sheet: s.id, key: 'color', value: c || '' }], { label: 'Tab colour' })) },
      { label: 'Move left', disabled: i === 0, onclick: () => applyOps([{ op: 'moveSheet', id: s.id, index: i - 1 }], { label: 'Move sheet' }) },
      { label: 'Move right', disabled: i === A.book.sheets.length - 1, onclick: () => applyOps([{ op: 'moveSheet', id: s.id, index: i + 1 }], { label: 'Move sheet' }) },
      { sep: true },
      {
        label: 'Delete',
        icon: 'trash',
        disabled: A.book.sheets.length < 2,
        onclick: async () => {
          if (await confirmDialog('Delete sheet', `Delete "${s.name}" and everything on it? Formulas that point to it will show #REF!.`)) applyOps([{ op: 'removeSheet', id: s.id }], { label: 'Delete sheet' })
        },
      },
    ],
    x,
    y,
  )
}
function addSheet() {
  applyOps([{ op: 'addSheet', sheet: { name: `Sheet${A.book.sheets.length + 1}` } }], { label: 'Add sheet' })
  switchSheet(A.book.sheets.length - 1)
}
async function renameSheet(i) {
  const s = A.book.sheets[i]
  const name = await promptDialog('Rename sheet', s.name, 'Rename')
  if (name && name !== s.name) applyOps([{ op: 'renameSheet', id: s.id, name }], { label: 'Rename sheet' })
}
function switchSheet(i) {
  if (i === A.si || i < 0 || i >= A.book.sheets.length) return
  if (A.edit) commitEdit()
  view().sx = scroller.scrollLeft
  view().sy = scroller.scrollTop
  A.si = i
  A.findHits = []
  buildAxes()
  const v = view()
  A.sel = { ...v.sel }
  scroller.scrollLeft = v.sx
  scroller.scrollTop = v.sy
  App.refreshAll()
  schedule()
}
function updateStats() {
  const g = norm()
  if (g.r1 === g.r2 && g.c1 === g.c2) return void (statsEl.textContent = '')
  const used = A.book.usedRange(sheet().id)
  if (!used) return void (statsEl.textContent = '')
  const r2 = Math.min(g.r2, used.r2)
  const c2 = Math.min(g.c2, used.c2)
  if ((r2 - g.r1 + 1) * (c2 - g.c1 + 1) > 300000) return void (statsEl.textContent = '')
  let n = 0
  let sum = 0
  let min = Infinity
  let max = -Infinity
  let cnt = 0
  for (let r = g.r1; r <= r2; r++)
    for (let c = g.c1; c <= c2; c++) {
      const v = A.book.get(A.si, r, c)
      if (v === null || v === undefined || v === '') continue
      cnt++
      if (typeof v === 'number') {
        n++
        sum += v
        min = Math.min(min, v)
        max = Math.max(max, v)
      }
    }
  const f = (v) => (Number.isInteger(v) ? v.toLocaleString() : (+v.toFixed(4)).toLocaleString())
  statsEl.replaceChildren(...(n ? [['Average', sum / n], ['Count', cnt], ['Sum', sum], ['Min', min], ['Max', max]] : [['Count', cnt]]).map(([k, v]) => h('span.stat', k, ': ', h('b', f(v)))))
}

// ── hooks the grid calls ─────────────────────────────────────────────────────────────────────────────────────────────────
App.onSelect = () => {
  const g = norm()
  const single = g.r1 === g.r2 && g.c1 === g.c2
  if (nameBox && document.activeElement !== nameBox) nameBox.value = single || mergeOf(A.sel.ar, A.sel.ac) ? A1(A.sel.ar, A.sel.ac) : XL.rangeText(g)
  if (!A.edit && formulaInput) {
    formulaInput.value = XL.rawText(sheet().cells.get(key(A.sel.ar, A.sel.ac)))
    growBar()
  }
  updateStats()
  refreshers.forEach((f) => f())
}
App.onChange = () => {
  renderTabs()
  App.onSelect()
}
App.refreshAll = () => {
  renderTabs()
  App.onSelect()
  App.refreshTop?.()
}

// ── context menu ─────────────────────────────────────────────────────────────────────────────────────────────────────────
App.contextMenu = (x, y, { onHeaderCol, onHeaderRow }) => {
  const g = norm()
  const nR = g.r2 - g.r1 + 1
  const nC = g.c2 - g.c1 + 1
  const wholeCols = g.r1 === 0 && g.r2 >= MAXR - 1
  const wholeRows = g.c1 === 0 && g.c2 >= MAXC - 1
  const items = [
    { label: 'Cut', icon: 'copy', key: 'Ctrl+X', onclick: () => document.execCommand('cut') || doCopy(null, true) },
    { label: 'Copy', icon: 'copy', key: 'Ctrl+C', onclick: () => document.execCommand('copy') || doCopy(null, false) },
    { label: 'Paste', icon: 'clipboard', key: 'Ctrl+V', disabled: !A.clip, onclick: () => doPaste() },
    { label: 'Paste values only', disabled: !A.clip, onclick: () => doPaste(undefined, 'values') },
    { label: 'Paste formatting only', disabled: !A.clip, onclick: () => doPaste(undefined, 'formats') },
    { sep: true },
  ]
  const rowsLabel = nR > 1 ? `${nR} rows` : 'row'
  const colsLabel = nC > 1 ? `${nC} columns` : 'column'
  if (!wholeCols) items.push({ label: `Insert ${rowsLabel} above`, onclick: () => structural('insert', 'row', g.r1, nR) }, { label: `Delete ${rowsLabel}`, onclick: () => structural('delete', 'row', g.r1, nR) })
  if (!wholeRows) items.push({ label: `Insert ${colsLabel} left`, onclick: () => structural('insert', 'col', g.c1, nC) }, { label: `Delete ${colsLabel}`, onclick: () => structural('delete', 'col', g.c1, nC) })
  items.push({ sep: true })
  if (onHeaderRow || wholeRows) items.push({ label: `Hide ${rowsLabel}`, onclick: () => hideAxis('hiddenRows', g.r1, g.r2) }, { label: 'Unhide rows', onclick: () => applyOps([{ op: 'prop', sheet: sheet().id, key: 'hiddenRows', value: [] }], { label: 'Unhide' }) })
  if (onHeaderCol || wholeCols) items.push({ label: `Hide ${colsLabel}`, onclick: () => hideAxis('hiddenCols', g.c1, g.c2) }, { label: 'Unhide columns', onclick: () => applyOps([{ op: 'prop', sheet: sheet().id, key: 'hiddenCols', value: [] }], { label: 'Unhide' }) })
  if (onHeaderCol) items.push({ label: 'Fit column to contents', onclick: () => autofit('col', g.c1) })
  if (onHeaderRow) items.push({ label: 'Fit row to contents', onclick: () => autofit('row', g.r1) })
  items.push(
    { sep: true },
    { label: 'Clear contents', key: 'Del', onclick: () => clearRange(g) },
    { label: 'Clear formatting', onclick: () => clearRange(g, 'formats') },
    { label: 'Format cells…', key: 'Ctrl+1', onclick: () => App.openFormatCells() },
    { label: 'Sort A to Z', icon: 'sortasc', onclick: () => sortBy(true) },
    { label: 'Sort Z to A', icon: 'sortdesc', onclick: () => sortBy(false) },
    { label: 'Create chart…', icon: 'chart', onclick: () => App.openChart() },
  )
  menu(items, x, y)
}
function structural(kind, dim, at, n) {
  applyOps([{ op: kind, sheet: sheet().id, dim, at, n }], { label: `${kind === 'insert' ? 'Insert' : 'Delete'} ${dim === 'row' ? 'rows' : 'columns'}` })
}
function hideAxis(prop, a, b) {
  const sh = sheet()
  applyOps([{ op: 'prop', sheet: sh.id, key: prop, value: [...new Set([...sh[prop], ...Array.from({ length: Math.min(b - a + 1, 5000) }, (_, i) => a + i)])] }], { label: 'Hide' })
}

// ── dialogs ──────────────────────────────────────────────────────────────────────────────────────────────────────────────
App.openFormatCells = (tab = 'number') => {
  const g = norm()
  const s0 = { ...curStyle() }
  const draft = { ...s0 }
  const tabs = ['number', 'alignment', 'font', 'fill', 'border']
  const body = h('.fc')
  const preview = h('.fc-preview')
  let cur = tab
  const sample = () => A.book.get(A.si, A.sel.ar, A.sel.ac)
  const refreshPreview = () => {
    const v = sample()
    preview.textContent = v === null || v === undefined || v === '' ? XL.formatValue(1234.5678, draft.nf) : XL.formatValue(v, draft.nf)
    preview.style.fontWeight = draft.b ? '700' : '400'
    preview.style.fontStyle = draft.i ? 'italic' : 'normal'
    preview.style.textDecoration = [draft.u && 'underline', draft.st && 'line-through'].filter(Boolean).join(' ')
    preview.style.color = draft.fg || ''
    preview.style.background = draft.bg || ''
    preview.style.textAlign = draft.al || ''
  }
  const draw = () => {
    body.replaceChildren()
    const tabBar = h('.fc-tabs', tabs.map((t) => h('button.fc-tab', { class: t === cur ? 'on' : '', onclick: () => ((cur = t), draw()) }, t[0].toUpperCase() + t.slice(1))))
    let pane
    if (cur === 'number') {
      const cats = [['General', 'General'], ['Number', '0.00'], ['Currency', '$#,##0.00'], ['Accounting', '#,##0.00;(#,##0.00)'], ['Percentage', '0.00%'], ['Date', 'yyyy-mm-dd'], ['Time', 'hh:mm:ss'], ['Scientific', '0.00E+00'], ['Text', '@'], ['Custom', draft.nf || '']]
      const code = h('input.input.mono', { value: draft.nf || 'General', style: { width: '100%' }, oninput: (e) => ((draft.nf = e.target.value === 'General' ? null : e.target.value), refreshPreview()) })
      const samples = {
        Number: ['0', '0.00', '#,##0', '#,##0.00', '0.000'],
        Currency: ['$#,##0', '$#,##0.00', '€#,##0.00', '£#,##0.00', '"Rs "#,##0', '$#,##0.00;[Red]($#,##0.00)'],
        Percentage: ['0%', '0.0%', '0.00%'],
        Date: ['m/d/yyyy', 'yyyy-mm-dd', 'd-mmm-yy', 'mmm d, yyyy', 'dddd, mmmm d, yyyy', 'mmmm yyyy'],
        Time: ['h:mm AM/PM', 'hh:mm', 'hh:mm:ss', '[h]:mm:ss'],
        Scientific: ['0.00E+00', '0.0E+00'],
        Text: ['@'],
        Accounting: ['#,##0.00;(#,##0.00)', '$#,##0.00;($#,##0.00)'],
        General: ['General'],
        Custom: ['#,##0.0,"K"', '0.0,,"M"', '[>=1000000]0.0,,"M";[>=1000]0.0,"K";0', '#,##0;[Red]-#,##0'],
      }
      const list = h('.fc-list')
      const examples = h('.fc-examples')
      const showEx = (name) => {
        examples.replaceChildren(...(samples[name] || []).map((f) => h('button.fc-ex', { onclick: () => ((draft.nf = f === 'General' ? null : f), (code.value = f), refreshPreview()) }, h('code', f), h('span', XL.formatValue(sample() ?? 1234.5678, f)))))
      }
      let cat = (cats.find((c) => c[1] === (draft.nf || 'General')) || cats[9])[0]
      const paintCats = () => {
        list.replaceChildren(...cats.map(([n]) => h('button.fc-cat', { class: n === cat ? 'on' : '', onclick: () => ((cat = n), paintCats(), showEx(n)) }, n)))
      }
      paintCats()
      showEx(cat)
      pane = h('.fc-pane.num', list, h('div', field('Format code', code, 'Examples below. Section separator ;  colour [Red]  text "…"'), examples))
    } else if (cur === 'alignment') {
      pane = h(
        '.fc-pane.col',
        field('Horizontal', sel([['', 'General'], ['left', 'Left'], ['center', 'Centre'], ['right', 'Right']], draft.al || '', { onchange: (e) => ((draft.al = e.target.value || null), refreshPreview()) })),
        field('Vertical', sel([['', 'Middle (default)'], ['top', 'Top'], ['middle', 'Middle'], ['bottom', 'Bottom']], draft.va || '', { onchange: (e) => (draft.va = e.target.value || null) })),
        h('label.check', h('input', { type: 'checkbox', checked: !!draft.wrap, onchange: (e) => (draft.wrap = e.target.checked || null) }), 'Wrap text'),
      )
    } else if (cur === 'font') {
      pane = h(
        '.fc-pane.col',
        field('Size', sel([8, 9, 10, 11, 12, 13, 14, 16, 18, 20, 24, 28, 36], draft.fs || 13, { onchange: (e) => (draft.fs = Number(e.target.value) === 13 ? null : Number(e.target.value)) })),
        h('.fc-row', ...[['b', 'Bold'], ['i', 'Italic'], ['u', 'Underline'], ['st', 'Strikethrough']].map(([k, l]) => h('label.check', h('input', { type: 'checkbox', checked: !!draft[k], onchange: (e) => ((draft[k] = e.target.checked || null), refreshPreview()) }), l))),
        field('Text colour', swatchField(draft.fg, (c) => ((draft.fg = c), refreshPreview(), draw()))),
      )
    } else if (cur === 'fill') {
      pane = h('.fc-pane.col', field('Background colour', swatchField(draft.bg, (c) => ((draft.bg = c), refreshPreview(), draw()))))
    } else {
      const bd = (draft.bd = { ...(draft.bd || {}) })
      const toggle = (k) => ((bd[k] = bd[k] ? null : 'thin'), draw())
      pane = h(
        '.fc-pane.col',
        h('.fc-row', [['t', 'Top'], ['b', 'Bottom'], ['l', 'Left'], ['r', 'Right']].map(([k, l]) => h('label.check', h('input', { type: 'checkbox', checked: !!bd[k], onchange: () => toggle(k) }), l))),
        h('.fc-row', h('button.btn', { onclick: () => ((draft.bd = { t: 'thin', b: 'thin', l: 'thin', r: 'thin' }), draw()) }, 'All'), h('button.btn', { onclick: () => ((draft.bd = null), draw()) }, 'None')),
        field('Style', sel([['thin', 'Thin'], ['thick', 'Thick'], ['dashed', 'Dashed']], Object.values(bd).find(Boolean) || 'thin', { onchange: (e) => Object.keys(bd).forEach((k) => bd[k] && (bd[k] = e.target.value)) })),
        field('Colour', swatchField(draft.bc, (c) => ((draft.bc = c), draw()))),
      )
    }
    body.append(tabBar, pane, h('.fc-prev-row', h('span.hint', 'Preview'), preview))
    refreshPreview()
  }
  draw()
  dialog({
    title: 'Format cells',
    width: 560,
    body,
    buttons: [
      { label: 'Cancel' },
      {
        label: 'Apply',
        primary: true,
        onclick: () => {
          const patch = {}
          for (const k of ['nf', 'al', 'va', 'wrap', 'fs', 'b', 'i', 'u', 'st', 'fg', 'bg', 'bd', 'bc']) {
            let v = draft[k]
            if (k === 'bd' && v && !Object.values(v).some(Boolean)) v = null
            if (k === 'bd' && v) v = Object.fromEntries(Object.entries(v).filter(([, x]) => x))
            if (JSON.stringify(v ?? null) !== JSON.stringify(s0[k] ?? null)) patch[k] = v ?? null
          }
          if (Object.keys(patch).length) styleRange(g, patch)
        },
      },
    ],
  })
}
function swatchField(current, onPick) {
  const b = h('button.btn.swatch-btn', { onclick: () => colorMenu(b, current, onPick) }, h('span.swatch', { style: { background: current || 'transparent' } }), current || 'None')
  return b
}

App.openFunctions = () => {
  const list = h('.fn-list')
  const detail = h('.fn-detail', h('span.hint', 'Pick a function'))
  const all = Object.keys(XL.FN).filter((n) => !n.startsWith('__')).sort()
  const search = h('input.input', { placeholder: 'Search functions', style: { width: '100%' }, oninput: paint })
  function paint() {
    const q = search.value.trim().toUpperCase()
    const hits = all.filter((n) => n.includes(q) || (FUNC_HELP[n]?.desc || '').toUpperCase().includes(q)).slice(0, 80)
    list.replaceChildren(...hits.map((n) => h('button.fn-row', { onclick: () => show(n), ondblclick: () => insert(n) }, n)))
  }
  const show = (n) => {
    const i = FUNC_HELP[n]
    detail.replaceChildren(h('code.sig', i ? i.sig : `${n}(…)`), h('p', i ? i.desc : 'See the function name; this engine follows the usual spreadsheet behaviour.'), h('button.btn.primary', { onclick: () => insert(n) }, 'Insert'))
  }
  const insert = (n) => {
    d.close()
    if (!A.edit) startEdit('=', 'edit')
    const el = editor
    const p = el.selectionStart
    el.setRangeText(`${n}(`, p, el.selectionEnd, 'end')
    onEditInput()
    el.focus()
  }
  paint()
  const d = dialog({ title: 'Insert function', width: 600, body: h('div', search, h('.fn-wrap', list, detail)) })
}

App.openFind = (replaceMode) => {
  let hits = []
  let idx = -1
  const find = h('input.input', { placeholder: 'Find', style: { width: '100%' } })
  const rep = h('input.input', { placeholder: 'Replace with', style: { width: '100%' } })
  const scope = sel([['sheet', 'This sheet'], ['book', 'Whole workbook']], 'sheet')
  const mc = h('input', { type: 'checkbox' })
  const whole = h('input', { type: 'checkbox' })
  const inF = h('input', { type: 'checkbox' })
  const info = h('.hint', '')
  const run = () => {
    if (!find.value) return ((hits = []), (A.findHits = []), (info.textContent = ''), schedule())
    const plan = A.book.planFind(scope.value === 'sheet' ? sheet().id : null, find.value, null, { matchCase: mc.checked, whole: whole.checked, formulas: inF.checked })
    hits = plan.matches
    A.findHits = hits.filter((m) => m.sheet === sheet().name).map((m) => {
      const p = XL.parseAddr(m.cell)
      return [p.r, p.c]
    })
    info.textContent = `${hits.length} match${hits.length === 1 ? '' : 'es'}`
    schedule()
  }
  const go = (dir) => {
    run()
    if (!hits.length) return
    idx = (idx + dir + hits.length) % hits.length
    const m = hits[idx]
    const si = A.book.sheetIndex(m.sheet)
    if (si !== A.si) switchSheet(si)
    const p = XL.parseAddr(m.cell)
    setSel(p.r, p.c)
    info.textContent = `${idx + 1} of ${hits.length}`
  }
  const doReplace = (all) => {
    const plan = A.book.planFind(scope.value === 'sheet' ? sheet().id : null, find.value, rep.value, { matchCase: mc.checked, whole: whole.checked, formulas: inF.checked, all })
    if (plan.ops.length) applyOps(plan.ops, { label: 'Replace' })
    toast(`${plan.matches.length} replaced`)
    run()
  }
  find.addEventListener('input', run)
  find.addEventListener('keydown', (e) => e.key === 'Enter' && go(e.shiftKey ? -1 : 1))
  const d = dialog({
    title: replaceMode ? 'Find and replace' : 'Find',
    width: 440,
    onclose: () => ((A.findHits = []), schedule()),
    body: h('.col', field('Find', find), replaceMode ? field('Replace with', rep) : '', field('Look in', scope), h('.fc-row', h('label.check', mc, 'Match case'), h('label.check', whole, 'Entire cell'), h('label.check', inF, 'Formulas')), info, h('.btn-row', h('button.btn', { onclick: () => go(-1) }, 'Previous'), h('button.btn', { onclick: () => go(1) }, 'Next'), replaceMode ? [h('button.btn', { onclick: () => doReplace(false) }, 'Replace'), h('button.btn.primary', { onclick: () => doReplace(true) }, 'Replace all')] : '')),
  })
  void d
}

App.openNames = () => {
  const draw = () => {
    box.replaceChildren(
      h('.names-list', Object.keys(A.book.names).length ? Object.entries(A.book.names).map(([n, ref]) => h('.name-row', h('b', n), h('code', ref), h('button.icon-btn', { title: 'Delete', onclick: () => ((delete A.book.names[n]), applyOps([{ op: 'names', value: { ...A.book.names } }], { label: 'Delete name' }), draw()) }, icon('trash', 14)))) : h('p.hint', 'No names yet. A name stands for a range or value, so formulas read =SUM(Sales) instead of =SUM(B2:B99).')),
      h('.name-new', nm, rf, h('button.btn.primary', { onclick: add }, 'Add')),
    )
  }
  const nm = h('input.input', { placeholder: 'Name', style: { width: '36%' } })
  const g0 = norm()
  const rf = h('input.input.mono', { placeholder: 'Sheet1!$A$1:$A$10', value: `${/^[A-Za-z_]\w*$/.test(sheet().name) ? sheet().name : `'${sheet().name}'`}!${XL.rangeText(g0).replace(/([A-Z]+)(\d+)/g, '$$$1$$$2')}`, style: { width: '56%' } })
  const add = () => {
    if (!/^[A-Za-z_][\w.]*$/.test(nm.value) || /^[A-Za-z]{1,3}\d+$/.test(nm.value)) return toast('A name starts with a letter and has no spaces.', 'warn')
    applyOps([{ op: 'names', value: { ...A.book.names, [nm.value]: rf.value.replace(/^=/, '') } }], { label: 'Define name' })
    nm.value = ''
    draw()
  }
  const box = h('div')
  draw()
  dialog({ title: 'Named ranges', width: 520, body: box, buttons: [{ label: 'Done', primary: true }] })
}

App.openSort = () => {
  const g = selectionRegion()
  const first = Array.from({ length: g.c2 - g.c1 + 1 }, (_, i) => A.book.display(A.si, g.r1, g.c1 + i))
  const header = h('input', { type: 'checkbox', checked: true })
  const rows = []
  const box = h('div')
  const add = () => rows.push({ col: 0, asc: true })
  add()
  const draw = () => {
    box.replaceChildren(
      ...rows.map((k, i) => h('.sort-row', h('span.hint', i ? 'Then by' : 'Sort by'), sel(first.map((n, j) => [j, header.checked ? n || `Column ${XL.colName(g.c1 + j)}` : `Column ${XL.colName(g.c1 + j)}`]), k.col, { onchange: (e) => (k.col = Number(e.target.value)) }), sel([['a', 'A to Z'], ['d', 'Z to A']], k.asc ? 'a' : 'd', { onchange: (e) => (k.asc = e.target.value === 'a') }))),
      h('button.btn', { onclick: () => (add(), draw()) }, 'Add level'),
    )
  }
  draw()
  dialog({
    title: 'Sort',
    width: 440,
    body: h('div', h('label.check', header, 'My data has headers'), box),
    buttons: [{ label: 'Cancel' }, { label: 'Sort', primary: true, onclick: () => (applyOps([A.book.planSort(sheet().id, g, rows, { header: header.checked })], { label: 'Sort' }), setSel(g.r1, g.c1, g.r2, g.c2)) }],
  })
}

// conditional formatting
App.openCF = () => {
  const sh = sheet()
  const g = norm()
  const box = h('div')
  const draw = () => {
    const mine = sheet().cf
    box.replaceChildren(
      mine.length
        ? h('.cf-list', mine.map((r, i) => h('.cf-row', h('span.cf-swatch', { style: { background: r.style?.bg || r.colors?.[1] || r.color || 'transparent', color: r.style?.fg || '' } }, 'Aa'), h('div', h('b', describeRule(r)), h('span.hint', ` ${r.range}`)), h('button.icon-btn', { title: 'Remove', onclick: () => applyOps([{ op: 'prop', sheet: sh.id, key: 'cf', value: sheet().cf.filter((_, j) => j !== i) }], { label: 'Remove rule' }) || draw() }, icon('trash', 14)))))
        : h('p.hint', 'No rules on this sheet.'),
      h('button.btn.primary', { onclick: () => newRule() }, 'New rule'),
    )
  }
  const describeRule = (r) => ({ cell: `Cell ${r.op} ${r.a}${r.op?.includes('etween') ? ` and ${r.b}` : ''}`, text: `Text ${r.op} "${r.text}"`, scale: 'Colour scale', bar: 'Data bars', formula: `Formula ${r.formula}`, dup: 'Duplicate values', top: `${r.bottom ? 'Bottom' : 'Top'} ${r.n}`, blank: 'Blank cells', error: 'Errors', avg: `${r.op === 'below' ? 'Below' : 'Above'} average` })[r.type] || r.type
  const newRule = () => {
    const type = sel([['cell', 'Cell value'], ['text', 'Text contains'], ['formula', 'Custom formula'], ['scale', 'Colour scale'], ['bar', 'Data bars'], ['dup', 'Duplicate values'], ['top', 'Top / bottom N'], ['avg', 'Above / below average'], ['blank', 'Blank cells'], ['error', 'Cells with errors']], 'cell')
    const range = h('input.input.mono', { value: XL.rangeText(g), style: { width: '100%' } })
    const op = sel([['>', 'greater than'], ['<', 'less than'], ['>=', 'at least'], ['<=', 'at most'], ['=', 'equal to'], ['<>', 'not equal to'], ['between', 'between'], ['notBetween', 'not between']], '>')
    const a = h('input.input', { placeholder: 'Value', style: { width: '45%' } })
    const b = h('input.input', { placeholder: 'and', style: { width: '45%' } })
    const textOp = sel([['contains', 'contains'], ['notContains', 'does not contain'], ['begins', 'begins with'], ['ends', 'ends with']], 'contains')
    const formula = h('input.input.mono', { placeholder: '=$B2>100', style: { width: '100%' } })
    const n = h('input.input', { type: 'number', value: 10, style: { width: '80px' } })
    const bottom = sel([['top', 'Top'], ['bottom', 'Bottom']], 'top')
    const avgOp = sel([['above', 'Above average'], ['below', 'Below average']], 'above')
    let style = { bg: '#fecaca', fg: '#991b1b' }
    let colors3 = ['#f8696b', '#ffeb84', '#63be7b']
    let barColor = '#5b9bd5'
    const presets = [
      ['Light red', { bg: '#fecaca', fg: '#991b1b' }],
      ['Yellow', { bg: '#fef08a', fg: '#854d0e' }],
      ['Green', { bg: '#bbf7d0', fg: '#166534' }],
      ['Blue', { bg: '#bfdbfe', fg: '#1e3a8a' }],
      ['Bold red', { fg: '#b91c1c', b: true }],
    ]
    const styleBox = h('.preset-row', presets.map(([l, s]) => h('button.btn.preset', { style: { background: s.bg || '', color: s.fg || '' }, onclick: () => (style = s) }, l)))
    const fields = h('div')
    const paint = () => {
      const t = type.value
      fields.replaceChildren(
        t === 'cell' ? h('.fc-row', op, a, b) : '',
        t === 'text' ? h('.fc-row', textOp, h('input.input', { id: 'cf-text', placeholder: 'text', style: { width: '50%' } })) : '',
        t === 'formula' ? formula : '',
        t === 'top' ? h('.fc-row', bottom, n) : '',
        t === 'avg' ? avgOp : '',
        ['cell', 'text', 'formula', 'dup', 'top', 'avg', 'blank', 'error'].includes(t) ? h('div', h('span.hint', 'Format with'), styleBox) : '',
        t === 'scale' ? h('.fc-row', ...colors3.map((c, i) => swatchField(c, (v) => ((colors3[i] = v || c), paint())))) : '',
        t === 'bar' ? swatchField(barColor, (v) => ((barColor = v || barColor), paint())) : '',
      )
    }
    type.addEventListener('change', paint)
    paint()
    dialog({
      title: 'New formatting rule',
      width: 480,
      body: h('.col', field('Rule', type), field('Apply to range', range), fields),
      buttons: [
        { label: 'Cancel' },
        {
          label: 'Add rule',
          primary: true,
          onclick: () => {
            const t = type.value
            const rule = { id: XL.newId(), type: t, range: range.value.replace(/\$/g, '').toUpperCase() }
            if (!XL.parseRange(rule.range)) return toast('That range is not valid.', 'warn') && false
            if (t === 'cell') Object.assign(rule, { op: op.value, a: a.value, b: b.value, style })
            else if (t === 'text') Object.assign(rule, { op: textOp.value, text: $('#cf-text').value, style })
            else if (t === 'formula') Object.assign(rule, { formula: formula.value, style })
            else if (t === 'top') Object.assign(rule, { n: Number(n.value) || 10, bottom: bottom.value === 'bottom', style })
            else if (t === 'avg') Object.assign(rule, { op: avgOp.value, style })
            else if (t === 'scale') rule.colors = colors3
            else if (t === 'bar') rule.color = barColor
            else rule.style = style
            applyOps([{ op: 'prop', sheet: sh.id, key: 'cf', value: [...sheet().cf, rule] }], { label: 'Add rule' })
            draw()
          },
        },
      ],
    })
  }
  draw()
  dialog({ title: 'Conditional formatting', width: 520, body: box, buttons: [{ label: 'Done', primary: true }] })
}

// data validation
App.openDV = () => {
  const sh = sheet()
  const g = norm()
  const existing = sh.dv.find((r) => XL.parseRange(r.range) && XL.inRange(XL.parseRange(r.range), g.r1, g.c1))
  const rule = existing ? { ...existing } : { type: 'list', allowBlank: true, strict: true }
  const type = sel([['any', 'Any value'], ['list', 'List'], ['number', 'Decimal'], ['whole', 'Whole number'], ['date', 'Date'], ['text', 'Text length'], ['custom', 'Custom formula']], existing ? rule.type : 'list')
  const range = h('input.input.mono', { value: existing?.range || XL.rangeText(g) })
  const body = h('.col')
  const items = h('input.input', { value: Array.isArray(rule.items) ? rule.items.join(', ') : rule.source || '', placeholder: 'Open, Closed, Pending   or   Sheet2!A1:A9' })
  const op = sel([['between', 'between'], ['notBetween', 'not between'], ['=', 'equal to'], ['<>', 'not equal to'], ['>', 'greater than'], ['<', 'less than'], ['>=', 'at least'], ['<=', 'at most']], rule.op || 'between')
  const a = h('input.input', { value: rule.a ?? '', placeholder: 'Minimum or value' })
  const b = h('input.input', { value: rule.b ?? '', placeholder: 'Maximum' })
  const formula = h('input.input.mono', { value: rule.formula || '', placeholder: '=ISNUMBER(A1)' })
  const message = h('input.input', { value: rule.message || '', placeholder: 'Shown when a value is rejected' })
  const strict = h('input', { type: 'checkbox', checked: rule.strict !== false })
  const blank = h('input', { type: 'checkbox', checked: rule.allowBlank !== false })
  const paint = () => {
    const t = type.value
    body.replaceChildren(
      field('Apply to', range),
      field('Allow', type),
      t === 'list' ? field('Items', items, 'Comma separated, or a range') : '',
      ['number', 'whole', 'date', 'text'].includes(t) ? h('.col', field('Condition', op), h('.fc-row', a, b)) : '',
      t === 'custom' ? field('Formula', formula, 'Relative to the first cell of the range') : '',
      t !== 'any' ? h('div', field('Error message', message), h('.fc-row', h('label.check', strict, 'Reject invalid values'), h('label.check', blank, 'Ignore blank'))) : '',
    )
  }
  type.addEventListener('change', paint)
  paint()
  dialog({
    title: 'Data validation',
    width: 460,
    body,
    buttons: [
      { label: 'Cancel' },
      existing ? { label: 'Remove', danger: true, onclick: () => applyOps([{ op: 'prop', sheet: sh.id, key: 'dv', value: sh.dv.filter((r) => r !== existing) }], { label: 'Remove validation' }) } : null,
      {
        label: 'Apply',
        primary: true,
        onclick: () => {
          const rest = sh.dv.filter((r) => r !== existing)
          const t = type.value
          if (t === 'any') return applyOps([{ op: 'prop', sheet: sh.id, key: 'dv', value: rest }], { label: 'Validation' })
          const out = { type: t, range: range.value.replace(/\$/g, '').toUpperCase(), message: message.value, strict: strict.checked, allowBlank: blank.checked }
          if (!XL.parseRange(out.range)) return toast('That range is not valid.', 'warn') && false
          if (t === 'list') {
            const v = items.value.trim()
            if (/^=?'?[\w ]+'?!|^=?\$?[A-Za-z]{1,3}\$?\d+(:|$)/.test(v)) out.source = v.replace(/^=/, '')
            else out.items = v.split(',').map((x) => x.trim()).filter(Boolean)
          } else if (t === 'custom') out.formula = formula.value
          else Object.assign(out, { op: op.value, a: a.value, b: b.value })
          applyOps([{ op: 'prop', sheet: sh.id, key: 'dv', value: [...rest, out] }], { label: 'Validation' })
        },
      },
    ].filter(Boolean),
  })
}
App.openListDropdown = (rule, b) => {
  const items = XL.listItems(A.book, A.si, rule).map(String)
  const r = scroller.getBoundingClientRect()
  menu(items.map((v) => ({ label: v, onclick: () => applyOps([opForText(A.sel.ar, A.sel.ac, v)], { label: 'Pick value' }) })), r.left + b.x, r.top + b.y + b.h, { minWidth: Math.max(120, b.w) })
}
// auto filter popover
App.openFilter = (off, b) => {
  const sh = sheet()
  const f = sh.filter
  const g = XL.parseRange(f.range)
  const col = g.c1 + off
  const counts = new Map()
  for (let r = g.r1 + 1; r <= g.r2; r++) {
    const t = A.book.display(A.si, r, col) || '(Blanks)'
    counts.set(t, (counts.get(t) || 0) + 1)
  }
  const values = [...counts.keys()].sort((x, y) => x.localeCompare(y, undefined, { numeric: true }))
  const allowed = new Set(f.cols?.[off]?.values || values)
  const list = h('.flt-list')
  const q = h('input.input', { placeholder: 'Search', style: { width: '100%' }, oninput: paint })
  function paint() {
    const s = q.value.toLowerCase()
    list.replaceChildren(
      ...values.filter((v) => v.toLowerCase().includes(s)).map((v) => h('label.check', h('input', { type: 'checkbox', checked: allowed.has(v), onchange: (e) => (e.target.checked ? allowed.add(v) : allowed.delete(v)) }), h('span.flt-v', v), h('span.hint', counts.get(v)))),
    )
  }
  paint()
  const apply = () => {
    const cols = { ...(f.cols || {}) }
    if (allowed.size === values.length) delete cols[off]
    else cols[off] = { values: [...allowed] }
    closeMenu()
    applyOps([{ op: 'prop', sheet: sh.id, key: 'filter', value: { ...f, cols } }], { label: 'Filter' })
  }
  const m = h(
    '.menu.flt',
    { onmousedown: (e) => e.stopPropagation() },
    h('button.menu-item', { onclick: () => (closeMenu(), applyOps([A.book.planSort(sh.id, { ...g, r1: g.r1 }, [{ col: off, asc: true }], { header: true })], { label: 'Sort' })) }, icon('sortasc', 15), h('span.mi-label', 'Sort A to Z')),
    h('button.menu-item', { onclick: () => (closeMenu(), applyOps([A.book.planSort(sh.id, { ...g, r1: g.r1 }, [{ col: off, asc: false }], { header: true })], { label: 'Sort' })) }, icon('sortdesc', 15), h('span.mi-label', 'Sort Z to A')),
    h('.menu-sep'),
    q,
    h('.flt-all', h('button.link', { onclick: () => (values.forEach((v) => allowed.add(v)), paint()) }, 'Select all'), h('button.link', { onclick: () => (allowed.clear(), paint()) }, 'Clear')),
    list,
    h('.dlg-foot', h('button.btn', { onclick: closeMenu }, 'Cancel'), h('button.btn.primary', { onclick: apply }, 'OK')),
  )
  closeMenu()
  document.body.append(m)
  const r = scroller.getBoundingClientRect()
  m.style.left = `${clamp(r.left + b.x + b.w - 240, 4, innerWidth - 250)}px`
  m.style.top = `${r.top + b.y + b.h}px`
  openMenu = m
  q.focus()
}

// export (copy as CSV) and import
App.openExport = async () => {
  const g = norm()
  const single = g.r1 === g.r2 && g.c1 === g.c2
  try {
    const r = await Host.tool('app_export', { id: A.book.id, sheet: sheet().id, range: single ? undefined : XL.rangeText(g), format: 'csv' })
    download(`${sheet().name}.csv`, r.text)
  } catch (e) {
    toast(e.message, 'err')
  }
}
App.openImport = () => {
  const text = h('textarea.input.mono', { rows: 12, style: { width: '100%' }, placeholder: 'Paste CSV or tab separated data here' })
  const at = h('input.input', { value: A1(A.sel.ar, A.sel.ac), style: { width: '100px' } })
  dialog({
    title: 'Import data',
    width: 600,
    body: h('div', field('Paste data', text), field('Put it at cell', at)),
    buttons: [
      { label: 'Cancel' },
      {
        label: 'Import',
        primary: true,
        onclick: async () => {
          if (!text.value.trim()) return false
          try {
            if (A.pending.length) await Save.flush()
            const r = await Host.tool('app_import', { id: A.book.id, sheet: sheet().id, at: at.value, csv: text.value })
            loadWorkbook(r.workbook, { keepView: true })
            toast('Imported')
          } catch (e) {
            toast(e.message, 'err')
            return false
          }
        },
      },
    ],
  })
}
App.openShortcuts = () => {
  const rows = [['Move', 'Arrow keys, Tab, Enter'], ['Jump to the edge of data', 'Ctrl + Arrow'], ['Select range', 'Shift + Arrow, or drag'], ['Select column / row', 'Ctrl + Space / Shift + Space'], ['Select all', 'Ctrl + A'], ['Edit cell', 'F2 or double-click'], ['New line in a cell', 'Alt + Enter'], ['Cycle $ in a reference', 'F4'], ['Copy, cut, paste', 'Ctrl + C / X / V'], ['Undo, redo', 'Ctrl + Z / Y'], ['Fill down / right', 'Ctrl + D / R'], ['Bold, italic, underline', 'Ctrl + B / I / U'], ['Format cells', 'Ctrl + 1'], ['Find, replace', 'Ctrl + F / H'], ['Show formulas', 'Ctrl + `'], ['Insert today', 'Ctrl + ;'], ['Zoom', 'Ctrl + scroll'], ['Run the query', 'Ctrl + Enter (database)']]
  dialog({ title: 'Keyboard shortcuts', width: 480, body: h('table.kbd', rows.map(([a, b]) => h('tr', h('td', a), h('td', h('kbd', b))))), buttons: [{ label: 'Close', primary: true }] })
}

// ── charts ───────────────────────────────────────────────────────────────────────────────────────────────────────────────
const CHART_COLORS = ['#2563eb', '#d97706', '#059669', '#db2777', '#7c3aed', '#0891b2', '#dc2626', '#65a30d']
function chartData(ch) {
  const book = A.book
  let ref
  try {
    ref = XL.tokenize(ch.range)[0].ref
  } catch {
    return null
  }
  const si = ref.sheet != null ? book.sheetIndex(ref.sheet) : A.si
  if (si < 0) return null
  const m = book.matrix(new XL.Rng(si, ref.r1, ref.c1, ref.r2, ref.c2))
  if (m.length < 1 || !m[0].length) return null
  const header = m.length > 1 && m[0].slice(1).some((v) => typeof v === 'string') || (m[0].length === 1 && typeof m[0][0] === 'string' && m.length > 1)
  const rows = header ? m.slice(1) : m
  const single = m[0].length === 1
  const cats = single ? rows.map((_, i) => String(i + 1)) : rows.map((r) => (r[0] instanceof XL.XlErr ? r[0].code : r[0] == null ? '' : typeof r[0] === 'number' ? fmtNum(r[0]) : String(r[0])))
  const xnum = rows.map((r) => (typeof r[0] === 'number' ? r[0] : NaN))
  const from = single ? 0 : 1
  const series = []
  for (let c = from; c < m[0].length; c++) series.push({ name: header && m[0][c] != null ? String(m[0][c]) : `Series ${series.length + 1}`, values: rows.map((r) => (typeof r[c] === 'number' ? r[c] : r[c] === null || r[c] === '' ? null : Number(r[c]) || 0)) })
  return { cats, xnum, series, header }
}
function niceScale(min, max, n = 5) {
  if (min === max) {
    max = min + 1
    min = min === 0 ? 0 : min - 1
  }
  const span = max - min
  const raw = span / n
  const mag = 10 ** Math.floor(Math.log10(raw))
  const step = [1, 2, 2.5, 5, 10].map((x) => x * mag).find((x) => x >= raw) || raw
  const lo = Math.floor(min / step) * step
  const hi = Math.ceil(max / step) * step
  const ticks = []
  for (let v = lo; v <= hi + step / 2; v += step) ticks.push(+v.toFixed(10))
  return { lo, hi, ticks }
}
const svgEl = (name, attrs = {}, ...kids) => {
  const e = document.createElementNS('http://www.w3.org/2000/svg', name)
  for (const [k, v] of Object.entries(attrs)) e.setAttribute(k, v)
  for (const k of kids.flat()) if (k !== null && k !== undefined) e.append(k instanceof Node ? k : document.createTextNode(String(k)))
  return e
}
function chartSvg(ch, W, H) {
  const d = chartData(ch)
  const svg = svgEl('svg', { viewBox: `0 0 ${W} ${H}`, width: '100%', height: '100%', class: 'chart-svg' })
  if (!d || !d.series.length) return (svg.append(svgEl('text', { x: W / 2, y: H / 2, 'text-anchor': 'middle', class: 'ch-muted' }, 'No data in the range')), svg)
  const type = ch.type || 'column'
  const title = ch.title ? 26 : 8
  const legendH = d.series.length > 1 || type === 'pie' ? 24 : 0
  const left = 46
  const bottom = 26
  const px = { l: type === 'pie' ? 10 : left, r: 14, t: title + 4, b: bottom + legendH }
  const iw = W - px.l - px.r
  const ih = H - px.t - px.b
  if (ch.title) svg.append(svgEl('text', { x: 12, y: 19, class: 'ch-title' }, ch.title))
  const color = (i) => CHART_COLORS[i % CHART_COLORS.length]
  if (type === 'pie') {
    const s = d.series[0]
    const vals = s.values.map((v) => Math.max(0, v || 0))
    const total = vals.reduce((a, b) => a + b, 0) || 1
    const cx = W / 2
    const cy = px.t + ih / 2
    const rad = Math.min(iw, ih) / 2 - 6
    let a0 = -Math.PI / 2
    vals.forEach((v, i) => {
      const a1 = a0 + (v / total) * Math.PI * 2
      const large = a1 - a0 > Math.PI ? 1 : 0
      const p = svgEl('path', { d: vals.length === 1 ? `M ${cx - rad} ${cy} a ${rad} ${rad} 0 1 0 ${rad * 2} 0 a ${rad} ${rad} 0 1 0 ${-rad * 2} 0` : `M ${cx} ${cy} L ${cx + rad * Math.cos(a0)} ${cy + rad * Math.sin(a0)} A ${rad} ${rad} 0 ${large} 1 ${cx + rad * Math.cos(a1)} ${cy + rad * Math.sin(a1)} Z`, fill: color(i), stroke: 'var(--c-surface)', 'stroke-width': 1.5 }, svgEl('title', {}, `${d.cats[i]}: ${fmtNum(v)} (${Math.round((v / total) * 100)}%)`))
      svg.append(p)
      if (v / total > 0.06) svg.append(svgEl('text', { x: cx + rad * 0.62 * Math.cos((a0 + a1) / 2), y: cy + rad * 0.62 * Math.sin((a0 + a1) / 2) + 4, 'text-anchor': 'middle', class: 'ch-pie-label' }, `${Math.round((v / total) * 100)}%`))
      a0 = a1
    })
    vals.forEach((_, i) => {
      const lx = 12 + i * 90
      if (lx + 80 < W) {
        svg.append(svgEl('rect', { x: lx, y: H - 17, width: 9, height: 9, fill: color(i), rx: 2 }))
        svg.append(svgEl('text', { x: lx + 13, y: H - 9, class: 'ch-legend' }, String(d.cats[i]).slice(0, 11)))
      }
    })
    return svg
  }
  const all = d.series.flatMap((s) => s.values).filter((v) => v !== null)
  const stacked = ch.stacked && (type === 'column' || type === 'bar' || type === 'area')
  let lo = Math.min(0, ...all)
  let hi = Math.max(...all, 0)
  if (stacked) {
    const sums = d.cats.map((_, i) => d.series.reduce((a, s) => a + (s.values[i] || 0), 0))
    hi = Math.max(...sums, 0)
    lo = Math.min(0, ...d.series.flatMap((s) => s.values).filter((v) => v !== null))
  }
  const isScatter = type === 'scatter'
  const xs = isScatter ? d.xnum.filter((v) => !Number.isNaN(v)) : []
  const sy = niceScale(type === 'line' || isScatter ? Math.min(...all) : lo, hi)
  const sx = isScatter && xs.length ? niceScale(Math.min(...xs), Math.max(...xs)) : null
  const horizontal = type === 'bar'
  const Y = (v) => px.t + ih - ((v - sy.lo) / (sy.hi - sy.lo)) * ih
  const Xh = (v) => px.l + ((v - sy.lo) / (sy.hi - sy.lo)) * iw
  // grid and axes
  for (const t of sy.ticks) {
    if (horizontal) {
      svg.append(svgEl('line', { x1: Xh(t), x2: Xh(t), y1: px.t, y2: px.t + ih, class: 'ch-grid' }), svgEl('text', { x: Xh(t), y: px.t + ih + 14, 'text-anchor': 'middle', class: 'ch-axis' }, axisNum(t)))
    } else {
      svg.append(svgEl('line', { x1: px.l, x2: px.l + iw, y1: Y(t), y2: Y(t), class: 'ch-grid' }), svgEl('text', { x: px.l - 6, y: Y(t) + 4, 'text-anchor': 'end', class: 'ch-axis' }, axisNum(t)))
    }
  }
  const n = d.cats.length
  if (isScatter && sx) {
    const X = (v) => px.l + ((v - sx.lo) / (sx.hi - sx.lo)) * iw
    for (const t of sx.ticks) svg.append(svgEl('text', { x: X(t), y: px.t + ih + 14, 'text-anchor': 'middle', class: 'ch-axis' }, axisNum(t)))
    d.series.forEach((s, si) => s.values.forEach((v, i) => v !== null && !Number.isNaN(d.xnum[i]) && svg.append(svgEl('circle', { cx: X(d.xnum[i]), cy: Y(v), r: 4, fill: color(si), 'fill-opacity': 0.8 }, svgEl('title', {}, `${s.name}: ${fmtNum(d.xnum[i])}, ${fmtNum(v)}`)))))
  } else if (horizontal) {
    const band = ih / n
    const bw = Math.min(26, (band * 0.7) / (stacked ? 1 : d.series.length))
    d.cats.forEach((c, i) => {
      svg.append(svgEl('text', { x: px.l - 6, y: px.t + band * i + band / 2 + 4, 'text-anchor': 'end', class: 'ch-axis' }, String(c).slice(0, 8)))
      let acc = 0
      d.series.forEach((s, si) => {
        const v = s.values[i] || 0
        const y0 = stacked ? px.t + band * i + (band - bw) / 2 : px.t + band * i + (band - bw * d.series.length) / 2 + si * bw
        const x0 = Xh(stacked ? acc : Math.min(0, v))
        const x1 = Xh(stacked ? acc + v : Math.max(0, v))
        svg.append(svgEl('rect', { x: x0, y: y0, width: Math.max(1, x1 - x0), height: bw - 1, fill: color(si), rx: 1.5 }, svgEl('title', {}, `${s.name} · ${c}: ${fmtNum(v)}`)))
        acc += v
      })
    })
  } else {
    const band = iw / n
    d.cats.forEach((c, i) => {
      if (n <= 14 || i % Math.ceil(n / 14) === 0) svg.append(svgEl('text', { x: px.l + band * i + band / 2, y: px.t + ih + 14, 'text-anchor': 'middle', class: 'ch-axis' }, String(c).slice(0, 9)))
    })
    if (type === 'column') {
      const bw = Math.min(34, (band * 0.72) / (stacked ? 1 : d.series.length))
      d.cats.forEach((c, i) => {
        let acc = 0
        d.series.forEach((s, si) => {
          const v = s.values[i]
          if (v === null) return
          const x0 = stacked ? px.l + band * i + (band - bw) / 2 : px.l + band * i + (band - bw * d.series.length) / 2 + si * bw
          const y0 = Y(stacked ? acc + v : Math.max(0, v))
          const y1 = Y(stacked ? acc : Math.min(0, v))
          svg.append(svgEl('rect', { x: x0, y: y0, width: bw - 1, height: Math.max(1, y1 - y0), fill: color(si), rx: 1.5 }, svgEl('title', {}, `${s.name} · ${c}: ${fmtNum(v)}`)))
          acc += v
        })
      })
    } else {
      d.series.forEach((s, si) => {
        const pts = s.values.map((v, i) => (v === null ? null : [px.l + band * i + band / 2, Y(v)]))
        const path = pts.filter(Boolean).map((p, i) => `${i ? 'L' : 'M'} ${p[0]} ${p[1]}`).join(' ')
        if (type === 'area') svg.append(svgEl('path', { d: `${path} L ${pts.filter(Boolean).slice(-1)[0][0]} ${Y(0)} L ${pts.filter(Boolean)[0][0]} ${Y(0)} Z`, fill: color(si), 'fill-opacity': 0.22 }))
        svg.append(svgEl('path', { d: path, fill: 'none', stroke: color(si), 'stroke-width': 2, 'stroke-linejoin': 'round' }))
        pts.forEach((p, i) => p && svg.append(svgEl('circle', { cx: p[0], cy: p[1], r: 3, fill: color(si) }, svgEl('title', {}, `${s.name} · ${d.cats[i]}: ${fmtNum(s.values[i])}`))))
      })
    }
  }
  if (legendH) {
    d.series.forEach((s, i) => {
      const lx = 12 + i * 96
      if (lx + 90 > W) return
      svg.append(svgEl('rect', { x: lx, y: H - 17, width: 9, height: 9, fill: color(i), rx: 2 }), svgEl('text', { x: lx + 13, y: H - 9, class: 'ch-legend' }, s.name.slice(0, 12)))
    })
  }
  return svg
}
let chartLayer
function buildChartLayer() {
  chartLayer = h('.chart-layer')
  return chartLayer
}
let lastChartSig = ''
function positionCharts() {
  if (!chartLayer) return
  const sh = sheet()
  const sig = JSON.stringify(sh.charts) + A.book.ver + A.si + A.zoom
  const have = new Map($$('.chart-box', chartLayer).map((el) => [el.dataset.id, el]))
  const ids = new Set(sh.charts.map((c) => c.id))
  for (const [id, el] of have) if (!ids.has(id)) el.remove()
  for (const ch of sh.charts) {
    const p = XL.parseAddr(ch.anchor) || { r: 0, c: 0 }
    const b = rect(p.r, p.c, false)
    let el = have.get(ch.id)
    const w = ch.w * A.zoom
    const hh = ch.h * A.zoom
    if (!el || el.dataset.sig !== sig) {
      if (!el) {
        el = h('.chart-box', { dataset: { id: ch.id } })
        chartLayer.append(el)
        wireChart(el, ch)
      }
      el.dataset.sig = sig
      el.querySelector('svg')?.remove()
      const svg = chartSvg(ch, ch.w, ch.h)
      el.prepend(svg)
    }
    el.style.left = `${b.x}px`
    el.style.top = `${b.y}px`
    el.style.width = `${w}px`
    el.style.height = `${hh}px`
    el.classList.toggle('sel', A.charts.sel === ch.id)
    el.style.display = b.x + w < HDR_W || b.y + hh < HDR_H || b.x > viewW || b.y > viewH ? 'none' : 'block'
  }
  lastChartSig = sig
}
function wireChart(el, ch) {
  el.append(h('button.chart-x', { title: 'Delete chart', onclick: (e) => (e.stopPropagation(), updateChart(ch.id, null)) }, icon('x', 12)), h('button.chart-edit', { title: 'Edit chart', onclick: (e) => (e.stopPropagation(), App.openChart(ch.id)) }, icon('edit', 12)), h('.chart-resize'))
  el.addEventListener('mousedown', (e) => {
    if (e.target.closest('button')) return
    A.charts.sel = ch.id
    const cur = () => sheet().charts.find((c) => c.id === ch.id)
    const start = { x: e.clientX, y: e.clientY, w: cur().w, h: cur().h, left: el.offsetLeft, top: el.offsetTop }
    const resizing = e.target.classList.contains('chart-resize')
    e.preventDefault()
    const move = (ev) => {
      const dx = ev.clientX - start.x
      const dy = ev.clientY - start.y
      if (resizing) {
        el.style.width = `${Math.max(180, start.w * A.zoom + dx)}px`
        el.style.height = `${Math.max(120, start.h * A.zoom + dy)}px`
      } else {
        el.style.left = `${start.left + dx}px`
        el.style.top = `${start.top + dy}px`
      }
    }
    const up = (ev) => {
      removeEventListener('mousemove', move)
      removeEventListener('mouseup', up)
      const c = cur()
      if (!c) return
      if (resizing) updateChart(c.id, { ...c, w: Math.round(Math.max(180, start.w + (ev.clientX - start.x) / A.zoom)), h: Math.round(Math.max(120, start.h + (ev.clientY - start.y) / A.zoom)) })
      else if (Math.abs(ev.clientX - start.x) + Math.abs(ev.clientY - start.y) > 3) {
        const r = scroller.getBoundingClientRect()
        const cell = hitCell(clamp(el.offsetLeft + 1, HDR_W + 1, viewW - 2) , clamp(el.offsetTop + 1, HDR_H + 1, viewH - 2))
        void r
        updateChart(c.id, { ...c, anchor: A1(cell.r, cell.c) })
      }
      schedule()
    }
    addEventListener('mousemove', move)
    addEventListener('mouseup', up)
    schedule()
  })
  el.addEventListener('dblclick', (e) => !e.target.closest('button') && App.openChart(ch.id))
}
function updateChart(id, next) {
  const sh = sheet()
  const list = next ? sh.charts.map((c) => (c.id === id ? next : c)) : sh.charts.filter((c) => c.id !== id)
  applyOps([{ op: 'prop', sheet: sh.id, key: 'charts', value: list }], { label: 'Chart' })
}
App.openChart = (id) => {
  const sh = sheet()
  const existing = id ? sh.charts.find((c) => c.id === id) : null
  const g = norm()
  const used = A.book.usedRange(sh.id)
  const guess = g.r1 === g.r2 && g.c1 === g.c2 ? (selectionRegion() ?? g) : g
  const sheetRef = (/^[A-Za-z_]\w*$/.test(sh.name) ? sh.name : `'${sh.name}'`)
  const type = sel([['column', 'Column'], ['bar', 'Bar'], ['line', 'Line'], ['area', 'Area'], ['pie', 'Pie'], ['scatter', 'Scatter']], existing?.type || 'column')
  const title = h('input.input', { value: existing?.title || '', placeholder: 'Chart title' })
  const range = h('input.input.mono', { value: existing?.range || `${sheetRef}!${XL.rangeText(guess)}` })
  const stacked = h('input', { type: 'checkbox', checked: !!existing?.stacked })
  const prev = h('.chart-preview')
  const paint = () => {
    prev.replaceChildren()
    try {
      prev.append(chartSvg({ type: type.value, title: title.value, range: range.value, stacked: stacked.checked }, 520, 300))
    } catch {
      prev.append(h('span.hint', 'Check the range'))
    }
  }
  for (const el of [type, title, range, stacked]) el.addEventListener('input', paint)
  paint()
  void used
  dialog({
    title: existing ? 'Edit chart' : 'Insert chart',
    width: 580,
    body: h('.col', h('.fc-row', field('Type', type), field('Title', title)), field('Data range', range, 'First row = series names, first column = labels'), h('label.check', stacked, 'Stacked'), prev),
    buttons: [
      { label: 'Cancel' },
      existing ? { label: 'Delete', danger: true, onclick: () => updateChart(existing.id, null) } : null,
      {
        label: existing ? 'Save' : 'Insert',
        primary: true,
        onclick: () => {
          const ch = { id: existing?.id || XL.newId(), type: type.value, title: title.value, range: range.value, stacked: stacked.checked, anchor: existing?.anchor || A1(Math.max(0, norm().r1), (XL.parseRange(range.value.replace(/^.*!/, ''))?.c2 ?? 3) + 2), w: existing?.w || 480, h: existing?.h || 300 }
          const list = existing ? sh.charts.map((c) => (c.id === ch.id ? ch : c)) : [...sh.charts, ch]
          applyOps([{ op: 'prop', sheet: sh.id, key: 'charts', value: list }], { label: 'Chart' })
          A.charts.sel = ch.id
        },
      },
    ].filter(Boolean),
  })
}

// ── data from a database, into the sheet ─────────────────────────────────────────────────────────────────────────────────
App.openDbImport = () => {
  if (!A.conns.length) {
    toast('Add a database connection first.', 'warn')
    return App.setMode('db')
  }
  const conn = sel(A.conns.map((c) => [c.id, `${c.name} (${c.type})`]), A.dbConn || A.conns[0].id)
  const sql = h('textarea.input.mono', { rows: 7, style: { width: '100%' }, placeholder: 'select * from customers limit 500' })
  const at = h('input.input', { value: A1(A.sel.ar, A.sel.ac), style: { width: '100px' } })
  const info = h('.hint', '')
  const g = norm()
  const sendSel = h('button.btn', { onclick: () => App.openSheetToTable(conn.value) }, 'Send selection to a table…')
  void g
  dialog({
    title: 'Data from a database',
    width: 640,
    body: h('.col', field('Connection', conn), field('Query (read-only)', sql), field('Place at cell', at), info, h('div', sendSel)),
    buttons: [
      { label: 'Cancel' },
      {
        label: 'Run and place on sheet',
        primary: true,
        onclick: async () => {
          if (!sql.value.trim()) return false
          info.textContent = 'Running…'
          try {
            if (A.pending.length) await Save.flush()
            const r = await Host.tool('app_db_to_sheet', { connection: conn.value, sql: sql.value, id: A.book.id, sheet: sheet().id, at: at.value })
            loadWorkbook(r.workbook, { keepView: true })
            toast(`${r.rows} rows placed`)
          } catch (e) {
            info.textContent = e.message
            return false
          }
        },
      },
    ],
  })
}
App.openSheetToTable = (connId) => {
  const g = selectionRegion()
  const conn = sel(A.conns.map((c) => [c.id, `${c.name} (${c.type})`]), connId || A.dbConn || A.conns[0]?.id)
  const table = h('input.input', { placeholder: 'table_name', style: { width: '100%' } })
  const range = h('input.input.mono', { value: XL.rangeText(g), style: { width: '100%' } })
  const replace = h('input', { type: 'checkbox' })
  dialog({
    title: 'Send range to a table',
    width: 460,
    body: h('.col', field('Connection', conn), field('Table name', table), field('Range (first row = column names)', range), h('label.check', replace, 'Replace the table if it exists')),
    buttons: [
      { label: 'Cancel' },
      {
        label: 'Create table',
        primary: true,
        onclick: async () => {
          try {
            if (A.pending.length) await Save.flush()
            const r = await Host.tool('app_sheet_to_table', { connection: conn.value, table: table.value.trim(), id: A.book.id, sheet: sheet().id, range: range.value, replace: replace.checked })
            toast(`${r.inserted} rows written to ${table.value}`)
          } catch (e) {
            toast(e.message, 'err')
            return false
          }
        },
      },
    ],
  })
}

// build the whole spreadsheet view
function buildSheetView() {
  const host = h('.grid-host')
  const view = h('.sheet-view', buildToolbar(), buildFormulaBar(), host, buildTabs())
  host.append(buildChartLayer())
  return { view, host }
}
