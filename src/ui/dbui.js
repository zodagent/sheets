// ── database workbench: connections, schema tree, SQL editor, result and table grids ──────────────────────────────────────────────────────
const D = { conn: null, schemas: {}, loading: {}, tabs: [], active: -1, history: [], saved: [], treeFilter: '', open: {}, seq: 0 }
const SQL_KEYWORDS = 'select from where and or not in is null like ilike between exists case when then else end as on join left right inner outer full cross using group by order having limit offset union all intersect except distinct insert into values update set delete create table view index alter add drop column primary key foreign references default unique check constraint asc desc with recursive returning explain analyze begin commit rollback true false cast count sum avg min max coalesce nullif over partition row_number rank lower upper length substr substring trim replace round abs date datetime now current_date current_timestamp interval if truncate pragma show describe'.split(' ')
const KW = new Set(SQL_KEYWORDS)
const DB_COLORS = ['#2563eb', '#059669', '#d97706', '#dc2626', '#7c3aed', '#0891b2', '#64748b']
const TYPE_LABEL = { sqlite: 'SQLite', postgres: 'PostgreSQL', mysql: 'MySQL' }
const connOf = (id) => A.conns.find((c) => c.id === id)
const curTab = () => D.tabs[D.active]
const ROW_H = 26

// ── result grid ──────────────────────────────────────────────────────────────────────────────────────────────────────────
function cellText(v) {
  if (v === null || v === undefined) return ''
  if (typeof v === 'object') return JSON.stringify(v)
  return String(v)
}
// opts: { columns, rows, types?, editable?, pk?, onSort?, sort?, onChanges? }
function resultGrid(opts) {
  const st = { sel: null, sortCol: opts.sort?.col ?? null, desc: !!opts.sort?.desc, edit: null, order: null }
  const cols = opts.columns
  const widths = cols.map((n, i) => {
    const sample = opts.rows.slice(0, 60).map((r) => cellText(r[i]).length)
    const head = n.length * 8 + (opts.types?.[i] ? opts.types[i].length * 6.4 + 10 : 0) + (opts.pkCols?.includes(n) ? 16 : 0) + 34
    return clamp(Math.max(head, Math.max(0, ...sample) * 7.2 + 18), 70, 340)
  })
  const root = h('.rg', { tabIndex: 0 })
  const head = h('.rg-head')
  const body = h('.rg-body')
  const spacer = h('.rg-spacer')
  const rowsEl = h('.rg-rows')
  body.append(spacer, rowsEl)
  root.append(head, body)
  const changes = (opts.changes ||= { upd: new Map(), del: new Set(), ins: [] })
  const total = () => view().length + changes.ins.length
  const view = () => {
    if (st.order) return st.order
    return opts.rows.map((_, i) => i)
  }
  const sortLocal = () => {
    if (opts.onSort || st.sortCol === null) return void (st.order = null)
    const c = st.sortCol
    st.order = opts.rows.map((_, i) => i).sort((a, b) => {
      const x = opts.rows[a][c]
      const y = opts.rows[b][c]
      if (x === y) return 0
      if (x === null || x === undefined) return 1
      if (y === null || y === undefined) return -1
      const d = typeof x === 'number' && typeof y === 'number' ? x - y : String(x).localeCompare(String(y), undefined, { numeric: true })
      return st.desc ? -d : d
    })
  }
  const totalW = () => 52 + widths.reduce((a, b) => a + b, 0)
  const paintHead = () => {
    head.replaceChildren(
      h('.rg-corner', { style: { width: '52px' } }),
      ...cols.map((n, i) =>
        h(
          '.rg-th',
          { style: { width: `${widths[i]}px` }, onclick: () => sortBy(i), title: `${n}${opts.types?.[i] ? ` · ${opts.types[i]}` : ''}` },
          h('span.rg-name', n),
          opts.pkCols?.includes(n) ? icon('key', 11) : '',
          opts.types?.[i] ? h('span.rg-type', opts.types[i]) : '',
          st.sortCol === i ? h('span.rg-sort', st.desc ? '↓' : '↑') : '',
          h('.rg-grip', {
            onmousedown: (e) => {
              e.stopPropagation()
              e.preventDefault()
              const x0 = e.clientX
              const w0 = widths[i]
              const mv = (ev) => ((widths[i] = clamp(w0 + ev.clientX - x0, 40, 700)), paintHead(), paint())
              const up = () => (removeEventListener('mousemove', mv), removeEventListener('mouseup', up))
              addEventListener('mousemove', mv)
              addEventListener('mouseup', up)
            },
            onclick: (e) => e.stopPropagation(),
          }),
        ),
      ),
    )
    head.style.width = `${totalW()}px`
  }
  function sortBy(i) {
    if (st.sortCol === i) {
      if (!st.desc) st.desc = true
      else ((st.sortCol = null), (st.desc = false))
    } else ((st.sortCol = i), (st.desc = false))
    if (opts.onSort) opts.onSort(st.sortCol === null ? null : { col: cols[st.sortCol], desc: st.desc })
    else {
      sortLocal()
      paintHead()
      paint()
    }
    paintHead()
  }
  const origIndex = (vi) => view()[vi]
  const rowKey = (ri) => (opts.pk || []).reduce((o, k) => ((o[k] = opts.rows[ri][cols.indexOf(k)]), o), {})
  function paint() {
    const n = total()
    spacer.style.height = `${n * ROW_H}px`
    spacer.style.width = `${totalW()}px`
    const top = body.scrollTop
    const first = Math.max(0, Math.floor(top / ROW_H) - 3)
    const last = Math.min(n, Math.ceil((top + body.clientHeight) / ROW_H) + 3)
    rowsEl.style.width = `${totalW()}px`
    const out = []
    for (let vi = first; vi < last; vi++) {
      const isIns = vi >= view().length
      const ri = isIns ? -1 : origIndex(vi)
      const ins = isIns ? changes.ins[vi - view().length] : null
      const deleted = !isIns && changes.del.has(ri)
      const upd = isIns ? null : changes.upd.get(ri)
      const tr = h('.rg-row', { class: `${vi % 2 ? 'odd' : ''} ${deleted ? 'del' : ''} ${isIns ? 'new' : ''}`, style: { top: `${vi * ROW_H}px` } }, h('.rg-num', { onclick: (e) => selectRow(vi, e) }, vi + 1 + (opts.offset || 0)))
      cols.forEach((name, ci) => {
        const raw = isIns ? ins[name] : upd && name in upd ? upd[name] : opts.rows[ri][ci]
        const isNull = raw === null || raw === undefined
        const selected = st.sel && vi >= Math.min(st.sel.r1, st.sel.r2) && vi <= Math.max(st.sel.r1, st.sel.r2) && ci >= Math.min(st.sel.c1, st.sel.c2) && ci <= Math.max(st.sel.c1, st.sel.c2)
        const td = h(
          '.rg-td',
          {
            class: `${typeof raw === 'number' ? 'num' : ''} ${isNull ? 'null' : ''} ${selected ? 'sel' : ''} ${upd && name in upd ? 'dirty' : ''}`,
            style: { width: `${widths[ci]}px` },
            title: cellText(raw),
            onmousedown: (e) => selectCell(vi, ci, e),
            ondblclick: () => beginEdit(vi, ci, td),
            oncontextmenu: (e) => (e.preventDefault(), selectCell(vi, ci, e, true), cellMenu(e, vi, ci)),
          },
          isNull ? 'NULL' : cellText(raw),
        )
        tr.append(td)
      })
      out.push(tr)
    }
    rowsEl.replaceChildren(...out)
  }
  function selectCell(vi, ci, e, keep) {
    if (e.shiftKey && st.sel) st.sel = { ...st.sel, r2: vi, c2: ci }
    else if (!(keep && st.sel && vi >= Math.min(st.sel.r1, st.sel.r2) && vi <= Math.max(st.sel.r1, st.sel.r2) && ci >= Math.min(st.sel.c1, st.sel.c2) && ci <= Math.max(st.sel.c1, st.sel.c2))) st.sel = { r1: vi, c1: ci, r2: vi, c2: ci }
    root.focus()
    paint()
  }
  function selectRow(vi, e) {
    st.sel = e.shiftKey && st.sel ? { ...st.sel, r2: vi, c2: cols.length - 1 } : { r1: vi, c1: 0, r2: vi, c2: cols.length - 1 }
    root.focus()
    paint()
  }
  const cellValue = (vi, ci) => {
    const isIns = vi >= view().length
    if (isIns) return changes.ins[vi - view().length][cols[ci]]
    const ri = origIndex(vi)
    const u = changes.upd.get(ri)
    return u && cols[ci] in u ? u[cols[ci]] : opts.rows[ri][ci]
  }
  const selection = () => {
    if (!st.sel) return null
    const { r1, r2, c1, c2 } = st.sel
    return { r1: Math.min(r1, r2), r2: Math.max(r1, r2), c1: Math.min(c1, c2), c2: Math.max(c1, c2) }
  }
  function selectedRows() {
    const s = selection()
    if (!s) return []
    const out = []
    for (let r = s.r1; r <= s.r2; r++) out.push(cols.map((_, c) => cellValue(r, c)))
    return out
  }
  const copySelection = (kind = 'tsv') => {
    const s = selection()
    if (!s) return
    const rows = []
    for (let r = s.r1; r <= s.r2; r++) rows.push(Array.from({ length: s.c2 - s.c1 + 1 }, (_, i) => cellValue(r, s.c1 + i)))
    const names = cols.slice(s.c1, s.c2 + 1)
    let text
    if (kind === 'json') text = JSON.stringify(rows.map((r) => Object.fromEntries(names.map((n, i) => [n, r[i]]))), null, 2)
    else if (kind === 'csv') text = [names, ...rows.map((r) => r.map(cellText))].map((r) => r.map((v) => (/[",\n]/.test(v) ? `"${v.replace(/"/g, '""')}"` : v)).join(',')).join('\n')
    else if (kind === 'md') text = `| ${names.join(' | ')} |\n| ${names.map(() => '---').join(' | ')} |\n${rows.map((r) => `| ${r.map((v) => cellText(v).replace(/\|/g, '\\|')).join(' | ')} |`).join('\n')}`
    else if (kind === 'insert') text = rows.map((r) => `INSERT INTO ${opts.table || 'table_name'} (${names.join(', ')}) VALUES (${r.map((v) => (v === null || v === undefined ? 'NULL' : typeof v === 'number' ? v : `'${String(v).replace(/'/g, "''")}'`)).join(', ')});`).join('\n')
    else text = rows.map((r) => r.map(cellText).join('\t')).join('\n')
    copyText(text)
    toast(`Copied ${rows.length} row${rows.length === 1 ? '' : 's'}`)
  }
  function cellMenu(e, vi, ci) {
    const s = selection()
    const items = [
      { label: 'Copy', icon: 'copy', key: 'Ctrl+C', onclick: () => copySelection('tsv') },
      { label: 'Copy as CSV', onclick: () => copySelection('csv') },
      { label: 'Copy as JSON', onclick: () => copySelection('json') },
      { label: 'Copy as Markdown', onclick: () => copySelection('md') },
      { label: 'Copy as INSERT', onclick: () => copySelection('insert') },
    ]
    if (opts.filterBy) items.push({ sep: true }, { label: 'Filter by this value', onclick: () => opts.filterBy(cols[ci], cellValue(vi, ci)) })
    if (opts.editable) {
      items.push({ sep: true }, { label: 'Edit cell', onclick: () => beginEdit(vi, ci) }, { label: 'Set to NULL', onclick: () => setCell(vi, ci, null) }, { label: `Delete ${s && s.r2 > s.r1 ? `${s.r2 - s.r1 + 1} rows` : 'row'}`, icon: 'trash', onclick: () => deleteRows() })
    }
    menu(items, e.clientX, e.clientY)
  }
  function setCell(vi, ci, value) {
    const name = cols[ci]
    if (vi >= view().length) changes.ins[vi - view().length][name] = value
    else {
      const ri = origIndex(vi)
      const orig = opts.rows[ri][ci]
      const u = changes.upd.get(ri) || {}
      if (value === orig || (value === '' && orig === null && false)) delete u[name]
      else u[name] = value
      if (Object.keys(u).length) changes.upd.set(ri, u)
      else changes.upd.delete(ri)
    }
    opts.onChanges?.()
    paint()
  }
  function beginEdit(vi, ci, td) {
    if (!opts.editable) return
    const cur = cellValue(vi, ci)
    const input = h('input.rg-input', { value: cur === null || cur === undefined ? '' : cellText(cur) })
    const cell = td || rowsEl.children[vi - Math.max(0, Math.floor(body.scrollTop / ROW_H) - 3)]?.children[ci + 1]
    if (!cell) return
    cell.replaceChildren(input)
    input.focus()
    input.select()
    let done = false
    const finish = (ok) => {
      if (done) return
      done = true
      if (ok) {
        const t = input.value
        const orig = cur
        let v = t
        if (typeof orig === 'number' && t.trim() !== '' && !Number.isNaN(Number(t))) v = Number(t)
        else if (t === '' && (orig === null || orig === undefined)) v = null
        setCell(vi, ci, v)
      } else paint()
    }
    input.addEventListener('keydown', (e) => {
      if (e.key === 'Enter') finish(true)
      if (e.key === 'Escape') finish(false)
      if (e.key === 'Tab') {
        e.preventDefault()
        finish(true)
        st.sel = { r1: vi, c1: Math.min(ci + 1, cols.length - 1), r2: vi, c2: Math.min(ci + 1, cols.length - 1) }
        paint()
      }
    })
    input.addEventListener('blur', () => finish(true))
  }
  function deleteRows() {
    const s = selection()
    if (!s) return
    for (let vi = s.r1; vi <= s.r2; vi++) {
      if (vi >= view().length) changes.ins.splice(vi - view().length, 1)
      else changes.del.add(origIndex(vi))
    }
    opts.onChanges?.()
    paint()
  }
  root.addEventListener('keydown', (e) => {
    if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'c') {
      e.preventDefault()
      copySelection('tsv')
    }
    const s = selection()
    if (!s) return
    const mv = (dr, dc) => {
      e.preventDefault()
      const r = clamp((st.sel.r2 ?? s.r1) + dr, 0, total() - 1)
      const c = clamp((st.sel.c2 ?? s.c1) + dc, 0, cols.length - 1)
      st.sel = e.shiftKey ? { ...st.sel, r2: r, c2: c } : { r1: r, c1: c, r2: r, c2: c }
      if (r * ROW_H < body.scrollTop) body.scrollTop = r * ROW_H
      if ((r + 1) * ROW_H > body.scrollTop + body.clientHeight) body.scrollTop = (r + 1) * ROW_H - body.clientHeight
      paint()
    }
    if (e.key === 'ArrowDown') mv(1, 0)
    else if (e.key === 'ArrowUp') mv(-1, 0)
    else if (e.key === 'ArrowRight') mv(0, 1)
    else if (e.key === 'ArrowLeft') mv(0, -1)
    else if (e.key === 'Enter' || e.key === 'F2') beginEdit(s.r1, s.c1)
    else if ((e.key === 'Delete') && opts.editable) deleteRows()
  })
  body.addEventListener('scroll', () => {
    head.style.transform = `translateX(${-body.scrollLeft}px)`
    paint()
  })
  new ResizeObserver(() => paint()).observe(body)
  paintHead()
  sortLocal()
  paint()
  return {
    el: root,
    changes,
    paint,
    addRow: () => {
      changes.ins.push({})
      opts.onChanges?.()
      paint()
      body.scrollTop = total() * ROW_H
    },
    count: () => changes.upd.size + changes.del.size + changes.ins.length,
    discard: () => {
      changes.upd.clear()
      changes.del.clear()
      changes.ins.length = 0
      opts.onChanges?.()
      paint()
    },
    collect: () => {
      const out = []
      for (const [ri, set] of changes.upd) if (!changes.del.has(ri)) out.push({ type: 'update', pk: rowKey(ri), set })
      for (const ri of changes.del) out.push({ type: 'delete', pk: rowKey(ri) })
      for (const v of changes.ins) out.push({ type: 'insert', values: v })
      return out
    },
    copySelection,
    selectedRows,
    selection,
  }
}

// ── schema tree ──────────────────────────────────────────────────────────────────────────────────────────────────────────
async function loadSchema(id, force = false) {
  if (D.loading[id]) return
  if (D.schemas[id] && !force) return
  D.loading[id] = true
  renderSide()
  try {
    D.schemas[id] = await Host.tool('app_db_schema', { connection: id })
    D.schemas[id].error = null
  } catch (e) {
    D.schemas[id] = { error: e.message, schemas: [] }
  } finally {
    D.loading[id] = false
    renderSide()
  }
}
function selectConn(id) {
  D.conn = id
  A.dbConn = id
  loadSchema(id)
  renderSide()
  renderMain()
}
const tablesOf = (id) => (D.schemas[id]?.schemas || []).flatMap((s) => s.tables.map((t) => ({ ...t, schema: s.name })))
const qident = (type, n) => (type === 'mysql' ? `\`${n}\`` : `"${n}"`)
const qtable = (c, t) => (c.type === 'sqlite' || !t.schema ? qident(c.type, t.name) : `${qident(c.type, t.schema)}.${qident(c.type, t.name)}`)
function renderSide() {
  const side = $('.db-side')
  if (!side) return
  const list = h(
    '.conn-list',
    A.conns.map((c) =>
      h(
        '.conn',
        { class: c.id === D.conn ? 'on' : '', onclick: () => selectConn(c.id), ondblclick: () => loadSchema(c.id, true) },
        h('span.conn-dot', { style: { background: c.color || '#64748b' } }),
        h('span.conn-name', c.name),
        h('span.conn-type', TYPE_LABEL[c.type]),
        c.readOnly ? h('span.conn-lock', { title: 'Read-only' }, icon('lock', 12)) : '',
        h('button.icon-btn.conn-more', { title: 'Options', onclick: (e) => (e.stopPropagation(), connMenu(c, e)) }, icon('more', 14)),
      ),
    ),
  )
  const c = connOf(D.conn)
  const tree = h('.tree')
  if (c) {
    const sc = D.schemas[c.id]
    if (D.loading[c.id] && !sc) tree.append(h('.tree-note', 'Connecting…'))
    else if (sc?.error) tree.append(h('.tree-note.err', sc.error), h('button.btn', { onclick: () => loadSchema(c.id, true) }, 'Try again'))
    else if (sc) {
      const q = D.treeFilter.toLowerCase()
      for (const s of sc.schemas) {
        const tabs = s.tables.filter((t) => !q || t.name.toLowerCase().includes(q) || t.columns.some((x) => x.name.toLowerCase().includes(q)))
        if (!tabs.length && q) continue
        const multi = sc.schemas.length > 1 || sc.type !== 'sqlite'
        if (multi) tree.append(h('.tree-schema', icon('layers', 13), s.name, h('span.hint', tabs.length)))
        for (const t of tabs) {
          const k = `${c.id}.${s.name}.${t.name}`
          const open = D.open[k] || (q && t.columns.some((x) => x.name.toLowerCase().includes(q)))
          tree.append(
            h(
              '.tree-table',
              { class: open ? 'open' : '' },
              h(
                '.tree-row',
                { ondblclick: () => openTable(c.id, s.name, t.name), oncontextmenu: (e) => (e.preventDefault(), tableMenu(c, { ...t, schema: s.name }, e)) },
                h('button.twisty', { onclick: () => ((D.open[k] = !D.open[k]), renderSide()) }, icon(open ? 'down' : 'right2', 12)),
                icon(t.kind === 'view' ? 'eye' : 'table', 14),
                h('span.tree-name', { onclick: () => openTable(c.id, s.name, t.name) }, t.name),
                h('span.hint', t.columns.length),
              ),
              open ? h('.tree-cols', t.columns.map((x) => h('.tree-col', x.pk ? icon('key', 11) : h('span.col-pad'), h('span', x.name), h('span.hint', x.type)))) : '',
            ),
          )
        }
      }
      if (!tree.children.length) tree.append(h('.tree-note', q ? 'No match' : 'This database has no tables yet.'))
    }
  }
  side.replaceChildren(
    h('.side-head', h('span', 'Connections'), h('button.icon-btn', { title: 'New connection', onclick: () => connForm() }, icon('plus', 16))),
    A.conns.length ? list : h('.side-empty', h('p', 'No connections yet.'), h('button.btn.primary', { onclick: () => connForm() }, 'Add a connection')),
    c
      ? h('.side-head.sub', h('span', c.name), h('button.icon-btn', { title: 'Refresh schema', onclick: () => loadSchema(c.id, true) }, icon('refresh', 14)), h('button.icon-btn', { title: 'New query', onclick: () => newQuery() }, icon('play', 14)))
      : '',
    c ? h('input.input.tree-search', { placeholder: 'Filter tables and columns', value: D.treeFilter, oninput: (e) => ((D.treeFilter = e.target.value), renderSide(), $('.tree-search')?.focus(), $('.tree-search')?.setSelectionRange(99, 99)) }) : '',
    tree,
  )
}
function connMenu(c, e) {
  menu(
    [
      { label: 'Refresh schema', icon: 'refresh', onclick: () => (selectConn(c.id), loadSchema(c.id, true)) },
      { label: 'New query', icon: 'play', onclick: () => (selectConn(c.id), newQuery()) },
      { label: 'Edit connection…', icon: 'edit', onclick: () => connForm(c) },
      { sep: true },
      {
        label: 'Remove',
        icon: 'trash',
        onclick: async () => {
          if (!(await confirmDialog('Remove connection', `Remove "${c.name}"? The database itself is not touched.`, 'Remove'))) return
          const r = await Host.tool('app_db_delete', { connection: c.id })
          A.conns = r.connections
          D.tabs = D.tabs.filter((t) => t.conn !== c.id)
          if (D.conn === c.id) D.conn = null
          D.active = Math.min(D.active, D.tabs.length - 1)
          renderDb()
        },
      },
    ],
    e.clientX,
    e.clientY,
  )
}
function tableMenu(c, t, e) {
  const full = qtable(c, t)
  menu(
    [
      { label: 'Open data', icon: 'table', onclick: () => openTable(c.id, t.schema, t.name) },
      { label: 'Structure', icon: 'columns', onclick: () => openTable(c.id, t.schema, t.name, 'structure') },
      { label: 'SELECT in a new query', icon: 'play', onclick: () => newQuery(`select *\nfrom ${full}\nlimit 100;`, true) },
      { label: 'Count rows', onclick: () => newQuery(`select count(*) as rows\nfrom ${full};`, true) },
      { sep: true },
      { label: 'Send to a sheet', icon: 'sheet', onclick: () => sendTableToSheet(c, t) },
      { label: 'Copy name', icon: 'copy', onclick: () => (copyText(t.name), toast('Copied')) },
    ],
    e.clientX,
    e.clientY,
  )
}
async function sendTableToSheet(c, t) {
  try {
    const r = await Host.tool('app_db_query', { connection: c.id, sql: `select * from ${qtable(c, t)}`, limit: 20000 })
    placeResultOnSheet(r, t.name)
  } catch (e) {
    toast(e.message, 'err')
  }
}
function placeResultOnSheet(result, label = 'result') {
  const at = h('input.input', { value: 'A1', style: { width: '90px' } })
  const where = sel([['current', `Current sheet (${sheet().name})`], ['new', 'A new sheet']], 'current')
  dialog({
    title: 'Send to a sheet',
    width: 420,
    body: h('.col', h('p.hint', `${result.rowCount} rows × ${result.columns.length} columns${result.truncated ? ' (truncated)' : ''}`), field('Where', where), field('Top-left cell', at)),
    buttons: [
      { label: 'Cancel' },
      {
        label: 'Place on sheet',
        primary: true,
        onclick: () => {
          const top = XL.parseAddr(at.value) || { r: 0, c: 0 }
          const ops = []
          if (where.value === 'new') {
            const nm = String(label).replace(/[[\]:*?/\\]/g, '').slice(0, 28) || 'Data'
            ops.push({ op: 'addSheet', sheet: { name: nm } })
            applyOps(ops, { label: 'Add sheet' })
            A.si = A.book.sheets.length - 1
          }
          const sh = sheet()
          const cells = {}
          result.columns.forEach((n, j) => (cells[A1(top.r, top.c + j)] = { v: n, s: { b: true, bg: '#eef2f7', bd: { b: 'thin' } } }))
          result.rows.forEach((row, i) =>
            row.forEach((v, j) => {
              if (v === null || v === undefined) return
              if (typeof v === 'string' && /^\d{4}-\d{2}-\d{2}( \d{2}:\d{2}(:\d{2})?)?$/.test(v.replace('T', ' ').replace(/Z$/, ''))) {
                const p = XL.parseInput(v.replace('T', ' ').replace(/Z$/, ''))
                if (p.cell && typeof p.cell.v === 'number') return void (cells[A1(top.r + 1 + i, top.c + j)] = { v: p.cell.v, s: { nf: p.nf || 'yyyy-mm-dd' } })
              }
              cells[A1(top.r + 1 + i, top.c + j)] = { v }
            }),
          )
          applyOps([{ op: 'set', sheet: sh.id, cells }], { label: 'Place query result' })
          setMode('sheet')
          setSel(top.r, top.c, top.r + result.rows.length, top.c + result.columns.length - 1)
          toast(`Placed ${result.rows.length} rows on ${sh.name}`)
        },
      },
    ],
  })
}

// ── connection form ──────────────────────────────────────────────────────────────────────────────────────────────────────
function connForm(existing, preset = {}) {
  const c = { type: 'sqlite', color: DB_COLORS[A.conns.length % DB_COLORS.length], ssl: false, ...preset, ...(existing || {}) }
  const body = h('.col')
  const info = h('.test-result')
  const inputs = {}
  const inp = (k, attrs = {}) => (inputs[k] = h('input.input', { value: c[k] ?? '', oninput: (e) => (c[k] = e.target.value), ...attrs }))
  const draw = () => {
    body.replaceChildren(
      h('.seg', Object.entries(TYPE_LABEL).map(([t, l]) => h('button.seg-btn', { class: c.type === t ? 'on' : '', disabled: !!existing, onclick: () => ((c.type = t), (c.port = ''), draw()) }, l))),
      field('Connection name', inp('name', { placeholder: c.type === 'sqlite' ? 'Sales database' : 'Production (read replica)' })),
      c.type === 'sqlite'
        ? field('Database file', inp('file', { placeholder: '/home/you/data/sales.db   or   :memory:' }), 'A path on the computer that runs Sheets. A missing file is created.')
        : h('.col', h('.fc-row', field('Host', inp('host', { placeholder: 'db.example.com' })), field('Port', inp('port', { placeholder: c.type === 'postgres' ? '5432' : '3306', style: { width: '90px' } }))), field('Database', inp('database')), h('.fc-row', field('User', inp('user', { autocomplete: 'off' })), field('Password', h('input.input', { type: 'password', autocomplete: 'new-password', placeholder: existing?.hasPassword ? '•••••••• (saved, leave blank to keep)' : '', oninput: (e) => (c.password = e.target.value) }))), field('Encryption', sel([['', 'None'], ['true', 'SSL / TLS'], ['insecure', 'SSL, skip certificate check']], c.ssl === true ? 'true' : c.ssl || '', { onchange: (e) => (c.ssl = e.target.value === 'true' ? true : e.target.value || false) }))),
      h('.fc-row', field('Colour tag', h('.color-tags', DB_COLORS.map((col) => h('button.swatch', { class: c.color === col ? 'on' : '', style: { background: col }, onclick: () => ((c.color = col), draw()) })))), h('label.check', h('input', { type: 'checkbox', checked: !!c.readOnly, onchange: (e) => (c.readOnly = e.target.checked) }), 'Read-only (never runs writes)')),
      info,
    )
  }
  draw()
  const payload = () => ({ id: existing?.id, type: c.type, name: c.name, color: c.color, readOnly: !!c.readOnly, file: c.file, host: c.host, port: c.port, database: c.database, user: c.user, ssl: c.ssl, password: c.password })
  dialog({
    title: existing ? `Edit ${existing.name}` : 'New connection',
    width: 520,
    body,
    buttons: [
      {
        label: 'Test',
        onclick: async () => {
          info.className = 'test-result'
          info.textContent = 'Testing…'
          try {
            const r = await Host.tool('app_db_test', { connection: payload() })
            info.className = `test-result ${r.ok ? 'ok' : 'bad'}`
            info.textContent = r.ok ? `Connected in ${r.ms} ms · ${r.server}` : r.error
          } catch (e) {
            info.className = 'test-result bad'
            info.textContent = e.message
          }
          return false
        },
      },
      { label: 'Cancel' },
      {
        label: 'Save',
        primary: true,
        onclick: async () => {
          try {
            const r = await Host.tool('app_db_save', { connection: payload() })
            A.conns = r.connections
            delete D.schemas[r.connection.id]
            renderDb()
            selectConn(r.connection.id)
          } catch (e) {
            info.className = 'test-result bad'
            info.textContent = e.message
            return false
          }
        },
      },
    ],
  })
}

// ── tabs ─────────────────────────────────────────────────────────────────────────────────────────────────────────────────
function newQuery(sql = '', run = false) {
  if (!D.conn) {
    if (!A.conns.length) return connForm()
    D.conn = A.conns[0].id
  }
  const t = { kind: 'query', id: ++D.seq, conn: D.conn, title: `Query ${D.tabs.filter((x) => x.kind === 'query').length + 1}`, sql, result: null, error: '', running: false }
  D.tabs.push(t)
  D.active = D.tabs.length - 1
  renderMain()
  if (run) runQuery(t)
}
function openTable(conn, schema, table, view = 'data') {
  let i = D.tabs.findIndex((t) => t.kind === 'table' && t.conn === conn && t.schema === schema && t.table === table)
  if (i < 0) {
    D.tabs.push({ kind: 'table', id: ++D.seq, conn, schema, table, title: table, view, page: 0, size: 200, sort: null, where: '', data: null, error: '', loading: false })
    i = D.tabs.length - 1
  } else D.tabs[i].view = view
  D.active = i
  renderMain()
  if (!D.tabs[i].data) loadTable(D.tabs[i])
}
async function loadTable(t) {
  t.loading = true
  t.error = ''
  renderMain()
  try {
    t.data = await Host.tool('app_db_rows', { connection: t.conn, schema: t.schema, table: t.table, limit: t.size, offset: t.page * t.size, sort: t.sort || undefined, where: t.where || undefined })
  } catch (e) {
    t.error = e.message
  }
  t.loading = false
  t.grid = null
  renderMain()
}
const closeTab = (i) => {
  const t = D.tabs[i]
  if (t.grid && t.grid.count() && !confirm('Discard unsaved changes?')) return
  D.tabs.splice(i, 1)
  D.active = Math.min(D.active, D.tabs.length - 1)
  if (D.active < 0 && D.tabs.length) D.active = 0
  renderMain()
}
function renderTabsBar() {
  return h(
    '.db-tabs',
    D.tabs.map((t, i) => h('.db-tab', { class: i === D.active ? 'on' : '', onclick: () => ((D.active = i), renderMain()), onauxclick: () => closeTab(i) }, icon(t.kind === 'query' ? 'play' : 'table', 13), h('span', t.title), h('button.icon-btn', { onclick: (e) => (e.stopPropagation(), closeTab(i)) }, icon('x', 12)))),
    h('button.db-tab-add', { title: 'New query', onclick: () => newQuery() }, icon('plus', 15)),
  )
}

// ── SQL editor ───────────────────────────────────────────────────────────────────────────────────────────────────────────
function highlight(sql) {
  const re = /(--[^\n]*|\/\*[\s\S]*?\*\/)|('(?:[^']|'')*'?)|("(?:[^"]|"")*"?|`[^`]*`?)|(\b\d+(?:\.\d+)?\b)|([A-Za-z_][\w$]*)/g
  let out = ''
  let last = 0
  let m
  while ((m = re.exec(sql))) {
    out += esc(sql.slice(last, m.index))
    const t = m[0]
    if (m[1]) out += `<i class="c">${esc(t)}</i>`
    else if (m[2]) out += `<i class="s">${esc(t)}</i>`
    else if (m[3]) out += `<i class="q">${esc(t)}</i>`
    else if (m[4]) out += `<i class="n">${esc(t)}</i>`
    else out += KW.has(t.toLowerCase()) ? `<i class="k">${esc(t)}</i>` : esc(t)
    last = m.index + t.length
  }
  return `${out + esc(sql.slice(last))}\n`
}
// the statement the caret is in, or the selection
function currentStatement(ta) {
  if (ta.selectionStart !== ta.selectionEnd) return ta.value.slice(ta.selectionStart, ta.selectionEnd).trim()
  const text = ta.value
  const stmts = []
  let start = 0
  let q = null
  for (let i = 0; i < text.length; i++) {
    const ch = text[i]
    if (q) {
      if (ch === q) q = null
    } else if (ch === "'" || ch === '"') q = ch
    else if (ch === '-' && text[i + 1] === '-') i = text.indexOf('\n', i) < 0 ? text.length : text.indexOf('\n', i)
    else if (ch === ';') {
      stmts.push([start, i + 1])
      start = i + 1
    }
  }
  stmts.push([start, text.length])
  const pos = ta.selectionStart
  const hit = stmts.find(([a, b]) => pos >= a && pos <= b && text.slice(a, b).trim()) || [...stmts].reverse().find(([a, b]) => text.slice(a, b).trim())
  return hit ? text.slice(hit[0], hit[1]).trim() : ''
}
function formatSql(sql) {
  const breaks = /\b(select|from|where|group by|order by|having|limit|offset|union all|union|left join|right join|inner join|full join|cross join|join|on|set|values|returning)\b/gi
  return sql
    .replace(/\s+/g, ' ')
    .trim()
    .replace(breaks, (m, k, off) => (off === 0 ? k.toUpperCase() : `\n${k.toUpperCase()}`))
    .replace(/\b(and|or)\b/gi, (m) => `\n  ${m.toUpperCase()}`)
}
function sqlEditor(t, onRun) {
  const pre = h('pre.sql-hl', { 'aria-hidden': true })
  const ta = h('textarea.sql-ta', { spellcheck: false, value: t.sql, placeholder: 'select * from …    (Ctrl+Enter runs the statement under the cursor)' })
  const gutter = h('.sql-gutter')
  const wrapEl = h('.sql-editor', gutter, h('.sql-stack', pre, ta))
  let ac = null
  let acList = []
  let acIdx = 0
  let acTouched = false
  const sync = () => {
    t.sql = ta.value
    pre.innerHTML = highlight(ta.value)
    const n = ta.value.split('\n').length
    gutter.textContent = Array.from({ length: n }, (_, i) => i + 1).join('\n')
    pre.style.minHeight = `${ta.scrollHeight}px`
  }
  const scroll = () => {
    pre.style.transform = `translate(${-ta.scrollLeft}px, ${-ta.scrollTop}px)`
    gutter.style.transform = `translateY(${-ta.scrollTop}px)`
  }
  ta.addEventListener('scroll', scroll)
  ta.addEventListener('input', () => {
    sync()
    suggest()
  })
  const closeAcList = () => {
    ac?.remove()
    ac = null
    acList = []
  }
  const caretXY = () => {
    const r = ta.getBoundingClientRect()
    const before = ta.value.slice(0, ta.selectionStart).split('\n')
    return { x: r.left + 12 + Math.min(before[before.length - 1].length * 7.4, r.width - 200) - ta.scrollLeft, y: r.top + 8 + before.length * 20 - ta.scrollTop }
  }
  function suggest(force) {
    closeAcList()
    const pos = ta.selectionStart
    const m = ta.value.slice(0, pos).match(/([A-Za-z_][\w$]*\.)?([A-Za-z_][\w$]*)$/)
    if (!m || (!force && m[2].length < 2 && !m[1])) return
    const q = m[2].toLowerCase()
    const tables = tablesOf(t.conn)
    let cands = []
    if (m[1]) {
      const tn = m[1].slice(0, -1).toLowerCase()
      const alias = new RegExp(`\\b(?:from|join)\\s+([\\w."\`]+)(?:\\s+as)?\\s+${tn}\\b`, 'i').exec(ta.value)
      const real = (alias ? alias[1] : tn).replace(/["`]/g, '').split('.').pop().toLowerCase()
      const tb = tables.find((x) => x.name.toLowerCase() === real)
      cands = tb ? tb.columns.map((c) => ({ n: c.name, d: c.type, k: 'col' })) : []
    } else {
      const mentioned = [...ta.value.matchAll(/\b(?:from|join|update|into)\s+([\w."`]+)/gi)].map((x) => x[1].replace(/["`]/g, '').split('.').pop().toLowerCase())
      const cols = tables.filter((x) => mentioned.includes(x.name.toLowerCase())).flatMap((x) => x.columns.map((c) => ({ n: c.name, d: `${x.name} · ${c.type}`, k: 'col' })))
      cands = [...cols, ...tables.map((x) => ({ n: x.name, d: x.kind, k: 'table' })), ...SQL_KEYWORDS.map((k) => ({ n: k.toUpperCase(), d: 'keyword', k: 'kw' }))]
    }
    const seen = new Set()
    acList = cands.filter((c) => c.n.toLowerCase().startsWith(q) && c.n.toLowerCase() !== q && !seen.has(c.n + c.k) && seen.add(c.n + c.k)).slice(0, 8)
    if (!acList.length) return
    acIdx = 0
    acTouched = false
    const { x, y } = caretXY()
    ac = h('.ac', { style: { left: `${x}px`, top: `${y}px` }, onmousedown: (e) => e.preventDefault() }, acList.map((c, i) => h('.ac-item', { class: i === 0 ? 'on' : '', onclick: () => accept(i) }, h('span.ac-name', c.n), h('span.ac-desc', c.d))))
    document.body.append(ac)
  }
  function accept(i) {
    const c = acList[i]
    if (!c) return
    const pos = ta.selectionStart
    const m = ta.value.slice(0, pos).match(/([A-Za-z_][\w$]*)$/)
    const start = pos - (m ? m[1].length : 0)
    const word = c.k === 'kw' ? c.n.toLowerCase() : c.n
    ta.setRangeText(word, start, pos, 'end')
    closeAcList()
    sync()
    ta.focus()
  }
  ta.addEventListener('keydown', (e) => {
    if (ac) {
      if (e.key === 'ArrowDown' || e.key === 'ArrowUp') {
        e.preventDefault()
        acTouched = true
        acIdx = (acIdx + (e.key === 'ArrowDown' ? 1 : -1) + acList.length) % acList.length
        $$('.ac-item', ac).forEach((x, i) => x.classList.toggle('on', i === acIdx))
        return
      }
      // Tab takes a suggestion; Enter only after the list was walked with the arrows, so Enter and Ctrl+Enter keep their own meaning
      if (e.key === 'Tab' || (e.key === 'Enter' && acTouched && !e.ctrlKey && !e.metaKey)) {
        e.preventDefault()
        return accept(acIdx)
      }
      if (e.key === 'Escape') return closeAcList()
      if (e.key === 'Enter') closeAcList()
    }
    if ((e.ctrlKey || e.metaKey) && e.key === 'Enter') {
      e.preventDefault()
      return onRun(currentStatement(ta))
    }
    if ((e.ctrlKey || e.metaKey) && e.key === ' ') {
      e.preventDefault()
      return suggest(true)
    }
    if (e.key === 'Tab') {
      e.preventDefault()
      ta.setRangeText('  ', ta.selectionStart, ta.selectionEnd, 'end')
      sync()
    }
    if ((e.ctrlKey || e.metaKey) && e.key === '/') {
      e.preventDefault()
      const a = ta.value.lastIndexOf('\n', ta.selectionStart - 1) + 1
      const line = ta.value.slice(a, ta.value.indexOf('\n', ta.selectionStart) < 0 ? undefined : ta.value.indexOf('\n', ta.selectionStart))
      const next = /^\s*--\s?/.test(line) ? line.replace(/^(\s*)--\s?/, '$1') : `-- ${line}`
      ta.setRangeText(next, a, a + line.length, 'end')
      sync()
    }
  })
  ta.addEventListener('blur', () => setTimeout(closeAcList, 150))
  sync()
  return { el: wrapEl, ta, sync, setSql: (s) => ((ta.value = s), sync()), focus: () => ta.focus() }
}

// ── running queries ──────────────────────────────────────────────────────────────────────────────────────────────────────
const DESTRUCTIVE = /^\s*(drop|truncate)\b|^\s*(delete|update)\b(?![\s\S]*\bwhere\b)/i
async function runQuery(t, stmt) {
  const sql = (stmt ?? t.editor?.ta?.value ?? t.sql).trim()
  if (!sql) return toast('Write a query first.', 'warn')
  t.running = true
  t.error = ''
  t.ranSql = sql
  renderMain()
  const t0 = Date.now()
  try {
    let r
    try {
      r = await Host.tool('app_db_query', { connection: t.conn, sql, limit: 5000 })
    } catch (e) {
      if (!/changes data|db_execute|read-only/i.test(e.message) || /marked read-only/i.test(e.message)) throw e
      if (!(await confirmDialog(DESTRUCTIVE.test(sql) ? 'Run a destructive statement?' : 'Run a statement that changes data?', h('div', h('p', `This runs on ${connOf(t.conn)?.name}:`), h('pre.sql-preview', sql.slice(0, 600))), 'Run', DESTRUCTIVE.test(sql)))) {
        t.running = false
        return renderMain()
      }
      r = await Host.tool('app_db_exec', { connection: t.conn, sql })
      if (!r.columns) r = { columns: [], rows: [], rowCount: 0, affected: r.affected ?? 0, ms: Date.now() - t0, readOnly: false, message: `Statement ran. ${r.affected ?? 0} rows affected.` }
      else if (/^\s*(create|drop|alter)/i.test(sql)) loadSchema(t.conn, true)
      else if (r.affected !== undefined) r.message = `${r.affected} rows affected.`
    }
    t.result = r
    t.resultGrid = null
    D.history.unshift({ sql, conn: t.conn, at: Date.now(), ms: r.ms, rows: r.rowCount })
    D.history = D.history.slice(0, 60)
    if (/^\s*(create|drop|alter)\b/i.test(sql)) loadSchema(t.conn, true)
  } catch (e) {
    t.error = e.message
    t.result = null
    D.history.unshift({ sql, conn: t.conn, at: Date.now(), error: true })
  }
  t.running = false
  renderMain()
}
async function loadSaved() {
  try {
    D.saved = (await Host.tool('app_queries_list', {})).queries
  } catch {
    D.saved = []
  }
}

function queryView(t) {
  if (!t.editor) {
    t.editor = sqlEditor(t, (stmt) => runQuery(t, stmt))
    t.height = t.height || 230
  }
  const c = connOf(t.conn)
  const connSel = sel(A.conns.map((x) => [x.id, x.name]), t.conn, { class: 'tb-select', onchange: (e) => ((t.conn = e.target.value), (D.conn = t.conn), loadSchema(t.conn), renderSide()) })
  const bar = h(
    '.q-bar',
    h('button.btn.primary.run', { disabled: t.running, onclick: () => runQuery(t, currentStatement(t.editor.ta)) }, icon('play', 14), t.running ? 'Running…' : 'Run'),
    h('span.hint', 'Ctrl+Enter'),
    h('span.tb-sep'),
    connSel,
    h('button.btn', { title: 'Format SQL', onclick: () => t.editor.setSql(formatSql(t.editor.ta.value)) }, icon('wand', 14), 'Format'),
    h('button.btn', { onclick: (e) => historyMenu(t, e.currentTarget) }, icon('history', 14), 'History'),
    h('button.btn', { onclick: (e) => savedMenu(t, e.currentTarget) }, icon('bookmark', 14), 'Saved'),
    h('span.grow'),
    c ? h('span.conn-pill', h('span.conn-dot', { style: { background: c.color || '#64748b' } }), c.name, c.readOnly ? icon('lock', 11) : '') : '',
  )
  const split = h('.q-split', { onmousedown: (e) => dragSplit(e, t) })
  t.editor.el.style.height = `${t.height}px`
  const res = h('.q-result', resultPanel(t))
  const pane = h('.q-pane', bar, t.editor.el, split, res)
  setTimeout(() => t.focusEditor !== false && (t.editor.focus(), (t.focusEditor = false)), 0)
  return pane
}
function dragSplit(e, t) {
  e.preventDefault()
  const y0 = e.clientY
  const h0 = t.height
  const mv = (ev) => {
    t.height = clamp(h0 + ev.clientY - y0, 90, 600)
    t.editor.el.style.height = `${t.height}px`
  }
  const up = () => (removeEventListener('mousemove', mv), removeEventListener('mouseup', up))
  addEventListener('mousemove', mv)
  addEventListener('mouseup', up)
}
function resultPanel(t) {
  if (t.running) return h('.q-note', h('span.spin'), 'Running…')
  if (t.error) return h('.q-note.err', h('b', 'Error'), h('pre', t.error))
  const r = t.result
  if (!r) return h('.q-note', h('p', 'Results appear here.'), h('p.hint', 'Tip: click a table on the left to open it, or write SQL above.'))
  if (!r.columns.length) return h('.q-note.ok', icon('check', 16), r.message || 'Done.', h('span.hint', ` ${r.ms} ms`))
  const g = (t.resultGrid ||= resultGrid({ columns: r.columns, rows: r.rows }))
  return h(
    '.q-res-wrap',
    g.el,
    h(
      '.q-foot',
      h('span', `${r.rowCount.toLocaleString()} row${r.rowCount === 1 ? '' : 's'}${r.truncated ? ' (limited)' : ''}`),
      h('span.hint', `${r.ms} ms`),
      h('span.grow'),
      h('button.btn', { onclick: () => g.copySelection('csv') }, icon('copy', 14), 'Copy CSV'),
      h('button.btn', { onclick: () => placeResultOnSheet(r, t.title) }, icon('sheet', 14), 'Send to sheet'),
    ),
  )
}
function historyMenu(t, el) {
  const items = D.history.length ? D.history.slice(0, 20).map((hh) => ({ label: `${hh.sql.replace(/\s+/g, ' ').slice(0, 60)}${hh.error ? '  (failed)' : ''}`, onclick: () => t.editor.setSql(hh.sql) })) : [{ label: 'Nothing run yet', disabled: true }]
  menuBelow(el, items, { minWidth: 340 })
}
async function savedMenu(t, el) {
  await loadSaved()
  const items = [
    { label: 'Save this query…', icon: 'bookmark', onclick: async () => {
      const name = await promptDialog('Save query', t.title, 'Save')
      if (!name) return
      try {
        D.saved = (await Host.tool('app_queries_save', { name, sql: t.editor.ta.value, connection: t.conn })).queries
        toast('Saved')
      } catch (e) {
        toast(e.message, 'err')
      }
    } },
    { sep: true },
    ...(D.saved.length ? D.saved.map((q) => ({ label: q.name, onclick: () => t.editor.setSql(q.sql) })) : [{ label: 'No saved queries', disabled: true }]),
  ]
  menuBelow(el, items, { minWidth: 260 })
}

// ── table tab ────────────────────────────────────────────────────────────────────────────────────────────────────────────
function tableView(t) {
  const c = connOf(t.conn)
  const sc = D.schemas[t.conn]?.schemas?.find((s) => s.name === t.schema)?.tables.find((x) => x.name === t.table)
  const pkCols = sc ? sc.columns.filter((x) => x.pk).map((x) => x.name) : []
  const canEdit = !!sc && sc.kind !== 'view' && pkCols.length > 0 && !c?.readOnly
  const seg = h('.seg.small', ['data', 'structure'].map((v) => h('button.seg-btn', { class: t.view === v ? 'on' : '', onclick: () => ((t.view = v), renderMain()) }, v === 'data' ? 'Data' : 'Structure')))
  if (t.view === 'structure') {
    return h('.q-pane', h('.q-bar', seg, h('span.grow'), h('button.btn', { onclick: () => newQuery(`select *\nfrom ${qtable(c, { schema: t.schema, name: t.table })}\nlimit 100;`, true) }, icon('play', 14), 'Query this table')), sc ? structureTable(sc) : h('.q-note', 'Loading structure…'))
  }
  const filter = h('input.input.filter-in', { placeholder: "Filter, e.g.  status = 'open' and total > 100", value: t.where, onkeydown: (e) => e.key === 'Enter' && ((t.where = e.target.value.trim()), (t.page = 0), loadTable(t)) })
  const info = t.data
  const pages = info?.total ? Math.ceil(info.total / t.size) : 1
  const dirty = t.grid?.count() || 0
  const bar = h(
    '.q-bar',
    seg,
    h('span.tb-sep'),
    filter,
    h('button.btn', { title: 'Reload', onclick: () => loadTable(t) }, icon('refresh', 14)),
    h('span.grow'),
    canEdit ? h('button.btn', { onclick: () => t.grid?.addRow() }, icon('plus', 14), 'Add row') : '',
    dirty ? h('button.btn', { onclick: () => (t.grid.discard(), renderMain()) }, 'Discard') : '',
    dirty ? h('button.btn.primary', { onclick: () => applyTableChanges(t) }, icon('check', 14), `Apply ${dirty} change${dirty === 1 ? '' : 's'}`) : '',
  )
  let body
  if (t.loading && !info) body = h('.q-note', h('span.spin'), 'Loading…')
  else if (t.error) body = h('.q-note.err', h('b', 'Error'), h('pre', t.error))
  else if (info) {
    const g = (t.grid ||= resultGrid({ columns: info.columns, rows: info.rows, editable: canEdit, pk: pkCols, pkCols, types: info.columns.map((n) => sc?.columns.find((x) => x.name === n)?.type || ''), offset: t.page * t.size, table: t.table, sort: t.sort, onSort: (s) => ((t.sort = s), (t.page = 0), loadTable(t)), onChanges: () => refreshTableBar(t), filterBy: (col, v) => ((t.where = v === null ? `${col} is null` : typeof v === 'number' ? `${col} = ${v}` : `${col} = '${String(v).replace(/'/g, "''")}'`), (t.page = 0), loadTable(t)) }))
    body = h(
      '.q-res-wrap',
      !canEdit && sc ? h('.banner', icon('info', 14), c?.readOnly ? 'This connection is read-only.' : sc.kind === 'view' ? 'Views cannot be edited.' : 'This table has no primary key, so rows cannot be edited safely. Use SQL.') : '',
      g.el,
      h('.q-foot', h('span', info.total !== null ? `${(t.page * t.size + 1).toLocaleString()}–${Math.min(info.total, t.page * t.size + info.rowCount).toLocaleString()} of ${info.total.toLocaleString()}` : `${info.rowCount} rows`), h('span.hint', `${info.ms} ms`), h('span.grow'), h('button.btn', { disabled: t.page === 0, onclick: () => ((t.page -= 1), loadTable(t)) }, 'Previous'), h('span.hint', `Page ${t.page + 1} of ${pages}`), h('button.btn', { disabled: t.page + 1 >= pages, onclick: () => ((t.page += 1), loadTable(t)) }, 'Next'), h('button.btn', { onclick: () => placeResultOnSheet(info, t.table) }, icon('sheet', 14), 'Send to sheet')),
    )
  } else body = h('.q-note', 'No data')
  t.barEl = bar
  return h('.q-pane', bar, body)
}
function refreshTableBar(t) {
  if (curTab() !== t) return
  const pane = $('.db-pane')
  if (!pane || !t.barEl) return
  const nb = tableView(t)
  // keep the grid element in place: only swap the toolbar
  t.barEl.replaceWith(nb.firstChild)
  t.barEl = pane.querySelector('.q-bar')
}
async function applyTableChanges(t) {
  const changes = t.grid.collect()
  if (!changes.length) return
  if (!(await confirmDialog('Apply changes', `Write ${changes.length} change${changes.length === 1 ? '' : 's'} to ${t.table}? This changes the database.`, 'Apply', false))) return
  try {
    const r = await Host.tool('app_db_exec', { connection: t.conn, schema: t.schema, table: t.table, changes })
    toast(`${r.applied} change${r.applied === 1 ? '' : 's'} saved`)
    t.grid = null
    loadTable(t)
  } catch (e) {
    toast(e.message, 'err')
  }
}
function structureTable(sc) {
  return h(
    '.struct',
    h('table.stbl', h('thead', h('tr', ['#', 'Column', 'Type', 'Nullable', 'Key', 'Default', 'References'].map((x) => h('th', x)))), h('tbody', sc.columns.map((c, i) => h('tr', h('td.hint', i + 1), h('td', h('b', c.name)), h('td', h('code', c.type || '—')), h('td', c.nullable ? 'yes' : 'no'), h('td', c.pk ? 'PRIMARY' : ''), h('td', h('code', c.default ?? '')), h('td', c.fk || ''))))),
  )
}

// ── main area ────────────────────────────────────────────────────────────────────────────────────────────────────────────
function renderMain() {
  const main = $('.db-main')
  if (!main) return
  const t = curTab()
  let pane
  if (!t) {
    pane = A.conns.length
      ? h('.db-empty', h('h2', 'Database workbench'), h('p', 'Browse tables, run SQL, edit rows, and send results to a sheet.'), h('.db-empty-actions', h('button.btn.primary', { onclick: () => newQuery() }, icon('play', 14), 'New query'), h('button.btn', { onclick: () => connForm() }, icon('plug', 14), 'New connection')), D.conn ? h('p.hint', 'Pick a table on the left to open it.') : h('p.hint', 'Pick a connection on the left to begin.'))
      : h('.db-empty', h('h2', 'Connect a database'), h('p', 'SQLite files, PostgreSQL and MySQL. Passwords are encrypted on disk and never shown to the assistant.'), h('.db-empty-cards', [['sqlite', 'SQLite file', 'A local .db or .sqlite file'], ['postgres', 'PostgreSQL', 'A server on your network or cloud'], ['mysql', 'MySQL / MariaDB', 'A server on your network or cloud']].map(([ty, name, sub]) => h('button.db-card', { onclick: () => connForm(null, { type: ty }) }, icon('database', 22), h('b', name), h('span.hint', sub)))))
  } else if (t.kind === 'query') pane = queryView(t)
  else pane = tableView(t)
  main.replaceChildren(renderTabsBar(), h('.db-pane', pane))
}
function renderDb() {
  renderSide()
  renderMain()
}
function buildDbView() {
  const v = h('.db-view', h('.db-side'), h('.db-main'))
  return v
}
