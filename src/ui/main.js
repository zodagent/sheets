// ── boot: the shell, the two modes, the theme, the host handshake ────────────────────────────────────────────────────────────────────────
let sheetHost
let chartLayerEl
// the frame is as tall as the screen allows, so the sheet fills the page below Zod's own header and scrolls with the page
const APP_HEIGHT = Math.max(640, Math.min(1000, (screen.availHeight || 900) - 230))

function setMode(m) {
  if (A.edit) commitEdit()
  A.mode = m
  $('.sheet-view').style.display = m === 'sheet' ? 'flex' : 'none'
  $('.db-view').style.display = m === 'db' ? 'grid' : 'none'
  $$('.mode-btn').forEach((b) => b.classList.toggle('on', b.dataset.mode === m))
  if (m === 'db') {
    if (!D.conn && A.conns.length) D.conn = A.dbConn || A.conns[0].id
    if (D.conn) loadSchema(D.conn)
    renderDb()
  } else {
    schedule()
    scroller.focus({ preventScroll: true })
  }
}
App.setMode = setMode

function applyTheme(ctx) {
  const root = document.documentElement
  const v = ctx?.styles?.variables || {}
  const bg = parseColor(v['--color-background-primary'])
  let dark = ctx?.theme ? ctx.theme === 'dark' : matchMedia('(prefers-color-scheme: dark)').matches
  if (bg) dark = bg.l < 0.5
  root.dataset.theme = dark ? 'dark' : 'light'
  root.style.colorScheme = dark ? 'dark' : 'light'
  // like Books: the host's background and text, and every other colour a tint of those two (see app.css)
  const set = (k, val) => (val ? root.style.setProperty(k, val) : root.style.removeProperty(k))
  set('--font', v['--font-sans'])
  set('--c-surface', v['--color-background-primary'])
  set('--c-text', v['--color-text-primary'])
  for (const [k, from] of [['--color-text-info', '--color-text-info'], ['--color-text-danger', '--color-text-danger'], ['--color-text-success', '--color-text-success'], ['--color-text-warning', '--color-text-warning']]) set(k, v[from])
  if (canvas) {
    readColors()
    schedule()
  }
}

function workbookMenu(el) {
  const items = [
    ...A.wbs.map((w) => ({ label: w.name, checked: w.id === A.book.id, onclick: () => switchWorkbook(w.id) })),
    { sep: true },
    {
      label: 'New workbook',
      icon: 'plus',
      onclick: async () => {
        const name = await promptDialog('New workbook', 'Untitled workbook', 'Create')
        if (name) workbookAction({ action: 'create', name })
      },
    },
    {
      label: 'Rename…',
      icon: 'edit',
      onclick: async () => {
        const name = await promptDialog('Rename workbook', A.book.name, 'Rename')
        if (name) workbookAction({ action: 'rename', id: A.book.id, name })
      },
    },
    { label: 'Duplicate', icon: 'copy', onclick: () => workbookAction({ action: 'duplicate', id: A.book.id }) },
    { label: 'Import pasted CSV…', icon: 'upload', onclick: () => App.openImport() },
    { label: 'Copy sheet as CSV…', icon: 'download', onclick: () => App.openExport() },
    { sep: true },
    {
      label: 'Delete workbook…',
      icon: 'trash',
      disabled: A.wbs.length < 2,
      onclick: async () => {
        if (await confirmDialog('Delete workbook', `Delete "${A.book.name}" for good? This cannot be undone.`)) workbookAction({ action: 'delete', id: A.book.id })
      },
    },
  ]
  menuBelow(el, items, { minWidth: 230 })
}
async function switchWorkbook(id) {
  if (A.pending.length) await Save.flush()
  const r = await Host.tool('app_state', { workbook: id })
  loadWorkbook(r.workbook)
  A.wbs = r.workbooks
  A.conns = r.connections
}
async function workbookAction(args) {
  try {
    if (A.pending.length) await Save.flush()
    const r = await Host.tool('app_workbook', args)
    A.wbs = r.workbooks
    loadWorkbook(r.workbook, { keepView: args.action === 'rename' })
  } catch (e) {
    toast(e.message, 'err')
  }
}
App.refreshTop = () => {
  const el = $('#wb-name')
  if (el) el.textContent = A.book?.name || ''
}

function buildShell() {
  const root = $('#root')
  const wbBtn = h('button.wb-btn', { onclick: (e) => workbookMenu(e.currentTarget), title: 'Workbooks' }, h('span#wb-name', ''), icon('down', 14))
  const top = h(
    '.topbar',
    wbBtn,
    h('.mode', h('button.mode-btn', { 'data-mode': 'sheet', onclick: () => setMode('sheet') }, 'Spreadsheet'), h('button.mode-btn', { 'data-mode': 'db', onclick: () => setMode('db') }, 'Database')),
    h('span.grow'),
    h('button#save-status.save-status', { onclick: () => Save.retry() }, 'All changes saved'),
    h('button.icon-btn.top-ico', { title: 'Keyboard shortcuts', onclick: () => App.openShortcuts() }, icon('info', 17)),
  )
  const sv = buildSheetView()
  const dv = buildDbView()
  root.append(h('.app', top, sv.view, dv))
  sheetHost = sv.host
  chartLayerEl = $('.chart-layer')
  gridInit(sheetHost)
  wrap.append(chartLayerEl)
  dv.style.display = 'none'
}

async function boot() {
  applyTheme(null)
  buildShell()
  Host.on('ui/notifications/host-context-changed', applyTheme)
  Host.on('ui/notifications/tool-result', (r) => {
    const s = r.structuredContent
    if (s?.workbook && !A.pending.length) {
      A.wbs = s.workbooks || A.wbs
      A.conns = s.connections || A.conns
      loadWorkbook(s.workbook)
      if (s.view === 'database') setMode('db')
    }
  })
  Host.on('ui/resource-teardown', () => {})
  let ctx = null
  try {
    const r = await Host.request('ui/initialize', { protocolVersion: '2025-11-21', appInfo: { name: 'zod-sheets', version: '1.0.0' }, appCapabilities: {} })
    ctx = r.hostContext
    Host.notify('ui/notifications/initialized', {})
    Host.ready = true
  } catch {
    document.body.textContent = 'This page runs inside an MCP host.'
    return
  }
  applyTheme(ctx)
  Host.resize(APP_HEIGHT)
  setMode('sheet')
  // an app opened from a tool call receives its workbook; one opened from the Apps list asks for it
  setTimeout(async () => {
    if (A.book) return
    try {
      const r = await Host.tool('app_state', {})
      if (A.book) return
      A.wbs = r.workbooks
      A.conns = r.connections
      loadWorkbook(r.workbook)
    } catch (e) {
      $('#root').prepend(h('.fatal', `Could not load: ${e.message}`))
    }
  }, 500)
}
addEventListener('beforeunload', () => A.pending.length && Save.flush())
boot()
