// ── the grid: a canvas drawn sheet with frozen panes, merges, formats, selection, editing, fill, clipboard and undo ────────────────────
const { MAXR, MAXC } = XL
const A = {
  book: null,
  wbs: [],
  conns: [],
  funcs: [],
  mode: 'sheet',
  si: 0,
  sel: { ar: 0, ac: 0, fr: 0, fc: 0 },
  views: {},
  zoom: 1,
  edit: null,
  clip: null,
  undo: [],
  redo: [],
  pending: [],
  saving: false,
  saveError: '',
  rev: 0,
  showFormulas: false,
  findHits: [],
  charts: { sel: null },
}
const DEF_W = 104
const DEF_H = 24
const HDR_W = 54
const HDR_H = 26
const sheet = () => A.book.sheets[A.si]
const key = (r, c) => r * MAXC + c
const A1 = (r, c) => XL.addr(r, c)

class Axis {
  constructor(def) {
    this.def = def
    this.n = 1000
    this.idx = []
    this.pre = [0]
    this.sz = new Map()
  }
  set(map) {
    this.sz = map
    this.idx = [...map.keys()].sort((a, b) => a - b)
    this.pre = [0]
    let s = 0
    for (const i of this.idx) {
      s += map.get(i) - this.def
      this.pre.push(s)
    }
  }
  cnt(i) {
    let lo = 0
    let hi = this.idx.length
    while (lo < hi) {
      const m = (lo + hi) >> 1
      if (this.idx[m] < i) lo = m + 1
      else hi = m
    }
    return lo
  }
  pos(i) {
    return i * this.def + this.pre[this.cnt(i)]
  }
  size(i) {
    const v = this.sz.get(i)
    return v === undefined ? this.def : v
  }
  at(p) {
    let lo = 0
    let hi = this.n - 1
    while (lo < hi) {
      const m = (lo + hi + 1) >> 1
      if (this.pos(m) <= p) lo = m
      else hi = m - 1
    }
    return lo
  }
}
const colAx = new Axis(DEF_W)
const rowAx = new Axis(DEF_H)
let canvas
let ctx
let scroller
let spacer
let wrap
let editor
let tipEl
let dpr = 1
let viewW = 0
let viewH = 0
let merges = []
let mergeAt = new Map()
let colors = {}
let dirty = false
let hiddenByFilter = new Set()
const Z = () => A.zoom

// ── view state ────────────────────────────────────────────────────────────────────────────────────────────────────────────
const norm = () => {
  const s = A.sel
  return { r1: Math.min(s.ar, s.fr), c1: Math.min(s.ac, s.fc), r2: Math.max(s.ar, s.fr), c2: Math.max(s.ac, s.fc) }
}
const readColors = () => {
  const cs = getComputedStyle(document.documentElement)
  const v = (n) => cs.getPropertyValue(n).trim()
  colors = { bg: v('--c-surface'), head: v('--c-head'), headText: v('--c-head-text'), grid: v('--c-grid'), text: v('--c-text'), accent: v('--c-accent'), accentSoft: v('--c-accent-soft'), headSel: v('--c-head-sel'), err: v('--c-err'), border: v('--c-border'), font: cs.getPropertyValue('--font').trim() || 'system-ui, sans-serif' }
}
function computeFilter() {
  hiddenByFilter = new Set()
  const sh = sheet()
  const f = sh.filter
  if (!f?.range) return
  const g = XL.parseRange(f.range)
  if (!g) return
  for (let r = g.r1 + 1; r <= g.r2; r++)
    for (const [off, crit] of Object.entries(f.cols || {})) {
      if (!crit?.values) continue
      const txt = A.book.display(A.si, r, g.c1 + Number(off))
      if (!crit.values.includes(txt === '' ? '(Blanks)' : txt)) {
        hiddenByFilter.add(r)
        break
      }
    }
}
function buildAxes() {
  const sh = sheet()
  const z = Z()
  colAx.def = DEF_W * z
  rowAx.def = DEF_H * z
  const cm = new Map()
  for (const [k, w] of Object.entries(sh.cols)) cm.set(XL.colIndex(k), w * z)
  for (const c of sh.hiddenCols) cm.set(c, 0)
  const rm = new Map()
  for (const [k, w] of Object.entries(sh.rows)) rm.set(Number(k) - 1, w * z)
  for (const r of sh.hiddenRows) rm.set(r, 0)
  computeFilter()
  for (const r of hiddenByFilter) rm.set(r, 0)
  colAx.set(cm)
  rowAx.set(rm)
  const used = A.book.usedRange(sh.id)
  const need = (used?.r2 ?? 0) + 200
  const needC = (used?.c2 ?? 0) + 30
  rowAx.n = clamp(Math.max(rowAx.n, need, 1000), 1, MAXR)
  colAx.n = clamp(Math.max(colAx.n, needC, 52), 1, MAXC)
  merges = sh.merges.map((m) => XL.parseRange(m)).filter(Boolean)
  mergeAt = new Map()
  merges.forEach((m, i) => {
    for (let r = m.r1; r <= Math.min(m.r2, m.r1 + 500); r++) for (let c = m.c1; c <= Math.min(m.c2, m.c1 + 200); c++) mergeAt.set(key(r, c), i)
  })
  spacer.style.width = `${HDR_W + colAx.pos(colAx.n) + 40}px`
  spacer.style.height = `${HDR_H + rowAx.pos(rowAx.n) + 40}px`
}
const mergeOf = (r, c) => {
  const i = mergeAt.get(key(r, c))
  return i === undefined ? null : merges[i]
}
// the box a cell (or its merged block) occupies on screen
function rect(r, c, merged = true) {
  const sh = sheet()
  const m = merged ? mergeOf(r, c) : null
  const r1 = m ? m.r1 : r
  const c1 = m ? m.c1 : c
  const r2 = m ? m.r2 : r
  const c2 = m ? m.c2 : c
  const fc = sh.freeze.cols
  const fr = sh.freeze.rows
  const sx = scroller.scrollLeft
  const sy = scroller.scrollTop
  const x = HDR_W + colAx.pos(c1) - (c1 >= fc ? sx : 0)
  const y = HDR_H + rowAx.pos(r1) - (r1 >= fr ? sy : 0)
  const x2 = HDR_W + colAx.pos(c2 + 1) - (c2 + 1 > fc ? sx : 0)
  const y2 = HDR_H + rowAx.pos(r2 + 1) - (r2 + 1 > fr ? sy : 0)
  return { x, y, w: Math.max(0, x2 - x), h: Math.max(0, y2 - y) }
}
function hitCell(px, py) {
  const sh = sheet()
  const FW = colAx.pos(sh.freeze.cols)
  const FH = rowAx.pos(sh.freeze.rows)
  const wx = px - HDR_W < FW ? px - HDR_W : px - HDR_W + scroller.scrollLeft
  const wy = py - HDR_H < FH ? py - HDR_H : py - HDR_H + scroller.scrollTop
  return { r: rowAx.at(Math.max(0, wy)), c: colAx.at(Math.max(0, wx)) }
}

// ── drawing ───────────────────────────────────────────────────────────────────────────────────────────────────────────────
const schedule = () => {
  if (dirty) return
  dirty = true
  requestAnimationFrame(() => {
    dirty = false
    render()
  })
}
const fontOf = (s, z) => `${s.i ? 'italic ' : ''}${s.b ? '600 ' : ''}${(s.fs || 13) * z}px ${colors.font}`
const textOf = (r, c) => {
  const book = A.book
  const cell = sheet().cells.get(key(r, c))
  if (A.showFormulas && cell?.f != null) return `=${cell.f}`
  return book.display(A.si, r, c)
}
function wrapLines(text, w, ctx2) {
  const out = []
  for (const para of String(text).split('\n')) {
    let line = ''
    for (const word of para.split(/(\s+)/)) {
      const t = line + word
      if (ctx2.measureText(t).width > w && line) {
        out.push(line.trimEnd())
        line = word.trimStart()
      } else line = t
    }
    out.push(line)
  }
  return out
}
function drawCell(r, c, x, y, w, h2, clipRight) {
  const book = A.book
  const si = A.si
  const sh = sheet()
  const cell = sh.cells.get(key(r, c))
  const v = book.get(si, r, c)
  const st = cell?.s || {}
  const z = Z()
  const cf = sh.cf.length ? XL.condStyle(book, si, r, c) : null
  const bg = cf?.bg || st.bg
  if (bg) {
    ctx.fillStyle = bg
    ctx.fillRect(x, y, w, h2)
  }
  if (cf?.bar) {
    ctx.fillStyle = cf.bar.color
    const bw = Math.max(1, (w - 4) * cf.bar.pct)
    ctx.globalAlpha = 0.45
    ctx.fillRect(x + 2, y + 3, bw, h2 - 6)
    ctx.globalAlpha = 1
  }
  if (v === null || v === undefined || v === '') {
    drawBorders(st, x, y, w, h2)
    return
  }
  let txt = textOf(r, c)
  if (txt === '') {
    drawBorders(st, x, y, w, h2)
    return
  }
  const bold = cf?.b ?? st.b
  const ital = cf?.i ?? st.i
  const s2 = { ...st, b: bold, i: ital }
  ctx.font = fontOf(s2, z)
  const isErr = v instanceof XL.XlErr
  ctx.fillStyle = cf?.fg || st.fg || (isErr ? colors.err : colors.text)
  const isNum = typeof v === 'number'
  const al = st.al || (isNum ? 'right' : typeof v === 'boolean' || isErr ? 'center' : 'left')
  const pad = 5 * z
  const lineH = (st.fs || 13) * z * 1.25
  let tw = w - pad * 2
  ctx.textBaseline = 'alphabetic'
  const fm = (st.fs || 13) * z
  if (st.wrap && typeof v === 'string') {
    const lines = wrapLines(txt, tw, ctx)
    const total = lines.length * lineH
    const va = st.va || 'top'
    let ty = va === 'middle' ? y + (h2 - total) / 2 : va === 'bottom' ? y + h2 - total - 3 : y + 3 * z
    ctx.save()
    ctx.beginPath()
    ctx.rect(x, y, w, h2)
    ctx.clip()
    for (const ln of lines) {
      const lw = ctx.measureText(ln).width
      const tx = al === 'right' ? x + w - pad - lw : al === 'center' ? x + (w - lw) / 2 : x + pad
      ctx.fillText(ln, tx, ty + fm)
      ty += lineH
    }
    ctx.restore()
    drawBorders(st, x, y, w, h2)
    return
  }
  let tx
  let width = ctx.measureText(txt).width
  if (isNum && width > tw) {
    txt = '#'.repeat(Math.max(1, Math.floor(tw / ctx.measureText('#').width)))
    width = ctx.measureText(txt).width
  }
  // text runs into empty neighbours
  let clipX = x
  let clipW = w
  if (!isNum && width > tw && al === 'left' && !st.wrap) {
    let extra = 0
    for (let cc = c + 1; cc < colAx.n && extra < width - tw && x + w + extra < clipRight; cc++) {
      const vv = book.get(si, r, cc)
      if (vv !== null && vv !== undefined && vv !== '') break
      if (mergeOf(r, cc)) break
      extra += colAx.size(cc)
    }
    clipW = w + extra
    tw = clipW - pad * 2
  }
  tx = al === 'right' ? x + w - pad - width : al === 'center' ? x + (w - width) / 2 : x + pad
  const va = st.va || 'middle'
  const base = va === 'top' ? y + 3 * z + fm : va === 'bottom' ? y + h2 - 5 * z : y + h2 / 2 + fm * 0.36
  ctx.save()
  ctx.beginPath()
  ctx.rect(clipX, y, clipW, h2)
  ctx.clip()
  ctx.fillText(txt, tx, base)
  if (st.u || st.st) {
    ctx.fillRect(tx, st.u ? base + 2 * z : base - fm * 0.32, width, Math.max(1, z))
  }
  ctx.restore()
  drawBorders(st, x, y, w, h2)
}
function drawBorders(st, x, y, w, h2) {
  const bd = st.bd
  if (!bd) return
  ctx.strokeStyle = st.bc || colors.text
  const line = (kind, x1, y1, x2, y2) => {
    if (!kind) return
    ctx.lineWidth = kind === 'thick' ? 2 : 1
    ctx.setLineDash(kind === 'dashed' ? [4, 3] : [])
    ctx.beginPath()
    ctx.moveTo(x1, y1)
    ctx.lineTo(x2, y2)
    ctx.stroke()
  }
  line(bd.t, x, y + 0.5, x + w, y + 0.5)
  line(bd.b, x, y + h2 - 0.5, x + w, y + h2 - 0.5)
  line(bd.l, x + 0.5, y, x + 0.5, y + h2)
  line(bd.r, x + w - 0.5, y, x + w - 0.5, y + h2)
  ctx.setLineDash([])
  ctx.lineWidth = 1
}
function drawRegion(rows, cols, clip) {
  const [rA, rB] = rows
  const [cA, cB] = cols
  if (rB < rA || cB < cA) return
  const sh = sheet()
  ctx.save()
  ctx.beginPath()
  ctx.rect(clip.x, clip.y, clip.w, clip.h)
  ctx.clip()
  const x0 = rect(rA, cA, false).x
  const y0 = rect(rA, cA, false).y
  const x1 = rect(rB, cB, false)
  const xe = x1.x + x1.w
  const ye = x1.y + x1.h
  ctx.fillStyle = colors.bg
  ctx.fillRect(clip.x, clip.y, clip.w, clip.h)
  if (sh.gridlines) {
    ctx.strokeStyle = colors.grid
    ctx.lineWidth = 1
    ctx.beginPath()
    for (let r = rA; r <= rB + 1; r++) {
      const y = Math.round(HDR_H + rowAx.pos(r) - (r >= sh.freeze.rows ? scroller.scrollTop : 0)) - 0.5
      ctx.moveTo(x0, y)
      ctx.lineTo(xe, y)
    }
    for (let c = cA; c <= cB + 1; c++) {
      const x = Math.round(HDR_W + colAx.pos(c) - (c >= sh.freeze.cols ? scroller.scrollLeft : 0)) - 0.5
      ctx.moveTo(x, y0)
      ctx.lineTo(x, ye)
    }
    ctx.stroke()
  }
  const clipRight = clip.x + clip.w
  for (let r = rA; r <= rB; r++) {
    const h2 = rowAx.size(r)
    if (!h2) continue
    for (let c = cA; c <= cB; c++) {
      if (mergeAt.has(key(r, c))) continue
      const w = colAx.size(c)
      if (!w) continue
      const b = rect(r, c, false)
      drawCell(r, c, b.x, b.y, b.w, b.h, clipRight)
    }
  }
  merges.forEach((m, i) => {
    if (m.r2 < rA || m.r1 > rB || m.c2 < cA || m.c1 > cB) return
    if (m.r1 < rA && rA > 0 && !(rA === sh.freeze.rows && false)) {
      /* a merge that starts above this region is still drawn from its own corner */
    }
    const b = rect(m.r1, m.c1, true)
    ctx.fillStyle = colors.bg
    ctx.fillRect(b.x, b.y, b.w, b.h)
    drawCell(m.r1, m.c1, b.x, b.y, b.w, b.h, clipRight)
    ctx.strokeStyle = colors.grid
    ctx.strokeRect(b.x - 0.5, b.y - 0.5, b.w, b.h)
    void i
  })
  ctx.restore()
}
const PALETTE = ['#2563eb', '#d97706', '#059669', '#db2777', '#7c3aed', '#0891b2', '#dc2626', '#65a30d']
function render() {
  if (!A.book || !canvas || A.mode !== 'sheet') return
  const sh = sheet()
  A.book.recalc()
  const W = viewW
  const H = viewH
  ctx.setTransform(dpr, 0, 0, dpr, 0, 0)
  ctx.fillStyle = colors.bg
  ctx.fillRect(0, 0, W, H)
  const fr = sh.freeze.rows
  const fc = sh.freeze.cols
  const FW = colAx.pos(fc)
  const FH = rowAx.pos(fr)
  const sx = scroller.scrollLeft
  const sy = scroller.scrollTop
  const c0 = Math.max(fc, colAx.at(FW + sx))
  const c1 = colAx.at(sx + W - HDR_W)
  const r0 = Math.max(fr, rowAx.at(FH + sy))
  const r1 = rowAx.at(sy + H - HDR_H)
  // grow the sheet when scrolled near its end
  if (r1 > rowAx.n - 40 && rowAx.n < MAXR) {
    rowAx.n = Math.min(MAXR, rowAx.n + 500)
    spacer.style.height = `${HDR_H + rowAx.pos(rowAx.n) + 40}px`
  }
  if (c1 > colAx.n - 6 && colAx.n < MAXC) {
    colAx.n = Math.min(MAXC, colAx.n + 26)
    spacer.style.width = `${HDR_W + colAx.pos(colAx.n) + 40}px`
  }
  const gx = HDR_W
  const gy = HDR_H
  drawRegion([r0, r1], [c0, c1], { x: gx + FW, y: gy + FH, w: W - gx - FW, h: H - gy - FH })
  if (fr) drawRegion([0, fr - 1], [c0, c1], { x: gx + FW, y: gy, w: W - gx - FW, h: FH })
  if (fc) drawRegion([r0, r1], [0, fc - 1], { x: gx, y: gy + FH, w: FW, h: H - gy - FH })
  if (fr && fc) drawRegion([0, fr - 1], [0, fc - 1], { x: gx, y: gy, w: FW, h: FH })
  // freeze lines
  ctx.strokeStyle = colors.border
  ctx.lineWidth = 1.5
  if (fc) line(gx + FW, gy, gx + FW, H)
  if (fr) line(gx, gy + FH, W, gy + FH)
  ctx.lineWidth = 1
  ctx.save()
  ctx.beginPath()
  ctx.rect(gx, gy, W - gx, H - gy)
  ctx.clip()
  drawOverlays()
  ctx.restore()
  drawHeaders(r0, r1, c0, c1)
  positionEditor()
  positionCharts()
  updateTip()
}
const line = (x1, y1, x2, y2) => {
  ctx.beginPath()
  ctx.moveTo(x1, y1)
  ctx.lineTo(x2, y2)
  ctx.stroke()
}
function drawHeaders(r0, r1, c0, c1) {
  const sh = sheet()
  const W = viewW
  const H = viewH
  const z = Z()
  const g = norm()
  const FW = colAx.pos(sh.freeze.cols)
  const FH = rowAx.pos(sh.freeze.rows)
  ctx.font = `${12 * z}px ${colors.font}`
  ctx.textBaseline = 'middle'
  ctx.textAlign = 'center'
  ctx.fillStyle = colors.head
  ctx.fillRect(0, 0, W, HDR_H)
  ctx.fillRect(0, 0, HDR_W, H)
  const colSel = (c) => c >= g.c1 && c <= g.c2
  const rowSel = (r) => r >= g.r1 && r <= g.r2
  const wholeCols = g.r1 === 0 && g.r2 >= MAXR - 1
  const wholeRows = g.c1 === 0 && g.c2 >= MAXC - 1
  const drawCol = (c) => {
    const w = colAx.size(c)
    if (!w) return
    const x = HDR_W + colAx.pos(c) - (c >= sh.freeze.cols ? scroller.scrollLeft : 0)
    if (x + w < HDR_W || x > W) return
    if (colSel(c)) {
      ctx.fillStyle = wholeCols ? colors.accent : colors.headSel
      ctx.fillRect(x, 0, w, HDR_H)
    }
    ctx.fillStyle = colSel(c) ? (wholeCols ? '#fff' : colors.accent) : colors.headText
    ctx.fillText(XL.colName(c), x + w / 2, HDR_H / 2 + 1)
    ctx.strokeStyle = colors.border
    line(Math.round(x + w) - 0.5, 0, Math.round(x + w) - 0.5, HDR_H)
    const f = sh.filter
    if (f?.range) {
      const fg = XL.parseRange(f.range)
      if (fg && c >= fg.c1 && c <= fg.c2) {
        /* filter buttons are drawn on the header row of the data, below */
      }
    }
  }
  const drawRow = (r) => {
    const hh = rowAx.size(r)
    if (!hh) return
    const y = HDR_H + rowAx.pos(r) - (r >= sh.freeze.rows ? scroller.scrollTop : 0)
    if (y + hh < HDR_H || y > H) return
    if (rowSel(r)) {
      ctx.fillStyle = wholeRows ? colors.accent : colors.headSel
      ctx.fillRect(0, y, HDR_W, hh)
    }
    ctx.fillStyle = rowSel(r) ? (wholeRows ? '#fff' : colors.accent) : colors.headText
    ctx.fillText(String(r + 1), HDR_W / 2, y + hh / 2 + 1)
    ctx.strokeStyle = colors.border
    line(0, Math.round(y + hh) - 0.5, HDR_W, Math.round(y + hh) - 0.5)
  }
  ctx.save()
  ctx.beginPath()
  ctx.rect(HDR_W + FW, 0, W, HDR_H)
  ctx.clip()
  for (let c = c0; c <= c1; c++) drawCol(c)
  ctx.restore()
  for (let c = 0; c < sh.freeze.cols; c++) drawCol(c)
  ctx.save()
  ctx.beginPath()
  ctx.rect(0, HDR_H + FH, HDR_W, H)
  ctx.clip()
  for (let r = r0; r <= r1; r++) drawRow(r)
  ctx.restore()
  for (let r = 0; r < sh.freeze.rows; r++) drawRow(r)
  ctx.strokeStyle = colors.border
  line(0, HDR_H - 0.5, W, HDR_H - 0.5)
  line(HDR_W - 0.5, 0, HDR_W - 0.5, H)
  ctx.fillStyle = colors.head
  ctx.fillRect(0, 0, HDR_W - 1, HDR_H - 1)
  ctx.fillStyle = colors.border
  ctx.beginPath()
  ctx.moveTo(HDR_W - 6, HDR_H - 6)
  ctx.lineTo(HDR_W - 6, HDR_H - 16)
  ctx.lineTo(HDR_W - 16, HDR_H - 6)
  ctx.fill()
  ctx.textAlign = 'left'
  void FH
}
function drawOverlays() {
  const sh = sheet()
  const g = norm()
  const b1 = rect(g.r1, g.c1, false)
  const rr = g.r2 >= MAXR - 1 ? viewH : null
  const b2 = rect(Math.min(g.r2, rowAx.n - 1), Math.min(g.c2, colAx.n - 1), false)
  const x = b1.x
  const y = b1.y
  const w = (g.c2 >= MAXC - 1 ? viewW + scroller.scrollLeft : b2.x + b2.w) - x
  const h2 = (rr ? viewH + scroller.scrollTop : b2.y + b2.h) - y
  const single = g.r1 === g.r2 && g.c1 === g.c2
  const act = rect(A.sel.ar, A.sel.ac, true)
  // formula references while typing
  if (A.edit?.refs) {
    for (const ref of A.edit.refs) {
      if (ref.sheet != null && ref.sheet.toLowerCase() !== sh.name.toLowerCase()) continue
      const a = rect(ref.r1, ref.c1, false)
      const b = rect(Math.min(ref.r2, rowAx.n - 1), Math.min(ref.c2, colAx.n - 1), false)
      ctx.strokeStyle = ref.color
      ctx.fillStyle = `${ref.color}22`
      ctx.lineWidth = 2
      ctx.fillRect(a.x, a.y, b.x + b.w - a.x, b.y + b.h - a.y)
      ctx.strokeRect(a.x + 1, a.y + 1, b.x + b.w - a.x - 2, b.y + b.h - a.y - 2)
    }
    ctx.lineWidth = 1
  }
  if (A.findHits.length) {
    ctx.fillStyle = 'rgba(250, 204, 21, 0.35)'
    for (const [r, c] of A.findHits) {
      const b = rect(r, c)
      ctx.fillRect(b.x, b.y, b.w, b.h)
    }
  }
  if (!single) {
    ctx.fillStyle = colors.accentSoft
    ctx.fillRect(x, y, w, h2)
    // keep the active cell clear
    ctx.fillStyle = colors.bg
    ctx.save()
    ctx.beginPath()
    ctx.rect(act.x, act.y, act.w, act.h)
    ctx.clip()
    drawCell(A.sel.ar, A.sel.ac, act.x, act.y, act.w, act.h, viewW)
    ctx.restore()
  }
  ctx.strokeStyle = colors.accent
  ctx.lineWidth = 2
  ctx.strokeRect(x + 1, y + 1, w - 2, h2 - 2)
  if (!single) ctx.strokeRect(act.x + 1, act.y + 1, act.w - 2, act.h - 2)
  ctx.lineWidth = 1
  if (g.r2 < MAXR - 1 && g.c2 < MAXC - 1 && !A.edit) {
    ctx.fillStyle = colors.accent
    ctx.fillRect(x + w - 5, y + h2 - 5, 7, 7)
    ctx.strokeStyle = colors.bg
    ctx.strokeRect(x + w - 5.5, y + h2 - 5.5, 8, 8)
  }
  if (A.clip && A.clip.sheet === sh.id) {
    const c = A.clip.range
    const a = rect(c.r1, c.c1, false)
    const b = rect(c.r2, c.c2, false)
    ctx.strokeStyle = colors.accent
    ctx.setLineDash([5, 4])
    ctx.lineWidth = 1.5
    ctx.strokeRect(a.x, a.y, b.x + b.w - a.x, b.y + b.h - a.y)
    ctx.setLineDash([])
    ctx.lineWidth = 1
  }
  if (A.drag?.kind === 'fill' && A.drag.target) {
    const t = A.drag.target
    const a = rect(t.r1, t.c1, false)
    const b = rect(t.r2, t.c2, false)
    ctx.strokeStyle = colors.accent
    ctx.setLineDash([4, 3])
    ctx.strokeRect(a.x, a.y, b.x + b.w - a.x, b.y + b.h - a.y)
    ctx.setLineDash([])
  }
  // filter buttons on the header row of an auto filter, and the list arrow of a validated cell
  const f = sh.filter
  if (f?.range) {
    const fg = XL.parseRange(f.range)
    if (fg) {
      for (let c = fg.c1; c <= fg.c2; c++) {
        const b = rect(fg.r1, c)
        if (b.x + b.w < HDR_W || b.x > viewW || !b.w) continue
        drawFilterButton(b.x + b.w - 18, b.y + (b.h - 15) / 2, !!f.cols?.[c - fg.c1]?.values)
      }
    }
  }
  const dv = validationAt(A.sel.ar, A.sel.ac)
  if (dv?.type === 'list' && !A.edit) {
    const b = rect(A.sel.ar, A.sel.ac)
    ctx.fillStyle = colors.head
    ctx.fillRect(b.x + b.w - 17, b.y + 2, 15, b.h - 4)
    ctx.strokeStyle = colors.border
    ctx.strokeRect(b.x + b.w - 17.5, b.y + 1.5, 15, b.h - 3)
    ctx.fillStyle = colors.headText
    ctx.beginPath()
    ctx.moveTo(b.x + b.w - 13, b.y + b.h / 2 - 2)
    ctx.lineTo(b.x + b.w - 5, b.y + b.h / 2 - 2)
    ctx.lineTo(b.x + b.w - 9, b.y + b.h / 2 + 3)
    ctx.fill()
  }
}
function drawFilterButton(x, y, active) {
  ctx.fillStyle = active ? colors.accent : colors.head
  ctx.fillRect(x, y, 15, 15)
  ctx.strokeStyle = active ? colors.accent : colors.border
  ctx.strokeRect(x + 0.5, y + 0.5, 14, 14)
  ctx.fillStyle = active ? '#fff' : colors.headText
  ctx.beginPath()
  ctx.moveTo(x + 4, y + 5)
  ctx.lineTo(x + 11, y + 5)
  ctx.lineTo(x + 7.5, y + 10)
  ctx.fill()
}
const validationAt = (r, c) => {
  for (const rule of sheet().dv) {
    const g = XL.parseRange(rule.range)
    if (g && XL.inRange(g, r, c)) return rule
  }
  return null
}

// ── selection and scrolling ───────────────────────────────────────────────────────────────────────────────────────────────
function expandMerges(g) {
  let ch = true
  const out = { ...g }
  while (ch) {
    ch = false
    for (const m of merges) {
      if (m.r2 < out.r1 || m.r1 > out.r2 || m.c2 < out.c1 || m.c1 > out.c2) continue
      if (m.r1 < out.r1) ((out.r1 = m.r1), (ch = true))
      if (m.c1 < out.c1) ((out.c1 = m.c1), (ch = true))
      if (m.r2 > out.r2) ((out.r2 = m.r2), (ch = true))
      if (m.c2 > out.c2) ((out.c2 = m.c2), (ch = true))
    }
  }
  return out
}
function setSel(ar, ac, fr = ar, fc = ac, { scroll = true, keepEdit = false } = {}) {
  if (A.edit && !keepEdit) commitEdit()
  ar = clamp(ar, 0, MAXR - 1)
  ac = clamp(ac, 0, MAXC - 1)
  fr = clamp(fr, 0, MAXR - 1)
  fc = clamp(fc, 0, MAXC - 1)
  if (fr >= rowAx.n - 1 && fr < MAXR - 1) rowAx.n = Math.min(MAXR, fr + 200)
  if (fc >= colAx.n - 1 && fc < MAXC - 1) colAx.n = Math.min(MAXC, fc + 30)
  A.sel = { ar, ac, fr, fc }
  if (merges.length && (ar !== fr || ac !== fc)) {
    const e = expandMerges({ r1: Math.min(ar, fr), c1: Math.min(ac, fc), r2: Math.max(ar, fr), c2: Math.max(ac, fc) })
    A.sel = { ar: ar <= fr ? e.r1 : e.r2, ac: ac <= fc ? e.c1 : e.c2, fr: ar <= fr ? e.r2 : e.r1, fc: ac <= fc ? e.c2 : e.c1 }
  }
  const m = mergeOf(A.sel.ar, A.sel.ac)
  if (m && A.sel.ar === A.sel.fr && A.sel.ac === A.sel.fc) A.sel = { ar: m.r1, ac: m.c1, fr: m.r2, fc: m.c2 }
  if (scroll) scrollTo(A.sel.fr, A.sel.fc)
  view().sel = { ...A.sel }
  schedule()
  App.onSelect?.()
}
const view = () => (A.views[sheet().id] ||= { sel: { ar: 0, ac: 0, fr: 0, fc: 0 }, sx: 0, sy: 0 })
function scrollTo(r, c) {
  const sh = sheet()
  const b = rect(r, c, true)
  const FW = colAx.pos(sh.freeze.cols)
  const FH = rowAx.pos(sh.freeze.rows)
  if (c >= sh.freeze.cols) {
    if (b.x < HDR_W + FW) scroller.scrollLeft -= HDR_W + FW - b.x
    else if (b.x + b.w > viewW) scroller.scrollLeft += b.x + b.w - viewW + 4
  }
  if (r >= sh.freeze.rows) {
    if (b.y < HDR_H + FH) scroller.scrollTop -= HDR_H + FH - b.y
    else if (b.y + b.h > viewH) scroller.scrollTop += b.y + b.h - viewH + 4
  }
}
// the next cell with data in a direction, like Ctrl+Arrow
function edgeJump(r, c, dr, dc) {
  const book = A.book
  const used = book.usedRange(sheet().id)
  const has = (rr, cc) => {
    const v = book.get(A.si, rr, cc)
    return v !== null && v !== undefined && v !== ''
  }
  const maxR = used ? used.r2 : 0
  const maxC = used ? used.c2 : 0
  let rr = r
  let cc = c
  const inb = (a, b) => a >= 0 && b >= 0 && a < MAXR && b < MAXC
  const next = () => [rr + dr, cc + dc]
  let [nr, nc] = next()
  if (!inb(nr, nc)) return { r, c }
  if (has(rr, cc) && has(nr, nc)) {
    while (inb(nr, nc) && has(nr, nc) && (dr ? nr <= maxR : nc <= maxC)) {
      rr = nr
      cc = nc
      ;[nr, nc] = next()
    }
    return { r: rr, c: cc }
  }
  rr = nr
  cc = nc
  while (inb(rr + dr, cc + dc) && !has(rr, cc) && (dr ? rr <= maxR : cc <= maxC)) {
    rr += dr
    cc += dc
  }
  if (!has(rr, cc)) {
    return dr > 0 ? { r: Math.max(rr, Math.min(maxR + 0, MAXR - 1)) > maxR ? Math.max(r, maxR) : rr, c } : dc > 0 ? { r, c: cc > maxC ? Math.max(c, maxC) : cc } : { r: dr < 0 ? 0 : rr, c: dc < 0 ? 0 : cc }
  }
  return { r: rr, c: cc }
}

// ── applying changes: every edit is an operation, kept for undo and sent to the server ────────────────────────────────────
function applyOps(ops, { undoable = true, label = '' } = {}) {
  if (!ops.length) return
  const book = A.book
  let inv
  try {
    inv = book.applyAll(ops)
  } catch (e) {
    toast(String(e.message || e), 'err')
    return false
  }
  if (undoable) {
    A.undo.push({ fwd: ops, inv, label, sel: { ...A.sel }, si: A.si })
    if (A.undo.length > 200) A.undo.shift()
    A.redo = []
  }
  A.pending.push(...ops)
  afterChange()
  Save.soon()
  return true
}
function afterChange() {
  const book = A.book
  if (A.si >= book.sheets.length) A.si = book.sheets.length - 1
  book.dirty = true
  buildAxes()
  schedule()
  App.onChange?.()
}
function undo() {
  const e = A.undo.pop()
  if (!e) return
  A.book.applyAll(e.inv)
  A.redo.push(e)
  A.pending.push(...e.inv)
  A.si = clamp(e.si, 0, A.book.sheets.length - 1)
  afterChange()
  setSel(e.sel.ar, e.sel.ac, e.sel.fr, e.sel.fc)
  Save.soon()
}
function redo() {
  const e = A.redo.pop()
  if (!e) return
  A.book.applyAll(e.fwd)
  A.undo.push(e)
  A.pending.push(...e.fwd)
  afterChange()
  Save.soon()
}
const cellJSON = (c) => {
  if (!c) return null
  const o = {}
  if (c.f != null) o.f = c.f
  else if (c.v !== undefined) o.v = c.v
  if (c.s) o.s = { ...c.s }
  return o
}
// set the cells of a range to a style patch ({b:true}, or null values to remove), as one change
function styleRange(g, patch) {
  const sh = sheet()
  const used = A.book.usedRange(sh.id)
  const r2 = g.r2 >= MAXR - 1 ? Math.max(used?.r2 ?? 0, 100) : g.r2
  const c2 = g.c2 >= MAXC - 1 ? Math.max(used?.c2 ?? 0, 26) : g.c2
  if ((r2 - g.r1 + 1) * (c2 - g.c1 + 1) > 60000) return toast('That is too many cells to format at once.', 'err')
  const cells = {}
  for (let r = g.r1; r <= r2; r++)
    for (let c = g.c1; c <= c2; c++) {
      const old = sh.cells.get(key(r, c))
      const s = { ...(old?.s || {}) }
      for (const [k, v] of Object.entries(patch)) {
        if (v === null || v === undefined || v === false || v === '') delete s[k]
        else s[k] = v
      }
      const base = old ? cellJSON(old) : {}
      delete base.s
      cells[A1(r, c)] = Object.keys(s).length ? { ...base, s } : Object.keys(base).length ? base : null
    }
  return applyOps([{ op: 'set', sheet: sh.id, cells }], { label: 'Format' })
}
function clearRange(g, what = 'contents') {
  const sh = sheet()
  const cells = {}
  for (const [k, c] of sh.cells) {
    const r = Math.floor(k / MAXC)
    const col = k % MAXC
    if (r < g.r1 || r > g.r2 || col < g.c1 || col > g.c2) continue
    cells[A1(r, col)] = what === 'all' ? null : what === 'formats' ? (c.f != null ? { f: c.f } : { v: c.v }) : c.s ? { s: c.s } : null
  }
  if (Object.keys(cells).length) applyOps([{ op: 'set', sheet: sh.id, cells }], { label: 'Clear' })
}
// what a typed text becomes
function opForText(r, c, text) {
  const sh = sheet()
  const old = sh.cells.get(key(r, c))
  const p = XL.parseInput(text)
  const s = { ...(old?.s || {}) }
  if (p.nf && !s.nf) s.nf = p.nf
  const cell = p.cell ? { ...p.cell, ...(Object.keys(s).length ? { s } : {}) } : Object.keys(s).length ? { s } : null
  return { op: 'set', sheet: sh.id, cells: { [A1(r, c)]: cell } }
}

// ── editing ───────────────────────────────────────────────────────────────────────────────────────────────────────────────
const REF_COLORS = ['#2563eb', '#dc2626', '#059669', '#d97706', '#7c3aed', '#0891b2', '#db2777']
function startEdit(initial, mode = 'edit') {
  if (A.edit) return
  const { ar: r, ac: c } = A.sel
  const cell = sheet().cells.get(key(r, c))
  const text = initial !== undefined && initial !== null ? initial : XL.rawText(cell)
  A.edit = { r, c, si: A.si, text, mode, orig: XL.rawText(cell) }
  editor.style.display = 'block'
  editor.value = text
  positionEditor()
  editor.focus()
  if (initial === undefined || initial === null) editor.setSelectionRange(text.length, text.length)
  onEditInput()
  schedule()
}
function onEditInput() {
  const e = A.edit
  if (!e) return
  e.text = editor.value
  e.refs = null
  if (e.text.startsWith('=')) {
    try {
      const toks = XL.tokenize(e.text.slice(1))
      e.refs = toks.filter((t) => t.t === 'ref').map((t, i) => ({ ...t.ref, color: REF_COLORS[i % REF_COLORS.length] }))
    } catch {
      /* half typed */
    }
  }
  sizeEditor()
  App.onEditInput?.(e.text)
  schedule()
}
function sizeEditor() {
  const b = rect(A.edit.r, A.edit.c)
  editor.style.height = 'auto'
  const need = clamp(editor.scrollHeight, b.h, 260)
  editor.style.height = `${need}px`
  const tmp = h('span', { style: { font: getComputedStyle(editor).font, position: 'absolute', visibility: 'hidden', whiteSpace: 'pre' } }, editor.value.split('\n').sort((x, y) => y.length - x.length)[0] || '')
  document.body.append(tmp)
  const tw = tmp.getBoundingClientRect().width + 28
  tmp.remove()
  editor.style.width = `${clamp(tw, b.w, Math.max(b.w, viewW - b.x - 8))}px`
}
function positionEditor() {
  if (!A.edit) return void (editor.style.display = 'none')
  if (A.edit.si !== A.si) return
  const b = rect(A.edit.r, A.edit.c)
  editor.style.left = `${b.x}px`
  editor.style.top = `${b.y}px`
  editor.style.minWidth = `${b.w}px`
  editor.style.minHeight = `${b.h}px`
  editor.style.font = `${13 * Z()}px ${colors.font}`
}
function commitEdit(move = null) {
  const e = A.edit
  if (!e) return true
  const text = editor.value
  const sh = A.book.sheets[e.si]
  if (e.si !== A.si) return cancelEdit()
  const same = text === e.orig
  if (!same) {
    const parsed = XL.parseInput(text)
    const val = parsed.cell ? (parsed.cell.f != null ? null : parsed.cell.v) : null
    if (parsed.cell && parsed.cell.f == null) {
      const check = XL.validate(A.book, e.si, e.r, e.c, text.trim() === '' ? null : text)
      if (!check.ok) {
        toast(check.message || 'That value is not allowed here.', 'err')
        editor.focus()
        return false
      }
    }
    void val
    A.edit = null
    editor.style.display = 'none'
    applyOps([opForText(e.r, e.c, text)], { label: 'Edit cell' })
  } else {
    A.edit = null
    editor.style.display = 'none'
  }
  void sh
  App.onEditInput?.(null)
  scroller.focus({ preventScroll: true })
  if (move) moveActive(move[0], move[1])
  schedule()
  return true
}
function cancelEdit() {
  if (!A.edit) return true
  A.edit = null
  editor.style.display = 'none'
  App.onEditInput?.(null)
  scroller.focus({ preventScroll: true })
  schedule()
  return true
}
// Enter / Tab move inside a multi-cell selection, otherwise one cell
function moveActive(dr, dc, extend = false) {
  const g = norm()
  const multi = g.r1 !== g.r2 || g.c1 !== g.c2
  let { ar, ac } = A.sel
  if (extend) return setSel(A.sel.ar, A.sel.ac, A.sel.fr + dr, A.sel.fc + dc)
  if (multi) {
    ar += dr
    ac += dc
    if (ar > g.r2) ((ar = g.r1), (ac = ac + 1))
    if (ar < g.r1) ((ar = g.r2), (ac = ac - 1))
    if (ac > g.c2) ((ac = g.c1), (ar = ar + 1))
    if (ac < g.c1) ((ac = g.c2), (ar = ar - 1))
    if (ar < g.r1 || ar > g.r2 || ac < g.c1 || ac > g.c2) {
      ar = g.r1
      ac = g.c1
    }
    A.sel = { ...A.sel, ar, ac }
    scrollTo(ar, ac)
    view().sel = { ...A.sel }
    schedule()
    App.onSelect?.()
    return
  }
  // skip hidden rows and columns
  let r = ar + dr
  let c = ac + dc
  while (r >= 0 && r < MAXR && rowAx.size(r) === 0) r += dr || 1
  while (c >= 0 && c < MAXC && colAx.size(c) === 0) c += dc || 1
  const m = mergeOf(ar, ac)
  if (m && dr > 0) r = m.r2 + 1
  if (m && dc > 0) c = m.c2 + 1
  setSel(r, c)
}

// ── clipboard ─────────────────────────────────────────────────────────────────────────────────────────────────────────────
function gatherCells(g) {
  const sh = sheet()
  const book = A.book
  const used = book.usedRange(sh.id)
  const r2 = g.r2 >= MAXR - 1 ? used?.r2 ?? 0 : g.r2
  const c2 = g.c2 >= MAXC - 1 ? used?.c2 ?? 0 : g.c2
  const cells = {}
  const rows = []
  for (let r = g.r1; r <= r2; r++) {
    const row = []
    for (let c = g.c1; c <= c2; c++) {
      const cell = sh.cells.get(key(r, c))
      cells[A1(r, c)] = cell ? { ...cellJSON(cell), _val: book.get(A.si, r, c) } : null
      row.push(book.display(A.si, r, c))
    }
    rows.push(row)
  }
  return { cells, rows, range: { r1: g.r1, c1: g.c1, r2, c2 } }
}
const tsv = (rows) => rows.map((r) => r.map((v) => (/[\t\n"]/.test(v) ? `"${v.replace(/"/g, '""')}"` : v)).join('\t')).join('\n')
function parseDelimited(text, d = '\t') {
  const rows = []
  let row = []
  let cell = ''
  let q = false
  for (let i = 0; i < text.length; i++) {
    const ch = text[i]
    if (q) {
      if (ch === '"') {
        if (text[i + 1] === '"') {
          cell += '"'
          i++
        } else q = false
      } else cell += ch
    } else if (ch === '"' && cell === '') q = true
    else if (ch === d) {
      row.push(cell)
      cell = ''
    } else if (ch === '\n' || ch === '\r') {
      if (ch === '\r' && text[i + 1] === '\n') i++
      row.push(cell)
      rows.push(row)
      row = []
      cell = ''
    } else cell += ch
  }
  if (cell !== '' || row.length) {
    row.push(cell)
    rows.push(row)
  }
  if (rows.length && rows[rows.length - 1].every((x) => x === '') && rows.length > 1) rows.pop()
  return rows
}
function doCopy(e, cut = false) {
  const g = norm()
  const got = gatherCells(g)
  const text = tsv(got.rows)
  A.clip = { sheet: sheet().id, range: got.range, cells: got.cells, top: { r: got.range.r1, c: got.range.c1 }, text, cut }
  if (e?.clipboardData) {
    e.clipboardData.setData('text/plain', text)
    e.preventDefault()
  } else copyText(text)
  schedule()
}
function doPaste(text, mode = 'all') {
  const sh = sheet()
  const g = norm()
  const top = { r: g.r1, c: g.c1 }
  const clip = A.clip
  let ops = []
  if (clip && (text === undefined || text === clip.text)) {
    const src = A.book.sheets.find((s) => s.id === clip.sheet)
    const srcCells = {}
    for (const [a, cell] of Object.entries(clip.cells)) srcCells[a] = cell
    const plan = A.book.planPaste(sh.id, srcCells, clip.top, top, { values: mode === 'values', formats: mode === 'formats' })
    ops.push(plan)
    if (clip.cut && src && mode === 'all') {
      const dst = { r1: top.r, c1: top.c, r2: top.r + (clip.range.r2 - clip.range.r1), c2: top.c + (clip.range.c2 - clip.range.c1) }
      const clear = {}
      for (let r = clip.range.r1; r <= clip.range.r2; r++)
        for (let c = clip.range.c1; c <= clip.range.c2; c++) {
          if (src === sh && XL.inRange(dst, r, c)) continue
          if (src.cells.has(key(r, c))) clear[A1(r, c)] = null
        }
      if (Object.keys(clear).length) ops = [{ op: 'set', sheet: src.id, cells: clear }, ...ops]
      A.clip = null
    }
    // pasted formulas that pointed into the cut range follow it: not tracked, as in most spreadsheets' simple mode
  } else if (text !== undefined) {
    const rows = parseDelimited(text, text.includes('\t') || !text.includes(',') ? '\t' : ',')
    const cells = {}
    rows.forEach((row, i) =>
      row.forEach((v, j) => {
        const p = XL.parseInput(v)
        const old = sh.cells.get(key(top.r + i, top.c + j))
        const s = { ...(old?.s || {}) }
        if (p.nf && !s.nf) s.nf = p.nf
        cells[A1(top.r + i, top.c + j)] = p.cell ? { ...p.cell, ...(Object.keys(s).length ? { s } : {}) } : Object.keys(s).length ? { s } : null
      }),
    )
    ops.push({ op: 'set', sheet: sh.id, cells })
    if (rows.length) setSelAfter(top.r + rows.length - 1, top.c + Math.max(...rows.map((r) => r.length)) - 1, top)
  }
  if (ops.length) applyOps(ops, { label: 'Paste' })
  schedule()
}
const setSelAfter = (r, c, top) => setSel(top.r, top.c, r, c)

// ── fill handle ───────────────────────────────────────────────────────────────────────────────────────────────────────────
function doFill(src, dst, series = true) {
  const plan = A.book.planFill(sheet().id, src, dst, { series })
  applyOps([plan], { label: 'Fill' })
  setSel(Math.min(src.r1, dst.r1), Math.min(src.c1, dst.c1), Math.max(src.r2, dst.r2), Math.max(src.c2, dst.c2))
}

// ── mouse ─────────────────────────────────────────────────────────────────────────────────────────────────────────────────
function localXY(e) {
  const b = scroller.getBoundingClientRect()
  return { x: e.clientX - b.left, y: e.clientY - b.top }
}
function headerEdge(x, y) {
  const sh = sheet()
  if (y < HDR_H && x >= HDR_W) {
    const { c } = hitCell(x, HDR_H + 1)
    for (const cc of [c, c - 1]) {
      if (cc < 0) continue
      const b = rect(0, cc, false)
      if (Math.abs(x - (b.x + b.w)) <= 3 && colAx.size(cc) > 0) return { kind: 'col', i: cc, start: b.x + b.w }
    }
  }
  if (x < HDR_W && y >= HDR_H) {
    const { r } = hitCell(HDR_W + 1, y)
    for (const rr of [r, r - 1]) {
      if (rr < 0) continue
      const b = rect(rr, 0, false)
      if (Math.abs(y - (b.y + b.h)) <= 3 && rowAx.size(rr) > 0) return { kind: 'row', i: rr, start: b.y + b.h }
    }
  }
  void sh
  return null
}
function onMouseDown(e) {
  if (e.button === 2) return
  closeMenu()
  const { x, y } = localXY(e)
  if (x > scroller.clientWidth || y > scroller.clientHeight) return // the scrollbars
  if (A.edit) {
    // pointing at a cell while typing a formula puts its address in
    if (pointingAllowed() && x >= HDR_W && y >= HDR_H) {
      e.preventDefault()
      const { r, c } = hitCell(x, y)
      beginPoint(r, c)
      A.drag = { kind: 'point', r, c }
      listenDrag()
      return
    }
    if (!commitEdit()) return
  }
  scroller.focus({ preventScroll: true })
  const edge = headerEdge(x, y)
  if (edge) {
    A.drag = { kind: 'resize', axis: edge.kind, i: edge.i, start: edge.kind === 'col' ? x : y, size: edge.kind === 'col' ? colAx.size(edge.i) / Z() : rowAx.size(edge.i) / Z() }
    listenDrag()
    return
  }
  if (x < HDR_W && y < HDR_H) return setSel(0, 0, MAXR - 1, MAXC - 1)
  if (y < HDR_H) {
    const { c } = hitCell(x, HDR_H + 1)
    const f = sheet().filter
    if (e.shiftKey) setSel(0, A.sel.ac, MAXR - 1, c)
    else setSel(0, c, MAXR - 1, c)
    A.drag = { kind: 'cols', c }
    listenDrag()
    void f
    return
  }
  if (x < HDR_W) {
    const { r } = hitCell(HDR_W + 1, y)
    if (e.shiftKey) setSel(A.sel.ar, 0, r, MAXC - 1)
    else setSel(r, 0, r, MAXC - 1)
    A.drag = { kind: 'rows', r }
    listenDrag()
    return
  }
  // fill handle
  const g = norm()
  if (g.r2 < MAXR - 1 && g.c2 < MAXC - 1) {
    const b2 = rect(g.r2, g.c2, false)
    const hx = b2.x + b2.w
    const hy = b2.y + b2.h
    if (Math.abs(x - hx) <= 6 && Math.abs(y - hy) <= 6) {
      A.drag = { kind: 'fill', src: g, target: null }
      listenDrag()
      return
    }
  }
  const { r, c } = hitCell(x, y)
  // the arrow of a validated list, the buttons of an auto filter
  const f = sheet().filter
  if (f?.range) {
    const fg = XL.parseRange(f.range)
    if (fg && r === fg.r1 && c >= fg.c1 && c <= fg.c2) {
      const b = rect(r, c)
      if (x >= b.x + b.w - 20) return App.openFilter?.(c - fg.c1, b)
    }
  }
  const dv = validationAt(A.sel.ar, A.sel.ac)
  if (dv?.type === 'list' && r === A.sel.ar && c === A.sel.ac) {
    const b = rect(r, c)
    if (x >= b.x + b.w - 18) return App.openListDropdown?.(dv, b)
  }
  if (e.shiftKey) setSel(A.sel.ar, A.sel.ac, r, c)
  else setSel(r, c)
  A.drag = { kind: 'cells' }
  listenDrag()
}
function listenDrag() {
  const move = (e) => onDragMove(e)
  const up = (e) => {
    removeEventListener('mousemove', move)
    removeEventListener('mouseup', up)
    clearInterval(A.drag?.timer)
    onDragEnd(e)
    A.drag = null
  }
  addEventListener('mousemove', move)
  addEventListener('mouseup', up)
}
function onDragMove(e) {
  const d = A.drag
  if (!d) return
  const { x, y } = localXY(e)
  // scroll when the pointer leaves the grid
  clearInterval(d.timer)
  if (d.kind !== 'resize') {
    const dx = x < HDR_W ? -30 : x > viewW ? 30 : 0
    const dy = y < HDR_H ? -30 : y > viewH ? 30 : 0
    if (dx || dy) d.timer = setInterval(() => ((scroller.scrollLeft += dx), (scroller.scrollTop += dy), onDragMove(e)), 40)
  }
  if (d.kind === 'resize') {
    const size = Math.max(8, d.size + ((d.axis === 'col' ? x : y) - d.start) / Z())
    d.live = size
    const sh = sheet()
    if (d.axis === 'col') sh.cols[XL.colName(d.i)] = Math.round(size)
    else sh.rows[String(d.i + 1)] = Math.round(size)
    buildAxes()
    schedule()
    return
  }
  const { r, c } = hitCell(clamp(x, HDR_W, viewW - 1), clamp(y, HDR_H, viewH - 1))
  if (d.kind === 'cells') setSel(A.sel.ar, A.sel.ac, r, c, { scroll: false })
  else if (d.kind === 'cols') setSel(0, d.c, MAXR - 1, c, { scroll: false })
  else if (d.kind === 'rows') setSel(d.r, 0, r, MAXC - 1, { scroll: false })
  else if (d.kind === 'point') pointExtend(d.r, d.c, r, c)
  else if (d.kind === 'fill') {
    const s = d.src
    const dv = r > s.r2 ? r - s.r2 : r < s.r1 ? s.r1 - r : 0
    const dh = c > s.c2 ? c - s.c2 : c < s.c1 ? s.c1 - c : 0
    if (!dv && !dh) d.target = null
    else if (dv >= dh) d.target = r > s.r2 ? { r1: s.r2 + 1, c1: s.c1, r2: r, c2: s.c2 } : { r1: r, c1: s.c1, r2: s.r1 - 1, c2: s.c2 }
    else d.target = c > s.c2 ? { r1: s.r1, c1: s.c2 + 1, r2: s.r2, c2: c } : { r1: s.r1, c1: c, r2: s.r2, c2: s.c1 - 1 }
    schedule()
  }
}
function onDragEnd() {
  const d = A.drag
  if (!d) return
  if (d.kind === 'resize') {
    const sh = sheet()
    const key2 = d.axis === 'col' ? 'cols' : 'rows'
    const next = { ...sh[key2] }
    const k = d.axis === 'col' ? XL.colName(d.i) : String(d.i + 1)
    const val = next[k]
    // the live edits already changed the sheet; record one operation with the real before value
    const before = { ...sh[key2] }
    if (d.size * 1 === val) return
    delete before[k]
    if (Math.abs(d.size - (d.axis === 'col' ? DEF_W : DEF_H)) > 0.5) before[k] = Math.round(d.size)
    A.undo.push({ fwd: [{ op: 'prop', sheet: sh.id, key: key2, value: next }], inv: [{ op: 'prop', sheet: sh.id, key: key2, value: before }], label: 'Resize', sel: { ...A.sel }, si: A.si })
    A.redo = []
    A.pending.push({ op: 'prop', sheet: sh.id, key: key2, value: next })
    Save.soon()
    return
  }
  if (d.kind === 'fill' && d.target) {
    doFill(d.src, d.target)
    return
  }
  if (d.kind === 'cells') App.onSelect?.()
}
function onDblClick(e) {
  const { x, y } = localXY(e)
  const edge = headerEdge(x, y)
  if (edge) return autofit(edge.kind, edge.i)
  if (x < HDR_W || y < HDR_H) return
  // the fill handle: fill down as far as the column beside it goes
  const g = norm()
  const b2 = rect(g.r2, g.c2, false)
  if (Math.abs(x - (b2.x + b2.w)) <= 6 && Math.abs(y - (b2.y + b2.h)) <= 6) {
    const book = A.book
    let last = g.r2
    const colNext = g.c1 > 0 ? g.c1 - 1 : g.c2 + 1
    for (let r = g.r2 + 1; r < MAXR; r++) {
      const v = book.get(A.si, r, colNext)
      if (v === null || v === undefined || v === '') break
      last = r
    }
    if (last > g.r2) doFill(g, { r1: g.r2 + 1, c1: g.c1, r2: last, c2: g.c2 })
    return
  }
  const { r, c } = hitCell(x, y)
  if (r === A.sel.ar && c === A.sel.ac) startEdit()
}
function autofit(axis, i) {
  const sh = sheet()
  const book = A.book
  ctx.font = `${13 * Z()}px ${colors.font}`
  const used = book.usedRange(sh.id)
  if (!used) return
  let best = 0
  if (axis === 'col') {
    for (let r = 0; r <= used.r2; r++) {
      const t = textOf(r, i)
      if (t) best = Math.max(best, ctx.measureText(t.split('\n')[0]).width)
    }
    const w = clamp(Math.round(best / Z() + 16), 30, 600)
    applyOps([{ op: 'prop', sheet: sh.id, key: 'cols', value: { ...sh.cols, [XL.colName(i)]: w } }], { label: 'Autofit' })
  } else {
    const lines = Math.max(1, ...Array.from({ length: used.c2 + 1 }, (_, c) => (sh.cells.get(key(i, c))?.s?.wrap ? wrapLines(textOf(i, c), colAx.size(c) - 10, ctx).length : 1)))
    applyOps([{ op: 'prop', sheet: sh.id, key: 'rows', value: { ...sh.rows, [String(i + 1)]: Math.round(lines * 18 + 6) } }], { label: 'Autofit' })
  }
}

// ── pointing: clicking cells while typing a formula ───────────────────────────────────────────────────────────────────────
function pointingAllowed() {
  const e = A.edit
  if (!e || !e.text.startsWith('=')) return false
  const pos = editor.selectionStart
  const before = e.text.slice(0, pos).replace(/\s+$/, '')
  return /[=+\-*/^&,(<>:;]$/.test(before) || e.point !== undefined && e.point !== null && pos === e.point.end
}
function beginPoint(r, c) {
  const e = A.edit
  const pos = e.point && editor.selectionStart === e.point.end ? e.point.start : editor.selectionStart
  const end = e.point && editor.selectionStart === e.point.end ? e.point.end : editor.selectionStart
  const text = `${A1(r, c)}`
  const v = editor.value
  editor.value = v.slice(0, pos) + text + v.slice(end)
  e.point = { start: pos, end: pos + text.length }
  editor.setSelectionRange(e.point.end, e.point.end)
  onEditInput()
}
function pointExtend(r0, c0, r, c) {
  const e = A.edit
  if (!e?.point) return
  const text = r0 === r && c0 === c ? A1(r, c) : `${A1(Math.min(r0, r), Math.min(c0, c))}:${A1(Math.max(r0, r), Math.max(c0, c))}`
  const v = editor.value
  editor.value = v.slice(0, e.point.start) + text + v.slice(e.point.end)
  e.point.end = e.point.start + text.length
  editor.setSelectionRange(e.point.end, e.point.end)
  onEditInput()
}

// ── keyboard ──────────────────────────────────────────────────────────────────────────────────────────────────────────────
function onKeyDown(e) {
  if (A.mode !== 'sheet' || e.target === editor || (e.target.closest && e.target.closest('.dlg, .formula-input, input, select, textarea'))) return
  const ctrl = e.ctrlKey || e.metaKey
  const k = e.key
  const g = norm()
  const nav = (dr, dc) => {
    e.preventDefault()
    if (ctrl) {
      const from = e.shiftKey ? { r: A.sel.fr, c: A.sel.fc } : { r: A.sel.ar, c: A.sel.ac }
      const t = edgeJump(from.r, from.c, dr, dc)
      return e.shiftKey ? setSel(A.sel.ar, A.sel.ac, t.r, t.c) : setSel(t.r, t.c)
    }
    if (e.shiftKey) {
      let r = A.sel.fr + dr
      let c = A.sel.fc + dc
      while (r >= 0 && rowAx.size(r) === 0 && r < MAXR) r += dr || 1
      while (c >= 0 && colAx.size(c) === 0 && c < MAXC) c += dc || 1
      return setSel(A.sel.ar, A.sel.ac, r, c)
    }
    const m = mergeOf(A.sel.ar, A.sel.ac)
    let r = (dr > 0 && m ? m.r2 : dr < 0 && m ? m.r1 : A.sel.ar) + dr
    let c = (dc > 0 && m ? m.c2 : dc < 0 && m ? m.c1 : A.sel.ac) + dc
    while (r >= 0 && r < MAXR && rowAx.size(r) === 0) r += dr || 1
    while (c >= 0 && c < MAXC && colAx.size(c) === 0) c += dc || 1
    setSel(r, c)
  }
  if (k === 'ArrowDown') return nav(1, 0)
  if (k === 'ArrowUp') return nav(-1, 0)
  if (k === 'ArrowRight') return nav(0, 1)
  if (k === 'ArrowLeft') return nav(0, -1)
  if (k === 'PageDown' || k === 'PageUp') {
    e.preventDefault()
    const rows = Math.max(1, Math.floor((viewH - HDR_H) / rowAx.def) - 1) * (k === 'PageDown' ? 1 : -1)
    return e.shiftKey ? setSel(A.sel.ar, A.sel.ac, A.sel.fr + rows, A.sel.fc) : setSel(A.sel.ar + rows, A.sel.ac)
  }
  if (k === 'Home') {
    e.preventDefault()
    return ctrl ? setSel(0, 0) : setSel(A.sel.ar, 0)
  }
  if (k === 'End' && ctrl) {
    e.preventDefault()
    const u = A.book.usedRange(sheet().id)
    return u ? setSel(u.r2, u.c2) : setSel(0, 0)
  }
  if (k === 'Enter') {
    e.preventDefault()
    return e.altKey ? startEdit() : moveActive(e.shiftKey ? -1 : 1, 0)
  }
  if (k === 'Tab') {
    e.preventDefault()
    return moveActive(0, e.shiftKey ? -1 : 1)
  }
  if (k === 'F2') {
    e.preventDefault()
    return startEdit()
  }
  if (k === 'Delete' || k === 'Backspace') {
    e.preventDefault()
    return clearRange(g)
  }
  if (k === 'Escape') {
    if (A.clip) {
      A.clip = null
      schedule()
    }
    A.findHits = []
    return
  }
  if (ctrl) {
    const lk = k.toLowerCase()
    if (lk === 'z') return (e.preventDefault(), e.shiftKey ? redo() : undo())
    if (lk === 'y') return (e.preventDefault(), redo())
    if (lk === 'b') return (e.preventDefault(), App.toggleStyle('b'))
    if (lk === 'i') return (e.preventDefault(), App.toggleStyle('i'))
    if (lk === 'u') return (e.preventDefault(), App.toggleStyle('u'))
    if (lk === 'a') return (e.preventDefault(), setSel(0, 0, MAXR - 1, MAXC - 1))
    if (lk === 'd') return (e.preventDefault(), g.r2 > g.r1 ? doFill({ ...g, r2: g.r1 }, { ...g, r1: g.r1 + 1 }, false) : g.r1 > 0 && doFill({ ...g, r1: g.r1 - 1, r2: g.r1 - 1 }, g, false))
    if (lk === 'r') return (e.preventDefault(), g.c2 > g.c1 ? doFill({ ...g, c2: g.c1 }, { ...g, c1: g.c1 + 1 }, false) : g.c1 > 0 && doFill({ ...g, c1: g.c1 - 1, c2: g.c1 - 1 }, g, false))
    if (lk === 'f') return (e.preventDefault(), App.openFind?.(false))
    if (lk === 'h') return (e.preventDefault(), App.openFind?.(true))
    if (lk === '1') return (e.preventDefault(), App.openFormatCells?.())
    if (lk === ' ') return (e.preventDefault(), setSel(0, A.sel.ac, MAXR - 1, A.sel.fc))
    if (lk === ';') return (e.preventDefault(), (startEdit(`${new Date().toISOString().slice(0, 10)}`), commitEdit()))
    if (lk === '`') return (e.preventDefault(), ((A.showFormulas = !A.showFormulas), schedule()))
    if (lk === 'c' || lk === 'x' || lk === 'v') return // the browser raises copy / cut / paste events
    return
  }
  if (e.shiftKey && k === ' ') {
    e.preventDefault()
    return setSel(A.sel.ar, 0, A.sel.fr, MAXC - 1)
  }
  if (k.length === 1 && !e.altKey) {
    e.preventDefault()
    startEdit(k, 'enter')
  }
}

// ── the cell editor's own keys ────────────────────────────────────────────────────────────────────────────────────────────
function onEditorKey(e) {
  const ed = A.edit
  if (!ed) return
  if (App.autocompleteKey?.(e)) return
  const k = e.key
  if (k === 'Escape') {
    e.preventDefault()
    return cancelEdit()
  }
  if (k === 'Enter' && !e.altKey) {
    e.preventDefault()
    return commitEdit([e.shiftKey ? -1 : 1, 0])
  }
  if (k === 'Enter' && e.altKey) {
    e.preventDefault()
    const p = editor.selectionStart
    editor.setRangeText('\n', p, editor.selectionEnd, 'end')
    return onEditInput()
  }
  if (k === 'Tab') {
    e.preventDefault()
    return commitEdit([0, e.shiftKey ? -1 : 1])
  }
  if (k === 'F4') {
    e.preventDefault()
    return cycleAbsolute()
  }
  if (['ArrowDown', 'ArrowUp', 'ArrowLeft', 'ArrowRight'].includes(k) && !e.shiftKey && !e.ctrlKey && !e.altKey) {
    if (pointingAllowed() && ed.text.startsWith('=')) {
      e.preventDefault()
      const dr = k === 'ArrowDown' ? 1 : k === 'ArrowUp' ? -1 : 0
      const dc = k === 'ArrowRight' ? 1 : k === 'ArrowLeft' ? -1 : 0
      const cur = ed.point ? A.pointAt || { r: A.sel.ar, c: A.sel.ac } : { r: A.sel.ar, c: A.sel.ac }
      const r = clamp(cur.r + dr, 0, MAXR - 1)
      const c = clamp(cur.c + dc, 0, MAXC - 1)
      A.pointAt = { r, c }
      beginPoint(r, c)
      return
    }
    if (ed.mode === 'enter' && (k === 'ArrowDown' || k === 'ArrowUp' || !editor.value.includes('\n'))) {
      e.preventDefault()
      return commitEdit([k === 'ArrowDown' ? 1 : k === 'ArrowUp' ? -1 : 0, k === 'ArrowRight' ? 1 : k === 'ArrowLeft' ? -1 : 0])
    }
  }
}
function cycleAbsolute() {
  const v = editor.value
  const pos = editor.selectionStart
  let m
  const re = /(\$?[A-Za-z]{1,3}\$?\d+)(?::(\$?[A-Za-z]{1,3}\$?\d+))?/g
  while ((m = re.exec(v))) {
    if (pos >= m.index && pos <= m.index + m[0].length) {
      const cyc = (s) => {
        const t = s.match(/^(\$?)([A-Za-z]{1,3})(\$?)(\d+)$/)
        const [ac, ar] = [t[1] === '$', t[3] === '$']
        const next = !ac && !ar ? [true, true] : ac && ar ? [false, true] : !ac && ar ? [true, false] : [false, false]
        return `${next[0] ? '$' : ''}${t[2]}${next[1] ? '$' : ''}${t[4]}`
      }
      const out = cyc(m[1]) + (m[2] ? `:${cyc(m[2])}` : '')
      editor.value = v.slice(0, m.index) + out + v.slice(m.index + m[0].length)
      editor.setSelectionRange(m.index + out.length, m.index + out.length)
      return onEditInput()
    }
  }
}

// ── tooltip with the cell's error or value ────────────────────────────────────────────────────────────────────────────────
function updateTip() {}

// ── context menu ──────────────────────────────────────────────────────────────────────────────────────────────────────────
function onContext(e) {
  e.preventDefault()
  const { x, y } = localXY(e)
  const g = norm()
  if (x >= HDR_W && y >= HDR_H) {
    const { r, c } = hitCell(x, y)
    if (!XL.inRange(g, r, c)) setSel(r, c)
  } else if (y < HDR_H && x >= HDR_W) {
    const { c } = hitCell(x, HDR_H + 1)
    if (!(g.c1 <= c && c <= g.c2 && g.r1 === 0 && g.r2 >= MAXR - 1)) setSel(0, c, MAXR - 1, c)
  } else if (x < HDR_W && y >= HDR_H) {
    const { r } = hitCell(HDR_W + 1, y)
    if (!(g.r1 <= r && r <= g.r2 && g.c1 === 0 && g.c2 >= MAXC - 1)) setSel(r, 0, r, MAXC - 1)
  }
  App.contextMenu?.(e.clientX, e.clientY, { onHeaderCol: y < HDR_H && x >= HDR_W, onHeaderRow: x < HDR_W && y >= HDR_H })
}

// ── setup ─────────────────────────────────────────────────────────────────────────────────────────────────────────────────
function gridInit(host) {
  wrap = h('.grid-wrap')
  canvas = h('canvas.grid-canvas')
  ctx = canvas.getContext('2d')
  spacer = h('.grid-spacer')
  scroller = h('.grid-scroller', { tabIndex: 0 }, spacer)
  editor = h('textarea.cell-editor', { spellcheck: false, rows: 1, style: { display: 'none' }, oninput: onEditInput, onkeydown: onEditorKey })
  wrap.append(canvas, scroller, editor)
  host.append(wrap)
  const resize = () => {
    dpr = window.devicePixelRatio || 1
    viewW = scroller.clientWidth + (scroller.offsetWidth - scroller.clientWidth)
    viewH = scroller.clientHeight + (scroller.offsetHeight - scroller.clientHeight)
    const w = wrap.clientWidth
    const hh = wrap.clientHeight
    viewW = w
    viewH = hh
    canvas.width = Math.round(w * dpr)
    canvas.height = Math.round(hh * dpr)
    canvas.style.width = `${w}px`
    canvas.style.height = `${hh}px`
    schedule()
  }
  new ResizeObserver(resize).observe(wrap)
  scroller.addEventListener('scroll', () => {
    const v = view()
    v.sx = scroller.scrollLeft
    v.sy = scroller.scrollTop
    schedule()
  })
  scroller.addEventListener('mousedown', onMouseDown)
  scroller.addEventListener('dblclick', onDblClick)
  scroller.addEventListener('contextmenu', onContext)
  scroller.addEventListener('keydown', onKeyDown)
  scroller.addEventListener('mousemove', (e) => {
    if (A.drag) return
    const { x, y } = localXY(e)
    const edge = headerEdge(x, y)
    scroller.style.cursor = edge ? (edge.kind === 'col' ? 'col-resize' : 'row-resize') : 'default'
    if (!edge && x >= HDR_W && y >= HDR_H) {
      const g = norm()
      if (g.r2 < MAXR - 1 && g.c2 < MAXC - 1) {
        const b2 = rect(g.r2, g.c2, false)
        if (Math.abs(x - (b2.x + b2.w)) <= 6 && Math.abs(y - (b2.y + b2.h)) <= 6) scroller.style.cursor = 'crosshair'
      }
    }
  })
  scroller.addEventListener(
    'wheel',
    (e) => {
      if (e.ctrlKey) {
        e.preventDefault()
        App.setZoom?.(clamp(A.zoom + (e.deltaY < 0 ? 0.1 : -0.1), 0.5, 2))
      }
    },
    { passive: false },
  )
  document.addEventListener('copy', (e) => {
    if (A.mode !== 'sheet' || A.edit || e.target.closest?.('input,textarea,.dlg')) return
    doCopy(e, false)
  })
  document.addEventListener('cut', (e) => {
    if (A.mode !== 'sheet' || A.edit || e.target.closest?.('input,textarea,.dlg')) return
    doCopy(e, true)
  })
  document.addEventListener('paste', (e) => {
    if (A.mode !== 'sheet' || A.edit || e.target.closest?.('input,textarea,.dlg')) return
    e.preventDefault()
    doPaste(e.clipboardData.getData('text/plain'))
  })
  readColors()
  resize()
}
