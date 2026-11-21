// The spreadsheet core: cell addresses, the formula parser, the evaluator (with dynamic-array spill), the workbook model and the
// operations that change it. Every change is an operation (plain JSON) that returns its inverse, so the page, the server and undo
// all use the same code. No dependencies.
import { E, XlErr, errByCode, formatValue, general, isDateFormat, parseTyped } from './format.js'

export const MAXR = 1048576
export const MAXC = 16384
export const FN = Object.create(null) // filled by functions.js

// ── addresses ────────────────────────────────────────────────────────────────────────────────────────────────────────────
export const colName = (c) => {
  let s = ''
  let n = c + 1
  while (n > 0) {
    const m = (n - 1) % 26
    s = String.fromCharCode(65 + m) + s
    n = Math.floor((n - 1) / 26)
  }
  return s
}
export const colIndex = (s) => {
  let n = 0
  for (const ch of String(s).toUpperCase()) n = n * 26 + ch.charCodeAt(0) - 64
  return n - 1
}
export const addr = (r, c) => colName(c) + (r + 1)
export const parseAddr = (a) => {
  const m = String(a).match(/^\$?([A-Za-z]{1,3})\$?(\d+)$/)
  if (!m) return null
  const c = colIndex(m[1])
  const r = Number(m[2]) - 1
  return c < MAXC && r >= 0 && r < MAXR ? { r, c } : null
}
// 'A1', 'A1:B3', 'A:A', '3:5' → { r1, c1, r2, c2 } (normalised, or null)
export const parseRange = (s) => {
  const t = String(s || '').replace(/\$/g, '').trim()
  let m = t.match(/^([A-Za-z]{1,3})(\d+)(?::([A-Za-z]{1,3})(\d+))?$/)
  if (m) {
    const a = parseAddr(m[1] + m[2])
    const b = m[3] ? parseAddr(m[3] + m[4]) : a
    if (!a || !b) return null
    return { r1: Math.min(a.r, b.r), c1: Math.min(a.c, b.c), r2: Math.max(a.r, b.r), c2: Math.max(a.c, b.c) }
  }
  m = t.match(/^([A-Za-z]{1,3}):([A-Za-z]{1,3})$/)
  if (m) {
    const a = colIndex(m[1])
    const b = colIndex(m[2])
    return { r1: 0, c1: Math.min(a, b), r2: MAXR - 1, c2: Math.max(a, b) }
  }
  m = t.match(/^(\d+):(\d+)$/)
  if (m) return { r1: Math.min(+m[1], +m[2]) - 1, c1: 0, r2: Math.max(+m[1], +m[2]) - 1, c2: MAXC - 1 }
  return null
}
export const rangeText = (g) => {
  if (g.r1 === 0 && g.r2 === MAXR - 1) return `${colName(g.c1)}:${colName(g.c2)}`
  if (g.c1 === 0 && g.c2 === MAXC - 1) return `${g.r1 + 1}:${g.r2 + 1}`
  return g.r1 === g.r2 && g.c1 === g.c2 ? addr(g.r1, g.c1) : `${addr(g.r1, g.c1)}:${addr(g.r2, g.c2)}`
}
export const newId = () => Math.random().toString(36).slice(2, 8)

// ── coercions ────────────────────────────────────────────────────────────────────────────────────────────────────────────
export class Rng {
  constructor(sh, r1, c1, r2, c2) {
    this.sh = sh
    this.r1 = r1
    this.c1 = c1
    this.r2 = r2
    this.c2 = c2
  }
}
export const isArr = (v) => Array.isArray(v)
export const toNum = (v) => {
  if (v instanceof XlErr) throw v
  if (typeof v === 'number') return v
  if (v == null || v === '') return 0
  if (typeof v === 'boolean') return v ? 1 : 0
  const t = String(v).trim()
  if (t === '') throw E.VALUE
  const p = parseTyped(t)
  if (p && typeof p.v === 'number') return p.v
  throw E.VALUE
}
export const toStr = (v) => {
  if (v instanceof XlErr) throw v
  if (v == null) return ''
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE'
  if (typeof v === 'number') return general(v)
  return String(v)
}
export const toBool = (v) => {
  if (v instanceof XlErr) throw v
  if (typeof v === 'boolean') return v
  if (typeof v === 'number') return v !== 0
  if (v == null || v === '') return false
  const t = String(v).trim().toUpperCase()
  if (t === 'TRUE') return true
  if (t === 'FALSE') return false
  throw E.VALUE
}
const rank = (v) => (typeof v === 'number' ? 0 : typeof v === 'string' ? 1 : 2)
// -1, 0, 1 like a spreadsheet: numbers < text < logicals, text without case, blanks as the other side's zero/empty
export const cmp = (a, b) => {
  if (a == null) a = typeof b === 'string' ? '' : typeof b === 'boolean' ? false : 0
  if (b == null) b = typeof a === 'string' ? '' : typeof a === 'boolean' ? false : 0
  const ra = rank(a)
  const rb = rank(b)
  if (ra !== rb) return ra < rb ? -1 : 1
  if (ra === 1) {
    const x = a.toLowerCase()
    const y = b.toLowerCase()
    return x < y ? -1 : x > y ? 1 : 0
  }
  return a < b ? -1 : a > b ? 1 : 0
}

// ── tokenizer ────────────────────────────────────────────────────────────────────────────────────────────────────────────
const CELL = String.raw`\$?[A-Za-z]{1,3}\$?\d+`
const RE_RANGE = new RegExp(`^(${CELL})(?::(${CELL}))?(?![\\w.(])`)
const RE_COLS = /^(\$?[A-Za-z]{1,3}):(\$?[A-Za-z]{1,3})(?![\w.(])/
const RE_ROWS = /^(\$?\d+):(\$?\d+)(?![\w.(])/
const RE_SHEET = /^(?:'((?:[^']|'')+)'|([A-Za-z_][\w.]*))!/
const RE_ERR = /^#(DIV\/0!|N\/A|NAME\?|NULL!|NUM!|REF!|VALUE!|SPILL!|CALC!|CIRC!|ERROR!)/i

const one = (s) => {
  const m = s.match(/^(\$?)([A-Za-z]{1,3})(\$?)(\d+)$/)
  const c = colIndex(m[2])
  const r = Number(m[4]) - 1
  return c < MAXC && r < MAXR && r >= 0 ? { c, r, ac: !!m[1], ar: !!m[3] } : null
}
// try a reference at the start of s; returns { ref, len } or null
function refAt(s, sheet) {
  let m = s.match(RE_RANGE)
  if (m) {
    const a = one(m[1])
    const b = m[2] ? one(m[2]) : a
    if (a && b) {
      return {
        len: m[0].length,
        ref: {
          sheet,
          kind: m[2] ? 'range' : 'cell',
          r1: Math.min(a.r, b.r),
          c1: Math.min(a.c, b.c),
          r2: Math.max(a.r, b.r),
          c2: Math.max(a.c, b.c),
          ar1: a.r <= b.r ? a.ar : b.ar,
          ac1: a.c <= b.c ? a.ac : b.ac,
          ar2: a.r <= b.r ? b.ar : a.ar,
          ac2: a.c <= b.c ? b.ac : a.ac,
        },
      }
    }
  }
  m = s.match(RE_COLS)
  if (m) {
    const a = colIndex(m[1].replace('$', ''))
    const b = colIndex(m[2].replace('$', ''))
    if (a < MAXC && b < MAXC)
      return {
        len: m[0].length,
        ref: { sheet, kind: 'col', r1: 0, r2: MAXR - 1, c1: Math.min(a, b), c2: Math.max(a, b), ac1: m[1][0] === '$', ac2: m[2][0] === '$', ar1: false, ar2: false },
      }
  }
  m = s.match(RE_ROWS)
  if (m) {
    const a = Number(m[1].replace('$', '')) - 1
    const b = Number(m[2].replace('$', '')) - 1
    if (a >= 0 && b >= 0 && a < MAXR && b < MAXR)
      return {
        len: m[0].length,
        ref: { sheet, kind: 'row', c1: 0, c2: MAXC - 1, r1: Math.min(a, b), r2: Math.max(a, b), ar1: m[1][0] === '$', ar2: m[2][0] === '$', ac1: false, ac2: false },
      }
  }
  return null
}

export class FormulaError extends Error {}

const OPS = ['<>', '<=', '>=', '+', '-', '*', '/', '^', '&', '=', '<', '>', '(', ')', ',', ';', '{', '}', '%', ':']
export function tokenize(src) {
  const toks = []
  const n = src.length
  let i = 0
  const operandBefore = () => {
    const p = toks[toks.length - 1]
    return !!p && (['num', 'str', 'ref', 'name', 'bool', 'err'].includes(p.t) || (p.t === 'op' && [')', '}', '%', '#'].includes(p.v)))
  }
  while (i < n) {
    const ch = src[i]
    if (/\s/.test(ch)) {
      i++
      continue
    }
    const start = i
    if (ch === '"') {
      let j = i + 1
      let s = ''
      for (;;) {
        if (j >= n) throw new FormulaError('unterminated text')
        if (src[j] === '"') {
          if (src[j + 1] === '"') {
            s += '"'
            j += 2
            continue
          }
          break
        }
        s += src[j++]
      }
      toks.push({ t: 'str', v: s, start, end: j + 1 })
      i = j + 1
      continue
    }
    if (ch === '#') {
      const m = src.slice(i).match(RE_ERR)
      if (m) {
        toks.push({ t: 'err', v: errByCode(`#${m[1].toUpperCase()}`), start, end: i + m[0].length })
        i += m[0].length
        continue
      }
      if (operandBefore()) {
        toks.push({ t: 'op', v: '#', start, end: i + 1 })
        i++
        continue
      }
      throw new FormulaError('unexpected #')
    }
    // sheet-qualified reference
    const sm = src.slice(i).match(RE_SHEET)
    if (sm) {
      const sheet = sm[1] != null ? sm[1].replace(/''/g, "'") : sm[2]
      const rest = src.slice(i + sm[0].length)
      if (/^#REF!/i.test(rest)) {
        toks.push({ t: 'err', v: E.REF, start, end: i + sm[0].length + 5 })
        i += sm[0].length + 5
        continue
      }
      const r = refAt(rest, sheet)
      if (r) {
        toks.push({ t: 'ref', ref: r.ref, start, end: i + sm[0].length + r.len })
        i += sm[0].length + r.len
        continue
      }
      throw new FormulaError('bad reference')
    }
    if (/[0-9.]/.test(ch)) {
      // a row range like 1:3 only where an operand may start
      if (!operandBefore()) {
        const rr = refAt(src.slice(i), null)
        if (rr && rr.ref.kind === 'row') {
          toks.push({ t: 'ref', ref: rr.ref, start, end: i + rr.len })
          i += rr.len
          continue
        }
      }
      const m = src.slice(i).match(/^(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?/)
      if (!m) throw new FormulaError('bad number')
      toks.push({ t: 'num', v: Number(m[0]), start, end: i + m[0].length })
      i += m[0].length
      continue
    }
    if (/[A-Za-z_$\\]/.test(ch)) {
      const rr = refAt(src.slice(i), null)
      if (rr) {
        toks.push({ t: 'ref', ref: rr.ref, start, end: i + rr.len })
        i += rr.len
        continue
      }
      const m = src.slice(i).match(/^[A-Za-z_\\][\w.]*/)
      if (!m) throw new FormulaError(`unexpected ${ch}`)
      const after = src.slice(i + m[0].length).match(/^\s*(.)/)
      const up = m[0].toUpperCase()
      if (after && after[1] === '(') toks.push({ t: 'fn', v: up, start, end: i + m[0].length })
      else if (up === 'TRUE' || up === 'FALSE') toks.push({ t: 'bool', v: up === 'TRUE', start, end: i + m[0].length })
      else toks.push({ t: 'name', v: m[0], start, end: i + m[0].length })
      i += m[0].length
      continue
    }
    const op = OPS.find((o) => src.startsWith(o, i))
    if (!op) throw new FormulaError(`unexpected ${ch}`)
    toks.push({ t: 'op', v: op, start, end: i + op.length })
    i += op.length
  }
  return toks
}

// ── parser: precedence  comparison < & < + - < * / < ^ < unary < % ───────────────────────────────────────────────────────
export function parse(src) {
  const toks = tokenize(src)
  let p = 0
  const peek = () => toks[p]
  const isOp = (v) => toks[p] && toks[p].t === 'op' && toks[p].v === v
  const expect = (v) => {
    if (!isOp(v)) throw new FormulaError(`expected ${v}`)
    p++
  }
  const level = (next, ops) => () => {
    let a = next()
    while (peek() && peek().t === 'op' && ops.includes(peek().v)) {
      const op = toks[p++].v
      a = { t: 'bin', op, a, b: next() }
    }
    return a
  }
  const primary = () => {
    const k = toks[p++]
    if (!k) throw new FormulaError('unexpected end')
    switch (k.t) {
      case 'num':
      case 'str':
      case 'bool':
        return { t: k.t, v: k.v }
      case 'err':
        return { t: 'err', v: k.v }
      case 'ref': {
        const node = { t: 'ref', ...k.ref }
        if (isOp('#')) {
          p++
          return { t: 'spill', ref: node }
        }
        return node
      }
      case 'name':
        return { t: 'name', n: k.v }
      case 'fn': {
        expect('(')
        const args = []
        if (isOp(')')) {
          p++
          return { t: 'call', name: k.v, args }
        }
        for (;;) {
          if (isOp(',') || isOp(')')) args.push({ t: 'empty' })
          else args.push(compare())
          if (isOp(',')) {
            p++
            continue
          }
          expect(')')
          break
        }
        return { t: 'call', name: k.v, args }
      }
      case 'op':
        if (k.v === '(') {
          const e = compare()
          expect(')')
          return { t: 'paren', a: e }
        }
        if (k.v === '{') {
          const rows = [[]]
          for (;;) {
            rows[rows.length - 1].push(unary())
            if (isOp(',')) {
              p++
              continue
            }
            if (isOp(';')) {
              p++
              rows.push([])
              continue
            }
            expect('}')
            break
          }
          return { t: 'arr', rows }
        }
        throw new FormulaError(`unexpected ${k.v}`)
      default:
        throw new FormulaError('unexpected token')
    }
  }
  const postfix = () => {
    let a = primary()
    while (isOp('%')) {
      p++
      a = { t: 'pct', a }
    }
    return a
  }
  const unary = () => {
    if (isOp('-') || isOp('+')) {
      const op = toks[p++].v
      return { t: 'un', op, a: unary() }
    }
    return postfix()
  }
  const power = level(unary, ['^'])
  const mul = level(power, ['*', '/'])
  const add = level(mul, ['+', '-'])
  const concat = level(add, ['&'])
  const compare = level(concat, ['=', '<>', '<', '>', '<=', '>='])
  const ast = compare()
  if (p < toks.length) throw new FormulaError(`unexpected ${toks[p].v ?? toks[p].t}`)
  return ast
}

// ── rewriting references inside formula text (insert/delete rows, paste, rename a sheet) ─────────────────────────────────
const quoteSheet = (s) => (/^[A-Za-z_][\w.]*$/.test(s) && !/^[A-Za-z]{1,3}\d+$/.test(s) ? s : `'${s.replace(/'/g, "''")}'`)
export function refText(g) {
  const pre = g.sheet != null ? `${quoteSheet(g.sheet)}!` : ''
  const a = (r, c, ar, ac) => (ac ? '$' : '') + colName(c) + (ar ? '$' : '') + (r + 1)
  if (g.kind === 'col') return `${pre}${g.ac1 ? '$' : ''}${colName(g.c1)}:${g.ac2 ? '$' : ''}${colName(g.c2)}`
  if (g.kind === 'row') return `${pre}${g.ar1 ? '$' : ''}${g.r1 + 1}:${g.ar2 ? '$' : ''}${g.r2 + 1}`
  const one1 = a(g.r1, g.c1, g.ar1, g.ac1)
  if (g.kind === 'cell') return pre + one1
  return `${pre}${one1}:${a(g.r2, g.c2, g.ar2, g.ac2)}`
}
// fn(ref) returns the new ref, or null for #REF!
export function mapRefs(src, fn) {
  let toks
  try {
    toks = tokenize(src)
  } catch {
    return src
  }
  let out = ''
  let last = 0
  for (const t of toks) {
    if (t.t !== 'ref') continue
    const next = fn({ ...t.ref })
    out += src.slice(last, t.start) + (next ? refText(next) : '#REF!')
    last = t.end
  }
  return out + src.slice(last)
}
// a formula moved by (dr, dc): relative parts move, $absolute parts stay
export function shiftFormula(src, dr, dc) {
  return mapRefs(src, (g) => {
    const n = { ...g }
    if (g.kind !== 'col' && g.kind !== 'row' || g.kind === 'row') {
      if (!g.ar1) n.r1 += dr
      if (!g.ar2) n.r2 += dr
    }
    if (g.kind !== 'row') {
      if (!g.ac1) n.c1 += dc
      if (!g.ac2) n.c2 += dc
    }
    if (g.kind === 'col') {
      n.r1 = 0
      n.r2 = MAXR - 1
    }
    if (g.kind === 'row') {
      n.c1 = 0
      n.c2 = MAXC - 1
    }
    if (n.r1 < 0 || n.c1 < 0 || n.r2 >= MAXR || n.c2 >= MAXC) return null
    return n
  })
}

// ── typed input ──────────────────────────────────────────────────────────────────────────────────────────────────────────
// text from a person or a tool → { cell: {v} | {f}, nf? }
export function parseInput(input) {
  if (input === null || input === undefined || input === '') return { cell: null }
  if (typeof input === 'number') return { cell: Number.isFinite(input) ? { v: input } : { v: String(input) } }
  if (typeof input === 'boolean') return { cell: { v: input } }
  const text = String(input)
  if (text.length > 1 && text[0] === '=') return { cell: { f: text.slice(1) } }
  if (text[0] === "'") return { cell: { v: text.slice(1) } }
  const t = text.trim()
  if (/^(true|false)$/i.test(t)) return { cell: { v: t.toLowerCase() === 'true' } }
  const typed = parseTyped(t)
  if (typed) return { cell: { v: typed.v }, nf: typed.nf }
  return { cell: { v: text } }
}
// what the editor shows for a cell
export function rawText(cell) {
  if (!cell) return ''
  if (cell.f != null) return `=${cell.f}`
  const v = cell.v
  if (v == null) return ''
  if (typeof v === 'string') return /^[='"]/.test(v) || (parseTyped(v) && v.trim() !== '') || /^(true|false)$/i.test(v.trim()) ? `'${v}` : v
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE'
  const nf = cell.s?.nf
  if (nf && /%/.test(nf)) return `${general(Math.round(v * 1e10) / 1e8)}%`
  if (nf && isDateFormat(nf)) return formatValue(v, nf)
  return general(v)
}

// ── evaluator ────────────────────────────────────────────────────────────────────────────────────────────────────────────
const matOf = (book, v) => (v instanceof Rng ? book.matrix(v) : isArr(v) ? v : [[v]])
export const toMat = matOf
const deref = (book, v) => {
  if (v instanceof Rng) {
    if (v.r1 === v.r2 && v.c1 === v.c2) return book.get(v.sh, v.r1, v.c1)
    return book.matrix(v)
  }
  return v
}
// apply f over scalars or arrays, broadcasting like a spreadsheet
export function broadcast(vals, f) {
  if (!vals.some(isArr)) return f(vals)
  const ms = vals.map((v) => (isArr(v) ? v : [[v]]))
  const R = Math.max(...ms.map((m) => (m.length === 1 ? 1 : m.length)))
  const C = Math.max(...ms.map((m) => (m[0].length === 1 ? 1 : m[0].length)))
  const out = []
  for (let r = 0; r < R; r++) {
    const row = []
    for (let c = 0; c < C; c++) {
      const els = []
      let bad = false
      for (const m of ms) {
        const rr = m.length === 1 ? 0 : r
        const cc = m[0].length === 1 ? 0 : c
        if (rr >= m.length || cc >= m[0].length) bad = true
        els.push(bad ? null : m[rr][cc])
      }
      if (bad) row.push(E.NA)
      else
        try {
          row.push(f(els))
        } catch (e) {
          if (e instanceof XlErr) row.push(e)
          else throw e
        }
    }
    out.push(row)
  }
  return out
}
const arith = {
  '+': (a, b) => toNum(a) + toNum(b),
  '-': (a, b) => toNum(a) - toNum(b),
  '*': (a, b) => toNum(a) * toNum(b),
  '/': (a, b) => {
    const d = toNum(b)
    if (d === 0) throw E.DIV0
    return toNum(a) / d
  },
  '^': (a, b) => {
    const x = toNum(a)
    const y = toNum(b)
    if (x === 0 && y < 0) throw E.DIV0
    const r = x ** y
    if (!Number.isFinite(r)) throw E.NUM
    return r
  },
  '&': (a, b) => toStr(a) + toStr(b),
  '=': (a, b) => (a instanceof XlErr ? thrower(a) : b instanceof XlErr ? thrower(b) : cmp(a, b) === 0),
  '<>': (a, b) => (a instanceof XlErr ? thrower(a) : b instanceof XlErr ? thrower(b) : cmp(a, b) !== 0),
  '<': (a, b) => (a instanceof XlErr ? thrower(a) : b instanceof XlErr ? thrower(b) : cmp(a, b) < 0),
  '>': (a, b) => (a instanceof XlErr ? thrower(a) : b instanceof XlErr ? thrower(b) : cmp(a, b) > 0),
  '<=': (a, b) => (a instanceof XlErr ? thrower(a) : b instanceof XlErr ? thrower(b) : cmp(a, b) <= 0),
  '>=': (a, b) => (a instanceof XlErr ? thrower(a) : b instanceof XlErr ? thrower(b) : cmp(a, b) >= 0),
}
const thrower = (e) => {
  throw e
}

export function ev(n, x) {
  const book = x.book
  switch (n.t) {
    case 'num':
    case 'str':
    case 'bool':
      return n.v
    case 'err':
      return n.v
    case 'empty':
      return null
    case 'paren':
      return ev(n.a, x)
    case 'ref': {
      const sh = n.sheet == null ? x.sh : book.sheetIndex(n.sheet)
      if (sh < 0) throw E.REF
      return new Rng(sh, n.r1, n.c1, n.r2, n.c2)
    }
    case 'spill': {
      const sh = n.ref.sheet == null ? x.sh : book.sheetIndex(n.ref.sheet)
      if (sh < 0) throw E.REF
      const cell = book.sheets[sh].cells.get(n.ref.r1 * MAXC + n.ref.c1)
      book.get(sh, n.ref.r1, n.ref.c1)
      if (!cell?._mat) throw E.REF
      return cell._mat
    }
    case 'name': {
      if (x.vars && n.n.toLowerCase() in x.vars) return x.vars[n.n.toLowerCase()]
      const def = book.nameDef(n.n)
      if (def == null) throw E.NAME
      return ev(def, { ...x, vars: null })
    }
    case 'un': {
      const v = deref(book, ev(n.a, x))
      return broadcast([v], ([a]) => (n.op === '-' ? -toNum(a) : a instanceof XlErr ? thrower(a) : a))
    }
    case 'pct': {
      const v = deref(book, ev(n.a, x))
      return broadcast([v], ([a]) => toNum(a) / 100)
    }
    case 'bin': {
      const a = deref(book, ev(n.a, x))
      const b = deref(book, ev(n.b, x))
      const f = arith[n.op]
      return broadcast([a, b], ([p, q]) => f(p, q))
    }
    case 'arr':
      return n.rows.map((row) => row.map((e) => deref(book, ev(e, x))))
    case 'call': {
      const def = FN[n.name]
      if (!def) throw E.NAME
      if (def.lazy) return def.f(n.args, x)
      const args = n.args.map((a) => {
        if (!def.raw) return ev(a, x)
        try {
          return ev(a, x)
        } catch (e) {
          if (e instanceof XlErr) return e
          throw e
        }
      })
      return def.f(args, x)
    }
    default:
      throw E.VALUE
  }
}

// ── the workbook ─────────────────────────────────────────────────────────────────────────────────────────────────────────
const SHEET_PROPS = ['cols', 'rows', 'hiddenCols', 'hiddenRows', 'freeze', 'merges', 'cf', 'dv', 'charts', 'filter', 'color', 'gridlines', 'name']
const clone = (v) => (v === undefined ? undefined : JSON.parse(JSON.stringify(v)))
const ck = (r, c) => r * MAXC + c

function mkSheet(j = {}) {
  const s = {
    id: j.id || newId(),
    name: j.name || 'Sheet1',
    cells: new Map(),
    cols: clone(j.cols) || {},
    rows: clone(j.rows) || {},
    hiddenCols: clone(j.hiddenCols) || [],
    hiddenRows: clone(j.hiddenRows) || [],
    freeze: { rows: 0, cols: 0, ...(j.freeze || {}) },
    merges: clone(j.merges) || [],
    cf: clone(j.cf) || [],
    dv: clone(j.dv) || [],
    charts: clone(j.charts) || [],
    filter: clone(j.filter) || null,
    color: j.color || '',
    gridlines: j.gridlines !== false,
    spill: new Map(),
    maxR: -1,
    maxC: -1,
  }
  for (const [a, cell] of Object.entries(j.cells || {})) {
    const p = parseAddr(a)
    if (p && cell) s.cells.set(ck(p.r, p.c), { ...clone(cell) })
  }
  return s
}

export class Book {
  constructor(data = {}) {
    this.load(data)
  }
  load(data = {}) {
    this.id = data.id || newId()
    this.name = data.name || 'Untitled workbook'
    this.rev = data.rev || 0
    this.created = data.created || Date.now()
    this.updated = data.updated || Date.now()
    this.names = clone(data.names) || {}
    this.sheets = (data.sheets?.length ? data.sheets : [{ name: 'Sheet1' }]).map(mkSheet)
    this.ver = 0
    this.dirty = true
    this._names = new Map()
  }
  toJSON() {
    return {
      id: this.id,
      name: this.name,
      rev: this.rev,
      created: this.created,
      updated: this.updated,
      names: clone(this.names),
      sheets: this.sheets.map((s) => this.sheetJSON(s)),
    }
  }
  sheetJSON(s) {
    const keys = [...s.cells.keys()].sort((a, b) => a - b)
    const cells = {}
    for (const k of keys) {
      const { v, f, s: st } = s.cells.get(k)
      const o = {}
      if (f != null) o.f = f
      else if (v !== undefined && v !== null) o.v = v
      if (st && Object.keys(st).length) o.s = st
      cells[addr(Math.floor(k / MAXC), k % MAXC)] = o
    }
    const out = { id: s.id, name: s.name, cells }
    for (const p of ['cols', 'rows']) if (Object.keys(s[p]).length) out[p] = clone(s[p])
    for (const p of ['hiddenCols', 'hiddenRows', 'merges', 'cf', 'dv', 'charts']) if (s[p].length) out[p] = clone(s[p])
    if (s.freeze.rows || s.freeze.cols) out.freeze = { ...s.freeze }
    if (s.filter) out.filter = clone(s.filter)
    if (s.color) out.color = s.color
    if (!s.gridlines) out.gridlines = false
    return out
  }

  // sheets
  sheetIndex(ref) {
    if (typeof ref === 'number') return ref >= 0 && ref < this.sheets.length ? ref : -1
    const t = String(ref ?? '')
    let i = this.sheets.findIndex((s) => s.id === t)
    if (i < 0) i = this.sheets.findIndex((s) => s.name === t)
    if (i < 0) i = this.sheets.findIndex((s) => s.name.toLowerCase() === t.toLowerCase())
    return i
  }
  sheet(ref) {
    const i = this.sheetIndex(ref)
    if (i < 0) throw new Error(`no sheet "${ref}"`)
    return this.sheets[i]
  }
  nameDef(n) {
    const key = n.toLowerCase()
    const found = Object.keys(this.names).find((k) => k.toLowerCase() === key)
    if (!found) return null
    const text = String(this.names[found]).replace(/^=/, '')
    let d = this._names.get(text)
    if (!d) {
      try {
        d = parse(text)
      } catch {
        return null
      }
      this._names.set(text, d)
    }
    return d
  }

  // values
  matrix(g) {
    const s = this.sheets[g.sh]
    this.extent(s)
    const r2 = Math.min(g.r2, Math.max(s.maxR, g.r1))
    const c2 = Math.min(g.c2, Math.max(s.maxC, g.c1))
    const out = []
    for (let r = g.r1; r <= r2; r++) {
      const row = []
      for (let c = g.c1; c <= c2; c++) row.push(this.get(g.sh, r, c))
      out.push(row)
    }
    return out
  }
  extent(s) {
    if (s._extVer === this.ver) return
    s._extVer = this.ver
    let mr = -1
    let mc = -1
    for (const k of s.cells.keys()) {
      const r = Math.floor(k / MAXC)
      const c = k % MAXC
      if (r > mr) mr = r
      if (c > mc) mc = c
    }
    for (const k of s.spill.keys()) {
      const r = Math.floor(k / MAXC)
      const c = k % MAXC
      if (r > mr) mr = r
      if (c > mc) mc = c
    }
    s.maxR = mr
    s.maxC = mc
  }
  usedRange(ref) {
    const s = this.sheet(ref)
    this.recalc()
    this.extent(s)
    return s.maxR < 0 ? null : { r1: 0, c1: 0, r2: s.maxR, c2: s.maxC }
  }
  get(sh, r, c) {
    const s = this.sheets[sh]
    const k = ck(r, c)
    const cell = s.cells.get(k)
    if (!cell) {
      const sp = s.spill.get(k)
      return sp ? sp.v : null
    }
    if (cell.f == null) return cell.v === undefined ? null : cell.v
    if (cell._ver === this.ver) return cell._val
    if (cell._busy) return E.CIRC
    cell._busy = true
    let val
    cell._mat = null
    try {
      if (!cell._ast) {
        try {
          cell._ast = parse(cell.f)
        } catch {
          cell._ast = { t: 'bad' }
        }
      }
      if (cell._ast.t === 'bad') val = E.PARSE
      else {
        let res = ev(cell._ast, { book: this, sh, r, c, vars: null })
        if (res instanceof Rng) res = res.r1 === res.r2 && res.c1 === res.c2 ? this.get(res.sh, res.r1, res.c1) : this.matrix(res)
        if (isArr(res)) {
          if (res.length === 1 && res[0].length === 1) val = res[0][0]
          else if (this.claim(s, r, c, res)) {
            cell._mat = res
            val = res[0][0]
          } else val = E.SPILL
        } else val = res
        if (val === undefined) val = null
      }
    } catch (e) {
      if (e instanceof XlErr) val = e
      else if (e instanceof RangeError) val = E.CALC
      else throw e
    } finally {
      cell._busy = false
    }
    cell._val = val
    cell._ver = this.ver
    return val
  }
  claim(s, r, c, mat) {
    const h = mat.length
    const w = mat[0].length
    if (r + h > MAXR || c + w > MAXC) return false
    const me = ck(r, c)
    for (let i = 0; i < h; i++)
      for (let j = 0; j < w; j++) {
        if (!i && !j) continue
        const k = ck(r + i, c + j)
        if (s.cells.has(k)) return false
        const e = s.spill.get(k)
        if (e && e.ver === this.ver && e.anchor !== me) return false
      }
    for (let i = 0; i < h; i++)
      for (let j = 0; j < w; j++) if (i || j) s.spill.set(ck(r + i, c + j), { anchor: me, v: mat[i][j], ver: this.ver })
    this._spilled = true
    return true
  }
  recalc() {
    if (!this.dirty) return
    const pass = () => {
      this.ver++
      this._spilled = false
      for (let sh = 0; sh < this.sheets.length; sh++)
        for (const [k, cell] of this.sheets[sh].cells) if (cell.f != null) this.get(sh, Math.floor(k / MAXC), k % MAXC)
    }
    pass()
    if (this._spilled || this.sheets.some((s) => s.spill.size)) pass() // a second pass lets earlier cells see later spills
    for (const s of this.sheets) for (const [k, e] of s.spill) if (e.ver < this.ver) s.spill.delete(k)
    this.dirty = false
  }
  // the value, as text, as the grid shows it
  value(ref, r, c) {
    this.recalc()
    return this.get(this.sheetIndex(ref), r, c)
  }
  cell(ref, r, c) {
    const s = this.sheet(ref)
    return s.cells.get(ck(r, c)) || null
  }
  display(sh, r, c) {
    const s = this.sheets[sh]
    const cell = s.cells.get(ck(r, c))
    const v = this.get(sh, r, c)
    let nf = cell?.s?.nf
    if (!nf && typeof v === 'number' && cell?.f != null) nf = this.inferFormat(cell)
    return formatValue(v, nf)
  }
  // a formula that returns a date keeps showing a date when it is a plain reference or TODAY/NOW/DATE
  inferFormat(cell) {
    const f = cell.f.trim().toUpperCase()
    if (/^(TODAY|DATE|EDATE|EOMONTH|DATEVALUE)\(/.test(f)) return 'yyyy-mm-dd'
    if (/^NOW\(/.test(f)) return 'yyyy-mm-dd hh:mm'
    if (/^TIME\(/.test(f)) return 'hh:mm:ss'
    return ''
  }
  // the evaluated result of a formula that is not stored in any cell (for tools and the formula bar)
  evaluate(formula, ref = 0) {
    this.recalc()
    const sh = this.sheetIndex(ref)
    try {
      let res = ev(parse(String(formula).replace(/^=/, '')), { book: this, sh: Math.max(sh, 0), r: 0, c: 0, vars: null })
      if (res instanceof Rng) res = res.r1 === res.r2 && res.c1 === res.c2 ? this.get(res.sh, res.r1, res.c1) : this.matrix(res)
      return res
    } catch (e) {
      if (e instanceof XlErr) return e
      if (e instanceof FormulaError) return E.PARSE
      throw e
    }
  }

  // ── operations ─────────────────────────────────────────────────────────────────────────────────────────────────────────
  // apply an operation, return the operation that undoes it
  apply(op) {
    this.dirty = true
    this.updated = Date.now()
    switch (op.op) {
      case 'set': {
        const s = this.sheet(op.sheet)
        const prev = {}
        for (const [a, cell] of Object.entries(op.cells)) {
          const p = parseAddr(a)
          if (!p) continue
          const k = ck(p.r, p.c)
          const old = s.cells.get(k)
          prev[a] = old ? this.cellOut(old) : null
          if (!cell || (cell.v == null && cell.f == null && !(cell.s && Object.keys(cell.s).length))) s.cells.delete(k)
          else {
            const n = {}
            if (cell.f != null) n.f = String(cell.f).replace(/^=/, '')
            else if (cell.v !== undefined && cell.v !== null) n.v = cell.v
            if (cell.s && Object.keys(cell.s).length) n.s = clone(cell.s)
            s.cells.set(k, n)
          }
        }
        return { op: 'set', sheet: s.id, cells: prev }
      }
      case 'prop': {
        const s = this.sheet(op.sheet)
        if (!SHEET_PROPS.includes(op.key)) throw new Error(`unknown sheet property ${op.key}`)
        const old = clone(s[op.key])
        if (op.key === 'name') return this.apply({ op: 'renameSheet', id: s.id, name: op.value }), { op: 'renameSheet', id: s.id, name: old }
        s[op.key] = clone(op.value) ?? (Array.isArray(old) ? [] : op.key === 'freeze' ? { rows: 0, cols: 0 } : {})
        return { op: 'prop', sheet: s.id, key: op.key, value: old }
      }
      case 'addSheet': {
        const j = op.sheet || {}
        let name = j.name || `Sheet${this.sheets.length + 1}`
        for (let n = 2; this.sheetIndex(name) >= 0; n++) name = `${j.name || 'Sheet'} ${n}`
        const s = mkSheet({ ...j, name })
        const at = Math.min(Math.max(op.index ?? this.sheets.length, 0), this.sheets.length)
        this.sheets.splice(at, 0, s)
        return { op: 'removeSheet', id: s.id, quiet: true }
      }
      case 'removeSheet': {
        if (this.sheets.length < 2) throw new Error('a workbook needs at least one sheet')
        const i = this.sheetIndex(op.id)
        if (i < 0) throw new Error('no such sheet')
        const before = this.snapshot()
        const name = this.sheets[i].name
        this.sheets.splice(i, 1)
        if (!op.quiet) this.rewrite((g, own) => (g.sheet != null && g.sheet.toLowerCase() === name.toLowerCase() ? null : g))
        return { op: 'snapshot', state: before }
      }
      case 'renameSheet': {
        const s = this.sheet(op.id)
        const name = String(op.name || '').trim()
        if (!name || name.length > 60 || /[[\]:*?/\\]/.test(name)) throw new Error('that is not a valid sheet name')
        const clash = this.sheets.find((x) => x !== s && x.name.toLowerCase() === name.toLowerCase())
        if (clash) throw new Error(`there is already a sheet called ${name}`)
        const old = s.name
        s.name = name
        this.rewrite((g) => (g.sheet != null && g.sheet.toLowerCase() === old.toLowerCase() ? { ...g, sheet: name } : g))
        return { op: 'renameSheet', id: s.id, name: old }
      }
      case 'moveSheet': {
        const i = this.sheetIndex(op.id)
        if (i < 0) throw new Error('no such sheet')
        const [s] = this.sheets.splice(i, 1)
        this.sheets.splice(Math.min(Math.max(op.index, 0), this.sheets.length), 0, s)
        return { op: 'moveSheet', id: s.id, index: i }
      }
      case 'insert':
      case 'delete': {
        const before = this.snapshot()
        this.structural(this.sheet(op.sheet), op.dim, op.at, op.n || 1, op.op === 'delete')
        return op.op === 'insert' ? { op: 'delete', sheet: op.sheet, dim: op.dim, at: op.at, n: op.n || 1, _undo: true } : { op: 'snapshot', state: before }
      }
      case 'names': {
        const old = clone(this.names)
        this.names = clone(op.value) || {}
        this._names.clear()
        return { op: 'names', value: old }
      }
      case 'rename': {
        const old = this.name
        this.name = String(op.name || '').trim().slice(0, 100) || old
        return { op: 'rename', name: old }
      }
      case 'snapshot': {
        const before = this.snapshot()
        this.restore(op.state)
        return { op: 'snapshot', state: before }
      }
      default:
        throw new Error(`unknown operation ${op.op}`)
    }
  }
  applyAll(ops) {
    const inv = []
    for (const o of ops) inv.unshift(this.apply(o))
    return inv
  }
  cellOut(c) {
    const o = {}
    if (c.f != null) o.f = c.f
    else if (c.v !== undefined) o.v = c.v
    if (c.s) o.s = clone(c.s)
    return o
  }
  snapshot() {
    return { sheets: this.sheets.map((s) => this.sheetJSON(s)), names: clone(this.names) }
  }
  restore(st) {
    this.sheets = st.sheets.map(mkSheet)
    this.names = clone(st.names) || {}
    this._names.clear()
  }
  // rewrite every reference in every formula (and defined names)
  rewrite(fn) {
    for (const s of this.sheets)
      for (const cell of s.cells.values())
        if (cell.f != null) {
          const nf = mapRefs(cell.f, (g) => fn(g, s))
          if (nf !== cell.f) {
            cell.f = nf
            cell._ast = null
          }
        }
    for (const k of Object.keys(this.names)) this.names[k] = mapRefs(String(this.names[k]), (g) => fn(g, null))
    this._names.clear()
  }
  structural(sheet, dim, at, n, del) {
    const row = dim === 'row'
    const lo = row ? 'r1' : 'c1'
    const hi = row ? 'r2' : 'c2'
    const skip = row ? 'col' : 'row'
    const limit = row ? MAXR : MAXC
    const shiftIdx = (i) => (!del ? (i >= at ? i + n : i) : i < at ? i : i >= at + n ? i - n : -1)
    const target = sheet.name.toLowerCase()
    // formulas anywhere that point into this sheet
    this.rewrite((g, own) => {
      const to = g.sheet != null ? g.sheet.toLowerCase() : own ? own.name.toLowerCase() : ''
      if (to !== target || g.kind === skip) return g
      const a = g[lo]
      const b = g[hi]
      const out = { ...g }
      if (!del) {
        out[lo] = a >= at ? a + n : a
        out[hi] = b >= at ? b + n : b
      } else {
        if (a >= at && b < at + n) return null
        out[lo] = a < at ? a : a >= at + n ? a - n : at
        out[hi] = b < at ? b : b >= at + n ? b - n : at - 1
      }
      if (out[hi] >= limit) out[hi] = limit - 1
      return out
    })
    // cells
    const moved = new Map()
    for (const [k, cell] of sheet.cells) {
      const r = Math.floor(k / MAXC)
      const c = k % MAXC
      const i = shiftIdx(row ? r : c)
      if (i < 0 || i >= limit) continue
      moved.set(row ? ck(i, c) : ck(r, i), cell)
    }
    sheet.cells = moved
    sheet.spill.clear()
    // sizes, hidden, freeze
    const sizes = row ? 'rows' : 'cols'
    const hid = row ? 'hiddenRows' : 'hiddenCols'
    const key = (i) => (row ? String(i) : colName(i))
    const idx = (k) => (row ? Number(k) - 1 : colIndex(k))
    const ns = {}
    for (const [k, v] of Object.entries(sheet[sizes])) {
      const i = shiftIdx(idx(k))
      if (i >= 0) ns[row ? String(i + 1) : colName(i)] = v
    }
    sheet[sizes] = ns
    void key
    sheet[hid] = sheet[hid].map(shiftIdx).filter((i) => i >= 0)
    const fz = row ? 'rows' : 'cols'
    if (sheet.freeze[fz] > at) sheet.freeze[fz] = del ? Math.max(at, sheet.freeze[fz] - n) : sheet.freeze[fz] + n
    // ranges kept as text: merges, conditional formats, validation, filter, charts
    const fixRange = (txt) => {
      const g = parseRange(String(txt).replace(/^.*!/, ''))
      if (!g) return txt
      let a = g[lo]
      let b = g[hi]
      if (!del) {
        if (a >= at) a += n
        if (b >= at) b += n
      } else {
        if (a >= at && b < at + n) return null
        a = a < at ? a : a >= at + n ? a - n : at
        b = b < at ? b : b >= at + n ? b - n : at - 1
      }
      return rangeText({ ...g, [lo]: a, [hi]: Math.min(b, limit - 1) })
    }
    sheet.merges = sheet.merges.map(fixRange).filter(Boolean)
    sheet.cf = sheet.cf.map((r) => ({ ...r, range: fixRange(r.range) })).filter((r) => r.range)
    sheet.dv = sheet.dv.map((r) => ({ ...r, range: fixRange(r.range) })).filter((r) => r.range)
    if (sheet.filter?.range) {
      const f = fixRange(sheet.filter.range)
      sheet.filter = f ? { ...sheet.filter, range: f } : null
    }
    for (const sh of this.sheets)
      sh.charts = sh.charts
        .map((ch) => {
          const nr = mapRefs(String(ch.range), (g) => {
            const to = g.sheet != null ? g.sheet.toLowerCase() : sh.name.toLowerCase()
            if (to !== target || g.kind === skip) return g
            const a = g[lo]
            const b = g[hi]
            const out = { ...g }
            if (!del) {
              out[lo] = a >= at ? a + n : a
              out[hi] = b >= at ? b + n : b
            } else {
              if (a >= at && b < at + n) return null
              out[lo] = a < at ? a : a >= at + n ? a - n : at
              out[hi] = b < at ? b : b >= at + n ? b - n : at - 1
            }
            return out
          })
          return /#REF!/.test(nr) ? null : { ...ch, range: nr }
        })
        .filter(Boolean)
  }

  // ── planners: they return a `set` operation, never change anything ────────────────────────────────────────────────────
  // copy src cells to dst top-left, formulas move relative. `from` is a Book sheet ref, cells = {A1: cell JSON} of the source
  planPaste(sheet, srcCells, srcTop, dstTop, { values = false, formats = false } = {}) {
    const s = this.sheet(sheet)
    const out = {}
    for (const [a, cell] of Object.entries(srcCells)) {
      const p = parseAddr(a)
      const dr = dstTop.r - srcTop.r
      const dc = dstTop.c - srcTop.c
      const r = p.r + dr
      const c = p.c + dc
      if (r < 0 || c < 0 || r >= MAXR || c >= MAXC) continue
      const old = s.cells.get(ck(r, c))
      const n = {}
      if (formats) {
        if (old) Object.assign(n, old.f != null ? { f: old.f } : old.v !== undefined ? { v: old.v } : {})
        if (cell?.s) n.s = clone(cell.s)
      } else if (values) {
        const v = cell ? (cell._val !== undefined ? cell._val : cell.v) : undefined
        if (v !== undefined && v !== null && !(v instanceof XlErr)) n.v = v
        if (old?.s) n.s = clone(old.s)
      } else if (cell) {
        if (cell.f != null) n.f = shiftFormula(cell.f, dr, dc)
        else if (cell.v !== undefined) n.v = cell.v
        if (cell.s) n.s = clone(cell.s)
      }
      out[addr(r, c)] = n
    }
    return { op: 'set', sheet: s.id, cells: out }
  }
  // fill the cells of `dst` from the cells of `src` (a range inside or beside it): copy, or continue a series
  planFill(sheet, src, dst, { series = true } = {}) {
    const s = this.sheet(sheet)
    this.recalc()
    const out = {}
    const down = dst.r1 > src.r2 || dst.r2 < src.r1 || (dst.r1 === src.r1 && dst.r2 === src.r2 ? false : dst.c1 >= src.c1 && dst.c2 <= src.c2)
    const horizontal = !down
    const lines = horizontal ? src.r2 - src.r1 + 1 : src.c2 - src.c1 + 1 // how many rows (or columns) are filled independently
    const srcLen = horizontal ? src.c2 - src.c1 + 1 : src.r2 - src.r1 + 1
    for (let l = 0; l < lines; l++) {
      const cellsOf = (i) => {
        const r = horizontal ? src.r1 + l : src.r1 + i
        const c = horizontal ? src.c1 + i : src.c1 + l
        return { r, c, cell: s.cells.get(ck(r, c)) }
      }
      const src1 = Array.from({ length: srcLen }, (_, i) => cellsOf(i))
      const nums = src1.map((x) => (x.cell && x.cell.f == null && typeof x.cell.v === 'number' ? x.cell.v : null))
      const allNum = series && nums.every((v) => v != null)
      let step = 0
      if (allNum) step = srcLen > 1 ? nums.reduce((acc, v, i) => (i ? acc + (v - nums[i - 1]) : acc), 0) / (srcLen - 1) : nums[0] !== undefined && srcLen === 1 ? 0 : 0
      const total = horizontal ? dst.c2 - dst.c1 + 1 : dst.r2 - dst.r1 + 1
      const before = horizontal ? src.c1 - dst.c1 : src.r1 - dst.r1
      const backwards = before > 0 && (horizontal ? dst.c2 < src.c1 : dst.r2 < src.r1)
      for (let i = 0; i < total; i++) {
        const pos = (horizontal ? dst.c1 : dst.r1) + i
        const r = horizontal ? dst.r1 + l : pos
        const c = horizontal ? pos : dst.c1 + l
        if (r >= src.r1 && r <= src.r2 && c >= src.c1 && c <= src.c2) continue
        const offset = pos - (horizontal ? src.c1 : src.r1)
        const k = ((offset % srcLen) + srcLen) % srcLen
        const from = src1[k]
        const n = {}
        const cell = from.cell
        if (cell) {
          if (cell.f != null) n.f = shiftFormula(cell.f, r - from.r, c - from.c)
          else if (allNum && srcLen >= 1 && (srcLen > 1 || cell.s?.nf || true)) {
            const lastNum = nums[srcLen - 1]
            const first = nums[0]
            n.v = srcLen === 1 ? first : backwards ? first + step * offset : lastNum + step * (offset - (srcLen - 1))
          } else if (cell.v !== undefined) {
            const m = typeof cell.v === 'string' && series ? cell.v.match(/^(.*?)(\d+)$/) : null
            if (m && srcLen === 1) n.v = `${m[1]}${String(Number(m[2]) + offset).padStart(m[2].length, '0')}`
            else n.v = nextInList(cell.v, offset, series && srcLen === 1)
          }
          if (cell.s) n.s = clone(cell.s)
        }
        out[addr(r, c)] = n
      }
    }
    return { op: 'set', sheet: s.id, cells: out }
  }
  planSort(sheet, range, keys, { header = false } = {}) {
    const s = this.sheet(sheet)
    this.recalc()
    const g = typeof range === 'string' ? parseRange(range) : range
    if (!g) throw new Error('that is not a range')
    this.extent(s)
    const r2 = Math.min(g.r2, Math.max(s.maxR, g.r1))
    const c2 = Math.min(g.c2, Math.max(s.maxC, g.c1))
    const r0 = g.r1 + (header ? 1 : 0)
    const sh = this.sheetIndex(sheet)
    const rows = []
    for (let r = r0; r <= r2; r++) {
      const cells = []
      for (let c = g.c1; c <= c2; c++) cells.push({ cell: s.cells.get(ck(r, c)), r, c })
      rows.push({ r, cells, keys: (keys || []).map((k) => this.get(sh, r, g.c1 + k.col)) })
    }
    const order = rows
      .map((x, i) => ({ x, i }))
      .sort((a, b) => {
        for (let ki = 0; ki < (keys || []).length; ki++) {
          const va = a.x.keys[ki]
          const vb = b.x.keys[ki]
          const ablank = va == null || va === ''
          const bblank = vb == null || vb === ''
          if (ablank !== bblank) return ablank ? 1 : -1 // blanks last, either direction
          const d = cmp(va, vb)
          if (d) return keys[ki].asc === false ? -d : d
        }
        return a.i - b.i
      })
    const out = {}
    order.forEach(({ x }, i) => {
      const dr = r0 + i - x.r
      for (const { cell, c } of x.cells) {
        const a = addr(r0 + i, c)
        if (!cell) out[a] = null
        else {
          const n = {}
          if (cell.f != null) n.f = shiftFormula(cell.f, dr, 0)
          else if (cell.v !== undefined) n.v = cell.v
          if (cell.s) n.s = clone(cell.s)
          out[a] = n
        }
      }
    })
    return { op: 'set', sheet: s.id, cells: out }
  }
  planFind(sheet, find, replace, { matchCase = false, whole = false, formulas = false, all = true, range = null } = {}) {
    const sheets = sheet == null ? this.sheets : [this.sheet(sheet)]
    const out = []
    const ops = []
    const needle = matchCase ? String(find) : String(find).toLowerCase()
    const g = range ? parseRange(range) : null
    for (const s of sheets) {
      const cells = {}
      for (const [k, cell] of [...s.cells].sort((a, b) => a[0] - b[0])) {
        const r = Math.floor(k / MAXC)
        const c = k % MAXC
        if (g && (r < g.r1 || r > g.r2 || c < g.c1 || c > g.c2)) continue
        const text = formulas || cell.f == null ? (cell.f != null ? `=${cell.f}` : cell.v == null ? '' : typeof cell.v === 'string' ? cell.v : general(cell.v)) : String(this.display(this.sheets.indexOf(s), r, c))
        const hay = matchCase ? text : text.toLowerCase()
        const hit = whole ? hay === needle : hay.includes(needle)
        if (!hit || text === '') continue
        out.push({ sheet: s.name, cell: addr(r, c), text })
        if (replace != null && (cell.f == null || formulas)) {
          let t
          if (whole) t = String(replace)
          else {
            const re = new RegExp(String(find).replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), matchCase ? 'g' : 'gi')
            t = text.replace(re, () => String(replace))
          }
          const p = parseInput(t)
          cells[addr(r, c)] = { ...(p.cell || {}), s: cell.s }
        }
        if (!all && out.length) break
      }
      if (Object.keys(cells).length) ops.push({ op: 'set', sheet: s.id, cells })
      if (!all && out.length) break
    }
    return { matches: out, ops }
  }
}

const LISTS = [
  ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday'],
  ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun'],
  ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'],
  ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'],
  ['Q1', 'Q2', 'Q3', 'Q4'],
]
function nextInList(v, offset, series) {
  if (!series || typeof v !== 'string') return v
  for (const l of LISTS) {
    const i = l.findIndex((x) => x.toLowerCase() === v.toLowerCase())
    if (i >= 0) return l[(((i + offset) % l.length) + l.length) % l.length]
  }
  return v
}

// ── conditional formats and validation (evaluated by the page; here so they are tested) ─────────────────────────────────

// numbers of a rule's range, with their extremes and mean, computed once per recalculation
function cfStats(book, sh, g) {
  const k = `${sh}:${rangeText(g)}`
  if (!book._cfc || book._cfc.ver !== book.ver) book._cfc = { ver: book.ver, map: new Map() }
  let st = book._cfc.map.get(k)
  if (!st) {
    const all = book.matrix(new Rng(sh, g.r1, g.c1, g.r2, g.c2)).flat()
    const nums = []
    let lo = Infinity
    let hi = -Infinity
    let sum = 0
    for (const v of all) if (typeof v === 'number') (nums.push(v), (lo = Math.min(lo, v)), (hi = Math.max(hi, v)), (sum += v))
    st = { all, nums, lo, hi, mean: nums.length ? sum / nums.length : 0, desc: null, asc: null }
    book._cfc.map.set(k, st)
  }
  return st
}
export const inRange = (g, r, c) => r >= g.r1 && r <= g.r2 && c >= g.c1 && c <= g.c2
// the style a cell gets from its sheet's conditional formats: { bg, fg, b, i, bar, scale }
export function condStyle(book, sh, r, c) {
  const s = book.sheets[sh]
  if (!s.cf.length) return null
  const v = book.get(sh, r, c)
  let out = null
  for (const rule of s.cf) {
    const g = parseRange(rule.range)
    if (!g || !inRange(g, r, c)) continue
    let hit = false
    const num = typeof v === 'number'
    switch (rule.type) {
      case 'cell': {
        const a = rule.a === '' || rule.a == null ? null : Number.isNaN(Number(rule.a)) ? rule.a : Number(rule.a)
        const b = rule.b === '' || rule.b == null ? null : Number.isNaN(Number(rule.b)) ? rule.b : Number(rule.b)
        if (v == null || v instanceof XlErr) break
        const d = cmp(v, a)
        hit = rule.op === '>' ? d > 0 : rule.op === '<' ? d < 0 : rule.op === '>=' ? d >= 0 : rule.op === '<=' ? d <= 0 : rule.op === '=' ? d === 0 : rule.op === '<>' ? d !== 0 : rule.op === 'between' ? d >= 0 && cmp(v, b) <= 0 : rule.op === 'notBetween' ? d < 0 || cmp(v, b) > 0 : false
        break
      }
      case 'text': {
        const t = String(v ?? '').toLowerCase()
        const q = String(rule.text ?? '').toLowerCase()
        hit = rule.op === 'contains' ? t.includes(q) : rule.op === 'notContains' ? !t.includes(q) : rule.op === 'begins' ? t.startsWith(q) : t.endsWith(q)
        if (v == null) hit = rule.op === 'notContains'
        break
      }
      case 'blank':
        hit = v == null || v === ''
        break
      case 'error':
        hit = v instanceof XlErr
        break
      case 'formula': {
        try {
          const f = shiftFormula(String(rule.formula).replace(/^=/, ''), r - g.r1, c - g.c1)
          const res = book.evaluate(f, sh)
          hit = res === true || (typeof res === 'number' && res !== 0)
        } catch {
          hit = false
        }
        break
      }
      case 'dup': {
        if (v == null) break
        hit = cfStats(book, sh, g).all.filter((x) => x != null && cmp(x, v) === 0).length > 1
        break
      }
      case 'top': {
        if (!num) break
        const st = cfStats(book, sh, g)
        const list = rule.bottom ? (st.asc ||= [...st.nums].sort((x, y) => x - y)) : (st.desc ||= [...st.nums].sort((x, y) => y - x))
        const n = rule.percent ? Math.max(1, Math.ceil((list.length * (rule.n || 10)) / 100)) : rule.n || 10
        hit = list.length > 0 && (rule.bottom ? v <= list[Math.min(n, list.length) - 1] : v >= list[Math.min(n, list.length) - 1])
        break
      }
      case 'avg': {
        if (!num) break
        const m = cfStats(book, sh, g).mean
        hit = rule.op === 'below' ? v < m : v > m
        break
      }
      case 'scale': {
        if (!num) break
        const { lo, hi } = cfStats(book, sh, g)
        const t = hi === lo ? 0.5 : (v - lo) / (hi - lo)
        const cols = rule.colors || ['#f8696b', '#ffeb84', '#63be7b']
        out = { ...(out || {}), bg: lerpColors(cols, t) }
        continue
      }
      case 'bar': {
        if (!num) break
        const st = cfStats(book, sh, g)
        const lo = Math.min(0, st.lo)
        const hi = Math.max(0, st.hi)
        out = { ...(out || {}), bar: { color: rule.color || '#5b9bd5', pct: hi === lo ? 0 : Math.abs(v) / Math.max(Math.abs(lo), Math.abs(hi)), neg: v < 0 } }
        continue
      }
      default:
    }
    if (hit) {
      out = { ...(out || {}), ...(rule.style || {}) }
      if (rule.stop) break
    }
  }
  return out
}
const hex = (h) => {
  const m = String(h).replace('#', '')
  return [0, 2, 4].map((i) => Number.parseInt(m.slice(i, i + 2), 16))
}
function lerpColors(cols, t) {
  const n = cols.length - 1
  const i = Math.min(n - 1, Math.floor(t * n))
  const f = t * n - i
  const a = hex(cols[i])
  const b = hex(cols[i + 1])
  return `#${a.map((x, k) => Math.round(x + (b[k] - x) * f).toString(16).padStart(2, '0')).join('')}`
}

// can this value go in this cell? → { ok, message }
export function validate(book, sh, r, c, value) {
  const s = book.sheets[sh]
  for (const rule of s.dv) {
    const g = parseRange(rule.range)
    if (!g || !inRange(g, r, c)) continue
    if (value == null || value === '') return { ok: !!rule.allowBlank || rule.allowBlank === undefined, message: rule.message || 'This cell cannot be empty.' }
    const msg = rule.message || 'That value is not allowed here.'
    const v = typeof value === 'string' && rule.type !== 'list' && rule.type !== 'text' && rule.type !== 'custom' ? (parseTyped(value)?.v ?? value) : value
    let ok = true
    if (rule.type === 'list') {
      const items = listItems(book, sh, rule)
      ok = items.some((x) => cmp(x, value) === 0 || String(x).toLowerCase() === String(value).toLowerCase())
    } else if (rule.type === 'number' || rule.type === 'whole' || rule.type === 'date') {
      ok = typeof v === 'number' && (rule.type !== 'whole' || Number.isInteger(v))
      const a = Number(rule.a)
      const b = Number(rule.b)
      if (ok)
        ok =
          rule.op === 'between' ? v >= a && v <= b : rule.op === 'notBetween' ? v < a || v > b : rule.op === '>' ? v > a : rule.op === '<' ? v < a : rule.op === '>=' ? v >= a : rule.op === '<=' ? v <= a : rule.op === '<>' ? v !== a : v === a
    } else if (rule.type === 'text') {
      const len = String(value).length
      const a = Number(rule.a)
      const b = Number(rule.b)
      ok = rule.op === 'between' ? len >= a && len <= b : rule.op === '<=' ? len <= a : rule.op === '>=' ? len >= a : len === a
    } else if (rule.type === 'custom') {
      try {
        const f = shiftFormula(String(rule.formula).replace(/^=/, ''), r - g.r1, c - g.c1)
        const res = book.evaluate(f, sh)
        ok = res === true || (typeof res === 'number' && res !== 0)
      } catch {
        ok = false
      }
    }
    if (!ok && rule.strict !== false) return { ok: false, message: msg }
  }
  return { ok: true }
}
export function listItems(book, sh, rule) {
  if (Array.isArray(rule.items)) return rule.items
  const src = String(rule.source || '')
  if (!src) return []
  const m = mapRefsFirst(src)
  if (!m) return String(src).split(',').map((x) => x.trim())
  const i = m.sheet != null ? book.sheetIndex(m.sheet) : sh
  if (i < 0) return []
  return book.matrix(new Rng(i, m.r1, m.c1, m.r2, m.c2)).flat().filter((x) => x != null && x !== '')
}
const mapRefsFirst = (s) => {
  try {
    const t = tokenize(String(s).replace(/^=/, ''))
    return t.length === 1 && t[0].t === 'ref' ? t[0].ref : null
  } catch {
    return null
  }
}
