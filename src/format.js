// Number formats and dates: the cell format codes of a spreadsheet (0.00, #,##0, $#,##0.00, 0%, yyyy-mm-dd, [h]:mm, 0.0E+00, "text", sections).
// No dependencies. Shared by the server and the page.

export class XlErr {
  constructor(code) {
    this.code = code
  }
  toString() {
    return this.code
  }
}
export const E = {
  DIV0: new XlErr('#DIV/0!'),
  NA: new XlErr('#N/A'),
  NAME: new XlErr('#NAME?'),
  NULL: new XlErr('#NULL!'),
  NUM: new XlErr('#NUM!'),
  REF: new XlErr('#REF!'),
  VALUE: new XlErr('#VALUE!'),
  SPILL: new XlErr('#SPILL!'),
  CALC: new XlErr('#CALC!'),
  CIRC: new XlErr('#CIRC!'),
  PARSE: new XlErr('#ERROR!'),
}
export const errByCode = (code) => Object.values(E).find((e) => e.code === code) || null

// ── dates: serial numbers, day 1 = 1900-01-01 (with the 1900 leap-year quirk spreadsheets keep) ─────────────────────────
const EPOCH = Date.UTC(1899, 11, 30)
export const toSerial = (y, m, d) => {
  const t = Date.UTC(y, m - 1, d)
  const n = Math.round((t - EPOCH) / 86400000)
  return n
}
export const fromSerial = (n) => {
  const whole = Math.floor(n)
  const dt = new Date(EPOCH + whole * 86400000)
  const frac = n - whole
  const secs = Math.round(frac * 86400)
  return {
    y: dt.getUTCFullYear(),
    m: dt.getUTCMonth() + 1,
    d: dt.getUTCDate(),
    wd: dt.getUTCDay(),
    h: Math.floor(secs / 3600) % 24,
    mi: Math.floor((secs % 3600) / 60),
    s: secs % 60,
  }
}
export const nowSerial = () => {
  const n = new Date()
  return (Date.now() - n.getTimezoneOffset() * 60000 - EPOCH) / 86400000
}

const MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December']
const DAYS = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday']
export const monthNames = MONTHS
const pad = (n, w = 2) => String(n).padStart(w, '0')

// ── General: up to 11 significant digits like the spreadsheet shows, 15 when a function turns a number into text ─────────
export function general(v, digits = 15) {
  if (!Number.isFinite(v)) return String(v)
  if (v === 0) return '0'
  const a = Math.abs(v)
  if (a >= 1e21 || a < 1e-9) return v.toExponential(Math.min(digits - 1, 5)).replace(/\.?0+e/, 'E').replace('e', 'E')
  const s = String(+v.toPrecision(digits))
  if (/e/.test(s)) return Number(s).toFixed(20).replace(/\.?0+$/, '')
  return s
}
const display = (v) => {
  // what a General cell shows: it rounds to fit, here to 10 significant digits
  if (!Number.isFinite(v)) return String(v)
  const a = Math.abs(v)
  if (a !== 0 && (a >= 1e11 || a < 1e-9)) return v.toExponential(5).replace(/\.?0+e/, 'E').replace('e', 'E')
  return String(+v.toPrecision(10)).replace(/e.*/, (m) => m)
}

// ── format codes ─────────────────────────────────────────────────────────────────────────────────────────────────────────
function splitSections(fmt) {
  const out = []
  let cur = ''
  let q = false
  let br = false
  for (let i = 0; i < fmt.length; i++) {
    const ch = fmt[i]
    if (ch === '\\') {
      cur += ch + (fmt[++i] || '')
      continue
    }
    if (ch === '"') q = !q
    else if (!q && ch === '[') br = true
    else if (!q && ch === ']') br = false
    if (ch === ';' && !q && !br) {
      out.push(cur)
      cur = ''
    } else cur += ch
  }
  out.push(cur)
  return out
}

// [Red] [Blue] [>100] [$€-407] at the start of a section
function sectionMeta(sec) {
  let color = ''
  let cond = null
  const rest = sec.replace(/\[([^\]]*)\]/g, (m, inner) => {
    if (/^(black|blue|cyan|green|magenta|red|white|yellow)$/i.test(inner)) {
      color = inner.toLowerCase()
      return ''
    }
    const c = inner.match(/^([<>=]+)\s*(-?[\d.]+)$/)
    if (c) {
      cond = { op: c[1], n: Number(c[2]) }
      return ''
    }
    const cur = inner.match(/^\$([^-\]]*)(?:-.*)?$/)
    if (cur) return cur[1] ? `"${cur[1]}"` : ''
    if (/^h+$|^m+$|^s+$/i.test(inner)) return m // elapsed time, kept
    return ''
  })
  return { text: rest, color, cond }
}

const DATE_TOKEN = /\[h+\]|\[m+\]|\[s+\]|yyyy|yy|mmmmm|mmmm|mmm|mm|m|dddd|ddd|dd|d|hh|h|ss|s|AM\/PM|am\/pm|A\/P|a\/p|\.0+|"[^"]*"|\\.|./gi
const isDateSection = (t) => {
  const stripped = t.replace(/"[^"]*"|\\./g, '').replace(/\[[^\]]*\]/g, (m) => (/^\[[hms]+\]$/i.test(m) ? 'h' : ''))
  return /[ymdhs]|AM\/PM|A\/P/i.test(stripped) && !/^[#0?.,%\s$€£()+-]*$/.test(stripped) && !/[#0]/.test(stripped.replace(/[ymdhs]/gi, ''))
}

function formatDate(n, text) {
  if (n < 0) return '#NUM!'
  const t = fromSerial(n)
  const toks = text.match(DATE_TOKEN) || []
  const ampm = toks.some((x) => /^(AM\/PM|A\/P)$/i.test(x))
  const sig = toks.filter((x) => /^[a-z\[]/i.test(x) && !/^"/.test(x))
  let out = ''
  for (let i = 0; i < toks.length; i++) {
    const x = toks[i]
    const l = x.toLowerCase()
    if (x[0] === '"') out += x.slice(1, -1)
    else if (x[0] === '\\') out += x[1] || ''
    else if (l === 'yyyy') out += pad(t.y, 4)
    else if (l === 'yy') out += pad(t.y % 100)
    else if (l === 'mmmmm') out += MONTHS[t.m - 1][0]
    else if (l === 'mmmm') out += MONTHS[t.m - 1]
    else if (l === 'mmm') out += MONTHS[t.m - 1].slice(0, 3)
    else if (l === 'mm' || l === 'm') {
      // a month, unless it follows hours or comes before seconds
      const at = sig.indexOf(x, 0)
      let minute = false
      for (let j = i - 1; j >= 0; j--) {
        if (/^\[?h+\]?$/i.test(toks[j])) minute = true
        if (/^[a-z]/i.test(toks[j]) && !/^h/i.test(toks[j]) && !/^\[h/i.test(toks[j])) break
        if (/^\[?h/i.test(toks[j])) break
      }
      for (let j = i + 1; j < toks.length && !minute; j++) {
        if (/^s+$/i.test(toks[j])) minute = true
        if (/^[a-z]/i.test(toks[j])) break
      }
      void at
      out += minute ? (l === 'mm' ? pad(t.mi) : t.mi) : l === 'mm' ? pad(t.m) : t.m
    } else if (l === 'dddd') out += DAYS[t.wd]
    else if (l === 'ddd') out += DAYS[t.wd].slice(0, 3)
    else if (l === 'dd') out += pad(t.d)
    else if (l === 'd') out += t.d
    else if (l === 'hh') out += pad(ampm ? t.h % 12 || 12 : t.h)
    else if (l === 'h') out += ampm ? t.h % 12 || 12 : t.h
    else if (/^\[h+\]$/.test(l)) out += Math.floor(n * 24)
    else if (/^\[m+\]$/.test(l)) out += Math.floor(n * 1440)
    else if (/^\[s+\]$/.test(l)) out += Math.floor(n * 86400)
    else if (l === 'ss') out += pad(t.s)
    else if (l === 's') out += t.s
    else if (l === 'am/pm') out += t.h < 12 ? 'AM' : 'PM'
    else if (l === 'a/p') out += t.h < 12 ? 'A' : 'P'
    else if (/^\.0+$/.test(x)) out += `.${pad(Math.round((n % 1) * 86400000) % 1000, 3).slice(0, x.length - 1)}`
    else out += x
  }
  return out
}

const addCommas = (s) => s.replace(/\B(?=(\d{3})+(?!\d))/g, ',')

function formatNumber(n, text, { sign }) {
  // percent signs outside quotes scale the number
  const bare = text.replace(/"[^"]*"|\\./g, '')
  let v = Math.abs(n)
  v *= 100 ** (bare.match(/%/g) || []).length
  // split into prefix, the digit pattern, suffix
  const lit = (s) => {
    let o = ''
    for (let i = 0; i < s.length; i++) {
      const ch = s[i]
      if (ch === '"') {
        const j = s.indexOf('"', i + 1)
        o += s.slice(i + 1, j < 0 ? s.length : j)
        i = j < 0 ? s.length : j
      } else if (ch === '\\') o += s[++i] || ''
      else if (ch === '_') {
        i++
        o += ' '
      } else if (ch === '*') i++
      else o += ch
    }
    return o
  }
  // mask quoted text so its characters are not taken for digit placeholders
  const masked = text.replace(/"[^"]*"|\\./g, (m) => '\u0000'.repeat(m.length))
  const first = masked.search(/[0#?.]/)
  if (first < 0) return (sign < 0 ? '-' : '') + lit(text)
  // the digit pattern: digits, point, commas, then an optional exponent
  const run = masked.slice(first).match(/^[0#?.,]+(?:E[+-]0+)?/i)[0]
  let pre = text.slice(0, first)
  let core = text.slice(first, first + run.length)
  let post = text.slice(first + run.length)
  const sci = core.match(/E([+-])(0+)$/i)
  if (sci) core = core.slice(0, sci.index)
  // commas after the last digit placeholder scale by 1000 each
  let scale = 0
  const trail = core.match(/,+$/)
  if (trail) {
    scale = trail[0].length
    core = core.slice(0, -scale)
  }
  const dot = core.indexOf('.')
  const intPat = dot < 0 ? core : core.slice(0, dot)
  const decPat = dot < 0 ? '' : core.slice(dot + 1)
  const thousands = /,/.test(intPat.replace(/,+$/, ''))
  const intDigits = intPat.replace(/,/g, '')
  const decDigits = decPat.replace(/,/g, '')
  v /= 1000 ** scale
  let body
  const minInt = (intDigits.match(/0/g) || []).length
  const decN = decDigits.length
  if (sci) {
    let exp = v === 0 ? 0 : Math.floor(Math.log10(v))
    const per = Math.max(1, intDigits.length)
    exp = Math.floor(exp / per) * per
    let mant = v / 10 ** exp
    mant = Number(mant.toFixed(decN))
    if (mant >= 10 ** per) {
      mant /= 10 ** per
      exp += per
    }
    let ms = mant.toFixed(decN)
    ms = ms.padStart(minInt + (decN ? decN + 1 : 0), '0')
    body = `${ms}E${exp < 0 ? '-' : sci[1] === '+' ? '+' : ''}${pad(Math.abs(exp), sci[2].length)}`
  } else {
    let s = roundTo(v, decN).toFixed(decN)
    let [ip, dp = ''] = s.split('.')
    if (ip === '0' && minInt === 0) ip = ''
    ip = ip.padStart(minInt, '0')
    // # and ? in the decimals: drop trailing zeros that # allows
    let d = dp
    for (let i = decN - 1; i >= 0 && decDigits[i] !== '0'; i--) if (d[i] === '0') d = d.slice(0, i) + (decDigits[i] === '?' ? ' ' : '\u0001') + d.slice(i + 1)
    d = d.replace(/\u0001/g, '')
    if (decN && d.trim() === '' && /^[#?]+$/.test(decDigits)) d = ''
    if (thousands) ip = addCommas(ip)
    s = ip + (decN && d !== '' ? `.${d}` : decN && d === '' && /0/.test(decDigits) ? '' : '')
    // keep the decimal point only when digits follow, or the pattern forces one
    body = s
    if (decN && d === '' && !/^[#?]+$/.test(decDigits)) body = `${ip}.${dp}`
    // ?-placeholder alignment is not drawn; a trailing space stays out
    body = body.replace(/ +$/, '')
  }
  void pre
  pre = lit(pre)
  post = lit(post)
  // a bare "0" pattern of a negative-section value keeps the pattern's own signs
  return (sign < 0 ? '-' : '') + pre + body + post
}
const roundTo = (v, d) => {
  const f = 10 ** d
  return Math.round((v + Number.EPSILON * v) * f) / f
}

// Format any value with a format code ('General' or empty means the default).
// Returns { text, color } when you need the colour; formatValue returns only the text.
export function formatFull(v, fmt) {
  if (v == null) return { text: '', color: '' }
  if (v instanceof XlErr) return { text: v.code, color: '' }
  if (typeof v === 'boolean') return { text: v ? 'TRUE' : 'FALSE', color: '' }
  const f = String(fmt || '').trim()
  if (typeof v === 'string') {
    if (!f || f === 'General') return { text: v, color: '' }
    const secs = splitSections(f)
    const t = secs[3] ?? (secs.length === 1 && /@/.test(secs[0]) ? secs[0] : null)
    if (t == null) return { text: v, color: '' }
    const m = sectionMeta(t)
    return { text: m.text.replace(/"([^"]*)"|\\(.)|@/g, (x, q, e) => (q != null ? q : e != null ? e : v)), color: m.color }
  }
  if (!Number.isFinite(v)) return { text: Number.isNaN(v) ? '#NUM!' : '#NUM!', color: '' }
  if (!f || f === 'General') return { text: display(v), color: '' }
  const secs = splitSections(f).map(sectionMeta)
  let pick = secs[0]
  let sign = 1
  let n = v
  if (secs.length >= 2) {
    // conditions first, otherwise positive; negative; zero
    const hit = secs.slice(0, 2).findIndex((s) => s.cond && cmpCond(v, s.cond))
    if (hit >= 0) pick = secs[hit]
    else if (secs[0].cond && !cmpCond(v, secs[0].cond)) pick = secs[1]
    else if (v < 0) {
      pick = secs[1]
      n = -v
      sign = 1
    } else if (v === 0 && secs.length >= 3) pick = secs[2]
    if (secs.length >= 3 && v === 0 && !secs.some((s) => s.cond)) pick = secs[2]
    if (secs.length >= 2 && v < 0 && pick === secs[1]) {
      n = -v
      sign = 1
    }
  } else if (v < 0) {
    sign = -1
    n = -v
  }
  if (isDateSection(pick.text)) return { text: formatDate(v, pick.text), color: pick.color }
  const text = formatNumber(n, pick.text.replace(/@/g, ''), { sign })
  return { text: text === '-' ? '' : text, color: pick.color }
}
const cmpCond = (v, { op, n }) => (op === '>' ? v > n : op === '<' ? v < n : op === '>=' ? v >= n : op === '<=' ? v <= n : op === '<>' ? v !== n : v === n)
export const formatValue = (v, fmt) => formatFull(v, fmt).text

// ── what a person types: a number, a percent, money, a date, a time ─────────────────────────────────────────────────────
const CURRENCY = /^([-+(]?)\s*([$€£¥₹])\s*([-+]?[\d,]*\.?\d+)\s*\)?$/
export function parseTyped(text) {
  const t = String(text).trim()
  if (t === '') return null
  if (/^[-+]?(\d+\.?\d*|\.\d+)([eE][-+]?\d+)?$/.test(t)) return { v: Number(t) }
  if (/^[-+]?\d{1,3}(,\d{3})+(\.\d+)?$/.test(t)) return { v: Number(t.replace(/,/g, '')), nf: '#,##0' + (/\./.test(t) ? `.${'0'.repeat(t.split('.')[1].length)}` : '') }
  const pct = t.match(/^([-+]?[\d,]*\.?\d+)\s*%$/)
  if (pct) {
    const d = (pct[1].split('.')[1] || '').length
    return { v: Number(pct[1].replace(/,/g, '')) / 100, nf: d ? `0.${'0'.repeat(d)}%` : '0%' }
  }
  const cur = t.match(CURRENCY)
  if (cur) {
    const num = Number(cur[3].replace(/,/g, ''))
    const neg = cur[1] === '-' || cur[1] === '(' ? -1 : 1
    const d = (cur[3].split('.')[1] || '').length
    return { v: num * neg, nf: `${cur[2]}#,##0${d ? `.${'0'.repeat(d)}` : cur[3].includes(',') || num >= 1000 ? '' : ''}` }
  }
  let m = t.match(/^(\d{4})-(\d{1,2})-(\d{1,2})(?:[ T](\d{1,2}):(\d{2})(?::(\d{2}))?)?$/)
  if (m) return dateOf(+m[1], +m[2], +m[3], m[4], m[5], m[6], 'yyyy-mm-dd')
  m = t.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?$/)
  if (m) return dateOf(+m[3] < 100 ? 2000 + +m[3] : +m[3], +m[1], +m[2], m[4], m[5], m[6], 'm/d/yyyy')
  m = t.match(/^(\d{1,2})[ -]([A-Za-z]{3,9})[ -,]*(\d{2,4})$/)
  if (m) {
    const mo = MONTHS.findIndex((x) => x.toLowerCase().startsWith(m[2].toLowerCase().slice(0, 3)))
    if (mo >= 0) return dateOf(+m[3] < 100 ? 2000 + +m[3] : +m[3], mo + 1, +m[1], null, null, null, 'd-mmm-yy')
  }
  m = t.match(/^(\d{1,2}):(\d{2})(?::(\d{2}))?\s*(AM|PM)?$/i)
  if (m) {
    let h = +m[1]
    if (m[4]) h = (h % 12) + (/pm/i.test(m[4]) ? 12 : 0)
    if (h < 24 && +m[2] < 60) return { v: (h * 3600 + +m[2] * 60 + +(m[3] || 0)) / 86400, nf: m[4] ? 'h:mm AM/PM' : m[3] ? 'hh:mm:ss' : 'hh:mm' }
  }
  return null
}
function dateOf(y, mo, d, h, mi, s, nf) {
  if (mo < 1 || mo > 12 || d < 1 || d > 31) return null
  const dt = new Date(Date.UTC(y, mo - 1, d))
  if (dt.getUTCMonth() !== mo - 1) return null
  let v = toSerial(y, mo, d)
  if (h != null && mi != null) {
    v += (+h * 3600 + +mi * 60 + +(s || 0)) / 86400
    nf = `${nf} hh:mm${s ? ':ss' : ''}`
  }
  return { v, nf }
}
// the format a typed or formula result suggests when the cell has none: dates stay dates
export const isDateFormat = (fmt) => !!fmt && isDateSection(splitSections(String(fmt))[0] || '')
