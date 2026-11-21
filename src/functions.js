// The function library: about 230 spreadsheet functions registered into FN (see core.js for how they are called).
import { E, XlErr, formatValue, fromSerial, general, monthNames, nowSerial, parseTyped, toSerial } from './format.js'
import { FN, MAXC, MAXR, Rng, addr, broadcast, cmp, colName, ev, isArr, parse, parseRange, toBool, toMat, toNum, toStr } from './core.js'

// ── registration ─────────────────────────────────────────────────────────────────────────────────────────────────────────
const def = (names, f, opt = {}) => {
  for (const n of [].concat(names)) FN[n] = { f, ...opt }
}
// a scalar function: given arrays or ranges it works element by element
const S = (names, f) =>
  def(names, (args, x) => {
    if (args.some((a) => a instanceof Rng || isArr(a))) {
      const book = x.book
      return broadcast(
        args.map((a) => (a instanceof Rng ? (a.r1 === a.r2 && a.c1 === a.c2 ? book.get(a.sh, a.r1, a.c1) : book.matrix(a)) : a)),
        (els) => f(...els),
      )
    }
    return f(...args)
  })
const lazy = (names, f) => def(names, f, { lazy: true })
const raw = (names, f) => def(names, f, { raw: true })

const one = (v, x) => (v instanceof Rng ? (v.r1 === v.r2 && v.c1 === v.c2 ? x.book.get(v.sh, v.r1, v.c1) : thrower(E.VALUE)) : isArr(v) ? v[0][0] : v)
const thrower = (e) => {
  throw e
}
const M = (v, x) => (v instanceof Rng ? x.book.matrix(v) : isArr(v) ? v : [[v]])
const flat = (v, x) => M(v, x).flat()
const err = (v) => {
  if (v instanceof XlErr) throw v
  return v
}
const opt = (v, d) => (v === undefined || v === null ? d : v)
const int = (v) => Math.trunc(toNum(v))
const numArg = (v, x, d) => (v === undefined ? d : toNum(one(v, x)))

// numbers of a list of arguments, the way SUM and friends read them: cells in ranges count if numeric, typed values are coerced
function* nums(args, x, { strict = false } = {}) {
  for (const a of args) {
    if (a instanceof Rng || isArr(a)) {
      for (const row of M(a, x))
        for (const v of row) {
          if (v instanceof XlErr) throw v
          if (typeof v === 'number') yield v
          else if (strict && typeof v === 'boolean') yield v ? 1 : 0
        }
    } else if (a instanceof XlErr) throw a
    else if (a === null || a === undefined) continue
    else yield toNum(a)
  }
}
const list = (args, x) => [...nums(args, x)]
const sum = (a) => a.reduce((p, q) => p + q, 0)

// ── criteria (COUNTIF, SUMIF…) ───────────────────────────────────────────────────────────────────────────────────────────
const wild = (s) => new RegExp(`^${s.replace(/~([*?~])|[.+^${}()|[\]\\]|(\*)|(\?)/g, (m, lit, star, q) => (lit ? `\\${lit}` : star ? '[\\s\\S]*' : q ? '[\\s\\S]' : `\\${m}`))}$`, 'i')
export function criterion(c) {
  if (c instanceof XlErr) return (v) => v instanceof XlErr && v.code === c.code
  if (typeof c === 'number') return (v) => (typeof v === 'number' ? v === c : typeof v === 'string' && parseTyped(v)?.v === c)
  if (typeof c === 'boolean') return (v) => v === c
  if (c == null) return (v) => v == null || v === ''
  const m = String(c).match(/^(<=|>=|<>|<|>|=)?([\s\S]*)$/)
  const op = m[1] || '='
  const rhs = m[2]
  const n = rhs.trim() === '' ? null : parseTyped(rhs.trim())?.v
  if (typeof n === 'number') {
    return (v) => {
      let a = v
      if (typeof a === 'string') a = op === '=' || op === '<>' ? parseTyped(a.trim())?.v ?? a : undefined
      if (typeof a !== 'number') return op === '<>'
      return op === '=' ? a === n : op === '<>' ? a !== n : op === '<' ? a < n : op === '>' ? a > n : op === '<=' ? a <= n : a >= n
    }
  }
  if (/^(true|false)$/i.test(rhs) && (op === '=' || op === '<>')) return (v) => (typeof v === 'boolean' && v === (rhs.toLowerCase() === 'true')) === (op === '=')
  if (rhs === '') return op === '=' ? (v) => v == null || v === '' : op === '<>' ? (v) => !(v == null || v === '') : () => false
  if ((op === '=' || op === '<>') && /[*?]/.test(rhs)) {
    const re = wild(rhs)
    return (v) => (typeof v === 'string' && re.test(v)) === (op === '=')
  }
  const low = rhs.toLowerCase()
  return (v) => {
    if (op === '<>') return !(typeof v === 'string' && v.toLowerCase() === low)
    if (typeof v !== 'string') return false
    const a = v.toLowerCase()
    return op === '=' ? a === low : op === '<' ? a < low : op === '>' ? a > low : op === '<=' ? a <= low : a >= low
  }
}
// pairs of (range, criterion) → the flat positions where every pair matches
function matching(pairs, x) {
  let keep = null
  let size = -1
  for (let i = 0; i < pairs.length; i += 2) {
    const cells = flat(pairs[i], x)
    if (size >= 0 && cells.length !== size) throw E.VALUE
    size = cells.length
    const p = criterion(one(pairs[i + 1], x))
    const ok = cells.map((v) => p(v))
    keep = keep ? keep.map((k, j) => k && ok[j]) : ok
  }
  return keep || []
}

// ── math ─────────────────────────────────────────────────────────────────────────────────────────────────────────────────
def('SUM', (a, x) => sum(list(a, x)))
def('PRODUCT', (a, x) => list(a, x).reduce((p, q) => p * q, 1))
def('SUMSQ', (a, x) => sum(list(a, x).map((v) => v * v)))
def('AVERAGE', (a, x) => {
  const l = list(a, x)
  if (!l.length) throw E.DIV0
  return sum(l) / l.length
})
def('AVERAGEA', (a, x) => {
  const l = a.flatMap((v) => flat(v, x)).map((v) => (v instanceof XlErr ? err(v) : typeof v === 'number' ? v : typeof v === 'boolean' ? (v ? 1 : 0) : typeof v === 'string' ? 0 : null)).filter((v) => v !== null)
  if (!l.length) throw E.DIV0
  return sum(l) / l.length
})
def('COUNT', (a, x) => {
  let n = 0
  for (const v of a) {
    if (v instanceof Rng || isArr(v)) n += flat(v, x).filter((q) => typeof q === 'number').length
    else if (typeof v === 'number' || (typeof v === 'string' && parseTyped(v.trim())) || typeof v === 'boolean') n++
  }
  return n
})
def('COUNTA', (a, x) => a.reduce((n, v) => n + (v instanceof Rng || isArr(v) ? flat(v, x).filter((q) => q != null && q !== '').length : v == null ? 0 : 1), 0))
def('COUNTBLANK', (a, x) => a.reduce((n, v) => n + flat(v, x).filter((q) => q == null || q === '').length, 0))
def('MIN', (a, x) => {
  const l = list(a, x)
  return l.length ? Math.min(...l) : 0
})
def('MAX', (a, x) => {
  const l = list(a, x)
  return l.length ? Math.max(...l) : 0
})
def('MEDIAN', (a, x) => {
  const l = list(a, x).sort((p, q) => p - q)
  if (!l.length) throw E.NUM
  const h = l.length >> 1
  return l.length % 2 ? l[h] : (l[h - 1] + l[h]) / 2
})
def(['MODE', 'MODE.SNGL'], (a, x) => {
  const l = list(a, x)
  const c = new Map()
  for (const v of l) c.set(v, (c.get(v) || 0) + 1)
  let best = null
  let n = 1
  for (const [v, k] of c) if (k > n) (best = v), (n = k)
  if (best === null) throw E.NA
  return best
})
const variance = (l, pop) => {
  if (l.length < (pop ? 1 : 2)) throw E.DIV0
  const m = sum(l) / l.length
  return sum(l.map((v) => (v - m) ** 2)) / (l.length - (pop ? 0 : 1))
}
def(['VAR', 'VAR.S'], (a, x) => variance(list(a, x), false))
def(['VARP', 'VAR.P'], (a, x) => variance(list(a, x), true))
def(['STDEV', 'STDEV.S'], (a, x) => Math.sqrt(variance(list(a, x), false)))
def(['STDEVP', 'STDEV.P'], (a, x) => Math.sqrt(variance(list(a, x), true)))
def('AVEDEV', (a, x) => {
  const l = list(a, x)
  const m = sum(l) / l.length
  return sum(l.map((v) => Math.abs(v - m))) / l.length
})
def('GEOMEAN', (a, x) => {
  const l = list(a, x)
  if (!l.length || l.some((v) => v <= 0)) throw E.NUM
  return Math.exp(sum(l.map(Math.log)) / l.length)
})
def('LARGE', ([r, k], x) => {
  const l = list([r], x).sort((p, q) => q - p)
  const i = int(one(k, x))
  if (i < 1 || i > l.length) throw E.NUM
  return l[i - 1]
})
def('SMALL', ([r, k], x) => {
  const l = list([r], x).sort((p, q) => p - q)
  const i = int(one(k, x))
  if (i < 1 || i > l.length) throw E.NUM
  return l[i - 1]
})
def(['RANK', 'RANK.EQ'], ([v, r, o], x) => {
  const n = toNum(one(v, x))
  const l = list([r], x)
  if (!l.includes(n)) throw E.NA
  const asc = o !== undefined && toNum(one(o, x)) !== 0
  return 1 + l.filter((q) => (asc ? q < n : q > n)).length
})
const pctile = (l, k) => {
  l.sort((p, q) => p - q)
  const pos = (l.length - 1) * k
  const lo = Math.floor(pos)
  return l[lo] + (l[Math.min(lo + 1, l.length - 1)] - l[lo]) * (pos - lo)
}
def(['PERCENTILE', 'PERCENTILE.INC'], ([r, k], x) => {
  const p = toNum(one(k, x))
  if (p < 0 || p > 1) throw E.NUM
  return pctile(list([r], x), p)
})
def(['QUARTILE', 'QUARTILE.INC'], ([r, q], x) => {
  const n = int(one(q, x))
  if (n < 0 || n > 4) throw E.NUM
  return pctile(list([r], x), n / 4)
})
const pairs = (a, b, x) => {
  const p = flat(a, x)
  const q = flat(b, x)
  if (p.length !== q.length) throw E.NA
  const out = []
  p.forEach((v, i) => typeof v === 'number' && typeof q[i] === 'number' && out.push([v, q[i]]))
  return out
}
const regress = (ys, xs, x) => {
  const d = pairs(ys, xs, x)
  const n = d.length
  if (n < 2) throw E.DIV0
  const my = sum(d.map((p) => p[0])) / n
  const mx = sum(d.map((p) => p[1])) / n
  const sxy = sum(d.map((p) => (p[0] - my) * (p[1] - mx)))
  const sxx = sum(d.map((p) => (p[1] - mx) ** 2))
  const syy = sum(d.map((p) => (p[0] - my) ** 2))
  if (sxx === 0) throw E.DIV0
  return { slope: sxy / sxx, intercept: my - (sxy / sxx) * mx, r: sxy / Math.sqrt(sxx * syy), mx, my }
}
def('SLOPE', ([y, xs], x) => regress(y, xs, x).slope)
def('INTERCEPT', ([y, xs], x) => regress(y, xs, x).intercept)
def('CORREL', ([a, b], x) => regress(a, b, x).r)
def('RSQ', ([y, xs], x) => regress(y, xs, x).r ** 2)
def('FORECAST', ([v, y, xs], x) => {
  const g = regress(y, xs, x)
  return g.intercept + g.slope * toNum(one(v, x))
})
def('SUMPRODUCT', (a, x) => {
  const ms = a.map((v) => M(v, x))
  const h = ms[0].length
  const w = ms[0][0].length
  let t = 0
  for (const m of ms) if (m.length !== h || m[0].length !== w) throw E.VALUE
  for (let r = 0; r < h; r++)
    for (let c = 0; c < w; c++) {
      let p = 1
      for (const m of ms) {
        const v = m[r][c]
        if (v instanceof XlErr) throw v
        p *= typeof v === 'number' ? v : typeof v === 'boolean' ? (v ? 1 : 0) : 0
      }
      t += p
    }
  return t
})
def('SUMIF', ([r, c, s], x) => {
  const cells = flat(r, x)
  const p = criterion(one(c, x))
  const vals = s === undefined ? cells : flat(s, x)
  return sum(cells.map((v, i) => (p(v) && typeof vals[i] === 'number' ? vals[i] : 0)))
})
def('SUMIFS', ([s, ...rest], x) => {
  const vals = flat(s, x)
  const keep = matching(rest, x)
  if (keep.length !== vals.length) throw E.VALUE
  return sum(vals.map((v, i) => (keep[i] && typeof v === 'number' ? v : 0)))
})
def('COUNTIF', ([r, c], x) => {
  const p = criterion(one(c, x))
  return flat(r, x).filter((v) => p(v)).length
})
def('COUNTIFS', (a, x) => matching(a, x).filter(Boolean).length)
def('AVERAGEIF', ([r, c, s], x) => {
  const cells = flat(r, x)
  const p = criterion(one(c, x))
  const vals = s === undefined ? cells : flat(s, x)
  const l = cells.map((v, i) => (p(v) && typeof vals[i] === 'number' ? vals[i] : null)).filter((v) => v !== null)
  if (!l.length) throw E.DIV0
  return sum(l) / l.length
})
def('AVERAGEIFS', ([s, ...rest], x) => {
  const vals = flat(s, x)
  const keep = matching(rest, x)
  const l = vals.filter((v, i) => keep[i] && typeof v === 'number')
  if (!l.length) throw E.DIV0
  return sum(l) / l.length
})
const ifs = (pick) => ([s, ...rest], x) => {
  const vals = flat(s, x)
  const keep = matching(rest, x)
  const l = vals.filter((v, i) => keep[i] && typeof v === 'number')
  return l.length ? pick(l) : 0
}
def('MAXIFS', ifs((l) => Math.max(...l)))
def('MINIFS', ifs((l) => Math.min(...l)))
const AGG = { 1: 'AVERAGE', 2: 'COUNT', 3: 'COUNTA', 4: 'MAX', 5: 'MIN', 6: 'PRODUCT', 7: 'STDEV', 8: 'STDEVP', 9: 'SUM', 10: 'VAR', 11: 'VARP' }
def('SUBTOTAL', ([n, ...rest], x) => {
  const k = int(one(n, x)) % 100
  const name = AGG[k]
  if (!name) throw E.VALUE
  return FN[name].f(rest, x)
})
def('AGGREGATE', ([n, , ...rest], x) => {
  const k = int(one(n, x))
  const names = { ...AGG, 12: 'MEDIAN', 14: 'LARGE', 15: 'SMALL' }
  if (!names[k]) throw E.VALUE
  return FN[names[k]].f(rest, x)
})

S('ABS', (a) => Math.abs(toNum(a)))
S('SIGN', (a) => Math.sign(toNum(a)))
S('SQRT', (a) => {
  const v = toNum(a)
  if (v < 0) throw E.NUM
  return Math.sqrt(v)
})
S('EXP', (a) => Math.exp(toNum(a)))
S('LN', (a) => {
  const v = toNum(a)
  if (v <= 0) throw E.NUM
  return Math.log(v)
})
S('LOG10', (a) => {
  const v = toNum(a)
  if (v <= 0) throw E.NUM
  return Math.log10(v)
})
S('LOG', (a, b) => {
  const v = toNum(a)
  const base = b === undefined ? 10 : toNum(b)
  if (v <= 0 || base <= 0 || base === 1) throw E.NUM
  return Math.log(v) / Math.log(base)
})
S('POWER', (a, b) => FN.__pow(toNum(a), toNum(b)))
FN.__pow = (p, q) => {
  if (p === 0 && q < 0) throw E.DIV0
  const r = p ** q
  if (!Number.isFinite(r)) throw E.NUM
  return r
}
S('MOD', (a, b) => {
  const n = toNum(a)
  const d = toNum(b)
  if (d === 0) throw E.DIV0
  return n - d * Math.floor(n / d)
})
S('QUOTIENT', (a, b) => {
  const d = toNum(b)
  if (d === 0) throw E.DIV0
  return Math.trunc(toNum(a) / d)
})
S('INT', (a) => Math.floor(toNum(a)))
const rnd = (v, d, mode) => {
  const f = 10 ** Math.trunc(d)
  const t = v * f
  const r = mode === 'up' ? Math.sign(t) * Math.ceil(Math.abs(t) - 1e-12) : mode === 'down' ? Math.trunc(t) : Math.sign(t) * Math.round(Math.abs(t) + 1e-9)
  return r / f
}
S('ROUND', (a, d) => rnd(toNum(a), d === undefined ? 0 : toNum(d), 'near'))
S('ROUNDUP', (a, d) => rnd(toNum(a), d === undefined ? 0 : toNum(d), 'up'))
S('ROUNDDOWN', (a, d) => rnd(toNum(a), d === undefined ? 0 : toNum(d), 'down'))
S('TRUNC', (a, d) => rnd(toNum(a), d === undefined ? 0 : toNum(d), 'down'))
S('MROUND', (a, m) => {
  const n = toNum(a)
  const k = toNum(m)
  if (k === 0) return 0
  if (n * k < 0) throw E.NUM
  return Math.round(n / k) * k
})
S(['CEILING', 'CEILING.MATH'], (a, s) => {
  const n = toNum(a)
  const k = s === undefined ? 1 : toNum(s)
  return k === 0 ? 0 : Math.ceil(n / k) * k
})
S(['FLOOR', 'FLOOR.MATH'], (a, s) => {
  const n = toNum(a)
  const k = s === undefined ? 1 : toNum(s)
  return k === 0 ? 0 : Math.floor(n / k) * k
})
S('EVEN', (a) => {
  const n = toNum(a)
  const r = 2 * Math.ceil(Math.abs(n) / 2)
  return n < 0 ? -r : r
})
S('ODD', (a) => {
  const n = toNum(a)
  const r = 2 * Math.ceil((Math.abs(n) + 1) / 2) - 1
  return n < 0 ? -r : r
})
const fact = (n) => {
  if (n < 0 || n > 170) throw E.NUM
  let r = 1
  for (let i = 2; i <= n; i++) r *= i
  return r
}
S('FACT', (a) => fact(Math.trunc(toNum(a))))
S('COMBIN', (n, k) => {
  const a = Math.trunc(toNum(n))
  const b = Math.trunc(toNum(k))
  if (a < 0 || b < 0 || b > a) throw E.NUM
  return Math.round(fact(a) / (fact(b) * fact(a - b)))
})
S('PERMUT', (n, k) => {
  const a = Math.trunc(toNum(n))
  const b = Math.trunc(toNum(k))
  if (a < 0 || b < 0 || b > a) throw E.NUM
  return Math.round(fact(a) / fact(a - b))
})
const gcd = (a, b) => (b ? gcd(b, a % b) : a)
def('GCD', (a, x) => list(a, x).map(Math.trunc).reduce(gcd, 0))
def('LCM', (a, x) => list(a, x).map(Math.trunc).reduce((p, q) => (p && q ? (p * q) / gcd(p, q) : 0), 1))
def('PI', () => Math.PI)
for (const [n, f] of Object.entries({ SIN: Math.sin, COS: Math.cos, TAN: Math.tan, ASIN: Math.asin, ACOS: Math.acos, ATAN: Math.atan, SINH: Math.sinh, COSH: Math.cosh, TANH: Math.tanh }))
  S(n, (a) => {
    const r = f(toNum(a))
    if (Number.isNaN(r)) throw E.NUM
    return r
  })
S('ATAN2', (a, b) => Math.atan2(toNum(b), toNum(a)))
S('DEGREES', (a) => (toNum(a) * 180) / Math.PI)
S('RADIANS', (a) => (toNum(a) * Math.PI) / 180)
def('RAND', () => Math.random())
def('RANDBETWEEN', ([a, b], x) => {
  const lo = Math.ceil(toNum(one(a, x)))
  const hi = Math.floor(toNum(one(b, x)))
  if (hi < lo) throw E.NUM
  return lo + Math.floor(Math.random() * (hi - lo + 1))
})
S('BASE', (a, r, m) => toNum(a).toString(toNum(r)).toUpperCase().padStart(m === undefined ? 0 : toNum(m), '0'))
S('DECIMAL', (a, r) => {
  const v = Number.parseInt(toStr(a), toNum(r))
  if (Number.isNaN(v)) throw E.NUM
  return v
})
S('ROMAN', (a) => {
  let n = Math.trunc(toNum(a))
  if (n < 0 || n > 3999) throw E.VALUE
  let s = ''
  for (const [v, r] of [[1000, 'M'], [900, 'CM'], [500, 'D'], [400, 'CD'], [100, 'C'], [90, 'XC'], [50, 'L'], [40, 'XL'], [10, 'X'], [9, 'IX'], [5, 'V'], [4, 'IV'], [1, 'I']])
    while (n >= v) {
      s += r
      n -= v
    }
  return s
})

// ── logic ────────────────────────────────────────────────────────────────────────────────────────────────────────────────
const sc = (n, x) => {
  const v = ev(n, x)
  return v instanceof Rng ? one(v, x) : v
}
lazy('IF', (n, x) => {
  const t = ev(n[0], x)
  const cond = t instanceof Rng || isArr(t)
  if (cond) {
    const m = M(t, x)
    const a = n[1] ? M(ev(n[1], x), x) : [[true]]
    const b = n[2] ? M(ev(n[2], x), x) : [[false]]
    return broadcast([m, a, b], ([c, p, q]) => (toBool(c) ? p : q))
  }
  if (toBool(t)) return n[1] ? (n[1].t === 'empty' ? 0 : ev(n[1], x)) : true
  return n[2] ? (n[2].t === 'empty' ? 0 : ev(n[2], x)) : false
})
lazy('IFS', (n, x) => {
  for (let i = 0; i + 1 < n.length; i += 2) if (toBool(sc(n[i], x))) return ev(n[i + 1], x)
  throw E.NA
})
lazy('IFERROR', (n, x) => {
  try {
    const v = ev(n[0], x)
    const s = v instanceof Rng ? one(v, x) : v
    if (s instanceof XlErr) throw s
    return v
  } catch (e) {
    if (e instanceof XlErr) return ev(n[1], x)
    throw e
  }
})
lazy('IFNA', (n, x) => {
  try {
    const v = ev(n[0], x)
    const s = v instanceof Rng ? one(v, x) : v
    if (s instanceof XlErr) throw s
    return v
  } catch (e) {
    if (e instanceof XlErr && e.code === '#N/A') return ev(n[1], x)
    throw e
  }
})
lazy('SWITCH', (n, x) => {
  const v = sc(n[0], x)
  let i = 1
  for (; i + 1 < n.length; i += 2) if (cmp(v, sc(n[i], x)) === 0) return ev(n[i + 1], x)
  if (i < n.length) return ev(n[i], x)
  throw E.NA
})
lazy('CHOOSE', (n, x) => {
  const i = int(sc(n[0], x))
  if (i < 1 || i >= n.length) throw E.VALUE
  return ev(n[i], x)
})
lazy('LET', (n, x) => {
  const vars = { ...(x.vars || {}) }
  for (let i = 0; i + 1 < n.length; i += 2) {
    if (n[i].t !== 'name') throw E.VALUE
    vars[n[i].n.toLowerCase()] = ev(n[i + 1], { ...x, vars })
  }
  return ev(n[n.length - 1], { ...x, vars })
})
const logical = (a, x) => a.flatMap((v) => (v instanceof Rng || isArr(v) ? flat(v, x).filter((q) => q != null && q !== '' && typeof q !== 'string').map(toBool) : [toBool(v)]))
def('AND', (a, x) => {
  const l = logical(a, x)
  if (!l.length) throw E.VALUE
  return l.every(Boolean)
})
def('OR', (a, x) => {
  const l = logical(a, x)
  if (!l.length) throw E.VALUE
  return l.some(Boolean)
})
def('XOR', (a, x) => logical(a, x).filter(Boolean).length % 2 === 1)
S('NOT', (a) => !toBool(a))
def('TRUE', () => true)
def('FALSE', () => false)
def('NA', () => {
  throw E.NA
})
raw('ISBLANK', ([v], x) => (v instanceof Rng ? x.book.get(v.sh, v.r1, v.c1) == null : v == null))
const isFn = (name, f) =>
  raw(name, ([a], x) => {
    const g = (v) => f(v instanceof Rng ? x.book.get(v.sh, v.r1, v.c1) : v)
    return isArr(a) || (a instanceof Rng && (a.r1 !== a.r2 || a.c1 !== a.c2)) ? M(a, x).map((r) => r.map(g)) : g(a)
  })
isFn('ISNUMBER', (v) => typeof v === 'number')
isFn('ISTEXT', (v) => typeof v === 'string')
isFn('ISNONTEXT', (v) => typeof v !== 'string')
isFn('ISLOGICAL', (v) => typeof v === 'boolean')
isFn('ISERROR', (v) => v instanceof XlErr)
isFn('ISERR', (v) => v instanceof XlErr && v.code !== '#N/A')
isFn('ISNA', (v) => v instanceof XlErr && v.code === '#N/A')
isFn('ISEVEN', (v) => Math.trunc(toNum(v)) % 2 === 0)
isFn('ISODD', (v) => Math.abs(Math.trunc(toNum(v))) % 2 === 1)
def('ISFORMULA', ([r], x) => (r instanceof Rng ? x.book.sheets[r.sh].cells.get(r.r1 * MAXC + r.c1)?.f != null : false), { raw: true })
def('ISREF', ([r]) => r instanceof Rng, { raw: true })
S('N', (v) => (typeof v === 'number' ? v : typeof v === 'boolean' ? +v : 0))
S('T', (v) => (typeof v === 'string' ? v : ''))
raw('TYPE', ([v], x) => {
  const s = v instanceof Rng ? one(v, x) : v
  return isArr(v) ? 64 : s instanceof XlErr ? 16 : typeof s === 'number' || s == null ? 1 : typeof s === 'string' ? 2 : 4
})
raw('ERROR.TYPE', ([v], x) => {
  const s = v instanceof Rng ? one(v, x) : v
  const i = ['#NULL!', '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A'].indexOf(s?.code)
  if (i < 0) throw E.NA
  return i + 1
})

// ── text ─────────────────────────────────────────────────────────────────────────────────────────────────────────────────
def('CONCAT', (a, x) => a.map((v) => flat(v, x).map(toStr).join('')).join(''))
def('CONCATENATE', (a, x) => a.map((v) => toStr(one(v, x))).join(''))
def('TEXTJOIN', ([d, ig, ...rest], x) => {
  const sep = toStr(one(d, x))
  const skip = toBool(one(ig, x))
  const items = rest.flatMap((v) => flat(v, x)).map((v) => (v instanceof XlErr ? err(v) : toStr(v)))
  return (skip ? items.filter((v) => v !== '') : items).join(sep)
})
S('LEFT', (t, n) => toStr(t).slice(0, n === undefined ? 1 : Math.max(0, int(n))))
S('RIGHT', (t, n) => {
  const s = toStr(t)
  const k = n === undefined ? 1 : Math.max(0, int(n))
  return k === 0 ? '' : s.slice(-k)
})
S('MID', (t, st, n) => {
  const a = int(st)
  const k = int(n)
  if (a < 1 || k < 0) throw E.VALUE
  return toStr(t).slice(a - 1, a - 1 + k)
})
S('LEN', (t) => toStr(t).length)
S('UPPER', (t) => toStr(t).toUpperCase())
S('LOWER', (t) => toStr(t).toLowerCase())
S('PROPER', (t) => toStr(t).toLowerCase().replace(/(^|[^a-z])([a-z])/g, (m, p, c) => p + c.toUpperCase()))
S('TRIM', (t) => toStr(t).replace(/ +/g, ' ').trim())
S('CLEAN', (t) => toStr(t).replace(/[\x00-\x1f]/g, ''))
S('EXACT', (a, b) => toStr(a) === toStr(b))
S('REPT', (t, n) => {
  const k = int(n)
  if (k < 0) throw E.VALUE
  if (toStr(t).length * k > 32767) throw E.VALUE
  return toStr(t).repeat(k)
})
S(['CHAR', 'UNICHAR'], (n) => {
  const k = int(n)
  if (k < 1 || k > 1114111) throw E.VALUE
  return String.fromCodePoint(k)
})
S(['CODE', 'UNICODE'], (t) => {
  const s = toStr(t)
  if (!s) throw E.VALUE
  return s.codePointAt(0)
})
S('FIND', (f, t, st) => {
  const i = toStr(t).indexOf(toStr(f), st === undefined ? 0 : int(st) - 1)
  if (i < 0) throw E.VALUE
  return i + 1
})
S('SEARCH', (f, t, st) => {
  const hay = toStr(t)
  const start = st === undefined ? 0 : int(st) - 1
  const m = wild(`*${toStr(f)}*`.replace(/^\*|\*$/g, '')).source
  const re = new RegExp(m.replace(/^\^|\$$/g, ''), 'i')
  const r = hay.slice(start).search(re)
  if (r < 0) throw E.VALUE
  return r + start + 1
})
S('SUBSTITUTE', (t, o, n, inst) => {
  const s = toStr(t)
  const old = toStr(o)
  if (!old) return s
  if (inst === undefined) return s.split(old).join(toStr(n))
  let k = int(inst)
  let out = ''
  let i = 0
  for (;;) {
    const j = s.indexOf(old, i)
    if (j < 0) return out + s.slice(i)
    if (--k === 0) return out + s.slice(i, j) + toStr(n) + s.slice(j + old.length)
    out += s.slice(i, j + old.length)
    i = j + old.length
  }
})
S('REPLACE', (t, st, n, r) => {
  const s = toStr(t)
  const a = int(st) - 1
  if (a < 0) throw E.VALUE
  return s.slice(0, a) + toStr(r) + s.slice(a + int(n))
})
S('TEXT', (v, f) => formatValue(typeof v === 'string' && parseTyped(v) ? parseTyped(v).v : v, toStr(f)))
S('VALUE', (t) => {
  if (typeof t === 'number') return t
  const p = parseTyped(toStr(t))
  if (!p || typeof p.v !== 'number') throw E.VALUE
  return p.v
})
S('NUMBERVALUE', (t, dec, grp) => {
  let s = toStr(t).replace(/\s/g, '')
  const d = dec === undefined ? '.' : toStr(dec)
  const g = grp === undefined ? ',' : toStr(grp)
  s = s.split(g).join('').replace(d, '.')
  const pct = /%+$/.exec(s)
  if (pct) s = s.slice(0, -pct[0].length)
  const n = Number(s)
  if (s === '' ) return 0
  if (Number.isNaN(n)) throw E.VALUE
  return pct ? n / 100 ** pct[0].length : n
})
S('FIXED', (v, d, noCommas) => {
  const k = d === undefined ? 2 : int(d)
  const s = toNum(v).toFixed(Math.max(0, k))
  return noCommas !== undefined && toBool(noCommas) ? s : s.replace(/\B(?=(\d{3})+(?!\d))/g, ',')
})
S('DOLLAR', (v, d) => formatValue(toNum(v), `$#,##0${(d === undefined ? 2 : int(d)) > 0 ? `.${'0'.repeat(d === undefined ? 2 : int(d))}` : ''}`))
S('TEXTBEFORE', (t, d, n) => {
  const s = toStr(t)
  const k = n === undefined ? 1 : int(n)
  let i = -1
  for (let j = 0; j < k; j++) {
    i = s.indexOf(toStr(d), i + 1)
    if (i < 0) throw E.NA
  }
  return s.slice(0, i)
})
S('TEXTAFTER', (t, d, n) => {
  const s = toStr(t)
  const k = n === undefined ? 1 : int(n)
  let i = -1
  for (let j = 0; j < k; j++) {
    i = s.indexOf(toStr(d), i + 1)
    if (i < 0) throw E.NA
  }
  return s.slice(i + toStr(d).length)
})
def('TEXTSPLIT', ([t, cd, rd], x) => {
  const s = toStr(one(t, x))
  const cs = cd === undefined || cd === null ? null : toStr(one(cd, x))
  const rs = rd === undefined || rd === null ? null : toStr(one(rd, x))
  const rows = rs ? s.split(rs) : [s]
  const out = rows.map((r) => (cs ? r.split(cs) : [r]))
  const w = Math.max(...out.map((r) => r.length))
  return out.map((r) => [...r, ...Array(w - r.length).fill(E.NA)])
})
S('HYPERLINK', (u, f) => (f === undefined ? toStr(u) : toStr(f)))
S('LENB', (t) => new TextEncoder().encode(toStr(t)).length)

// ── dates ────────────────────────────────────────────────────────────────────────────────────────────────────────────────
const dateOf = (v) => {
  if (typeof v === 'string') {
    const p = parseTyped(v.trim())
    if (p && typeof p.v === 'number') return p.v
    throw E.VALUE
  }
  const n = toNum(v)
  if (n < 0) throw E.NUM
  return n
}
const ymd = (y, m, d) => toSerial(y, m, d)
def('TODAY', () => Math.floor(nowSerial()))
def('NOW', () => nowSerial())
S('DATE', (y, m, d) => {
  let yy = int(y)
  if (yy >= 0 && yy < 1900) yy += 1900
  const dt = new Date(Date.UTC(yy, int(m) - 1, int(d)))
  if (yy < 1900 && yy > 9999) throw E.NUM
  return ymd(dt.getUTCFullYear(), dt.getUTCMonth() + 1, dt.getUTCDate())
})
S('YEAR', (d) => fromSerial(dateOf(d)).y)
S('MONTH', (d) => fromSerial(dateOf(d)).m)
S('DAY', (d) => fromSerial(dateOf(d)).d)
S('HOUR', (d) => fromSerial(dateOf(d)).h)
S('MINUTE', (d) => fromSerial(dateOf(d)).mi)
S('SECOND', (d) => fromSerial(dateOf(d)).s)
S('TIME', (h, m, s) => (int(h) * 3600 + int(m) * 60 + int(s)) / 86400 - Math.floor((int(h) * 3600 + int(m) * 60 + int(s)) / 86400))
S('DATEVALUE', (t) => Math.floor(dateOf(toStr(t))))
S('TIMEVALUE', (t) => {
  const v = dateOf(toStr(t))
  return v - Math.floor(v)
})
S('WEEKDAY', (d, t) => {
  const wd = fromSerial(dateOf(d)).wd
  const type = t === undefined ? 1 : int(t)
  return type === 2 ? ((wd + 6) % 7) + 1 : type === 3 ? (wd + 6) % 7 : wd + 1
})
S('WEEKNUM', (d, t) => {
  const s = Math.floor(dateOf(d))
  const { y } = fromSerial(s)
  const jan1 = ymd(y, 1, 1)
  const startDay = t !== undefined && int(t) === 2 ? 1 : 0
  const wd1 = fromSerial(jan1).wd
  return Math.floor((s - jan1 + ((wd1 - startDay + 7) % 7)) / 7) + 1
})
S('ISOWEEKNUM', (d) => {
  const s = Math.floor(dateOf(d))
  const wd = (fromSerial(s).wd + 6) % 7
  const th = s - wd + 3
  const y = fromSerial(th).y
  return Math.floor((th - ymd(y, 1, 1)) / 7) + 1
})
const addMonths = (s, n) => {
  const { y, m, d } = fromSerial(Math.floor(s))
  const total = y * 12 + (m - 1) + n
  const ny = Math.floor(total / 12)
  const nm = (total % 12) + 1
  const last = new Date(Date.UTC(ny, nm, 0)).getUTCDate()
  return ymd(ny, nm, Math.min(d, last))
}
S('EDATE', (d, n) => addMonths(dateOf(d), int(n)))
S('EOMONTH', (d, n) => {
  const s = addMonths(dateOf(d), int(n))
  const { y, m } = fromSerial(s)
  return ymd(y, m, new Date(Date.UTC(y, m, 0)).getUTCDate())
})
S('DAYS', (e, s) => Math.floor(dateOf(e)) - Math.floor(dateOf(s)))
S('DATEDIF', (a, b, u) => {
  const s = Math.floor(dateOf(a))
  const e = Math.floor(dateOf(b))
  if (e < s) throw E.NUM
  const A = fromSerial(s)
  const B = fromSerial(e)
  let months = (B.y - A.y) * 12 + (B.m - A.m) - (B.d < A.d ? 1 : 0)
  switch (toStr(u).toUpperCase()) {
    case 'D':
      return e - s
    case 'M':
      return months
    case 'Y':
      return Math.floor(months / 12)
    case 'MD':
      return B.d >= A.d ? B.d - A.d : e - addMonths(s, months)
    case 'YM':
      months = ((months % 12) + 12) % 12
      return months
    case 'YD': {
      const yrs = Math.floor(months / 12)
      return e - addMonths(s, yrs * 12)
    }
    default:
      throw E.NUM
  }
})
S('YEARFRAC', (a, b, basis) => {
  const s = dateOf(a)
  const e = dateOf(b)
  const k = basis === undefined ? 0 : int(basis)
  const [lo, hi] = s <= e ? [s, e] : [e, s]
  if (k === 0 || k === 4) {
    const A = fromSerial(lo)
    const B = fromSerial(hi)
    const d1 = Math.min(A.d, 30)
    const d2 = B.d === 31 && d1 >= 30 ? 30 : B.d
    return ((B.y - A.y) * 360 + (B.m - A.m) * 30 + (d2 - d1)) / 360
  }
  return (hi - lo) / (k === 2 ? 360 : 365)
})
const holidays = (h, x) => new Set(h === undefined ? [] : flat(h, x).filter((v) => v != null).map((v) => Math.floor(dateOf(v))))
const isWork = (s, hol) => {
  const wd = fromSerial(s).wd
  return wd !== 0 && wd !== 6 && !hol.has(s)
}
def('NETWORKDAYS', ([a, b, h], x) => {
  const s = Math.floor(dateOf(one(a, x)))
  const e = Math.floor(dateOf(one(b, x)))
  const hol = holidays(h, x)
  let n = 0
  for (let d = Math.min(s, e); d <= Math.max(s, e); d++) if (isWork(d, hol)) n++
  return s <= e ? n : -n
})
def('WORKDAY', ([a, n, h], x) => {
  let s = Math.floor(dateOf(one(a, x)))
  let k = int(one(n, x))
  const hol = holidays(h, x)
  const step = k < 0 ? -1 : 1
  while (k !== 0) {
    s += step
    if (isWork(s, hol)) k -= step
  }
  return s
})

// ── lookup and reference ─────────────────────────────────────────────────────────────────────────────────────────────────
const refOf = (v) => {
  if (!(v instanceof Rng)) throw E.VALUE
  return v
}
const eq = (a, b) => {
  if (a == null || b == null) return a == null && b == null ? true : false
  return cmp(a, b) === 0 && typeof a === typeof b
}
const findExact = (cells, key) => {
  if (typeof key === 'string' && /[*?~]/.test(key)) {
    const re = wild(key)
    return cells.findIndex((v) => typeof v === 'string' && re.test(v))
  }
  return cells.findIndex((v) => eq(v, key))
}
const findApprox = (cells, key, desc) => {
  let best = -1
  for (let i = 0; i < cells.length; i++) {
    const v = cells[i]
    if (v == null || typeof v !== typeof key) continue
    const d = cmp(v, key)
    if (desc ? d >= 0 : d <= 0) best = i
    else if (!desc) break
  }
  return best
}
def('VLOOKUP', ([k, t, c, a], x) => {
  const key = err(one(k, x))
  const m = M(t, x)
  const col = int(one(c, x))
  if (col < 1 || col > m[0].length) throw E.REF
  const approx = a === undefined ? true : toBool(one(a, x))
  const first = m.map((r) => r[0])
  const i = approx ? findApprox(first, key) : findExact(first, key)
  if (i < 0) throw E.NA
  return err(m[i][col - 1])
})
def('HLOOKUP', ([k, t, rr, a], x) => {
  const key = err(one(k, x))
  const m = M(t, x)
  const row = int(one(rr, x))
  if (row < 1 || row > m.length) throw E.REF
  const approx = a === undefined ? true : toBool(one(a, x))
  const i = approx ? findApprox(m[0], key) : findExact(m[0], key)
  if (i < 0) throw E.NA
  return err(m[row - 1][i])
})
def('LOOKUP', ([k, l, r], x) => {
  const key = err(one(k, x))
  const lm = M(l, x)
  const cells = lm.length === 1 ? lm[0] : lm.map((q) => q[0])
  const i = findApprox(cells, key)
  if (i < 0) throw E.NA
  if (r === undefined) return lm.length === 1 ? lm[0][i] : lm[lm.length - 1][lm[0].length - 1 === 0 ? 0 : lm[0].length - 1]
  const rm = M(r, x)
  return rm.length === 1 ? rm[0][i] : rm[i][0]
})
def('MATCH', ([k, a, t], x) => {
  const key = err(one(k, x))
  const m = M(a, x)
  const cells = m.length === 1 ? m[0] : m.map((q) => q[0])
  const type = t === undefined ? 1 : toNum(one(t, x))
  const i = type === 0 ? findExact(cells, key) : findApprox(cells, key, type < 0)
  if (i < 0) throw E.NA
  return i + 1
})
def('XMATCH', ([k, a, mm, sm], x) => xfind(one(k, x), flat(a, x), mm === undefined ? 0 : toNum(one(mm, x)), sm === undefined ? 1 : toNum(one(sm, x))) + 1)
function xfind(key, cells, mode, search) {
  err(key)
  const idx = cells.map((_, i) => i)
  if (search < 0) idx.reverse()
  if (mode === 0) {
    for (const i of idx) if (eq(cells[i], key)) return i
  } else if (mode === 2) {
    const re = wild(String(key))
    for (const i of idx) if (typeof cells[i] === 'string' && re.test(cells[i])) return i
  } else {
    let best = -1
    for (const i of idx) {
      const v = cells[i]
      if (v == null || typeof v !== typeof key) continue
      const d = cmp(v, key)
      if (d === 0) return i
      if (mode < 0 ? d < 0 : d > 0) if (best < 0 || (mode < 0 ? cmp(v, cells[best]) > 0 : cmp(v, cells[best]) < 0)) best = i
    }
    if (best >= 0) return best
  }
  throw E.NA
}
def('XLOOKUP', (args, x) => {
  const [k, la, ra, nf, mm, sm] = args
  const key = one(k, x)
  const lm = M(la, x)
  const rm = M(ra, x)
  const vertical = lm[0].length === 1
  const cells = vertical ? lm.map((r) => r[0]) : lm[0]
  let i
  try {
    i = xfind(key, cells, mm === undefined ? 0 : toNum(one(mm, x)), sm === undefined ? 1 : toNum(one(sm, x)))
  } catch (e) {
    if (e instanceof XlErr && e.code === '#N/A' && nf !== undefined && nf !== null) return nf
    throw e
  }
  if (vertical) return rm[i].length === 1 ? rm[i][0] : [rm[i]]
  const col = rm.map((r) => r[i])
  return col.length === 1 ? col[0] : col.map((v) => [v])
})
def('INDEX', ([a, r, c], x) => {
  const m = M(a, x)
  let ri = r === undefined || r === null ? 0 : int(one(r, x))
  let ci = c === undefined || c === null ? 0 : int(one(c, x))
  if (m.length === 1 && c === undefined && r !== undefined) {
    ci = ri
    ri = 1
  }
  if (ri < 0 || ci < 0 || ri > m.length || ci > m[0].length) throw E.REF
  if (ri === 0 && ci === 0) return m
  if (ri === 0) return m.map((row) => [row[ci - 1]])
  if (ci === 0) return [m[ri - 1]]
  return err(m[ri - 1][ci - 1])
})
def('OFFSET', ([r, dr, dc, h, w], x) => {
  const g = refOf(r)
  const hh = h === undefined ? g.r2 - g.r1 + 1 : int(one(h, x))
  const ww = w === undefined ? g.c2 - g.c1 + 1 : int(one(w, x))
  const r1 = g.r1 + int(one(dr, x))
  const c1 = g.c1 + int(one(dc, x))
  if (r1 < 0 || c1 < 0 || hh < 1 || ww < 1 || r1 + hh > MAXR || c1 + ww > MAXC) throw E.REF
  return new Rng(g.sh, r1, c1, r1 + hh - 1, c1 + ww - 1)
})
def('INDIRECT', ([t], x) => {
  const text = toStr(one(t, x))
  const m = text.match(/^(?:'((?:[^']|'')+)'|([^'!]+))!(.+)$/)
  const sh = m ? x.book.sheetIndex((m[1] ?? m[2]).replace(/''/g, "'")) : x.sh
  if (sh < 0) throw E.REF
  const g = parseRange(m ? m[3] : text)
  if (!g) {
    const def2 = x.book.nameDef(text)
    if (def2) return ev(def2, x)
    throw E.REF
  }
  return new Rng(sh, g.r1, g.c1, g.r2, g.c2)
})
def('ROW', ([r], x) => (r === undefined ? x.r + 1 : Array.from({ length: refOf(r).r2 - r.r1 + 1 }, (_, i) => [r.r1 + i + 1]).slice(0, 1e5).reduce((a, v, i) => (i === 0 && r.r1 === r.r2 ? v[0] : a.concat([v])), [])))
def('COLUMN', ([r], x) => (r === undefined ? x.c + 1 : r.c1 === r.c2 ? r.c1 + 1 : [Array.from({ length: Math.min(r.c2, 16383) - r.c1 + 1 }, (_, i) => r.c1 + i + 1)]))
def('ROWS', ([r], x) => (r instanceof Rng ? Math.min(r.r2, MAXR - 1) - r.r1 + 1 : M(r, x).length))
def('COLUMNS', ([r], x) => (r instanceof Rng ? r.c2 - r.c1 + 1 : M(r, x)[0].length))
def('ADDRESS', ([r, c, a, a1, s], x) => {
  const ri = int(one(r, x))
  const ci = int(one(c, x))
  const abs = a === undefined ? 1 : int(one(a, x))
  const t = `${abs === 1 || abs === 3 ? '$' : ''}${colName(ci - 1)}${abs === 1 || abs === 2 ? '$' : ''}${ri}`
  return s === undefined ? t : `${toStr(one(s, x))}!${t}`
})
def('TRANSPOSE', ([a], x) => {
  const m = M(a, x)
  return m[0].map((_, c) => m.map((r) => r[c]))
})
def('HSTACK', (a, x) => {
  const ms = a.map((v) => M(v, x))
  const h = Math.max(...ms.map((m) => m.length))
  return Array.from({ length: h }, (_, r) => ms.flatMap((m) => (r < m.length ? m[r] : Array(m[0].length).fill(E.NA))))
})
def('VSTACK', (a, x) => {
  const ms = a.map((v) => M(v, x))
  const w = Math.max(...ms.map((m) => m[0].length))
  return ms.flatMap((m) => m.map((r) => [...r, ...Array(w - r.length).fill(E.NA)]))
})
def('TAKE', ([a, r, c], x) => {
  let m = M(a, x)
  if (r !== undefined && r !== null) {
    const n = int(one(r, x))
    m = n >= 0 ? m.slice(0, n) : m.slice(n)
  }
  if (c !== undefined && c !== null) {
    const n = int(one(c, x))
    m = m.map((row) => (n >= 0 ? row.slice(0, n) : row.slice(n)))
  }
  if (!m.length || !m[0].length) throw E.CALC
  return m
})
def('DROP', ([a, r, c], x) => {
  let m = M(a, x)
  if (r !== undefined && r !== null) {
    const n = int(one(r, x))
    m = n >= 0 ? m.slice(n) : m.slice(0, n)
  }
  if (c !== undefined && c !== null) {
    const n = int(one(c, x))
    m = m.map((row) => (n >= 0 ? row.slice(n) : row.slice(0, n)))
  }
  if (!m.length || !m[0].length) throw E.CALC
  return m
})
def('TOCOL', ([a], x) => M(a, x).flat().map((v) => [v]))
def('TOROW', ([a], x) => [M(a, x).flat()])
def('CHOOSECOLS', ([a, ...cs], x) => {
  const m = M(a, x)
  const idx = cs.flatMap((c) => flat(c, x)).map((v) => int(v) - 1)
  if (idx.some((i) => i < 0 || i >= m[0].length)) throw E.VALUE
  return m.map((r) => idx.map((i) => r[i]))
})
def('FILTER', ([a, inc, ifEmpty], x) => {
  const m = M(a, x)
  const im = M(inc, x)
  const byRow = im[0].length === 1 && im.length === m.length
  const byCol = im.length === 1 && im[0].length === m[0].length
  if (!byRow && !byCol) throw E.VALUE
  const out = byRow ? m.filter((_, i) => toBool(im[i][0])) : m.map((r) => r.filter((_, j) => toBool(im[0][j])))
  if (!out.length || !out[0].length) {
    if (ifEmpty !== undefined) return ifEmpty
    throw E.CALC
  }
  return out
})
const sortRows = (m, keys, orders) =>
  m
    .map((r, i) => ({ r, i }))
    .sort((a, b) => {
      for (let k = 0; k < keys.length; k++) {
        const d = cmp(keys[k][a.i], keys[k][b.i])
        if (d) return orders[k] < 0 ? -d : d
      }
      return a.i - b.i
    })
    .map((o) => o.r)
def('SORT', ([a, si, so, bc], x) => {
  let m = M(a, x)
  const byCol = bc !== undefined && toBool(one(bc, x))
  if (byCol) m = m[0].map((_, c) => m.map((r) => r[c]))
  const idx = si === undefined ? 1 : int(one(si, x))
  if (idx < 1 || idx > m[0].length) throw E.VALUE
  const order = so === undefined ? 1 : toNum(one(so, x))
  let out = sortRows(m, [m.map((r) => r[idx - 1])], [order])
  if (byCol) out = out[0].map((_, c) => out.map((r) => r[c]))
  return out
})
def('SORTBY', ([a, ...rest], x) => {
  const m = M(a, x)
  const keys = []
  const orders = []
  for (let i = 0; i < rest.length; i += 2) {
    const k = flat(rest[i], x)
    if (k.length !== m.length) throw E.VALUE
    keys.push(k)
    orders.push(rest[i + 1] === undefined ? 1 : toNum(one(rest[i + 1], x)))
  }
  return sortRows(m, keys, orders)
})
def('UNIQUE', ([a, bc, once], x) => {
  let m = M(a, x)
  const byCol = bc !== undefined && toBool(one(bc, x))
  if (byCol) m = m[0].map((_, c) => m.map((r) => r[c]))
  const seen = new Map()
  for (const r of m) {
    const k = JSON.stringify(r.map((v) => (typeof v === 'string' ? v.toLowerCase() : v)))
    seen.set(k, { r, n: (seen.get(k)?.n || 0) + 1 })
  }
  let out = [...seen.values()].filter((e) => once === undefined || !toBool(one(once, x)) || e.n === 1).map((e) => e.r)
  if (!out.length) throw E.CALC
  if (byCol) out = out[0].map((_, c) => out.map((r) => r[c]))
  return out
})
def('SEQUENCE', ([r, c, s, st], x) => {
  const h = int(one(r, x))
  const w = c === undefined ? 1 : int(one(c, x))
  const start = s === undefined ? 1 : toNum(one(s, x))
  const step = st === undefined ? 1 : toNum(one(st, x))
  if (h < 1 || w < 1 || h * w > 1e6) throw E.CALC
  return Array.from({ length: h }, (_, i) => Array.from({ length: w }, (_, j) => start + (i * w + j) * step))
})
def('RANDARRAY', ([r, c, lo, hi, wh], x) => {
  const h = r === undefined ? 1 : int(one(r, x))
  const w = c === undefined ? 1 : int(one(c, x))
  const a = lo === undefined ? 0 : toNum(one(lo, x))
  const b = hi === undefined ? 1 : toNum(one(hi, x))
  const whole = wh !== undefined && toBool(one(wh, x))
  return Array.from({ length: h }, () => Array.from({ length: w }, () => (whole ? a + Math.floor(Math.random() * (b - a + 1)) : a + Math.random() * (b - a))))
})
def('MMULT', ([a, b], x) => {
  const p = M(a, x)
  const q = M(b, x)
  if (p[0].length !== q.length) throw E.VALUE
  return p.map((r) => q[0].map((_, j) => sum(r.map((v, k) => toNum(v) * toNum(q[k][j])))))
})
def('MDETERM', ([a], x) => {
  const m = M(a, x).map((r) => r.map(toNum))
  const n = m.length
  if (m.some((r) => r.length !== n)) throw E.VALUE
  let det = 1
  for (let i = 0; i < n; i++) {
    let p = i
    for (let r = i + 1; r < n; r++) if (Math.abs(m[r][i]) > Math.abs(m[p][i])) p = r
    if (m[p][i] === 0) return 0
    if (p !== i) {
      ;[m[p], m[i]] = [m[i], m[p]]
      det = -det
    }
    det *= m[i][i]
    for (let r = i + 1; r < n; r++) {
      const f = m[r][i] / m[i][i]
      for (let c = i; c < n; c++) m[r][c] -= f * m[i][c]
    }
  }
  return det
})

// ── finance ──────────────────────────────────────────────────────────────────────────────────────────────────────────────
const pmt = (rate, n, pv, fv, type) => (rate === 0 ? -(pv + fv) / n : (-(pv * (1 + rate) ** n + fv) * rate) / ((1 + rate * type) * ((1 + rate) ** n - 1)))
const fvf = (rate, n, p, pv, type) => (rate === 0 ? -(pv + p * n) : -(pv * (1 + rate) ** n + (p * (1 + rate * type) * ((1 + rate) ** n - 1)) / rate))
S('PMT', (r, n, pv, fv, t) => pmt(toNum(r), toNum(n), toNum(pv), fv === undefined ? 0 : toNum(fv), t === undefined ? 0 : toNum(t)))
S('FV', (r, n, p, pv, t) => fvf(toNum(r), toNum(n), toNum(p), pv === undefined ? 0 : toNum(pv), t === undefined ? 0 : toNum(t)))
S('PV', (r, n, p, fv, t) => {
  const rate = toNum(r)
  const nn = toNum(n)
  const pm = toNum(p)
  const f = fv === undefined ? 0 : toNum(fv)
  const ty = t === undefined ? 0 : toNum(t)
  return rate === 0 ? -(f + pm * nn) : -(f + (pm * (1 + rate * ty) * ((1 + rate) ** nn - 1)) / rate) / (1 + rate) ** nn
})
S('NPER', (r, p, pv, fv, t) => {
  const rate = toNum(r)
  const pm = toNum(p)
  const v = toNum(pv)
  const f = fv === undefined ? 0 : toNum(fv)
  const ty = t === undefined ? 0 : toNum(t)
  if (rate === 0) return -(v + f) / pm
  const a = pm * (1 + rate * ty) - f * rate
  const b = pm * (1 + rate * ty) + v * rate
  if (a / b <= 0) throw E.NUM
  return Math.log(a / b) / Math.log(1 + rate)
})
// interest part of payment number k, by walking the balance (signs as the spreadsheet: a loan is positive, payments negative)
const interestOf = (rate, k, n, pv, fv, ty) => {
  if (k < 1 || k > n) throw E.NUM
  const walk = (upTo) => {
    const p = pmt(rate, n, pv, fv, 0)
    let bal = pv
    let interest = 0
    for (let i = 1; i <= upTo; i++) {
      interest = -bal * rate
      bal = bal * (1 + rate) + p
    }
    return interest
  }
  if (ty !== 1) return walk(k)
  return k === 1 ? 0 : walk(k - 1) / (1 + rate)
}
const loanArgs = (r, n, pv, fv, t) => [toNum(r), toNum(n), toNum(pv), fv === undefined ? 0 : toNum(fv), t === undefined ? 0 : toNum(t)]
S('IPMT', (r, per, n, pv, fv, t) => {
  const [rate, nn, v, f, ty] = loanArgs(r, n, pv, fv, t)
  return interestOf(rate, toNum(per), nn, v, f, ty)
})
S('PPMT', (r, per, n, pv, fv, t) => {
  const [rate, nn, v, f, ty] = loanArgs(r, n, pv, fv, t)
  return pmt(rate, nn, v, f, ty) - interestOf(rate, toNum(per), nn, v, f, ty)
})
S('SLN', (c, s, l) => (toNum(c) - toNum(s)) / toNum(l))
S('EFFECT', (r, n) => (1 + toNum(r) / toNum(n)) ** toNum(n) - 1)
S('NOMINAL', (r, n) => toNum(n) * ((1 + toNum(r)) ** (1 / toNum(n)) - 1))
def('NPV', ([r, ...v], x) => {
  const rate = toNum(one(r, x))
  return list(v, x).reduce((s, c, i) => s + c / (1 + rate) ** (i + 1), 0)
})
const irr = (cfs, guess = 0.1) => {
  let r = guess
  for (let i = 0; i < 100; i++) {
    let f = 0
    let d = 0
    cfs.forEach((c, k) => {
      f += c / (1 + r) ** k
      d -= (k * c) / (1 + r) ** (k + 1)
    })
    if (d === 0) break
    const n = r - f / d
    if (Math.abs(n - r) < 1e-10) return n
    r = n
  }
  throw E.NUM
}
def('IRR', ([v, g], x) => {
  const l = list([v], x)
  if (!l.some((c) => c > 0) || !l.some((c) => c < 0)) throw E.NUM
  return irr(l, g === undefined ? 0.1 : toNum(one(g, x)))
})
def('XNPV', ([r, v, d], x) => {
  const rate = toNum(one(r, x))
  const vals = list([v], x)
  const ds = flat(d, x).map(dateOf)
  return vals.reduce((s, c, i) => s + c / (1 + rate) ** ((ds[i] - ds[0]) / 365), 0)
})
def('XIRR', ([v, d, g], x) => {
  const vals = list([v], x)
  const ds = flat(d, x).map(dateOf)
  let r = g === undefined ? 0.1 : toNum(one(g, x))
  for (let i = 0; i < 100; i++) {
    let f = 0
    let df = 0
    vals.forEach((c, k) => {
      const t = (ds[k] - ds[0]) / 365
      f += c / (1 + r) ** t
      df -= (t * c) / (1 + r) ** (t + 1)
    })
    const n = r - f / df
    if (Math.abs(n - r) < 1e-10) return n
    r = n
  }
  throw E.NUM
})
S('RATE', (n, p, pv, fv, t, g) => {
  const nn = toNum(n)
  const pm = toNum(p)
  const v = toNum(pv)
  const f = fv === undefined ? 0 : toNum(fv)
  const ty = t === undefined ? 0 : toNum(t)
  let r = g === undefined ? 0.1 : toNum(g)
  for (let i = 0; i < 100; i++) {
    const fn = (q) => v * (1 + q) ** nn + pm * (1 + q * ty) * (((1 + q) ** nn - 1) / q) + f
    const h = 1e-7
    const y = fn(r)
    const d = (fn(r + h) - y) / h
    const nr = r - y / d
    if (Math.abs(nr - r) < 1e-10) return nr
    r = nr
  }
  throw E.NUM
})
S('NORM.S.DIST', (z, cum) => {
  const v = toNum(z)
  if (cum !== undefined && !toBool(cum)) return Math.exp(-v * v / 2) / Math.sqrt(2 * Math.PI)
  return normcdf(v)
})
S('NORM.DIST', (v, m, s, cum) => {
  const sd = toNum(s)
  if (sd <= 0) throw E.NUM
  const z = (toNum(v) - toNum(m)) / sd
  return toBool(cum) ? normcdf(z) : Math.exp(-z * z / 2) / (sd * Math.sqrt(2 * Math.PI))
})
function normcdf(z) {
  const t = 1 / (1 + 0.2316419 * Math.abs(z))
  const d = 0.3989422804014327 * Math.exp((-z * z) / 2)
  const p = d * t * (0.31938153 + t * (-0.356563782 + t * (1.781477937 + t * (-1.821255978 + t * 1.330274429))))
  return z > 0 ? 1 - p : p
}

export const FUNCTION_NAMES = () => Object.keys(FN).filter((n) => !n.startsWith('__')).sort()
void general
void monthNames
void addr
void parse
void toMat
