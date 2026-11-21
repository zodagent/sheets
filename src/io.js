// Files in and out: CSV/TSV, and .xlsx (a zip of XML) read and written with no dependencies except node:zlib.
import { deflateRawSync, inflateRawSync } from 'node:zlib'
import { XlErr, errByCode, formatValue } from './format.js'
import { Book, addr, colIndex, colName, parseAddr, parseInput } from './core.js'

// ── CSV ──────────────────────────────────────────────────────────────────────────────────────────────────────────────────
export function parseCSV(text, delimiter) {
  const src = String(text).replace(/^﻿/, '')
  const d = delimiter || guessDelimiter(src)
  const rows = []
  let row = []
  let cell = ''
  let q = false
  for (let i = 0; i < src.length; i++) {
    const ch = src[i]
    if (q) {
      if (ch === '"') {
        if (src[i + 1] === '"') {
          cell += '"'
          i++
        } else q = false
      } else cell += ch
    } else if (ch === '"' && cell === '') q = true
    else if (ch === d) {
      row.push(cell)
      cell = ''
    } else if (ch === '\n' || ch === '\r') {
      if (ch === '\r' && src[i + 1] === '\n') i++
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
  return rows
}
function guessDelimiter(s) {
  const head = s.split(/\r?\n/).slice(0, 5).join('\n')
  const count = (d) => head.split(d).length
  return [',', '\t', ';', '|'].sort((a, b) => count(b) - count(a))[0]
}
export function toCSV(rows, delimiter = ',') {
  const esc = (v) => {
    const s = v == null ? '' : String(v)
    return /["\n\r]/.test(s) || s.includes(delimiter) ? `"${s.replace(/"/g, '""')}"` : s
  }
  return rows.map((r) => r.map(esc).join(delimiter)).join('\n')
}
// rows of text → `set` operation cells, typed the way a person would type them
export function cellsFromRows(rows, top = { r: 0, c: 0 }, { typed = true } = {}) {
  const cells = {}
  rows.forEach((row, i) =>
    row.forEach((v, j) => {
      if (v === '' || v == null) return
      const p = typed ? parseInput(v) : { cell: { v: String(v) } }
      if (!p.cell) return
      cells[addr(top.r + i, top.c + j)] = p.nf ? { ...p.cell, s: { nf: p.nf } } : p.cell
    }),
  )
  return cells
}
// a range of a book as rows of text (display) or values
export function rowsOf(book, sheet, g, { raw = false } = {}) {
  book.recalc()
  const sh = book.sheetIndex(sheet)
  const out = []
  for (let r = g.r1; r <= g.r2; r++) {
    const row = []
    for (let c = g.c1; c <= g.c2; c++) {
      const v = book.get(sh, r, c)
      row.push(raw ? (v instanceof XlErr ? v.code : v) : book.display(sh, r, c))
    }
    out.push(row)
  }
  return out
}

// ── zip ──────────────────────────────────────────────────────────────────────────────────────────────────────────────────
const CRC = (() => {
  const t = new Uint32Array(256)
  for (let n = 0; n < 256; n++) {
    let c = n
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1
    t[n] = c >>> 0
  }
  return t
})()
const crc32 = (buf) => {
  let c = 0xffffffff
  for (let i = 0; i < buf.length; i++) c = CRC[(c ^ buf[i]) & 0xff] ^ (c >>> 8)
  return (c ^ 0xffffffff) >>> 0
}
export function zip(files) {
  const parts = []
  const central = []
  let offset = 0
  for (const [name, text] of Object.entries(files)) {
    const data = Buffer.from(text, 'utf8')
    const comp = deflateRawSync(data)
    const nm = Buffer.from(name)
    const h = Buffer.alloc(30)
    h.writeUInt32LE(0x04034b50, 0)
    h.writeUInt16LE(20, 4)
    h.writeUInt16LE(0x0800, 6)
    h.writeUInt16LE(8, 8)
    h.writeUInt32LE(crc32(data), 14)
    h.writeUInt32LE(comp.length, 18)
    h.writeUInt32LE(data.length, 22)
    h.writeUInt16LE(nm.length, 26)
    parts.push(h, nm, comp)
    const c = Buffer.alloc(46)
    c.writeUInt32LE(0x02014b50, 0)
    c.writeUInt16LE(20, 4)
    c.writeUInt16LE(20, 6)
    c.writeUInt16LE(0x0800, 8)
    c.writeUInt16LE(8, 10)
    c.writeUInt32LE(crc32(data), 16)
    c.writeUInt32LE(comp.length, 20)
    c.writeUInt32LE(data.length, 24)
    c.writeUInt16LE(nm.length, 28)
    c.writeUInt32LE(offset, 42)
    central.push(c, nm)
    offset += h.length + nm.length + comp.length
  }
  const cd = Buffer.concat(central)
  const end = Buffer.alloc(22)
  end.writeUInt32LE(0x06054b50, 0)
  end.writeUInt16LE(Object.keys(files).length, 8)
  end.writeUInt16LE(Object.keys(files).length, 10)
  end.writeUInt32LE(cd.length, 12)
  end.writeUInt32LE(offset, 16)
  return Buffer.concat([...parts, cd, end])
}
export function unzip(buf) {
  let e = buf.length - 22
  while (e >= 0 && buf.readUInt32LE(e) !== 0x06054b50) e--
  if (e < 0) throw new Error('that is not an .xlsx file (no zip index)')
  const n = buf.readUInt16LE(e + 10)
  let p = buf.readUInt32LE(e + 16)
  const out = {}
  for (let i = 0; i < n; i++) {
    if (buf.readUInt32LE(p) !== 0x02014b50) throw new Error('the zip index is damaged')
    const method = buf.readUInt16LE(p + 10)
    const csize = buf.readUInt32LE(p + 20)
    const nl = buf.readUInt16LE(p + 28)
    const xl = buf.readUInt16LE(p + 30)
    const cl = buf.readUInt16LE(p + 32)
    const lo = buf.readUInt32LE(p + 42)
    const name = buf.toString('utf8', p + 46, p + 46 + nl)
    const lnl = buf.readUInt16LE(lo + 26)
    const lxl = buf.readUInt16LE(lo + 28)
    const start = lo + 30 + lnl + lxl
    const raw = buf.subarray(start, start + csize)
    if (/\.(xml|rels)$/.test(name)) out[name] = (method === 8 ? inflateRawSync(raw) : raw).toString('utf8')
    p += 46 + nl + xl + cl
  }
  return out
}

// ── xlsx: write ──────────────────────────────────────────────────────────────────────────────────────────────────────────
const xe = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;')
const argb = (h) => `FF${String(h || '000000').replace('#', '').toUpperCase().padEnd(6, '0')}`

export function writeXlsx(book) {
  book.recalc()
  const fonts = ['<font><sz val="11"/><name val="Calibri"/></font>']
  const fills = ['<fill><patternFill patternType="none"/></fill>', '<fill><patternFill patternType="gray125"/></fill>']
  const borders = ['<border><left/><right/><top/><bottom/><diagonal/></border>']
  const numFmts = []
  const xfs = ['<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>']
  const seen = new Map()
  const idOf = (list, xml) => {
    let i = list.indexOf(xml)
    if (i < 0) i = list.push(xml) - 1
    return i
  }
  const BUILTIN = { '0': 1, '0.00': 2, '#,##0': 3, '#,##0.00': 4, '0%': 9, '0.00%': 10, 'm/d/yyyy': 14 }
  const styleId = (s) => {
    if (!s || !Object.keys(s).length) return 0
    const key = JSON.stringify(s)
    if (seen.has(key)) return seen.get(key)
    const f = `<font>${s.b ? '<b/>' : ''}${s.i ? '<i/>' : ''}${s.u ? '<u/>' : ''}${s.st ? '<strike/>' : ''}<sz val="${s.fs || 11}"/>${s.fg ? `<color rgb="${argb(s.fg)}"/>` : ''}<name val="${xe(s.ff || 'Calibri')}"/></font>`
    const fi = s.bg ? idOf(fills, `<fill><patternFill patternType="solid"><fgColor rgb="${argb(s.bg)}"/><bgColor indexed="64"/></patternFill></fill>`) : 0
    const side = (n, v) => (v ? `<${n} style="${v === 'thick' ? 'thick' : v === 'dashed' ? 'dashed' : 'thin'}"><color rgb="${argb(s.bc || '000000')}"/></${n}>` : `<${n}/>`)
    const b = s.bd ? idOf(borders, `<border>${side('left', s.bd.l)}${side('right', s.bd.r)}${side('top', s.bd.t)}${side('bottom', s.bd.b)}<diagonal/></border>`) : 0
    let nf = 0
    if (s.nf && s.nf !== 'General') {
      if (BUILTIN[s.nf]) nf = BUILTIN[s.nf]
      else {
        const i = numFmts.indexOf(s.nf)
        nf = 164 + (i < 0 ? numFmts.push(s.nf) - 1 : i)
      }
    }
    const al = s.al || s.va || s.wrap || s.ind ? `<alignment${s.al ? ` horizontal="${s.al}"` : ''}${s.va ? ` vertical="${s.va === 'middle' ? 'center' : s.va}"` : ''}${s.wrap ? ' wrapText="1"' : ''}${s.ind ? ` indent="${s.ind}"` : ''}/>` : ''
    const x = `<xf numFmtId="${nf}" fontId="${idOf(fonts, f)}" fillId="${fi}" borderId="${b}" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1"${al ? ' applyAlignment="1">' + al + '</xf>' : '/>'}`
    const id = xfs.push(x) - 1
    seen.set(key, id)
    return id
  }
  const sheetsXml = book.sheets.map((s, si) => {
    const rows = new Map()
    for (const [k, cell] of s.cells) {
      const r = Math.floor(k / 16384)
      if (!rows.has(r)) rows.set(r, [])
      rows.get(r).push([k % 16384, cell])
    }
    let body = ''
    for (const r of [...rows.keys()].sort((a, b) => a - b)) {
      const h = s.rows[String(r + 1)]
      body += `<row r="${r + 1}"${h ? ` ht="${(h * 0.75).toFixed(1)}" customHeight="1"` : ''}>`
      for (const [c, cell] of rows.get(r).sort((a, b) => a[0] - b[0])) {
        const ref = addr(r, c)
        const sid = styleId(cell.s)
        const sa = sid ? ` s="${sid}"` : ''
        if (cell.f != null) {
          const v = book.get(si, r, c)
          const t = typeof v === 'string' ? ' t="str"' : typeof v === 'boolean' ? ' t="b"' : v instanceof XlErr ? ' t="e"' : ''
          const vv = v == null ? '' : v instanceof XlErr ? v.code : typeof v === 'boolean' ? (v ? 1 : 0) : v
          body += `<c r="${ref}"${sa}${t}><f>${xe(cell.f)}</f>${vv === '' ? '' : `<v>${xe(vv)}</v>`}</c>`
        } else if (typeof cell.v === 'number') body += `<c r="${ref}"${sa}><v>${cell.v}</v></c>`
        else if (typeof cell.v === 'boolean') body += `<c r="${ref}"${sa} t="b"><v>${cell.v ? 1 : 0}</v></c>`
        else if (cell.v != null) body += `<c r="${ref}"${sa} t="inlineStr"><is><t xml:space="preserve">${xe(cell.v)}</t></is></c>`
        else body += `<c r="${ref}"${sa}/>`
      }
      body += '</row>'
    }
    const cols = Object.entries(s.cols)
      .map(([k, w]) => ({ i: colIndex(k), w }))
      .sort((a, b) => a.i - b.i)
      .map((c) => `<col min="${c.i + 1}" max="${c.i + 1}" width="${(c.w / 7).toFixed(2)}" customWidth="1"/>`)
      .join('')
    const pane = s.freeze.rows || s.freeze.cols ? `<pane${s.freeze.cols ? ` xSplit="${s.freeze.cols}"` : ''}${s.freeze.rows ? ` ySplit="${s.freeze.rows}"` : ''} topLeftCell="${addr(s.freeze.rows, s.freeze.cols)}" activePane="${s.freeze.rows && s.freeze.cols ? 'bottomRight' : s.freeze.rows ? 'bottomLeft' : 'topRight'}" state="frozen"/>` : ''
    const merges = s.merges.length ? `<mergeCells count="${s.merges.length}">${s.merges.map((m) => `<mergeCell ref="${m}"/>`).join('')}</mergeCells>` : ''
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetViews><sheetView workbookViewId="0"${s.gridlines ? '' : ' showGridLines="0"'}${si === 0 ? ' tabSelected="1"' : ''}>${pane}</sheetView></sheetViews>${cols ? `<cols>${cols}</cols>` : ''}<sheetData>${body}</sheetData>${merges}</worksheet>`
  })
  const styles = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">${numFmts.length ? `<numFmts count="${numFmts.length}">${numFmts.map((f, i) => `<numFmt numFmtId="${164 + i}" formatCode="${xe(f)}"/>`).join('')}</numFmts>` : ''}<fonts count="${fonts.length}">${fonts.join('')}</fonts><fills count="${fills.length}">${fills.join('')}</fills><borders count="${borders.length}">${borders.join('')}</borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="${xfs.length}">${xfs.join('')}</cellXfs></styleSheet>`
  const names = Object.entries(book.names)
    .map(([n, ref]) => `<definedName name="${xe(n)}">${xe(String(ref).replace(/^=/, ''))}</definedName>`)
    .join('')
  const files = {
    '[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>${book.sheets.map((_, i) => `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join('')}</Types>`,
    '_rels/.rels': '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>',
    'xl/workbook.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>${book.sheets.map((s, i) => `<sheet name="${xe(s.name)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join('')}</sheets>${names ? `<definedNames>${names}</definedNames>` : ''}</workbook>`,
    'xl/_rels/workbook.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${book.sheets.map((_, i) => `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`).join('')}<Relationship Id="rId${book.sheets.length + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`,
    'xl/styles.xml': styles,
  }
  sheetsXml.forEach((x, i) => {
    files[`xl/worksheets/sheet${i + 1}.xml`] = x
  })
  return zip(files)
}

// ── xlsx: read ───────────────────────────────────────────────────────────────────────────────────────────────────────────
const unx = (s) => String(s).replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&#(\d+);/g, (m, n) => String.fromCodePoint(+n)).replace(/&#x([0-9a-f]+);/gi, (m, n) => String.fromCodePoint(Number.parseInt(n, 16))).replace(/&amp;/g, '&')
const attr = (tag, name) => {
  const m = tag.match(new RegExp(`\\b${name}="([^"]*)"`))
  return m ? unx(m[1]) : null
}
const BUILTIN_FMT = { 1: '0', 2: '0.00', 3: '#,##0', 4: '#,##0.00', 9: '0%', 10: '0.00%', 14: 'm/d/yyyy', 15: 'd-mmm-yy', 16: 'd-mmm', 17: 'mmm-yy', 18: 'h:mm AM/PM', 19: 'h:mm:ss AM/PM', 20: 'h:mm', 21: 'h:mm:ss', 22: 'm/d/yyyy h:mm', 37: '#,##0;(#,##0)', 38: '#,##0;[Red](#,##0)', 39: '#,##0.00;(#,##0.00)', 40: '#,##0.00;[Red](#,##0.00)', 45: 'mm:ss', 46: '[h]:mm:ss', 47: 'mm:ss.0', 48: '##0.0E+0' }
const rgb = (tag) => {
  const m = tag && attr(tag, 'rgb')
  return m ? `#${m.slice(-6).toLowerCase()}` : ''
}
export function readXlsx(buf, name = 'Imported workbook') {
  const z = unzip(buf)
  if (!z['xl/workbook.xml']) throw new Error('that is not an .xlsx workbook')
  const shared = []
  for (const m of (z['xl/sharedStrings.xml'] || '').matchAll(/<si>([\s\S]*?)<\/si>/g)) shared.push(unx([...m[1].matchAll(/<t[^>]*>([\s\S]*?)<\/t>/g)].map((t) => t[1]).join('')))
  // styles
  const st = z['xl/styles.xml'] || ''
  const custom = {}
  for (const m of st.matchAll(/<numFmt\s[^>]*>/g)) custom[attr(m[0], 'numFmtId')] = attr(m[0], 'formatCode')
  const fonts = [...(st.match(/<fonts[\s\S]*?<\/fonts>/)?.[0] || '').matchAll(/<font>([\s\S]*?)<\/font>|<font\/>/g)].map((m) => {
    const f = m[1] || ''
    return { b: /<b\/>|<b val="1"/.test(f), i: /<i\/>|<i val="1"/.test(f), u: /<u\b/.test(f), fg: rgb(f.match(/<color[^>]*>/)?.[0]), fs: Number(attr(f.match(/<sz[^>]*>/)?.[0] || '', 'val')) || 0 }
  })
  const fills = [...(st.match(/<fills[\s\S]*?<\/fills>/)?.[0] || '').matchAll(/<fill>([\s\S]*?)<\/fill>/g)].map((m) => (/patternType="solid"/.test(m[1]) ? rgb(m[1].match(/<fgColor[^>]*>/)?.[0]) : ''))
  const xfBlock = st.match(/<cellXfs[\s\S]*?<\/cellXfs>/)?.[0] || ''
  const xfs = [...xfBlock.matchAll(/<xf\b([^>]*?)(?:\/>|>([\s\S]*?)<\/xf>)/g)].map((m) => {
    const tag = m[1]
    const s = {}
    const nfid = attr(tag, 'numFmtId')
    const code = custom[nfid] || BUILTIN_FMT[nfid]
    if (code && code !== 'General') s.nf = code
    const f = fonts[Number(attr(tag, 'fontId')) || 0]
    if (f?.b) s.b = true
    if (f?.i) s.i = true
    if (f?.u) s.u = true
    if (f?.fg && f.fg !== '#000000') s.fg = f.fg
    if (f?.fs && f.fs !== 11) s.fs = f.fs
    const bg = fills[Number(attr(tag, 'fillId')) || 0]
    if (bg) s.bg = bg
    const al = m[2]?.match(/<alignment[^>]*>/)?.[0]
    if (al) {
      const h = attr(al, 'horizontal')
      if (h && h !== 'general') s.al = h
      const v = attr(al, 'vertical')
      if (v) s.va = v === 'center' ? 'middle' : v
      if (attr(al, 'wrapText') === '1') s.wrap = true
    }
    return s
  })
  // sheets
  const rels = {}
  for (const m of (z['xl/_rels/workbook.xml.rels'] || '').matchAll(/<Relationship\s[^>]*>/g)) rels[attr(m[0], 'Id')] = attr(m[0], 'Target')
  const sheets = []
  for (const m of z['xl/workbook.xml'].matchAll(/<sheet\s[^>]*>/g)) {
    const target = rels[attr(m[0], 'r:id')] || ''
    const path = target.startsWith('/') ? target.slice(1) : `xl/${target}`
    const xml = z[path]
    if (!xml) continue
    const cells = {}
    for (const c of xml.matchAll(/<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g)) {
      const ref = attr(c[1], 'r')
      const p = ref && parseAddr(ref)
      if (!p) continue
      const t = attr(c[1], 't')
      const body = c[2] || ''
      const f = body.match(/<f[^>]*>([\s\S]*?)<\/f>/)
      const vm = body.match(/<v>([\s\S]*?)<\/v>/)
      const is = body.match(/<is>([\s\S]*?)<\/is>/)
      const out = {}
      if (f && f[1]) out.f = unx(f[1])
      else if (t === 's' && vm) out.v = shared[Number(vm[1])] ?? ''
      else if (t === 'inlineStr' && is) out.v = unx([...is[1].matchAll(/<t[^>]*>([\s\S]*?)<\/t>/g)].map((x) => x[1]).join(''))
      else if (t === 'str' && vm) out.v = unx(vm[1])
      else if (t === 'b' && vm) out.v = vm[1] === '1'
      else if (t === 'e' && vm) out.v = unx(vm[1])
      else if (vm && vm[1] !== '') out.v = Number(vm[1])
      const sx = xfs[Number(attr(c[1], 's')) || 0]
      if (sx && Object.keys(sx).length) out.s = { ...sx }
      if (out.f != null || out.v !== undefined || out.s) cells[ref] = out
    }
    const colsOut = {}
    for (const c of xml.matchAll(/<col\s[^>]*>/g)) {
      const w = Number(attr(c[0], 'width'))
      const lo = Number(attr(c[0], 'min'))
      const hi = Math.min(Number(attr(c[0], 'max')), lo + 60)
      if (w && attr(c[0], 'customWidth') === '1') for (let i = lo; i <= hi; i++) colsOut[colName(i - 1)] = Math.round(w * 7)
    }
    const rowsOut = {}
    for (const r of xml.matchAll(/<row\s[^>]*>/g)) {
      const h = Number(attr(r[0], 'ht'))
      if (h && attr(r[0], 'customHeight') === '1') rowsOut[attr(r[0], 'r')] = Math.round(h / 0.75)
    }
    const pane = xml.match(/<pane\s[^>]*>/)?.[0]
    sheets.push({
      name: attr(m[0], 'name'),
      cells,
      cols: colsOut,
      rows: rowsOut,
      merges: [...xml.matchAll(/<mergeCell\s+ref="([^"]+)"/g)].map((x) => x[1]),
      freeze: pane && attr(pane, 'state') === 'frozen' ? { rows: Number(attr(pane, 'ySplit')) || 0, cols: Number(attr(pane, 'xSplit')) || 0 } : undefined,
      gridlines: !/showGridLines="0"/.test(xml),
    })
  }
  const names = {}
  for (const m of z['xl/workbook.xml'].matchAll(/<definedName\s+name="([^"]+)"[^>]*>([\s\S]*?)<\/definedName>/g)) if (!m[1].startsWith('_xlnm')) names[unx(m[1])] = unx(m[2])
  if (!sheets.length) throw new Error('that workbook has no sheets')
  return new Book({ name, sheets, names })
}
void errByCode
void formatValue
