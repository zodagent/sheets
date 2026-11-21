import { expect, test } from 'bun:test'
import { Book, XlErr, colName, colIndex, parseAddr, parseRange, shiftFormula, FUNCTION_NAMES, formatValue } from '../src/index.js'

const sheet = (cells, extra = {}) => {
  const b = new Book({ sheets: [{ name: 'Sheet1', cells: Object.fromEntries(Object.entries(cells).map(([a, v]) => [a, typeof v === 'string' && v[0] === '=' ? { f: v.slice(1) } : { v }])) }, ...(extra.more || [])] })
  b.recalc()
  return b
}
const val = (b, a, sh = 0) => {
  const p = parseAddr(a)
  const v = b.get(sh, p.r, p.c)
  return v instanceof XlErr ? v.code : v
}
const calc = (f, cells = {}) => {
  const b = sheet({ ...cells, Z99: `=${f}` })
  return val(b, 'Z99')
}

test('addresses', () => {
  expect(colName(0)).toBe('A')
  expect(colName(27)).toBe('AB')
  expect(colIndex('XFD')).toBe(16383)
  expect(parseAddr('$B$3')).toEqual({ r: 2, c: 1 })
  expect(parseRange('B2:A1')).toMatchObject({ r1: 0, c1: 0, r2: 1, c2: 1 })
  expect(shiftFormula('A1+$B$2+C:C', 1, 1)).toBe('B2+$B$2+D:D')
})

test('arithmetic, precedence, text, comparison', () => {
  expect(calc('1+2*3')).toBe(7)
  expect(calc('-2^2')).toBe(4)
  expect(calc('2^3^2')).toBe(64)
  expect(calc('10/4')).toBe(2.5)
  expect(calc('50%')).toBe(0.5)
  expect(calc('"a"&"b"&1')).toBe('ab1')
  expect(calc('1/0')).toBe('#DIV/0!')
  expect(calc('"x">5')).toBe(true)
  expect(calc('1=1.0')).toBe(true)
  expect(calc('NOSUCH(1)')).toBe('#NAME?')
  expect(calc('1+')).toBe('#ERROR!')
})

test('references, ranges, cross-sheet, names', () => {
  const b = new Book({
    names: { Rate: 'Sheet1!$B$1' },
    sheets: [
      { name: 'Sheet1', cells: { A1: { v: 10 }, A2: { v: 20 }, A3: { v: 30 }, B1: { v: 0.5 }, C1: { f: 'SUM(A1:A3)*Rate' }, C2: { f: "'My Data'!A1+1" } } },
      { name: 'My Data', cells: { A1: { v: 99 } } },
    ],
  })
  b.recalc()
  expect(val(b, 'C1')).toBe(30)
  expect(val(b, 'C2')).toBe(100)
})

test('circular references are flagged, not hung', () => {
  const b = sheet({ A1: '=B1+1', B1: '=A1+1' })
  expect(val(b, 'A1')).toBe('#CIRC!')
})

test('aggregates and criteria', () => {
  const d = { A1: 1, A2: 2, A3: 'x', A4: 4, B1: 'a', B2: 'b', B3: 'a', B4: 'a', C1: 10, C2: 20, C3: 30, C4: 40 }
  expect(calc('SUM(A1:A4)', d)).toBe(7)
  expect(calc('AVERAGE(A1:A4)', d)).toBeCloseTo(7 / 3)
  expect(calc('COUNT(A1:A4)', d)).toBe(3)
  expect(calc('COUNTA(A1:A4)', d)).toBe(4)
  expect(calc('SUMIF(B1:B4,"a",C1:C4)', d)).toBe(80)
  expect(calc('COUNTIF(C1:C4,">15")', d)).toBe(3)
  expect(calc('SUMIFS(C1:C4,B1:B4,"a",C1:C4,">10")', d)).toBe(70)
  expect(calc('COUNTIF(B1:B4,"?")', d)).toBe(4)
  expect(calc('MAXIFS(C1:C4,B1:B4,"a")', d)).toBe(40)
  expect(calc('SUMPRODUCT(A1:A2,C1:C2)', d)).toBe(50)
  expect(calc('MEDIAN(C1:C4)', d)).toBe(25)
  expect(calc('LARGE(C1:C4,2)', d)).toBe(30)
  expect(calc('STDEV(C1:C4)', d)).toBeCloseTo(12.9099, 3)
  expect(calc('SUM(A:A)', d)).toBe(7)
})

test('logic and errors', () => {
  expect(calc('IF(1>2,"a","b")')).toBe('b')
  expect(calc('IFERROR(1/0,"oops")')).toBe('oops')
  expect(calc('IFS(1>2,"a",2>1,"b")')).toBe('b')
  expect(calc('SWITCH(2,1,"one",2,"two","other")')).toBe('two')
  expect(calc('AND(TRUE,1,NOT(FALSE))')).toBe(true)
  expect(calc('ISERROR(1/0)')).toBe(true)
  expect(calc('ISNUMBER("a")')).toBe(false)
  expect(calc('LET(x,5,y,x*2,x+y)')).toBe(15)
})

test('text functions', () => {
  expect(calc('UPPER("abc")&LEFT("hello",2)&RIGHT("hello",3)&MID("hello",2,3)')).toBe('ABChelloell')
  expect(calc('SUBSTITUTE("a-b-c","-","+",2)')).toBe('a-b+c')
  expect(calc('TEXT(1234.5,"#,##0.00")')).toBe('1,234.50')
  expect(calc('TEXT(0.256,"0%")')).toBe('26%')
  expect(calc('TEXTJOIN(", ",TRUE,"a","","b")')).toBe('a, b')
  expect(calc('TRIM("  a   b ")')).toBe('a b')
  expect(calc('FIND("l","hello")')).toBe(3)
  expect(calc('PROPER("hello wORLD")')).toBe('Hello World')
  expect(calc('LEN(REPT("ab",3))')).toBe(6)
  expect(calc('VALUE("1,234.5")')).toBe(1234.5)
})

test('lookup', () => {
  const d = { A1: 'k1', B1: 10, A2: 'k2', B2: 20, A3: 'k3', B3: 30 }
  expect(calc('VLOOKUP("k2",A1:B3,2,FALSE)', d)).toBe(20)
  expect(calc('VLOOKUP("zz",A1:B3,2,FALSE)', d)).toBe('#N/A')
  expect(calc('INDEX(B1:B3,MATCH("k3",A1:A3,0))', d)).toBe(30)
  expect(calc('XLOOKUP("k1",A1:A3,B1:B3)', d)).toBe(10)
  expect(calc('XLOOKUP("none",A1:A3,B1:B3,"-")', d)).toBe('-')
  expect(calc('SUM(OFFSET(B1,1,0,2,1))', d)).toBe(50)
  expect(calc('INDIRECT("B2")', d)).toBe(20)
  expect(calc('VLOOKUP(25,B1:B3,1,TRUE)', d)).toBe(20)
})

test('dates', () => {
  expect(calc('DATE(2024,2,29)')).toBe(45351)
  expect(calc('YEAR(DATE(2024,14,1))')).toBe(2025)
  expect(calc('EOMONTH(DATE(2024,1,15),1)')).toBe(45351)
  expect(calc('DATEDIF(DATE(2020,1,1),DATE(2024,3,1),"Y")')).toBe(4)
  expect(calc('WEEKDAY(DATE(2024,1,1),2)')).toBe(1)
  expect(calc('NETWORKDAYS(DATE(2024,1,1),DATE(2024,1,14))')).toBe(10)
  expect(calc('TEXT(DATE(2024,3,5),"dd mmm yyyy")')).toBe('05 Mar 2024')
})

test('finance', () => {
  expect(calc('PMT(0.05/12,360,200000)')).toBeCloseTo(-1073.64, 2)
  expect(calc('FV(0.06,10,-100)')).toBeCloseTo(1318.08, 2)
  expect(calc('NPV(0.1,100,200,300)')).toBeCloseTo(481.59, 2)
  expect(calc('IRR(A1:A4)', { A1: -100, A2: 50, A3: 50, A4: 50 })).toBeCloseTo(0.2337, 3)
  expect(calc('IPMT(0.05/12,1,360,200000)')).toBeCloseTo(-833.33, 2)
  expect(calc('PPMT(0.05/12,1,360,200000)')).toBeCloseTo(-240.31, 2)
})

test('dynamic arrays spill', () => {
  const d = { A1: 3, A2: 1, A3: 2, B1: 'x', B2: 'y', B3: 'z' }
  const b = sheet({ ...d, D1: '=SORT(A1:A3)', F1: '=SEQUENCE(2,2)', H1: '=UNIQUE(A1:A3)', J1: '=FILTER(B1:B3,A1:A3>1)', L1: '=SUM(D1#)' })
  expect([val(b, 'D1'), val(b, 'D2'), val(b, 'D3')]).toEqual([1, 2, 3])
  expect([val(b, 'F1'), val(b, 'G1'), val(b, 'F2'), val(b, 'G2')]).toEqual([1, 2, 3, 4])
  expect([val(b, 'J1'), val(b, 'J2')]).toEqual(['x', 'z'])
  expect(val(b, 'L1')).toBe(6)
  const blocked = sheet({ A1: 3, A2: 1, D1: '=SORT(A1:A2)', D2: 'in the way' })
  expect(val(blocked, 'D1')).toBe('#SPILL!')
  expect(calc('SUM({1,2;3,4})')).toBe(10)
  expect(calc('SUM(A1:A3*2)', { A1: 1, A2: 2, A3: 3 })).toBe(12)
})

test('inserting and deleting rows rewrites formulas, ranges and other sheets', () => {
  const b = new Book({
    sheets: [
      { name: 'Sheet1', cells: { A1: { v: 1 }, A2: { v: 2 }, A3: { v: 3 }, B1: { f: 'SUM(A1:A3)' }, B2: { f: 'A3*2' } }, merges: ['A2:B2'], cf: [{ type: 'cell', op: '>', a: 1, range: 'A1:A3', style: {} }] },
      { name: 'Two', cells: { A1: { f: 'Sheet1!A3+Sheet1!B1' } } },
    ],
  })
  const sid = b.sheets[0].id
  const undo = b.apply({ op: 'insert', sheet: sid, dim: 'row', at: 1, n: 2 })
  b.recalc()
  expect(b.sheets[0].cells.size).toBe(5)
  expect(b.sheets[0].cells.get(0 * 16384 + 1).f).toBe('SUM(A1:A5)')
  expect(b.sheets[1].cells.get(0).f).toBe('Sheet1!A5+Sheet1!B1')
  expect(b.sheets[0].merges).toEqual(['A4:B4'])
  expect(b.sheets[0].cf[0].range).toBe('A1:A5')
  b.apply(undo)
  b.recalc()
  expect(b.sheets[1].cells.get(0).f).toBe('Sheet1!A3+Sheet1!B1')
  // deleting a row a formula points at gives #REF!, and undo brings it back
  const un2 = b.apply({ op: 'delete', sheet: sid, dim: 'row', at: 2, n: 1 })
  b.recalc()
  expect(b.sheets[0].cells.get(1 * 16384 + 1).f).toBe('#REF!*2')
  expect(b.sheets[0].cells.get(0 * 16384 + 1).f).toBe('SUM(A1:A2)')
  b.apply(un2)
  expect(b.sheets[0].cells.get(1 * 16384 + 1).f).toBe('A3*2')
})

test('columns, rename a sheet, and every op has an inverse', () => {
  const b = new Book({ sheets: [{ name: 'Data', cells: { A1: { v: 5 }, C1: { f: 'A1+B1' } } }, { name: 'Calc', cells: { A1: { f: 'Data!C1' } } }] })
  const before = JSON.stringify(b.toJSON())
  const ops = [
    { op: 'insert', sheet: 'Data', dim: 'col', at: 1, n: 1 },
    { op: 'renameSheet', id: 'Data', name: 'My Data' },
    { op: 'set', sheet: 'My Data', cells: { A1: { v: 9, s: { b: true } } } },
    { op: 'prop', sheet: 'My Data', key: 'freeze', value: { rows: 1, cols: 0 } },
    { op: 'addSheet', sheet: { name: 'New' } },
  ]
  const inv = b.applyAll(ops)
  b.recalc()
  expect(b.sheets[1].cells.get(0).f).toBe("'My Data'!D1")
  for (const o of inv) b.apply(o)
  expect(JSON.stringify(b.toJSON())).toBe(before.replace(/"updated":\d+/, `"updated":${b.updated}`).replace(/"id":"[^"]+"/g, (m) => m))
})

test('sort, fill, paste and find plans', () => {
  const b = sheet({ A1: 'n', A2: 3, A3: 1, A4: 2, B2: '=A2*10', B3: '=A3*10', B4: '=A4*10' })
  const sid = b.sheets[0].id
  b.apply(b.planSort(sid, 'A1:B4', [{ col: 0, asc: true }], { header: true }))
  b.recalc()
  expect([val(b, 'A2'), val(b, 'A3'), val(b, 'A4'), val(b, 'B2')]).toEqual([1, 2, 3, 10])
  const f = sheet({ A1: 1, A2: 2 })
  f.apply(f.planFill(f.sheets[0].id, { r1: 0, c1: 0, r2: 1, c2: 0 }, { r1: 2, c1: 0, r2: 4, c2: 0 }))
  f.recalc()
  expect([val(f, 'A3'), val(f, 'A5')]).toEqual([3, 5])
  const g = sheet({ A1: '=B1*2', B1: 4 })
  g.apply(g.planFill(g.sheets[0].id, { r1: 0, c1: 0, r2: 0, c2: 0 }, { r1: 1, c1: 0, r2: 2, c2: 0 }))
  expect(g.sheets[0].cells.get(2 * 16384).f).toBe('B3*2')
  const r = b.planFind(null, 'x', 'y')
  expect(r.matches.length).toBe(0)
})

test('number formats', () => {
  expect(formatValue(1234567.891, '#,##0.00')).toBe('1,234,567.89')
  expect(formatValue(-5, '#,##0;(#,##0)')).toBe('(5)')
  expect(formatValue(0.5, '0%')).toBe('50%')
  expect(formatValue(1e12, 'General')).toBe('1E+12')
  expect(formatValue(1234567, '0.00E+00')).toBe('1.23E+06')
  expect(formatValue(45000, 'yyyy-mm-dd')).toBe('2023-03-15')
})

test('the function list is big', () => {
  expect(FUNCTION_NAMES().length).toBeGreaterThan(220)
})
