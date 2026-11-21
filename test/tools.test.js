import { afterAll, expect, test } from 'bun:test'
import { mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

const dir = mkdtempSync(join(tmpdir(), 'zod-sheets-tools-'))
process.env.SHEETS_DIR = dir
const { handle } = await import('../server.js')
const db = await import('../src/db.js')
afterAll(() => {
  db.closeAll()
  rmSync(dir, { recursive: true, force: true })
})

let n = 0
const rpc = async (method, params) => (await handle({ jsonrpc: '2.0', id: ++n, method, params })).result
const tool = async (name, args = {}) => {
  const r = await rpc('tools/call', { name, arguments: args })
  return { ...r, text: r.content?.[0]?.text, data: r.structuredContent }
}

test('the server describes itself, lists tools and serves the page', async () => {
  const init = await rpc('initialize', {})
  expect(init.serverInfo.name).toBe('zod-sheets')
  expect(init.instructions).toMatch(/write_cells/)
  const { tools } = await rpc('tools/list')
  const names = tools.map((t) => t.name)
  expect(names.length).toBeGreaterThan(40)
  expect(tools.find((t) => t.name === 'open_sheets')._meta.ui.resourceUri).toBe('ui://sheets/app')
  // pages-only tools are hidden from the model; reads are marked read-only so they never ask
  expect(tools.find((t) => t.name === 'app_apply')._meta.ui.visibility).toEqual(['app'])
  expect(tools.find((t) => t.name === 'db_query').annotations.readOnlyHint).toBe(true)
  expect(tools.find((t) => t.name === 'db_execute').annotations?.readOnlyHint).toBeUndefined()
  expect(tools.find((t) => t.name === 'write_cells').annotations?.readOnlyHint).toBeUndefined()
  const page = await rpc('resources/read', { uri: 'ui://sheets/app' })
  expect(page.contents[0].mimeType).toBe('text/html;profile=mcp-app')
  expect(page.contents[0].text).toMatch(/<div id="root">/)
  expect(page.contents[0].text.length).toBeGreaterThan(100000)
})

test('build a workbook with tools: write, read, errors, format, sort, structure, undoable ops', async () => {
  const c = await tool('create_workbook', { name: 'Budget', sheets: [{ name: 'Data', rows: [['Item', 'Cost'], ['Rent', 1200], ['Food', 450], ['Fun', 150]] }] })
  expect(c.data.sheets[0]).toMatchObject({ name: 'Data', used: 'A1:B4' })
  const w = await tool('write_cells', { workbook: 'Budget', sheet: 'Data', cells: { A5: 'Total', B5: '=SUM(B2:B4)', C1: '=1/0' } })
  expect(w.data.errors).toEqual(['C1 #DIV/0!'])
  expect((await tool('read_range', { workbook: 'Budget', sheet: 'Data', range: 'B5' })).data.values).toEqual([[1800]])
  expect((await tool('read_range', { workbook: 'Budget', sheet: 'Data', range: 'B5', mode: 'formulas' })).data.values).toEqual([['=SUM(B2:B4)']])
  await tool('format_range', { workbook: 'Budget', sheet: 'Data', range: 'B2:B5', style: { number_format: '$#,##0', bold: true } })
  expect((await tool('read_range', { workbook: 'Budget', sheet: 'Data', range: 'B2:B3', mode: 'display' })).data.values).toEqual([['$1,200'], ['$450']])
  await tool('sort_range', { workbook: 'Budget', sheet: 'Data', range: 'A2:B4', keys: [{ column: 'B', ascending: true }] })
  expect((await tool('read_range', { workbook: 'Budget', sheet: 'Data', range: 'A2:A4' })).data.values).toEqual([['Fun'], ['Food'], ['Rent']])
  await tool('edit_structure', { workbook: 'Budget', sheet: 'Data', action: 'insert_rows', at: '2', count: 2 })
  expect((await tool('read_range', { workbook: 'Budget', sheet: 'Data', range: 'B7', mode: 'formulas' })).data.values).toEqual([['=SUM(B4:B6)']])
  expect((await tool('read_range', { workbook: 'Budget', sheet: 'Data', range: 'B7' })).data.values).toEqual([[1800]])
  const f = await tool('find_replace', { workbook: 'Budget', find: 'rent', replace: 'Housing' })
  expect(f.text).toMatch(/1 match/)
  expect((await tool('evaluate', { formula: '=SUMIFS(B:B,A:A,"Housing")', workbook: 'Budget', sheet: 'Data' })).data.value).toBe(1200)
  const bad = await tool('read_range', { workbook: 'Budget', sheet: 'Nope', range: 'A1' })
  expect(bad.isError).toBe(true)
  expect(bad.text).toMatch(/no sheet/)
})

test('a page edit and a model edit both land, and the page can tell something changed', async () => {
  const b = (await tool('app_state', { workbook: 'Budget' })).data
  const sid = b.workbook.sheets[0].id
  const r = await tool('app_apply', { id: b.workbook.id, ops: [{ op: 'set', sheet: sid, cells: { E1: { v: 'from the page' } } }] })
  expect(r.data.rev).toBe(r.data.before + 1)
  const p = await tool('app_poll', { id: b.workbook.id, rev: b.workbook.rev })
  expect(p.data.changed).toBe(true)
  expect(p.data.workbook.sheets[0].cells.E1.v).toBe('from the page')
  expect((await tool('app_poll', { id: b.workbook.id, rev: p.data.rev })).data.changed).toBe(false)
})

test('charts, rules and validation are stored and survive a reload', async () => {
  await tool('add_chart', { workbook: 'Budget', sheet: 'Data', type: 'pie', range: 'A5:B6', title: 'Mix' })
  await tool('conditional_format', { workbook: 'Budget', sheet: 'Data', range: 'B:B', rule: { type: 'cell', op: '>', a: 1000, style: { bg: '#fecaca' } } })
  await tool('data_validation', { workbook: 'Budget', sheet: 'Data', range: 'D1:D9', rule: { type: 'list', items: ['Open', 'Closed'] } })
  const info = (await tool('workbook_info', { workbook: 'Budget' })).data
  expect(info.sheets[0]).toMatchObject({ charts: 1, conditional_formats: 1, validations: 1 })
  const { forget } = await import('../src/store.js')
  forget()
  expect((await tool('workbook_info', { workbook: 'Budget' })).data.sheets[0].charts).toBe(1)
})

test('csv in and out, and an xlsx file survives the round trip with formulas, formats and sheets', async () => {
  await tool('create_workbook', { name: 'Trip' })
  await tool('import_data', { workbook: 'Trip', csv: 'city,pop,date\nLahore,"13,000,000",2024-03-05\nKarachi,16000000,3/6/2024\n', at: 'A1' })
  expect((await tool('read_range', { workbook: 'Trip', range: 'B2:B3' })).data.values).toEqual([[13000000], [16000000]])
  expect((await tool('read_range', { workbook: 'Trip', range: 'C2', mode: 'display' })).data.values).toEqual([['2024-03-05']])
  await tool('write_cells', { workbook: 'Trip', cells: { E1: '=SUM(B2:B3)', E2: 'say "hi", ok' } })
  await tool('format_range', { workbook: 'Trip', range: 'E1', style: { number_format: '#,##0', bold: true, fill: '#ffeeaa' } })
  await tool('manage_sheets', { workbook: 'Trip', action: 'add', name: 'Notes' })
  await tool('write_cells', { workbook: 'Trip', sheet: 'Notes', cells: { A1: "=Sheet1!E1*2" } })
  const csv = (await tool('export_data', { workbook: 'Trip', range: 'A1:E3', format: 'csv' })).text
  expect(csv.split('\n')[0]).toBe('city,pop,date,,"29,000,000"') // exported as shown on screen
  expect(csv).toContain('"say ""hi"", ok"')
  const path = join(dir, 'trip.xlsx')
  const w = await tool('export_data', { workbook: 'Trip', format: 'xlsx', path })
  expect(w.text).toMatch(/Wrote/)
  const imp = await tool('import_data', { path, name: 'Trip again' })
  expect(imp.data.sheets.map((s) => s.name)).toEqual(['Sheet1', 'Notes'])
  expect((await tool('read_range', { workbook: 'Trip again', range: 'E1' })).data.values).toEqual([[29000000]])
  expect((await tool('read_range', { workbook: 'Trip again', range: 'E1', mode: 'formulas' })).data.values).toEqual([['=SUM(B2:B3)']])
  expect((await tool('read_range', { workbook: 'Trip again', range: 'E1', mode: 'display' })).data.values).toEqual([['29,000,000']])
  expect((await tool('read_range', { workbook: 'Trip again', sheet: 'Notes', range: 'A1' })).data.values).toEqual([[58000000]])
  const zip = readFileSync(path)
  expect(zip.subarray(0, 2).toString()).toBe('PK')
})

test('database tools: reads run freely, writes and credentials do not slip through', async () => {
  const file = join(dir, 'shop.db')
  writeFileSync(file, '')
  expect((await tool('db_add_sqlite', { name: 'Shop', file })).text).toMatch(/Saved/)
  expect((await tool('db_execute', { connection: 'Shop', sql: 'create table t (id integer primary key, name text, qty integer)' })).isError).toBeUndefined()
  await tool('db_execute', { connection: 'Shop', sql: "insert into t (name, qty) values ('a', 3), ('b', 5)" })
  const q = await tool('db_query', { connection: 'Shop', sql: 'select name, qty from t order by qty desc' })
  expect(q.data.rows).toEqual([['b', 5], ['a', 3]])
  const w = await tool('db_query', { connection: 'Shop', sql: 'delete from t' })
  expect(w.isError).toBe(true)
  expect((await tool('db_query', { connection: 'Shop', sql: 'select count(*) from t' })).data.rows[0][0]).toBe(2)
  expect((await tool('db_schema', { connection: 'Shop' })).text).toMatch(/t \(table\): id integer PK, name text, qty integer/i)
  await tool('create_workbook', { name: 'DB' })
  await tool('db_query_to_sheet', { connection: 'Shop', sql: 'select * from t', workbook: 'DB', at: 'B2' })
  expect((await tool('read_range', { workbook: 'DB', range: 'B2:D4' })).data.values).toEqual([['id', 'name', 'qty'], [1, 'a', 3], [2, 'b', 5]])
  await tool('write_cells', { workbook: 'DB', cells: { E3: '=SUM(D3:D4)' } })
  await tool('sheet_to_table', { connection: 'Shop', table: 'copy_of_t', workbook: 'DB', range: 'B2:D4' })
  expect((await tool('db_query', { connection: 'Shop', sql: 'select sum(qty) from copy_of_t' })).data.rows[0][0]).toBe(8)
  // there is no model tool that takes a password, and none lists one
  const names = (await rpc('tools/list')).tools.filter((t) => !t._meta?.ui?.visibility).map((t) => t.name)
  expect(names).not.toContain('app_db_save')
  expect(JSON.stringify((await tool('db_connections')).data)).not.toMatch(/password"?:\s*"[^"]/)
})
