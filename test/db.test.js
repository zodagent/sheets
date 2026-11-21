import { afterAll, expect, test } from 'bun:test'
import { mkdtempSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

const dir = mkdtempSync(join(tmpdir(), 'zod-sheets-db-'))
process.env.SHEETS_DIR = dir
const db = await import('../src/db.js')
afterAll(() => {
  db.closeAll()
  rmSync(dir, { recursive: true, force: true })
})

test('read-only detection', () => {
  expect(db.isReadOnlySql('select * from t')).toBe(true)
  expect(db.isReadOnlySql('  WITH a AS (select 1) select * from a;')).toBe(true)
  expect(db.isReadOnlySql("select 'drop table x' as s")).toBe(true)
  expect(db.isReadOnlySql('select 1; drop table t')).toBe(false)
  expect(db.isReadOnlySql('delete from t')).toBe(false)
  expect(db.isReadOnlySql('with d as (delete from t returning *) select * from d')).toBe(false)
  expect(db.isReadOnlySql('select * into backup from t')).toBe(false)
  expect(db.isReadOnlySql('pragma table_info(t)')).toBe(true)
  expect(db.isReadOnlySql('pragma writable_schema=1')).toBe(false)
})

test('a SQLite connection: save, schema, query, edit, write a table, secrets stay hidden', async () => {
  const file = join(dir, 'shop.db')
  const c = db.saveConnection({ name: 'Shop', type: 'sqlite', file })
  expect(db.listConnections()[0]).toMatchObject({ name: 'Shop', type: 'sqlite' })
  expect(() => db.saveConnection({ name: 'shop', type: 'sqlite', file })).toThrow(/already/)
  await db.runQuery('Shop', 'create table items (id integer primary key, name text, price real, made text)', { write: true })
  await db.runQuery('Shop', "insert into items (name, price, made) values ('Pen', 1.5, '2024-03-05'), ('Book', 12, null)", { write: true })
  await expect(db.runQuery('Shop', 'delete from items')).rejects.toThrow(/db_execute/)
  const r = await db.runQuery('Shop', 'select * from items order by id')
  expect(r.columns).toEqual(['id', 'name', 'price', 'made'])
  expect(r.rows).toEqual([[1, 'Pen', 1.5, '2024-03-05'], [2, 'Book', 12, null]])
  const s = await db.schema('Shop')
  expect(s.schemas[0].tables[0]).toMatchObject({ name: 'items', kind: 'table' })
  expect(s.schemas[0].tables[0].columns.find((x) => x.name === 'id').pk).toBe(true)
  const td = await db.tableData(c.id, 'main', 'items', { sort: { col: 'price', desc: true }, where: 'price > 1' })
  expect(td.rows[0][1]).toBe('Book')
  expect(td.total).toBe(2)
  await expect(db.tableData(c.id, 'main', 'items', { where: '1=1; drop table items' })).rejects.toThrow()
  await db.applyChanges(c.id, 'main', 'items', [{ type: 'update', pk: { id: 1 }, set: { price: 2 } }, { type: 'insert', values: { name: 'Ink', price: 3 } }, { type: 'delete', pk: { id: 2 } }])
  expect((await db.runQuery('Shop', 'select name, price from items order by id')).rows).toEqual([['Pen', 2], ['Ink', 3]])
  const w = await db.writeTable('Shop', 'main', 'from_sheet', [['city', 'pop'], ['Lahore', 13000000], ['Karachi', 16000000]])
  expect(w.inserted).toBe(2)
  expect((await db.runQuery('Shop', 'select sum(pop) from from_sheet')).rows[0][0]).toBe(29000000)
  const cells = db.cellsFromResult(r, { r: 0, c: 0 })
  expect(cells.A1.v).toBe('id')
  expect(cells.D2.s.nf).toBe('yyyy-mm-dd')
  db.saveConnection({ name: 'Locked', type: 'sqlite', file, readOnly: true })
  await expect(db.runQuery('Locked', 'insert into items (name) values (1)', { write: true })).rejects.toThrow(/read-only/)
})

test('a remote password is encrypted on disk and never listed', async () => {
  const c = db.saveConnection({ name: 'Prod', type: 'postgres', host: '127.0.0.1', port: 1, database: 'x', user: 'u', password: 'hunter2' })
  expect(JSON.stringify(db.listConnections())).not.toContain('hunter2')
  expect(c.hasPassword).toBe(true)
  expect(await Bun.file(join(dir, 'connections.json')).text()).not.toContain('hunter2')
  const t = await db.testConnection({ id: c.id })
  expect(t.ok).toBe(false)
  expect(t.error).toMatch(/refused|timed out|could not/i)
  db.saveConnection({ id: c.id, name: 'Prod', host: '127.0.0.1' }) // an edit without a password keeps it
  expect(db.listConnections().find((x) => x.name === 'Prod').hasPassword).toBe(true)
})
