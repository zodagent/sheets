// Databases, in the manner of Beekeeper Studio: saved connections (SQLite file, PostgreSQL, MySQL/MariaDB), a schema browser,
// a SQL runner, table data that can be edited, and sheets that can be written to a table. Uses Bun's built-in SQL client,
// so nothing is installed. Passwords are encrypted on disk and never leave this file: lists show `hasPassword`, not the secret.
import { SQL } from 'bun'
import { createCipheriv, createDecipheriv, randomBytes } from 'node:crypto'
import { chmodSync, existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs'
import { join, resolve } from 'node:path'
import { homedir } from 'node:os'
import { parseTyped, toSerial } from './format.js'

export const TYPES = { sqlite: 'SQLite', postgres: 'PostgreSQL', mysql: 'MySQL / MariaDB' }
const DEFAULT_PORT = { postgres: 5432, mysql: 3306 }
export const dataDir = () => {
  const d = process.env.SHEETS_DIR || join(homedir(), '.zod-sheets')
  mkdirSync(d, { recursive: true })
  return d
}

// ── saved connections ────────────────────────────────────────────────────────────────────────────────────────────────────
const keyFile = () => join(dataDir(), '.key')
function key() {
  const f = keyFile()
  if (!existsSync(f)) {
    writeFileSync(f, randomBytes(32).toString('hex'), { mode: 0o600 })
    try {
      chmodSync(f, 0o600)
    } catch {
      /* not every file system has modes */
    }
  }
  return Buffer.from(readFileSync(f, 'utf8').trim(), 'hex')
}
const seal = (text) => {
  const iv = randomBytes(12)
  const c = createCipheriv('aes-256-gcm', key(), iv)
  const enc = Buffer.concat([c.update(String(text), 'utf8'), c.final()])
  return `enc:${Buffer.concat([iv, c.getAuthTag(), enc]).toString('base64')}`
}
const unseal = (blob) => {
  if (!String(blob).startsWith('enc:')) return String(blob || '')
  const b = Buffer.from(String(blob).slice(4), 'base64')
  const d = createDecipheriv('aes-256-gcm', key(), b.subarray(0, 12))
  d.setAuthTag(b.subarray(12, 28))
  return Buffer.concat([d.update(b.subarray(28)), d.final()]).toString('utf8')
}
const file = () => join(dataDir(), 'connections.json')
const readAll = () => {
  try {
    return JSON.parse(readFileSync(file(), 'utf8'))
  } catch {
    return []
  }
}
const writeAll = (list) => {
  writeFileSync(file(), JSON.stringify(list, null, 2), { mode: 0o600 })
  try {
    chmodSync(file(), 0o600)
  } catch {
    /* ignore */
  }
}
const pub = (c) => {
  const { password, ...rest } = c
  return { ...rest, hasPassword: !!password, label: describe(c) }
}
export function describe(c) {
  if (c.type === 'sqlite') return c.file
  return `${c.user ? `${c.user}@` : ''}${c.host}:${c.port || DEFAULT_PORT[c.type]}/${c.database || ''}`
}
export const listConnections = () => readAll().map(pub)
export const getConnection = (idOrName) => {
  const t = String(idOrName ?? '')
  const c = readAll().find((x) => x.id === t) || readAll().find((x) => x.name.toLowerCase() === t.toLowerCase())
  if (!c) throw new Error(`no saved connection "${idOrName}". Use db_connections to list them.`)
  return c
}
function clean(input, old = {}) {
  const type = input.type || old.type
  if (!TYPES[type]) throw new Error(`type must be one of: ${Object.keys(TYPES).join(', ')}`)
  const c = {
    id: old.id || Math.random().toString(36).slice(2, 10),
    name: String(input.name ?? old.name ?? '').trim().slice(0, 60),
    type,
    color: input.color ?? old.color ?? '',
    readOnly: input.readOnly ?? old.readOnly ?? false,
  }
  if (!c.name) throw new Error('give the connection a name')
  if (type === 'sqlite') {
    const f = String(input.file ?? old.file ?? '').trim()
    if (!f) throw new Error('a SQLite connection needs a file path (use :memory: for a scratch database)')
    c.file = f === ':memory:' ? f : resolve(f.replace(/^~(?=\/|$)/, homedir()))
  } else {
    c.host = String(input.host ?? old.host ?? '').trim()
    if (!c.host) throw new Error('a host is needed')
    c.port = Number(input.port ?? old.port) || DEFAULT_PORT[type]
    c.database = String(input.database ?? old.database ?? '').trim()
    c.user = String(input.user ?? old.user ?? '').trim()
    c.ssl = input.ssl ?? old.ssl ?? false
    // an empty password in an edit keeps the saved one
    const pw = input.password
    c.password = pw === undefined || pw === '' ? old.password || '' : seal(pw)
  }
  return c
}
export function saveConnection(input) {
  const all = readAll()
  const old = input.id ? all.find((x) => x.id === input.id) : null
  const c = clean(input, old || {})
  const clash = all.find((x) => x.id !== c.id && x.name.toLowerCase() === c.name.toLowerCase())
  if (clash) throw new Error(`there is already a connection called ${c.name}`)
  writeAll(old ? all.map((x) => (x.id === c.id ? c : x)) : [...all, c])
  closeConnection(c.id)
  return pub(c)
}
export function deleteConnection(idOrName) {
  const c = getConnection(idOrName)
  writeAll(readAll().filter((x) => x.id !== c.id))
  closeConnection(c.id)
  return { deleted: c.name }
}

// ── live clients ─────────────────────────────────────────────────────────────────────────────────────────────────────────
const open = new Map()
export function closeConnection(id) {
  for (const k of [id, `${id}:ro`]) {
    const s = open.get(k)
    if (s) {
      open.delete(k)
      Promise.resolve(s.close?.()).catch(() => {})
    }
  }
}
export function closeAll() {
  for (const id of [...open.keys()]) closeConnection(id.replace(/:ro$/, ''))
}
function client(c, readonly = false) {
  const k = readonly && c.type === 'sqlite' ? `${c.id}:ro` : c.id
  if (open.has(k)) return open.get(k)
  let s
  if (c.type === 'sqlite') {
    if (c.file !== ':memory:' && !existsSync(c.file) && readonly) throw new Error(`the file ${c.file} does not exist yet`)
    s = new SQL({ adapter: 'sqlite', filename: c.file, readonly: readonly && c.file !== ':memory:', create: true })
  } else {
    s = new SQL({
      adapter: c.type,
      hostname: c.host,
      port: c.port,
      database: c.database || undefined,
      username: c.user || undefined,
      password: c.password ? unseal(c.password) : undefined,
      tls: c.ssl ? (c.ssl === 'insecure' ? { rejectUnauthorized: false } : true) : false,
      max: 4,
      connectionTimeout: 15,
      idleTimeout: 60,
    })
  }
  open.set(k, s)
  return s
}
const q = (c, name) => (c.type === 'mysql' ? `\`${String(name).replace(/`/g, '``')}\`` : `"${String(name).replace(/"/g, '""')}"`)
const ph = (c, i) => (c.type === 'postgres' ? `$${i}` : '?')
const clip = (v) => {
  if (v == null) return null
  if (typeof v === 'bigint') return Number.isSafeInteger(Number(v)) ? Number(v) : String(v)
  if (v instanceof Date) return Number.isNaN(v.getTime()) ? null : v.toISOString().replace(/\.000Z$/, 'Z').replace(/T00:00:00Z$/, '')
  if (v instanceof Uint8Array) return `[binary ${v.length} bytes]`
  if (typeof v === 'object') return JSON.stringify(v)
  return v
}

export async function testConnection(input) {
  const c = clean({ ...input, name: input.name || 'test' }, input.id ? getConnection(input.id) : {})
  c.id = `test-${c.id}`
  const t0 = Date.now()
  try {
    const s = client(c)
    const r = await s.unsafe(c.type === 'sqlite' ? 'select sqlite_version() as v' : c.type === 'postgres' ? 'select version() as v' : 'select version() as v')
    return { ok: true, ms: Date.now() - t0, server: String(clip(Object.values(r[0] || {})[0]) || '').slice(0, 120) }
  } catch (e) {
    return { ok: false, error: friendly(e) }
  } finally {
    closeConnection(c.id)
  }
}
const friendly = (e) => {
  const m = String(e?.message || e)
  if (/ECONNREFUSED|CONNECTION_REFUSED/.test(m + (e?.code || ''))) return 'Connection refused: nothing is listening at that host and port.'
  if (/ENOTFOUND|getaddrinfo/.test(m)) return 'That host name could not be found.'
  if (/password|authentication|access denied/i.test(m)) return 'The database refused the user name or password.'
  if (/timeout|timed out/i.test(m + (e?.code || ''))) return 'The connection timed out.'
  return m.replace(/\s+/g, ' ').slice(0, 300)
}

// ── what can be run without asking ───────────────────────────────────────────────────────────────────────────────────────
const READ_START = /^\s*(?:\(\s*)*(select|with|explain|show|describe|desc|pragma\s+(?:table_info|table_xinfo|index_list|index_info|foreign_key_list|database_list|table_list)|values)\b/i
const WRITE_WORDS = /\b(insert|update|delete|drop|alter|create|truncate|replace|merge|grant|revoke|attach|detach|vacuum|reindex|call|copy|set\s+\w+\s*=)\b/i
export function isReadOnlySql(sql) {
  const s = stripSql(sql)
  if (!s) return false
  if (/;\s*\S/.test(s)) return false // more than one statement
  if (!READ_START.test(s)) return false
  if (/^\s*(explain|with)\b/i.test(s) && WRITE_WORDS.test(s)) return false // a data-changing CTE or EXPLAIN ANALYZE of a write
  if (/\binto\b/i.test(s)) return false // SELECT ... INTO creates a table or writes a file
  return true
}
// the SQL with comments and string contents blanked, so keywords inside text are not mistaken for statements
function stripSql(sql) {
  return String(sql)
    .replace(/--[^\n]*/g, ' ')
    .replace(/\/\*[\s\S]*?\*\//g, ' ')
    .replace(/'(?:[^']|'')*'/g, "''")
    .replace(/"(?:[^"]|"")*"/g, '""')
    .trim()
    .replace(/;\s*$/, '')
}

export async function runQuery(idOrName, sql, { params = [], limit = 1000, write = false } = {}) {
  const c = getConnection(idOrName)
  const text = String(sql || '').trim()
  if (!text) throw new Error('write a query first')
  const readOnly = isReadOnlySql(text)
  if (!readOnly && !write) throw new Error('That statement changes data or the schema. Use db_execute (it asks first) for anything that is not a read.')
  if (!readOnly && c.readOnly) throw new Error(`${c.name} is marked read-only, so it only runs reads.`)
  const s = client(c, readOnly)
  const t0 = Date.now()
  try {
    let rows
    let truncated = false
    if (readOnly) {
      const body = text.replace(/;\s*$/, '')
      const wrap = /^\s*(select|with)\b/i.test(body) && !/\blimit\s+\d+\s*(offset\s+\d+)?\s*$/i.test(stripSql(body))
      const run = async (conn) => (wrap ? conn.unsafe(`select * from (${body}) as _zod_q limit ${Number(limit) + 1}`, params) : conn.unsafe(body, params))
      try {
        rows = c.type === 'sqlite' ? await run(s) : await inReadOnly(c, s, run)
      } catch (e) {
        if (!wrap) throw e
        rows = c.type === 'sqlite' ? await s.unsafe(body, params) : await inReadOnly(c, s, (conn) => conn.unsafe(body, params))
      }
      if (rows.length > limit) {
        rows = Array.prototype.slice.call(rows, 0, limit)
        truncated = true
      }
    } else rows = await s.unsafe(text.replace(/;\s*$/, ''), params)
    const list = Array.prototype.slice.call(rows)
    const columns = list.length ? Object.keys(list[0]) : rows.columns?.map?.((x) => x.name) || []
    return {
      columns,
      rows: list.map((r) => columns.map((k) => clip(r[k]))),
      rowCount: list.length,
      affected: readOnly ? 0 : Number(rows.affectedRows ?? rows.count ?? 0),
      truncated,
      ms: Date.now() - t0,
      readOnly,
    }
  } catch (e) {
    throw new Error(friendly(e))
  }
}
// PostgreSQL reads run in a read-only transaction, so the server itself refuses a write. MySQL and the rest rely on isReadOnlySql.
async function inReadOnly(c, s, fn) {
  if (c.type !== 'postgres' || typeof s.begin !== 'function') return fn(s)
  return s.begin('read only', (tx) => fn(tx))
}

// ── schema ───────────────────────────────────────────────────────────────────────────────────────────────────────────────
export async function schema(idOrName) {
  const c = getConnection(idOrName)
  const s = client(c, c.type === 'sqlite' && c.file !== ':memory:' && existsSync(c.file))
  try {
    if (c.type === 'sqlite') {
      const objs = await s.unsafe("select name, type from sqlite_master where type in ('table','view') and name not like 'sqlite_%' order by type, name")
      const tables = []
      for (const o of objs) {
        const cols = await s.unsafe(`pragma table_info(${q(c, o.name)})`)
        const fks = await s.unsafe(`pragma foreign_key_list(${q(c, o.name)})`).catch(() => [])
        tables.push({
          name: o.name,
          kind: o.type,
          columns: cols.map((x) => ({ name: x.name, type: x.type || '', nullable: !x.notnull && !x.pk, pk: x.pk > 0, default: x.dflt_value ?? null, fk: fks.find((f) => f.from === x.name) ? `${fks.find((f) => f.from === x.name).table}.${fks.find((f) => f.from === x.name).to}` : null })),
        })
      }
      return { type: c.type, schemas: [{ name: 'main', tables }] }
    }
    const filter = c.type === 'postgres' ? "table_schema not in ('pg_catalog','information_schema') and table_schema not like 'pg_toast%'" : 'table_schema = database()'
    const cols = await s.unsafe(`select table_schema as sch, table_name as tbl, column_name as col, data_type as typ, is_nullable as nul, column_default as def, ordinal_position as pos from information_schema.columns where ${filter} order by table_schema, table_name, ordinal_position`)
    const kinds = await s.unsafe(`select table_schema as sch, table_name as tbl, table_type as kind from information_schema.tables where ${filter}`)
    const pks = await s.unsafe(
      c.type === 'postgres'
        ? "select kcu.table_schema as sch, kcu.table_name as tbl, kcu.column_name as col from information_schema.table_constraints tc join information_schema.key_column_usage kcu on tc.constraint_name = kcu.constraint_name and tc.table_schema = kcu.table_schema where tc.constraint_type = 'PRIMARY KEY'"
        : "select table_schema as sch, table_name as tbl, column_name as col from information_schema.key_column_usage where constraint_name = 'PRIMARY' and table_schema = database()",
    )
    const pkSet = new Set(pks.map((p) => `${p.sch}.${p.tbl}.${p.col}`))
    const by = new Map()
    for (const k of kinds) by.set(`${k.sch}.${k.tbl}`, { name: k.tbl, sch: k.sch, kind: /view/i.test(k.kind) ? 'view' : 'table', columns: [] })
    for (const x of cols) {
      const t = by.get(`${x.sch}.${x.tbl}`)
      if (t) t.columns.push({ name: x.col, type: x.typ, nullable: x.nul === 'YES', pk: pkSet.has(`${x.sch}.${x.tbl}.${x.col}`), default: x.def ?? null, fk: null })
    }
    const schemas = new Map()
    for (const t of by.values()) {
      if (!schemas.has(t.sch)) schemas.set(t.sch, { name: t.sch, tables: [] })
      schemas.get(t.sch).tables.push({ name: t.name, kind: t.kind, columns: t.columns })
    }
    return { type: c.type, schemas: [...schemas.values()].map((x) => ({ ...x, tables: x.tables.sort((a, b) => a.name.localeCompare(b.name)) })) }
  } catch (e) {
    throw new Error(friendly(e))
  }
}
const qualified = (c, sch, table) => (c.type === 'sqlite' || !sch ? q(c, table) : `${q(c, sch)}.${q(c, table)}`)

// rows of a table, with sort and a simple filter, for the data grid
export async function tableData(idOrName, sch, table, { limit = 200, offset = 0, sort = null, where = '' } = {}) {
  const c = getConnection(idOrName)
  const t = qualified(c, sch, table)
  const order = sort?.col ? ` order by ${q(c, sort.col)} ${sort.desc ? 'desc' : 'asc'}` : ''
  if (where && !/^[\w\s"'`.,=<>!%()_-]+$/.test(where.replace(/'(?:[^']|'')*'/g, "''"))) throw new Error('that filter has characters a plain filter does not take; use the query editor')
  if (/;|--|\/\*|\b(insert|update|delete|drop|alter|create|attach|pragma)\b/i.test(where.replace(/'(?:[^']|'')*'/g, "''"))) throw new Error('a filter can only be a condition, like age > 30 and city = \'Lahore\'')
  const sql = `select * from ${t}${where ? ` where ${where}` : ''}${order} limit ${Math.min(Number(limit) || 200, 5000)} offset ${Math.max(0, Number(offset) || 0)}`
  const r = await runQuery(idOrName, sql, { limit: 5000 })
  let total = null
  try {
    total = Number((await runQuery(idOrName, `select count(*) as n from ${t}${where ? ` where ${where}` : ''}`)).rows[0][0])
  } catch {
    /* a count is nice to have */
  }
  return { ...r, total }
}

// edits from the data grid: [{ type:'update', pk:{id:1}, set:{name:'x'} }, { type:'insert', values:{...} }, { type:'delete', pk:{id:1} }]
export async function applyChanges(idOrName, sch, table, changes) {
  const c = getConnection(idOrName)
  if (c.readOnly) throw new Error(`${c.name} is marked read-only.`)
  const s = client(c)
  const t = qualified(c, sch, table)
  const stmts = []
  for (const ch of changes) {
    const vals = []
    const add = (v) => {
      vals.push(v === undefined ? null : v)
      return ph(c, vals.length)
    }
    const where = (pk) => {
      const keys = Object.keys(pk || {})
      if (!keys.length) throw new Error('this table has no primary key, so a row cannot be edited safely; use SQL')
      return keys.map((k) => (pk[k] === null ? `${q(c, k)} is null` : `${q(c, k)} = ${add(pk[k])}`)).join(' and ')
    }
    if (ch.type === 'update') {
      const set = Object.keys(ch.set || {})
      if (!set.length) continue
      stmts.push({ sql: `update ${t} set ${set.map((k) => `${q(c, k)} = ${add(ch.set[k])}`).join(', ')} where ${where(ch.pk)}`, vals })
    } else if (ch.type === 'insert') {
      const keys = Object.keys(ch.values || {})
      stmts.push({ sql: keys.length ? `insert into ${t} (${keys.map((k) => q(c, k)).join(', ')}) values (${keys.map((k) => add(ch.values[k])).join(', ')})` : `insert into ${t} default values`, vals })
    } else if (ch.type === 'delete') stmts.push({ sql: `delete from ${t} where ${where(ch.pk)}`, vals })
    else throw new Error(`unknown change ${ch.type}`)
  }
  try {
    const run = async (tx) => {
      let n = 0
      for (const st of stmts) {
        const r = await tx.unsafe(st.sql, st.vals)
        n += Number(r.affectedRows ?? r.count ?? 1)
      }
      return n
    }
    const affected = typeof s.begin === 'function' ? await s.begin(run) : await run(s)
    return { applied: stmts.length, affected }
  } catch (e) {
    throw new Error(friendly(e))
  }
}

// ── sheets ↔ tables ──────────────────────────────────────────────────────────────────────────────────────────────────────
// rows from a query as sheet cells: numbers stay numbers, ISO dates become real dates, text stays text
export function cellsFromResult(result, top = { r: 0, c: 0 }, { header = true } = {}) {
  const cells = {}
  const colL = (n) => {
    let s = ''
    let k = n + 1
    while (k > 0) {
      s = String.fromCharCode(65 + ((k - 1) % 26)) + s
      k = Math.floor((k - 1) / 26)
    }
    return s
  }
  const at = (r, c) => `${colL(c)}${r + 1}`
  let r0 = top.r
  if (header) {
    result.columns.forEach((n, j) => {
      cells[at(r0, top.c + j)] = { v: n, s: { b: true, bg: '#eef2f7' } }
    })
    r0++
  }
  result.rows.forEach((row, i) =>
    row.forEach((v, j) => {
      if (v === null || v === undefined) return
      if (typeof v === 'string' && /^\d{4}-\d{2}-\d{2}(?:[ T]\d{2}:\d{2}(?::\d{2})?Z?)?$/.test(v)) {
        const d = parseTyped(v.replace(/Z$/, '').replace('T', ' '))
        if (d && typeof d.v === 'number') {
          cells[at(r0 + i, top.c + j)] = { v: d.v, s: { nf: d.nf || 'yyyy-mm-dd' } }
          return
        }
      }
      cells[at(r0 + i, top.c + j)] = { v }
    }),
  )
  return cells
}
const guessType = (c, values) => {
  const vs = values.filter((v) => v !== null && v !== undefined && v !== '')
  if (vs.length && vs.every((v) => typeof v === 'number' && Number.isInteger(v))) return c.type === 'postgres' ? 'bigint' : c.type === 'mysql' ? 'bigint' : 'integer'
  if (vs.length && vs.every((v) => typeof v === 'number')) return c.type === 'postgres' ? 'double precision' : c.type === 'mysql' ? 'double' : 'real'
  if (vs.length && vs.every((v) => typeof v === 'boolean')) return c.type === 'postgres' ? 'boolean' : c.type === 'mysql' ? 'tinyint(1)' : 'integer'
  return c.type === 'mysql' ? 'text' : 'text'
}
// write rows (first row = column names) to a table, creating it when asked
export async function writeTable(idOrName, sch, table, rows, { create = true, replace = false } = {}) {
  const c = getConnection(idOrName)
  if (c.readOnly) throw new Error(`${c.name} is marked read-only.`)
  if (rows.length < 2) throw new Error('a table needs a header row and at least one data row')
  const names = rows[0].map((n, i) => String(n ?? '').trim() || `column_${i + 1}`)
  if (new Set(names.map((n) => n.toLowerCase())).size !== names.length) throw new Error('the header row has the same column name twice')
  const data = rows.slice(1)
  const s = client(c)
  const t = qualified(c, sch, table)
  try {
    const run = async (tx) => {
      if (replace) await tx.unsafe(`drop table if exists ${t}`)
      if (create) await tx.unsafe(`create table ${replace ? '' : 'if not exists '}${t} (${names.map((n, j) => `${q(c, n)} ${guessType(c, data.map((r) => r[j]))}`).join(', ')})`)
      let n = 0
      for (let i = 0; i < data.length; i += 200) {
        const chunk = data.slice(i, i + 200)
        const vals = []
        const tuples = chunk.map((r) => `(${names.map((_, j) => {
          vals.push(r[j] === undefined ? null : r[j])
          return ph(c, vals.length)
        }).join(', ')})`)
        await tx.unsafe(`insert into ${t} (${names.map((x) => q(c, x)).join(', ')}) values ${tuples.join(', ')}`, vals)
        n += chunk.length
      }
      return n
    }
    const inserted = typeof s.begin === 'function' ? await s.begin(run) : await run(s)
    return { table, columns: names.length, inserted }
  } catch (e) {
    throw new Error(friendly(e))
  }
}
void toSerial

// ── saved queries (what the SQL editor's Saved menu keeps) ────────────────────────────────────────────────────────────────
const qfile = () => join(dataDir(), 'queries.json')
export const listQueries = () => {
  try {
    return JSON.parse(readFileSync(qfile(), 'utf8'))
  } catch {
    return []
  }
}
export function saveQuery(name, sql, connection) {
  const n = String(name || '').trim().slice(0, 80)
  if (!n || !String(sql || '').trim()) throw new Error('a saved query needs a name and some SQL')
  const list = listQueries().filter((q) => q.name !== n)
  list.unshift({ name: n, sql: String(sql).slice(0, 100000), connection: connection || '', at: Date.now() })
  writeFileSync(qfile(), JSON.stringify(list.slice(0, 200), null, 2))
  return list.slice(0, 200)
}
export function removeQuery(name) {
  const list = listQueries().filter((q) => q.name !== name)
  writeFileSync(qfile(), JSON.stringify(list, null, 2))
  return list
}
