// Workbooks on disk: one JSON file each in <data dir>/workbooks, written atomically, kept in memory while the server runs.
import { existsSync, mkdirSync, readFileSync, readdirSync, renameSync, rmSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'
import { Book, newId } from './index.js'
import { dataDir } from './db.js'

const dir = () => {
  const d = join(dataDir(), 'workbooks')
  mkdirSync(d, { recursive: true })
  return d
}
const cache = new Map()

function read(id) {
  const f = join(dir(), `${id}.json`)
  if (!existsSync(f)) return null
  const b = new Book(JSON.parse(readFileSync(f, 'utf8')))
  cache.set(b.id, b)
  return b
}
export function allIds() {
  return readdirSync(dir())
    .filter((f) => f.endsWith('.json'))
    .map((f) => f.slice(0, -5))
}
export function list() {
  return allIds()
    .map((id) => cache.get(id) || read(id))
    .filter(Boolean)
    .map((b) => summary(b))
    .sort((a, b) => b.updated - a.updated)
}
export const summary = (b) => ({ id: b.id, name: b.name, sheets: b.sheets.map((s) => s.name), updated: b.updated, rev: b.rev })

// a workbook by id or name; with no argument, the one touched last (made if there is none)
export function get(ref) {
  if (ref === undefined || ref === null || ref === '') {
    const first = list()[0]
    return first ? get(first.id) : create('My workbook')
  }
  const t = String(ref)
  if (cache.has(t)) return cache.get(t)
  const byId = read(t.replace(/[^\w-]/g, ''))
  if (byId) return byId
  const hit = list().find((x) => x.name.toLowerCase() === t.toLowerCase())
  if (hit) return get(hit.id)
  throw new Error(`no workbook "${ref}". Use list_workbooks to see them, or create_workbook.`)
}
export function save(b) {
  b.rev++
  b.updated = Date.now()
  const f = join(dir(), `${b.id}.json`)
  const tmp = `${f}.${process.pid}.tmp`
  writeFileSync(tmp, JSON.stringify(b.toJSON()))
  renameSync(tmp, f)
  cache.set(b.id, b)
  return b
}
export function create(name, data = {}) {
  const b = new Book({ ...data, id: newId() + newId(), name: String(name || 'Untitled workbook').slice(0, 100) })
  b.rev = 0
  return save(b)
}
export function remove(ref) {
  const b = get(ref)
  rmSync(join(dir(), `${b.id}.json`), { force: true })
  cache.delete(b.id)
  return b.name
}
export function duplicate(ref, name) {
  const b = get(ref)
  const data = b.toJSON()
  data.sheets = data.sheets.map((s) => ({ ...s, id: undefined }))
  return create(name || `${b.name} copy`, { ...data, id: undefined, rev: 0 })
}
export const forget = () => cache.clear()
