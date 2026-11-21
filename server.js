#!/usr/bin/env bun
// Zod Sheets: a spreadsheet and database workbench as an MCP App. Tools for the model, a `ui://sheets/app` page for the person.
//   stdio   bun server.js                          Zod starts it (`bunx github:OWNER/sheets`) and talks over stdin/stdout
//   HTTP    PORT=8080 MCP_TOKEN=secret bun server.js   you host it (Docker, Dokploy, a VPS) and Zod connects by address
// Data lives in SHEETS_DIR (default ~/.zod-sheets): workbooks as JSON, database connections encrypted.
import { createInterface } from 'node:readline'
import { TOOLS, UI_URI } from './src/tools.js'
import { buildPage } from './src/page.js'
import { closeAll } from './src/db.js'

const MIME = 'text/html;profile=mcp-app'
const INSTRUCTIONS = `Sheets is a spreadsheet (formulas, formatting, charts, many sheets, 230+ functions) with a database workbench (SQLite, PostgreSQL, MySQL).
- open_sheets shows the app. Everything the person edits on screen is saved; what you write with tools appears there within a few seconds.
- Work on cells with write_cells (text starting with = is a formula), read_range, format_range, sort_range, add_chart, conditional_format, data_validation. Check the "errors" a write returns and fix them.
- Prefer formulas over pasted numbers so the sheet stays live. Use evaluate to test a formula first.
- Databases: db_connections, db_schema, db_query (read-only), db_execute (asks the person), db_query_to_sheet, sheet_to_table. New remote connections are added by the person in the app so passwords stay out of the chat.`

export async function handle(m) {
  if (m.id == null) return null
  const ok = (result) => ({ jsonrpc: '2.0', id: m.id, result })
  const fail = (code, message) => ({ jsonrpc: '2.0', id: m.id, error: { code, message } })
  switch (m.method) {
    case 'initialize':
      return ok({ protocolVersion: '2025-03-26', capabilities: { tools: {}, resources: {} }, serverInfo: { name: 'zod-sheets', version: '1.0.0' }, instructions: INSTRUCTIONS })
    case 'ping':
      return ok({})
    case 'tools/list':
      return ok({ tools: TOOLS.map(({ run, ...t }) => t) })
    case 'resources/list':
      return ok({ resources: [{ uri: UI_URI, name: 'Sheets', title: 'Sheets', description: 'Spreadsheet and database workbench', mimeType: MIME }] })
    case 'resources/read':
      return m.params?.uri === UI_URI ? ok({ contents: [{ uri: UI_URI, mimeType: MIME, text: buildPage() }] }) : fail(-32002, 'no such resource')
    case 'tools/call': {
      const t = TOOLS.find((x) => x.name === m.params?.name)
      if (!t) return ok({ isError: true, content: [{ type: 'text', text: `no tool ${m.params?.name}` }] })
      try {
        return ok(await t.run(m.params.arguments || {}))
      } catch (e) {
        return ok({ isError: true, content: [{ type: 'text', text: String(e?.message || e) }] })
      }
    }
    default:
      return fail(-32601, `unknown method ${m.method}`)
  }
}

if (import.meta.main) {
  const port = Number(process.env.PORT || 0)
  if (port) {
    const token = process.env.MCP_TOKEN || ''
    Bun.serve({
      port,
      hostname: '0.0.0.0',
      async fetch(req) {
        const url = new URL(req.url)
        if (url.pathname === '/health') return new Response('ok')
        if (url.pathname !== '/mcp') return new Response('not found', { status: 404 })
        if (token && req.headers.get('authorization') !== `Bearer ${token}`) return new Response('unauthorized', { status: 401 })
        if (req.method !== 'POST') return new Response('POST only', { status: 405 })
        const body = await req.json()
        if (Array.isArray(body)) {
          const out = (await Promise.all(body.map(handle))).filter(Boolean)
          return out.length ? Response.json(out) : new Response(null, { status: 202 })
        }
        const reply = await handle(body)
        return reply ? Response.json(reply) : new Response(null, { status: 202 })
      },
    })
    console.error(`zod-sheets on http://0.0.0.0:${port}/mcp${token ? ' (token required)' : ' (NO TOKEN: set MCP_TOKEN)'}`)
  } else {
    let chain = Promise.resolve()
    createInterface({ input: process.stdin }).on('line', (line) => {
      chain = chain.then(async () => {
        try {
          const r = await handle(JSON.parse(line))
          if (r) process.stdout.write(`${JSON.stringify(r)}\n`)
        } catch {
          /* not JSON: ignore */
        }
      })
    })
    process.on('exit', closeAll)
  }
}
