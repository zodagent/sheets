# Zod Sheets

A spreadsheet and a database workbench, built as an [MCP App](https://modelcontextprotocol.io) for [Zod](https://github.com/zodagent). Zod shows its screens in a sandboxed frame, and the assistant can drive the same workbooks with tools.

- **Spreadsheet**: formulas (230+ functions, dynamic arrays, cross-sheet references, named ranges), number formats, styles, borders, merged cells, frozen panes, conditional formatting, data validation, auto filter, sort, find and replace, charts, undo and redo, fill handle, copy and paste, many sheets. Rendered on a canvas, so large sheets stay fast.
- **Database** (in the manner of Beekeeper Studio): connections to SQLite files, PostgreSQL and MySQL/MariaDB; schema browser; SQL editor with highlighting, autocomplete, history and saved queries; table data you can edit and save; query results sent to a sheet, and sheet ranges written to a table.
- **Files**: CSV in and out, `.xlsx` in and out (formulas, formats, sheets, names), all with no dependencies.
- **For the assistant**: 29 tools to read, write, format, chart, sort, import, export and query. Things you edit on screen and things the assistant writes meet in the same workbook (the screen refreshes within seconds).

Runs on [Bun](https://bun.sh). No packages to install.

## Add it to Zod

**From GitHub** (runs on your computer). In Zod: **Connectors → Add custom → Add from a GitHub link or address**, paste

```
github.com/zodagent/sheets/tree/main
```

Zod runs it with `bunx github:zodagent/sheets#main`. Name the branch (`/tree/main`): Bun cannot fetch a bare `github:` spec. Pin a version with `github.com/zodagent/sheets/tree/v1.0.0`.

**From a clone** (needed if you fork this privately: `bunx github:` fetches an unauthenticated tarball, so it only works for a public repository). Clone it, then in the same form paste a command:

```
bun /path/to/zod-sheets/server.js
```

**Saving without a prompt.** Edits you make on screen are saved by a tool the page calls, and Zod asks Allow / Deny for any tool that changes something. Open **Connectors → Sheets** and set the server to *trusted* (or set `app_apply` to *allow*) so saving is silent. Reads are never asked about.

**Self-hosted** (Docker, Dokploy, a VPS):

```bash
docker build -t zod-sheets .
docker run -p 8080:8080 -v sheets-data:/data -e MCP_TOKEN=choose-a-long-secret zod-sheets
```

In Zod, paste `https://your-host/mcp` into the same form and put the token in **Access token**. Without `MCP_TOKEN` anyone who can reach the port can use the server: always set one. The database files and paths you give it are the ones **on the server**.

## Where things are kept

`SHEETS_DIR` (default `~/.zod-sheets`):

| file | what |
| --- | --- |
| `workbooks/<id>.json` | one workbook each, written atomically |
| `connections.json` | saved database connections; passwords are AES-256-GCM encrypted with `.key` (both mode 600) |
| `queries.json` | saved SQL queries |

Passwords are typed in the app (the **Database** tab), never through the chat, and are never returned to the page or the assistant: lists show only that one is saved. Set `SHEETS_NO_FILES=1` to stop the tools from reading or writing files on the host's disk.

## Database safety

- `db_query` and the editor's Run only run reads (`SELECT`, `WITH`, `EXPLAIN`, `SHOW`, safe `PRAGMA`s). SQLite reads use a read-only connection; PostgreSQL reads run inside a read-only transaction; for MySQL the statement check is the guard, so use a database user without write rights (or mark the connection read-only) when it matters. Several statements in one call, `SELECT … INTO`, and data-changing CTEs are refused.
- Anything else goes through `db_execute` (the assistant) or a confirmation (the app), and Zod asks first as well. A connection marked **read-only** never runs writes.
- Editing table data needs a primary key; changes are applied in one transaction.

## Tools the assistant gets

`open_sheets`, `list_workbooks`, `create_workbook`, `manage_workbook`, `workbook_info`, `read_range`, `write_cells`, `clear_range`, `format_range`, `edit_structure`, `manage_sheets`, `sort_range`, `find_replace`, `fill_range`, `define_name`, `add_chart`, `conditional_format`, `data_validation`, `evaluate`, `function_list`, `import_data`, `export_data`, `db_connections`, `db_add_sqlite`, `db_schema`, `db_query`, `db_execute`, `db_query_to_sheet`, `sheet_to_table`.

Tools whose names start with `app_` are for the page only and are hidden from the model.

## Develop

```bash
bun test                      # engine, formats, database, tools, xlsx round trip
bun run harness               # a stand-in host on http://127.0.0.1:19400 to try the page in a browser
bun server.js                 # stdio, as Zod runs it
PORT=8080 MCP_TOKEN=x bun server.js
```

Layout: `src/format.js` (number formats, dates), `src/core.js` (references, parser, evaluator, workbook, operations), `src/functions.js` (the function library), `src/io.js` (CSV, xlsx), `src/db.js` (connections, schema, queries), `src/tools.js` (MCP tools), `src/ui/` (the page; `src/page.js` bundles the engine and the UI into one HTML file).

Every change to a workbook is a small JSON operation that returns its own inverse, which is what undo, saving from the page and the tools all share.

## Limits

No pivot tables, no cell comments, no macros, no printing. Formulas recalculate in full on each change (fast to tens of thousands of formula cells). The page saves through MCP tool calls, so it needs a host that supports MCP Apps.

## Licence

MIT
