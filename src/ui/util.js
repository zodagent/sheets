// ── helpers: DOM, icons, host bridge, dialogs, menus, toasts ───────────────────────────────────────────────────────────────
const $ = (sel, root = document) => root.querySelector(sel)
const $$ = (sel, root = document) => [...root.querySelectorAll(sel)]
// h('div.cls#id', { onclick, style: {..}, title }, child, 'text', [children])
function h(spec, attrs, ...kids) {
  const m = String(spec).match(/^([\w-]*)((?:[.#][\w-]+)*)$/)
  const el = document.createElement(m[1] || 'div')
  for (const part of m[2].match(/[.#][\w-]+/g) || []) part[0] === '.' ? el.classList.add(part.slice(1)) : (el.id = part.slice(1))
  if (attrs && typeof attrs === 'object' && !(attrs instanceof Node) && !Array.isArray(attrs)) {
    for (const [k, v] of Object.entries(attrs)) {
      if (v === undefined || v === null || v === false) continue
      if (k === 'style' && typeof v === 'object') Object.assign(el.style, v)
      else if (k === 'class') el.className += ` ${v}`
      else if (k.startsWith('on')) el.addEventListener(k.slice(2), v)
      else if (k === 'html') el.innerHTML = v
      else if (k in el && k !== 'list' && k !== 'form') el[k] = v
      else el.setAttribute(k, v === true ? '' : v)
    }
  } else if (attrs !== undefined) kids.unshift(attrs)
  for (const k of kids.flat(Infinity)) if (k !== null && k !== undefined && k !== false) el.append(k instanceof Node ? k : document.createTextNode(String(k)))
  return el
}
const esc = (s) => String(s ?? '').replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[c])
const clamp = (v, lo, hi) => Math.max(lo, Math.min(hi, v))
const debounce = (f, ms) => {
  let t
  return (...a) => {
    clearTimeout(t)
    t = setTimeout(() => f(...a), ms)
  }
}

const ICONS = {
  undo: 'M3 7v6h6 M21 17a9 9 0 0 0-9-9 9 9 0 0 0-6 2.3L3 13',
  redo: 'M21 7v6h-6 M3 17a9 9 0 0 1 9-9 9 9 0 0 1 6 2.3L21 13',
  bold: 'M6 4h8a4 4 0 0 1 0 8H6z M6 12h9a4 4 0 0 1 0 8H6z',
  italic: 'M19 4h-9 M14 20H5 M15 4 9 20',
  underline: 'M6 4v6a6 6 0 0 0 12 0V4 M4 20h16',
  strike: 'M16 4H9a3 3 0 0 0-2.83 4 M14 12a4 4 0 0 1 0 8H6 M4 12h16',
  left: 'M21 6H3 M15 12H3 M17 18H3',
  center: 'M21 6H3 M17 12H7 M19 18H5',
  right: 'M21 6H3 M21 12H9 M21 18H7',
  fill: 'm19 11-8-8-8.6 8.6a2 2 0 0 0 0 2.8l5.2 5.2c.8.8 2 .8 2.8 0L19 11Z M5 2l5 5 M2 13h15',
  textcolor: 'M4 7V4h16v3 M9 20h6 M12 4v16',
  borders: 'M3 3h18v18H3z M3 9h18 M3 15h18 M9 3v18 M15 3v18',
  merge: 'M8 3H5a2 2 0 0 0-2 2v3 M21 8V5a2 2 0 0 0-2-2h-3 M3 16v3a2 2 0 0 0 2 2h3 M16 21h3a2 2 0 0 0 2-2v-3 M8 12h8 M10 10l-2 2 2 2 M14 10l2 2-2 2',
  wrap: 'M3 6h18 M3 12h15a3 3 0 1 1 0 6h-4 M16 16l-2 2 2 2 M3 18h7',
  percent: 'M19 5 5 19 M6.5 9a2.5 2.5 0 1 0 0-5 2.5 2.5 0 0 0 0 5z M17.5 20a2.5 2.5 0 1 0 0-5 2.5 2.5 0 0 0 0 5z',
  dollar: 'M12 2v20 M17 5H9.5a3.5 3.5 0 0 0 0 7h5a3.5 3.5 0 0 1 0 7H6',
  comma: 'M5 18c1 0 2-.8 2-2s-1-2-2-2 M10 6h10 M10 12h10 M10 18h10',
  sortasc: 'm3 8 4-4 4 4 M7 4v16 M11 12h4 M11 16h7 M11 20h10',
  sortdesc: 'm3 16 4 4 4-4 M7 20V4 M11 4h10 M11 8h7 M11 12h4',
  filter: 'M22 3H2l8 9.46V19l4 2v-8.54z',
  freeze: 'M3 3h18v18H3z M3 9h18 M9 3v18',
  chart: 'M3 3v18h18 M7 16V9 M12 16V5 M17 16v-5',
  search: 'm21 21-4.3-4.3 M11 18a7 7 0 1 0 0-14 7 7 0 0 0 0 14z',
  plus: 'M12 5v14 M5 12h14',
  x: 'M18 6 6 18 M6 6l12 12',
  trash: 'M3 6h18 M8 6V4h8v2 M19 6l-1 14H6L5 6',
  database: 'M12 8c4.4 0 8-1.3 8-3s-3.6-3-8-3-8 1.3-8 3 3.6 3 8 3z M4 5v14c0 1.7 3.6 3 8 3s8-1.3 8-3V5 M4 12c0 1.7 3.6 3 8 3s8-1.3 8-3',
  table: 'M3 3h18v18H3z M3 9h18 M3 15h18 M12 3v18',
  play: 'M6 4l14 8-14 8z',
  check: 'M20 6 9 17l-5-5',
  right2: 'm9 18 6-6-6-6',
  down: 'm6 9 6 6 6-6',
  refresh: 'M21 12a9 9 0 1 1-3-6.7 M21 3v6h-6',
  download: 'M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4 M7 10l5 5 5-5 M12 15V3',
  upload: 'M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4 M17 8l-5-5-5 5 M12 3v12',
  plug: 'M9 2v6 M15 2v6 M6 8h12v4a6 6 0 0 1-12 0z M12 18v4',
  more: 'M12 12h.01 M19 12h.01 M5 12h.01',
  copy: 'M9 9h11v11H9z M5 15H4V4h11v1',
  wand: 'M15 4V2 M15 16v-2 M8 9h2 M20 9h2 M17.8 11.8 19 13 M15 9h.01 M17.8 6.2 19 5 M3 21l9-9 M12.2 6.2 11 5',
  history: 'M3 12a9 9 0 1 0 3-6.7L3 8 M3 3v5h5 M12 7v5l4 2',
  bookmark: 'm19 21-7-4-7 4V5a2 2 0 0 1 2-2h10a2 2 0 0 1 2 2z',
  lock: 'M5 11h14v10H5z M8 11V7a4 4 0 0 1 8 0v4',
  info: 'M12 22a10 10 0 1 0 0-20 10 10 0 0 0 0 20z M12 16v-4 M12 8h.01',
  rules: 'M3 3h7v7H3z M14 3h7v7h-7z M3 14h7v7H3z M17.5 14v7 M14 17.5h7',
  validate: 'm9 11 3 3L22 4 M21 12v7a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11',
  tag: 'M12 2H2v10l9.3 9.3a1 1 0 0 0 1.4 0l8.6-8.6a1 1 0 0 0 0-1.4z M7 7h.01',
  sigma: 'M18 7V4H6l6 8-6 8h12v-3',
  file: 'M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z M14 2v6h6',
  maximize: 'M8 3H5a2 2 0 0 0-2 2v3 M21 8V5a2 2 0 0 0-2-2h-3 M3 16v3a2 2 0 0 0 2 2h3 M16 21h3a2 2 0 0 0 2-2v-3',
  minimize: 'M8 3v3a2 2 0 0 1-2 2H3 M21 8h-3a2 2 0 0 1-2-2V3 M3 16h3a2 2 0 0 1 2 2v3 M16 21v-3a2 2 0 0 1 2-2h3',
  external: 'M15 3h6v6 M10 14 21 3 M18 13v6a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h6',
  folder: 'M20 20a2 2 0 0 0 2-2V8a2 2 0 0 0-2-2h-7.9a2 2 0 0 1-1.69-.9L9.6 3.9A2 2 0 0 0 7.93 3H4a2 2 0 0 0-2 2v13a2 2 0 0 0 2 2Z',
  eye: 'M2 12s3.6-7 10-7 10 7 10 7-3.6 7-10 7S2 12 2 12z M12 15a3 3 0 1 0 0-6 3 3 0 0 0 0 6z',
  key: 'M21 2l-2 2m-7.6 7.6a5.5 5.5 0 1 1-7.8 7.8 5.5 5.5 0 0 1 7.8-7.8z m0 0L15.5 7.5m0 0 3 3L22 7l-3-3m-3.5 3.5L19 4',
  edit: 'M12 20h9 M16.5 3.5a2.1 2.1 0 0 1 3 3L7 19l-4 1 1-4z',
  zoom: 'm21 21-4.3-4.3 M11 18a7 7 0 1 0 0-14 7 7 0 0 0 0 14z M8 11h6 M11 8v6',
  grid: 'M3 3h18v18H3z M3 9h18 M3 15h18 M9 3v18 M15 3v18',
  layers: 'm12 2 10 5-10 5L2 7z M2 17l10 5 10-5 M2 12l10 5 10-5',
  sheet: 'M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z M14 2v6h6 M8 13h8 M8 17h8 M8 9h2',
  columns: 'M12 3h7a2 2 0 0 1 2 2v14a2 2 0 0 1-2 2h-7m0-18H5a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h7m0-18v18',
  clipboard: 'M16 4h2a2 2 0 0 1 2 2v14a2 2 0 0 1-2 2H6a2 2 0 0 1-2-2V6a2 2 0 0 1 2-2h2 M9 2h6v4H9z',
}
const icon = (name, size = 16) => {
  const svg = document.createElementNS('http://www.w3.org/2000/svg', 'svg')
  svg.setAttribute('viewBox', '0 0 24 24')
  svg.setAttribute('width', size)
  svg.setAttribute('height', size)
  svg.setAttribute('fill', 'none')
  svg.setAttribute('stroke', 'currentColor')
  svg.setAttribute('stroke-width', '1.8')
  svg.setAttribute('stroke-linecap', 'round')
  svg.setAttribute('stroke-linejoin', 'round')
  svg.classList.add('ico')
  for (const d of (ICONS[name] || '').split(' M').map((s, i) => (i ? `M${s}` : s))) {
    if (!d) continue
    const p = document.createElementNS('http://www.w3.org/2000/svg', 'path')
    p.setAttribute('d', d)
    svg.append(p)
  }
  return svg
}

// ── host bridge: the MCP Apps protocol over postMessage ────────────────────────────────────────────────────────────────────
// Inside Zod the page talks to its host over postMessage. Opened on its own in a tab (the link the server prints) it is the top-level
// page and talks straight to the server, with the secret token from the link.
const STANDALONE = window.parent === window
const TOKEN = new URLSearchParams(location.search).get('t') || ''
const Host = (() => {
  let id = 0
  const pending = new Map()
  const listeners = {}
  const send = (msg) => parent.postMessage({ jsonrpc: '2.0', ...msg }, '*')
  addEventListener('message', (e) => {
    const m = e.data || {}
    if (m.jsonrpc !== '2.0') return
    if (m.id !== undefined && pending.has(m.id) && m.method === undefined) {
      const p = pending.get(m.id)
      pending.delete(m.id)
      return m.error ? p.fail(new Error(m.error.message || 'request failed')) : p.ok(m.result)
    }
    if (m.method) listeners[m.method]?.(m.params || {}, m)
  })
  return {
    on: (method, f) => {
      listeners[method] = f
    },
    request(method, params) {
      if (STANDALONE) {
        if (method === 'ui/initialize') return Promise.resolve({ hostContext: { theme: matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light', styles: { variables: {} } } })
        if (method === 'ui/open-link') return Promise.resolve(window.open(params.url, '_blank', 'noopener'))
        return fetch('/mcp', { method: 'POST', headers: { 'content-type': 'application/json', 'x-sheets-token': TOKEN }, body: JSON.stringify({ jsonrpc: '2.0', id: 1, method, params }) })
          .then((r) => (r.ok ? r.json() : Promise.reject(new Error('this tab lost its connection to Sheets; open it again from Zod'))))
          .then((j) => (j.error ? Promise.reject(new Error(j.error.message)) : j.result))
      }
      return new Promise((ok, fail) => {
        const i = ++id
        pending.set(i, { ok, fail })
        send({ id: i, method, params })
      })
    },
    notify: (method, params) => !STANDALONE && send({ method, params }),
    // call a server tool; resolves to its structuredContent (or text), throws with the tool's own message when it fails
    async tool(name, args = {}) {
      const r = await Host.request('tools/call', { name, arguments: args })
      if (r?.isError) throw new Error((r.content || []).map((c) => c.text).join(' ') || 'that did not work')
      return r?.structuredContent ?? { text: (r?.content || []).map((c) => c.text).join('\n') }
    },
    resize: (height) => Host.notify('ui/notifications/size-changed', { height }),
  }
})()

// ── toasts, menus, dialogs ─────────────────────────────────────────────────────────────────────────────────────────────────
function toast(msg, kind = '') {
  let box = $('#toasts')
  if (!box) box = document.body.appendChild(h('#toasts'))
  const t = h('.toast', { class: kind }, msg)
  box.append(t)
  setTimeout(() => t.classList.add('out'), 3800)
  setTimeout(() => t.remove(), 4300)
}
let openMenu = null
function closeMenu() {
  openMenu?.remove()
  openMenu = null
}
addEventListener('mousedown', (e) => {
  if (openMenu && !openMenu.contains(e.target)) closeMenu()
})
addEventListener('blur', closeMenu)
// items: [{ label, icon, key, onclick, disabled, checked, sep, header, children? }]
function menu(items, x, y, { minWidth = 190 } = {}) {
  closeMenu()
  const m = h('.menu', { style: { minWidth: `${minWidth}px` }, onmousedown: (e) => e.stopPropagation() })
  for (const it of items) {
    if (!it) continue
    if (it.sep) m.append(h('.menu-sep'))
    else if (it.header) m.append(h('.menu-head', it.header))
    else
      m.append(
        h(
          '.menu-item',
          {
            class: it.disabled ? 'disabled' : '',
            onclick: () => {
              if (it.disabled) return
              closeMenu()
              it.onclick?.()
            },
          },
          h('span.mi-ico', it.checked ? icon('check', 14) : it.icon ? icon(it.icon, 15) : ''),
          h('span.mi-label', it.label),
          it.key ? h('span.mi-key', it.key) : '',
        ),
      )
  }
  document.body.append(m)
  const r = m.getBoundingClientRect()
  m.style.left = `${clamp(x, 4, innerWidth - r.width - 4)}px`
  m.style.top = `${clamp(y, 4, innerHeight - r.height - 4)}px`
  openMenu = m
  return m
}
const menuBelow = (el, items, opts) => {
  const r = el.getBoundingClientRect()
  menu(items, r.left, r.bottom + 2, opts)
}
// a modal: body is a node, buttons [{ label, primary, onclick -> false keeps it open }]
function dialog({ title, body, buttons = [], width = 460, onclose }) {
  const back = h('.dlg-back')
  const close = () => {
    back.remove()
    onclose?.()
  }
  const box = h(
    '.dlg',
    { style: { width: `${width}px` }, role: 'dialog' },
    h('.dlg-head', h('span', title), h('button.icon-btn', { title: 'Close', onclick: close }, icon('x', 16))),
    h('.dlg-body', body),
    buttons.length
      ? h(
          '.dlg-foot',
          buttons.map((b) =>
            h(
              'button.btn',
              {
                class: b.primary ? 'primary' : b.danger ? 'danger' : '',
                onclick: async () => {
                  if ((await b.onclick?.()) !== false) close()
                },
              },
              b.label,
            ),
          ),
        )
      : '',
  )
  back.append(box)
  back.addEventListener('mousedown', (e) => e.target === back && close())
  box.addEventListener('keydown', (e) => {
    if (e.key === 'Escape') close()
    if (e.key === 'Enter' && e.target.tagName === 'INPUT' && !e.target.closest('.no-enter')) box.querySelector('.btn.primary')?.click()
  })
  document.body.append(back)
  ;(box.querySelector('input,select,textarea') || box).focus?.()
  return { close, box }
}
const confirmDialog = (title, message, label = 'Delete', danger = true) =>
  new Promise((ok) => {
    let done = false
    dialog({ title, width: 400, body: h('p.msg', message), buttons: [{ label: 'Cancel', onclick: () => ok(false) }, { label, primary: !danger, danger, onclick: () => ((done = true), ok(true)) }], onclose: () => done || ok(false) })
  })
const promptDialog = (title, value = '', label = 'OK', hint = '') =>
  new Promise((ok) => {
    const input = h('input.input', { value, style: { width: '100%' } })
    let done = false
    const d = dialog({ title, width: 400, body: h('div', hint ? h('p.hint', hint) : '', input), buttons: [{ label: 'Cancel' }, { label, primary: true, onclick: () => ((done = true), ok(input.value)) }], onclose: () => done || ok(null) })
    setTimeout(() => (input.focus(), input.select()), 0)
    void d
  })
// a form row: label + control
const field = (label, control, hint) => h('label.field', h('span.f-label', label), control, hint ? h('span.hint', hint) : '')
const sel = (options, value, attrs = {}) =>
  h('select.input', attrs, options.map((o) => (Array.isArray(o) ? h('option', { value: o[0], selected: String(o[0]) === String(value) }, o[1]) : h('option', { value: o, selected: String(o) === String(value) }, o))))
const download = (name, text) => {
  // a sandboxed frame cannot save files; show the text to copy instead
  dialog({ title: name, width: 640, body: h('textarea.input.mono', { readOnly: true, rows: 16, style: { width: '100%' }, value: text, onfocus: (e) => e.target.select() }), buttons: [{ label: 'Copy', primary: true, onclick: () => (copyText(text), toast('Copied'), false) }, { label: 'Close' }] })
}
function copyText(text) {
  const t = h('textarea', { value: text, style: { position: 'fixed', opacity: 0, left: '-9999px' } })
  document.body.append(t)
  t.select()
  try {
    document.execCommand('copy')
  } catch {
    /* the host may block it; the text is in the dialog */
  }
  t.remove()
}
// a CSS colour (hsl(), rgb(), #hex) as { h, s, l } with s and l from 0 to 1, or null
function parseColor(text) {
  const t = String(text || '').trim()
  let r
  let g
  let b
  let m = t.match(/^hsla?\(\s*([\d.]+)(?:deg)?[\s,]+([\d.]+)%[\s,]+([\d.]+)%/i)
  if (m) return { h: +m[1], s: +m[2] / 100, l: +m[3] / 100 }
  if ((m = t.match(/^#([0-9a-f]{3}|[0-9a-f]{6})\b/i))) {
    const x = m[1].length === 3 ? [...m[1]].map((c) => c + c).join('') : m[1]
    ;[r, g, b] = [0, 2, 4].map((i) => Number.parseInt(x.slice(i, i + 2), 16) / 255)
  } else if ((m = t.match(/^rgba?\(\s*([\d.]+)[\s,]+([\d.]+)[\s,]+([\d.]+)/i))) [r, g, b] = [+m[1] / 255, +m[2] / 255, +m[3] / 255]
  else return null
  const max = Math.max(r, g, b)
  const min = Math.min(r, g, b)
  const l = (max + min) / 2
  const d = max - min
  if (!d) return { h: 0, s: 0, l }
  const s = d / (1 - Math.abs(2 * l - 1))
  const h = max === r ? ((g - b) / d) % 6 : max === g ? (b - r) / d + 2 : (r - g) / d + 4
  return { h: (h * 60 + 360) % 360, s, l }
}
const fmtNum = (n) => (typeof n === 'number' ? (Number.isInteger(n) ? String(n) : String(+n.toPrecision(10))) : String(n))
// short axis labels: 12k, 3.5M
const axisNum = (n) => {
  const a = Math.abs(n)
  if (a >= 1e9) return `${+(n / 1e9).toFixed(1)}B`
  if (a >= 1e6) return `${+(n / 1e6).toFixed(1)}M`
  if (a >= 1e4) return `${+(n / 1e3).toFixed(1)}k`
  return fmtNum(+n.toFixed(6))
}
