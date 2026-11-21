// The same app in an ordinary browser tab. While the server runs under Zod (stdio) it also listens on 127.0.0.1, on a port the system
// picks, and every request must carry a secret token that only this process and the link it hands to the page know. So the full-screen
// tab can be opened from Zod, and no other program or web page can use it. SHEETS_NO_WEB=1 turns it off.
import { randomBytes, timingSafeEqual } from 'node:crypto'

export const web = { url: '', token: randomBytes(16).toString('hex') }
const same = (a, b) => {
  const x = Buffer.from(String(a))
  const y = Buffer.from(String(b))
  return x.length === y.length && timingSafeEqual(x, y)
}
export function startWeb(handle, buildPage) {
  if (process.env.SHEETS_NO_WEB) return null
  try {
    const s = Bun.serve({
      port: Number(process.env.SHEETS_WEB_PORT || 0),
      hostname: '127.0.0.1',
      async fetch(req) {
        const url = new URL(req.url)
        if (url.pathname === '/mcp') {
          if (req.method !== 'POST' || !same(req.headers.get('x-sheets-token') || '', web.token)) return new Response('forbidden', { status: 403 })
          return Response.json(await handle(await req.json()))
        }
        if (url.pathname === '/' && same(url.searchParams.get('t') || '', web.token)) return new Response(buildPage(), { headers: { 'content-type': 'text/html; charset=utf-8', 'cache-control': 'no-store' } })
        return new Response('This link needs the secret part. Open Sheets from Zod to get a fresh one.', { status: 403 })
      },
    })
    web.url = `http://127.0.0.1:${s.port}/?t=${web.token}`
    return s
  } catch {
    return null // a busy port or no permission: the app still works inside Zod
  }
}
