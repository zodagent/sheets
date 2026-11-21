// A stand-in for the MCP host, for trying the page in a browser: serves a page that embeds the app in a sandboxed frame
// (sandbox="allow-scripts", the same as Zod) and answers its postMessage calls by POSTing to this server's own /mcp.
// bun test/harness.js [port]   then open http://localhost:<port>/
import { handle } from '../server.js'
const port = Number(process.argv[2] || process.env.PORT || 19400)
const page = `<!doctype html><meta charset=utf-8><title>host</title><style>html,body{margin:0;height:100%;background:#cfd6e0}iframe{border:0;width:100%;display:block}</style>
<iframe id=f sandbox="allow-scripts" style="height:100vh"></iframe>
<script>
const f=document.getElementById('f')
fetch('/mcp',{method:'POST',body:JSON.stringify({jsonrpc:'2.0',id:1,method:'resources/read',params:{uri:'ui://sheets/app'}})}).then(r=>r.json()).then(r=>{
  const html=r.result.contents[0].text
  const csp="default-src 'none'; script-src 'unsafe-inline'; style-src 'unsafe-inline'; img-src data: blob:; font-src data:; media-src data: blob:; connect-src 'none'; frame-src 'none'; form-action 'none'; base-uri 'none'"
  f.srcdoc=html.replace('<head>','<head><meta http-equiv="Content-Security-Policy" content="'+csp+'">')
})
let theme=new URLSearchParams(location.search).get('theme')||'light'
addEventListener('message',async e=>{
  if(e.source!==f.contentWindow)return
  const m=e.data||{}
  if(m.jsonrpc!=='2.0'||!m.method)return
  const send=x=>f.contentWindow.postMessage({jsonrpc:'2.0',...x},'*')
  if(m.method==='ui/initialize')return send({id:m.id,result:{protocolVersion:'2025-11-21',hostInfo:{name:'harness'},hostCapabilities:{serverTools:{}},hostContext:{theme,styles:{variables:{}},displayMode:'inline'}}})
  if(m.method==='notifications/message'){document.title='ERR '+JSON.stringify(m.params.data);console.error(m.params.data);return}
  if(m.method==='ui/notifications/size-changed')return
  if(m.method==='tools/call'){
    const r=await fetch('/mcp',{method:'POST',body:JSON.stringify({jsonrpc:'2.0',id:1,method:'tools/call',params:m.params})}).then(r=>r.json())
    return send({id:m.id,result:r.result})
  }
})
</script>`
Bun.serve({
  port,
  hostname: '127.0.0.1',
  async fetch(req) {
    const url = new URL(req.url)
    if (url.pathname === '/mcp') return Response.json(await handle(await req.json()))
    return new Response(page, { headers: { 'content-type': 'text/html' } })
  },
})
console.error(`harness on http://127.0.0.1:${port}/`)
