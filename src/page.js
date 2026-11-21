// Builds the single HTML page of the app: the engine (format, core, functions) bundled inline, plus the UI files in src/ui.
import { readFileSync } from 'node:fs'
import { join } from 'node:path'

const here = import.meta.dir
const read = (f) => readFileSync(join(here, f), 'utf8')

// Each engine module becomes its own scope that returns what it exports; imports become destructuring from those scopes.
// The result is one script that exposes everything as `XL`.
function engine() {
  let out = ''
  const mods = []
  for (const f of ['format', 'core', 'functions']) {
    const names = []
    let src = read(`${f}.js`)
    src = src.replace(/^import \{([^}]*)\} from '\.\/(\w+)\.js'.*$/gm, (m, list, from) => `const {${list}} = M_${from}`)
    src = src.replace(/^export (const|let|function|class) (\w+)/gm, (m, kind, name) => {
      names.push(name)
      return `${kind} ${name}`
    })
    src = src.replace(/^export \{[^}]*\}.*$/gm, '')
    out += `\nconst M_${f} = (() => {${src}\nreturn { ${names.join(', ')} }\n})()\n`
    mods.push(`M_${f}`)
  }
  return `${out}\nconst XL = { ${mods.map((m) => `...${m}`).join(', ')} }`
}

let cached = null
export function buildPage() {
  if (cached && !process.env.SHEETS_DEV) return cached
  const html = read('ui/app.html')
    .replace('/*CSS*/', () => read('ui/app.css'))
    .replace('/*ENGINE*/', () => engine())
    .replace('/*APP*/', () => ['util', 'grid', 'sheet', 'dbui', 'main'].map((f) => read(`ui/${f}.js`)).join('\n'))
  cached = html
  return html
}
