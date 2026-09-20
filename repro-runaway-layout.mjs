// Repro: opening the attached DOCX never finishes layout; the Node heap
// grows without bound until V8 dies with "JavaScript heap out of memory".
//
//   npm i @silurus/ooxml@0.87.0 @napi-rs/canvas
//   node --max-old-space-size=1024 repro-runaway-layout.mjs
//
// Reproduces on @silurus/ooxml 0.86.1 and 0.87.0.
import { readFileSync } from 'node:fs'
import { openDocxDocument } from '@silurus/ooxml/node'
import { createCanvas } from '@napi-rs/canvas'

const bytes = readFileSync(new URL('./repro-runaway-layout.docx', import.meta.url))
const t0 = performance.now()
const timer = setInterval(() => {
  const m = process.memoryUsage()
  console.log(((performance.now() - t0) / 1000).toFixed(0) + 's',
    'rss', (m.rss / 2 ** 20).toFixed(0), 'MB')
}, 5000)

try {
  const doc = await openDocxDocument(bytes, {
    factory: { createCanvas: (w, h) => createCanvas(w, h) },
  })
  console.log('opened in', ((performance.now() - t0) / 1000).toFixed(2), 's, pages:', doc.pageCount)
} catch (err) {
  console.log('failed in', ((performance.now() - t0) / 1000).toFixed(2), 's:', err?.message)
}
clearInterval(timer)
