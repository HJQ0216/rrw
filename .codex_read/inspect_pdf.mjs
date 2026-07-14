import fs from 'node:fs';
import path from 'node:path';
import zlib from 'node:zlib';

const [input, outDir] = process.argv.slice(2);
if (!input || !outDir) throw new Error('Usage: node inspect_pdf.mjs input.pdf outDir');
fs.mkdirSync(outDir, { recursive: true });
const buf = fs.readFileSync(input);
const src = buf.toString('latin1');
const objects = new Map();
const objectRe = /(\d+)\s+(\d+)\s+obj\b([\s\S]*?)\bendobj\b/g;
let match;
while ((match = objectRe.exec(src))) {
  objects.set(Number(match[1]), { gen: Number(match[2]), body: match[3], start: match.index });
}

let pageCount = 0;
let streamCount = 0;
let flateCount = 0;
let textLikeCount = 0;
let cmapCount = 0;
const inventory = [];
for (const [id, obj] of objects) {
  if (/\/Type\s*\/Page\b/.test(obj.body)) pageCount += 1;
  const s = obj.body.indexOf('stream');
  const e = obj.body.lastIndexOf('endstream');
  let decoded = null;
  if (s >= 0 && e > s) {
    streamCount += 1;
    let p = s + 6;
    if (obj.body[p] === '\r' && obj.body[p + 1] === '\n') p += 2;
    else if (obj.body[p] === '\n' || obj.body[p] === '\r') p += 1;
    let q = e;
    while (q > p && (obj.body[q - 1] === '\r' || obj.body[q - 1] === '\n')) q -= 1;
    const raw = Buffer.from(obj.body.slice(p, q), 'latin1');
    if (/\/FlateDecode\b/.test(obj.body.slice(0, s))) {
      try { decoded = zlib.inflateSync(raw); flateCount += 1; } catch {}
    } else decoded = raw;
    if (decoded) {
      const text = decoded.toString('latin1');
      const isCmap = /begincmap|beginbfchar|beginbfrange/.test(text);
      const isTextLike = /\bBT\b|\bTj\b|\bTJ\b/.test(text);
      if (isCmap) cmapCount += 1;
      if (isTextLike) textLikeCount += 1;
      if (isCmap || isTextLike) fs.writeFileSync(path.join(outDir, `object_${id}.txt`), decoded);
      if (/\/Type\s*\/ObjStm\b/.test(obj.body.slice(0, s))) fs.writeFileSync(path.join(outDir, `object_${id}_objstm.txt`), decoded);
      inventory.push({ id, bytes: decoded.length, flate: /\/FlateDecode\b/.test(obj.body.slice(0,s)), isCmap, isTextLike, head: obj.body.slice(0, s).replace(/\s+/g,' ').trim().slice(0,240) });
    }
  }
}
fs.writeFileSync(path.join(outDir, 'inventory.json'), JSON.stringify(inventory, null, 2));
console.log(JSON.stringify({ bytes: buf.length, objects: objects.size, pageCount, streamCount, flateCount, textLikeCount, cmapCount }, null, 2));
