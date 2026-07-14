import fs from 'node:fs';
import path from 'node:path';
import zlib from 'node:zlib';

const [input, outDir] = process.argv.slice(2);
if (!input || !outDir) throw new Error('Usage: node extract_pdf.mjs input.pdf outDir');
fs.mkdirSync(outDir, { recursive: true });
const imageDir = path.join(outDir, 'images');
fs.mkdirSync(imageDir, { recursive: true });
const source = fs.readFileSync(input).toString('latin1');
const objects = new Map();

function streamParts(body) {
  const s = body.indexOf('stream');
  const e = body.lastIndexOf('endstream');
  if (s < 0 || e <= s) return null;
  let p = s + 6;
  if (body[p] === '\r' && body[p + 1] === '\n') p += 2;
  else if (body[p] === '\r' || body[p] === '\n') p += 1;
  let q = e;
  while (q > p && /[\r\n]/.test(body[q - 1])) q -= 1;
  return { dict: body.slice(0, s), raw: Buffer.from(body.slice(p, q), 'latin1') };
}

function decodeStream(body) {
  const parts = streamParts(body);
  if (!parts) return null;
  if (/\/FlateDecode\b/.test(parts.dict)) {
    try { return zlib.inflateSync(parts.raw); } catch { return null; }
  }
  return parts.raw;
}

const directRe = /(\d+)\s+(\d+)\s+obj\b([\s\S]*?)\bendobj\b/g;
let m;
while ((m = directRe.exec(source))) objects.set(Number(m[1]), { body: m[3], direct: true });

for (const [id, obj] of [...objects]) {
  if (!/\/Type\s*\/ObjStm\b/.test(obj.body)) continue;
  const decoded = decodeStream(obj.body);
  if (!decoded) continue;
  const text = decoded.toString('latin1');
  const n = Number((obj.body.match(/\/N\s+(\d+)/) || [])[1]);
  const first = Number((obj.body.match(/\/First\s+(\d+)/) || [])[1]);
  const nums = text.slice(0, first).trim().split(/\s+/).map(Number);
  for (let i = 0; i < n; i += 1) {
    const objId = nums[i * 2];
    const offset = nums[i * 2 + 1];
    const next = i + 1 < n ? nums[(i + 1) * 2 + 1] : text.length - first;
    objects.set(objId, { body: text.slice(first + offset, first + next).trim(), direct: false, objectStream: id });
  }
}

function utf16be(hex) {
  const clean = hex.replace(/\s+/g, '');
  let out = '';
  for (let i = 0; i + 3 < clean.length; i += 4) out += String.fromCharCode(parseInt(clean.slice(i, i + 4), 16));
  return out;
}

function parseCMap(text) {
  const map = new Map();
  let sourceBytes = 2;
  const code = text.match(/begincodespacerange[\s\S]*?<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>/);
  if (code) sourceBytes = code[1].length / 2;
  for (const line of text.split(/\r?\n/)) {
    let x = line.match(/^\s*<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>\s*$/);
    if (x) { map.set(x[1].toUpperCase(), utf16be(x[2])); continue; }
    x = line.match(/^\s*<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>\s*$/);
    if (x) {
      const start = parseInt(x[1], 16), end = parseInt(x[2], 16), dst = parseInt(x[3], 16);
      const width = x[1].length;
      for (let c = start; c <= end; c += 1) map.set(c.toString(16).padStart(width, '0').toUpperCase(), String.fromCharCode(dst + c - start));
      continue;
    }
    x = line.match(/^\s*<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>\s*\[(.*)\]\s*$/);
    if (x) {
      const start = parseInt(x[1], 16), width = x[1].length;
      const dests = [...x[3].matchAll(/<([0-9A-Fa-f]+)>/g)];
      for (let j = 0; j < dests.length; j += 1) map.set((start + j).toString(16).padStart(width, '0').toUpperCase(), utf16be(dests[j][1]));
    }
  }
  return { map, sourceBytes };
}

const cmaps = new Map();
for (const [id, obj] of objects) {
  const decoded = decodeStream(obj.body);
  if (decoded && /begincmap|beginbfchar|beginbfrange/.test(decoded.toString('latin1'))) cmaps.set(id, parseCMap(decoded.toString('latin1')));
}

const fontMaps = new Map();
for (const [id, obj] of objects) {
  if (!/\/Type\s*\/Font\b/.test(obj.body)) continue;
  const ref = obj.body.match(/\/ToUnicode\s+(\d+)\s+\d+\s+R/);
  fontMaps.set(id, ref && cmaps.get(Number(ref[1])) ? cmaps.get(Number(ref[1])) : null);
}

function decodeHex(hex, cmap) {
  const clean = hex.replace(/\s+/g, '').toUpperCase();
  if (cmap) {
    const width = cmap.sourceBytes * 2;
    let out = '';
    for (let i = 0; i < clean.length; i += width) out += cmap.map.get(clean.slice(i, i + width)) ?? '�';
    return out;
  }
  let out = '';
  for (let i = 0; i + 1 < clean.length; i += 2) {
    const b = parseInt(clean.slice(i, i + 2), 16);
    out += b >= 32 && b < 127 ? String.fromCharCode(b) : '';
  }
  return out;
}

function decodeLiteral(s) {
  return s.replace(/\\([nrtbf()\\])/g, (_, c) => ({n:'\n',r:'\r',t:'\t',b:'\b',f:'\f','(':'(',')':')','\\':'\\'})[c] ?? c)
    .replace(/\\([0-7]{1,3})/g, (_,o)=>String.fromCharCode(parseInt(o,8)));
}

function extractGlyphs(content, pageFontRefs) {
  const lines = content.toString('latin1').split(/\r?\n/);
  let state = { fontName: null, fontSize: 10, x: 0, y: 0 };
  const stack = [];
  const glyphs = [];
  for (const raw of lines) {
    const line = raw.trim();
    if (line === 'q') { stack.push({...state}); continue; }
    if (line === 'Q') { state = stack.pop() ?? state; continue; }
    let x = line.match(/^\/(\S+)\s+([-+\d.]+)\s+Tf$/);
    if (x) { state.fontName = x[1]; state.fontSize = Number(x[2]); continue; }
    x = line.match(/^([-+\d.]+)\s+([-+\d.]+)\s+([-+\d.]+)\s+([-+\d.]+)\s+([-+\d.]+)\s+([-+\d.]+)\s+cm$/);
    if (x) { state.x = Number(x[5]); state.y = Number(x[6]); continue; }
    const fontId = pageFontRefs.get(state.fontName);
    const cmap = fontMaps.get(fontId);
    const chunks = [];
    if (/TJ$/.test(line)) {
      for (const h of line.matchAll(/<([0-9A-Fa-f\s]+)>/g)) chunks.push(decodeHex(h[1], cmap));
      for (const l of line.matchAll(/\(([^()]*(?:\\.[^()]*)*)\)/g)) chunks.push(decodeLiteral(l[1]));
    } else {
      x = line.match(/^<([0-9A-Fa-f\s]+)>\s*Tj$/); if (x) chunks.push(decodeHex(x[1], cmap));
      x = line.match(/^\((.*)\)\s*Tj$/); if (x) chunks.push(decodeLiteral(x[1]));
    }
    const text = chunks.join('');
    if (text) glyphs.push({ x: state.x, y: state.y, size: state.fontSize, text, font: state.fontName });
  }
  return glyphs;
}

function glyphsToText(glyphs) {
  const sorted = [...glyphs].sort((a,b) => (b.y - a.y) || (a.x - b.x));
  const lines = [];
  for (const g of sorted) {
    let line = lines.find(l => Math.abs(l.y - g.y) <= 1.8);
    if (!line) { line = { y: g.y, items: [] }; lines.push(line); }
    line.items.push(g);
  }
  lines.sort((a,b)=>b.y-a.y);
  return lines.map(line => {
    line.items.sort((a,b)=>a.x-b.x);
    let out = '', prev = null;
    for (const item of line.items) {
      if (prev) {
        const gap = item.x - prev.x;
        const bothAscii = /^[\x20-\x7E]+$/.test(prev.text + item.text);
        if (bothAscii && gap > Math.max(4.5, prev.size * 0.75)) out += ' ';
      }
      out += item.text;
      prev = item;
    }
    return out.replace(/�+/g,'').trim();
  }).filter(Boolean).join('\n');
}

const pagesRoot = [...objects].find(([,o])=>/\/Type\s*\/Pages\b/.test(o.body));
let pageIds = [];
if (pagesRoot) pageIds = [...pagesRoot[1].body.matchAll(/(\d+)\s+\d+\s+R/g)].map(x=>Number(x[1])).filter(id=>/\/Type\s*\/Page\b/.test(objects.get(id)?.body ?? ''));
if (!pageIds.length) pageIds = [...objects].filter(([,o])=>/\/Type\s*\/Page\b/.test(o.body)).map(([id])=>id).sort((a,b)=>a-b);

const pageTexts = [];
for (let i = 0; i < pageIds.length; i += 1) {
  const body = objects.get(pageIds[i]).body;
  const contentRef = Number((body.match(/\/Contents\s+(\d+)\s+\d+\s+R/)||[])[1]);
  const fontRefs = new Map([...body.matchAll(/\/(F\S+)\s+(\d+)\s+\d+\s+R/g)].map(x=>[x[1],Number(x[2])]));
  const content = decodeStream(objects.get(contentRef)?.body ?? '');
  const text = content ? glyphsToText(extractGlyphs(content, fontRefs)) : '';
  pageTexts.push(`===== PAGE ${i+1} =====\n${text}`);
}
fs.writeFileSync(path.join(outDir,'fulltext.txt'), pageTexts.join('\n\n'), 'utf8');

const imageInventory = [];
for (const [id,obj] of objects) {
  const parts = streamParts(obj.body);
  if (!parts || !/\/DCTDecode\b/.test(parts.dict) || !/\/Subtype\s*\/Image\b/.test(parts.dict)) continue;
  const width = Number((parts.dict.match(/\/Width\s+(\d+)/)||[])[1]);
  const height = Number((parts.dict.match(/\/Height\s+(\d+)/)||[])[1]);
  const file = `image_object_${id}.jpg`;
  fs.writeFileSync(path.join(imageDir,file), parts.raw);
  imageInventory.push({id,file,width,height,bytes:parts.raw.length});
}
fs.writeFileSync(path.join(outDir,'images.json'),JSON.stringify(imageInventory,null,2));
console.log(JSON.stringify({objects:objects.size,pages:pageIds.length,cmaps:cmaps.size,fonts:fontMaps.size,images:imageInventory.length,textChars:pageTexts.join('').length},null,2));
