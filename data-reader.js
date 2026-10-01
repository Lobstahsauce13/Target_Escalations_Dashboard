// Reads the data file (Power BI's export) in the background, one row at a time, and keeps only the rows the page asks for.
// The export's one sheet unpacks to hundreds of MB of text, too much to hold at once: reading it whole froze the tab for
// half a minute and could run Chrome out of memory. Here the sheet is unzipped as a stream and each row is dropped
// unless it matches, so memory stays small and the page can show progress.
//
// In:  { file, sheet, keep: { column: regex source } }. A row is kept when every named column matches its pattern.
// Out: { progress: [rows read, rows kept, fraction of the file read] } now and then, then { header, rows } or { error }.
// Cell values come out as SheetJS's sheet_to_json(raw: true, defval: '') gives them: numbers as numbers, text as text, blanks as ''.

const td = new TextDecoder();
const XML_ENT = { lt: '<', gt: '>', amp: '&', quot: '"', apos: "'" };
const unxml = s =>
  s
    .replace(/&(lt|gt|amp|quot|apos);|&#x([0-9a-f]+);|&#(\d+);/gi, (m, n, h, d) =>
      n ? XML_ENT[n] : String.fromCodePoint(h ? parseInt(h, 16) : +d)
    )
    .replace(/_x([0-9a-f]{4})_/gi, (m, h) => String.fromCharCode(parseInt(h, 16)));
// The text of every <t> inside a string item (a rich text cell has several runs).
const textOf = xml => unxml((xml.match(/<(?:\w+:)?t\b[^>]*>[^<]*<\/(?:\w+:)?t>/g) ?? []).map(t => t.replace(/<[^>]+>/g, '')).join(''));

// The zip's file list, from its central directory.
function zipEntries(u8) {
  const dv = new DataView(u8.buffer, u8.byteOffset, u8.byteLength);
  let e = u8.length - 22;
  while (e >= 0 && dv.getUint32(e, true) !== 0x06054b50) e--;
  if (e < 0) throw Error("This isn't an Excel (.xlsx) file.");
  const out = new Map();
  for (let i = 0, p = dv.getUint32(e + 16, true), n = dv.getUint16(e + 10, true); i < n; i++) {
    const nl = dv.getUint16(p + 28, true),
      off = dv.getUint32(p + 42, true);
    out.set(td.decode(u8.subarray(p + 46, p + 46 + nl)), {
      method: dv.getUint16(p + 10, true),
      size: dv.getUint32(p + 20, true),
      unpacked: dv.getUint32(p + 24, true),
      // The data starts after the local header, whose name and extra field lengths can differ from the central directory's.
      start: off + 30 + dv.getUint16(off + 26, true) + dv.getUint16(off + 28, true)
    });
    p += 46 + nl + dv.getUint16(p + 30, true) + dv.getUint16(p + 32, true);
  }
  return out;
}

// One zip entry as a stream of text. onBytes is told what fraction of the unpacked entry has been read so far.
function entryText(u8, ent, onBytes) {
  let read = 0;
  const packed = new Blob([u8.subarray(ent.start, ent.start + ent.size)]).stream();
  if (ent.method !== 0 && ent.method !== 8) throw Error("This Excel file is packed in a way the dashboard can't read.");
  return (ent.method === 8 ? packed.pipeThrough(new DecompressionStream('deflate-raw')) : packed)
    .pipeThrough(
      new TransformStream({
        transform(chunk, c) {
          onBytes((read += chunk.length) / ent.unpacked);
          c.enqueue(chunk);
        }
      })
    )
    .pipeThrough(new TextDecoderStream());
}
async function wholeText(u8, ent) {
  let s = '';
  for await (const c of entryText(u8, ent, () => {})) s += c;
  return s;
}

// The path of the sheet with this name, from the workbook and its relationships.
async function sheetPath(u8, zip, name) {
  const wb = await wholeText(u8, zip.get('xl/workbook.xml')),
    rels = await wholeText(u8, zip.get('xl/_rels/workbook.xml.rels')),
    tag = [...wb.matchAll(/<(?:\w+:)?sheet\b[^>]*>/g)].map(m => m[0]).find(t => unxml(/\bname="([^"]*)"/.exec(t)?.[1] ?? '') === name);
  if (!tag) throw Error(`Missing “${name}” sheet.`);
  const id = /\br:id="([^"]*)"|\bid="([^"]*)"/.exec(tag.replace(/\bsheetId="[^"]*"/, '')),
    rid = id[1] ?? id[2],
    rel = [...rels.matchAll(/<(?:\w+:)?Relationship\b[^>]*>/g)].map(m => m[0]).find(t => t.includes(`Id="${rid}"`)),
    target = /\bTarget="([^"]*)"/.exec(rel)[1];
  return target.startsWith('/') ? target.slice(1) : 'xl/' + target;
}

// "AB12" -> 27 (zero-based column).
const colOf = ref => {
  let n = 0;
  for (const ch of /^[A-Z]+/.exec(ref)[0]) n = n * 26 + ch.charCodeAt(0) - 64;
  return n - 1;
};
const CELL = /<(?:\w+:)?c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/(?:\w+:)?c>)/g;
function cells(row, shared) {
  const out = [];
  let m;
  CELL.lastIndex = 0;
  while ((m = CELL.exec(row))) {
    const at = m[1],
      inner = m[2] ?? '',
      ref = /\br="([A-Z]+)\d*"/.exec(at),
      t = /\bt="(\w+)"/.exec(at)?.[1],
      v = /<(?:\w+:)?v>([^<]*)<\/(?:\w+:)?v>/.exec(inner)?.[1];
    if (ref) while (out.length < colOf(ref[1])) out.push('');
    out.push(
      t === 'inlineStr'
        ? textOf(inner)
        : v == null
          ? ''
          : t === 's'
            ? shared[+v]
            : t === 'str' || t === 'e'
              ? unxml(v)
              : t === 'b'
                ? v === '1'
                : +v
    );
  }
  return out;
}

self.onmessage = async ({ data: { file, sheet, keep } }) => {
  try {
    const u8 = new Uint8Array(await file.arrayBuffer()),
      zip = zipEntries(u8),
      ent = zip.get(await sheetPath(u8, zip, sheet));
    const sst = zip.get('xl/sharedStrings.xml'),
      shared = sst ? (await wholeText(u8, sst)).match(/<(?:\w+:)?si>[\s\S]*?<\/(?:\w+:)?si>/g)?.map(textOf) ?? [] : [];
    let header = null,
      tests = null,
      seen = 0,
      done = 0,
      lastPost = 0,
      rest = '';
    const kept = [];
    const row = xml => {
      const c = cells(xml, shared);
      if (!header) {
        header = c.map(h => String(h).trim());
        tests = Object.entries(keep).map(([k, rx]) => [header.indexOf(k), new RegExp(rx, 'i')]);
        return;
      }
      if (!c.some(x => x !== '')) return;
      seen++;
      if (tests.every(([i, rx]) => i >= 0 && rx.test(String(c[i] ?? '')))) {
        while (c.length < header.length) c.push('');
        kept.push(c);
      }
    };
    const END = /<\/(?:\w+:)?row>|<(?:\w+:)?row\b[^>]*\/>/g;
    for await (const chunk of entryText(u8, ent, f => (done = f))) {
      rest += chunk;
      let from = 0,
        m;
      END.lastIndex = 0;
      while ((m = END.exec(rest))) {
        const seg = rest.slice(from, END.lastIndex),
          i = seg.search(/<(?:\w+:)?row\b/);
        from = END.lastIndex;
        if (i >= 0 && !/\/>$/.test(m[0])) row(seg.slice(i));
      }
      rest = rest.slice(from);
      const t = Date.now();
      if (t - lastPost > 150) {
        lastPost = t;
        postMessage({ progress: [seen, kept.length, done] });
      }
    }
    if (!header) throw Error('The sheet is empty.');
    postMessage({ progress: [seen, kept.length, 1] });
    postMessage({ header, rows: kept });
  } catch (e) {
    postMessage({ error: e.message });
  }
};
