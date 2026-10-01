// Writes cell changes back into an Excel (.xlsx) file without touching anything else in it.
// An .xlsx is a zip of XML parts. Only the changed sheets' XML is rewritten; every other part (styles, shared strings,
// filters, conditional formatting, column widths, other sheets) is copied byte for byte, so the file keeps its formatting.
//
// patchXlsx(bytes, edits) -> Promise<Uint8Array>
//   edits: { "Sheet name": [{ ref: "K12", value, style }] }
//   value: a number (written as a number; dates are Excel serial day numbers) or text ('' clears the cell).
//   style: the style index for a cell that doesn't exist yet (a new row copies the row above's); an existing cell keeps its own.

const XLSX_TD = new TextDecoder(),
  XLSX_TE = new TextEncoder();

// The zip's entries, from its central directory, with what's needed to copy each one unchanged.
function zipDirectory(u8) {
  const dv = new DataView(u8.buffer, u8.byteOffset, u8.byteLength);
  let e = u8.length - 22;
  while (e >= 0 && dv.getUint32(e, true) !== 0x06054b50) e--;
  if (e < 0) throw Error("This isn't an Excel (.xlsx) file.");
  const out = [];
  for (let i = 0, p = dv.getUint32(e + 16, true), n = dv.getUint16(e + 10, true); i < n; i++) {
    const nl = dv.getUint16(p + 28, true),
      off = dv.getUint32(p + 42, true);
    out.push({
      name: XLSX_TD.decode(u8.subarray(p + 46, p + 46 + nl)),
      nameBytes: u8.slice(p + 46, p + 46 + nl),
      method: dv.getUint16(p + 10, true),
      time: dv.getUint16(p + 12, true),
      date: dv.getUint16(p + 14, true),
      crc: dv.getUint32(p + 16, true),
      size: dv.getUint32(p + 20, true),
      unpacked: dv.getUint32(p + 24, true),
      attrs: dv.getUint32(p + 38, true),
      // The data starts after the local header, whose name and extra field lengths can differ from the central directory's.
      start: off + 30 + dv.getUint16(off + 26, true) + dv.getUint16(off + 28, true)
    });
    p += 46 + nl + dv.getUint16(p + 30, true) + dv.getUint16(p + 32, true);
  }
  return out;
}

async function streamBytes(stream) {
  return new Uint8Array(await new Response(stream).arrayBuffer());
}
async function entryBytes(u8, ent) {
  const packed = u8.subarray(ent.start, ent.start + ent.size);
  if (ent.method === 0) return packed;
  if (ent.method !== 8) throw Error("This Excel file is packed in a way the dashboard can't write.");
  return streamBytes(new Blob([packed]).stream().pipeThrough(new DecompressionStream('deflate-raw')));
}
const deflate = bytes => streamBytes(new Blob([bytes]).stream().pipeThrough(new CompressionStream('deflate-raw')));

const CRC_TABLE = (() => {
  const t = new Uint32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    t[n] = c >>> 0;
  }
  return t;
})();
function crc32(u8) {
  let c = 0xffffffff;
  for (let i = 0; i < u8.length; i++) c = CRC_TABLE[(c ^ u8[i]) & 0xff] ^ (c >>> 8);
  return (c ^ 0xffffffff) >>> 0;
}

// A new zip of the same entries in the same order. replaced: entry name -> its new uncompressed bytes.
async function rebuildZip(u8, entries, replaced) {
  const parts = [],
    central = [];
  let offset = 0;
  for (const ent of entries) {
    let data = u8.subarray(ent.start, ent.start + ent.size),
      { method, crc, unpacked } = ent;
    if (replaced.has(ent.name)) {
      const raw = replaced.get(ent.name);
      data = await deflate(raw);
      method = 8;
      crc = crc32(raw);
      unpacked = raw.length;
    }
    const local = new DataView(new ArrayBuffer(30));
    local.setUint32(0, 0x04034b50, true);
    local.setUint16(4, 20, true);
    local.setUint16(6, 0x0800, true); // UTF-8 names, sizes in this header (no data descriptor)
    local.setUint16(8, method, true);
    local.setUint16(10, ent.time, true);
    local.setUint16(12, ent.date, true);
    local.setUint32(14, crc, true);
    local.setUint32(18, data.length, true);
    local.setUint32(22, unpacked, true);
    local.setUint16(26, ent.nameBytes.length, true);
    parts.push(new Uint8Array(local.buffer), ent.nameBytes, data);
    const cd = new DataView(new ArrayBuffer(46));
    cd.setUint32(0, 0x02014b50, true);
    cd.setUint16(4, 20, true);
    cd.setUint16(6, 20, true);
    cd.setUint16(8, 0x0800, true);
    cd.setUint16(10, method, true);
    cd.setUint16(12, ent.time, true);
    cd.setUint16(14, ent.date, true);
    cd.setUint32(16, crc, true);
    cd.setUint32(20, data.length, true);
    cd.setUint32(24, unpacked, true);
    cd.setUint16(28, ent.nameBytes.length, true);
    cd.setUint32(38, ent.attrs, true);
    cd.setUint32(42, offset, true);
    central.push(new Uint8Array(cd.buffer), ent.nameBytes);
    offset += 30 + ent.nameBytes.length + data.length;
  }
  const cdSize = central.reduce((n, p) => n + p.length, 0),
    end = new DataView(new ArrayBuffer(22));
  end.setUint32(0, 0x06054b50, true);
  end.setUint16(8, entries.length, true);
  end.setUint16(10, entries.length, true);
  end.setUint32(12, cdSize, true);
  end.setUint32(16, offset, true);
  return new Uint8Array(await new Blob([...parts, ...central, new Uint8Array(end.buffer)]).arrayBuffer());
}

const xmlText = s => String(s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[c]);
const xmlUnescape = s => s.replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&amp;/g, '&');
// "AB12" -> [27 (zero-based column), 12].
function splitRef(ref) {
  const m = /^([A-Z]+)(\d+)$/.exec(ref);
  let c = 0;
  for (const ch of m[1]) c = c * 26 + ch.charCodeAt(0) - 64;
  return [c - 1, +m[2]];
}

// The path of the sheet with this name, from the workbook and its relationships.
function sheetPathIn(files, name) {
  const wb = files.get('xl/workbook.xml'),
    rels = files.get('xl/_rels/workbook.xml.rels'),
    tag = [...wb.matchAll(/<(?:\w+:)?sheet\b[^>]*>/g)].map(m => m[0]).find(t => xmlUnescape(/\bname="([^"]*)"/.exec(t)?.[1] ?? '') === name);
  if (!tag) throw Error(`The file has no "${name}" sheet any more.`);
  const rid = /\br:id="([^"]*)"/.exec(tag)[1],
    rel = [...rels.matchAll(/<(?:\w+:)?Relationship\b[^>]*>/g)].map(m => m[0]).find(t => t.includes(`Id="${rid}"`)),
    target = /\bTarget="([^"]*)"/.exec(rel)[1];
  return target.startsWith('/') ? target.slice(1) : 'xl/' + target;
}

// The XML of one cell. A number goes in as is; text goes in inline, so the shared strings part stays untouched.
function cellXml(ref, style, value) {
  const s = style == null ? '' : ` s="${style}"`;
  if (typeof value === 'number') return `<c r="${ref}"${s}><v>${value}</v></c>`;
  if (value === '') return `<c r="${ref}"${s}/>`;
  return `<c r="${ref}"${s} t="inlineStr"><is><t xml:space="preserve">${xmlText(value)}</t></is></c>`;
}

// The sheet's XML with one cell set. An existing cell keeps its style; a missing cell or row is inserted in order.
function setCell(xml, ref, value, style) {
  const [col, rowN] = splitRef(ref),
    rowRe = new RegExp(`<row\\b[^>]*\\br="${rowN}"[^>]*?(/>|>)`),
    rm = rowRe.exec(xml);
  if (!rm) {
    // No such row: insert it before the first row numbered after it, or at the end of the sheet data.
    const after = [...xml.matchAll(/<row\b[^>]*\br="(\d+)"/g)].find(m => +m[1] > rowN),
      at = after ? after.index : xml.indexOf('</sheetData>');
    if (at < 0) throw Error('This sheet has no data area.');
    return xml.slice(0, at) + `<row r="${rowN}">${cellXml(ref, style, value)}</row>` + xml.slice(at);
  }
  if (rm[1] === '/>') {
    const open = rm[0].slice(0, -2) + '>';
    return xml.slice(0, rm.index) + open + cellXml(ref, style, value) + '</row>' + xml.slice(rm.index + rm[0].length);
  }
  const bodyStart = rm.index + rm[0].length,
    bodyEnd = xml.indexOf('</row>', bodyStart),
    body = xml.slice(bodyStart, bodyEnd),
    cells = [...body.matchAll(/<c\b[^>]*?(?:\/>|>[\s\S]*?<\/c>)/g)];
  for (const m of cells) {
    const r = /\br="([A-Z]+\d+)"/.exec(m[0])[1],
      [c] = splitRef(r);
    if (c === col) {
      const keep = /\bs="(\d+)"/.exec(m[0])?.[1] ?? style;
      const next = body.slice(0, m.index) + cellXml(ref, keep, value) + body.slice(m.index + m[0].length);
      return xml.slice(0, bodyStart) + next + xml.slice(bodyEnd);
    }
    if (c > col) {
      const next = body.slice(0, m.index) + cellXml(ref, style, value) + body.slice(m.index);
      return xml.slice(0, bodyStart) + next + xml.slice(bodyEnd);
    }
  }
  return xml.slice(0, bodyEnd) + cellXml(ref, style, value) + xml.slice(bodyEnd);
}

// The style index of a cell in the sheet, or null when the cell has none.
function cellStyle(xml, ref) {
  const m = new RegExp(`<c\\b[^>]*\\br="${ref}"[^>]*>`).exec(xml) ?? new RegExp(`<c\\b[^>]*\\br="${ref}"[^>]*/>`).exec(xml);
  return m ? (/\bs="(\d+)"/.exec(m[0])?.[1] ?? null) : null;
}

// The dimension ref grows to take in rows added below it.
function growDimension(xml, lastRow) {
  // Excel writes <dimension ref="A1:T20"/>; other tools put a space before the "/>".
  return xml.replace(/<dimension ref="([A-Z]+)(\d+):([A-Z]+)(\d+)"\s*\/>/, (m, c0, r0, c1, r1) =>
    +r1 >= lastRow ? m : `<dimension ref="${c0}${r0}:${c1}${lastRow}"/>`
  );
}

async function patchXlsx(bytes, edits) {
  const u8 = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes),
    entries = zipDirectory(u8),
    byName = new Map(entries.map(e => [e.name, e])),
    files = new Map();
  for (const n of ['xl/workbook.xml', 'xl/_rels/workbook.xml.rels']) files.set(n, XLSX_TD.decode(await entryBytes(u8, byName.get(n))));
  const replaced = new Map();
  for (const [sheet, list] of Object.entries(edits)) {
    if (!list.length) continue;
    const path = sheetPathIn(files, sheet);
    let xml = XLSX_TD.decode(await entryBytes(u8, byName.get(path))),
      last = 0;
    for (const e of list) {
      // A new cell copies its style from the cell above it, so an added row looks like the rows before it.
      const [, r] = splitRef(e.ref),
        above = e.ref.replace(/\d+$/, String(r - 1)),
        style = e.style !== undefined ? e.style : cellStyle(xml, above);
      xml = setCell(xml, e.ref, e.value, style);
      last = Math.max(last, r);
    }
    replaced.set(path, XLSX_TE.encode(growDimension(xml, last)));
  }
  return rebuildZip(u8, entries, replaced);
}
