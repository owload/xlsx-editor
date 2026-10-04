/**
 * Minimal ZIP reader/writer on top of the built-in DecompressionStream/CompressionStream.
 * No third-party dependencies; with limits against zip bombs.
 */

export const LIMITS = {
  maxEntries: 2000,
  maxEntryBytes: 200 * 1024 * 1024,
  maxTotalBytes: 500 * 1024 * 1024,
};

const u16 = (d: DataView, o: number) => d.getUint16(o, true);
const u32 = (d: DataView, o: number) => d.getUint32(o, true);

const CRC_TABLE = (() => {
  const t = new Uint32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    t[n] = c >>> 0;
  }
  return t;
})();

export function crc32(buf: Uint8Array): number {
  let c = 0xffffffff;
  for (let i = 0; i < buf.length; i++) c = CRC_TABLE[(c ^ buf[i]) & 0xff] ^ (c >>> 8);
  return (c ^ 0xffffffff) >>> 0;
}

async function pipe(data: Uint8Array, stream: { writable: WritableStream; readable: ReadableStream<Uint8Array> }, cap: number) {
  const writer = stream.writable.getWriter();
  writer.write(data).catch(() => {});
  writer.close().catch(() => {});
  const reader = stream.readable.getReader();
  const chunks: Uint8Array[] = [];
  let total = 0;
  for (;;) {
    const { done, value } = await reader.read();
    if (done) break;
    total += value.length;
    if (total > cap) {
      await reader.cancel();
      throw new Error('File is too large after decompression');
    }
    chunks.push(value);
  }
  const out = new Uint8Array(total);
  let off = 0;
  for (const c of chunks) {
    out.set(c, off);
    off += c.length;
  }
  return out;
}

export async function unzip(buf: ArrayBuffer | Uint8Array): Promise<Map<string, Uint8Array>> {
  const bytes = buf instanceof Uint8Array ? buf : new Uint8Array(buf);
  const dv = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  let eocd = -1;
  for (let i = bytes.length - 22; i >= Math.max(0, bytes.length - 22 - 65535); i--) {
    if (u32(dv, i) === 0x06054b50) {
      eocd = i;
      break;
    }
  }
  if (eocd < 0) throw new Error('Not a ZIP/XLSX file');
  const count = u16(dv, eocd + 10);
  let p = u32(dv, eocd + 16);
  if (count > LIMITS.maxEntries) throw new Error('Too many files in the archive');
  const files = new Map<string, Uint8Array>();
  let total = 0;
  const dec = new TextDecoder();
  for (let n = 0; n < count; n++) {
    if (p + 46 > bytes.length || u32(dv, p) !== 0x02014b50) throw new Error('Corrupted archive');
    const flags = u16(dv, p + 8);
    const method = u16(dv, p + 10);
    const crc = u32(dv, p + 16);
    const csize = u32(dv, p + 20);
    const usize = u32(dv, p + 24);
    const nlen = u16(dv, p + 28), elen = u16(dv, p + 30), clen = u16(dv, p + 32);
    const lho = u32(dv, p + 42);
    const name = dec.decode(bytes.subarray(p + 46, p + 46 + nlen));
    p += 46 + nlen + elen + clen;
    if (flags & 1) throw new Error('Encrypted xlsx files are not supported');
    if (name.endsWith('/')) continue;
    if (usize > LIMITS.maxEntryBytes || (total += usize) > LIMITS.maxTotalBytes) throw new Error('Archive is too large');
    if (lho + 30 > bytes.length || u32(dv, lho) !== 0x04034b50) throw new Error('Corrupted archive');
    const start = lho + 30 + u16(dv, lho + 26) + u16(dv, lho + 28);
    if (start + csize > bytes.length) throw new Error('Corrupted archive');
    const raw = bytes.subarray(start, start + csize);
    let data: Uint8Array;
    if (method === 0) data = raw.slice();
    else if (method === 8) data = await pipe(raw, new DecompressionStream('deflate-raw'), Math.min(usize, LIMITS.maxEntryBytes));
    else throw new Error('Unsupported compression method');
    if (data.length !== usize || crc32(data) !== crc) throw new Error('Checksum mismatch');
    // the name is only a Map key; it never reaches a file system
    files.set(name.replace(/^\/+/, ''), data);
  }
  return files;
}

export async function zip(files: Map<string, Uint8Array>): Promise<Uint8Array> {
  const enc = new TextEncoder();
  const parts: Uint8Array[] = [];
  const central: Uint8Array[] = [];
  let offset = 0;
  for (const [name, data] of files) {
    const nameB = enc.encode(name);
    const comp = await pipe(data, new CompressionStream('deflate-raw'), Number.MAX_SAFE_INTEGER);
    const useDeflate = comp.length < data.length;
    const body = useDeflate ? comp : data;
    const crc = crc32(data);
    const lh = new DataView(new ArrayBuffer(30));
    lh.setUint32(0, 0x04034b50, true);
    lh.setUint16(4, 20, true);
    lh.setUint16(6, 0x0800, true);
    lh.setUint16(8, useDeflate ? 8 : 0, true);
    lh.setUint32(14, crc, true);
    lh.setUint32(18, body.length, true);
    lh.setUint32(22, data.length, true);
    lh.setUint16(26, nameB.length, true);
    const ch = new DataView(new ArrayBuffer(46));
    ch.setUint32(0, 0x02014b50, true);
    ch.setUint16(4, 20, true);
    ch.setUint16(6, 20, true);
    ch.setUint16(8, 0x0800, true);
    ch.setUint16(10, useDeflate ? 8 : 0, true);
    ch.setUint32(16, crc, true);
    ch.setUint32(20, body.length, true);
    ch.setUint32(24, data.length, true);
    ch.setUint16(28, nameB.length, true);
    ch.setUint32(42, offset, true);
    parts.push(new Uint8Array(lh.buffer), nameB, body);
    central.push(new Uint8Array(ch.buffer), nameB);
    offset += 30 + nameB.length + body.length;
  }
  const cdSize = central.reduce((a, b) => a + b.length, 0);
  const end = new DataView(new ArrayBuffer(22));
  end.setUint32(0, 0x06054b50, true);
  end.setUint16(8, files.size, true);
  end.setUint16(10, files.size, true);
  end.setUint32(12, cdSize, true);
  end.setUint32(16, offset, true);
  const all = [...parts, ...central, new Uint8Array(end.buffer)];
  const out = new Uint8Array(all.reduce((a, b) => a + b.length, 0));
  let o = 0;
  for (const a of all) {
    out.set(a, o);
    o += a.length;
  }
  return out;
}
