import { describe, expect, test } from 'vitest';
import { LIMITS, unzip, zip } from '../lib/zip';

const enc = new TextEncoder();
const files = (o: Record<string, string>) => new Map(Object.entries(o).map(([k, v]) => [k, enc.encode(v)]));
const toBuf = (u: Uint8Array) => u.buffer.slice(u.byteOffset, u.byteOffset + u.byteLength) as ArrayBuffer;

describe('zip', () => {
  test('round-trips compressible and incompressible entries', async () => {
    const input = files({ 'a.txt': 'hello '.repeat(1000), 'dir/b.bin': 'x', 'empty.txt': '' });
    const out = await unzip(toBuf(await zip(input)));
    expect([...out.keys()].sort()).toEqual(['a.txt', 'dir/b.bin', 'empty.txt']);
    expect(new TextDecoder().decode(out.get('a.txt'))).toBe('hello '.repeat(1000));
    expect(out.get('empty.txt')!.length).toBe(0);
  });

  test('accepts a Uint8Array view with a non-zero offset', async () => {
    const z = await zip(files({ 'a.txt': 'abc' }));
    const padded = new Uint8Array(z.length + 7);
    padded.set(z, 7);
    const out = await unzip(padded.subarray(7));
    expect(new TextDecoder().decode(out.get('a.txt'))).toBe('abc');
  });

  test('rejects non-zip input', async () => {
    await expect(unzip(new Uint8Array([1, 2, 3, 4]).buffer)).rejects.toThrow('Not a ZIP');
  });

  test('detects corrupted data through the checksum', async () => {
    const bytes = await zip(files({ 'a.txt': 'some content that is stored as is' }));
    // flip a byte inside the (stored) payload, after the 30-byte header and the name
    bytes[30 + 'a.txt'.length + 3] ^= 0xff;
    await expect(unzip(toBuf(bytes))).rejects.toThrow();
  });

  test('rejects an archive whose declared size exceeds the limit', async () => {
    const bytes = await zip(files({ 'a.txt': 'x' }));
    const dv = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    const eocd = bytes.length - 22;
    const cd = dv.getUint32(eocd + 16, true);
    dv.setUint32(cd + 24, LIMITS.maxEntryBytes + 1, true);
    await expect(unzip(toBuf(bytes))).rejects.toThrow('too large');
  });
});
