export function decodeCol(s: string): number {
  let n = 0;
  for (const ch of s.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n - 1;
}

export function encodeCol(c: number): string {
  let s = '';
  for (let n = c + 1; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s;
  return s;
}

export const encodeAddr = (r: number, c: number) => encodeCol(c) + (r + 1);

export function decodeAddr(a: string): { r: number; c: number } | null {
  const m = /^([A-Za-z]{1,3})(\d+)$/.exec(a);
  return m ? { c: decodeCol(m[1]), r: Number(m[2]) - 1 } : null;
}
