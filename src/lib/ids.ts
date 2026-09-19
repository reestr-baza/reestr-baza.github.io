const ALPHABET = '0123456789abcdefghijklmnopqrstuvwxyz';

/** Короткий случайный идентификатор: 10 символов base36 ≈ 51 бит энтропии. */
export function uid(len = 10): string {
  const bytes = crypto.getRandomValues(new Uint8Array(len));
  let out = '';
  for (const b of bytes) out += ALPHABET[b % 36];
  return out;
}
