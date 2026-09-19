/** A → 0, Z → 25, AA → 26 … */
export function lettersToCol(letters: string): number {
  let n = 0;
  for (const ch of letters.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n - 1;
}

export function colToLetters(col: number): string {
  let s = '';
  let n = col + 1;
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

export function cellAddress(row: number, col: number): string {
  return `${colToLetters(col)}${row + 1}`;
}

/** Кириллические буквы, визуально совпадающие с латинскими: «А1» → «A1». */
const LOOKALIKE: Record<string, string> = {
  А: 'A', В: 'B', С: 'C', Е: 'E', Н: 'H', К: 'K', М: 'M', О: 'O', Р: 'P', Т: 'T', Х: 'X',
  а: 'A', в: 'B', с: 'C', е: 'E', н: 'H', к: 'K', м: 'M', о: 'O', р: 'P', т: 'T', х: 'X',
};

export function normalizeLookalikes(s: string): string {
  let out = '';
  for (const ch of s) out += LOOKALIKE[ch] ?? ch;
  return out;
}

export function isLookalike(ch: string): boolean {
  return ch in LOOKALIKE;
}
