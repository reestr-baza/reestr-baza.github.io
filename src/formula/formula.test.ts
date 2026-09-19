import { describe, expect, it } from 'vitest';
import { evaluateToScalar, type EvalHost } from './evaluate';
import { parseFormula } from './parse';
import { adjustFormula, remapColumns, shiftFormula } from './refs';
import { ERROR_TEXT, isErr, type Scalar } from './values';

function hostOf(grid: Scalar[][], names: Record<string, Scalar> = {}): EvalHost {
  return {
    sheetByName: (n) => (n.toLowerCase() === 'лист1' ? 's1' : null),
    cellValue: (_s, r, c) => grid[r]?.[c] ?? null,
    rowCount: () => grid.length,
    colCount: () => Math.max(...grid.map((r) => r.length)),
    namedValue: (n) => names[n.toUpperCase()],
    rowHidden: (_s, r) => ({ manual: false, filtered: r === 2 }),
  };
}

function calc(src: string, grid: Scalar[][] = [], names?: Record<string, Scalar>, baseRow = 0): Scalar {
  const v = evaluateToScalar(parseFormula(src), { host: hostOf(grid, names), sheetId: 's1', row: 0, col: 0, baseRow });
  return isErr(v) ? ERROR_TEXT[v.code] : v;
}

const G: Scalar[][] = [
  ['Артикул', 'Товар', 'Цена', 'Курс'],
  [3707, 'Джинсы', 2500, 12.75],
  [3708, 'Худи', 1800, 12.75],
  [3709, 'Кепка', 300, 12.75],
];

describe('арифметика и операторы', () => {
  it('приоритеты', () => {
    expect(calc('=1+2*3')).toBe(7);
    expect(calc('=(1+2)*3')).toBe(9);
    expect(calc('=-2^2')).toBe(4);
    expect(calc('=2^3^2')).toBe(64);
    expect(calc('=10%')).toBe(0.1);
    expect(calc('=0.1+0.2')).toBe(0.3);
  });
  it('русская запись чисел и разделителей', () => {
    expect(calc('=12,5*2')).toBe(25);
    expect(calc('=ОКРУГЛ(2,345;2)')).toBe(2.35);
    expect(calc('=ROUND(2.345,2)')).toBe(2.35);
  });
  it('текст и сравнения', () => {
    expect(calc('="a"&1&ИСТИНА')).toBe('a1ИСТИНА');
    expect(calc('="abc"="ABC"')).toBe(true);
    expect(calc('=1/0')).toBe('#ДЕЛ/0!');
    expect(calc('="x"+1')).toBe('#ЗНАЧ!');
  });
});

describe('ссылки', () => {
  it('ячейки и диапазоны', () => {
    expect(calc('=C2*D2', G)).toBe(31875);
    expect(calc('=СУММ(C2:C4)', G)).toBe(4600);
    expect(calc('=SUM(C:C)', G)).toBe(4600);
    expect(calc('=$C$2', G)).toBe(2500);
    expect(calc('=Лист1!C3', G)).toBe(1800);
  });
  it('кириллические буквы в адресах', () => {
    expect(calc('=С2', G)).toBe(2500);
  });
  it('формула столбца со сдвигом строки', () => {
    expect(calc('=C1*D1', G, {}, 2)).toBe(1800 * 12.75);
  });
  it('именованные значения', () => {
    expect(calc('=C2*КУРС_ЮАНЯ', G, { КУРС_ЮАНЯ: 13 })).toBe(32500);
    expect(calc('=НЕТ_ТАКОГО', G)).toBe('#ИМЯ?');
  });
});

describe('функции', () => {
  it('логика', () => {
    expect(calc('=ЕСЛИ(C2>2000;"дорого";"ок")', G)).toBe('дорого');
    expect(calc('=IF(C4>2000,"дорого","ок")', G)).toBe('ок');
    expect(calc('=ЕСЛИОШИБКА(1/0;"—")')).toBe('—');
    expect(calc('=И(1;ИСТИНА)')).toBe(true);
  });
  it('условные агрегаты', () => {
    expect(calc('=СУММЕСЛИ(C2:C4;">1000")', G)).toBe(4300);
    expect(calc('=СЧЁТЕСЛИ(B2:B4;"Д*")', G)).toBe(1);
    expect(calc('=SUMIFS(C2:C4,B2:B4,"<>Кепка",C2:C4,">0")', G)).toBe(4300);
  });
  it('поиск', () => {
    expect(calc('=ВПР(3708;A2:D4;3;0)', G)).toBe(1800);
    expect(calc('=ВПР(9999;A2:D4;3;ЛОЖЬ)', G)).toBe('#Н/Д');
    expect(calc('=ИНДЕКС(B2:B4;ПОИСКПОЗ("Кепка";B2:B4;0))', G)).toBe('Кепка');
    expect(calc('=XLOOKUP("Худи",B2:B4,C2:C4)', G)).toBe(1800);
  });
  it('текст', () => {
    expect(calc('=ЛЕВСИМВ("Джинсы";3)')).toBe('Джи');
    expect(calc('=ПРОПИСН("prime")')).toBe('PRIME');
    expect(calc('=ТЕКСТ(31875;"# ##0")')).toBe('31 875');
    expect(calc('=СЦЕП("A";"-";1)')).toBe('A-1');
  });
  it('даты', () => {
    expect(calc('=ГОД(ДАТА(2026;9;18))')).toBe(2026);
    expect(calc('=ТЕКСТ(ДАТА(2026;9;18);"ДД.ММ.ГГГГ")')).toBe('18.09.2026');
  });
  it('итоги по видимым строкам', () => {
    expect(calc('=ПРОМЕЖУТОЧНЫЕ.ИТОГИ(9;C2:C4)', G)).toBe(2800);
  });
  it('неизвестная функция', () => {
    expect(calc('=ФУНКЦИЯ(1)')).toBe('#ИМЯ?');
  });
});

describe('перезапись ссылок', () => {
  it('сдвиг при копировании', () => {
    expect(shiftFormula('=G2*H2', 1, 0)).toBe('=G3*H3');
    expect(shiftFormula('=G2*$H$2', 3, 1)).toBe('=H5*$H$2');
    expect(shiftFormula('=СУММ(A1:A3;B$1)', 2, 0)).toBe('=СУММ(A3:A5;B$1)');
    expect(shiftFormula('=A1', -1, 0)).toBe('=#ССЫЛКА!');
  });
  it('вставка и удаление строк', () => {
    const all = () => true;
    expect(adjustFormula('=A5+A2', { axis: 'row', at: 2, count: 1 }, all)).toBe('=A6+A2');
    expect(adjustFormula('=СУММ(A2:A10)', { axis: 'row', at: 4, count: 2 }, all)).toBe('=СУММ(A2:A12)');
    expect(adjustFormula('=A5', { axis: 'row', at: 4, count: -1 }, all)).toBe('=#ССЫЛКА!');
    expect(adjustFormula('=СУММ(A2:A10)', { axis: 'row', at: 4, count: -3 }, all)).toBe('=СУММ(A2:A7)');
    expect(adjustFormula('=G1*$H$1', { axis: 'row', at: 0, count: 1 }, all, true)).toBe('=G1*$H$2');
  });
  it('столбцы', () => {
    const all = () => true;
    expect(adjustFormula('=C1*D1', { axis: 'col', at: 3, count: 1 }, all)).toBe('=C1*E1');
    expect(adjustFormula('=C1*D1', { axis: 'col', at: 2, count: -1 }, all)).toBe('=#ССЫЛКА!*C1');
    expect(remapColumns('=A1+C1', (c) => (c === 0 ? 2 : c === 2 ? 0 : c), all)).toBe('=C1+A1');
  });
  it('ссылки на другие листы не трогаются', () => {
    const onlyLocal = (s: string | undefined) => !s;
    expect(adjustFormula('=A5+Лист2!A5', { axis: 'row', at: 0, count: 1 }, onlyLocal)).toBe('=A6+Лист2!A5');
  });
});

describe('обмен с Excel', async () => {
  const { toRuFormula, fromExcelFormula, toExcelFormula } = await import('./excel');
  it('английская запись → русская', () => {
    expect(toRuFormula('=IF(N3>0,(J3/(N3*$J$1+$K$1)),"?")')).toBe('=ЕСЛИ(N3>0;(J3/(N3*$J$1+$K$1));"?")');
    expect(toRuFormula('=ROUND(A1*1.5,2)')).toBe('=ОКРУГЛ(A1*1,5;2)');
  });
  it('из файла со сдвигом строк заголовка', () => {
    expect(fromExcelFormula('J3-N3*$J$1-$K$1', -2)).toBe('=J1-N1*#ССЫЛКА!-#ССЫЛКА!');
  });
  it('обратно в Excel', () => {
    expect(toExcelFormula('=ЕСЛИ(N1>0;J1/(N1*КУРС+ДОСТАВКА);"?")', 2)).toBe('IF(N3>0,J3/(N3*КУРС+ДОСТАВКА),"?")');
    expect(toExcelFormula('=СУММ(A1:A3)*1,5', 1)).toBe('SUM(A2:A4)*1.5');
  });
});
