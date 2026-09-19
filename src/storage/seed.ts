import { uid } from '../lib/ids';
import { dateToSerial } from '../formula/functions';
import type { Cell, Column, Row, RowId, Sheet, WorkbookMeta } from '../model/types';
import { importImage } from './images';

const WASHES = ['indigo', 'black', 'lightwash', 'grey', 'ecru', 'olive', 'sand', 'darkblue'] as const;

const MODELS = ['PRIME', 'NORD', 'RIDER', 'LOFT', 'ONYX', 'MOSS', 'DUNE', 'HARBOR', 'KITE', 'SOHO'];
const FITS = ['2 широкие', 'прямые', 'клёш', 'мом', 'бананы', 'палаццо', 'бойфренд', 'слим'];
const WASH_NAMES: Record<(typeof WASHES)[number], string> = {
  indigo: 'индиго',
  black: 'чёрный',
  lightwash: 'светлая варка',
  grey: 'графит',
  ecru: 'экрю',
  olive: 'олива',
  sand: 'песок',
  darkblue: 'тёмный деним',
};
const SUPPLIERS = [
  ['Guangzhou Baiyun', 'https://detail.1688.com/offer/'],
  ['Shishi Denim Co.', 'https://detail.1688.com/offer/'],
  ['Xintang Jeans', 'https://www.alibaba.com/product-detail/'],
] as const;
const STATUSES = ['В наличии', 'В пути', 'Заказать', 'Нет у поставщика'];
const STATUS_COLORS: Record<string, string | undefined> = {
  'В наличии': undefined,
  'В пути': '#fff6c7',
  Заказать: '#dce9fb',
  'Нет у поставщика': '#fde2e1',
};

/** Демо-база: пример из ТЗ (3707 · Джинсы · PRIME · 2 широкие · ¥2 500 × 12,75) и ещё 47 позиций. */
export async function buildDemo(baseUrl: string): Promise<{ meta: WorkbookMeta; sheets: Sheet[] }> {
  const images: Record<string, string> = {};
  await Promise.all(
    WASHES.map(async (w) => {
      try {
        const res = await fetch(`${baseUrl}demo/jeans_${w}.webp`);
        if (res.ok) images[w] = await importImage(await res.blob(), `jeans_${w}.webp`);
      } catch {
        /* демо-фото не обязательны */
      }
    }),
  );

  const col = (name: string, w: number, extra: Partial<Column> = {}): Column => ({ id: uid(6), name, w, ...extra });
  const cSku = col('Артикул', 108, { st: { b: true } });
  const cPhoto = col('Фото', 92, { st: { ha: 'center' } });
  const cBrand = col('Марка', 110);
  const cModel = col('Модель', 104);
  const cGen = col('Поколение', 124);
  const cCountry = col('Страна', 96);
  const cYuan = col('Стоимость юани', 150, { st: { nf: { k: 'currency', c: 'CNY', d: 0 } } });
  const cRate = col('Курс', 80, { st: { nf: { k: 'number', d: 2 } } });
  const cRub = col('Стоимость рубли', 158, { st: { nf: { k: 'currency', c: 'RUB', d: 0 }, b: true } });
  const cColor = col('Цвет', 124);
  const cSizes = col('Размеры', 112);
  const cStock = col('Остаток', 96, { st: { nf: { k: 'number', d: 0 } } });
  const cStatus = col('Статус', 164);
  const cSupplier = col('Поставщик', 176);
  const cDate = col('Дата заказа', 118, { st: { nf: { k: 'date' } } });
  const cNote = col('Комментарий', 220, { st: { wrap: false } });
  const columns = [cSku, cPhoto, cBrand, cModel, cGen, cCountry, cYuan, cRate, cRub, cColor, cSizes, cStock, cStatus, cSupplier, cDate, cNote];
  // «Стоимость рубли» — формула столбца, как в ТЗ: юани × курс
  cRub.formula = `=G1*H1`;

  const rows = new Map<RowId, Row>();
  const order: RowId[] = [];
  const today = new Date(2026, 8, 18);
  let seed = 7;
  const rnd = () => {
    seed = (seed * 1103515245 + 12345) & 0x7fffffff;
    return seed / 0x7fffffff;
  };
  const pick = <T,>(xs: readonly T[]) => xs[Math.floor(rnd() * xs.length)];

  for (let i = 0; i < 48; i++) {
    const id = uid(8);
    const wash = i === 0 ? 'indigo' : WASHES[(i * 3 + Math.floor(i / 8)) % WASHES.length];
    const model = i === 0 ? 'PRIME' : MODELS[(i * 7) % MODELS.length];
    const fit = i === 0 ? '2 широкие' : FITS[(i * 5 + 3) % FITS.length];
    const yuan = i === 0 ? 2500 : Math.round((1400 + rnd() * 2600) / 10) * 10;
    const status = i === 0 ? 'В наличии' : pick(STATUSES);
    const supplier = SUPPLIERS[i % SUPPLIERS.length];
    const cells: Record<string, Cell> = {
      [cSku.id]: { v: 3707 + i },
      [cBrand.id]: { v: 'Джинсы' },
      [cModel.id]: { v: model },
      [cGen.id]: { v: fit },
      [cCountry.id]: { v: 'Китай' },
      [cYuan.id]: { v: yuan },
      [cRate.id]: { v: 12.75 },
      [cColor.id]: { v: WASH_NAMES[wash] },
      [cSizes.id]: { v: pick(['25–31', '26–32', '24–30', '27–34', 'XS–L', 'S–XL']) },
      [cStock.id]: { v: status === 'Нет у поставщика' ? 0 : Math.floor(rnd() * 40) },
      [cStatus.id]: { v: status, st: STATUS_COLORS[status] ? { bg: STATUS_COLORS[status] } : undefined },
      [cSupplier.id]: { v: supplier[0], href: `${supplier[1]}${6780000000 + i * 7919}.html` },
      [cDate.id]: { v: dateToSerial(new Date(today.getTime() - Math.floor(rnd() * 120) * 86400000)) },
    };
    if (images[wash] && i % 6 !== 5) cells[cPhoto.id] = { img: images[wash] };
    if (i === 0) cells[cNote.id] = { v: 'Пример из ТЗ: ¥2 500 × 12,75 = 31 875 ₽' };
    if (i === 3) cells[cNote.id] = { v: 'Проверить посадку на размере 27' };
    if (i === 9) cells[cNote.id] = { v: 'Поставщик обещал скидку от 50 шт.', st: { fg: '#1a5fd1' } };
    if (i % 6 === 5) cells[cNote.id] = { v: 'Нет фото — перетащите файл в ячейку' , st: { fg: '#737982', i: true } };
    const row: Row = { id, cells };
    if (i === 0 && images.indigo) {
      row.card = [
        { img: images.lightwash, text: 'Та же модель — светлая варка', st: { ha: 'center' } },
        { text: 'Ткань: 100% хлопок, 13 oz.\nПосадка высокая, штанина от бедра.', st: { va: 'top' } },
        { text: 'Размерная сетка', st: { va: 'top', b: true, bg: '#fff6c7' } },
        { text: '25 · 26 · 27 · 28 · 29 · 30 · 31', st: { va: 'middle', ha: 'center' } },
        { text: 'Мин. партия — 20 шт.\nСрок пошива 12 дней.', st: { va: 'top', b: true } },
        {},
        { text: 'Упаковка: пакет + бирка', st: { va: 'bottom', ha: 'right', i: true } },
      ];

    }
    rows.set(id, row);
    order.push(id);
  }
  // запас пустых строк для ввода
  for (let i = 0; i < 12; i++) {
    const id = uid(8);
    rows.set(id, { id, cells: {} });
    order.push(id);
  }

  const sheet: Sheet = {
    id: uid(8),
    name: 'Пример из ТЗ',
    columns,
    rowOrder: order,
    rows,
    keyColId: cSku.id,
    frozen: 2,
    density: 'M',
    filters: {},
  };

  const parts = buildPartsSheet();
  const meta: WorkbookMeta = {
    id: uid(8),
    title: 'База запчастей — пример',
    sheetIds: [parts.id, sheet.id],
    activeSheet: parts.id,
    names: [
      { name: 'КУРС', value: 12.1, note: 'Рублей за 1 юань' },
      { name: 'ДОСТАВКА', value: 3000, note: 'Доставка одной позиции, ₽' },
    ],
    createdAt: Date.now(),
    demo: true,
  };
  return { meta, sheets: [parts, sheet] };
}

const CARS: [string, string, string][] = [
  ['AUDI', 'A1', '8X (2010—2015)'],
  ['AUDI', 'A3', '8V (2012—2016)'],
  ['AUDI', 'A4', 'B9 (2015—2020)'],
  ['AUDI', 'Q5', 'FY (2017—2020)'],
  ['BMW', '3 серия', 'G20 (2018—2022)'],
  ['BMW', 'X5', 'G05 (2018—2023)'],
  ['MERCEDES-BENZ', 'E-класс', 'W213 (2016—2020)'],
  ['MERCEDES-BENZ', 'C-класс', 'W205 (2014—2018)'],
  ['VOLKSWAGEN', 'Tiguan', 'II (2016—2020)'],
  ['VOLKSWAGEN', 'Polo', 'VI (2020—2025)'],
  ['TOYOTA', 'Camry', 'XV70 (2017—2021)'],
  ['TOYOTA', 'RAV4', 'XA50 (2018—2025)'],
  ['KIA', 'Sportage', 'IV (2016—2022)'],
  ['LEXUS', 'RX', 'IV (2015—2019)'],
];
const PARTS = ['Фара галоген', 'Фара ксенон', 'Фара LED', 'Фонарь внешний LED', 'Фонарь внутренний LED', 'ПТФ левая', 'Решётка радиатора', 'Зеркало левое'];
const TRIMS = ['', '', 'S-LINE', 'MATRIX', 'Адаптив', 'Рестайлинг', 'Хэтч 3 = Хэтч 5', 'Седан'];
const SELLERS = ['Guangzhou Auto Light', 'Danyang Lamp', 'Xin Da Parts', 'Hengsheng'];

/** Лист как у клиента: цены от / принято / до, цена в Китае, проценка и маржа формулами от КУРС и ДОСТАВКА. */
function buildPartsSheet(): Sheet {
  const col = (name: string, w: number, extra: Partial<Column> = {}): Column => ({ id: uid(6), name, w, ...extra });
  const rub = { nf: { k: 'currency' as const, c: 'RUB' as const, d: 0 }, ha: 'center' as const };
  const cols = [
    col('Артикул', 108, { st: { ha: 'center', b: true } }),
    col('Марка', 128),
    col('Модель', 104),
    col('Поколение', 150, { st: { wrap: true } }),
    col('Фото', 110, { st: { ha: 'center' } }),
    col('Запчасть', 170, { st: { wrap: true } }),
    col('Исполнение', 140, { st: { wrap: true } }),
    col('Цена от', 100, { st: rub }),
    col('Принято', 104, { st: { ...rub, bg: '#f5edc9' } }),
    col('Цена до', 100, { st: rub }),
    col('Продавцы', 170, { st: { wrap: true, va: 'top' } }),
    col('Глубина', 84, { st: { ha: 'center' } }),
    col('CHN_1', 96, { st: { nf: { k: 'currency', c: 'CNY', d: 0 } } }),
    col('Проценка', 96, { st: { nf: { k: 'number', d: 1 }, ha: 'center' } }),
    col('Маржа', 110, { st: rub }),
    col('Дата проценки', 118, { st: { nf: { k: 'date' }, ha: 'center' } }),
  ];
  // как в Excel клиента: проценка = принято / (цена в Китае × курс + доставка), маржа = принято − себестоимость
  cols[13].formula = '=ЕСЛИ(M1>0;I1/(M1*КУРС+ДОСТАВКА);"?")';
  cols[14].formula = '=I1-M1*КУРС-ДОСТАВКА';

  let s = 11;
  const rnd = () => {
    s = (s * 1103515245 + 12345) & 0x7fffffff;
    return s / 0x7fffffff;
  };
  const rows = new Map<RowId, Row>();
  const order: RowId[] = [];
  const today = new Date(2026, 8, 18);
  let n = 1;
  for (const [make, model, gen] of CARS) {
    const count = 2 + Math.floor(rnd() * 3);
    for (let k = 0; k < count; k++) {
      const part = PARTS[(n * 3 + k) % PARTS.length];
      const base = /LED|ксенон/.test(part) ? 26000 + rnd() * 90000 : 9000 + rnd() * 30000;
      const accepted = Math.round(base / 1000) * 1000;
      const from = rnd() < 0.2 ? '-' : Math.round((accepted * (0.6 + rnd() * 0.3)) / 1000) * 1000;
      const to = rnd() < 0.15 ? '-' : Math.round((accepted * (1.1 + rnd() * 0.5)) / 1000) * 1000;
      const chn = Math.round((accepted / 12.1) * (0.25 + rnd() * 0.45) / 10) * 10;
      const cells: Record<string, Cell> = {
        [cols[0].id]: { v: n },
        [cols[1].id]: { v: make },
        [cols[2].id]: { v: model },
        [cols[3].id]: { v: gen },
        [cols[5].id]: { v: part },
        [cols[7].id]: { v: from },
        [cols[8].id]: { v: accepted },
        [cols[9].id]: { v: to },
        [cols[10].id]: { v: `${SELLERS[n % SELLERS.length]}\n${SELLERS[(n + 1) % SELLERS.length]}` },
        [cols[11].id]: { v: Math.floor(rnd() * 4) },
        [cols[12].id]: { v: chn },
        [cols[15].id]: { v: dateToSerial(new Date(today.getTime() - Math.floor(rnd() * 40) * 86400000)) },
      };
      const trim = TRIMS[(n * 5 + k) % TRIMS.length];
      if (trim) cells[cols[6].id] = { v: trim };
      if (n === 4) cells[cols[12].id] = { v: chn, note: 'Цена у продавца без коробки, с коробкой +¥80' };
      rows.set(`p${n}`, { id: `p${n}`, cells, h: 64 });
      order.push(`p${n}`);
      n++;
    }
  }
  for (let i = 0; i < 8; i++) {
    const id = uid(8);
    rows.set(id, { id, cells: {} });
    order.push(id);
  }
  return {
    id: uid(8),
    name: 'Запчасти',
    columns: cols,
    rowOrder: order,
    rows,
    keyColId: cols[0].id,
    frozen: 1,
    density: 'M',
    filters: {},
  };
}

/** Лист для проверки скорости: 20 000 строк × 14 столбцов, формула столбца. */
export function buildStressSheet(rowsCount = 20000): Sheet {
  const names = ['Артикул', 'Марка', 'Модель', 'Поколение', 'Страна', 'Стоимость юани', 'Курс', 'Стоимость рубли', 'Цвет', 'Размер', 'Остаток', 'Статус', 'Поставщик', 'Комментарий'];
  const columns: Column[] = names.map((n) => ({ id: uid(6), name: n, w: n.length > 10 ? 132 : 100 }));
  columns[5].st = { nf: { k: 'currency', c: 'CNY', d: 0 } };
  columns[6].st = { nf: { k: 'number', d: 2 } };
  columns[7].st = { nf: { k: 'currency', c: 'RUB', d: 0 }, b: true };
  columns[7].formula = '=F1*G1';
  const brands = ['Джинсы', 'Худи', 'Футболка', 'Куртка', 'Кроссовки', 'Кепка', 'Сумка', 'Рубашка'];
  const colors = ['чёрный', 'белый', 'индиго', 'графит', 'олива', 'песок', 'экрю', 'синий'];
  const statuses = ['В наличии', 'В пути', 'Заказать', 'Нет у поставщика'];
  const rows = new Map<RowId, Row>();
  const order: RowId[] = [];
  let s = 42;
  const rnd = () => {
    s = (s * 1103515245 + 12345) & 0x7fffffff;
    return s / 0x7fffffff;
  };
  for (let i = 0; i < rowsCount; i++) {
    const id = uid(8);
    const b = brands[Math.floor(rnd() * brands.length)];
    const cells: Record<string, Cell> = {
      [columns[0].id]: { v: 100000 + i },
      [columns[1].id]: { v: b },
      [columns[2].id]: { v: MODELS[Math.floor(rnd() * MODELS.length)] },
      [columns[3].id]: { v: FITS[Math.floor(rnd() * FITS.length)] },
      [columns[4].id]: { v: rnd() < 0.85 ? 'Китай' : 'Турция' },
      [columns[5].id]: { v: Math.round(100 + rnd() * 4000) },
      [columns[6].id]: { v: 12.75 },
      [columns[8].id]: { v: colors[Math.floor(rnd() * colors.length)] },
      [columns[9].id]: { v: ['XS', 'S', 'M', 'L', 'XL'][Math.floor(rnd() * 5)] },
      [columns[10].id]: { v: Math.floor(rnd() * 200) },
      [columns[11].id]: { v: statuses[Math.floor(rnd() * statuses.length)] },
      [columns[12].id]: { v: SUPPLIERS[Math.floor(rnd() * SUPPLIERS.length)][0] },
    };
    if (rnd() < 0.1) cells[columns[13].id] = { v: 'Проверить качество швов' };
    rows.set(id, { id, cells });
    order.push(id);
  }
  return {
    id: uid(8),
    name: `Тест ${rowsCount.toLocaleString('ru')}`,
    columns,
    rowOrder: order,
    rows,
    keyColId: columns[0].id,
    frozen: 1,
    density: 'S',
    filters: {},
  };
}
