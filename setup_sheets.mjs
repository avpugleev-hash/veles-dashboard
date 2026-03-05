import { google } from 'googleapis';
import { readFileSync } from 'fs';

const SA = JSON.parse(readFileSync('/home/user/veles-dashboard/service_account.json', 'utf8'));
const SPREADSHEET_ID = '13Xrmh-cfWFoR3-9TGFq2OoYZGmKxDhUhfqXWXC_SMKM';

const auth = new google.auth.GoogleAuth({
  credentials: SA,
  scopes: ['https://www.googleapis.com/auth/spreadsheets'],
});
const sheets = google.sheets({ version: 'v4', auth });

// ── helpers ────────────────────────────────────────────────────────────────

function hex(h) {
  const r = parseInt(h.slice(0,2),16)/255;
  const g = parseInt(h.slice(2,4),16)/255;
  const b = parseInt(h.slice(4,6),16)/255;
  return { red: r, green: g, blue: b };
}
function border(c='999999') {
  return { style:'SOLID', color: hex(c) };
}
function borders(c) {
  return { top: border(c), bottom: border(c), left: border(c), right: border(c) };
}

function headerFmt(bg) {
  return {
    backgroundColor: hex(bg),
    textFormat: { bold: true },
    borders: borders('999999'),
  };
}
function cellFmt(bg) {
  return {
    ...(bg ? { backgroundColor: hex(bg) } : {}),
    borders: borders('dddddd'),
  };
}

function repeatCell(sheetId, r1, r2, c1, c2, fmt) {
  return {
    repeatCell: {
      range: { sheetId, startRowIndex: r1, endRowIndex: r2,
               startColumnIndex: c1, endColumnIndex: c2 },
      cell: { userEnteredFormat: fmt },
      fields: 'userEnteredFormat(backgroundColor,textFormat,borders)',
    }
  };
}

function freeze(sheetId, rows=1) {
  return {
    updateSheetProperties: {
      properties: { sheetId, gridProperties: { frozenRowCount: rows } },
      fields: 'gridProperties.frozenRowCount',
    }
  };
}

function autoResize(sheetId, count) {
  return {
    autoResizeDimensions: {
      dimensions: { sheetId, dimension:'COLUMNS', startIndex:0, endIndex:count }
    }
  };
}

async function write(sheetName, range, values) {
  await sheets.spreadsheets.values.update({
    spreadsheetId: SPREADSHEET_ID,
    range: `'${sheetName}'!${range}`,
    valueInputOption: 'USER_ENTERED',
    requestBody: { values },
  });
}

async function clear(sheetName) {
  await sheets.spreadsheets.values.clear({
    spreadsheetId: SPREADSHEET_ID,
    range: `'${sheetName}'!A1:Z2000`,
    requestBody: {},
  });
}

async function batchUpdate(requests) {
  await sheets.spreadsheets.batchUpdate({
    spreadsheetId: SPREADSHEET_ID,
    requestBody: { requests },
  });
}

// ── get/create sheets ──────────────────────────────────────────────────────

const meta = await sheets.spreadsheets.get({ spreadsheetId: SPREADSHEET_ID });
const existing = {};
for (const s of meta.data.sheets) {
  existing[s.properties.title] = s.properties.sheetId;
}
console.log('Существующие листы:', Object.keys(existing));

const needed = ['МАРШРУТЫ', 'РЕЙСЫ', 'API_TRIPS'];
const toCreate = needed.filter(n => !(n in existing));
if (toCreate.length) {
  const resp = await batchUpdate(toCreate.map(title => ({ addSheet: { properties: { title } } })));
  // re-fetch IDs
  const m2 = await sheets.spreadsheets.get({ spreadsheetId: SPREADSHEET_ID });
  for (const s of m2.data.sheets) existing[s.properties.title] = s.properties.sheetId;
  console.log('Созданы:', toCreate);
}

// ══════════════════════════════════════════════════════════════════════════
// МАРШРУТЫ
// ══════════════════════════════════════════════════════════════════════════
const SID_R = existing['МАРШРУТЫ'];
await clear('МАРШРУТЫ');

const routesHeaders = [
  'Код маршрута','Маршрут',
  'Станция отправления','Станция назначения',
  'Время отправления (план)','Дней в пути',
  'Время прибытия (план)','Расстояние, км'
];
const routesData = [
  ['RSH-01','Красный Гуляй → Шакша','ст. Красный Гуляй','ст. Шакша','08:00',8,'16:00',1850],
  ['RSH-02','Ульяновск → Шакша','ст. Ульяновск-Центральный','ст. Шакша','14:00',7,'10:00',1430],
  ['RSH-03','Самара → Шакша','ст. Самара','ст. Шакша','10:00',6,'14:00',1200],
  ['RSH-04','Тольятти → Шакша','ст. Тольятти','ст. Шакша','12:00',7,'08:00',1280],
  ['RSH-05','Пенза → Шакша','ст. Пенза-1','ст. Шакша','09:00',5,'11:00',980],
  ['RSH-06','Уфа → Шакша','ст. Уфа','ст. Шакша','18:00',4,'08:00',640],
];

await write('МАРШРУТЫ', 'A1', [routesHeaders, ...routesData]);
await batchUpdate([
  freeze(SID_R),
  repeatCell(SID_R, 0, 1, 0, routesHeaders.length, headerFmt('cfe2ff')),
  repeatCell(SID_R, 1, routesData.length+1, 0, routesHeaders.length, cellFmt()),
  autoResize(SID_R, routesHeaders.length),
]);
console.log('✅ МАРШРУТЫ заполнен');

// ══════════════════════════════════════════════════════════════════════════
// РЕЙСЫ
// ══════════════════════════════════════════════════════════════════════════
const SID_T = existing['РЕЙСЫ'];
await clear('РЕЙСЫ');

const tripsHeaders = [
  'Код рейса',             // A 0  ★ авто
  'Маршрут (код)',         // B 1
  'Маршрут',               // C 2  ★ авто
  'Дата отправления',      // D 3
  'Время отправления',     // E 4  ★ авто
  'Дата прибытия (план)',  // F 5  ★ авто
  'Время прибытия (план)', // G 6  ★ авто
  'Вагонов всего',         // H 7
  'Состояние',             // I 8
  'Статус погрузки',       // J 9
  'Погружено вагонов',     // K 10
  'Текущая станция',       // L 11 ★ n8n
  'Осталось, км',          // M 12 ★ n8n
  'Прогноз прибытия',      // N 13 ★ n8n
  'Факт отправления',      // O 14
  'Факт прибытия',         // P 15
  'Дислокация обновлена',  // Q 16 ★ n8n
  'Примечание',            // R 17
];

const tripsDemo = [
  ['RSH-01-0315-01','RSH-01','Красный Гуляй → Шакша',
   '15.03.2026','08:00','23.03.2026','16:00',
   45,'В пути','','',
   'ст. Канаш',620,'23.03.2026 16:00',
   '15.03.2026','','05.03.2026 14:30',''],

  ['RSH-02-0318-01','RSH-02','Ульяновск → Шакша',
   '18.03.2026','14:00','25.03.2026','10:00',
   32,'На погрузке','Задерживается',28,
   '','','',
   '','','','Часть вагонов не подана. Ожидается 19.03.'],

  ['RSH-03-0322-01','RSH-03','Самара → Шакша',
   '22.03.2026','10:00','28.03.2026','14:00',
   50,'В плане','В плане','',
   '','','',
   '','','',''],
];

await write('РЕЙСЫ', 'A1', [tripsHeaders, ...tripsDemo]);

// авто-колонки: 0,2,4,5,6,11,12,13,16
const autoCols = [0,2,4,5,6,11,12,13,16];
await batchUpdate([
  freeze(SID_T),
  repeatCell(SID_T, 0, 1, 0, tripsHeaders.length, headerFmt('e8f0fe')),
  ...autoCols.map(c => repeatCell(SID_T, 0, 1, c, c+1, headerFmt('d0e8ff'))),
  repeatCell(SID_T, 1, tripsDemo.length+1, 0, tripsHeaders.length, cellFmt()),
  autoResize(SID_T, tripsHeaders.length),
]);
console.log('✅ РЕЙСЫ заполнен');

// ══════════════════════════════════════════════════════════════════════════
// API_TRIPS
// ══════════════════════════════════════════════════════════════════════════
const SID_A = existing['API_TRIPS'];
await clear('API_TRIPS');

const apiHeaders = [
  'trip_id','route_name','current_state','loading_status',
  'wagons_total','wagons_loaded',
  'departure_plan','arrival_plan',
  'actual_departure','actual_arrival',
  'current_station','km_remaining','eta',
  'dislocation_updated_at','alert_text','overall_status'
];

const apiData = [
  ['RSH-01-0315-01','Красный Гуляй → Шакша','В пути','',
   45,'',
   '15.03.2026 08:00','23.03.2026 16:00',
   '15.03.2026','',
   'ст. Канаш',620,'23.03.2026 16:00',
   '05.03.2026 14:30','','OK'],

  ['RSH-02-0318-01','Ульяновск → Шакша','На погрузке','Задерживается',
   32,28,
   '18.03.2026 14:00','25.03.2026 10:00',
   '','',
   '','','',
   '','Часть вагонов не подана. Ожидается 19.03.','WARN'],

  ['RSH-03-0322-01','Самара → Шакша','В плане','В плане',
   50,'',
   '22.03.2026 10:00','28.03.2026 14:00',
   '','',
   '','','',
   '','','OK'],
];

await write('API_TRIPS', 'A1', [apiHeaders, ...apiData]);
await batchUpdate([
  freeze(SID_A),
  repeatCell(SID_A, 0, 1, 0, apiHeaders.length, headerFmt('fce8b2')),
  repeatCell(SID_A, 1, apiData.length+1, 0, apiHeaders.length, cellFmt()),
  autoResize(SID_A, apiHeaders.length),
]);
console.log('✅ API_TRIPS заполнен');

console.log('\n🎉 Готово! Все листы созданы и заполнены.');
