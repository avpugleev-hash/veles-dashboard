/**
 * ВЕЛЕС • Sisecam Dispatcher — Apps Script
 *
 * Листы:
 *   МАРШРУТЫ  — справочник маршрутов (заполняется один раз)
 *   РЕЙСЫ     — план рейсов на месяц (диспетчер ведёт вручную)
 *   API_TRIPS — публичный CSV для дашборда (скрипт генерирует; n8n дополняет дислокацию)
 *
 * Установка:
 *   1. Расширения → Apps Script → вставьте этот файл
 *   2. Сохраните, запустите installAll() один раз
 *   3. Опубликуйте API_TRIPS как CSV: Файл → Поделиться → Опубликовать в интернете → API_TRIPS → CSV
 *   4. Вставьте URL в дашборд
 */

// ─── Имена листов ─────────────────────────────────────────────────────────────

const ROUTES_SHEET = 'МАРШРУТЫ';
const TRIPS_SHEET  = 'РЕЙСЫ';
const API_SHEET    = 'API_TRIPS';

// ─── Колонки листа РЕЙСЫ (1-based) ──────────────────────────────────────────
// Авто-заполняемые колонки помечены ★

const T = {
  TRIP_ID           : 1,   // A  Код рейса           ★ авто
  ROUTE_ID          : 2,   // B  Маршрут (код)          диспетчер выбирает из списка
  ROUTE_NAME        : 3,   // C  Маршрут (название)   ★ авто из справочника
  DEPARTURE_DATE    : 4,   // D  Дата отправления       диспетчер
  DEPARTURE_TIME    : 5,   // E  Время отправления     ★ авто из справочника
  ARRIVAL_DATE      : 6,   // F  Дата прибытия (план)  ★ авто = D + дней в пути
  ARRIVAL_TIME      : 7,   // G  Время прибытия (план) ★ авто из справочника
  WAGONS_TOTAL      : 8,   // H  Вагонов всего           диспетчер
  CURRENT_STATE     : 9,   // I  Состояние               диспетчер / n8n
  LOADING_STATUS    : 10,  // J  Статус погрузки          диспетчер (при состоянии «На погрузке»)
  WAGONS_LOADED     : 11,  // K  Погружено вагонов        диспетчер (если не все)
  CURRENT_STATION   : 12,  // L  Текущая станция        ★ n8n
  KM_REMAINING      : 13,  // M  Осталось, км           ★ n8n
  ETA               : 14,  // N  Прогноз прибытия       ★ n8n
  ACTUAL_DEPARTURE  : 15,  // O  Факт отправления         диспетчер
  ACTUAL_ARRIVAL    : 16,  // P  Факт прибытия          ★ n8n / диспетчер
  DISLO_UPDATED_AT  : 17,  // Q  Дислокация обновлена   ★ n8n
  ALERT_TEXT        : 18,  // R  Примечание               диспетчер
};

// ─── Колонки листа МАРШРУТЫ (0-based для массива) ────────────────────────────
const R = {
  ID          : 0,  // Код маршрута
  NAME        : 1,  // Маршрут
  FROM        : 2,  // Станция отправления
  TO          : 3,  // Станция назначения
  DEP_TIME    : 4,  // Время отправления (план)
  TRANSIT_DAYS: 5,  // Дней в пути
  ARR_TIME    : 6,  // Время прибытия (план)
  DISTANCE_KM : 7,  // Расстояние, км
};

const STATES          = ['В плане','На погрузке','В пути','Под выгрузкой','Выполнен'];
const LOADING_STATUSES = ['В плане','Задерживается'];

// ─── Установка ───────────────────────────────────────────────────────────────

function installAll() {
  createRoutesSheet();
  createTripsSheet();
  createApiSheet();
  addMenu();
  SpreadsheetApp.getActiveSpreadsheet()
    .toast('✅ Готово! Заполните МАРШРУТЫ, затем РЕЙСЫ.', 'ВЕЛЕС Dispatcher', 10);
}

// ─── Лист МАРШРУТЫ ───────────────────────────────────────────────────────────

function createRoutesSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(ROUTES_SHEET);
  if (!sh) sh = ss.insertSheet(ROUTES_SHEET);
  sh.clearContents();

  const headers = [
    'Код маршрута','Маршрут',
    'Станция отправления','Станция назначения',
    'Время отправления (план)','Дней в пути',
    'Время прибытия (план)','Расстояние, км'
  ];
  sh.getRange(1,1,1,headers.length)
    .setValues([headers]).setFontWeight('bold').setBackground('#cfe2ff');

  const demo = [
    ['RSH-01','Красный Гуляй → Шакша','ст. Красный Гуляй','ст. Шакша','08:00',8,'16:00',1850],
    ['RSH-02','Ульяновск → Шакша','ст. Ульяновск-Центральный','ст. Шакша','14:00',7,'10:00',1430],
    ['RSH-03','Самара → Шакша','ст. Самара','ст. Шакша','10:00',6,'14:00',1200],
    ['RSH-04','Тольятти → Шакша','ст. Тольятти','ст. Шакша','12:00',7,'08:00',1280],
    ['RSH-05','Пенза → Шакша','ст. Пенза-1','ст. Шакша','09:00',5,'11:00',980],
    ['RSH-06','Уфа → Шакша','ст. Уфа','ст. Шакша','18:00',4,'08:00',640],
  ];
  sh.getRange(2,1,demo.length,headers.length).setValues(demo);
  sh.setFrozenRows(1);
  sh.autoResizeColumns(1,headers.length);
}

// ─── Лист РЕЙСЫ ──────────────────────────────────────────────────────────────

function createTripsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(TRIPS_SHEET);
  if (!sh) sh = ss.insertSheet(TRIPS_SHEET);
  sh.clearContents();

  // Заголовки
  const headers = [
    'Код рейса',            // A  T.TRIP_ID          ★ авто
    'Маршрут (код)',        // B  T.ROUTE_ID
    'Маршрут',              // C  T.ROUTE_NAME        ★ авто
    'Дата отправления',     // D  T.DEPARTURE_DATE
    'Время отправления',    // E  T.DEPARTURE_TIME    ★ авто
    'Дата прибытия (план)', // F  T.ARRIVAL_DATE      ★ авто
    'Время прибытия (план)',// G  T.ARRIVAL_TIME      ★ авто
    'Вагонов всего',        // H  T.WAGONS_TOTAL
    'Состояние',            // I  T.CURRENT_STATE
    'Статус погрузки',      // J  T.LOADING_STATUS
    'Погружено вагонов',    // K  T.WAGONS_LOADED
    'Текущая станция',      // L  T.CURRENT_STATION   ★ n8n
    'Осталось, км',         // M  T.KM_REMAINING      ★ n8n
    'Прогноз прибытия',     // N  T.ETA               ★ n8n
    'Факт отправления',     // O  T.ACTUAL_DEPARTURE
    'Факт прибытия',        // P  T.ACTUAL_ARRIVAL
    'Дислокация обновлена', // Q  T.DISLO_UPDATED_AT  ★ n8n
    'Примечание',           // R  T.ALERT_TEXT
  ];

  const hRange = sh.getRange(1,1,1,headers.length);
  hRange.setValues([headers]).setFontWeight('bold').setBackground('#e8f0fe');

  // Выделить авто-колонки голубым
  const autoCols = [T.TRIP_ID, T.ROUTE_NAME, T.DEPARTURE_TIME, T.ARRIVAL_DATE, T.ARRIVAL_TIME,
                    T.CURRENT_STATION, T.KM_REMAINING, T.ETA, T.DISLO_UPDATED_AT];
  autoCols.forEach(c => sh.getRange(1,c).setBackground('#d0e8ff'));

  sh.setFrozenRows(1);

  // Валидация: Маршрут (код) — из МАРШРУТЫ!A
  const routeVal = SpreadsheetApp.newDataValidation()
    .requireValueInRange(ss.getSheetByName(ROUTES_SHEET).getRange('A2:A200'), true)
    .setAllowInvalid(false).build();
  sh.getRange(2, T.ROUTE_ID, 500).setDataValidation(routeVal);

  // Валидация: Состояние
  sh.getRange(2, T.CURRENT_STATE, 500)
    .setDataValidation(SpreadsheetApp.newDataValidation()
      .requireValueInList(STATES, true).setAllowInvalid(false).build());

  // Валидация: Статус погрузки
  sh.getRange(2, T.LOADING_STATUS, 500)
    .setDataValidation(SpreadsheetApp.newDataValidation()
      .requireValueInList(LOADING_STATUSES, true).setAllowInvalid(false).build());

  // Форматы
  sh.getRange(2, T.DEPARTURE_DATE, 500).setNumberFormat('dd.MM.yyyy');
  sh.getRange(2, T.ARRIVAL_DATE,   500).setNumberFormat('dd.MM.yyyy');
  sh.getRange(2, T.ACTUAL_DEPARTURE, 500).setNumberFormat('dd.MM.yyyy HH:mm');
  sh.getRange(2, T.ACTUAL_ARRIVAL,   500).setNumberFormat('dd.MM.yyyy HH:mm');

  sh.autoResizeColumns(1, headers.length);
}

// ─── Лист API_TRIPS ──────────────────────────────────────────────────────────

function createApiSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(API_SHEET);
  if (!sh) sh = ss.insertSheet(API_SHEET);
  sh.clearContents();

  // Английские имена — дашборд читает по ним
  const headers = [
    'trip_id','route_name','current_state','loading_status',
    'wagons_total','wagons_loaded',
    'departure_plan','arrival_plan',
    'actual_departure','actual_arrival',
    'current_station','km_remaining','eta',
    'dislocation_updated_at','alert_text','overall_status'
  ];
  sh.getRange(1,1,1,headers.length).setValues([headers])
    .setFontWeight('bold').setBackground('#fce8b2');
  sh.setFrozenRows(1);
  sh.autoResizeColumns(1, headers.length);
}

// ─── onEdit: авто-заполнение ─────────────────────────────────────────────────

function onEdit(e) {
  const sh = e.range.getSheet();
  if (sh.getName() !== TRIPS_SHEET) return;
  const row = e.range.getRow();
  const col = e.range.getColumn();
  if (row < 2) return;

  if (col === T.ROUTE_ID || col === T.DEPARTURE_DATE) {
    fillAutoColumns(sh, row);
  }
}

function fillAutoColumns(sh, row) {
  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const routeSh  = ss.getSheetByName(ROUTES_SHEET);
  const routeId  = sh.getRange(row, T.ROUTE_ID).getValue();
  const depDateV = sh.getRange(row, T.DEPARTURE_DATE).getValue();
  if (!routeId) return;

  // Найти маршрут в справочнике
  const routeData = routeSh.getDataRange().getValues();
  let route = null;
  for (let i = 1; i < routeData.length; i++) {
    if (String(routeData[i][R.ID]) === String(routeId)) { route = routeData[i]; break; }
  }
  if (!route) return;

  sh.getRange(row, T.ROUTE_NAME).setValue(route[R.NAME]);
  sh.getRange(row, T.DEPARTURE_TIME).setValue(route[R.DEP_TIME]);
  sh.getRange(row, T.ARRIVAL_TIME).setValue(route[R.ARR_TIME]);

  // Авто-ID рейса если пустой
  if (!sh.getRange(row, T.TRIP_ID).getValue()) {
    const month = depDateV ? Utilities.formatDate(new Date(depDateV),'GMT+3','MMdd') : '0000';
    sh.getRange(row, T.TRIP_ID).setValue(`${routeId}-${month}-${String(row-1).padStart(2,'0')}`);
  }

  // Дата прибытия = дата отправления + дней в пути
  if (depDateV) {
    const dep  = new Date(depDateV);
    const days = parseInt(route[R.TRANSIT_DAYS]) || 0;
    dep.setDate(dep.getDate() + days);
    sh.getRange(row, T.ARRIVAL_DATE).setValue(dep).setNumberFormat('dd.MM.yyyy');
  }
}

// ─── Синхронизация РЕЙСЫ → API_TRIPS ────────────────────────────────────────

function syncToApi() {
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const tripsSh = ss.getSheetByName(TRIPS_SHEET);
  const apiSh   = ss.getSheetByName(API_SHEET);
  if (!tripsSh || !apiSh) {
    SpreadsheetApp.getUi().alert('Листы РЕЙСЫ или API_TRIPS не найдены. Запустите installAll().');
    return;
  }

  const tripsData = tripsSh.getDataRange().getValues();
  if (tripsData.length < 2) return;

  // Читаем n8n-данные из API_TRIPS чтобы не затереть при синхронизации
  const apiData = apiSh.getDataRange().getValues();
  const apiH    = apiData[0];
  const n8nMap  = {}; // trip_id → n8n-поля
  const A = (name) => apiH.indexOf(name);

  for (let i = 1; i < apiData.length; i++) {
    const id = apiData[i][0];
    if (!id) continue;
    n8nMap[id] = {
      current_station       : apiData[i][A('current_station')],
      km_remaining          : apiData[i][A('km_remaining')],
      eta                   : apiData[i][A('eta')],
      dislocation_updated_at: apiData[i][A('dislocation_updated_at')],
    };
  }

  const rows = [];
  for (let i = 1; i < tripsData.length; i++) {
    const r      = tripsData[i];
    const tripId = r[T.TRIP_ID - 1];
    if (!tripId) continue;

    const n8n    = n8nMap[tripId] || {};
    const state  = r[T.CURRENT_STATE  - 1] || '';
    const depPlan = joinDateTime(r[T.DEPARTURE_DATE - 1], r[T.DEPARTURE_TIME - 1]);
    const arrPlan = joinDateTime(r[T.ARRIVAL_DATE   - 1], r[T.ARRIVAL_TIME   - 1]);

    // overall_status для алертов
    let overall = 'OK';
    if (r[T.LOADING_STATUS - 1] === 'Задерживается') overall = 'WARN';
    if (r[T.ALERT_TEXT - 1])                         overall = 'BAD';

    rows.push([
      tripId,
      r[T.ROUTE_NAME       - 1] || '',
      state,
      r[T.LOADING_STATUS   - 1] || '',
      r[T.WAGONS_TOTAL     - 1] || '',
      r[T.WAGONS_LOADED    - 1] || '',
      depPlan,
      arrPlan,
      formatDateTime(r[T.ACTUAL_DEPARTURE - 1]),
      formatDateTime(r[T.ACTUAL_ARRIVAL   - 1]),
      // n8n-поля: берём из API_TRIPS если там свежее, иначе из РЕЙСЫ
      n8n.current_station        || r[T.CURRENT_STATION  - 1] || '',
      n8n.km_remaining           || r[T.KM_REMAINING     - 1] || '',
      n8n.eta                    || r[T.ETA              - 1] || '',
      n8n.dislocation_updated_at || r[T.DISLO_UPDATED_AT - 1] || '',
      r[T.ALERT_TEXT       - 1] || '',
      overall,
    ]);
  }

  // Перезаписать данные (заголовок сохраняем)
  const lastRow = apiSh.getLastRow();
  if (lastRow > 1) apiSh.getRange(2, 1, lastRow - 1, apiSh.getLastColumn()).clearContent();
  if (rows.length > 0) {
    apiSh.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
  }

  SpreadsheetApp.getActiveSpreadsheet()
    .toast(`✅ Синхронизировано: ${rows.length} рейс(ов)`, 'ВЕЛЕС', 5);
}

// ─── Меню ─────────────────────────────────────────────────────────────────────

function addMenu() {
  SpreadsheetApp.getUi().createMenu('🚂 ВЕЛЕС')
    .addItem('Синхронизировать API_TRIPS', 'syncToApi')
    .addItem('Заполнить демо-рейсы (3 шт)', 'fillDemoTrips')
    .addSeparator()
    .addItem('Переустановить листы', 'installAll')
    .addToUi();
}

function onOpen() { addMenu(); }

// ─── Демо-рейсы ───────────────────────────────────────────────────────────────

function fillDemoTrips() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(TRIPS_SHEET);
  if (!sh) {
    SpreadsheetApp.getUi().alert('Сначала запустите installAll()');
    return;
  }

  // trip_id, route_id, route_name, dep_date, dep_time, arr_date, arr_time,
  // wagons, state, load_status, wagons_loaded,
  // cur_station, km_rem, eta, act_dep, act_arr, dislo_upd, alert
  const demo = [
    ['RSH-01-0315-01','RSH-01','Красный Гуляй → Шакша',
     pd('15.03.2026'),'08:00',pd('23.03.2026'),'16:00',
     45,'В пути','','',
     'ст. Канаш',620,'23.03.2026 16:00',
     pd('15.03.2026'),'','05.03.2026 14:30',''],

    ['RSH-02-0318-01','RSH-02','Ульяновск → Шакша',
     pd('18.03.2026'),'14:00',pd('25.03.2026'),'10:00',
     32,'На погрузке','Задерживается',28,
     '','','',
     '','','','Часть вагонов не подана. Ожидается 19.03.'],

    ['RSH-03-0322-01','RSH-03','Самара → Шакша',
     pd('22.03.2026'),'10:00',pd('28.03.2026'),'14:00',
     50,'В плане','В плане','',
     '','','',
     '','','',''],
  ];

  sh.getRange(2, 1, demo.length, demo[0].length).setValues(demo);
  sh.getRange(2, T.DEPARTURE_DATE, demo.length).setNumberFormat('dd.MM.yyyy');
  sh.getRange(2, T.ARRIVAL_DATE,   demo.length).setNumberFormat('dd.MM.yyyy');

  syncToApi();
}

// ─── Утилиты ──────────────────────────────────────────────────────────────────

/** dd.MM.yyyy → Date */
function pd(str) {
  if (!str) return '';
  const p = str.split('.');
  if (p.length === 3) return new Date(+p[2], +p[1]-1, +p[0]);
  return str;
}

function formatDate(val) {
  if (!val) return '';
  try { return Utilities.formatDate(new Date(val), 'GMT+3', 'dd.MM.yyyy'); }
  catch(e) { return String(val); }
}

function formatDateTime(val) {
  if (!val) return '';
  try {
    const d = new Date(val);
    if (isNaN(d.getTime())) return String(val);
    return Utilities.formatDate(d, 'GMT+3', 'dd.MM.yyyy HH:mm');
  }
  catch(e) { return String(val); }
}

/** Склеить дату + время в строку "dd.MM.yyyy HH:mm" */
function joinDateTime(dateVal, timeStr) {
  const d = formatDate(dateVal);
  if (!d) return '';
  return timeStr ? `${d} ${timeStr}` : d;
}
