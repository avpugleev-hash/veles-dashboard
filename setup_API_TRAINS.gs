/************************************************************
 * VELES — Variant A (правильная схема)
 * Apps Script НЕ считает логику дислокации.
 * Он только:
 * 1) создаёт/чинит лист API_TRAINS (заголовки)
 * 2) (опционально) создаёт демо-строки
 * 3) даёт кнопку "очистить витрину"
 *
 * Логику расчёта и запись строк делает n8n.
 ************************************************************/
const CFG = {
  API_SHEET: "API_TRAINS",
  API_HEADERS: [
    "train_name",
    "train_id",
    "route_name",
    "current_state",
    "current_station",
    "dislocation_updated_at",
    "wagons_total",
    "wagons_in_loading",
    "wagons_in_transit",
    "wagons_in_unloading",
    "norm_unload_days",
    "fact_unload_days",
    "overall_status",
    // новые колонки для дашборда
    "departure_plan",
    "departure_fact",
    "arrival_plan",
    "arrival_fact",
    "arrival_eta",
    "wagon_ids",
    "alert_text",
  ],
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("VELES API")
    .addItem("Инициализировать API_TRAINS", "initApiTrains")
    .addItem("Очистить данные API_TRAINS", "clearApiTrainsData")
    .addSeparator()
    .addItem("Заполнить демо (3 рейса)", "fillDemoApiTrains")
    .addToUi();
}

/**
 * Создаёт/чинит лист API_TRAINS и заголовки (строка 1)
 */
function initApiTrains() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(CFG.API_SHEET) || ss.insertSheet(CFG.API_SHEET);
  sh.getRange(1, 1, 1, CFG.API_HEADERS.length).setValues([CFG.API_HEADERS]);
  sh.setFrozenRows(1);
  sh.autoResizeColumns(1, CFG.API_HEADERS.length);
  SpreadsheetApp.flush();
}

/**
 * Очищает только данные (со 2 строки), заголовки оставляет
 */
function clearApiTrainsData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(CFG.API_SHEET);
  if (!sh) return;
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return;
  sh.getRange(2, 1, lastRow - 1, CFG.API_HEADERS.length).clearContent();
}

/**
 * Демо-данные — 3 рейса с новыми полями дат
 */
function fillDemoApiTrains() {
  initApiTrains();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(CFG.API_SHEET);
  const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd.MM.yyyy HH:mm");

  // train_name, train_id, route_name, current_state, current_station,
  // dislocation_updated_at, wagons_total, wagons_in_loading, wagons_in_transit,
  // wagons_in_unloading, norm_unload_days, fact_unload_days, overall_status,
  // departure_plan, departure_fact, arrival_plan, arrival_fact, arrival_eta,
  // wagon_ids, alert_text
  const demo = [
    [
      "Рейс №1", "ПКРС-01", "Красный Гуляй → Шакша", "В пути", "ст. Ульяновск",
      now, 50, 0, 50, 0, 1.5, 0, "BAD",
      "10.11.2025", "12.11.2025", "18.11.2025", "", "16.11",
      "НКЛ-2025-08451;НКЛ-2025-08452;НКЛ-2025-08453",
      "Поезд отправился 12.11 вместо 10.11. Прибытие сдвинуто на 16.11."
    ],
    [
      "Рейс №2", "ПКРС-02", "Красный Гуляй → Шакша", "Погрузка", "ст. Красный Гуляй",
      now, 50, 50, 0, 0, 1.5, 0, "OK",
      "20.11.2025", "", "28.11.2025", "", "",
      "НКЛ-2025-12001;НКЛ-2025-12002",
      ""
    ],
    [
      "Рейс №3", "ПКРС-03", "Красный Гуляй → Шакша", "Под выгрузкой", "ст. Шакша",
      now, 50, 0, 0, 32, 1.5, 0.8, "WARN",
      "05.11.2025", "05.11.2025", "12.11.2025", "13.11.2025", "",
      "НКЛ-2025-03001;НКЛ-2025-03002",
      "Ожидается завершение выгрузки. Подготовьте персонал."
    ],
  ];

  clearApiTrainsData();
  sh.getRange(2, 1, demo.length, CFG.API_HEADERS.length).setValues(demo);
}
