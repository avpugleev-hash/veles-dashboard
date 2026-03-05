/**
 * ВЕЛЕС — Настройка листа API_TRAINS
 *
 * Как использовать:
 * 1. Откройте вашу таблицу Google Sheets
 * 2. Меню → Расширения → Apps Script
 * 3. Вставьте этот код, нажмите ▶ Run (функция setupSheet)
 * 4. Разрешите доступ при запросе
 * 5. Лист "API_TRAINS" создастся автоматически с заголовками и примерами
 */

function setupSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Удаляем старый лист если есть, создаём новый
  let sheet = ss.getSheetByName("API_TRAINS");
  if (sheet) {
    const ok = Browser.msgBox(
      "Лист API_TRAINS уже существует",
      "Пересоздать (все данные будут удалены)?",
      Browser.Buttons.YES_NO
    );
    if (ok !== Browser.Buttons.YES) return;
    ss.deleteSheet(sheet);
  }
  sheet = ss.insertSheet("API_TRAINS");

  // ── ЗАГОЛОВКИ ──────────────────────────────────────────────────────────────
  const HEADERS = [
    ["train_id",              "ID рейса (напр. №РСХ-08)"],
    ["train_name",            "Название рейса"],
    ["route_name",            "Маршрут (напр. Самара → Шакша)"],
    ["current_state",         "Статус: В пути / Под выгрузкой / Выполнен / В плане"],
    ["current_station",       "Текущая станция (из дислокации)"],
    ["departure_plan",        "Плановая дата отправки (дд.мм.гггг)"],
    ["departure_fact",        "Фактическая дата отправки (дд.мм.гггг)"],
    ["arrival_plan",          "Плановое прибытие на выгрузку (дд.мм.гггг)"],
    ["arrival_fact",          "Фактическое прибытие (дд.мм.гггг)"],
    ["arrival_eta",           "Ожидаемое прибытие при задержке (дд.мм или дд.мм.гггг)"],
    ["wagons_total",          "Всего вагонов в рейсе"],
    ["wagons_in_loading",     "Вагонов на погрузке"],
    ["wagons_in_transit",     "Вагонов в пути"],
    ["wagons_in_unloading",   "Вагонов выгружено"],
    ["norm_unload_days",      "Норма выгрузки (суток)"],
    ["fact_unload_days",      "Факт выгрузки (суток)"],
    ["wagon_ids",             "Номера вагонов через ; (напр. НКЛ-2025-08451;НКЛ-2025-08452)"],
    ["overall_status",        "Итоговый статус: OK / WARN / BAD / ALERT"],
    ["dislocation_updated_at","Дата обновления дислокации"],
    ["alert_text",            "Текст для баннера-предупреждения (если overall_status = BAD/WARN)"],
  ];

  const headerRow  = HEADERS.map(h => h[0]);
  const commentRow = HEADERS.map(h => h[1]);

  // Строка 1 — машиночитаемые ключи (используются дашбордом)
  sheet.getRange(1, 1, 1, headerRow.length).setValues([headerRow]);

  // Строка 2 — комментарии (серый, мелкий, не влияет на данные — но УДАЛИТЕ её перед публикацией
  //             или начинайте данные с 3-й строки и измените slice(1) на slice(2) в index.html)
  sheet.getRange(2, 1, 1, commentRow.length).setValues([commentRow]);

  // ── ПРИМЕРЫ ДАННЫХ (строки 3-8) ───────────────────────────────────────────
  const SAMPLE = [
    ["№РСХ-08","Рейс РСХ-08","Красный Гуляй → Шакша","В пути","ст. КАНАШ",
     "10.11.2025","12.11.2025","18.11.2025","","16.11",
     45,0,45,0,3,"","НКЛ-2025-08451;НКЛ-2025-08452;НКЛ-2025-08453",
     "BAD","12.11.2025 14:30","Поезд отправился 12.11 вместо планового 10.11. Прибытие сдвинуто на 16.11."],

    ["№РСХ-12","Рейс РСХ-12","Ульяновск → Шакша","Под выгрузкой","ст. Шакша",
     "08.11.2025","08.11.2025","16.11.2025","13.11.2025","",
     32,0,14,18,3,3,"НКЛ-2025-12001;НКЛ-2025-12002",
     "WARN","13.11.2025 09:00","Ожидается прибытие на ст. Шакша в течение 2 часов. Подготовьте персонал."],

    ["№РСХ-05","Рейс РСХ-05","Самара → Шакша","Выполнен","",
     "01.11.2025","01.11.2025","09.11.2025","09.11.2025","",
     50,0,0,50,3,3,"НКЛ-2025-05001;НКЛ-2025-05002;НКЛ-2025-05003;НКЛ-2025-05004",
     "OK","09.11.2025 18:00",""],

    ["№РСХ-15","Рейс РСХ-15","Пенза → Шакша","В плане","",
     "20.11.2025","","28.11.2025","","",
     38,38,0,0,3,"","НКЛ-2025-15001;НКЛ-2025-15002",
     "OK","",""],

    ["№РСХ-03","Рейс РСХ-03","Тольятти → Шакша","В пути","ст. Пенза-1",
     "05.11.2025","05.11.2025","12.11.2025","","12.11",
     60,0,60,0,4,"","НКЛ-2025-03001;НКЛ-2025-03002;НКЛ-2025-03003",
     "OK","11.11.2025 20:15",""],

    ["№РСХ-20","Рейс РСХ-20","Уфа → Шакша","В плане","",
     "25.11.2025","","04.12.2025","","",
     44,44,0,0,3,"","НКЛ-2025-20001;НКЛ-2025-20002",
     "OK","",""],
  ];

  sheet.getRange(3, 1, SAMPLE.length, SAMPLE[0].length).setValues(SAMPLE);

  // ── ФОРМАТИРОВАНИЕ ─────────────────────────────────────────────────────────
  const totalCols = headerRow.length;
  const totalRows = 2 + SAMPLE.length;

  // Заголовок (строка 1): синий фон, белый текст, жирный
  const headerRange = sheet.getRange(1, 1, 1, totalCols);
  headerRange.setBackground("#1d4ed8").setFontColor("#ffffff").setFontWeight("bold");

  // Комментарии (строка 2): светло-серый, мелкий курсив
  const commentRange = sheet.getRange(2, 1, 1, totalCols);
  commentRange.setBackground("#f1f5f9").setFontColor("#94a3b8")
    .setFontSize(9).setFontStyle("italic");

  // Данные: чередующиеся строки
  for (let i = 3; i <= totalRows; i++) {
    sheet.getRange(i, 1, 1, totalCols)
      .setBackground(i % 2 === 0 ? "#f8fafc" : "#ffffff");
  }

  // Подсветка overall_status (колонка 18)
  const statusCol = 18;
  for (let i = 3; i <= totalRows; i++) {
    const cell   = sheet.getRange(i, statusCol);
    const val    = cell.getValue();
    if (val === "BAD"  || val === "ALERT") cell.setBackground("#fee2e2").setFontColor("#b91c1c").setFontWeight("bold");
    if (val === "WARN" || val === "WARNING") cell.setBackground("#fef9c3").setFontColor("#92400e").setFontWeight("bold");
    if (val === "OK")  cell.setBackground("#dcfce7").setFontColor("#15803d");
  }

  // Закрепить первую строку, авторазмер колонок
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, totalCols);

  // Установить ширину для длинных колонок вручную
  sheet.setColumnWidth(3,  220); // route_name
  sheet.setColumnWidth(17, 280); // wagon_ids
  sheet.setColumnWidth(20, 320); // alert_text

  // ── ИНСТРУКЦИЯ НА ОТДЕЛЬНОМ ЛИСТЕ ─────────────────────────────────────────
  let infoSheet = ss.getSheetByName("📋 Инструкция");
  if (!infoSheet) infoSheet = ss.insertSheet("📋 Инструкция");
  infoSheet.clearContents();

  const instructions = [
    ["ВЕЛЕС — Инструкция по заполнению листа API_TRAINS"],
    [""],
    ["ВАЖНО: Строка 1 — заголовки (не менять!). Строка 2 — комментарии (можно удалить). Данные с строки 3."],
    [""],
    ["Поле", "Обязательное?", "Значения / Формат"],
    ["train_id",            "✅ Да",  "Уникальный ID рейса, напр. №РСХ-08"],
    ["train_name",          "Нет",   "Произвольное название"],
    ["route_name",          "✅ Да",  "Маршрут, напр. Самара → Шакша"],
    ["current_state",       "✅ Да",  "Ровно одно из: В пути / Под выгрузкой / Выполнен / В плане"],
    ["current_station",     "Нет",   "Название станции из дислокации, напр. ст. КАНАШ"],
    ["departure_plan",      "Нет",   "дд.мм.гггг — плановая дата отправки"],
    ["departure_fact",      "Нет",   "дд.мм.гггг — фактическая дата отправки"],
    ["arrival_plan",        "Нет",   "дд.мм.гггг — плановое прибытие на выгрузку"],
    ["arrival_fact",        "Нет",   "дд.мм.гггг — фактическое прибытие"],
    ["arrival_eta",         "Нет",   "дд.мм или дд.мм.гггг — ожидаемое прибытие при задержке"],
    ["wagons_total",        "✅ Да",  "Число — всего вагонов в рейсе"],
    ["wagons_in_loading",   "Нет",   "Число — вагонов на погрузке"],
    ["wagons_in_transit",   "Нет",   "Число — вагонов в пути"],
    ["wagons_in_unloading", "Нет",   "Число — вагонов выгружено (для прогресс-бара)"],
    ["norm_unload_days",    "Нет",   "Число суток — норма выгрузки"],
    ["fact_unload_days",    "Нет",   "Число суток — факт выгрузки"],
    ["wagon_ids",           "Нет",   "Номера вагонов через ; — НКЛ-2025-08451;НКЛ-2025-08452"],
    ["overall_status",      "Нет",   "OK — норма, WARN — предупреждение, BAD или ALERT — критично"],
    ["dislocation_updated_at","Нет", "Дата и время обновления, напр. 12.11.2025 14:30"],
    ["alert_text",          "Нет",   "Текст баннера-предупреждения (при WARN/BAD/ALERT)"],
    [""],
    ["Как опубликовать CSV для дашборда:"],
    ["1. Файл → Поделиться → Опубликовать в интернете"],
    ["2. Выбрать лист API_TRAINS → формат CSV → Опубликовать"],
    ["3. Скопировать URL и вставить в CSV_URL в index Sisecam.html"],
    ["   (или добавить &gid=XXXXXXX к уже имеющемуся URL)"],
  ];

  infoSheet.getRange(1, 1, instructions.length, 3).setValues(
    instructions.map(r => r.length < 3 ? [...r, ...Array(3 - r.length).fill("")] : r)
  );
  infoSheet.getRange(1,1).setFontWeight("bold").setFontSize(13);
  infoSheet.getRange(5,1,1,3).setFontWeight("bold").setBackground("#dbeafe");
  infoSheet.autoResizeColumns(1, 3);

  // Активировать лист API_TRAINS
  ss.setActiveSheet(sheet);

  Browser.msgBox("✅ Готово!", "Лист API_TRAINS создан с заголовками и примерами.\nОткройте лист '📋 Инструкция' для справки.", Browser.Buttons.OK);
}
