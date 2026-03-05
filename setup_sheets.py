"""
Создаёт листы МАРШРУТЫ, РЕЙСЫ, API_TRIPS в Google Sheets
и заполняет демо-данными согласно setup_SISECAM_DISPATCHER.gs
"""

import json
from google.oauth2 import service_account
from googleapiclient.discovery import build

SPREADSHEET_ID = '13Xrmh-cfWFoR3-9TGFq2OoYZGmKxDhUhfqXWXC_SMKM'
SA_FILE = '/home/user/veles-dashboard/service_account.json'
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

creds = service_account.Credentials.from_service_account_file(SA_FILE, scopes=SCOPES)
service = build('sheets', 'v4', credentials=creds)
ss = service.spreadsheets()

# ── Получить существующие листы ───────────────────────────────────────────────
meta = ss.get(spreadsheetId=SPREADSHEET_ID).execute()
existing = {s['properties']['title']: s['properties']['sheetId'] for s in meta['sheets']}
print("Существующие листы:", list(existing.keys()))

# ── Создать отсутствующие листы ───────────────────────────────────────────────
needed = ['МАРШРУТЫ', 'РЕЙСЫ', 'API_TRIPS']
requests = []
for name in needed:
    if name not in existing:
        requests.append({'addSheet': {'properties': {'title': name}}})

if requests:
    resp = ss.batchUpdate(spreadsheetId=SPREADSHEET_ID, body={'requests': requests}).execute()
    print("Созданы листы:", [r['addSheet']['properties']['title']
                              for r in resp['replies'] if 'addSheet' in r])
    # Обновить карту ID
    meta = ss.get(spreadsheetId=SPREADSHEET_ID).execute()
    existing = {s['properties']['title']: s['properties']['sheetId'] for s in meta['sheets']}

# ── Вспомогательные функции ───────────────────────────────────────────────────

def color(hex_str):
    h = hex_str.lstrip('#')
    r, g, b = int(h[0:2],16)/255, int(h[2:4],16)/255, int(h[4:6],16)/255
    return {'red': r, 'green': g, 'blue': b}

def header_fmt(hex_color):
    return {
        'backgroundColor': color(hex_color),
        'textFormat': {'bold': True},
        'borders': {k: {'style': 'SOLID', 'color': color('999999')}
                    for k in ['top','bottom','left','right']},
    }

def cell_fmt(hex_color=None):
    fmt = {'borders': {k: {'style': 'SOLID', 'color': color('dddddd')}
                        for k in ['top','bottom','left','right']}}
    if hex_color:
        fmt['backgroundColor'] = color(hex_color)
    return fmt

def repeat_cell(sheet_id, start_row, end_row, start_col, end_col, fmt):
    return {
        'repeatCell': {
            'range': {
                'sheetId': sheet_id,
                'startRowIndex': start_row,
                'endRowIndex': end_row,
                'startColumnIndex': start_col,
                'endColumnIndex': end_col,
            },
            'cell': {'userEnteredFormat': fmt},
            'fields': 'userEnteredFormat(backgroundColor,textFormat,borders)',
        }
    }

def freeze(sheet_id, rows=1):
    return {
        'updateSheetProperties': {
            'properties': {'sheetId': sheet_id, 'gridProperties': {'frozenRowCount': rows}},
            'fields': 'gridProperties.frozenRowCount',
        }
    }

def auto_resize(sheet_id, col_count):
    return {
        'autoResizeDimensions': {
            'dimensions': {'sheetId': sheet_id, 'dimension': 'COLUMNS',
                           'startIndex': 0, 'endIndex': col_count}
        }
    }

def write_values(sheet_name, range_a1, values):
    ss.values().update(
        spreadsheetId=SPREADSHEET_ID,
        range=f"'{sheet_name}'!{range_a1}",
        valueInputOption='USER_ENTERED',
        body={'values': values}
    ).execute()

def clear_sheet(sheet_name):
    ss.values().clear(
        spreadsheetId=SPREADSHEET_ID,
        range=f"'{sheet_name}'!A1:Z2000",
        body={}
    ).execute()

# ══════════════════════════════════════════════════════════════════════════════
# МАРШРУТЫ
# ══════════════════════════════════════════════════════════════════════════════
SID_R = existing['МАРШРУТЫ']
clear_sheet('МАРШРУТЫ')

routes_headers = [
    'Код маршрута','Маршрут',
    'Станция отправления','Станция назначения',
    'Время отправления (план)','Дней в пути',
    'Время прибытия (план)','Расстояние, км'
]
routes_data = [
    ['RSH-01','Красный Гуляй → Шакша','ст. Красный Гуляй','ст. Шакша','08:00',8,'16:00',1850],
    ['RSH-02','Ульяновск → Шакша','ст. Ульяновск-Центральный','ст. Шакша','14:00',7,'10:00',1430],
    ['RSH-03','Самара → Шакша','ст. Самара','ст. Шакша','10:00',6,'14:00',1200],
    ['RSH-04','Тольятти → Шакша','ст. Тольятти','ст. Шакша','12:00',7,'08:00',1280],
    ['RSH-05','Пенза → Шакша','ст. Пенза-1','ст. Шакша','09:00',5,'11:00',980],
    ['RSH-06','Уфа → Шакша','ст. Уфа','ст. Шакша','18:00',4,'08:00',640],
]

write_values('МАРШРУТЫ', 'A1', [routes_headers] + routes_data)

fmt_requests = [
    freeze(SID_R),
    repeat_cell(SID_R, 0, 1, 0, len(routes_headers), header_fmt('cfe2ff')),
    repeat_cell(SID_R, 1, len(routes_data)+1, 0, len(routes_headers), cell_fmt()),
    auto_resize(SID_R, len(routes_headers)),
]
ss.batchUpdate(spreadsheetId=SPREADSHEET_ID, body={'requests': fmt_requests}).execute()
print("✅ МАРШРУТЫ заполнен")

# ══════════════════════════════════════════════════════════════════════════════
# РЕЙСЫ
# ══════════════════════════════════════════════════════════════════════════════
SID_T = existing['РЕЙСЫ']
clear_sheet('РЕЙСЫ')

trips_headers = [
    'Код рейса',            # A  ★ авто
    'Маршрут (код)',        # B  диспетчер
    'Маршрут',              # C  ★ авто
    'Дата отправления',     # D
    'Время отправления',    # E  ★ авто
    'Дата прибытия (план)', # F  ★ авто
    'Время прибытия (план)',# G  ★ авто
    'Вагонов всего',        # H
    'Состояние',            # I
    'Статус погрузки',      # J
    'Погружено вагонов',    # K
    'Текущая станция',      # L  ★ n8n
    'Осталось, км',         # M  ★ n8n
    'Прогноз прибытия',     # N  ★ n8n
    'Факт отправления',     # O
    'Факт прибытия',        # P
    'Дислокация обновлена', # Q  ★ n8n
    'Примечание',           # R
]

# Демо-рейсы (18 колонок)
trips_demo = [
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
]

write_values('РЕЙСЫ', 'A1', [trips_headers] + trips_demo)

# Автоколонки (0-based): A=0, C=2, E=4, F=5, G=6, L=11, M=12, N=13, Q=16
auto_cols = [0, 2, 4, 5, 6, 11, 12, 13, 16]

fmt_r2 = [
    freeze(SID_T),
    # Все заголовки — синий фон
    repeat_cell(SID_T, 0, 1, 0, len(trips_headers), header_fmt('e8f0fe')),
    # Авто-колонки — светло-голубой заголовок
    *[repeat_cell(SID_T, 0, 1, c, c+1, header_fmt('d0e8ff')) for c in auto_cols],
    # Данные
    repeat_cell(SID_T, 1, 10, 0, len(trips_headers), cell_fmt()),
    auto_resize(SID_T, len(trips_headers)),
]
ss.batchUpdate(spreadsheetId=SPREADSHEET_ID, body={'requests': fmt_r2}).execute()
print("✅ РЕЙСЫ заполнен")

# ══════════════════════════════════════════════════════════════════════════════
# API_TRIPS
# ══════════════════════════════════════════════════════════════════════════════
SID_A = existing['API_TRIPS']
clear_sheet('API_TRIPS')

api_headers = [
    'trip_id','route_name','current_state','loading_status',
    'wagons_total','wagons_loaded',
    'departure_plan','arrival_plan',
    'actual_departure','actual_arrival',
    'current_station','km_remaining','eta',
    'dislocation_updated_at','alert_text','overall_status'
]

# overall_status: OK / WARN / BAD
api_data = [
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
]

write_values('API_TRIPS', 'A1', [api_headers] + api_data)

fmt_r3 = [
    freeze(SID_A),
    repeat_cell(SID_A, 0, 1, 0, len(api_headers), header_fmt('fce8b2')),
    repeat_cell(SID_A, 1, 10, 0, len(api_headers), cell_fmt()),
    auto_resize(SID_A, len(api_headers)),
]
ss.batchUpdate(spreadsheetId=SPREADSHEET_ID, body={'requests': fmt_r3}).execute()
print("✅ API_TRIPS заполнен")

print("\n🎉 Готово! Все листы созданы и заполнены.")
