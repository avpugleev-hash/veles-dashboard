# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Архитектура

**Стек:** чистый HTML + Vanilla JS + CSS, без сборки, без npm, без фреймворков. Каждая страница — один самодостаточный `.html` файл (~800–4200 строк).

**Страницы:**
- `index.html` — главный дашборд (погрузка, ДС, ЕЛС, задолженность, бюджет, ремонты)
- `index_managers.html` — версия дашборда для менеджеров
- `debt.html` — детализация дебиторской задолженности с комментариями и генерацией писем
- `creditors.html` — кредиторская задолженность (задолженность перед поставщиками)

## Источники данных

Два типа источников данных, всегда `fetch(..., {cache:'no-store'})`:

**1. n8n webhook-агрегатор** (`index.html`, `index_managers.html`):
```js
const DASHBOARD_API_URL = 'https://alekseipugleev.app.n8n.cloud/webhook/dashboard-data';
```
Возвращает `{ ok: true, csv: { raw, payments, clients, ap_suppliers, loading_journal, cash, budget_prod, comm_kpi, comm_nat, comm_clients, repairs } }`. Все основные данные дашборда идут через этот один запрос.

**2. Google Sheets напрямую** (браузер → Sheets):
- `index.html` / `index_managers.html`: ЕЛС — `gid=985785018` (лист "ЕЛС ЭТРАН", строки по субсчётам)
- `debt.html`: задолженность — spreadsheet `1D9-uViFll7h3SUT01lvlHKTF0VnQCx46h8569zSpaCQ`, `gid=1164788937`
- `creditors.html`: кредиторы — spreadsheet `1LWplKTlBOJB-5TycHmsReHX-jLDik1TnerdByVCxiUo`

## Авторизация

Общий localStorage-ключ `veles_dashboard_auth_v1`, TTL 8 часов. Валидация через:
```js
const LOGIN_API_URL = 'https://alekseipugleev.app.n8n.cloud/webhook/dashboard-login';
```
Флаг `isAuthenticated` проверяется перед каждым `refreshData()`.

## n8n webhooks

| Webhook | Назначение |
|---|---|
| `webhook/dashboard-data` | Главный агрегатор всех CSV-данных |
| `webhook/dashboard-login` | Авторизация |
| `webhook/get-comments` | Комментарии менеджеров по клиентам |
| `webhook/save-comment` | Сохранение комментария |
| `webhook/get-comment-history` | История комментариев по клиенту |
| `webhook/generate-email` | Генерация письма через Groq LLM |

Воркфлоу в этом же репозитории: `n8n_daily_report.json`, `n8n_els_morning_report.json`, `n8n_generate_email.json`, `n8n_get_comment_history.json`, `n8n_save_comment.json`.

## ЕЛС — структура листа (gid=985785018)

Колонки: `Дата запроса, Дата данных, Организация, ОКПО, Код ЕЛС, ЦФТО, Субсчёт, Нач.сальдо, Дебет, Кредит, Сальдо`

Парсинг: найти последний срез по `Дата запроса` (формат `DD.MM.YYYY HH:MM`), затем:
- субсчёт `"облагаются НДС"` → ЕЛС 22%
- субсчёт `"НДС 0%"` → ЕЛС 0%

## Разработка и деплой

Нет системы сборки. Редактировать `.html` файлы напрямую, проверять в браузере, пушить в `main`.

```bash
# Локальная проверка
python3 -m http.server 8080
# Открыть: http://localhost:8080/index.html
```

Репозиторий подключён к GitHub Pages или используется напрямую через raw/GitHub URL.

## Важные паттерны

- Числа из Google Sheets приходят в **русской локали** (пробел = тысячи, запятая = десятичная) → парсить через `.replace(/[\s\u00a0]/g,'').replace(',','.')`
- Даты из старых листов в формате `M/D/YYYY` (нужна конвертация), из листа ЕЛС ЭТРАН — `DD.MM.YYYY` (готово к отображению)
- `index_managers.html` — устаревшая копия `index.html`, может отставать по изменениям (например, ЕЛС там всё ещё читает `gid=1639673932`)
