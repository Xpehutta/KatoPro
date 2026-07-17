# KatoPro

Сервис автоматической генерации ежемесячных отчётных книг Excel для торговых точек.

На основе заполненного файла за предыдущий месяц создаётся чистый лист на следующий месяц: даты, подсветка выходных и праздников РФ, формулы, ссылки на предыдущий месяц и защита ячеек.

## Возможности

- генерация листа Excel на выбранный год и месяц;
- произвольное число торговых точек (добавление через веб-интерфейс или конфиг);
- пакетная загрузка исходных `.xlsx` через UI (с подтверждением замены при совпадении имени);
- просмотр, скачивание и удаление файлов из `data/` и `generated/` (удаление — в `trash/`);
- проверка наличия файла предыдущего периода в `data/` с рекомендацией и корректировкой периода по точке;
- подтверждение перед пересозданием уже существующего отчёта в `generated/`;
- получение праздников РФ из интернета (с резервным списком при недоступности API);
- REST API и русскоязычный веб-интерфейс;
- запуск в Docker / Docker Compose.

## Быстрый старт

Кратко:

```bash
git clone https://github.com/Xpehutta/KatoPro.git
cd KatoPro
docker compose up -d --build
```

Откройте в браузере: [http://localhost:8000](http://localhost:8000)

Полная пошаговая инструкция для новичка (установка Docker и Git, `git clone` / `git pull`, подготовка Excel, генерация, типовые ошибки):

**→ [USER_GUIDE.md — Быстрый старт](USER_GUIDE.md#2-быстрый-старт-для-новичка)**

## Локальный запуск без Docker

Требования: Python 3.10+.

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt

uvicorn src.main:app --host 0.0.0.0 --port 8000
```

## Структура проекта

```
KatoPro/
├── AGENT.md                 # техническое задание
├── USER_GUIDE.md            # руководство пользователя
├── README.md
├── config.yaml              # базовые настройки
├── docker-compose.yml
├── Dockerfile
├── requirements.txt
├── data/                    # исходные Excel и points.yaml (локально, не в Git)
├── generated/               # готовые отчёты (локально, не в Git)
├── trash/                   # корзина удалённых файлов
├── logs/                    # логи
└── src/
    ├── main.py              # FastAPI + веб-интерфейс
    ├── generator.py         # генерация Excel
    ├── holidays.py          # праздники РФ
    ├── config_loader.py     # конфигурация и точки
    ├── models.py
    ├── utils.py
    ├── templates/           # HTML
    └── static/              # CSS
```

Содержимое `data/` и `generated/` в репозиторий не попадает — только пустые каталоги-заглушки (`.gitkeep`). Рабочие Excel и `points.yaml` остаются у вас на диске.

## Конфигурация

Основные параметры — в `config.yaml`:

| Параметр | Описание |
|----------|----------|
| `points` | Начальный список точек (далее обычно ведётся в `data/points.yaml`) |
| `output_dir` | Каталог для сгенерированных файлов |
| `protection_password` | Пароль защиты листа (`null` / `""` — без пароля) |
| `holidays_api` | URL API производственного календаря и запасной источник |

Список точек, добавленных через веб-интерфейс, сохраняется в `data/points.yaml` и имеет приоритет над `points` из `config.yaml`.

## API (кратко)

| Метод | Путь | Описание |
|-------|------|----------|
| `GET` | `/` | Веб-интерфейс |
| `GET` | `/health` | Проверка работоспособности |
| `POST` | `/generate` | Генерация отчётов |
| `GET` | `/api/points` | Список точек |
| `GET` | `/api/storage` | Список файлов в `data/` и `generated/` |
| `DELETE` | `/api/storage/{kind}/{filename}` | Переместить файл в `trash/` (`kind`: `data` или `generated`) |
| `POST` | `/api/storage/clear-session` | Удалить файлы: параметры `clear_data` / `clear_generated` (Excel → `trash/`; при `data/` также очищаются точки) |
| `POST` | `/api/storage/generated/{filename}/to-data` | Переместить файл из `generated/` в `data/` (`replace=true` при конфликте имени) |
| `GET` | `/download/{kind}/{filename}` | Скачать файл (`kind`: `data` или `generated`) |
| `GET` | `/api/points/scan-data` | Найти в `data/` точки, которых ещё нет в списке |
| `POST` | `/api/points/sync-data` | Добавить найденные новые точки из `data/` |
| `POST` | `/api/points/upload` | Добавить точки (один или несколько файлов; `replace=true` — перезаписать существующие имена) |
| `POST` | `/api/points/import-data` | Зарегистрировать незанятые Excel из `data/` |

Пример генерации:

```bash
curl -X POST http://localhost:8000/generate \
  -H "Content-Type: application/json" \
  -d '{"year": 2026, "month": 5, "points": ["Смола"]}'
```

Если `points` не указан — обрабатываются все точки из конфигурации.

Ответы `409` у `/generate`:

| `detail.code` | Когда | Что делать |
|---------------|--------|------------|
| `missing_previous` | В `data/` нет файла предыдущего месяца (`МесОтчYYYYMM_Точка.xlsx`) | Передать `point_overrides` с рекомендованным периодом или `skip_previous_check: true` |
| `already_exists` | Файл уже есть в `generated/` | Передать `"force": true` для перезаписи |

Пример с индивидуальным периодом для точки:

```bash
curl -X POST http://localhost:8000/generate \
  -H "Content-Type: application/json" \
  -d '{
    "year": 2026,
    "month": 6,
    "points": ["Смола"],
    "point_overrides": [{"name": "Смола", "year": 2026, "month": 5}]
  }'
```

## Тома Docker

| Хост | Контейнер | Назначение |
|------|-----------|------------|
| `./data` | `/app/data` | Исходные Excel и `points.yaml` |
| `./generated` | `/app/generated` | Результаты генерации |
| `./trash` | `/app/trash` | Корзина удалённых файлов |
| `./config.yaml` | `/app/config.yaml` | Настройки |
| `./logs` | `/app/logs` | Логи |

## Технологии

- Python 3.10+
- FastAPI, Uvicorn, Jinja2
- openpyxl
- httpx, PyYAML, Pydantic, Loguru
- Docker

## Документация

- [USER_GUIDE.md](USER_GUIDE.md) — руководство пользователя
- [USER_GUIDE.md — Быстрый старт](USER_GUIDE.md#2-быстрый-старт-для-новичка) — установка, clone/pull, первый запуск
- [AGENT.md](AGENT.md) — исходное техническое задание

## Лицензия

Внутренний проект. Условия использования определяются владельцем репозитория.
