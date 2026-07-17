# KatoPro

Сервис автоматической генерации ежемесячных отчётных книг Excel для торговых точек.

На основе заполненного файла за предыдущий месяц создаётся чистый лист на следующий месяц: даты, подсветка выходных и праздников РФ, формулы, ссылки на предыдущий месяц и защита ячеек.

## Возможности

- генерация листа Excel на выбранный год и месяц;
- произвольное число торговых точек (добавление через веб-интерфейс или конфиг);
- загрузка исходных `.xlsx` через UI;
- получение праздников РФ из интернета (с резервным списком при недоступности API);
- REST API и русскоязычный веб-интерфейс;
- запуск в Docker / Docker Compose.

## Быстрый старт (Docker)

Требования: Docker и Docker Compose.

```bash
# 1. Положите исходные Excel-файлы точек в папку data/
#    например: data/МесОтч202603_Смола.xlsx

# 2. Запустите сервис
docker compose up -d --build

# 3. Откройте веб-интерфейс
open http://localhost:8000
```

Документация API (Swagger): http://localhost:8000/docs

Остановка:

```bash
docker compose down
```

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
├── data/                    # исходные Excel и points.yaml
├── generated/               # готовые отчёты
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
| `POST` | `/api/points/upload` | Добавить точку (загрузка файла) |
| `POST` | `/api/points/manual` | Добавить точку (файл уже в `data/`) |
| `DELETE` | `/api/points/{name}` | Удалить точку из списка |
| `GET` | `/download/{filename}` | Скачать файл из `generated/` |

Пример генерации:

```bash
curl -X POST http://localhost:8000/generate \
  -H "Content-Type: application/json" \
  -d '{"year": 2026, "month": 5, "points": ["Смола"]}'
```

Если `points` не указан — обрабатываются все точки из конфигурации.

## Тома Docker

| Хост | Контейнер | Назначение |
|------|-----------|------------|
| `./data` | `/app/data` | Исходные Excel и `points.yaml` |
| `./generated` | `/app/generated` | Результаты генерации |
| `./config.yaml` | `/app/config.yaml` | Настройки |
| `./logs` | `/app/logs` | Логи |

## Технологии

- Python 3.10+
- FastAPI, Uvicorn, Jinja2
- openpyxl
- httpx, PyYAML, Pydantic, Loguru
- Docker

## Документация

- [USER_GUIDE.md](USER_GUIDE.md) — как пользоваться сервисом
- [AGENT.md](AGENT.md) — исходное техническое задание

## Лицензия

Внутренний проект. Условия использования определяются владельцем репозитория.
