# Установка и запуск

## Быстрый старт

```bash
# 1. Клонировать репозиторий
git clone https://github.com/AEvstratov/teams-chat-exporter.git
cd teams-chat-exporter

# 2. Установить зависимости
pip install -r requirements.txt

# 3. Запустить
python TEAMS_explorer.py
```

## Требования

- **Python 3.9+** — [скачать](https://python.org)
- **Microsoft Edge** (рекомендуется) или Google Chrome
- Корпоративный аккаунт Microsoft с доступом к Teams

## Зависимости

| Пакет | Назначение |
|-------|-----------|
| `requests` | HTTP запросы к Graph API |
| `websocket-client` | CDP подключение к браузеру |

## Сборка в .exe (опционально)

```bash
pip install nuitka
nuitka --onefile --enable-plugin=tk-inter --assume-yes-for-downloads TEAMS_explorer.py
```
