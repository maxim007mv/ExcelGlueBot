# 🤖 Price Comparator Bot

Telegram-бот для автоматического сравнения цен из Excel-файлов. Анализирует данные, формирует отчеты и выявляет лучшие предложения.

[![Python](https://img.shields.io/badge/Python-3.11%2B-blue?logo=python)](https://python.org)
[![Render](https://img.shields.io/badge/Deploy%20on-Render-46B3E6?logo=render)](https://render.com)

![Пример работы бота](https://via.placeholder.com/800x400.png?text=Demo+GIF+Placeholder)

## 🌟 Особенности
- **Поддержка форматов**: XLSX, XLS, XLSB, ODS, XLSM
- **Автоанализ данных**: автоматическое определение колонок с ценами и названиями
- **Умные отчеты**:
  - Сводное сравнение цен
  - Детальная аналитика (мин/макс/среднее)
  - Визуализация данных
- **Безопасность**: ограничение доступа по ID пользователя
- **Логирование**: полный трекинг операций

## 🛠 Технологии
- `python-telegram-bot 20.3` — работа с Telegram API
- `pandas 2.0.3` — обработка данных
- `openpyxl 3.1.2` — работа с Excel-файлами
- `SQLite` — хранение метаданных
- `Render` — хостинг

## 🚀 Быстрый старт

### 1. Локальная установка
```bash
# Клонировать репозиторий
git clone https://github.com/yourusername/price-comparator-bot.git
cd price-comparator-bot

# Создать виртуальное окружение
python -m venv .venv
source .venv/bin/activate  # Linux/Mac
# .venv\Scripts\activate  # Windows

# Установить зависимости
pip install -r requirements.txt

# Установка зависимостей без requirements.txt

Для установки всех необходимых библиотек выполните **одну из команд** ниже:

---

## 1. Минимальная установка (основные зависимости)
```bash
pip install --upgrade pip && pip install \
python-telegram-bot==20.3 \
pandas==2.0.3 \
numpy==1.24.3 \
openpyxl==3.1.2 \
xlrd==2.0.1 \
pyxlsb==1.0.10 \
odfpy==1.4.1 \
python-dotenv==1.0.0

## 2. Полная установка (включая все транзитивные зависимости)
pip install --upgrade pip && pip install \
python-telegram-bot==20.3 \
pandas==2.0.3 \
numpy==1.24.3 \
openpyxl==3.1.2 \
xlrd==2.0.1 \
pyxlsb==1.0.10 \
odfpy==1.4.1 \
python-dotenv==1.0.0 \
httpx==0.24.1 \
httpcore==0.17.3 \
anyio==4.8.0 \
typing-extensions==4.12.2 \
idna==3.10 \
h11==0.14.0 \
sniffio==1.3.1 \
certifi==2024.12.14 \
python-dateutil==2.9.0.post0 \
pytz==2024.2 \
tzdata==2025.1 \
defusedxml==0.7.1 \
et-xmlfile==2.0.0 \
six==1.17.0

# Создать файл .env
echo "TELEGRAM_BOT_TOKEN=ваш_токен" > .env

# Запустить бота
python main.py
