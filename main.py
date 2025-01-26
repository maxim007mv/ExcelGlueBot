import os
import logging
import sqlite3
import shutil
import time
from datetime import datetime
import pandas as pd
from io import BytesIO
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import (
    Application,
    MessageHandler,
    filters,
    CallbackContext,
    CommandHandler,
    CallbackQueryHandler
)
from dotenv import load_dotenv
load_dotenv()  # Загружает переменные из .env

# Настройка логирования
logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    level=logging.INFO,
    filename='bot.log',
    filemode='a'
)
logger = logging.getLogger(__name__)

# Конфигурация
TOKEN = os.getenv("TELEGRAM_BOT_TOKEN")
TEMP_DIR = "temp_files"
RESULT_FILE = "result.xlsx"
DB_NAME = "files_db.sqlite"
MAX_FILES_PER_USER = 10

# Поддерживаемые форматы Excel
EXCEL_EXTENSIONS = ['xlsx', 'xls', 'xlsm', 'xlsb', 'odf', 'ods']
SUPPORTED_MIME_TYPES = [
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    'application/vnd.ms-excel',
    'application/vnd.ms-excel.sheet.macroEnabled.12',
    'application/vnd.oasis.opendocument.spreadsheet'
]


def handle_remove_readonly(func, path, exc):
    """Обработчик ошибок доступа для Windows"""
    import stat
    os.chmod(path, stat.S_IRWXU | stat.S_IRWXG | stat.S_IRWXO)  # 0777
    func(path)


# Инициализация БД
def init_db():
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    cursor.execute('''CREATE TABLE IF NOT EXISTS users
                     (id INTEGER PRIMARY KEY,
                      username TEXT,
                      created_at TIMESTAMP)''')
    cursor.execute('''CREATE TABLE IF NOT EXISTS files
                     (id INTEGER PRIMARY KEY AUTOINCREMENT,
                      user_id INTEGER,
                      filename TEXT,
                      uploaded_at TIMESTAMP,
                      FOREIGN KEY(user_id) REFERENCES users(id))''')
    conn.commit()
    conn.close()


init_db()


# Функции работы с БД
def log_user(user_id, username):
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    cursor.execute('''INSERT OR IGNORE INTO users (id, username, created_at)
                     VALUES (?, ?, ?)''',
                   (user_id, username, datetime.now()))
    conn.commit()
    conn.close()


def log_file(user_id, filename):
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    cursor.execute('''INSERT INTO files (user_id, filename, uploaded_at)
                     VALUES (?, ?, ?)''',
                   (user_id, filename, datetime.now()))
    conn.commit()
    conn.close()


def get_file_count(user_id):
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    cursor.execute('''SELECT COUNT(*) FROM files WHERE user_id = ?''', (user_id,))
    count = cursor.fetchone()[0]
    conn.close()
    return count


def delete_user_files(user_id):
    """Удаляет все файлы пользователя из БД"""
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    cursor.execute('''DELETE FROM files WHERE user_id = ?''', (user_id,))
    conn.commit()
    conn.close()
    logger.info(f"Удалены файлы пользователя {user_id} из БД")


# Обработка Excel
def process_excel(file_path: str) -> pd.DataFrame:
    """Обработка Excel-файла с поддержкой разных форматов"""
    try:
        ext = os.path.splitext(file_path)[1].lower()
        engine_map = {
            '.xlsx': 'openpyxl',
            '.xls': 'xlrd',
            '.xlsm': 'openpyxl',
            '.xlsb': 'pyxlsb',
            '.ods': 'odf',
        }
        df = pd.read_excel(file_path, engine=engine_map.get(ext, None))
        logger.info(f"Прочитан файл: {file_path}")

        # Поиск столбцов
        name_col = next((col for col in df.columns
                         if any(kw in col.lower() for kw in ["наименование", "название", "товар"])), None)
        price_col = next((col for col in df.columns
                          if any(kw in col.lower() for kw in ["цена", "стоимость"])), None)

        if not name_col or not price_col:
            raise ValueError("Отсутствуют обязательные столбцы (Наименование и Цена)")

        # Обработка данных
        df = df.rename(columns={name_col: "Наименование", price_col: "Цена"})
        df["Наименование"] = df["Наименование"].astype(str).str.strip()

        if "Количество" not in df.columns:
            df["Количество"] = 0

        return df[["Наименование", "Количество", "Цена"]]

    except Exception as e:
        logger.error(f"Ошибка обработки {file_path}: {str(e)}", exc_info=True)
        raise


# Обработчики команд
async def handle_document(update: Update, context: CallbackContext) -> None:
    """Обработка документов"""
    try:
        user = update.message.from_user
        document = update.message.document
        log_user(user.id, user.username)

        # Проверка расширения файла
        file_ext = os.path.splitext(document.file_name)[1].lower()[1:]
        if file_ext not in EXCEL_EXTENSIONS:
            await update.message.reply_text("❌ Поддерживаются только Excel-файлы!")
            return

        # Проверка лимита файлов
        file_count = get_file_count(user.id)
        if file_count >= MAX_FILES_PER_USER:
            await update.message.reply_text(f"❌ Лимит файлов ({MAX_FILES_PER_USER}) исчерпан!")
            return

        # Скачивание файла
        file = await context.bot.get_file(document)
        user_dir = os.path.join(TEMP_DIR, str(user.id))
        file_path = os.path.join(user_dir, document.file_name)

        os.makedirs(user_dir, exist_ok=True, mode=0o777)
        await file.download_to_drive(file_path)

        log_file(user.id, document.file_name)
        logger.info(f"Файл сохранен: {file_path}")

        # Обработка файла
        df = process_excel(file_path)
        context.user_data.setdefault('files', {})[file_path] = df
        await update.message.reply_text(
            f"✅ Файл обработан! ({file_count + 1}/{MAX_FILES_PER_USER}) Отправьте следующий или /report")

    except Exception as e:
        await update.message.reply_text(f"❌ Ошибка: {str(e)}")
        logger.error(f"Ошибка обработки документа: {str(e)}", exc_info=True)


async def send_report(update: Update, context: CallbackContext) -> None:
    """Генерация основного отчета"""
    try:
        user = update.message.from_user
        if not context.user_data.get('files'):
            await update.message.reply_text("⚠ Сначала отправьте файлы!")
            return

        # Создание основного отчета
        with pd.ExcelWriter(RESULT_FILE, engine='openpyxl') as writer:
            start_col = 0
            for file_path, df in context.user_data['files'].items():
                pd.DataFrame([[os.path.basename(file_path)]]).to_excel(
                    writer,
                    startrow=0,
                    startcol=start_col,
                    index=False,
                    header=False
                )
                df.to_excel(
                    writer,
                    startrow=3,
                    startcol=start_col,
                    index=False,
                    header=False
                )
                start_col += 4

        # Создаем кнопку для подробного отчета
        keyboard = [[InlineKeyboardButton("📈 Подробный анализ", callback_data='detailed_report')]]
        reply_markup = InlineKeyboardMarkup(keyboard)

        # Отправка файла с кнопкой
        await update.message.reply_document(
            document=open(RESULT_FILE, 'rb'),
            caption="📊 Ваш отчет готов!\nДля нового сравнения нажмите /newreport",
            filename="Сравнение_цен.xlsx",
            reply_markup=reply_markup
        )

        # Сохраняем сырые данные для возможного подробного отчета
        context.user_data['raw_data'] = {
            file_path: df.copy()
            for file_path, df in context.user_data['files'].items()
        }

        # Очистка временных файлов
        for file_path in context.user_data['files']:
            if os.path.exists(file_path):
                try:
                    os.remove(file_path)
                except Exception as e:
                    logger.error(f"Ошибка удаления файла {file_path}: {str(e)}")

        if os.path.exists(RESULT_FILE):
            os.remove(RESULT_FILE)

        del context.user_data['files']
        delete_user_files(user.id)
        logger.info(f"Отчет отправлен для пользователя {user.id}")

    except Exception as e:
        await update.message.reply_text(f"🔥 Ошибка генерации отчета: {str(e)}")
        logger.error(f"Ошибка отчета: {str(e)}", exc_info=True)


async def detailed_report_callback(update: Update, context: CallbackContext) -> None:
    """Обработка нажатия на кнопку подробного отчета"""
    query = update.callback_query
    await query.answer()

    try:
        user = query.from_user
        if 'raw_data' not in context.user_data:
            await query.message.reply_text("⚠ Данные для анализа больше недоступны!")
            return

        # Собираем все данные в один DataFrame
        all_data = []
        for file_path, df in context.user_data['raw_data'].items():
            temp_df = df.copy()
            temp_df['Источник'] = os.path.basename(file_path)
            all_data.append(temp_df)

        full_df = pd.concat(all_data, ignore_index=True)

        # Анализ данных
        analysis = full_df.groupby('Наименование').agg({
            'Цена': ['min', 'max', 'mean', 'count'],
            'Источник': lambda x: ', '.join(x)
        }).reset_index()

        analysis.columns = [
            'Наименование',
            'Минимальная цена',
            'Максимальная цена',
            'Средняя цена',
            'Количество предложений',
            'Источники'
        ]

        # Добавляем информацию о продавцах с мин/макс ценами
        min_sources = full_df.loc[full_df.groupby('Наименование')['Цена'].idxmin()][['Наименование', 'Источник']]
        max_sources = full_df.loc[full_df.groupby('Наименование')['Цена'].idxmax()][['Наименование', 'Источник']]

        analysis = analysis.merge(
            min_sources.rename(columns={'Источник': 'Продавец с мин. ценой'}),
            on='Наименование'
        ).merge(
            max_sources.rename(columns={'Источник': 'Продавец с макс. ценой'}),
            on='Наименование'
        )

        # Создаем Excel-файл с анализом
        report_file = "detailed_analysis.xlsx"
        with pd.ExcelWriter(report_file, engine='openpyxl') as writer:
            analysis.to_excel(writer, sheet_name='Сводка', index=False)
            full_df.to_excel(writer, sheet_name='Все данные', index=False)

            stats = pd.DataFrame({
                'Метрика': [
                    'Всего товаров',
                    'Товары с одним предложением',
                    'Средний разброс цен',
                    'Максимальный разброс цен'
                ],
                'Значение': [
                    len(analysis),
                    sum(analysis['Количество предложений'] == 1),
                    (analysis['Максимальная цена'] - analysis['Минимальная цена']).mean(),
                    (analysis['Максимальная цена'] - analysis['Минимальная цена']).max()
                ]
            })
            stats.to_excel(writer, sheet_name='Статистика', index=False)

        # Отправляем результаты
        await context.bot.send_document(
            chat_id=query.message.chat_id,
            document=open(report_file, 'rb'),
            caption="📊 Подробный анализ цен\n"
                    "Содержит:\n"
                    "1. Сводку по товарам\n"
                    "2. Все исходные данные\n"
                    "3. Общую статистику",
            filename="Подробный_анализ_цен.xlsx"
        )

        # Очистка
        os.remove(report_file)
        del context.user_data['raw_data']

    except Exception as e:
        await query.message.reply_text(f"🔥 Ошибка генерации подробного отчета: {str(e)}")
        logger.error(f"Ошибка detailed_report: {str(e)}", exc_info=True)


async def new_report(update: Update, context: CallbackContext) -> None:
    """Сброс текущей сессии"""
    try:
        user = update.message.from_user
        user_dir = os.path.join(TEMP_DIR, str(user.id))

        if os.path.exists(user_dir):
            try:
                time.sleep(1)
                shutil.rmtree(user_dir, onerror=handle_remove_readonly)
                logger.info(f"Директория {user_dir} удалена")
            except Exception as e:
                logger.error(f"Ошибка удаления директории: {str(e)}")
                raise

        context.user_data.pop('files', None)
        delete_user_files(user.id)

        await update.message.reply_text(
            "🆕 Новая сессия начата!\n"
            f"Лимит файлов: {MAX_FILES_PER_USER}\n"
            "Отправьте первый файл"
        )

    except Exception as e:
        await update.message.reply_text(f"❌ Ошибка сброса: {str(e)}")
        logger.error(f"Ошибка new_report: {str(e)}", exc_info=True)


async def start(update: Update, context: CallbackContext) -> None:
    """Обработчик команды /start"""
    user = update.message.from_user
    file_count = get_file_count(user.id)

    message_text = (
            "📎 Отправьте Excel-файлы для сравнения цен\n"
            f"Текущий прогресс: {file_count}/{MAX_FILES_PER_USER}\n"
            "Поддерживаемые форматы: " + ", ".join(EXCEL_EXTENSIONS) + "\n\n"
                                                                       "Команды:\n"
                                                                       "/newreport - начать заново\n"
                                                                       "/report - сформировать отчет"
    )

    if file_count > 0:
        message_text += "\n\n⚠ Обнаружены предыдущие файлы. Для очистки используйте /newreport"

    await update.message.reply_text(message_text)


def main():
    application = Application.builder().token(TOKEN).build()

    # Фильтры для документов
    ext_filters = filters.Document.FileExtension(EXCEL_EXTENSIONS[0])
    for ext in EXCEL_EXTENSIONS[1:]:
        ext_filters |= filters.Document.FileExtension(ext)
    doc_filter = ext_filters | filters.Document.MimeType(SUPPORTED_MIME_TYPES)

    application.add_handler(CommandHandler("start", start))
    application.add_handler(CommandHandler("report", send_report))
    application.add_handler(CommandHandler("newreport", new_report))
    application.add_handler(MessageHandler(doc_filter, handle_document))
    application.add_handler(CallbackQueryHandler(detailed_report_callback, pattern='^detailed_report$'))

    logger.info("Бот запущен")
    application.run_polling()


if __name__ == '__main__':
    os.makedirs(TEMP_DIR, exist_ok=True, mode=0o777)
    main()