import telebot
import pandas as pd
from io import BytesIO

# Ваш токен для бота
TOKEN = '7193167480:AAHW-uClzruZMfGwiV-za3FZSoG6UMG63Ko'
bot = telebot.TeleBot(TOKEN)

# Функция для нормализации столбцов
def normalize_columns(df):
    column_mapping = {
        'Тема урока': 'Тема урока'
    }
    # Переименовываем столбцы по маппингу
    df.rename(columns=column_mapping, inplace=True)
    return df

# Проверка корректности темы урока
def check_lesson_topic(file):
    try:
        df = pd.read_excel(file, sheet_name=0)
        df = normalize_columns(df)

        if 'Тема урока' not in df.columns:
            return "Ошибка: В файле отсутствует столбец 'Тема урока'."

        incorrect_topics = df[~df['Тема урока'].str.startswith('Урок ', na=False)]

        messages = [
            f"Все темы уроков заполнены корректно. {row['Тема урока']}"
            for _, row in incorrect_topics.iterrows()
        ]
        return messages or ["Все темы уроков заполнены корректно."]
    except Exception as e:
        return f"Ошибка при обработке файла: {e}"

# Обработчик команды /start
@bot.message_handler(commands=['start'])
def start(message):
    bot.send_message(message.chat.id, "Привет! Отправь мне Excel.")

# Обработчик получения файлов
@bot.message_handler(content_types=['document'])
def handle_document(message):
    try:
        file_id = message.document.file_id
        file_info = bot.get_file(file_id)
        file = bot.download_file(file_info.file_path)

        # Проверяем правильность темы уроков
        result_topics = check_lesson_topic(BytesIO(file))
        send_result(message, result_topics, "Результаты проверки тем уроков")

    except Exception as e:
        bot.send_message(message.chat.id, f"Произошла ошибка: {e}")

# Функция отправки результата
def send_result(message, result, caption):
    if isinstance(result, str):
        bot.send_message(message.chat.id, result)
    else:
        # Если результат - список сообщений, то отправляем каждое по отдельности
        for res in result:
            bot.send_message(message.chat.id, res)

# Запуск бота
if __name__ == '__main__':
    bot.polling(none_stop=True)
