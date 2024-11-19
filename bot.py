import time
import threading
import telebot
import pandas as pd
from datetime import datetime, timedelta, date
from openpyxl import Workbook
import json
import os
import sys

# Функция для записи в файл с меткой времени
def log_print(*args, **kwargs):
    # Получаем текущую дату и время
    current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    # Формируем строку для записи в файл
    message = f"{current_time} - " + " ".join(map(str, args))

    # Записываем в файл
    with open("log.txt", "a", encoding="utf-8") as log_file:
        log_file.write(message + "\n")

    # Используем встроенную функцию print для вывода на экран
    __builtins__.print(*args, **kwargs)


# Пример использования функции
log_print("==================")
log_print("НАЧАЛО ЛОГИРОВАНИЯ")


# Загрузка данных из Excel файла
excel_file_path = r'E:\SOFT\PyProject\FKBot\FKBot\Список сотрудников.xlsx'
df = pd.read_excel(excel_file_path)

# Добавляем столбец 'Нажал на кнопку' и заполняем его значениями 'Да' по умолчанию
df['Нажал на кнопку'] = 'Да'
print(df.columns)

# Токен вашего бота
TOKEN = 'ТЕЛЕГРАМ_ТОКЕН'
log_print("Получен токен бота")
# ID чата №1 и №2
CHAT_ID_1 = ID_ЧАТА_1
CHAT_ID_2 = ID_ЧАТА_1

bot = telebot.TeleBot(TOKEN)

late_minutes_allowed = 15 # Время дедлайна
# ==================
# ИЗМЕНЕНИЯ ОТ 18.11.2024
HOLIDAYS_FILE = "holidays.json"
def load_holidays():
    if os.path.exists(HOLIDAYS_FILE):
        try:
            with open(HOLIDAYS_FILE, "r") as file:
                return set(json.load(file))
        except Exception as e:
            print(f"Ошибка загрузки выходных из файла: {e}")
            log_print(f"Ошибка загрузки выходных из файла: {e}")
    return set()

# Функция сохранения выходных в файл
def save_holidays(holidays):
    try:
        with open(HOLIDAYS_FILE, "w") as file:
            json.dump(list(holidays), file)
    except Exception as e:
        print(f"Ошибка сохранения выходных в файл: {e}")
        log_print(f"Ошибка загрузки выходных в файл: {e}")

# Список государственных выходных
holidays = load_holidays()

def add_holiday(date):
    if date not in holidays:
        holidays.add(date)
        save_holidays(holidays)  # Сохранение изменений
        print(f"Добавлен выходной день: {date}")
        log_print("Была вызвана команда /add_holidays")
        log_print(f"Добавлен выходной день: {date}")
    else:
        print(f"День {date} уже является выходным.")
        log_print(f"Добавлен выходной день: {date}")

def remove_holiday(date):
    if date in holidays:
        holidays.remove(date)
        save_holidays(holidays)  # Сохранение изменений
        print(f"Удалён выходной день: {date}")
        log_print("Была вызвана команда /remove_holidays")
        log_print(f"Удалён выходной день: {date}")
    else:
        print(f"День {date} не найден в списке выходных.")
        log_print(f"День {date} не найден в списке выходных.")

# Проверка, является ли день выходным
def is_holiday():
    today = date.today().strftime("%Y-%m-%d")
    return today in holidays
# КОНЕЦ ИЗМЕНЕНИЙ ОТ 18.11.2024
# ==================

# Функция для отправки сообщений с кнопкой в чат №1

def send_start_message():
    if is_holiday():
        print("Сегодня государственный выходной. Сообщение не будет отправлено.")
        log_print("Сегодня государственный выходной. Сообщение не будет отправлено.")
        return
    keyboard = telebot.types.InlineKeyboardMarkup()
    keyboard.row(telebot.types.InlineKeyboardButton("Начать работу", callback_data='start_work'))
    bot.send_message(CHAT_ID_1, "Доброе утро! Нажмите на кнопку, чтобы отметить начало работы.", reply_markup=keyboard)
    log_print("Отправлено сообщение в чат №1 о начале работы.")

# Вызов функции для отправки сообщения с кнопкой в чат №1
send_start_message()

# Словарь для отслеживания времени последнего отправленного уведомления об опоздании для каждого пользователя
last_late_notification_time = {}

# Список исключенных участников
excluded_users = []

# Функция для отправки сообщений об опоздании в чат №2 и личные сообщения сотруднику
def send_late_notification(user_id, user_name):
    if is_holiday():
        print(f"Сегодня выходной. Уведомление об опоздании {user_name} (ID: {user_id}) не отправлено.")
        log_print(f"Сегодня выходной. Уведомление об опоздании {user_name} (ID: {user_id}) не отправлено.")
        return
    try:
        bot.send_message(user_id, f"Внимание, {user_name}! Вы опоздали на работу.")
        log_print("Отправлено личное сообщение опоздавшему сотруднику")
        bot.send_message(CHAT_ID_2, f"Сотрудник {user_name} (ID: {user_id}) опоздал на работу.")
        log_print(f"Сотрудник {user_name} (ID: {user_id}) опоздал на работу.")
    except Exception as e:
        error_message = (f"Не удалось отправить сообщение {user_name}.\n"
                         f"(ID: {user_id}): {str(e)}")
        print(error_message)
        log_print((f"Не удалось отправить сообщение {user_name}.\n"
                         f"(ID: {user_id}): {str(e)}"))
        bot.send_message(CHAT_ID_2, error_message)

# Функция для создания Excel-файла с данными о опозданиях
def create_excel_report(file_name):
    wb = Workbook()
    ws = wb.active
    ws.append(["Имя", "ID", "Во сколько подключились", "На сколько опоздали"])

    for _, row in df.iterrows():
        user_name = row['Имя']
        user_id = row['ID']
        start_time_str = row['Начало работы']

        if 'Нажал на кнопку' in df.columns and row['Нажал на кнопку'] == 'Да':
            connected_time = last_late_notification_time.get(user_id)
            if connected_time:
                connected_time_str = connected_time.strftime('%H:%M')
                late_minutes = (datetime.now() - connected_time).total_seconds() / 60
                late_minutes = round(late_minutes)
            else:
                connected_time_str = "Вовремя"
                late_minutes = "Вовремя"
        else:
            connected_time_str = "Не нажал на кнопку"
            late_minutes = "Не нажал на кнопку"

        ws.append([user_name, user_id, connected_time_str, late_minutes])

    wb.save(file_name)


# Функция для отправки предупреждения об опоздании
def send_warning_message(telegram_id, user_name):
    if is_holiday():
        print(f"Сегодня выходной. Предупреждение {user_name} (ID: {telegram_id}) не отправлено.")
        log_print(f"Сегодня выходной. Предупреждение {user_name} (ID: {telegram_id}) не отправлено.")
        return
    try:
        bot.send_message(telegram_id, f"Внимание, {user_name}! Вы опоздаете через минуту, если не отметите начало работы.")
        log_print(f"Отправлено сообщение: Внимание, {user_name}! Вы опоздаете через минуту, если не отметите начало работы.")
    except Exception as e:
        bot.send_message(CHAT_ID_2, f"Не удалось отправить предупреждение сотруднику {user_name} (ID: {telegram_id}). Ошибка: {e}")
        log_print(f"Не удалось отправить предупреждение сотруднику {user_name} (ID: {telegram_id}). Ошибка: {e}")

# Словарь для отслеживания сотрудников и даты последнего уведомления
notified_late_employees = {}

# Создание блокировки для синхронизации доступа к df
lock = threading.Lock()

def check_users():
    while True:
        current_time = datetime.now()
        print("=============")
        log_print("=============")
        print("Начало цикла")
        log_print("Начало цикла")
        print("Текущее время:", current_time)

        with lock:
            for _, row in df.iterrows():
                start_time_str = row['Начало работы']
                start_time = datetime.strptime(start_time_str, '%H:%M')
                start_datetime = datetime.combine(datetime.now().date(), start_time.time())
                ten_minutes_after_start = start_datetime + timedelta(minutes=10)  # Время для уведомления в личные сообщения
                deadline_time = start_datetime + timedelta(minutes=15)  # Время для уведомления в чат №2
                user_id = row['ID']
                user_name = row['Имя']
                vacation_start = row['Начало отпуска']
                vacation_end = row['Конец отпуска']

                print(f"Сотрудник: {user_name} (ID: {user_id})")
                log_print(f"Сотрудник: {user_name} (ID: {user_id})")
                print("Время начала работы:", start_datetime)
                log_print("Время начала работы:", start_datetime)
                print("Время уведомления в личные сообщения (10 минут после начала):", ten_minutes_after_start)
                log_print("Время уведомления в личные сообщения (10 минут после начала):", ten_minutes_after_start)
                print("Время дедлайна (уведомление в чат №2 через 15 минут):", deadline_time)
                log_print("Время дедлайна (уведомление в чат №2 через 15 минут):", deadline_time)
                print("+++++++++++++")
                log_print("+++++++++++++")

                if pd.notna(vacation_start) and pd.notna(vacation_end):
                    vacation_start = pd.to_datetime(vacation_start)
                    vacation_end = pd.to_datetime(vacation_end)

                    if vacation_start <= current_time <= vacation_end:
                        continue

                # Уведомление в личные сообщения в промежутке между 10 и 15 минутами после начала рабочего дня
                if ten_minutes_after_start <= current_time < deadline_time:
                    if 'Нажал на кнопку' in df.columns and row['Нажал на кнопку'] != 'Да':
                        if user_id not in last_late_notification_time or last_late_notification_time[user_id] < ten_minutes_after_start:
                            send_late_notification(user_id, user_name)
                            last_late_notification_time[user_id] = current_time
                            #print(f"-------------\nОтправлено уведомление об опоздании {user_name} в личные сообщения.") #залупа вышла. Это лишний лог.
                            #log_print(f"-------------\nОтправлено уведомление об опоздании {user_name} в личные сообщения.") #залупа вышла. Это лишний лог.

                # Уведомление в чат №2 через 15 минут после начала рабочего дня
                if current_time >= deadline_time:
                    if user_id not in last_late_notification_time:
                        send_warning_message(user_id, user_name)  # Изменено на send_warning_message
                        last_late_notification_time[user_id] = current_time
                        #print(f"-------------\nОтправлено уведомление об опоздании {user_name} в чат №2.") #залупа вышла. Это лишний лог.
                        #log_print(f"-------------\nОтправлено уведомление об опоздании {user_name} в чат №2.") #залупа вышла. Это лишний лог.

        current_datetime = datetime.now()
        file_name = f"График опозданий {current_datetime.strftime('%Y-%m-%d')}.xlsx"
        create_excel_report(file_name)

        if current_time.strftime('%H:%M') == '10:30':
            with open(file_name, 'rb') as file:
                bot.send_document(CHAT_ID_2, file, caption="#Отчёт_о_опозданиях")
            print("Отправлен файл в чат №2")
            log_print("Отправлен файл в чат №2")

        time.sleep(60)

# Функция для создания Excel-файла с данными о опозданиях
def create_excel_report(file_name):
    wb = Workbook()
    ws = wb.active
    ws.append(["Имя", "ID", "Во сколько подключились", "На сколько опоздали"])

    for _, row in df.iterrows():
        user_name = row['Имя']
        user_id = row['ID']
        start_time_str = row['Начало работы']

        if 'Нажал на кнопку' in df.columns and row['Нажал на кнопку'] == 'Да':
            connected_time = last_late_notification_time.get(user_id)
            if connected_time:
                connected_time_str = connected_time.strftime('%H:%M')
                late_minutes = (datetime.now() - connected_time).total_seconds() / 60
                late_minutes = round(late_minutes)
            else:
                connected_time_str = "Вовремя"
                late_minutes = "Вовремя"
        else:
            connected_time_str = "Не нажал на кнопку"
            late_minutes = "Не нажал на кнопку"

        ws.append([user_name, user_id, connected_time_str, late_minutes])

    wb.save(file_name)

# Генерация отчета
current_datetime = datetime.now()
file_name = f"График опозданий {current_datetime.strftime('%Y-%m-%d')}.xlsx"
create_excel_report(file_name)

# Запускаем проверку пользователей в отдельном потоке
threading.Thread(target=check_users, daemon=True).start()

# Команда для добавления даты в список выходных
@bot.message_handler(commands=['add_holiday'])
def add_holiday_command(message):
    if message.chat.id == CHAT_ID_2:
        try:
            holiday_date = message.text.split()[1]
            if holiday_date not in holidays:
                holidays.add(holiday_date)
                save_holidays(holidays)  # Сохранение изменений
                bot.reply_to(message, f"Дата {holiday_date} добавлена в список выходных.")
                log_print("fДата {holiday_date} добавлена в список выходных.")
            else:
                bot.reply_to(message, f"Дата {holiday_date} уже есть в списке выходных.")
                log_print(f"Дата {holiday_date} уже есть в списке выходных.")
        except IndexError:
            bot.reply_to(message, "Пожалуйста, укажите дату в формате YYYY-MM-DD.")
            log_print("Пожалуйста, укажите дату в формате YYYY-MM-DD.")
    else:
        bot.reply_to(message, "Эта команда недоступна в этом чате.")
        log_print("Эта команда недоступна в этом чате.")

# Команда для удаления даты из списка выходных
@bot.message_handler(commands=['remove_holiday'])
def remove_holiday_command(message):
    if message.chat.id == CHAT_ID_2:
        try:
            holiday_date = message.text.split()[1]
            if holiday_date in holidays:
                holidays.remove(holiday_date)
                save_holidays(holidays)  # Сохранение изменений
                bot.reply_to(message, f"Дата {holiday_date} удалена из списка выходных.")
                log_print(f"Дата {holiday_date} удалена из списка выходных.")
            else:
                bot.reply_to(message, f"Дата {holiday_date} отсутствует в списке выходных.")
                log_print(f"Дата {holiday_date} отсутствует в списке выходных.")
        except IndexError:
            bot.reply_to(message, "Пожалуйста, укажите дату в формате YYYY-MM-DD.")
            log_print("Пожалуйста, укажите дату в формате YYYY-MM-DD.")
    else:
        bot.reply_to(message, "Эта команда недоступна в этом чате.")
        log_print("Эта команда недоступна в этом чате.")

# Команда для отображения списка всех выходных дней
@bot.message_handler(commands=['list_holidays'])
def list_holidays(message):
    if message.chat.id == CHAT_ID_2:
        if holidays:
            holiday_list = "\n".join(holidays)
            bot.reply_to(message, f"Список выходных дней:\n{holiday_list}")
            log_print("Была вызвана команда /list_holidays")
            log_print(f"Список выходных дней:\n{holiday_list}")
        else:
            bot.reply_to(message, "Список выходных дней пуст.")
            log_print("Список выходных дней пуст.")
    else:
        bot.reply_to(message, "Эта команда недоступна в этом чате.")
        log_print("Эта команда недоступна в этом чате.")

# Обработчик нажатия кнопки "Начать работу"
@bot.callback_query_handler(func=lambda call: call.data == 'start_work')
def start_work_callback(call):
    user_id = call.from_user.id
    if not df.empty:
        user_name = df.loc[df['ID'] == user_id, 'Имя'].values
        if len(user_name) > 0:
            user_name = user_name[0]
            bot.answer_callback_query(call.id, "Отлично, начинаем работу!")
            print(f"Сотрудник {user_name} с ID {user_id} нажал на кнопку 'Начать работу'.")
            log_print(f"Сотрудник {user_name} с ID {user_id} нажал на кнопку 'Начать работу'.")
            last_late_notification_time[user_id] = datetime.now()
        else:
            print(f"Не удалось найти участника с ID {user_id} в файле Excel.")
            log_print(f"Не удалось найти участника с ID {user_id} в файле Excel.")
    else:
        print("Файл Excel не содержит данных.")
        log_print("Файл Excel не содержит данных.")

""" 
==============
ФУНКЦИОНАЛ ОТПУСКОВ
НАЧАЛО
==============
"""
# Загрузить Excel файл
def load_employee_data():
    return pd.read_excel('Список сотрудников.xlsx')

# Сохранить изменения в Excel файл
def save_employee_data(df):
    df.to_excel('Список сотрудников.xlsx', index=False)

# Команда для добавления сотрудника в отпуск
@bot.message_handler(commands=['vacation'])
def vacation_command(message):
    # Проверка чата
    if message.chat.id != CHAT_ID_2:
        bot.reply_to(message, "Эта команда недоступна в этом чате.")
        log_print("Попытка вызова команды /vacation вне разрешенного чата.")
        return

    args = message.text.split()

    if len(args) != 4:
        bot.reply_to(message,
                     "Используйте формат: /vacation <Имя_сотрудника> <Дата_начала_отпуска> <Дата_конца_отпуска>")
        log_print("Была вызвана команда /vacation с некорректным форматом.")
        return

    employee_name = args[1]
    start_date = args[2]
    end_date = args[3]

    # Преобразование дат
    try:
        start_date = pd.to_datetime(start_date)
        end_date = pd.to_datetime(end_date)
    except ValueError:
        bot.reply_to(message, "Некорректный формат даты. Используйте формат YYYY-MM-DD.")
        log_print("Некорректный формат даты. Используйте формат YYYY-MM-DD.")
        return

    # Загрузка данных сотрудников
    df = load_employee_data()

    # Поиск ID сотрудника по имени
    employee_row = df[df['Имя'].str.lower() == employee_name.lower()]

    if employee_row.empty:
        bot.reply_to(message, f"Сотрудник с именем {employee_name} не найден.")
        log_print(f"Сотрудник с именем {employee_name} не найден.")
        return

    # Обновление столбцов "Начало отпуска" и "Конец отпуска"
    employee_id = employee_row.iloc[0]['ID']
    df.loc[df['ID'] == employee_id, 'Начало отпуска'] = start_date
    df.loc[df['ID'] == employee_id, 'Конец отпуска'] = end_date

    # Сохранение изменений
    save_employee_data(df)

    bot.reply_to(message, f"Сотруднику {employee_name} добавлен отпуск с {start_date.date()} по {end_date.date()}.")
    log_print(f"Сотруднику {employee_name} добавлен отпуск с {start_date.date()} по {end_date.date()}.")


""" 
==============
ФУНКЦИОНАЛ ОТПУСКОВ
КОНЕЦ
==============
"""

# Обработчик команды "/set_late_minutes"
@bot.message_handler(commands=['set_late_minutes'])
def set_late_minutes(message):
    global late_minutes_allowed
    if message.chat.id == CHAT_ID_2:
        try:
            new_value = int(message.text.split()[1])
            if new_value >= 0:
                late_minutes_allowed = new_value
                bot.reply_to(message, f"Время опоздания установлено на {late_minutes_allowed} минут.")
                log_print("Была вызвана команда /set_late_minutes - установление допустимого времени опоздания")
                log_print(f"Время опоздания установлено на {late_minutes_allowed} минут.")
            else:
                bot.reply_to(message, "Значение должно быть неотрицательным.")
                log_print("Значение должно быть неотрицательным.")
        except (ValueError, IndexError):
            bot.reply_to(message, "Пожалуйста, укажите корректное значение в минутах.")
            log_print("Пожалуйста, укажите корректное значение в минутах.")
    else:
        bot.reply_to(message, "Это заклинание запрещено в этом чате.")
        log_print("Это заклинание запрещено в этом чате.")

# Обработчик команды "/start"
@bot.message_handler(commands=['start'])
def start(message):
    if message.chat.id == CHAT_ID_2:
        log_print("Была вызвана команда /start")
        keyboard = telebot.types.InlineKeyboardMarkup()
        keyboard.row(telebot.types.InlineKeyboardButton("Начать работу", callback_data='start_work'))
        bot.send_message(message.chat.id, "Доброе утро! Нажмите на кнопку, чтобы отметить начало работы.",
                         reply_markup=keyboard)
    else:
        bot.reply_to(message, "Это заклинание запрещено в этом чате)")

# Обработчик команды "/help"
@bot.message_handler(commands=['help'])
def help_command(message):
    if message.chat.id == CHAT_ID_2:
        help_text = (
            "Команды бота:\n"
            "/start - Начать работу и получить кнопку для отметки начала работы.\n"
            "/vacation <Имя_сотрудника> <Дата_начала_отпуска> <Дата_конца_отпуска> - Добавить сотруднику отпуск.\n"
            "/set_late_minutes [количество минут] - Установить допустимое время опоздания.\n"
            "/help - Получить информацию о функционале бота.\n"
            "/add_holiday {Г.М.Д} - добавить дату в выходные.\n"
            "/remove_holiday {Г.М.Д} - удалить дату из выходных.\n"
            "/list_holidays - показать выходные дни.\n\n"
            f"Текущее допустимое время опоздания: {late_minutes_allowed} минут.\n"
            "Бот уведомляет о опоздании за минуту до фактического начала работы.\n"
        )
        log_print("Был вызван /help")
        bot.reply_to(message, help_text)
    else:
        bot.reply_to(message, "Эта команда не доступна в этом чате.")

# Обработчик для всех остальных сообщений
@bot.message_handler(func=lambda message: True)
def handle_messages(message):
    if message.chat.id == CHAT_ID_2 and message.text.startswith('/'):
        command = message.text.split()[0].lower()
        if command == '/start':
            start(message)
        #elif command == '/exclude_user':
            #exclude_user(message)
        #elif command == '/include_user':
            #include_user(message)
        elif command == '/set_late_minutes':
            set_late_minutes(message)
        else:
            bot.reply_to(message, "Неизвестная мне команда.")
    else:
        pass

# Запуск бота
bot.polling()
