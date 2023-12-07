import tkinter
import sys
import pyautogui
import threading
import time
import os
from google.oauth2 import service_account
from googleapiclient.discovery import build
from tkinter import *
from tkinter import Text
import tkinter.messagebox as mb
from googleapiclient.http import MediaFileUpload
import smtplib
from plyer import notification
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from ttkbootstrap.scrolled import ScrolledFrame

# Путь к файлу учетных данных сервисного аккаунта

base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
json_path = os.path.join(base_path, 'cruds.json')
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
SERVICE_ACCOUNT_FILE = 'cruds.json'
creds = service_account.Credentials.from_service_account_file(
SERVICE_ACCOUNT_FILE, scopes=SCOPES)
service = build('sheets', 'v4', credentials=creds)
FOLDER_ID = '1UunF1EMwfKvCvGaLXPEuLUDzMSMZPXIT'
spreadsheet_id = '1-C3TivYBTi2U2-lQr8CVGfw7Z4PbmcT68PHMSG3COwI'



class User:
    def __init__(self, name, folder_id, email_login=None, email_password=None):
        self.name = name
        self.folder_id = folder_id
        self.email_login = email_login
        self.email_password = email_password

class Quest:
    def __init__(self, name, type, timer,message):
        self.name = name
        self.type = type
        self.timer = timer
        self.message = message


lenti = []
periki = []
object = []
sheet = service.spreadsheets()

root1 = ttk.Window()
root1.geometry("400x400")
root1.title("Screenshot App")



# Создание фрейма и его размещение на области прокрутки
root  = ScrolledFrame(root1,autohide =False,bootstyle ="light")
root.pack(fill= BOTH, expand=YES)

send_object_button = ttk.Button(root, text="Отправить письмо")

periki_xl = sheet.values().get(spreadsheetId=spreadsheet_id, range='QuestDate!A:E').execute()
values_periki = periki_xl.get('values', [])
for i in values_periki:
        if i[0] == 'P':
            periki.append(Quest(i[1],i[2],i[3],i[4]))
        else:
            lenti.append(Quest(i[1],i[2],i[3],i[4]))


login_xl = sheet.values().get(spreadsheetId=spreadsheet_id, range='loginDate!A:D').execute()
values_login = login_xl.get('values', [])

dict_of_login_date = {}
for i in values_login:
    dict_of_login_date[i[0]] = User(i[0],i[1],i[2],i[3])

position_dict ={
    'Перекрестки' : periki,
    'Ленты\Метро' : lenti
}

def get_object_list():
    position = position_combox.get()
    global dict_of_contacts
    global object_combox
    global entry_order
    global entry_comment
    global label_en
    fl_comments = ttk.LabelFrame(root, text="Журнал + Письма", bootstyle="dark")
    if position == 'Перекрестки':
        result3 = sheet.values().get(spreadsheetId=spreadsheet_id, range='PerikiList!A:B').execute()
        values3 = result3.get('values', [])
        dict_of_contacts = {}
        for i in values3:
            dict_of_contacts[i[0]] = i[1]
        send_object_button = ttk.Button(fl_comments, text="Отправить письмо", command=send_message)
        send_jornal_button = ttk.Button(fl_comments, text="Добавить в журнал", command=add_data_to_jornal)
        send_object_button.configure(state='disable')
    else:
        result3 = sheet.values().get(spreadsheetId=spreadsheet_id, range='Contacts!A:B').execute()
        values3 = result3.get('values', [])
        dict_of_contacts = {}
        for i in values3:
            dict_of_contacts[i[0]] = i[1]
        send_object_button = ttk.Button(fl_comments, text="Отправить письмо", command=send_message)
        send_jornal_button = ttk.Button(fl_comments, text="Добавить в журнал", command=add_data_to_jornal)


    label_object = ttk.Label(fl_comments, text='Выбери объект')
    object_combox = ttk.Combobox(fl_comments, values= list(dict_of_contacts.keys()))
    object_combox.configure(height= 10,width=70)
    object_combox.bind('<KeyRelease>',search_object)
    object_combox.set("Начни вводить номер объекта")
    label_alarm = ttk.Label(fl_comments, text='Напиши суть аварии (пример: 13 горка не в сети)')
    entry_order = ttk.Text(fl_comments, height=10, width=70)
    object_combox.bind('<KeyRelease>', search_object)
    label_comment= ttk.Label(fl_comments, text='Напиши комментарий если он есть. Можно оставить пустым')
    entry_comment = ttk.Text(fl_comments, height=10, width=70)
    label_en = ttk.Label(fl_comments, text='___________________________________')
    entry_comment.bind("<Control-v>", paste_text)
    entry_order.bind('<Control-v>', paste_text)

    label_object.grid(row=0, column=0,pady=5)
    object_combox.grid(row=1, column=0,pady=5,padx=5)
    label_alarm.grid(row=2, column=0,pady=5)
    entry_order.grid(row=3, column=0,pady=5)
    label_comment.grid(row=4, column=0,pady=5)
    entry_comment.grid(row=5, column=0,pady=5)
    send_jornal_button.grid(row=6, column=0,pady=5)
    send_object_button.grid(row=7, column=0,pady=5)
    label_en.grid(row=9, column=0,pady=5)

    fl_comments.pack(pady=10)
def search_object(event):
    value = event.widget.get()
    if value == '':
        object_combox['value'] = list(dict_of_contacts.keys())
    else:
        data = []
        for item in list(dict_of_contacts.keys()):
            if value.lower() in item.lower():
                data.append(item)
        object_combox['values'] = data

def paste_text(event):
    text = root.clipboard_get()
    if isinstance(event.widget, ttk.Text):
        event.widget.insert(tkinter.INSERT, text)
    elif isinstance(event.widget, ttk.Entry):
        event.widget.insert(0, text)
def add_qd_to_sheet(data):
    sheet = service.spreadsheets()
    time.sleep(1)
    current_time = time.strftime("%d/%m/%y %H:%M", time.localtime())
    try:
        user_name = user_combox.get()
        poss_name = position_combox.get()
        body = {
            'values': [[
                user_name,
                data.name,
                current_time,
                poss_name,
            ]]
        }
        result = sheet.values().append(
            spreadsheetId=spreadsheet_id, range='QJornal!A:D',
            valueInputOption='USER_ENTERED', body=body).execute()
        return print(body)
    except Exception as e:
        print('An error occurred:', str(e))
        time.sleep(60)
        # Попытка повторного выполнения функции add_email_to_sheet
        add_qd_to_sheet(data)






list_labels = []

def start_smena():
    global FOLDER_ID
    global list_labels
    user_combox.configure(state="disabled")
    position_combox.configure(state="disable")
    check.configure(state="disabled")
    pos_choice = position_combox.get()
    pos = position_dict[pos_choice]
    list_labels = make_pos_button(pos)
    value = user_combox.get()
    FOLDER_ID = dict_of_login_date[value].folder_id
    get_object_list()
    root1.geometry("800x1600")
    take_screenshot()



def send_message_start():
    value = user_combox.get()
    user = dict_of_login_date[value]
    current_timer = time.strftime("%d.%m.%y")
    subject = f"Инженер мониторинга {current_timer}"
    recipients =  ['ok@es-company.ru','monitoring@es-company.ru']
    current_time = time.strftime("%H:%M")

    day_time_min = time.strptime('07:00', '%H:%M')
    day_time_max = time.strptime('19:00', '%H:%M')

    if day_time_min <= time.strptime(current_time, "%H:%M") <= day_time_max:
        body = f"{user.name} на смену заступил {current_timer} c 8-00 до 20-00 (1 смена)"
    else:
        body = f"{user.name} на смену заступил {current_timer} c 20-00 до 08-00 (2 смена)"

    # Настройки SMTP сервера
    smtp_server = 'smtp.mail.ru'
    smtp_port = 465
    sender_email = user.email_login
    sender_password = user.email_password

    # Создаем MIMEMultipart объект и устанавливаем заголовки
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = ', '.join(recipients)

    msg['Subject'] = subject

    # Добавляем текст сообщения
    msg.attach(MIMEText(body, 'plain'))

    # Устанавляем SMTP соединение
    server = smtplib.SMTP_SSL(smtp_server, smtp_port)

    # Авторизуемся на сервере
    server.login(sender_email, sender_password)

    # Отправляем письмо
    server.send_message(msg)
    # Закрываем соединение с сервером
    server.quit()

    # Выводим сообщение об успешной отправке
    label_mail.config (text = '_________Письмо_успешно_отправлено___________')
    send_button.configure(state="disable")

def send_message_to(recipients,subjects):
    value = user_combox.get()
    user = dict_of_login_date[value]
    body = entry_order.get("1.0", "end-1c")
    print(body)
    current_timer = time.strftime("%d.%m.%y")
    subject = f"[МОНИТОРИНГ] Заявка на объект - {subjects}"




    # Настройки SMTP сервера
    smtp_server = 'smtp.mail.ru'
    smtp_port = 465
    sender_email = user.email_login
    sender_password = user.email_password

    # Создаем MIMEMultipart объект и устанавливаем заголовки
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = ', '.join(recipients)
    msg['Subject'] = subject

    # Добавляем текст сообщения
    msg.attach(MIMEText(body, 'plain'))

    # Устанавляем SMTP соединение
    server = smtplib.SMTP_SSL(smtp_server, smtp_port)

    # Авторизуемся на сервере
    server.login(sender_email, sender_password)

    # Отправляем письмо
    server.send_message(msg)
    # Закрываем соединение с сервером
    server.quit()

    # Выводим сообщение об успешной отправке

    label_en.config (text = f'____Письмо на объект {subjects} отправлено___')


def add_data_to_jornal():
    id_of_jornal = '1PxNTX2GTStj_8YRag2jWBvcpnZrCmA1Qbnvphv-6gTo'
    sheet = service.spreadsheets()
    time.sleep(1)
    current_time = time.strftime("%d/%m/%y %H:%M", time.localtime())
    try:
        object = object_combox.get()
        alarm = entry_order.get("1.0", "end-1c")
        commment = entry_comment.get("1.0", "end-1c")

        body = {
            'values': [[
                object,
                alarm,
                commment,
                current_time,
            ]]
        }
        poss_name = position_combox.get()
        if poss_name == 'Перекрестки':
            result = sheet.values().append(
                spreadsheetId=id_of_jornal, range='Periki!A:D',
                valueInputOption='USER_ENTERED', body=body).execute()
            label_en.config(text=f'____Заметка на объект {object} добавлена___')
        else:
            result = sheet.values().append(
                spreadsheetId=id_of_jornal, range='Len/Metro!A:D',
                valueInputOption='USER_ENTERED', body=body).execute()
            label_en.config(text=f'____Заметка на объект {object} добавлена___')

    except Exception as e:
        print('An error occurred:', str(e))
        time.sleep(60)
        # Попытка повторного выполнения функции add_email_to_sheet
        add_qd_to_sheet(data)



def start_timer(i,val):
    if val.type == 'interval':
        z = 0
        while output_flags[i]:
            if z < int(val.timer):
                z = z + 1
                time.sleep(1)
            else:
                while output_flags[i]:
                    if z < 300:
                        z = z + 1
                        time.sleep(1)
    else:
        while output_flags[i]:
            current_time = time.localtime()
            target_time = time.strptime(val.timer, "%H:%M:%S")
            if (current_time.tm_hour == target_time.tm_hour and
                    current_time.tm_min == target_time.tm_min and
                    current_time.tm_sec >= target_time.tm_sec):
                notification1(val)
                break
            time.sleep(1)
# Инициализация графического интерфейса

global output_flags
global threads
output_flags = [False, False, False, False, False, False, False, False, False, False, False, False, False, False]
threads = [None, None, None, None, None, None, None, None, None, None, None, None, None, None]

# Функция для обработки нажатия на кнопку
def button_click(i, val):
        global output_flags
        global threads
        current_time = time.strftime("%d_%m_%H:%M", time.localtime())
        add_qd_to_sheet(val)
        list_labels[i].configure(text = f" Последнее время: {current_time} ")
        if not output_flags[i]:
            # Создание и запуск потока с функцией start_timer
            output_flags[i] = True
            threads[i] = threading.Thread(target=start_timer, args=(i,val,))
            threads[i].start()

        else:
            if val.type == 'interval':
                output_flags[i] = False
                threads[i].join()
                threads[i] = threading.Thread(target=start_timer, args=(i,val,))
                threads[i].start()
                output_flags[i] = True
            else:
                output_flags[i] = False




def send_message():
    object = object_combox.get()
    recipients = ['alarm@es-company.ru','monitoring@es-company.ru', dict_of_contacts[object]]
    subject = f"{object}"
    send_message_to(recipients, subject)  # вызов функции для отправки сообщения
frame_login = ttk.LabelFrame(root, text="Выбери свои данные",bootstyle="dark")
label_surname = ttk.Label(frame_login, text= "Выбрать фамилию")
label_position = ttk.Label(frame_login,text= 'Выбрать позицию')
user_combox = ttk.Combobox(frame_login, values=list(dict_of_login_date.keys()))
position_combox = ttk.Combobox(frame_login, values= list(position_dict.keys()))
check = ttk.Button(frame_login,text = "Готово", command=start_smena)

label_surname.grid(row=0, column=0,pady=5,padx=5)
label_position.grid(row=2, column=0,pady=5,padx=5)
user_combox.grid(row=1, column=0,pady=5,padx=5)
position_combox.grid(row=3, column=0,pady=5,padx=5)
check.grid(row=4, column=0,pady=5,padx=5)

position_combox.current(0)
user_combox.current(0)

label = ttk.Label(root)
task_button = ttk.Button(root, text= "")
label_frame = ttk.LabelFrame(root, text="Отправить письмо о начале смены в О.К.?",bootstyle="secondary")
send_button = ttk.Button(label_frame, text="Отправить письмо", command=send_message_start)

send_button.grid(row=0, column=0, sticky="w",pady=5,padx=70)





label.pack()
frame_login.pack(pady=10)
label_frame.pack(pady=10)



def make_pos_button(possition):
    labels = []

    for i,val in enumerate(possition):
        label_frame = ttk.LabelFrame(root, text=f"{val.name}",bootstyle="dark")
        if val.type == 'interval':
            button = ttk.Button(label_frame, text=f"Чек {val.name}", command=lambda i=i, v=val: button_click(i, v))
            button.grid(row=0, column=0, sticky="w",pady=5,padx=5)
        else:
            button = ttk.Button(label_frame, text=f"Выполнить {val.name}", command=lambda i=i, v=val: button_click(i, v))
            button.grid(row=0, column=0, sticky="w",pady=5,padx=5)
        label = ttk.Label(label_frame,text=f"Последнее время: ")
        label.grid(row=1, column=0,pady=3)
        label_frame.pack(pady=5)
        labels.append(label)
    return labels


    # Функция для выполнения скриншота
def take_screenshot():
    # Получение текущего времени в формате "чч_мм_сс"
    current_time = time.strftime("%d_%m_%H:%M", time.localtime())

    # Создание скриншота
    screenshot = pyautogui.screenshot()

    # Сохранение скриншота в файл
    screenshot_path = f"{current_time}.png"
    screenshot.save(screenshot_path)

    # Добавление скриншота в папку на Google Drive
    upload_to_google_drive(screenshot_path)

    # Удаление локального файла скриншота
    os.remove(screenshot_path)

    # Запуск функции через каждые 3 минуты
    root.after(180000, take_screenshot)
    

def current_time():
    current_timer = time.strftime("%d/%m_%H:%M")
    label.config (text = current_timer)
    label.after(60000, current_time)


def notification1(quest):
    name = quest.name
    msg = quest.message
    notification.notify(message = msg,app_name = 'BBWFY', title = name,timeout = 2)
def notification2(quest):
    name = quest.name
    msg = quest.message
    print(f"{name} и {msg}")

# Функция для загрузки файла на Google Drive
def upload_to_google_drive(file_path):
    # Создание объекта авторизации
    creds = service_account.Credentials.from_service_account_file(json_path, scopes=['https://www.googleapis.com/auth/drive'])

    # Создание объекта Google Drive API
    drive_service = build('drive', 'v3', credentials=creds)

    # Загрузка файла на Google Drive
    file_metadata = {
        'name': os.path.basename(file_path),
        'parents': [FOLDER_ID]
    }

    media = MediaFileUpload(file_path)
    drive_service.files().create(body=file_metadata, media_body=media, fields='id').execute()





root1.mainloop()