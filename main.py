import tkinter
import sys
import pyautogui
import threading
import time
import os
from google.oauth2 import service_account
from googleapiclient.discovery import build
from tkinter import *
from tkinter import ttk
import tkinter.messagebox as mb
from googleapiclient.http import MediaFileUpload
import smtplib
from plyer import notification
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart


# Путь к файлу учетных данных сервисного аккаунта
base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
json_path = os.path.join(base_path, 'cruds.json')

# ID папки на Google Drive, в которую будут добавляться скриншоты
FOLDER_ID = '1UunF1EMwfKvCvGaLXPEuLUDzMSMZPXIT'

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

skubin = User("Скубин","1KPhIQ7_9aArAgOwJuTUc8xJ4A5LYcSRI","skubin.se@es-company.ru","zA4hmMgSubh9XAeqr5px")
bludov = User("Блудов","1Pe6iwPx1523Pe5FCNCjAeZzR-IDmr3bZ")
averkiev = User("Аверкиев","1m5ocXtkPLfHAEnqM8Y9OgBCuSjJikOud")
malahov = User("Малахов","1oj-7_LyQVXJZjNJ1axRlDYIWp2slE48O")
maliha = User("Малыха","17kl1OLHXPxv7rXacso9EI0FNVweEhGTY")
luka = User("Лукьянченко","1ECCP7dZ9v-woSHjODVr1Y3LAdSa07E9f")
andreev = User("Андреев","1VlqvGSBYLnV-MBX-HeUq3H_LD5bYtHSd")
virot = User("Вироцкий","1o2FQdWCCj9GRuN3PdibkqdV3Sd4x4h8A")
test = User("test", "1UunF1EMwfKvCvGaLXPEuLUDzMSMZPXIT")

dixel = Quest("Dixel","interval",3600, "Время проверить Диксель/Юнисервер")
lenta = Quest("Lenta","interval",3600, "Время проверить Ленты")
zabbix = Quest("Zabbix","interval",3600, "Время проверить Забикс на новые оффлайны")
ashan = Quest("Ashan","interval",3600, "Время проверить Ашаны")
indis = Quest("Indis","interval",3600, "Время проверить Индис")


rostovD = Quest("Danfoss Rostov","interval",25, "Время проверить Danfoss Ростов")
stavD = Quest("Danfoss StavHolod","interval",3600, "Время проверить Danfoss Ставрополь")
spbD = Quest("Danfoss Spb","interval",3600, "Время проверить Danfoss СПБ")

rostovС = Quest("Carel Rostov","interval",3600, "Время проверить Carel Ростов")
stavС = Quest("Carel StavHolod","interval",3600, "Время проверить Carel Ставрополь")
spbС = Quest("Carel","interval",3600, "Время проверить Carel СПБ")

upo = Quest("UPO magnit","interval",4000, "Время провести заявки Магнита в УПО")
vkusV = Quest("VkusVill","interval",3600, "Время проверить ВкусВиллы")


periki = [rostovD, stavD, spbD, rostovС, stavС,spbС,upo,vkusV]
lenti = [dixel, lenta, zabbix, ashan, indis]



dict = {
    'Скубин' : skubin,
    'Блудов' : bludov,
    'Аверкиев' : averkiev,
    'Малахов' : malahov,
    'Малыха' : maliha,
    'Лукьянченко' : luka,
    'Андреев' : andreev,
    'Вироцкий' : virot,
    'test' : test
}

position_dict ={
    'Перекрестки' : periki,
    'Ленты\Метро' : lenti
}

# Создание окна tkinter

root = Tk()
root.geometry("400x700")
root.title("Screenshot App")
list_labels = []
def disable_combobox():
    global FOLDER_ID
    global list_labels
    user_combox.configure(state="disabled")
    position_combox.configure(state="disable")
    check.configure(state="disabled")
    pos_choice = position_combox.get()
    pos = position_dict[pos_choice]
    list_labels = make_pos_button(pos)
    value = user_combox.get()
    FOLDER_ID = dict[value].folder_id
    take_screenshot()



def send_message_from():
    value = user_combox.get()
    user = dict[value]
    current_timer = time.strftime("%d.%m.%y")
    subject = f"Инженер мониторинга {current_timer}"
    recipient = "cthutq1597539@mail.ru"
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
    msg['To'] = recipient
    msg['Subject'] = subject

    # Добавляем текст сообщения
    msg.attach(MIMEText(body, 'plain'))

    # Устанавливаем SMTP соединение
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


def start_timer(i,val):
    # Через 3 минуты
    time.sleep(val.timer)
    # Выводим сообщение каждую секунду, если флаг вывода равен True
    while output_flags[i]:
        notification1(val)
        time.sleep(5)

# Инициализация графического интерфейса

global output_flags
global threads
output_flags = [False, False, False, False, False, False, False, False]
threads = [None, None, None, None, None, None, None, None]

# Функция для обработки нажатия на кнопку
def button_click(i, val):
        global output_flags
        global threads
        current_time = time.strftime("%d_%m_%H:%M", time.localtime())
        list_labels[i].configure(text = f" Последнее время: {current_time} ")
        if not output_flags[i]:
            # Создание и запуск потока с функцией start_timer
            threads[i] = threading.Thread(target=start_timer, args=(i,val,))
            threads[i].start()
            output_flags[i] = True
        else:
            output_flags[i] = False
            threads[i].join()
            threads[i] = threading.Thread(target=start_timer, args=(i,val,))
            threads[i].start()
            output_flags[i] = True




label_surname = tkinter.Label(root, text= "Выбрать фамилию")
label_position = tkinter.Label(root,text= 'Выбрать позицию')
user_combox = ttk.Combobox(root, values=list(dict.keys()))
position_combox = ttk.Combobox(root, values= list(position_dict.keys()))
check = tkinter.Button(root,text = "Готово", command=disable_combobox)
position_combox.current(0)
user_combox.current(0)
label = tkinter.Label(root)
task_button = ttk.Button(root, text= "")
label_mail = tkinter.Label(root,text= '___________________________________')
send_button = tkinter.Button(root, text="Отправить письмо", command=send_message_from)
label_qs = tkinter.Label(root,text= '___________________________________')



label.pack()
label_surname.pack()
user_combox.pack()
label_position.pack()
position_combox.pack()
check.pack()
label_mail.pack()
send_button.pack()
label_qs.pack()
def make_pos_button(possition):
    labels = []
    for i,val in enumerate(possition):
        button = tkinter.Button(root, text=f"Чек {val.name}", command=lambda i=i, v=val: button_click(i, v))
        label = tkinter.Label(root,text=f"Last check: ")
        labeler = tkinter.Label(root, text='___________________________________')
        button.pack()
        label.pack()
        labeler.pack()
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





# Запуск функции
current_time()

# Запуск основного цикла tkinter
root.mainloop()