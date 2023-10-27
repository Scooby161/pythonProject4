import tkinter
import sys
import pyautogui
import time
import os
from google.oauth2 import service_account
from googleapiclient.discovery import build
from tkinter import *
from tkinter import ttk
import tkinter.messagebox as mb
from googleapiclient.http import MediaFileUpload

# Путь к файлу учетных данных сервисного аккаунта
base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
json_path = os.path.join(base_path, 'cruds.json')

# ID папки на Google Drive, в которую будут добавляться скриншоты
FOLDER_ID = '1UunF1EMwfKvCvGaLXPEuLUDzMSMZPXIT'

class User:
    def __init__(self, name, folder_id):
        self.name = name
        self.folder_id = folder_id
        self.password = 0
        self.email_login = 0
        self.email_password = 0



skubin = User("Скубин","1KPhIQ7_9aArAgOwJuTUc8xJ4A5LYcSRI")
bludov = User("Блудов","1Pe6iwPx1523Pe5FCNCjAeZzR-IDmr3bZ")
averkiev = User("Аверкиев","1m5ocXtkPLfHAEnqM8Y9OgBCuSjJikOud")
malahov = User("Малахов","1oj-7_LyQVXJZjNJ1axRlDYIWp2slE48O")
maliha = User("Малыха","17kl1OLHXPxv7rXacso9EI0FNVweEhGTY")
luka = User("Лукьянченко","1ECCP7dZ9v-woSHjODVr1Y3LAdSa07E9f")
andreev = User("Андреев","1VlqvGSBYLnV-MBX-HeUq3H_LD5bYtHSd")
virot = User("Вироцкий","1o2FQdWCCj9GRuN3PdibkqdV3Sd4x4h8A")
test =  User("test", "1UunF1EMwfKvCvGaLXPEuLUDzMSMZPXIT")
dict = {
    'Скубин' : skubin,
    'Блудов' : bludov,
    'Аверкиев' : averkiev,
    'Малахов' : malahov,
    'Малыха' : maliha,
    'Лукьянченко' : luka,
    'Андреев' : andreev,
    'Вироцкий' : virot,
    'Выбери фамилию' : test
}

position_dict ={
    'Перекрестки' : 'periki',
    'Ленты\Метро' : 'lenti'
}

# Создание окна tkinter

root = Tk()
root.geometry("400x400")
root.title("Screenshot App")

def disable_combobox():
    global FOLDER_ID
    user_combox.configure(state="disabled")
    position_combox.configure(state="disable")
    check.configure(state="disabled")
    value = user_combox.get()
    FOLDER_ID = dict[value].folder_id
    take_screenshot()
    notification()


user_combox = ttk.Combobox(root, values=list(dict.keys()))
position_combox = ttk.Combobox(root, values= list(position_dict.keys()))
check = tkinter.Button(root,text = "Выбрать сотрудника", command=disable_combobox)
position_combox.current(0)
user_combox.current(8)
label = tkinter.Label(root)

label.pack()
user_combox.pack()
position_combox.pack()
check.pack()






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


def notification():
    msg = "Проверь диксель"
    msg2 = " Проверь ВкусВиллы"

    if position_combox.get() == "Перекрестки":
        mb.showinfo("Информация", msg2)
    else:
        mb.showinfo("Информация", msg)

    root.after(3600000, notification)

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