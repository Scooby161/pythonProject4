import tkinter

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
SERVICE_ACCOUNT_CRED = '/Users/sergejskubin/PycharmProjects/pythonProject4/cruds.json'

# ID папки на Google Drive, в которую будут добавляться скриншоты
FOLDER_ID = '1UunF1EMwfKvCvGaLXPEuLUDzMSMZPXIT'

dict_surname = {
    'Скубин' : '1KPhIQ7_9aArAgOwJuTUc8xJ4A5LYcSRI',
    'Блудов' : '1Pe6iwPx1523Pe5FCNCjAeZzR-IDmr3bZ'
}
def disable_combobox():
    global FOLDER_ID
    combo.configure(state="disabled")
    check.configure(state="disabled")
    value = combo.get()
    FOLDER_ID = dict_surname[value]
    take_screenshot()
    notification()


# Создание окна tkinter
surname = ("Блудов","Скубин")
root = Tk()
root.geometry("400x400")
root.title("Screenshot App")
combo = ttk.Combobox(root, values=surname)
check = tkinter.Button(root,text = "Выбрать сотрудника", command=disable_combobox)

combo.pack()


check.pack()
# Функция для выполнения скриншота
def take_screenshot():
    # Получение текущего времени в формате "чч_мм_сс"
    current_time = time.strftime("%d:%m_%H_%M", time.localtime())

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
    mb.showinfo("Информация", msg)
    root.after(65000, notification)
# Функция для загрузки файла на Google Drive
def upload_to_google_drive(file_path):
    # Создание объекта авторизации
    creds = service_account.Credentials.from_service_account_file(SERVICE_ACCOUNT_CRED, scopes=['https://www.googleapis.com/auth/drive'])

    # Создание объекта Google Drive API
    drive_service = build('drive', 'v3', credentials=creds)

    # Загрузка файла на Google Drive
    file_metadata = {
        'name': os.path.basename(file_path),
        'parents': [FOLDER_ID]
    }

    media = MediaFileUpload(file_path)
    drive_service.files().create(body=file_metadata, media_body=media, fields='id').execute()

label = tkinter.Label(root)
label.pack()

# Запуск функции
current_time()


# Запуск основного цикла tkinter
root.mainloop()