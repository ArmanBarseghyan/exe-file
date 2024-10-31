import os
import requests
from win32com.client import Dispatch

# Определяем текущую директорию, где расположен .exe или .py файл
current_directory = os.path.dirname(os.path.abspath(__file__))

# Название папки и ссылка на сайт
folder_name = "МояПапка"
folder_path = os.path.join(current_directory, folder_name)
url = "https://admin.spectrepro.ru/martin/events-line"  # Ваша ссылка
icon_url = "https://drive.google.com/file/d/1mh6E9WVtV6692ORe8LfeyeugyziYiDwh/view?pli=1"  # Ссылка на .ico файл

# Создаем папку в текущей директории
os.makedirs(folder_path, exist_ok=True)

# Загружаем .ico файл
icon_path = os.path.join(folder_path, "icon.ico")
response = requests.get(icon_url)
with open(icon_path, "wb") as icon_file:
    icon_file.write(response.content)

# Создаем ярлык с иконкой и ссылкой
shortcut_path = os.path.join(folder_path, "МойЯрлык.lnk")
shell = Dispatch("WScript.Shell")
shortcut = shell.CreateShortcut(shortcut_path)
shortcut.TargetPath = url
shortcut.IconLocation = icon_path
shortcut.Save()

print("Папка и ярлык созданы в той же папке, где находится скрипт.")
