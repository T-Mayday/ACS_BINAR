import os
import fnmatch
import shutil
import time
from datetime import datetime
from openpyxl import load_workbook

# Подключение к ldap
from connect.ldapConnect import ActiveDirectoryConnector
connector = ActiveDirectoryConnector()

# Инициализация папок
input_dir = connector.getInput()
output_dir = connector.getOutput()
waste_dir = connector.getWaste()

# Подключение файла create.py
from actions.create import create_user

# Подключение файла change.py
from actions.change import change_user

# Подключение файла blocking.py
from actions.blocking import blocking_user

# Подключение файла holiday.py
from actions.holiday import holiday

# Подключение файла сообщения
from message.message import send_msg, log


def move_file(file_path,output_dir):
    if os.path.exists( output_dir+os.path.basename(file_path) ):
        dest_name = os.path.join(output_dir,datetime.now().strftime('%H%M%S')+os.path.basename(file_path))
    else:
        dest_name = output_dir
    shutil.move(file_path, dest_name)


def process_file(file_path):
    try:
        userData = load_workbook(file_path).active

        action = userData["M2"].value

        if action == "Создание":
            created = create_user(file_path)
            if created:
                move_file(file_path, output_dir)
                send_msg(f'Файл {os.path.basename(file_path)} обработан')
            else:
                raise ValueError("Ошибка при создании пользователя")
        elif action == "Изменение":
            changed = change_user(file_path)
            if changed:
                move_file(file_path, output_dir)
                send_msg(f'Файл {os.path.basename(file_path)} обработан')
            else:
                raise ValueError("Ошибка при изменении пользователя")
        elif action == "Блокировка":
            blocking_user(file_path)
            move_file(file_path, output_dir)
            send_msg(f'Файл {os.path.basename(file_path)} обработан')

        elif action in ["Отпуск", "больничный", "командировка"]:
            holiday(file_path)
            move_file(file_path, output_dir)
            send_msg(f'Файл {os.path.basename(file_path)} обработан')
    except Exception as e:
        move_file(file_path, waste_dir)
        send_msg(f'Ошибка обработки файла {file_path}: {str(e)}')

def main():
    ver = 'V.17.09.2024'

    if connector.getState() == "1":
        mode = 'Боевой режим!'
    else:
        mode = 'Тестовый режим!'

#    log.info(f"Старт Версии {ver} {mode}")
    send_msg(f"Старт Версии {ver} {mode}")
    n = 0
    while True:
        for root, dirs, files in os.walk(input_dir):
            for file in fnmatch.filter(files, "*.xlsx"):
                log.info(os.path.join(root, file))
                process_file(os.path.join(root, file))
        time.sleep(60)
        print(datetime.now())            

if __name__ == '__main__':
    main()