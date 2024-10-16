import os
import fnmatch
import shutil
import time
from datetime import datetime
from openpyxl import load_workbook

# Подключение к ldap
from connect.ldapConnect import ActiveDirectoryConnector
connector = ActiveDirectoryConnector()

from connect.bitrixConnect import Bitrix24Connector
bitrix_connector = Bitrix24Connector()

# Инициализация папок
input_dir = connector.getInput()
output_dir = connector.getOutput()
waste_dir = connector.getWaste()
error_dir = connector.getError()

# Подключение файла create.py
from actions.create import create_user

# Подключение файла change.py
from actions.change import change_user

# Подключение файла blocking.py
from actions.blocking import blocking_user

# Подключение файла holiday.py
from actions.holiday import holiday

# Подключение файла сообщения
from message.message import log


def move_file(file_path,output_dir):
    if os.path.exists( output_dir+os.path.basename(file_path) ):
        dest_name = os.path.join(output_dir,datetime.now().strftime('%H%M%S')+os.path.basename(file_path))
    else:
        dest_name = output_dir
    shutil.move(file_path, dest_name)

def move_file_waste(file_path,output_dir):
    if not os.path.exists( output_dir+os.path.basename(file_path) ):
        shutil.move(file_path, output_dir)

# Перемещаем файлы из waste обратно в input через час
def move_back():

    current_time = time.time()
    files_to_move = []  # Список для хранения файлов, которые будут перемещены

    for root, dirs, files in os.walk(waste_dir):
        for file in fnmatch.filter(files, "*.xlsx"):
            file_path = os.path.join(root, file)

            time_in_waste = current_time - os.path.getmtime(file_path)

            if time_in_waste > 3600:
                files_to_move.append(file_path)

    if files_to_move:
        bitrix_connector.send_msg(
            f"Перенесены следующие файлы обратно в input: {', '.join([os.path.basename(f) for f in files_to_move])}")
        for file_path in files_to_move:
            move_file(file_path, input_dir)


# Проверка валидации данных
def validate_user_data(workbook):
    errors = []
    required_fields = {
        'inn': {'cell': 'A2', 'label': 'ИНН', 'check': lambda v: v and v.isdigit() and len(v) == 12,
                'error': "Некорректный ИНН (должен быть 12 цифр)."},
        'last_name': {'cell': 'B2', 'label': 'Фамилия', 'check': lambda v: v, 'error': "Отсутствует фамилия."},
        'first_name': {'cell': 'C2', 'label': 'Имя', 'check': lambda v: v, 'error': "Отсутствует имя."},
        'middle_name': {'cell': 'D2', 'label': 'Отчество', 'check': lambda v: True},
        'employment_date': {'cell': 'E2', 'label': 'Дата трудоустройства', 'check': lambda v: v,
                            'error': "Отсутствует дата трудоустройства."},
        'organization': {'cell': 'F2', 'label': 'Организация', 'check': lambda v: v,
                         'error': "Отсутствует организация."},
        'department': {'cell': 'G2', 'label': 'Отдел', 'check': lambda v: v, 'error': "Отсутствует отдел."},
        'bitrix_department_code': {'cell': 'H2', 'label': 'Код отдела Битрикса', 'check': lambda v: v,
                                   'error': "Отсутствует код отдела Битрикса."},
        'subdivision': {'cell': 'I2', 'label': 'Подразделение', 'check': lambda v: True},

        'position': {'cell': 'J2', 'label': 'Должность', 'check': lambda v: v, 'error': "Отсутствует должность."},
        # 'mobile_number': {'cell': 'K2', 'label': 'Номер мобильного телефона', 'check': lambda v: True,
        #                   'error': "Некорректный номер мобильного телефона."},
        # 'birth_date': {'cell': 'L2', 'label': 'Дата рождения', 'check': lambda v: v,
        #                'error': "Отсутствует дата рождения."},
        # 'mobile_number': {'cell': 'K2', 'label': 'Номер мобильного телефона', 'check': lambda v: v,
        #                   'error': "Отсутствует номер мобильного телефона."},
        'status': {'cell': 'M2', 'label': 'Статус',
                   'check': lambda v: v in ['Создание', 'Изменение', 'Блокировка', 'Отпуск', 'больничный',
                                            'командировка'],
                   'error': "Некорректный статус."},
        'stores': {'cell': 'N2', 'label': 'Магазины', 'check': lambda v: True}
    }
    user_data = {}
    try:
        for field, props in required_fields.items():
            value = workbook[props['cell']].value
            user_data[field] = value
            if not props['check'](value):
                errors.append(props['error'])
    except Exception as e:
        errors.append(f"Ошибка при чтении файла: {str(e)}")

    return errors, user_data

def process_file(file_path):
    try:
        workbook = load_workbook(file_path).active

        errors, user_data = validate_user_data(workbook)
        if errors:
            move_file(file_path,error_dir)
            bitrix_connector.send_msg(f"Ошибки валидации: {', '.join(errors)}")

        action = user_data.get("status")

        if action == "Создание":
            created = create_user(file_path)
            if created:
                move_file(file_path, output_dir)
            else:
                raise ValueError("Ошибка при создании пользователя")
        elif action == "Изменение":
            changed = change_user(file_path)
            if changed:
                move_file(file_path, output_dir)
            else:
                raise ValueError("Ошибка при изменении пользователя")
        elif action == "Блокировка":
            bloсking = blocking_user(file_path)
            if bloсking:
                move_file(file_path, output_dir)
            else:
                raise ValueError("Ошибка при блокировке пользователя")

        elif action in ["Отпуск", "больничный", "командировка"]:
            holiday_user = holiday(file_path)
            if holiday_user:
                move_file(file_path, output_dir)
            else:
                raise ValueError("Ошибка при назначении отпуска сотруднику")
    except Exception as e:
        move_file_waste(file_path, waste_dir)
        bitrix_connector.send_msg(f'Ошибка обработки файла {file_path}: {str(e)}')

def main():
    ver = 'V.16.10.2024'

    if connector.getState() == "1":
        mode = 'Боевой режим!'
    else:
        mode = 'Тестовый режим!'

#    bitrix_connector.send_msg_adm(f"Старт Версии {ver} {mode}")
    bitrix_connector.send_msg(f"Старт Версии {ver} {mode}")
    n = 0
    while True:
        for root, dirs, files in os.walk(input_dir):
            for file in fnmatch.filter(files, "*.xlsx"):
                log.info(os.path.join(root, file))
                process_file(os.path.join(root, file))

        time.sleep(60)
        move_back()
        print(datetime.now())            

if __name__ == '__main__':
    main()