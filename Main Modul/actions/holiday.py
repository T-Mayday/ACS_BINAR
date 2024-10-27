from openpyxl import load_workbook
import random
import string
from datetime import datetime
import pandas as pd

# Подключение ldapConnect
from connect.ldapConnect import ActiveDirectoryConnector
connector = ActiveDirectoryConnector()
base_dn = connector.getBaseDn()
state = connector.getState()


# подключение файла поиска
from outher.search import user_verification

# Подключение файла сообщения
from message.message import log


# Подключение BitrixConnect
from connect.bitrixConnect import Bitrix24Connector

bitrix_connector = Bitrix24Connector()
bx24, tokens = bitrix_connector.connect()


def generate_random_string(length=12):
   characters = string.ascii_letters + string.digits
   random_string = ''.join(random.choice(characters) for _ in range(length))
   return random_string


def check_existing_holiday(user_id, start_date, end_date):
    try:
        existing_holidays = bx24.call('lists.element.get', {
            'IBLOCK_TYPE_ID': 'bitrix_processes',
            'IBLOCK_ID': '52',
            'FILTER': {
                'CREATED_BY': user_id
            }
        })

        if 'result' in existing_holidays and existing_holidays['result']:
            for holiday in existing_holidays['result']:
                holiday_start = next(iter(holiday['PROPERTY_320'].values()), None)
                holiday_end = next(iter(holiday['PROPERTY_322'].values()), None)

                if holiday_start == start_date and holiday_end == end_date:
                    return True
        return False
    except Exception as e:
        bitrix_connector.send_msg_error(f"Ошибка при проверке существующего отпуска: {e}")
        return False


def holiday(file_path):
    global random_string, state

    excel_data = load_workbook(file_path).active
    df_users = pd.read_excel(file_path)
    df_roles = pd.read_excel('info.xlsx')
    # поиск по info.xlsx
    flags = user_verification(df_roles, df_users)
    random_string = generate_random_string()

    def format_date(date):
        if isinstance(date, datetime):
            return date.strftime("%d.%m.%Y")
        elif isinstance(date, str):
            try:
                date_obj = datetime.strptime(date, "%d.%m.%Y %H:%M:%S")
                return date_obj.strftime("%d.%m.%Y")
            except ValueError:
                raise ValueError("Invalid date format. Expected format: 'dd.mm.yyyy HH:MM:SS'")
        else:
            raise ValueError("Input must be a string or a datetime object")

    lastname = excel_data['B2'].value
    firstname = excel_data['C2'].value
    surname = excel_data['D2'].value
    start_holiday = format_date(excel_data['N2'].value)
    end_holiday = format_date(excel_data['O2'].value)

    state_holiday = excel_data['P2'].value
    if state_holiday is None:
        state_holiday = excel_data['M2'].value
    # Виды отпуска
    type_holiday = {
        'отпуск ежегодный': '332',
        'командировка': '334',
        'больничный': '336',
        'декретный': '338',
        'за свой счет': '340',
        'другое': '342'
    }
    # Виды сообщений
    message_templates = {
        'отпуск ежегодный': f'Здравствуйте, Ваш ежегодный отпуск зарегистрирован в Графике отсутствия. \n\nДаты: с {start_holiday} по {end_holiday}.',
        'командировка': f'Здравствуйте, Ваша командировка зарегистрирована в Графике отсутствия. \n\nДаты: с {start_holiday} по {end_holiday}.',
        'больничный': f'Здравствуйте, Ваш больничный зарегистрирован в Графике отсутствия. \n\nДаты: с {start_holiday} по {end_holiday}.',
        'декретный': f'Здравствуйте, Ваш декретный отпуск зарегистрирован в Графике отсутствия. \n\nДаты: с {start_holiday} по {end_holiday}.',
        'за свой счет': f'Здравствуйте, Ваш отпуск за свой счет зарегистрирован в Графике отсутствия. \n\nДаты: с {start_holiday} по {end_holiday}.',
        'другое': f'Здравствуйте, Ваше отсутствие зарегистрировано в Графике отсутствия. \n\nДаты: с {start_holiday} по {end_holiday}.'
    }

    bx24_success = True

    if flags['BX24'] and flags['Normal_account']:
        user_id = bitrix_connector.search_user(bx24,lastname, firstname, surname)
        if not (user_id is None):
            if state == "1":
                if state_holiday.lower() in type_holiday:
                    existence_date = check_existing_holiday(user_id, start_holiday, end_holiday)
                    if not existence_date:
                        result = type_holiday[state_holiday.lower()]
                        date = {
                            'IBLOCK_TYPE_ID': 'bitrix_processes',
                            'IBLOCK_ID': '52',
                            'ELEMENT_CODE': random_string,
                            'FIELDS': {
                                'NAME': 'ОТПУСК',
                                'CREATED_BY': f'{user_id}',
                                'PROPERTY_320': f'{start_holiday}',
                                'PROPERTY_322': f'{end_holiday}',
                                'PROPERTY_324': [
                                    f'{result}'
                                ]
                            }
                        }
                        try:
                            bx24.refresh_tokens()
                            result = bx24.call('lists.element.add', date)
                            if result.get('error'):
                                error_message = result.get('error_description')
                                bitrix_connector.send_msg_error(
                                    f"BX24. {state_holiday.upper()}: Сотрудник {lastname, firstname, surname}, должность {excel_data['J2'].value} Ошибка: {error_message} {date}")
                                bx24_success = False
                            if result.get('result'):
                                msg = message_templates.get(state_holiday.lower(), 'Ваш отпуск зарегистрирован.')
                                bitrix_connector.send_msg_user(user_id, msg)
                                bitrix_connector.send_msg(
                                    f"BX24. {state_holiday.upper()}: Сотрудник {lastname, firstname, surname}, должность {excel_data['J2'].value}. Выполнено")
                                bx24_success = True
                        except Exception as e:
                            bx24_success = False
                            bitrix_connector.send_msg_error(f"BX24. {state_holiday.upper()}: Сотрудник {lastname, firstname, surname}, должность {excel_data['J2'].value} Ошибка: {str(e)} {date}")
            else:
                bitrix_connector.send_msg(
                    f"BX24. {state_holiday.upper()} (Тест): Сотрудник {lastname, firstname, surname}, должность {excel_data['J2'].value}. Выполнено")
        else:
            bitrix_connector.send_msg(
                f"BX24. {state_holiday.upper()}: Сотрудник {lastname, firstname, surname}, должность {excel_data['J2'].value}. Ошибка. Сотрудник не найден по ФИО или не активен.")
    return bx24_success

