from openpyxl import load_workbook
import pandas as pd

# Подключение файла create.py
from actions.create import create_user

# подключение файла поиска
from outher.search import user_verification

# # Подключение файла для работы с Базой данных
# from connect.SQLConnect import DatabaseConnector
# sql_connector = DatabaseConnector()

# Подключение Person
from outher.person import Person, encrypt_inn

# Подключение BitrixConnect
from connect.bitrixConnect import Bitrix24Connector
bitrix_connector = Bitrix24Connector()

# подключение connect1c
from connect.connect1C import Connector1C
connector_1c = Connector1C()

# Подключение ldapConnect
from connect.ldapConnect import ActiveDirectoryConnector

connector = ActiveDirectoryConnector()
base_dn = connector.getBaseDn()
state = connector.getState()


def change_user(file_path):
    global base_dn, state

    userData = load_workbook(file_path).active

    df_users = pd.read_excel(file_path)
    df_roles = pd.read_excel('info.xlsx')

    # Создание объекта сотрудника
    employee = Person(userData['C2'].value, userData['B2'].value, userData["D2"].value)

    # поиск по info.xlsx
    flags = user_verification(df_roles, df_users)
    # flags = sql_connector.user_verification(userData['G2'].value, userData['J2'].value)

    # Зашифровка ИНН
    INN = encrypt_inn(userData['A2'].value)

    # Новые данные
    name_atrr = {
        'sn': userData["B2"].value.encode('utf-8'),
        'givenName': userData["C2"].value.encode('utf-8'),
        'department': userData["G2"].value.encode('utf-8'),
        'division': userData["I2"].value.encode('utf-8'),
        'company': userData["F2"].value.encode('utf-8'),
        'title': userData["J2"].value.encode('utf-8'),
    }

    new_data = {
        "NAME": str(userData['C2'].value),
        "LAST_NAME": str(userData['B2'].value),
        "UF_DEPARTMENT": str(userData['H2'].value),
        "ACTIVE": "Y",
        "WORK_POSITION": str(userData['J2'].value)
    }

    # Изменение в 1С
    c1_success = True
    if flags['ZUP'] or flags['RTL'] or flags['ERP'] and flags['Normal_account']:
        action = "Изменение"

        existence = connector.search_in_ad(INN)
        if not (existence is None) and len(existence) > 0:
            user_dn, attributes = existence[0]
            cn = attributes.get('cn', [b''])[0].decode('utf-8')
            login = f"\\BINLTD\{cn}"
            # Поиск друга сотрудника одной должности
            friendly = bitrix_connector.find_jobfriend(userData['J2'].value, userData['H2'].value)

            ZUP_value, RTL_value, ERP_value = (
                1 if flags['ZUP'] else 0, 1 if flags['RTL'] else 0, 1 if flags['ERP'] else 0)

            url = connector_1c.getUrlChanges()
            data = {
                'full_name': f'{userData["B2"].value} {userData["C2"].value} {userData["D2"].value}',
                'domain': login,
                'ERP': ERP_value,
                'RTL': RTL_value,
                'ZUP': ZUP_value,
                'job_friend': friendly
            }
            c1_success = connector_1c.send_rq(url, data, employee, userData,action)

    ad_success = False
    update_status_ad = False
    if flags['AD'] and flags['Normal_account']:
        existence = connector.search_in_ad(INN)
        if existence:
            if state == '1':
                try:
                    ad_success, update_status_ad = connector.update_user(existence, name_atrr, employee, userData)
                except Exception as e:
                    bitrix_connector.send_msg_error(
                        f"AD. Изменение. Сотрудник {employee.lastname} {employee.firstname} {employee.surname} из отдела {userData['G2'].value} на должность {userData['J2'].value}. Ошибка в обновлении - {e}")
            else:
                bitrix_connector.send_msg(
                    f"AD. Изменение (Тест): Сотрудник {employee.lastname} {employee.firstname} {employee.surname}. Выполнено"
                )
        else:
            ad_success = create_user(file_path)
    else:
        ad_success = True

    bx_success = False
    update_status_bx = False
    if flags['AD'] and flags['BX24'] and flags['Normal_account']:
        existence = connector.search_in_ad(INN)
        if existence:
            user_dn, attributes = existence[0]
            mail = attributes.get('mail', [b''])[0].decode('utf-8')
            user_info = bitrix_connector.search_email(mail)
            if user_info:
                if state == '1':
                    bx_success, update_status_bx = bitrix_connector.update_user(user_info, new_data, employee, userData)
                else:
                    bitrix_connector.send_msg(
                        f"BX24. Изменение (Тест): Сотрудник {employee.lastname} {employee.firstname} {employee.surname} {user_info.get('ID')}. Выполнено")
            else:
                bx_success = bitrix_connector.create_user(mail, employee, userData)
        else:
            user_info = bitrix_connector.search_user(employee.lastname, employee.firstname, employee.surname)
            if user_info:
                if state == '1':
                    bx_success, update_status_bx = bitrix_connector.update_user(user_info.get('ID'), new_data, employee,
                                                                                userData)
                else:
                    bitrix_connector.send_msg(
                        f"BX24. Изменение (Тест): Сотрудник {employee.lastname} {employee.firstname} {employee.surname} {user_info.get('ID')}. Выполнено")
            else:
                bitrix_connector.send_msg_error(
                    f"BX24. Изменение: Сотрудник {employee.lastname} {employee.firstname} {employee.surname} не найден.")
    else:
        bx_success = True

    if update_status_ad and update_status_bx and flags['Normal_account']:
        existence = connector.search_in_ad(INN)
        user_dn, attributes = existence[0]
        mail = attributes.get('mail', [b''])[0].decode('utf-8')
        user_info = bitrix_connector.search_email(mail)

        # Проверка в каких 1С системах открыли-закрыли доступ
        zup_enabled = flags['ZUP']
        rtl_enabled = flags['RTL']
        erp_enabled = flags['ERP']

        # Формирование сообщения для 1С
        c1_message = "Изменения в 1С:\n"
        c1_message += "- Включен доступ к системе 1С:ЗУП\n" if zup_enabled else "- Отключен доступ к системе 1С:ЗУП\n"
        c1_message += "- Включен доступ к системе 1С:Розница\n" if rtl_enabled else "- Отключен доступ к системе 1С:Розница\n"
        c1_message += "- Включен доступ к системе 1С:ERP\n" if erp_enabled else "- Отключен доступ к системе 1С:ERP\n"

        # Основное сообщение
        message = "Здравствуйте, ваши данные в системе были успешно обновлены в следующих системах:\n"

        if update_status_bx:
            message += "- Bitrix24\n"
        if c1_success:
            message += "- 1С\n" + c1_message

        message += f"\nОбновленные данные:\nОтдел: {userData['G2'].value} \nДолжность: {userData['J2'].value}"

        # Отправка сообщения сотруднику и дублирование
        try:
            if user_info and user_info.get('ID'):
                user_id = user_info.get('ID')
                bitrix_connector.send_msg_user(user_id, message)

                log_message = f"Сообщение отправлено сотруднику {employee.lastname} {employee.firstname} ID = {user_id}:\n'{message}'"
                bitrix_connector.send_msg(log_message)

        except Exception as e:
            error_message = f"Ошибка при отправке уведомления сотруднику {employee.lastname} {employee.firstname}: {e}"
            bitrix_connector.send_msg_error(error_message)
    
    if ad_success and bx_success and c1_success:
        return True
    else:
        return ad_success and bx_success and c1_success






