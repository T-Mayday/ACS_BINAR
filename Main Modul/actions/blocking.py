from openpyxl import load_workbook
import pandas as pd


# подключение файла поиска
from outher.search import user_verification

# # Подключение файла для работы с Базой данных
# from connect.SQLConnect import DatabaseConnector
# sql_connector = DatabaseConnector()

# подключение файла сообщений
# from message.message import log

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

# Подключение к SM
from connect.SMConnect import SMConnect

sm_conn = SMConnect()
sm_conn.connect_SM()
test_role_id = sm_conn.getRoleID()

# Данные МД АУДИТА
from connect.MDConnect import MDAUIDConnect
MD_AUDIT = MDAUIDConnect()



def blocking_user(file_path):
    global base_dn, state

    userData = load_workbook(file_path).active

    df_users = pd.read_excel(file_path)
    df_roles = pd.read_excel('info.xlsx')

    # поиск по info.xlsx
    flags = user_verification(df_roles, df_users)
    # flags = sql_connector.user_verification(userData['G2'].value, userData['J2'].value)

    # Создание объекта сотрудника
    employee = Person(userData['C2'].value, userData['B2'].value, userData["D2"].value)

    # Зашифровка ИНН
    INN = encrypt_inn(userData['A2'].value)

    # Блокировка в 1C
    c1_success = True
    if (flags['ZUP'] or flags['RTL'] or flags['ERP']) and flags['Normal_account']:
        action = 'Блокировка'

        # Поиск друга сотрудника одной должности
        friendly = bitrix_connector.find_jobfriend(userData['J2'].value, userData['G2'].value)

        ZUP_value, RTL_value, ERP_value = (1 if flags['ZUP'] else 0, 1 if flags['RTL'] else 0, 1 if flags['ERP'] else 0)

        url = connector_1c.getUrlBlock()
        data = {
            'full_name': f'{userData["B2"].value} {userData["C2"].value} {userData["D2"].value}',
            'ERP': ERP_value,
            'RTL': RTL_value,
            'ZUP': ZUP_value,
            'job_friend': friendly
        }
        if state == '1':
            c1_success = connector_1c.send_rq(url, data, employee, userData, action)
        else:
            c1_success = True
            bitrix_connector.send_msg(
                f"1С. Блокировка (Тест): Сотрудник {employee.lastname, employee.firstname, employee.surname}. Выполнено")

    # Блокировка в SM Глобальный
    sm_success = True
    if flags['SM_GEN'] and flags['Normal_account']:
        existence = connector.search_in_ad(INN)
        if existence:
            user_dn, attributes = existence[0]
            login = attributes.get('sAMAccountName', [b''])[0].decode('utf-8')
            sm_login = employee.transform_login(login)
            user_exists = sm_conn.user_exists(sm_login)
            if user_exists:
                try:
                    if state == '1':
                        if user_exists[0] == '1':
                            sm_success = sm_conn.block_user(employee.sm_login)
                            bitrix_connector.send_msg(f"СуперМаг Глобальный. Блокировка: {employee.lastname, employee.firstname, employee.surname} {sm_login}. Выполнено")
                    else:
                        bitrix_connector.send_msg(
                            f"СуперМаг Глобальный (Тест). Блокировка: {employee.lastname, employee.firstname, employee.surname} {sm_login}. Выполнено")
                except Exception as e:
                    sm_success = False
                    bitrix_connector.send_msg_error(f'СуперМаг Глобальный. Блокировка: {employee.lastname, employee.firstname, employee.surname} из отдела {userData["G2"].value} на должность {userData["J2"].value} {sm_login}. Не выполнено')

    sm_local_success = True
    if flags['SM_LOCAL'] and flags['Normal_account']:
        sm_local_success = True

    # Блокировка в Bitrix24
    bx24_success = True
    if flags['AD'] and flags['BX24'] and flags['Normal_account']:
        existence = connector.search_in_ad(INN)
        if len(existence) > 0:
            user_dn, attributes = existence[0]
            mail = attributes.get('mail', [b''])[0].decode('utf-8')
            user_info = bitrix_connector.search_email(mail)
            if user_info:
                if state == '1':
                    if user_info.get('Active') == 'True':
                        bx24_success = bitrix_connector.block_user(user_info.get('ID'), employee, userData)
                else:
                    bitrix_connector.send_msg(
                        f"BX24. Блокировка (Тест): {employee.firstname} {employee.lastname} {employee.surname}. Выполнено")
        else:
            bx24_success = True
            bitrix_connector.send_msg_error(
                f'BX24. Блокировка: {employee.firstname} {employee.lastname} {employee.surname}. Пользователь не найден в AD. Не выполнено.')
    else:
        bx24_success = True

    # Блокировка в МД АУДИТ
    md_success = True
    if flags['AD'] and flags['Normal_account']:
        existence = connector.search_in_ad(INN)
        if existence:
            user_dn, attributes = existence[0]
            mail = attributes.get('mail', [b''])[0].decode('utf-8')
            user_info = MD_AUDIT.find_user_by_email(mail)
            if user_info:
                if state == '1':
                    if user_info.get('active') == 'True':
                        md_success = MD_AUDIT.block_user(user_info.get('id'), employee.lastname, employee.firstname, employee.surname, userData['G2'].value, userData['J2'].value)
                else:
                    bitrix_connector.send_msg(
                        f"MD_AUDIT. Блокировка (Тест): {employee.firstname} {employee.lastname} {employee.surname}. Выполнено")

    # Блокировка в Active Directory
    ad_success = True
    if flags['AD'] and flags['Normal_account']:
        existence = connector.search_in_ad(INN)
        if existence:
            if state == '1':
                ad_success = connector.block_user(existence, employee, userData)
            else:
                bitrix_connector.send_msg(
                    f"AD. Блокировка (Тест): Сотрудник {employee.lastname, employee.firstname, employee.surname}. Выполнено")
        else:
            ad_success = False
    
    if ad_success and bx24_success and c1_success and sm_success and sm_local_success and md_success:
        return True
    else:
        return False







