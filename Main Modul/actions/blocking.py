from openpyxl import load_workbook
import pandas as pd
import ldap
import requests
import random
import string
import time


# подключение файла поиска
from outher.search import user_verification

# подключение файла сообщений
from message.message import log

# Подключение BitrixConnect
from connect.bitrixConnect import Bitrix24Connector

bitrix_connector = Bitrix24Connector()
bx24, tokens = bitrix_connector.connect()

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

# Матрица перевода
transliteration_dict = {
            'А': 'A', 'К': 'K', 'Х': 'KH',
            'Б': 'B', 'Л': 'L', 'Ц': 'TS',
            'В': 'V', 'М': 'M', 'Ч': 'CH',
            'Г': 'G', 'Н': 'N', 'Ш': 'SH',
            'Д': 'D', 'О': 'O', 'Щ': 'SHCH',
            'Е': 'E', 'Ё': 'E', 'Ж': 'ZH',
            'П': 'P', 'Ъ': '',
            'Р': 'R', 'Ы': 'Y',
            'З': 'Z', 'Т': 'T', 'Э': 'E',
            'И': 'I', 'У': 'U',
            'Й': 'Y', 'Ф': 'F',
            'С': 'S', 'Ь': '',
            'Ю': 'YU', 'Я': 'YA',
}

# Словарь для шифрования и обратный
cipher_dict = {
    "0": "g",
    "1": "M",
    "2": "k",
    "3": "A",
    "4": "r",
    "5": "X",
    "6": "b",
    "7": "@",
    "8": "#",
    "9": "!"
}
reverse_cipher_dict = {v: k for k, v in cipher_dict.items()}
def encrypt_inn(inn):
    encrypted_inn = ''
    for digit in inn:
        if digit in cipher_dict:
            encrypted_inn += cipher_dict[digit]
        else:
            log.info(f"Ошибка шифрования: Символ '{digit}' присутсвует в ИНН ")
            encrypted_inn += digit
    return encrypted_inn

# Класс создания нужных атрибутов
class Person:
    def __init__(self, firstname, lastname, surname):
        self.firstname = firstname
        self.lastname = lastname
        self.surname = self.check_surname(surname)
        self.simple_login = self.create_simple_login()
        self.long_login = self.create_long_login()
        self.full_login = self.create_full_login()
        self.sm_login = self.create_sm_login()
        self.sm_login_login = self.create_sm_long_login()
        self.sm_full_login = self.create_sm_full_login()
        self.password = self.generate_password(12)

    def check_surname(self, surname):
        return surname if surname else 'Нету'

    def custom_transliterate(self, text):
        result = []
        for item in text:
            if item.upper() in transliteration_dict:
                result.append(transliteration_dict[item.upper()].upper() if item.isupper() else transliteration_dict[
                    item.upper()].lower())
            else:
                result.append(item)
        return ''.join(result)

    def create_simple_login(self):
        firstname = self.custom_transliterate(self.firstname)
        lastname = self.custom_transliterate(self.lastname)
        if firstname and lastname:
            return f"{firstname[0].lower()}.{lastname.lower()}"
        return None

    def create_long_login(self):
        firstname = self.custom_transliterate(self.firstname)
        lastname = self.custom_transliterate(self.lastname)
        if firstname and lastname:
            return f"{firstname.lower()}.{lastname.lower()}"
        return None

    def create_full_login(self):
        firstname = self.custom_transliterate(self.firstname)
        lastname = self.custom_transliterate(self.lastname)
        surname = self.custom_transliterate(self.surname)
        if firstname and lastname and surname:
            return f"{firstname.lower()}.{surname[0].lower()}.{lastname.lower()}"
        return None

    def create_email(self, login):
        return f"{login}@{connector.getAdress()}" if login else None

    def create_sm_login(self):
        firstname = self.custom_transliterate(self.firstname)
        lastname = self.custom_transliterate(self.lastname)
        if firstname and lastname:
            return f"{firstname[0].lower()}_{lastname.lower()}"
        return None

    def create_sm_long_login(self):
        firstname = self.custom_transliterate(self.firstname)
        lastname = self.custom_transliterate(self.lastname)
        if firstname and lastname:
            return f"{firstname.lower()}_{lastname.lower()}"
        return None

    def create_sm_full_login(self):
        firstname = self.custom_transliterate(self.firstname)
        lastname = self.custom_transliterate(self.lastname)
        surname = self.custom_transliterate(self.surname)
        if firstname and lastname and surname:
            return f"{firstname.lower()}_{surname[0].lower()}_{lastname.lower()}"
        return None

    def generate_password(self, length):
        lower = random.choice(string.ascii_lowercase)
        upper = random.choice(string.ascii_uppercase)
        digit = random.choice(string.digits)
        all_characters = lower + upper + digit + ''.join(
            random.choices(string.ascii_letters + string.digits, k=length - 3))
        return ''.join(random.sample(all_characters, len(all_characters)))

# Блокировка в AD
def block_ad_user(conn, user, employee, userData):
        block_attr = {
            'userAccountControl': b'514'
        }
        user_dn, user_attrs = user[0]
        for attr_name, attr_value in block_attr.items():
            if attr_name in user_attrs and user_attrs[attr_name][0] != attr_value:
                mod_attrs = [(ldap.MOD_REPLACE, attr_name, attr_value)]
                try:
                    conn.modify_s(user_dn, mod_attrs)
                    bitrix_connector.send_msg(
                        f"AD. Блокировка: Сотрудник {employee.lastname} {employee.firstname} {employee.surname}. Выполнено")
                    return True
                except Exception as e:
                    bitrix_connector.send_msg_error(
                        f"AD. Блокировка: Сотрудник {employee.lastname} {employee.firstname} {employee.surname} из отдела {userData['G2'].value} на должность {userData['J2'].value}. Не выполнено - ошибка {str(e)}")
                    return False
                finally:
                    connector.disconnect_ad(conn)

# Блокировка в BX24
def block_user_bitrix(bx24, user_id, employee, userData):
        try:
            bx24.refresh_tokens()
            result = bx24.call('user.update', {
                'ID': user_id,
                'ACTIVE': 'N'
            })
            bitrix_connector.send_msg(
                f"BX24. Блокировка: {employee.lastname, employee.firstname, employee.surname} {user_id}. Выполнено")
            return True
        except Exception as e:

            bitrix_connector.send_msg_error(f"BX24. Блокировка: {employee.lastname, employee.firstname, employee.surname} из отдела {userData['G2'].value} на должность {userData['J2'].value}. {user_id} {result}. Ошибка {e}")
            return False

# Отправка в 1с
def send_in_1c(url, data, employee, userData):
        try:
            headers = {'Content-Type': 'application/json'}
            response = requests.post(url, json=data, headers=headers)
            if response.status_code == 200:
                result = response.text
                if state == '1':
                    bitrix_connector.send_msg(f"1C. Блокировка : {employee.lastname, employee.firstname, employee.surname} Выполнено")
                    return True
                else:
                    bitrix_connector.send_msg(f"1C. Блокировка (Тест) : {employee.lastname, employee.firstname, employee.surname} Выполнено")
                    return False
            else:

                bitrix_connector.send_msg_error(f"1C. Блокировка. У сотруднка {employee.lastname, employee.firstname, employee.surname} из отдела {userData['G2'].value} на должность {userData['J2'].value}. Ошибка: {url} {data} {response.status_code}")
                return False
        except requests.exceptions.RequestException as e:
            bitrix_connector.send_msg_error(f'1C. Блокировка. У сотруднка {employee.lastname, employee.firstname, employee.surname} из отдела {userData["G2"].value} на должность {userData["J2"].value}. Ошибка: {url} {data} {response.status_code}')
            return False



def blocking_user(file_path):
    global base_dn, state, cipher_dict

    userData = load_workbook(file_path).active

    df_users = pd.read_excel(file_path)
    df_roles = pd.read_excel('info.xlsx')

    # подключение к ldap
    conn = connector.connect_ad()

    # поиск по info.xlsx
    flags = user_verification(df_roles, df_users)

    # Создание объекта сотрудника
    employee = Person(userData['C2'].value, userData['B2'].value, userData["D2"].value)

    # Зашифровка ИНН
    INN = encrypt_inn(userData['A2'].value)

    c1_success = True
    if (flags['ZUP'] or flags['RTL'] or flags['ERP']) and flags['Normal_account']:
        # Поиск друга сотрудника одной должности
        friendly = bitrix_connector.find_jobfriend(bx24,userData['J2'].value, userData['G2'].value)

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
            c1_success = send_in_1c(url, data, employee, userData)
        else:
            c1_success = True
            bitrix_connector.send_msg(
                f"1С. Блокировка (Тест): Сотрудник {employee.firstname, employee.lastname, employee.surname}. Выполнено")

    sm_success = True
    if flags['SM_GEN'] and flags['Normal_account']:
        # поиск по логинам в SM
        sm_login = sm_conn.user_exists(employee.sm_login) 
        sm_long_login = sm_conn.user_exists(employee.sm_login_login) 
        sm_full_login = sm_conn.user_exists(employee.sm_full_login) 
        if sm_login:
            try:
                if state == '1':
                    sm_success = sm_conn.block_user(employee.sm_login)
                    bitrix_connector.send_msg(f"СуперМаг Глобальный. Блокировка: {employee.lastname, employee.firstname, employee.surname} {employee.sm_login}. Выполнено")
                else:
                    bitrix_connector.send_msg(
                        f"СуперМаг Глобальный (Тест). Блокировка: {employee.lastname, employee.firstname, employee.surname} {employee.sm_login}. Выполнено")
            except Exception as e:
                sm_success = False
                bitrix_connector.send_msg_error(f'СуперМаг Глобальный. Блокировка: {employee.lastname, employee.firstname, employee.surname} из отдела {userData["G2"].value} на должность {userData["J2"].value} {employee.sm_login}. Не выполнено')
        # elif sm_long_login:
        #     try:
        #         if state == '1':
        #             sm_success = sm_conn.block_user(sm_long_login)
        #             bitrix_connector.send_msg(
        #                 f"СуперМаг Глобальный. Блокировка: {employee.lastname, employee.firstname, employee.surname} {sm_long_login}. Выполнено")
        #         else:
        #             bitrix_connector.send_msg(
        #                 f"СуперМаг Глобальный (Тест). Блокировка: {employee.lastname, employee.firstname, employee.surname} {sm_long_login}. Выполнено")
        #     except Exception as e:
        #         sm_success = False
        #         bitrix_connector.send_msg_error(f"СуперМаг Глобальный. Блокировка: {employee.lastname, employee.firstname, employee.surname} из отдела {userData['G2'].value} на должность {userData['J2'].value} {sm_long_login}. Не выполнено")
        # elif sm_full_login:
        #     try:
        #         if state == '1':
        #             sm_success = sm_conn.block_user(sm_full_login)
        #             bitrix_connector.send_msg(
        #                 f"СуперМаг Глобальный. Блокировка: {employee.lastname, employee.firstname, employee.surname} {sm_full_login}. Выполнено")
        #         else:
        #             bitrix_connector.send_msg(
        #                 f"СуперМаг Глобальный (Тест). Блокировка: {employee.lastname, employee.firstname, employee.surname} {sm_full_login}. Выполнено")
        #     except Exception as e:
        #         sm_success = False
        #         bitrix_connector.send_msg_error(f"СуперМаг Глобальный. Блокировка: {employee.lastname, employee.firstname, employee.surname} из отдела {userData['G2'].value} на должность {userData['J2'].value} {sm_full_login}. Не выполнено")
        else:
            bitrix_connector.send_msg_error(f"СуперМаг Глобальный. Блокировка: {employee.lastname, employee.firstname, employee.surname} {employee.sm_login}. Не выполнено")

    sm_local_success = True
    if flags['SM_LOCAL'] and flags['Normal_account']:
        sm_local_success = True
#        return sm_local_success

    bx24_success = True
    if flags['AD'] and flags['BX24'] and flags['Normal_account']:
        existence = connector.search_in_ad(INN)
        if len(existence) > 0:
            user_dn, attributes = existence[0]
            mail = attributes.get('mail', [b''])[0].decode('utf-8')
            user_id = bitrix_connector.search_email(bx24, mail)
            if state == '1':
                bx24.refresh_tokens()
                bx24_success = block_user_bitrix(bx24, user_id, employee, userData)
            else:
                bitrix_connector.send_msg(
                    f"BX24. Блокировка (Тест): {employee.firstname} {employee.lastname} {employee.surname}. Выполнено")
        else:
            bx24_success = True
            bitrix_connector.send_msg_error(
                f'BX24. Блокировка: {employee.firstname} {employee.lastname} {employee.surname}. Пользователь не найден в AD. Не выполнено.')
    else:
        bx24_success = True
#        return bx24_success

    # Блокировка МД АУДИТ
    md_success = True
    if flags['AD'] and flags['Normal_account']:
        if len(existence) > 0:
            user_dn, attributes = existence[0]
            mail = attributes.get('mail', [b''])[0].decode('utf-8')
            user_id = MD_AUDIT.find_user_by_email(mail)
            if user_id:
                if state == '1':
                    md_success = MD_AUDIT.block_user(user_id, employee.lastname, employee.firstname, employee.surname, userData['G2'].value, userData['J2'].value)
                else:
                    bitrix_connector.send_msg(
                        f"MD_AUDIT. Блокировка (Тест): {employee.firstname} {employee.lastname} {employee.surname}. Выполнено")
            else:
                md_success = True
#                return md_success
        else:
            md_success = True
#            return md_success
    else:
        md_success = True
#        return md_success

    ad_success = True
    if flags['AD'] and flags['Normal_account']:

        if existence:
            if state == '1':
                ad_success = block_ad_user(conn, existence, employee, userData)
            else:
                log.info(
                    f"AD. Блокировка (Тест): Сотрудник {employee.lastname, employee.firstname, employee.surname}. Выполнено")
        else:
            ad_success = False
#            return ad_success
    # else:
    #     ad_success = True
#        return ad_success
    
    if ad_success and bx24_success and c1_success and sm_success and sm_local_success and md_success:
        return True
    else:
        return False







