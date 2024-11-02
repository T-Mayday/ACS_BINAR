from openpyxl import load_workbook
import pandas as pd
import ldap
import requests
import random
import string

# Подключение файла create.py
from actions.create import create_user, create_in_BX24

# подключение файла поиска
from outher.search import user_verification

# подключение файла сообщений
from message.message import log

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


# Функция для шифрование ИНН
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


# Обновление в AD
def update_ad_attributes(conn, user, new_atrr, employee):
        user_dn, user_attrs = user[0]
        success = True
        for attr_name, attr_value in new_atrr.items():
            if attr_name in user_attrs and user_attrs[attr_name][0] != attr_value:
                mod_attrs = [(ldap.MOD_REPLACE, attr_name, attr_value)]
                try:
                    conn.modify_s(user_dn, mod_attrs)
                    bitrix_connector.send_msg(
                        f"AD. Изменение: Сотрудник {employee.lastname} {employee.firstname} {employee.surname}. Обновление атрибута {attr_name}. Выполнено"
                    )
                except Exception as e:
                    bitrix_connector.send_msg_error(
                        f"AD. Изменение: Сотрудник {employee.lastname} {employee.firstname} {employee.surname}. Не выполнено. Ошибка при обновлении атрибута {attr_name}: {str(e)}"
                    )
                    success = False
                finally:
                    connector.disconnect_ad(conn)
        return success

# Обновление в BX24
def bitrix_call(bx24, user_id, new_data, employee, userData):
        try:
            bx24.refresh_tokens()
            result = bx24.call('user.update', {'ID': user_id, **new_data})
            if result.get('error'):
                bitrix_connector.send_msg_error(
                    f"BX24. Ошибка при изменении пользователя: Сотрудник {employee.lastname} {employee.firstname} {employee.surname} из отдела {userData['G2'].value} на должность {userData['J2'].value}. Ошибка: {result.get('error_description')}")
                success = False
            else:
                bitrix_connector.send_msg(
                    f"BX24. Изменение: Сотрудник {employee.lastname} {employee.firstname} {employee.surname} из отдела {userData['G2'].value} на должность {userData['J2'].value}. Выполнено")
                success = True
        except Exception as e:
            bitrix_connector.send_msg_error(f"BX24. Изменение: Сотрудник {employee.lastname} {employee.firstname} {employee.surname} из отдела {userData['G2'].value} на должность {userData['J2'].value}. Ошибка при изменение пользователя в Битрикс24: {e}")
            success = False
        return success

# Отправка в 1с
def send_in_1c(url, data, employee, userData):
        try:
            if state == '1':
                headers = {'Content-Type': 'application/json'}
                response = requests.post(url, json=data, headers=headers)
                if response.status_code == 200:
                    bitrix_connector.send_msg(
                        f"1С. Изменение: Сотрудник {employee.lastname, employee.firstname, employee.surname}. Выполнено")
                    return True
                else:
                    result = response.text
                    bitrix_connector.send_msg_error(
                        f"1С. Изменение: Сотрудник {employee.lastname, employee.firstname, employee.surname} из отдела {userData['G2'].value} на должность {userData['J2'].value}. Не выполнено. Данные {data} отправлены {url}, результат {response.status_code} {result}")
                    return False
            else:
                bitrix_connector.send_msg(
                    f"1С. Изменение (Тест): Сотрудник {employee.lastname, employee.firstname, employee.surname}. Выполнено")
                return True
        except requests.exceptions.RequestException as e:
            bitrix_connector.send_msg_error(
                f"1С. Изменение: Сотрудник {employee.lastname, employee.firstname, employee.surname} из отдела {userData['G2'].value} на должность {userData['J2'].value}. Не выполнено. Ошибка {e}")

            return False

def change_user(file_path):
    global base_dn, state

    userData = load_workbook(file_path).active

    df_users = pd.read_excel(file_path)
    df_roles = pd.read_excel('info.xlsx')

    # Создание объекта сотрудника
    employee = Person(userData['C2'].value, userData['B2'].value, userData["D2"].value)

    # Подключение к ldap
    conn = connector.connect_ad()

    # Подключение к битрикс
    bx24, tokens = bitrix_connector.connect()

    # поиск по info.xlsx
    flags = user_verification(df_roles, df_users)

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

    def update_1c():
        if (flags['ZUP'] or flags['RTL'] or flags['ERP']) and flags['Normal_account']:
            existence = connector.search_in_ad(INN)
            if not (existence is None) and len(existence) > 0:
                user_dn, attributes = existence[0]
                cn = attributes.get('cn', [b''])[0].decode('utf-8')
                login = f"\\BINLTD\{cn}"
                # Поиск друга сотрудника одной должности
                friendly = bitrix_connector.find_jobfriend(bx24,userData['J2'].value, userData['H2'].value)

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
                return send_in_1c(url, data, employee, userData)
        else:
            return True

    ad_success = False
    if flags['AD'] and flags['Normal_account']:
        existence = connector.search_in_ad(INN)
        if existence:
            if state == '1':
                ad_success = update_ad_attributes(conn, existence, name_atrr, employee)
            else:
                bitrix_connector.send_msg(
                    f"AD. Изменение (Тест): Сотрудник {employee.lastname} {employee.firstname} {employee.surname}. Выполнено"
                )
        else:
            ad_success = create_user(file_path)
#            return ad_success
    else:
        ad_success = True
#        return ad_success

    bx_success = False
    if flags['AD'] and flags['BX24'] and flags['Normal_account']:
        existence = connector.search_in_ad(INN)
        if len(existence) > 0:
            user_dn, attributes = existence[0]
            mail = attributes.get('mail', [b''])[0].decode('utf-8')

            user_info = bitrix_connector.search_email(bx24, mail)
            if user_info:
                if state == '1':
                    bx_success = bitrix_call(bx24, user_info.get('ID'), new_data, employee, userData)
                    # bitrix_connector.send_msg(
                    #     f"BX24. Изменение: Сотрудник {employee.lastname} {employee.firstname} {employee.surname} {user_info.get('ID')}. Выполнено")
#                    return bx_success
                else:
                    bitrix_connector.send_msg(
                        f"BX24. Изменение (Тест): Сотрудник {employee.lastname} {employee.firstname} {employee.surname} {user_info.get('ID')}. Выполнено")
            else:
                bx_success = create_in_BX24(mail, bx24, employee, userData, conn)
        else:
            user_info = bitrix_connector.search_user(bx24, employee.lastname, employee.firstname, employee.surname)
            if user_info:
                if state == '1':
                    bx_success = bitrix_call(bx24, user_info.get('ID'), new_data, employee, userData)
#                    return bx_success
                else:
                    bitrix_connector.send_msg(
                        f"BX24. Изменение (Тест): Сотрудник {employee.lastname} {employee.firstname} {employee.surname} {user_info.get('ID')}. Выполнено")
            else:
                bitrix_connector.send_msg_error(
                    f"BX24. Изменение: Сотрудник {employee.lastname} {employee.firstname} {employee.surname} не найден.")
#                bx_success = create_user(file_path)
#                return bx_success
    else:
        bx_success = True
#        return bx_success
    
    if ad_success and bx_success and update_1c():
        return True
    else:
        return False






