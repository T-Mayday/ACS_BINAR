from openpyxl import load_workbook
import pandas as pd
import ldap
import requests
import random
import string

# Подключение файла create.py
# from actions.create import create_user

# подключение файла поиска
from outher.search import user_verification, find_jobfriend, search_in_AD, search_email_bx

# подключение файла сообщений
from message.message import send_msg, send_msg_error, log

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

# подключение к ldap
conn = connector.connect_ad()

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

# Подключение файла create.py
# from actions.create import create_user


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


def change_user(file_path):
    global base_dn, state


    userData = load_workbook(file_path).active

    df_users = pd.read_excel(file_path)
    df_roles = pd.read_excel('info.xlsx')

    # Создание объекта сотрудника
    employee = Person(userData['C2'].value, userData['B2'].value, userData["D2"].value)

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
        "NAME": userData['C2'].value,
        "LAST_NAME": userData['B2'].value,
        "UF_DEPARTMENT": userData['H2'].value,
        "ACTIVE": "Y",
        "WORK_POSITION": userData['J2'].value
    }


    def update_ad_attributes(conn, user, new_atrr):
        user_dn, user_attrs = user[0]
        for attr_name, attr_value in new_atrr.items():
            if attr_name in user_attrs and user_attrs[attr_name][0] != attr_value:
                mod_attrs = [(ldap.MOD_REPLACE, attr_name, attr_value)]
                try:
                    conn.modify_s(user_dn, mod_attrs)
                    log.info(
                        f"AD. Изменение: Сотрудник {employee.lastname}, {employee.firstname}, {employee.surname}. Обновление атрибута {attr_name}. Выполнено"
                    )
                    return True
                except Exception as e:
                    log.info(
                        f"AD. Изменение: Сотрудник {employee.lastname}, {employee.firstname}, {employee.surname}. Не выполнено. Ошибка при обновлении атрибута {attr_name}: {str(e)}"
                    )
                    return False

    def bitrix_call(user_id, new_data):
        try:
            bx24.refresh_tokens()
            result = bx24.call('user.update', {'ID': user_id, **new_data})
            send_msg(
                f"BX24. Изменение: Сотрудник {employee.lastname, employee.firstname, employee.surname} из отдела {userData['G2'].value} на должность {userData['J2'].value}. {result}. Выполнено")
            return True
        except Exception as e:
            send_msg_error(

                f"BX24. Изменение: Сотрудник {employee.lastname, employee.firstname, employee.surname} из отдела {userData['G2'].value} на должность {userData['J2'].value}. Ошибка при изменение пользователя в Битрикс24: {e}")


            return False

    # Функция для изменения пользователя в 1C
    def send_in_1c(url, data):
        try:
            if state == '1':
                headers = {'Content-Type': 'application/json'}
                response = requests.post(url, json=data, headers=headers)
                if response.status_code == 200:
                    send_msg(
                        f"1С. Изменение: Сотрудник {employee.lastname, employee.firstname, employee.surname}. Выполнено")
                    return True
                else:
                    result = response.text
                    send_msg_error(

                        f"1С. Изменение: Сотрудник {employee.lastname, employee.firstname, employee.surname} из отдела {userData['G2'].value} на должность {userData['J2'].value}. Не выполнено. Данные {data} отправлены, результат {response.status_code} {result}")
                    return False
            else:
                send_msg(
                    f"1С. Изменение (Тест): Сотрудник {employee.lastname, employee.firstname, employee.surname}. Выполнено")
                return True
        except requests.exceptions.RequestException as e:
            send_msg_error(

                f"1С. Изменение: Сотрудник {employee.lastname, employee.firstname, employee.surname} из отдела {userData['G2'].value} на должность {userData['J2'].value}. Не выполнено. Ошибка {e}")

            return False

    ad_success = False
    if flags['AD'] and flags['Normal_account']:
        simple_email = search_in_AD(employee.create_email(employee.simple_login), conn, base_dn)
        long_email = search_in_AD(employee.create_email(employee.long_login), conn, base_dn)
        full_email = search_in_AD(employee.create_email(employee.full_login), conn, base_dn)

        if simple_email and len(simple_email) > 0:
            if state == '1':
                ad_success = update_ad_attributes(conn, simple_email, name_atrr)
            else:
                send_msg(
                    f"AD. Изменение (Тест): Сотрудник {employee.lastname}, {employee.firstname}, {employee.surname}. Выполнено"
                )
        elif long_email and len(long_email) > 0:
            if state == '1':
                ad_success = update_ad_attributes(conn, long_email, name_atrr)
            else:
                send_msg(
                    f"AD. Изменение (Тест): Сотрудник {employee.lastname}, {employee.firstname}, {employee.surname}. Выполнено"
                )
        elif full_email and len(full_email) > 0:
            if state == '1':
                ad_success = update_ad_attributes(conn, full_email, name_atrr)
            else:
                send_msg(
                    f"AD. Изменение (Тест): Сотрудник {employee.lastname}, {employee.firstname}, {employee.surname}. Выполнено"
                )
        else:
            send_msg(f'AD.Изменение: Не нашел сотрудника {employee.lastname}, {employee.firstname}, {employee.surname} ему следует создать аккаунт.')
            ad_success = False
            return ad_success
    else:
        ad_success = True
        return ad_success

    bx_success = False
    if flags['AD'] and flags['BX24'] and flags['Normal_account']:
        if len(simple_email) > 0:
            user_dn, user_info = simple_email[0]
            email_ad = user_info.get('mail', [None])[0]
            ID_BX24 = search_email_bx(email_ad.decode('utf-8'))
            if ID_BX24 and state == '1':
                bx_success = bitrix_call(ID_BX24.decode('utf-8'), new_data)
            else:
                send_msg(
                    f"BX24. Изменение (Тест): Сотрудник {employee.lastname, employee.firstname, employee.surname}. Выполнено")
        elif len(long_email) > 0:
            user_dn, user_info = long_email[0]
            email_ad = user_info.get('mail', [None])[0]
            ID_BX24 = search_email_bx(email_ad.decode('utf-8'))

            if ID_BX24 and state == '1':
                bx_success = bitrix_call(ID_BX24.decode('utf-8'), new_data)
            else:
                send_msg(
                    f"BX24. Изменение (Тест): Сотрудник {employee.lastname, employee.firstname, employee.surname}. Выполнено")
        elif len(full_email) > 0:
            user_dn, user_info = full_email[0]
            email_ad = user_info.get('mail', [None])[0]
            ID_BX24 = search_email_bx(email_ad.decode('utf-8'))

            if ID_BX24 and state == '1':
                bx_success = bitrix_call(ID_BX24.decode('utf-8'), new_data)
            else:
                send_msg(
                    f"BX24. Изменение (Тест): Сотрудник {employee.lastname, employee.firstname, employee.surname}. Выполнено")
        else:
            bx_success = False
            return bx_success
    else:
        bx_success = True
        return bx_success

    def update_1c():
        if flags['ZUP'] or flags['RTL'] or flags['ERP'] and flags['Normal_account']:
            # Поиск друга сотрудника одной должности
            friendly = find_jobfriend(userData['J2'].value, userData['H2'].value)

            ZUP_value, RTL_value, ERP_value = (
                1 if flags['ZUP'] else 0, 1 if flags['RTL'] else 0, 1 if flags['ERP'] else 0)

            url = connector_1c.getUrlChanges()
            data = {
                'full_name': f'{userData["B2"].value} {userData["C2"].value} {userData["D2"].value}',
                'ERP': ERP_value,
                'RTL': RTL_value,
                'ZUP': ZUP_value,
                'job_friend': friendly
            }
            return send_in_1c(url, data)
        else:
            return True

    if ad_success and bx_success and update_1c():
        return True
    else:
        return False






