from openpyxl import load_workbook
import pandas as pd
import ldap
import requests
import random
import string


# подключение файла поиска
from outher.search import user_verification, search_in_AD, find_jobfriend

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

# Подключение к SM
from connect.SMConnect import SMConnect

sm_conn = SMConnect()
sm_conn.connect_SM()
test_role_id = sm_conn.getRoleID()



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
                result.append(transliteration_dict[item.upper()].upper() if item.isupper() else transliteration_dict[item.upper()].lower())
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

    def create_email(self, connector, login):
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
        all_characters = lower + upper + digit + ''.join(random.choices(string.ascii_letters + string.digits, k=length - 3))
        return ''.join(random.sample(all_characters, len(all_characters)))


flags = {
        'AD': False,
        'BX24': False,
        'ZUP': False,
        'RTL': False,
        'ERP': False,
        'SM_GEN': False,
        'SM_LOCAL': False,
        'Normal_account': False,
        'Shop_account': False
    }


def blocking_user(file_path):
    global flags, base_dn, state, cipher_dict, reverse_cipher_dict, name_domain

    userData = load_workbook(file_path).active

    df_users = pd.read_excel(file_path)
    df_roles = pd.read_excel('info.xlsx')

    # подключение к ldap
    conn = connector.connect_ad()
    if not (bx24 and conn):
        send_msg('Блокировка.Ошибка подключения к AD и BX24')

    # поиск по info.xlsx
    flags = user_verification(df_roles, df_users)

    # Создание объекта сотрудника
    employee = Person(userData['C2'].value, userData['B2'].value, userData["D2"].value)

    # Поиск друга сотрудника одной должности
    friendly = find_jobfriend(userData['J2'].value, userData['H2'].value)

    # Зашифровка ИНН
    INN = encrypt_inn(userData['A2'].value)

    # поиск по INN
    exists_in_AD = search_in_AD(INN, conn,base_dn)


    # поиск по логинам в SM
    sm_login = sm_conn.user_exists(employee.sm_login) == -1
    sm_long_login = sm_conn.user_exists(employee.sm_login_login) == -1
    sm_full_login = sm_conn.user_exists(employee.sm_full_login) == -1

    name_atrr = {
        'userAccountControl': b'514'
    }
    new_data = {
        "ACTIVE": "N",
    }


    # Функция для создания пользователя в 1C
    def send_in_1c(url, data):
        try:
            headers = {'Content-Type': 'application/json'}
            response = requests.post(url, json=data, headers=headers)
            if response.status_code == 200:
                result = response.text
                if state == '1':
                    send_msg(f'1C. Блокировка : Сотрудник {employee.lastname, employee.firstname, employee.surname} Выполнено')
                    return result
                else:
                    send_msg(f'1C. Блокировка (Тест) : Сотрудник {employee.lastname, employee.firstname, employee.surname} Выполнено')
            else:
                return send_msg_error(f'1C. Блокировка. Ошибка : {response.status_code}')
        except requests.exceptions.RequestException as e:
            return send_msg_error(f'1C. Блокировка. Ошибка: {e}')


    if flags['AD'] and flags['BX24']:
        if exists_in_AD:
            user_dn, user_attrs = exists_in_AD[0]
            for attr_name, attr_value in name_atrr.items():
                if attr_name in user_attrs and user_attrs[attr_name][0] != attr_value:
                    mod_attrs = [(ldap.MOD_REPLACE, attr_name, attr_value)]
                    if state == '1':
                        conn.modify_s(user_dn, mod_attrs)
                        send_msg(
                            f"AD. Блокировка: Сотрудник {employee.lastname, employee.firstname, employee.lastname}. Выполнено")
                    else:
                        send_msg(
                            f"AD. Блокировка (Тест): Сотрудник {employee.lastname, employee.firstname, employee.lastname}. Выполнено")


    if flags['AD'] and flags['BX24']:
        if exists_in_AD:
            user_dn, user_info = exists_in_AD[0]
            id_user_bx = user_info.get("pager", [None])[0]
            if state == '1':
                response = bx24.call('user.get', {'ID': id_user_bx.decode('utf-8')})
                if response:
                    bx24.call('user.update', {'ID': id_user_bx.decode('utf-8'), **new_data, })
                    send_msg(
                        f"BX24. Блокировка: Сотрудник {employee.lastname, employee.firstname, employee.lastname}. Выполнено")
                else:
                    send_msg_error(f"BX24. Блокировка: Сотрудник {employee.lastname, employee.firstname, employee.lastname}. Не выполнено.")
                    # log.error(f"BX24. Блокировка: У сотрудника {employee.lastname, employee.firstname, employee.lastname} ID {id_user_bx.decode('utf-8')} не найден в Битрикс24.")
            else:
                send_msg(
                    f"BX24. Блокировка: Сотрудник {employee.lastname, employee.firstname, employee.lastname}. Выполнено")
        else:
            send_msg_error(
                f'BX24. Блокировка: Сотрудник {employee.lastname, employee.firstname, employee.lastname}. Не выполнено')
            # log.error(f'BX24 и AD. Блокировка: Сотрудник {employee.lastname, employee.firstname, employee.lastname}. Поиск не нашел в домене сотрудника')




    if flags['ZUP'] or flags['RTL'] or flags['ERP']:

        ZUP_value, RTL_value, ERP_value = (1 if flags['ZUP'] else 0, 1 if flags['RTL'] else 0, 1 if flags['ERP'] else 0)

        url = connector_1c.getUrlBlock()
        data = {
            'full_name': f'{userData["B2"].value} {userData["C2"].value} {userData["D2"].value}',
            'ERP': ERP_value,
            'RTL': RTL_value,
            'ZUP': ZUP_value,
            'job_friend': friendly
        }
        send_in_1c(url,data)


    if flags['SM_GEN']:
        if sm_login:
            try:
                if state == '1':
                    sm_conn.block_user(sm_login)
                    send_msg(f'СуперМаг Глобальный. Блокировка: Сотрудник {employee.lastname, employee.firstname, employee.lastname}. Выполнено')
                else:
                    send_msg(
                        f'СуперМаг Глобальный (Тест). Блокировка: Сотрудник {employee.lastname, employee.firstname, employee.lastname}. Выполнено')
            except Exception as e:
                send_msg_error(f'СуперМаг Глобальный. Блокировка: Сотрудник {employee.lastname, employee.firstname, employee.lastname}. Не выполнено')
                # log.error(f'СуперМаг Глобальный: Ошибка при блокировке у сотрудника {employee.lastname, employee.firstname, employee.lastname}. Ошибка {str(e)}')


        elif sm_long_login:
            try:
                if state == '1':
                    sm_conn.block_user(sm_long_login)
                    send_msg(
                        f'СуперМаг Глобальный. Блокировка: Сотрудник {employee.lastname, employee.firstname, employee.lastname}. Выполнено')
                else:
                    send_msg(
                        f'СуперМаг Глобальный (Тест). Блокировка: Сотрудник {employee.lastname, employee.firstname, employee.lastname}. Выполнено')
            except Exception as e:
                send_msg_error(
                    f'СуперМаг Глобальный. Блокировка: Сотрудник {employee.lastname, employee.firstname, employee.lastname}. Не выполнено')
                # log.error(
                #     f'СуперМаг Глобальный: Ошибка при блокировке у сотрудника {employee.lastname, employee.firstname, employee.lastname}. Ошибка {str(e)}')

        elif sm_full_login:
            try:
                if state == '1':
                    sm_conn.block_user(sm_full_login)
                    send_msg(
                        f'СуперМаг Глобальный. Блокировка: Сотрудник {employee.lastname, employee.firstname, employee.lastname}. Выполнено')
                else:
                    send_msg(
                        f'СуперМаг Глобальный (Тест). Блокировка: Сотрудник {employee.lastname, employee.firstname, employee.lastname}. Выполнено')
            except Exception as e:
                send_msg_error(
                    f'СуперМаг Глобальный. Блокировка: Сотрудник {employee.lastname, employee.firstname, employee.lastname}. Не выполнено')
                # log.error(
                #     f'СуперМаг Глобальный: Ошибка при блокировке у сотрудника {employee.lastname, employee.firstname, employee.lastname}. Ошибка {str(e)}')
        else:
            send_msg_error(f'СуперМаг Глобальный. Блокировка: Сотрудник {employee.lastname, employee.firstname, employee.lastname}. Не выполнено')
            # log.error(f'Поиск выдал что не по одному из логинов у сотрудника {employee.lastname, employee.firstname, employee.lastname} на должности {userData["G2"].value, userData["J2"].value} не нашел')








