from openpyxl import load_workbook
import pandas as pd
import ldap
import requests
import random
import string


# Подключение файла create.py
# from actions.create import create_user

# подключение файла поиска
from outher.search import user_verification, find_jobfriend, search_in_AD,search_login

# подключение файла сообщений
from message.message import send_msg, send_msg_error,log

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
from actions.create import create_user

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

# Инициализация флагов доступа
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
def change_user(file_path):
    global base_dn, state, flags

    userData = load_workbook(file_path).active

    df_users = pd.read_excel(file_path)
    df_roles = pd.read_excel('info.xlsx')


    # Создание объекта сотрудника
    employee = Person(userData['C2'].value, userData['B2'].value, userData["D2"].value)
    company_department = userData['G2'].value
    post_job = userData['J2'].value
    phone = userData['K2'].value if userData['K2'].value is None else 'not phone'

    # поиск по info.xlsx
    flags = user_verification(df_roles, df_users)

    # Поиск друга сотрудника одной должности
    friendly = find_jobfriend(userData['J2'].value, userData['H2'].value)

    # Зашифровка ИНН
    INN = encrypt_inn(userData['A2'].value)

    # поиск по логинам в AD
    first_login = search_login(employee.simple_login, conn, base_dn)
    second_login = search_login(employee.long_login, conn, base_dn)
    tried_login = search_login(employee.full_login, conn, base_dn)

    # поиск по INN
    exists_in_AD = search_in_AD(INN, conn, base_dn)

    # Функция для изменения пользователя в 1C
    def send_in_1c(url, data):
        try:
            if state == '1':
                headers = {'Content-Type': 'application/json'}
                response = requests.post(url, json=data, headers=headers)
                if response.status_code == 200:
                    send_msg(
                            f'1С. Изменение: Сотрудник {employee.lastname, employee.firstname, employee.surname}. Выполнено')
                    return True
                else:
                    result = response.text
                    send_msg_error(f'1С. Изменение: Сотрудник {employee.lastname, employee.firstname, employee.surname}. Не выполнено - {response.status_code}')
                    log.error(f'1С. Изменение: Сотрудник {employee.lastname, employee.firstname, employee.surname}. Данные {data} отправлены, результат {result}')
                    return False
            else:
                send_msg(f'1С. Изменение (Тест): Сотрудник {employee.lastname, employee.firstname, employee.surname}. Выполнено')
                return True
        except requests.exceptions.RequestException as e:
            send_msg_error(f'1С. Изменение: Сотрудник {employee.lastname, employee.firstname, employee.surname}. Не выполнено')
            log.error(f'1С. Изменение: Ошибка {e} у сотрудника {employee.lastname, employee.firstname, employee.surname}')
            return False

    def update_ad_and_bx24():
        # флаги Обновления
        AD_update,BX24_update = False 
        
        name_atrr = {
            'sn': userData["B2"].value.encode('utf-8'),
            'givenName': userData["C2"].value.encode('utf-8'),
            'department': userData["G2"].value.encode('utf-8'),
            'division': userData["I2"].value.encode('utf-8'),
            'telephoneNumber': str(phone).encode('utf-8'),
            'company': userData["F2"].value.encode('utf-8'),
            'title': userData["J2"].value.encode('utf-8'),
        }
        new_data = {
            "NAME": userData['C2'].value,
            "LAST_NAME":  userData['B2'].value,
            "UF_DEPARTMENT": userData['H2'].value,
            "ACTIVE": "Y",
            "WORK_POSITION": userData['J2'].value,
            "PERSONAL_MOBILE": userData['K2'].value
        }

        if flags['AD'] and flags['BX24']:
            if exists_in_AD:
                user_dn, user_attrs = exists_in_AD[0]
                for attr_name, attr_value in name_atrr.items():
                    if attr_name in user_attrs and user_attrs[attr_name][0] != attr_value:
                        mod_attrs = [(ldap.MOD_REPLACE, attr_name, attr_value)]
                        if state == '1':
                            try:
                                conn.modify_s(user_dn, mod_attrs)
                                send_msg(
                                    f"AD. Изменение: Сотрудник {employee.lastname, employee.firstname, employee.surname}. Выполнено ")
                                AD_update =  False
                                return AD_update
                            except Exception as e:
                                send_msg_error(f"AD. Изменение: Сотрудник {employee.lastname, employee.firstname, employee.surname}. Не выполнено")
                                log.error(f'AD. Изменение: Ошибка при обновлении атрибута {attr_name} в домене у Сотрудника {employee.lastname, employee.firstname, employee.surname} - {e}')
                                AD_update = False
                        else:
                            send_msg(
                                f"AD. Изменение (Тест): Сотрудник {employee.lastname, employee.firstname, employee.surname}. Выполнено")
            else:
                pass

        else:
            return AD_update

        if flags['AD'] and flags['BX24']:
            if exists_in_AD:
                user_dn, user_info = exists_in_AD[0]
                id_user_bx = user_info.get("pager", [None])[0]
                if not id_user_bx or len(id_user_bx) <= 0:
                    send_msg_error(
                        f'BX24. Изменение: У сотрудника {employee.lastname, employee.firstname, employee.surname} не записан ID BX24 в атрибуите pager AD')
                    BX24_update = False
                    return BX24_update
                else:
                    if state == '1':
                        response = bx24.call('user.get', {'ID': id_user_bx.decode('utf-8')})
                        if response:
                            bx24.call('user.update', {'ID': id_user_bx.decode('utf-8'), **new_data})
                            send_msg(
                                f"BX24. Изменение: Сотрудник {employee.lastname, employee.firstname, employee.surname}. Выполнено")
                            BX24_update = True
                            return BX24_update
                        else:
                            pass
                            # send_msg_error(
                            #     f"КО Изменение: Сотрудника {employee.lastname, employee.firstname, employee.surname} с ID {id_user_bx.decode('utf-8')} не существует в BX24")
                            # BX24_update = False
                            # return BX24_update
                    else:
                        send_msg(
                            f"BX24. Изменение (Тест): Сотрудник {employee.lastname, employee.firstname, employee.surname}. Выполнено")
            else:
                pass
                # send_msg(
                #     f'КО Изменение: Cотрудник {employee.lastname, employee.firstname, employee.surname} не был найден в домене для изменения в BX24')

        if AD_update and BX24_update:
            return True
        else:
            return False

    def update_1c():
        if flags['ZUP'] or flags['RTL'] or flags['ERP']:
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
            pass


    if update_ad_and_bx24() and update_1c():
        return True
    else:
        return False


    




