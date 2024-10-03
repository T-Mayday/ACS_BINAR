import time
from openpyxl import load_workbook, Workbook
import pandas as pd
import ldap
import ldap.modlist as modlist
import requests
import random
import string

# подключение файла поиска
from outher.search import user_verification, search_in_AD, search_login, find_jobfriend
# подключение файла сообщений
from message.message import send_msg, send_msg_error, send_msg_adm, log
# Подключение BitrixConnect
from connect.bitrixConnect import Bitrix24Connector
bitrix_connector = Bitrix24Connector()
bx24, tokens = bitrix_connector.connect()
# Подключение connect1c
from connect.connect1C import Connector1C
connector_1c = Connector1C()
# Подключение ldapConnect
from connect.ldapConnect import ActiveDirectoryConnector
connector = ActiveDirectoryConnector()
base_dn = connector.getBaseDn()
state = connector.getState()
name_domain = connector.domain_name
# Подключение к SM
from connect.SMConnect import SMConnect
sm_conn = SMConnect()
sm_conn.connect_SM()
test_role_id = sm_conn.getRoleID()
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


# Функция для расшифрование ИНН
def decrypt_inn(encrypted_inn):
    decrypted_inn = ''
    for char in encrypted_inn:
        if char in reverse_cipher_dict:
            decrypted_inn += reverse_cipher_dict[char]
        else:
            log.info(f"Ошибка: Цифра '{char}' присутсвует в ИНН ")
            decrypted_inn += char
    return decrypted_inn

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
        self.full_name = self.full_name()

    def full_name(self):
        return self.lastname + '' + self.firstname + '' + self.surname

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
        all_characters = lower + upper + digit + ''.join(random.choices(string.ascii_letters + string.digits, k=length - 3))
        return ''.join(random.sample(all_characters, len(all_characters)))


def create_user(file_path):
    global base_dn, state, cipher_dict, reverse_cipher_dict, name_domain, base_dn, conn

    userData = load_workbook(file_path).active

    df_users = pd.read_excel(file_path)
    df_roles = pd.read_excel('info.xlsx')

    # поиск по info.xlsx
    flags = user_verification(df_roles, df_users)

    # Создание объекта сотрудника
    employee = Person(userData['C2'].value, userData['B2'].value, userData["D2"].value)
    phone = userData['K2'].value if not (userData['K2'].value is None) else ''

    # Зашифровка ИНН
    INN = encrypt_inn(userData['A2'].value)

    ad_success = False
    bx24_success = False
    c1_success = False
    sm_success = False
    sm_local_success = False

    # Функция записи логина и пароля
    def save_login(phone_number, full_name , login):
        send_msg_adm(f"{phone_number} {full_name} {login}")
        file_path = 'logins.xlsx'
        try:
            workbook = load_workbook(file_path)
            sheet = workbook.active
        except FileNotFoundError:
            workbook = Workbook()
            sheet = workbook.active
            sheet.append(['Номер телефона', 'ФИО', 'Логин'])
        sheet.append([phone_number, full_name, login])
        workbook.save(file_path)
        workbook.close()

    # Функция для создания пользователя в AD
    def create_in_AD(login_type):
        try:
            dn = connector.getNewUserDn().format(login_type)
            user_dn = f"CN={login_type},{dn}"
            attrs = [
                ('objectClass', [b'top', b'person', b'organizationalPerson', b'user']),
                ('cn', [login_type.encode('utf-8')]),
                ('givenName', [employee.firstname.encode('utf-8')]),
                ('sAMAccountName', [login_type.encode('utf-8')]),
                ('userPrincipalName', [login_type.encode('utf-8') + b'@BINLTD.local']),
#                ('displayName', [login_type.encode('utf-8')]),
                ('displayName', [str(f"{employee.lastname} {employee.firstname} {employee.surname}").encode('utf-8')]),
                ('department', [userData['G2'].value.encode('utf-8')]),
                ('mail', [employee.create_email(login_type).encode('utf-8')]),
                ('sn', [employee.lastname.encode('utf-8')]),
                ('employeeID', [str(INN).encode('utf-8')]),
                ('company', [userData['F2'].value.encode('utf-8')]),
                ('userAccountControl', [b'512']),
                ('middleName', [employee.surname.encode('utf-8')]),
                ('title', [userData['J2'].value.encode('utf-8')]),
                ('userPassword', [employee.password.encode()]),
                ('pwdLastSet', [b'0']),
            ]
            if state == "1":
#                ldif = modlist.addModlist(attrs)
                conn.add_s(user_dn, attrs)
                created = search_in_AD(INN, conn, base_dn)
                save_login(phone,employee.full_name,login_type)
                if len(created) != 0:
                    send_msg(
                        f"AD. Создание: Сотруднику {employee.firstname, employee.lastname, employee.surname} {user_dn}. Выполнено")
                    return True
                else:
                    send_msg_error(
                        f"AD. Создание: Сотрудник {employee.firstname, employee.lastname, employee.surname} из отдела {userData['H2'].value} на должность {userData['J2'].value}. {user_dn}. Не выполнено")
                    return False
            else:
                send_msg(
                    f"AD. Создание (Tест): Сотруднику {employee.firstname, employee.lastname, employee.surname} {user_dn}. Выполнено")
                return True
        except Exception as e:
            send_msg_error(f'AD. Создание: Сотрудник {employee.firstname, employee.lastname, employee.surname} из отдела {userData['H2'].value} на должность {userData['J2'].value}. Ошибка при создании {user_dn} {attrs} {str(e)}')
            return False

    # Функция для создания пользователя в BX24
    def create_in_BX24(email):
        try:
            user_data = {
                "NAME": employee.firstname,
                "LAST_NAME": employee.lastname,
                "SECOND_NAME": employee.surname,
                "EMAIL": email,
                "UF_DEPARTMENT": str(userData['H2'].value),
                "ACTIVE": "Y",
                "WORK_POSITION": str(userData['J2'].value),
            }
            if state == '1':
                bx24.refresh_tokens()
                createBX = bx24.call('user.add', user_data)

                if createBX.get('error'):
                    error_message = createBX.get('error_description')
                    send_msg_error(
                        f"BX24. Создание: Сотрудник {employee.firstname, employee.lastname, employee.surname} из отдела {userData['H2'].value} на должность {userData['J2'].value}. Не выполнено. {user_data} {error_message}")
                    return False

                if createBX.get('result'):
                    user_id = createBX.get('result')
                    search_filter = f"(employeeID={INN})"
                    search_base = connector.getSearchBase()
                    result = conn.search_s(search_base, ldap.SCOPE_SUBTREE, search_filter)

                    if result:
                        user_dn, user_attrs = result[0]
                        attr = [(ldap.MOD_REPLACE, 'pager', str(user_id).encode('utf-8'))]
                    if state == "1":
                        conn.modify_s(user_dn, attr)
                    send_msg(
                        f"BX24. Создание: Сотрудник {employee.firstname, employee.lastname, employee.surname} ID={user_id}. Выполнено")
                    return True
                else:
                    send_msg_error(
                        f"BX24. Создание: Сотрудник {employee.firstname, employee.lastname, employee.surname}. Ошибка: {createBX.get('result')}")
                    return False
            else:
                send_msg(
                    f"BX24. Создание (Тест): Сотрудник {employee.firstname, employee.lastname, employee.surname}. Выполнено")
                return True
        except Exception as e:
            send_msg_error(f'BX24. Создание: Сотрудник {employee.firstname, employee.lastname, employee.surname}. Ошибка:{e}')
            return False

    def send_in_1c(url, data):
        try:
            if state == '1':
                headers = {'Content-Type': 'application/json'}
                response = requests.post(url, json=data, headers=headers)
                if response.status_code == 200:
                    result = response.text
                    send_msg(
                        f"1С. Создание: Сотрудник {employee.firstname, employee.lastname, employee.surname}. {response.status_code} Выполнено")
                    return True
                else:
                    result = response.text
                    send_msg_error(f'1С. Создание: Сотрудник {employee.firstname, employee.lastname, employee.surname} из отдела {userData['H2'].value} на должность {userData['J2'].value}. Не выполнено. Ошибки - {response.status_code} {url} {data}')
                    return False
            else:
                send_msg(
                    f'1С. Создание: Сотрудник(Тест): Сотрудник {employee.lastname, employee.firstname, employee.surname}. Выполнено')
                return True
        except requests.exceptions.RequestException as e:
            send_msg_error(f'1С. Создание: Сотрудник {employee.lastname, employee.firstname, employee.surname}. Ошибка {url} {data} Error: {e}')
            return False

    # Получение название локальной базы
    def get_storeId():
        stores_id = userData['N2'].value
        if stores_id is None:
            sm_local_success = False
            return sm_local_success

        if isinstance(stores_id, str):
            stores = [int(num.strip()) for num in stores_id.split(',') if num.strip().isdigit()]
            store_names = []
            for store in stores:
                name_store = sm_conn.get_store(store)
                if name_store:
                    store_names.append(name_store)
            return store_names
        return []

    store_names = get_storeId()

    # Основная логика
    if flags['AD'] and flags['Normal_account']:
        # поиск по INN
        exists_in_AD = search_in_AD(INN, conn, base_dn)
        if len(exists_in_AD) == 0:
            # поиск по логинам в AD

            first_login = search_login(employee.simple_login, conn, base_dn)
            second_login = search_login(employee.long_login, conn, base_dn)
            tried_login = search_login(employee.full_login, conn, base_dn)

            if len(first_login) == 0:
                try:
                    ad_success = create_in_AD(employee.simple_login)
                except Exception as e:
                    send_msg_error(
                        f'AD. Создание: Ошибка при создании первичного логина у сотрудника {employee.firstname, employee.lastname, employee.surname} из отдела {userData['H2'].value} на должность {userData['J2'].value}.Ошибка {e}')
            elif len(second_login) == 0:
                try:
                    ad_success = create_in_AD(employee.long_login)
                except Exception as e:
                    send_msg_error(
                        f'AD. Создание: Ошибка при создании вторичного логина у сотрудника {employee.firstname, employee.lastname, employee.surname} из отдела {userData['H2'].value} на должность {userData['J2'].value}.Ошибка {e}')
            elif len(tried_login) == 0:
                try:
                    ad_success = create_in_AD(employee.full_login)
                except Exception as e:
                    send_msg_error(
                        f'AD. Создание: Ошибка при создании третичного логина у сотрудника {employee.firstname, employee.lastname, employee.surname} из отдела {userData['H2'].value} на должность {userData['J2'].value}.Ошибка {e}')
        else:
            send_msg_error(
                f'AD. Создание: У сотрудника {employee.firstname, employee.lastname, employee.surname} из отдела {userData['H2'].value} на должность {userData['J2'].value}. Поиск по ИНН выдал что такой пользователь уже существует в AD ')
    else:
        ad_success = True
        return ad_success

    if flags['AD'] and flags['BX24'] and flags['Normal_account']:
        # поиск по INN
        exists_in_AD = search_in_AD(INN, conn, base_dn)
    
        if exists_in_AD:
            user_dn, user_info = exists_in_AD[0]
            id_user_bx = user_info.get("pager", [None])[0]
            if not id_user_bx or len(id_user_bx) <= 0:
                try:
                    bx24_success = create_in_BX24(employee.create_email(employee.simple_login))
                except Exception as e:
                    send_msg_error(
                        f'BX24. Создание: Ошибка при создании первичного логина у сотрудника {employee.firstname, employee.lastname, employee.surname} из отдела {userData['H2'].value} на должность {userData['J2'].value} Ошибка {e}')
        else:
            send_msg_error(
                f'BX24. Создание: У сотрудника {employee.firstname, employee.lastname, employee.surname} из отдела {userData['H2'].value} на должность {userData['J2'].value}. Пользователь не найден в AD')
    else:
        bx24_success = True
        return bx24_success

    if flags['ZUP'] or flags['RTL'] or flags['ERP'] and flags['Normal_account']:

        # Поиск друга сотрудника одной должности
        friendly = find_jobfriend(userData['J2'].value, userData['H2'].value)

        ZUP_value, RTL_value, ERP_value = (
            1 if flags['ZUP'] else 0, 1 if flags['RTL'] else 0, 1 if flags['ERP'] else 0)

        url = connector_1c.getUrlCreate()
        data = {
            'full_name': f'{employee.lastname} {employee.firstname} {employee.surname}',
            'password': 'qwerty32',
            'domain': 'binltd.local',
            'ERP': ERP_value,
            'RTL': RTL_value,
            'ZUP': ZUP_value,
            'job_friend': friendly
        }

        if state == '1':
            c1_success = send_in_1c(url, data)
        else:
            c1_success = True
            send_msg(
                f"1С. Создание (Тест): Сотрудник {employee.firstname, employee.lastname, employee.surname}. Выполнено")
    else:
        c1_success = True
        return c1_success

    if flags['SM_GEN'] and flags['Normal_account']:
        # поиск по логинам в SM
        sm_login = sm_conn.user_exists(employee.sm_login) == -1
        sm_long_login = sm_conn.user_exists(employee.sm_login_login) == -1
        sm_full_login = sm_conn.user_exists(employee.sm_full_login) == -1
        if sm_login:
            try:
                sm_success = sm_conn.create_user(sm_login, employee.password, test_role_id)
                send_msg(f'СуперМаг Глобальный. Создание: Сотрудник {employee.firstname, employee.lastname, employee.surname}. Выполнено')

            except Exception as e:
                send_msg_error(
                    f'СуперМаг Глобальный. Создание: Ошибка при создании первичного логина в SM у сотрудника {employee.firstname, employee.lastname, employee.surname} из отдела {userData['H2'].value} на должность {userData['J2'].value} - {e}')
        elif sm_long_login:
            try:
                sm_success = sm_conn.create_user(sm_long_login, employee.password, test_role_id)
                send_msg(f'СуперМаг Глобальный. Создание: Сотрудник {employee.firstname, employee.lastname, employee.surname}. Выполнено')

            except Exception as e:
                send_msg_error(
                    f'СуперМаг Глобальный. Создание: Ошибка при создании вторичного логина в SM у сотрудника {employee.firstname, employee.lastname, employee.surname} из отдела {userData['H2'].value} на должность {userData['J2'].value} - {e}')

        elif sm_full_login:
            try:
                sm_success = sm_conn.create_user(sm_full_login, employee.password, test_role_id)
                send_msg(f'СуперМаг Глобальный. Создание: Сотрудник {employee.firstname, employee.lastname, employee.surname}. Выполнено')

            except Exception as e:
                send_msg_error(
                    f'СуперМаг Глобальный.Создание: Ошибка при создании третичного логина в SM у сотрудника {employee.firstname, employee.lastname, employee.surname} из отдела {userData['H2'].value} на должность {userData['J2'].value} - {e}')
        else:
            send_msg(
                f'СуперМаг Глобальный. Создание: У сотрудника {employee.firstname, employee.lastname, employee.surname} все логины {sm_login}, {sm_long_login}, {sm_full_login} уже существуют')
    else:
        sm_success = True
        return sm_success

    if flags['SM_LOCAL']:
        for dbname in store_names:
            sm_conn.create_user_in_local_db(dbname,employee.sm_login,employee.password, test_role_id)
            send_msg(f'СуперМаг Локальный. Создание: Сотруднику {employee.firstname, employee.lastname, employee.surname} на должность {userData["J2"].value} создан аккаунт в {dbname} c логином {employee.sm_login}')
            sm_local_success = True
            return sm_local_success
    else:
        sm_local_success = False
        return sm_local_success

    if ad_success and bx24_success and c1_success and sm_success and sm_local_success:
        return True
    else:
        return False
