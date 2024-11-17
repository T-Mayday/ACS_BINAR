# Функции для работа с Bitrix24

import configparser
import requests as req
from urllib.parse import urlparse, parse_qs
from requests.auth import HTTPBasicAuth
from pybitrix24 import Bitrix24

from message.message import log

class Bitrix24Connector:
    #  Получение данных из ini файла
    def __init__(self):
        self.config = configparser.ConfigParser()
        self.config.read('connect_domain.ini')
        self.user_login = self.config['Bitrix24']['user_login']
        self.user_password = self.config['Bitrix24']['user_password']
        self.linkBx24 = self.config.get('Bitrix24', 'linkBx24')
        self.clientId = self.config.get('Bitrix24', 'clientId')
        self.clientSecret = self.config.get('Bitrix24', 'clientSecret')
        self.chatID = self.config.get('Bitrix24', 'chatID')
        self.chatadmID = self.config.get('Bitrix24', 'chatadmID')
    def getChatID(self):
        return self.chatID

    def parse_query_param(self, param_name, url):
        parsed_url = urlparse(url)
        query_params = parse_qs(parsed_url.query)
        return query_params.get(param_name)[0]

    # Подключение к Bitrix24
    def connect(self):
        bx24 = Bitrix24(self.linkBx24, self.clientId, self.clientSecret)
        rsp = req.get(bx24.build_authorization_url(), auth=HTTPBasicAuth(self.user_login, self.user_password))
        auth_code = self.parse_query_param('code', rsp.url)
        tokens = bx24.obtain_tokens(auth_code, scope='')
        return bx24, tokens

    # Отправка сообщений в чат для трассировки
    def send_msg(self, msg):
        bx24, tokens = self.connect()
        try:
            log.info(msg)
            bx24.refresh_tokens()
            res = bx24.call('im.message.add', {'DIALOG_ID': self.chatID, 'MESSAGE': msg, 'URL_PREVIEW': 'N'})
        except Exception as e:
            log.exception("Error sending message", e)

    # Отправка сообщений в чат для трассировки
    def send_msg_error(self, msg):
        bx24, tokens = self.connect()
        try:
            log.exception(msg)
            bx24.refresh_tokens()
            res = bx24.call('im.message.add', {'DIALOG_ID': self.chatID, 'MESSAGE': msg, 'URL_PREVIEW': 'N'})
        except Exception as e:
            log.exception(e)

    # Отправка сообщений в чат адмнистраторов
    def send_msg_adm(self, msg):
        bx24, tokens = self.connect()
        try:
            log.info(msg)
            bx24.refresh_tokens()
            res = bx24.call('im.message.add', {'DIALOG_ID': self.chatadmID, 'MESSAGE': msg, 'URL_PREVIEW': 'N'})
        except Exception as e:
            log.exception("Error sending message", e)

    # Отправка сообщение пользователям Bitrix24
    def send_msg_user(self, user_id, msg):
        bx24, tokens = self.connect()
        try:

            bx24.refresh_tokens()
            message_data = {
                'DIALOG_ID': user_id,
                'MESSAGE': msg,
                'SYSTEM': 'N',
                'URL_PREVIEW': 'N'
            }
            res = bx24.call('im.message.add', message_data)
        except Exception as e:
            log.exception("Error sending message", e)

    # Поиск сосодрудника для 1с
    def find_jobfriend(self, post_job, codeBX24):
        bx24, tokens = self.connect()
        filter_params = {
            'FILTER': {
                'WORK_POSITION': post_job,
                'UF_DEPARTMENT': codeBX24
            }
        }
        try:
            employees = bx24.call('user.get', filter_params)
            if 'result' in employees and employees['result']:
                employee_info = employees['result'][0]
                return f"{employee_info['LAST_NAME']} {employee_info['NAME']} {employee_info['SECOND_NAME']}"
        except Exception as e:
            self.send_msg_error(f"BX24: Ошибка поиска сотрудника {post_job} в департаменте {codeBX24}: {e}")
        return None

    # Поиск по Почте
    def search_email(self, email):
        bx24, tokens = self.connect()
        try:
            bx24.refresh_tokens()
            result = bx24.call("user.get", {"EMAIL": email})
            if 'result' in result and result['result']:
                user_info = result['result'][0]
#                return user_info.get('ID')
                return user_info
            else:
                return []
        except Exception as e:
            self.send_msg_error(f"BX24: Ошибка при поиске пользователя по email '{email}': {e}")
            return []

    # Поиск по ФИО
    def search_user(self, last_name, name, second_name):
        bx24, tokens = self.connect()
        try:
            bx24.refresh_tokens()
            result = bx24.call("user.get", {"LAST_NAME": last_name, "NAME": name, "SECOND_NAME": second_name})
            if result.get('result'):
                user_info = result.get('result')[0]
                user_id = user_info.get('ID')
                user_active = user_info.get('ACTIVE')
                if user_active == 'Y':
                    return user_id
                else:
                    return None
            if result.get('error'):
                self.send_msg_error(
                    f"BX24. Пользователь с ФИО '{last_name} {name} {second_name}' не найден. {result.get('error')[0]}")
                return None
        except Exception as e:
            self.send_msg_error(f"BX24. Ошибка при получении пользователей: {last_name} {name} {second_name} {e}")
            return None

    # Обновление пользователя
    def update_user(self, user_info, new_data, employee, userData):
        bx24, tokens = self.connect()
        update_attr = []
        try:
            for key in user_info.keys() & new_data.keys():
                if user_info[key] != new_data[key]:
                    update_attr.append((key, new_data[key]))
                    user_id = user_info.get('ID')
                    result = bx24.call('user.update', {'ID': user_id, **new_data})
                    if result.get('error'):
                        self.send_msg_error(
                            f"BX24. Ошибка при изменении пользователя: Сотрудник {employee.lastname} {employee.firstname} {employee.surname} из отдела {userData['G2'].value} на должность {userData['J2'].value}. Ошибка: {result.get('error_description')}")
                        return False, False
                    if result.get('result'):
                        self.send_msg(
                            f"BX24. Обновление данных сотрудника {employee.firstname} {employee.lastname} {employee.surname} ID = {user_id}. Выполнено.")
                        if len(update_attr) > 0:
                            return True, True
                else:
                    return True, False
        except Exception as e:
            self.send_msg_error(f"BX24. Ошибка при обновлении данных {user_info.get('ID')} {new_data} Ошибка {e}")
            return False, False

    # Создания пользователя
    def create_user(self, email, employee, userData):
        bx24, tokens = self.connect()
        try:
            user_data = {
                "NAME": employee.firstname,
                "LAST_NAME": employee.lastname,
                "SECOND_NAME": employee.surname,
                "EMAIL": email,
                "UF_DEPARTMENT": str(userData['H2'].value),
                "ACTIVE": "Y",
                "WORK_POSITION": str(userData["J2"].value),
            }
            bx24.refresh_tokens()
            create = bx24.call('user.add', user_data)
            if create.get('error'):
                error_message = create.get('error_description')
                self.send_msg_error(
                    f"BX24.Создание !Не выполнено: Сотрудник {employee.firstname, employee.lastname, employee.surname} из отдела {userData['H2'].value} на должность {userData['J2'].value}. {user_data} {error_message}")
                return False
            if create.get('result'):
                user_id = create.get('result')
                self.send_msg(
                        f"BX24.Создание: Сотрудник {employee.firstname, employee.lastname, employee.surname} ID = {user_id} {userData['H2'].value}. Выполнено")
                return True
            else:
                self.send_msg_error(
                        f"BX24.Создание !Не выполнено: Сотрудник {employee.firstname, employee.lastname, employee.surname}. Ошибка: {create.get('result')}")
                return False
        except Exception as e:
            self.send_msg_error(
                f"BX24. Создание: Сотрудник {employee.firstname, employee.lastname, employee.surname} из отдела {userData['G2'].value} на должность {userData['J2'].value}. Ошибка: {e}")
            return False

    # Блокировка пользователя
    def block_user(self, user_id, employee, userData):
        bx24, tokens = self.connect()
        try:
            bx24.refresh_tokens()
            result = bx24.call('user.update', {
                'ID': user_id,
                'ACTIVE': 'N'
            })
            self.send_msg(
                f"BX24. Блокировка: {employee.lastname, employee.firstname, employee.surname} {user_id}. Выполнено")
            return True
        except Exception as e:
            self.send_msg_error(f"BX24. Блокировка: {employee.lastname, employee.firstname, employee.surname} из отдела {userData['G2'].value} на должность {userData['J2'].value}. {user_id} {result}. Ошибка {e}")
            return False








        