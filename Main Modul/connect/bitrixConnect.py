import configparser
import requests as req
from urllib.parse import urlparse, parse_qs
from requests.auth import HTTPBasicAuth
from pybitrix24 import Bitrix24

from message.message import send_msg_error


class Bitrix24Connector:
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

    def connect(self):
        bx24 = Bitrix24(self.linkBx24, self.clientId, self.clientSecret)
        rsp = req.get(bx24.build_authorization_url(), auth=HTTPBasicAuth(self.user_login, self.user_password))
        auth_code = self.parse_query_param('code', rsp.url)
        tokens = bx24.obtain_tokens(auth_code, scope='')
        return bx24, tokens

    def find_jobfriend(self, bx24, post_job, codeBX24):
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
            send_msg_error(f"BX24: Ошибка поиска сотрудника {post_job} в департаменте {codeBX24}: {e}")
        return None
    def search_email(self, bx24, email):
        try:
            bx24.refresh_tokens()

            result = bx24.call("user.get", {"EMAIL": email})

            if 'result' in result and result['result']:
                user_info = result['result'][0]
                return user_info.get('ID')
            else:
                result = []
                return result

        except Exception as e:
            send_msg_error(f"BX24: Ошибка при поиске пользователя по email '{email}': {e}")
            return None

    def search_user(self, bx24, last_name, name, second_name):
        try:
            bx24.refresh_tokens()
            result = bx24.call("user.get", {"LAST_NAME": last_name, "NAME": name, "SECOND_NAME": second_name})
            if result.get('result'):
                r = result.get('result')[0]
                return r.get('ID')
            if result.get('error'):
                send_msg_error(
                    f"BX24. Пользователь с ФИО '{last_name} {name} {second_name}' не найден. {result.get('error')[0]}")
                return None
        except Exception as e:
            send_msg_error(f"BX24. Ошибка при получении пользователей: {last_name} {name} {second_name} {e}")
            return None