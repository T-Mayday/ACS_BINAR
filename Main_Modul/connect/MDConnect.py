import requests
import configparser

from connect.bitrixConnect import Bitrix24Connector
bitrix_connector = Bitrix24Connector()

class MDAUIDConnect:
    def __init__(self):
        try:
            self.config = configparser.ConfigParser()
            self.config.read('connect_domain.ini')
            self.base_url = self.config['MD_AUDIT']['base_url']
            self.api_token = self.config['MD_AUDIT']['api_token']
        except Exception as e:
            log.exception("Error read config", e)

    def find_user_by_email(self, email):
        url = f"{self.base_url}/orgstruct/users"
        headers = {
            "Authorization": f"Bearer {self.api_token}",
            "Content-Type": "application/json"
        }
        params = {
            "email": email
        }
        response = requests.get(url, headers=headers, params=params)
        if response.status_code == 200:
            data = response.json()
            if data:
                return data[0]
            else:
                bitrix_connector.send_msg_error(f"MD_AUDIT. ПОИСК. Пользователь с таким {email} не найден.")
                return []
        else:
            bitrix_connector.send_msg_error(f"MD_AUDIT. ПОИСК. Ошибка {response.status_code}: {response.text}")
            return []

    def block_user(self, user_id, lastname, firstname, surname, department, postjob):
        url = f"{self.base_url}/orgstruct/users/{user_id}"
        headers = {
            "Authorization": f"Bearer {self.api_token}",
            "Content-Type": "application/json"
        }
        data = {
            "active": False
        }

        response = requests.patch(url, headers=headers, json=data)

        if response.status_code == 200:
            bitrix_connector.send_msg(f"MD_AUDIT.Блокировка. Сотрудник {lastname} {firstname} {surname} из {department} с должностью {postjob}. Выполнено.")
            return True
        else:
            bitrix_connector.send_msg_error(f"MD_AUDIT.Блокировка.Ошибка у сотрудника {lastname} {firstname} {surname} из {department} с должностью {postjob}. {response.status_code}: {response.text}")
            return False