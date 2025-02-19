import configparser
import requests

from connect.bitrixConnect import Bitrix24Connector
bitrix_connector = Bitrix24Connector()

class Connector1C:
    def __init__(self):
        self.config = configparser.ConfigParser()
        self.config.read('connect_domain.ini')
        self.url_create = self.config.get('INFO1C', 'url_create')
        self.url_changes = self.config.get('INFO1C', 'url_changes')
        self.url_block = self.config.get('INFO1C', 'url_block')

    def getUrlCreate(self):
        return self.url_create

    def getUrlChanges(self):
        return self.url_changes

    def getUrlBlock(self):
        return self.url_block

    def send_rq(self, url, data, employee, userData, action):
        try:
            headers = {'Content-Type': 'application/json'}
            response = requests.post(url, json=data, headers=headers)
            if response.status_code == 200:
                bitrix_connector.send_msg(
                    f"1С. {action}. Сотрудник {employee.lastname, employee.firstname, employee.surname}. Выполнено")
                return True
            elif response.status_code == 500:
                bitrix_connector.send_msg(
                    f"1С. {action}. Сотрудник {employee.lastname, employee.firstname, employee.surname}.  Уже зарегистрирован")
                return True

            else:
                result = response.text
                bitrix_connector.send_msg_error(
                    f'1С. {action}. Сотрудник {employee.lastname, employee.firstname, employee.surname} из отдела {userData["G2"].value} на должность {userData["J2"].value}. Не выполнено. Ошибки - {response.status_code} {url} {data} {result}')
                return False
        except requests.exceptions.RequestException as e:
            bitrix_connector.send_msg_error(
                f"1С. {action}. Сотрудник {employee.lastname, employee.firstname, employee.surname}. Ошибка {url} {data} Error: {e}")
            return False
