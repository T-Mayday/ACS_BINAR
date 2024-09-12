import configparser
import requests as req
from urllib.parse import urlparse, parse_qs
from requests.auth import HTTPBasicAuth
from pybitrix24 import Bitrix24



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