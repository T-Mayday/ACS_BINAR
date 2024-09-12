import ldap
import configparser

from message.message import log
class ActiveDirectoryConnector:
    def __init__(self):
        self.config = configparser.ConfigParser()
        self.config.read('connect_domain.ini')
        self.domain_name = self.config.get('Domain', 'domain_name')
        self.ip = self.config.get('Domain', 'ip')
        self.username = self.config.get('Domain', 'username')
        self.password = self.config.get('Domain', 'password')
        self.state = self.config.get('Domain', 'state')
        self.base_dn = self.config.get('Domain', 'base_dn')
        self.dn = self.config.get( 'Domain', 'dn')
        self.search_base = self.config.get( 'Domain', 'search_base')
        self.adress = self.config.get('Domain', 'ardress')
        self.dir_input = self.config.get('Domain', 'input')
        self.dir_output =self.config.get('Domain', 'output')
        self.dir_waste = self.config.get('Domain', 'waste')

    def getBaseDn(self):
        return self.base_dn

    def getDn(self):
        return self.dn

    def getSearchBase(self):
        return self.search_base

    def getState(self):
        return self.state
    def getAdress(self):
        return self.adress
    def getInput(self):
        return self.dir_input
    def getOutput(self):
        return self.dir_output
    def getWaste(self):
        return self.dir_waste
    def connect_ad(self):
        ldap.set_option(ldap.OPT_X_TLS_REQUIRE_CERT, ldap.OPT_X_TLS_NEVER)
        ldap.set_option(ldap.OPT_REFERRALS, 0)

        try:
            conn = ldap.initialize(f'ldaps://{self.ip}:636')
            conn.simple_bind_s(f'{self.domain_name}\\{self.username}', self.password)
            return conn
        except ldap.LDAPError as e:
            log.error(f"AD.Ошибка подключения. При установке соединения с LDAP: {e}")
            return None