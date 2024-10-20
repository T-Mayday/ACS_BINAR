import ldap
from ldap.filter import escape_filter_chars
import configparser

from message.message import log

from connect.bitrixConnect import Bitrix24Connector
bitrix_connector = Bitrix24Connector()


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
        self.newuser_dn = self.config.get( 'Domain', 'newuser_dn')
        self.search_base = self.config.get( 'Domain', 'search_base')
        self.adress = self.config.get('Domain', 'ardress')
        self.dir_input = self.config.get('Domain', 'input')
        self.dir_output =self.config.get('Domain', 'output')
        self.dir_waste = self.config.get('Domain', 'waste')
        self.dir_error = self.config.get('Domain', 'error')

    def getBaseDn(self):
        return self.base_dn

    def getDn(self):
        return self.dn

    def getNewUserDn(self):
        return self.newuser_dn

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
    def getError(self):
        return self.dir_error
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

    def disconnect_ad(self,conn):
        try:
            conn.unbind_s()
        except ldap.LDAPError as e:
            log.error(f"AD. Ошибка при отключении от LDAP: {e}")

    def search_in_ad(self, INN):
        conn = self.connect_ad()
        if not conn:
            return None
        search_filter = f"(employeeID={escape_filter_chars(INN)})"
        try:
            result = conn.search_s(self.base_dn, ldap.SCOPE_SUBTREE, search_filter)
            return result

        except Exception as e:
            bitrix_connector.send_msg_error(f'LDAP Ошибка поиcка по INN: {search_filter} {str(e)} ')
            return None
        finally:
            self.disconnect_ad(conn)

    def search_by_mail(self, mail):
        conn = self.connect_ad()
        if not conn:
            return None
        search_filter = f"(mail={escape_filter_chars(mail)})"
        try:
            result = conn.search_s(self.base_dn, ldap.SCOPE_SUBTREE, search_filter)
            return result
        except Exception as e:
            bitrix_connector.send_msg_error(f'LDAP Ошибка поиcка по mail: {search_filter} {str(e)} ')
            return None
        finally:
            self.disconnect_ad(conn)

    def search_by_fullname(self, full_name):
        conn = self.connect_ad()
        if not conn:
            return []
        search_filter = f"(displayName={escape_filter_chars(full_name)})"
        try:
            result = conn.search_s(self.base_dn, ldap.SCOPE_SUBTREE, search_filter)
            return result
        except Exception as e:
            bitrix_connector.send_msg_error(f'LDAP Ошибка поиcка по ФИО: {search_filter} {str(e)} ')
            return None
        finally:
            self.disconnect_ad(conn)


