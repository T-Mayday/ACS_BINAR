# Функции для работы с LDAP Сервером
import ldap
from ldap.filter import escape_filter_chars
import configparser

from message.message import log

from connect.bitrixConnect import Bitrix24Connector
bitrix_connector = Bitrix24Connector()

class ActiveDirectoryConnector:
    #  Получение данных из ini файла
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
        self.adress = self.config.get('Domain', 'address')
        self.dir_input = self.config.get('Domain', 'input')
        self.dir_output = self.config.get('Domain', 'output')
        self.dir_waste = self.config.get('Domain', 'waste')
        self.dir_error = self.config.get('Domain', 'error')
        self.dbinfo = self.config.get('Domain', 'dbinfo')

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

    # Подключение к LDAP Серверу
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

    # Отключение от LDAP Сервера
    def disconnect_ad(self, conn):
        try:
            conn.unbind_s()
        except ldap.LDAPError as e:
            log.error(f"AD. Ошибка при отключении от LDAP: {e}")

    # Поиск по атрибуту (EmployyID)
    def search_in_ad(self, INN):
        conn = self.connect_ad()
        if not conn:
            return []
        search_filter = f"(employeeID={escape_filter_chars(INN)})"
        
        try:
            result = conn.search_s(self.base_dn, ldap.SCOPE_SUBTREE, search_filter)
            return result
        except Exception as e:
            bitrix_connector.send_msg_error(f'LDAP Ошибка поиcка по INN: {search_filter} {str(e)} ')
            return []
        finally:
            self.disconnect_ad(conn)

    # Поиск по атрибуту (mail)
    def search_by_mail(self, mail):
        conn = self.connect_ad()
        if not conn:
            return []
        search_filter = f"(mail={escape_filter_chars(mail)})"
        try:
            result = conn.search_s(self.base_dn, ldap.SCOPE_SUBTREE, search_filter)
            return result
        except Exception as e:
            bitrix_connector.send_msg_error(f'LDAP Ошибка поиcка по mail: {search_filter} {str(e)} ')
            return []
        finally:
            self.disconnect_ad(conn)

    # Поиск по атрибуту (DisplayName)
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
            return []
        finally:
            self.disconnect_ad(conn)

    # Создание аккаунта в Active Directory
    def create_user(self, login_type, employee, userData, INN):
        conn = self.connect_ad()
        if not conn:
            return []
        try:
            dn = self.getNewUserDn().format(login_type)
            user_dn = f"CN={login_type},{dn}"
            attrs = [
                ('objectClass', [b'top', b'person', b'organizationalPerson', b'user']),
                ('cn', [login_type.encode('utf-8')]),
                ('givenName', [employee.firstname.encode('utf-8')]),
                ('sAMAccountName', [login_type.encode('utf-8')]),
                ('userPrincipalName', [login_type.encode('utf-8') + b'@BINLTD.local']),
                #           ('displayName', [login_type.encode('utf-8')]),
                ('displayName', [str(f"{employee.lastname} {employee.firstname} {employee.surname}").encode('utf-8')]),
                ('department', [userData['G2'].value.encode('utf-8')]),
                ('mail', [employee.create_email(login_type).encode('utf-8')]),
                ('sn', [employee.lastname.encode('utf-8')]),
                ('employeeID', [str(INN).encode('utf-8')]),
                ('company', [userData['F2'].value.encode('utf-8')]),
                ('userAccountControl', [b'512']),
                ('middleName', [employee.surname.encode('utf-8')]),
                ('title', [userData['J2'].value.encode('utf-8')]),
                ('unicodePwd', [str(f"\"{employee.password}\"").encode('utf-16-le')]),
                ('pwdLastSet', [b'0'])
            ]
            conn.add_s(user_dn, attrs)
            created = self.search_in_ad(INN)
            if created is not None and len(created) != 0:
                bitrix_connector.send_msg_adm(
                    f"{employee.lastname, employee.firstname, employee.surname} {user_dn} {employee.password}")
                bitrix_connector.send_msg(
                    f"AD. Создание: Сотруднику {employee.lastname, employee.firstname, employee.surname} {user_dn}. Выполнено")
                return True
            else:
                bitrix_connector.send_msg_error(
                    f"AD. Создание: Сотрудник {employee.lastname, employee.firstname, employee.surname} из отдела {userData['G2'].value} на должность {userData['J2'].value}. {user_dn}. Не выполнено")
                return False
        except Exception as e:
            bitrix_connector.send_msg_error(
                f"AD. Создание: Сотрудник {employee.lastname, employee.firstname, employee.surname} из отдела {userData['G2'].value} на должность {userData['J2'].value}. Ошибка при создании {str(user_dn)} {str(attrs)} {str(e)}")
            return False
        finally:
            self.disconnect_ad(conn)

    # Активация аккаунта в Active Directory
    def activate_user(self, user_dn, employee, userData):
        conn = self.connect_ad()
        if not conn:
            return []
        try:
            mod_attrs = [(ldap.MOD_REPLACE, 'userAccountControl', b'512')]
            conn.modify_s(user_dn, mod_attrs)
            bitrix_connector.send_msg(
                f"AD. Активация: Учетная запись сотрудника {employee.firstname} {employee.lastname} {employee.surname} из отдела {userData['G2'].value} на должность {userData['J2'].value}: {str(user_dn)} была успешно активирована.")
            return True
        except Exception as e:
            bitrix_connector.send_msg_error(
                f"AD. Ошибка! При активации учетной записи {employee.firstname} {employee.lastname} {employee.surname} из отдела {userData['G2'].value} на должность {userData['J2'].value}: {str(user_dn)} {e}")
            return False
        finally:
            self.disconnect_ad(conn)

    # Обновление атрибутов аккаунта в Active Directory
    def update_user(self, user, new_attr, employee, userData):
        conn = self.connect_ad()
        if not conn:
            return []

        user_dn, user_attrs = user[0]
        updated_attrs = []
        for attr_name, attr_value in new_attr.items():
            if attr_name in user_attrs and user_attrs[attr_name][0] != attr_value:
                mod_attrs = [(ldap.MOD_REPLACE, attr_name, attr_value)]
                try:
                    conn.modify_s(user_dn, mod_attrs)
                    updated_attrs.append(attr_name)
                except Exception as e:
                    bitrix_connector.send_msg_error(
                        f"AD. Изменение: Сотрудник {employee.lastname} {employee.firstname} {employee.surname} из отдела {userData['G2'].value} на должность {userData['J2'].value}. Не выполнено. Ошибка при обновлении атрибута {attr_name}: {str(e)}"
                    )
                    return False
        self.disconnect_ad(conn)
        if updated_attrs:
            updated_attrs_str = ', '.join(updated_attrs)
            bitrix_connector.send_msg(
                f"AD. Изменение: Сотрудник {employee.lastname} {employee.firstname} {employee.surname}. Обновление атрибута {updated_attrs_str}. Выполнено"
            )
            return True
        else:
            return True

    # Блокировка аккаунта в Active Directory
    def block_user(self,user, employee, userData):
        conn = self.connect_ad()
        if not conn:
            return []

        block_attr = {
                    'userAccountControl': b'514'
                }
        user_dn, user_attrs = user[0]

        if 'userAccountControl' in user_attrs and user_attrs['userAccountControl'][0] == b'514':
            bitrix_connector.send_msg(
                f"AD. Блокировка: Сотрудник {employee.lastname} {employee.firstname} {employee.surname} уже заблокирован."
            )
            return True
        for attr_name, attr_value in block_attr.items():
            if attr_name in user_attrs and user_attrs[attr_name][0] != attr_value:
                mod_attrs = [(ldap.MOD_REPLACE, attr_name, attr_value)]
                try:
                    conn.modify_s(user_dn, mod_attrs)
                    bitrix_connector.send_msg(
                        f"AD. Блокировка: Сотрудник {employee.lastname} {employee.firstname} {employee.surname} {user_dn}. Выполнено")
                    return True
                except Exception as e:
                    bitrix_connector.send_msg_error(
                        f"AD. Блокировка: Сотрудник {employee.lastname} {employee.firstname} {employee.surname} из отдела {userData['G2'].value} на должность {userData['J2'].value} {user_dn}. Не выполнено - ошибка {str(e)}")
                    return False
        self.disconnect_ad(conn)



