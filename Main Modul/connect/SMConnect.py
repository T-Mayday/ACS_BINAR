import cx_Oracle
import configparser

from message.message import log

from connect.bitrixConnect import Bitrix24Connector
bitrix_connector = Bitrix24Connector()

class SMConnect:
    def __init__(self):
        self.config = configparser.ConfigParser()
        self.config.read('connect_domain.ini')
        self.service_name = self.config['SM']['service_name']
        self.username = self.config['SM']['username']
        self.password = self.config['SM']['password']
        self.role_id = self.config['SM']['role_id']
        self.connection = None
        self.cursor = None

    def connect_SM(self):
        try:
            # dsn = cx_Oracle.makedsn(self.service_name)
            self.connection = cx_Oracle.connect(self.username, self.password, self.service_name)
            self.cursor = self.connection.cursor()
        except cx_Oracle.DatabaseError as e:
            bitrix_connector.send_msg_error(f"SM^ Ошибка подключения: {self.username}@{self.service_name} {e}")
            raise

    def connect_SM_LOCAL(self, service_name):
        try:
            dsn = cx_Oracle.makedsn(service_name)
            self.connection = cx_Oracle.connect(self.username, self.password, dsn)
            self.cursor = self.connection.cursor()
        except cx_Oracle.DatabaseError as e:
            bitrix_connector.send_msg_error(f"SM LOCAL: Ошибка подключения: {e}")
            raise

    def close(self):
        if self.cursor:
            self.cursor.close()
        if self.connection:
            self.connection.close()

    def execute_query(self, query):
        try:
            self.cursor.execute(query)
            result = self.cursor.fetchall() or []
            return result
        except cx_Oracle.DatabaseError as e:
            bitrix_connector.send_msg_error(f"SM: Ошибка выполнения запроса: {query} {e}")
            return []

    def execute_update(self, query):
        try:
            self.cursor.execute(query)
            self.connection.commit()
        except cx_Oracle.DatabaseError as e:
            bitrix_connector.send_msg_error(f"SM: Ошибка выполнения обновления: {query} {e}")
            raise

    def execute_procedure(self, procedure_name, params):
        try:
            self.cursor.callproc(procedure_name, params)
            self.connection.commit()
        except cx_Oracle.DatabaseError as e:
            bitrix_connector.send_msg_error(f"SM: Ошибка выполнения процедуры: {procedure_name} {params} {e}")
            raise

    def user_exists(self, surname):
        query = f"""
        SELECT NVL((SELECT userenabled FROM supermag.smstaff WHERE UPPER(surname) = UPPER('{surname}')), -1) AS user_exists
        FROM dual
        """
        result = self.execute_query(query)
        if result:
            r = result[0]
        else:
            r = []
        return r

    def create_user(self, username, password, role):
        try:
            query = f"""
            DECLARE
                pUser VARCHAR2(255) := '{username}'; -- Логин
                password VARCHAR2(255) := '{password}'; -- Пароль
                pDol VARCHAR2(20) := '{role}'; -- Должность
            BEGIN
                SUPERMAG.BIN_CreateUser(pUser, password, pDol, 1, 1);
            END;
            """
            self.execute_update(query)
            bitrix_connector.send_msg(f"SM: Создание. Пользователь {username}. Выполено успешно.")
            return True
        except cx_Oracle.DatabaseError as e:
            bitrix_connector.send_msg_error(f"SM: Создание. Пользователь {username}. Ошибка : {e}")
#            raise
            return False

    def block_user(self, login):
        try:
            query = f"""
            BEGIN
                EXECUTE IMMEDIATE 'REVOKE SUPERMAG_USER FROM {login}';

                UPDATE Supermag.SMStaff
                SET UserEnabled = '0'
                WHERE surname = '{login}';

                COMMIT;
            END;
            """
            self.execute_update(query)
            bitrix_connector.send_msg(f"SM Блокировка: {login}. Выполнено успешно.")
            return True
        except cx_Oracle.DatabaseError as e:
            bitrix_connector.send_msg_error(f"SM Блокировка: {login}. Ошибка {e}")
#            raise
            return False

    def unblock_user(self, login):
        try:
            query = f"""
            BEGIN
                EXECUTE IMMEDIATE 'GRANT SUPERMAG_USER TO {login}';

                UPDATE Supermag.SMStaff
                SET UserEnabled = '1'
                WHERE ID = {login};

                COMMIT;
            END;
            """
            self.execute_update(query)
            bitrix_connector.send_msg(f"SM: Разблокировка. {login}. Выполнено успешно.")
            return True
        except cx_Oracle.DatabaseError as e:
            bitrix_connector.send_msg_error(f"SM: Разблокировка. {login}. Ошибка: {e}")
#            raise
            return False

    def get_store(self, store_id):
        try:
            with self.cursor as cursor:
                cursor.execute("""
                    SELECT
                        l.id,
                        l.name,
                        t.dbname
                    FROM supermag.smstorelocations l
                    JOIN (SELECT storeloc, propval AS dbname 
                          FROM supermag.smstoreproperties 
                          WHERE propid = 'REP.DBNAME') t
                    ON l.id = t.storeloc
                    WHERE l.id = :store_id
                    ORDER BY l.name
                """, store_id=store_id)

                result = cursor.fetchone()
#                result = cursor.fetchall()
#                cursor.close()
                return {
                    "id": result[0],
                    "name": result[1],
                    "dbname": result[2]
                } if result else None
        except cx_Oracle.DatabaseError as e:
            bitrix_connector.send_msg_error(f"SM: Ошибка получения данных по store_id={store_id}: {e}")
            raise
        except Exception as e:
            bitrix_connector.send_msg_error(f"SM: Ошибка получения данных по store_id={store_id}: {e}")
            raise

    def create_user_in_local_db(self, dbname, user_login, user_password, user_role):
        """Создание пользователя в локальной базе данных."""
        local_dsn = cx_Oracle.makedsn('localhost', '1521', service_name=dbname)
        try:
            local_connection = cx_Oracle.connect(self.username, self.password, local_dsn)
            with local_connection.cursor() as cursor:
                cursor.execute("""
                    DECLARE
                        pUser VARCHAR2(255) := :user_login; 
                        password VARCHAR2(255) := :user_password;                     
                        pDol VARCHAR2(20) := :user_role;
                    BEGIN 
                        SUPERMAG.BIN_CreateUser(pUser, password, pDol, 1, 1);
                    END;
                """, user_login=user_login, user_password=user_password, user_role=user_role)

                local_connection.commit()
                log.info(f"SM LOCAL: Пользователь {user_login} в базе данных {dbname} успешно создан")
                return True
        except Exception as e:
            bitrix_connector.send_msg_error(f"SM LOCAL: Пользователь {user_login} в базе данных {dbname}. Не удалось создать, ошибка: {e}")
            return False
        finally:
            if local_connection:
                local_connection.close()

    def getRoleID(self):
        return self.role_id



