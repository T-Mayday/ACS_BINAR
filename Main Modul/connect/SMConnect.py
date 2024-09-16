import cx_Oracle
import configparser

from message.message import send_msg, send_msg_error, log


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
    def getRoleID(self):
        return self.role_id

    def connect_SM(self):
        try:
            # dsn = cx_Oracle.makedsn(self.service_name)
            self.connection = cx_Oracle.connect(self.username, self.password, self.service_name)
            self.cursor = self.connection.cursor()
        except cx_Oracle.DatabaseError as e:
            send_msg_error(f"SM Ошибка подключения: {self.username}@{self.service_name} {e}")
            raise

    def connect_SM_LOCAL(self, service_name):
        try:
            dsn = cx_Oracle.makedsn(service_name)
            self.connection = cx_Oracle.connect(self.username, self.password, dsn)
            self.cursor = self.connection.cursor()
        except cx_Oracle.DatabaseError as e:
            send_msg_error(f"SM Ошибка подключения: {self.username}@{service_name} {e}")
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
            send_msg_error(f"SM Ошибка выполнения запроса: {query} {e}")
            return []

    def execute_update(self, query):
        try:
            self.cursor.execute(query)
            self.connection.commit()
        except cx_Oracle.DatabaseError as e:
            send_msg_error(f"SM Ошибка выполнения обновления: {query} {e}")
            raise

    def execute_procedure(self, procedure_name, params):
        try:
            self.cursor.callproc(procedure_name, params)
            self.connection.commit()
        except cx_Oracle.DatabaseError as e:
            send_msg_error(f"SM Ошибка выполнения процедуры: {query} {e}")
            raise

    def user_exists(self, username):
        query = f"SELECT * FROM users WHERE username = '{username}'"
        result = self.execute_query(query)

        if result and isinstance(result, (list, tuple)):
            return result[0]
        else:
            return None

    def create_user(self, username, password, role_id):
        query = f"INSERT INTO users (username, password, role_id) VALUES ('{username}', '{password}', '{role_id}')"
        self.execute_update(query)

    def block_user(self, username):
        query = f"UPDATE users SET status='blocked' WHERE username='{username}'"
        self.execute_update(query)

    def unblock_user(self, username, user_id):
        query = f"UPDATE users SET status='active' WHERE username='{username}' AND id='{user_id}'"
        self.execute_update(query)

    def get_store(self, store_id):
        try:
            with self.cursor as cursor:
                cursor.execute("""
                    SELECT t.dbname
                    FROM supermag.smstorelocations l
                    JOIN (SELECT storeloc, propval AS dbname FROM supermag.smstoreproperties WHERE propid='REP.DBNAME') t
                    ON l.id = t.storeloc
                    WHERE l.id = :store_id
                """, store_id=store_id)
                result = cursor.fetchone()
                return result[0] if result else None
        except cx_Oracle.DatabaseError as e:
            send_msg_error(f"SM Ошибка получения dbname по store_id={store_id}: {e}")
            raise

    def create_user_in_local_db(self, dbname, user_login, user_password, role_id):
        """Создание пользователя в локальной базе данных."""
        local_dsn = cx_Oracle.makedsn('localhost', '1521', service_name=dbname)
        try:
            local_connection = cx_Oracle.connect(self.username, self.password, local_dsn)
            with local_connection.cursor() as cursor:
                cursor.execute("""
                    DECLARE
                        pUser VARCHAR2(255) := :user_login; 
                        password VARCHAR2(255) := :user_password;                     
                        pDolID NUMBER := :role_id;
                        pDolzhnost VARCHAR2(20);
                    BEGIN 
                        SELECT orarole INTO pDolzhnost FROM supermag.smoffcfg WHERE id = pDolID;                    
                        Supermag.BIN_CreateUser(pUser, password, pDolzhnost, 1, 1);
                    END;            
                """, user_login=user_login, user_password=user_password, role_id=role_id)

                local_connection.commit()
                send_msg(f"Пользователь {user_login} успешно создан в базе данных {dbname}.")
        except Exception as e:
            send_msg_error(f"SM Не удалось создать пользователя {user_login} в базе данных {dbname}: {e}")
        finally:
            if local_connection:
                local_connection.close()


