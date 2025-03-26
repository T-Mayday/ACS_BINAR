import psycopg2
import configparser
from message.message import log
from connect.bitrixConnect import Bitrix24Connector  # Подключение Bitrix24
bitrix_connector = Bitrix24Connector()

class DatabaseConnector:
    def __init__(self):
        self.config = configparser.ConfigParser(interpolation=None)  # Отключаем интерполяцию для '%'
        self.config.read('connect_domain.ini', encoding='utf-8')
        self.host = self.config.get('DATABASE', 'host')
        self.user = self.config.get('DATABASE', 'user')
        self.password = self.config.get('DATABASE', 'password')
        self.db = self.config.get('DATABASE', 'db')
        self.port = self.config.get('DATABASE', 'port')
        self.connection = self.connect()

    def connect(self):
        try:
            DB_CONFIG = {
                "dbname": self.db,
                "user": self.user,
                "password": self.password,
                "host": self.host,
                "port": self.port,
                "client_encoding": "UTF8"
            }
            conn = psycopg2.connect(**DB_CONFIG)
            return conn
        except psycopg2.Error as e:
            log.exception("Ошибка подключения к базе данных:", e)
            bitrix_connector.send_msg_error(f"Ошибка подключения к базе данных: {str(e)}")
            return None

    def user_verification(self, department, position):
        # Флаги с нужными названиями
        flags = {
            'AD': False,
            'BX24': False,
            'ZUP': False,
            'RTL': False,
            'ERP': False,
            'SM_GEN': False,
            'SM_LOCAL': False,
            'Normal_account': False,
            'Shop_account': False
        }

        # Сопоставление названий систем из БД с нужными флагами
        system_mapping = {
            'AD': 'AD',
            'BITRIX24': 'BX24',
            'SM_BINUU00': 'SM_GEN',
            'SM_LOCAL': 'SM_LOCAL',
            '1CERP': 'ERP',
            '1CRTL': 'RTL',
            '1CZUP': 'ZUP'
        }

        # Подключение к Bitrix24
        bitrix_connector = Bitrix24Connector()

        # SQL-запрос для получения систем
        sql_query = """
        SELECT s.name
        FROM access_rights ar
        JOIN departments d ON ar.department_id = d.id
        JOIN positions p ON ar.position_id = p.id
        JOIN systems s ON ar.system_id = s.id
        WHERE d.name = %s
          AND p.name = %s;
        """

        # SQL-запрос для получения user_type
        query_get_user_type = """
        SELECT ar.user_type 
        FROM access_rights ar
        JOIN departments d ON ar.department_id = d.id
        JOIN positions p ON ar.position_id = p.id
        WHERE d.name = %s
            AND p.name = %s
        LIMIT 1;
        """

        try:
            with self.connection.cursor() as cursor:
                # Проверка существования отдела
                cursor.execute("SELECT id FROM departments WHERE name = %s;", (department,))
                department_exists = cursor.fetchone()

                if not department_exists:
                    bitrix_connector.send_msg_error(f'Ошибка: отдел {department} не найден в БД.')
                    return flags

                # Проверка существования должности в отделе
                cursor.execute("""
                    SELECT id FROM positions 
                    WHERE name = %s 
                    AND id IN (SELECT position_id FROM access_rights WHERE department_id = (SELECT id FROM departments WHERE name = %s));
                """, (position, department))
                position_exists = cursor.fetchone()

                if not position_exists:
                    bitrix_connector.send_msg_error(f'Ошибка: должность {position} в отделе {department} не найдена в БД.')
                    return flags

                # Получение доступных систем
                cursor.execute(sql_query, (department, position))
                result = cursor.fetchall()
                for row in result:
                    system_name = row[0]
                    if system_name in system_mapping:
                        mapped_name = system_mapping[system_name]
                        flags[mapped_name] = True

                # Получение user_type
                cursor.execute(query_get_user_type, (department, position))
                result = cursor.fetchone()
                if result:
                    if result[0] == "ФИО":
                        flags['Normal_account'] = True
                    elif result[0] == "Должность+Магазин":
                        flags['Shop_account'] = True

        except psycopg2.Error as e:
            bitrix_connector.send_msg_error("Ошибка выполнения SQL-запроса:", e)

        return flags


