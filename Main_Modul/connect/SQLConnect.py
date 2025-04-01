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

        query = """
            SELECT 
                d.name AS department,
                p.name AS position,
                ar.user_type AS user_type,
                MAX(CASE WHEN s.name = 'AD' THEN 1 ELSE 0 END) AS AD,
                MAX(CASE WHEN s.name = 'BITRIX24' THEN 1 ELSE 0 END) AS BX24,
                MAX(CASE WHEN s.name = 'SM_BINUU00' THEN 1 ELSE 0 END) AS SM_GEN,
                MAX(CASE WHEN s.name = 'SM_LOCAL' THEN 1 ELSE 0 END) AS SM_LOCAL,
                MAX(CASE WHEN s.name = '1CERP' THEN 1 ELSE 0 END) AS ERP,
                MAX(CASE WHEN s.name = '1CRTL' THEN 1 ELSE 0 END) AS RTL,
                MAX(CASE WHEN s.name = '1CZUP' THEN 1 ELSE 0 END) AS ZUP
            FROM access_rights ar
            JOIN departments d ON ar.department_id = d.id
            JOIN positions p ON ar.position_id = p.id
            JOIN systems s ON ar.system_id = s.id
            WHERE d.name = %s AND p.name = %s
            GROUP BY d.name, p.name, ar.user_type;
        """

        try:
            with self.connection.cursor() as cursor:
                cursor.execute(query, (department, position))
                result = cursor.fetchone()

                if not result:
                    # Добавляем отдел
                    cursor.execute("SELECT id FROM departments WHERE name = %s;", (department,))
                    dept_id = cursor.fetchone()
                    if not dept_id:
                        cursor.execute("INSERT INTO departments(name) VALUES (%s) RETURNING id;", (department,))
                        dept_id = cursor.fetchone()
                        log.info(f'Добавлен новый отдел: "{department}"')

                    # Добавляем должность
                    cursor.execute("SELECT id FROM positions WHERE name = %s;", (position,))
                    pos_id = cursor.fetchone()
                    if not pos_id:
                        cursor.execute("INSERT INTO positions(name) VALUES (%s) RETURNING id;", (position,))
                        pos_id = cursor.fetchone()
                        log.info(f'Добавлена новая должность: "{position}"')

                    self.connection.commit()
                    bitrix_connector.send_msg(
                        f'⚠ Нет настроек доступа для: Отдел "{department}", Должность "{position}". Запись добавлена, необходимо настроить права.')

                    # отправка сообщения лично
                    bitrix_connector.send_msg_user(
                        f'⚠️ Нет настроек доступа для: Отдел "{department}", Должность "{position}". Запись добавлена, необходимо настроить права.',
                        '135')
                    return flags

                # Распаковка результата
                (
                    dept_name, pos_name, user_type,
                    ad, bx24, sm_gen, sm_local,
                    erp, rtl, zup
                ) = result

                flags['AD'] = bool(ad)
                flags['BX24'] = bool(bx24)
                flags['SM_GEN'] = bool(sm_gen)
                flags['SM_LOCAL'] = bool(sm_local)
                flags['ERP'] = bool(erp)
                flags['RTL'] = bool(rtl)
                flags['ZUP'] = bool(zup)

                if user_type == "ФИО":
                    flags['Normal_account'] = True
                elif user_type == "Должность+Магазин":
                    flags['Shop_account'] = True

        except psycopg2.Error as e:
            bitrix_connector.send_msg_error(f"Ошибка выполнения SQL-запроса: {e}")

        return flags


