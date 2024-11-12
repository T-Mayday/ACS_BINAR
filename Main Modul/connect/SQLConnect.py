import pymysql
import configparser

from connect.bitrixConnect import Bitrix24Connector
bitrix_connector = Bitrix24Connector()

class DatabaseConnector:
    def __init__(self):
        self.config = configparser.ConfigParser()
        self.config.read('connect_domain.ini', encoding='utf-8')

        self.host = self.config.get('DATABASE', 'host')
        self.user = self.config.get('DATABASE', 'user')
        self.password = self.config.get('DATABASE', 'password')
        self.db = self.config.get('DATABASE', 'db')
        self.charset = self.config.get('DATABASE', 'charset')

        self.connection = self.connect()

    def connect(self):
        try:
            connection = pymysql.connect(
                host=self.host,
                user=self.user,
                password=self.password,
                db=self.db,
                charset=self.charset,
                cursorclass=pymysql.cursors.DictCursor
            )
            return connection
        except pymysql.MySQLError as e:
            bitrix_connector.send_msg_error("Ошибка подключения к базе данных:", e)
            return None

    def get_connection(self):
        return self.connection

    def close_conn(self):
        if self.connection:
            self.connection.close()

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

        sql_query = """
                    SELECT
                        a.ad,
                        a.bitrix24,
                        a.sm_gl,
                        a.sm_local,
                        a.1CERP,
                        a.1CRTL,
                        a.1CZUP,
                        CASE
                            WHEN a.user_type = 'ФИО' THEN 1
                            ELSE 0
                        END AS Normal_account,
                        CASE
                            WHEN a.user_type = 'Должность+Магазин' THEN 1
                            ELSE 0
                        END AS Shop_account
                    FROM
                        position p
                    LEFT JOIN
                        access a ON p.id = a.position_id
                    JOIN
                        department d ON p.department_id = d.id
                    WHERE
                        d.name = %s AND
                        p.name = %s;
                    """
        try:
            with self.connection.cursor() as cursor:
                cursor.execute(sql_query, (department, position))
                result = cursor.fetchone()

                if result:
                    flags['AD'] = bool(result['ad'])
                    flags['BX24'] = bool(result['bitrix24'])
                    flags['ZUP'] = bool(result['1CZUP'])
                    flags['RTL'] = bool(result['1CRTL'])
                    flags['ERP'] = bool(result['1CERP'])
                    flags['SM_GEN'] = bool(result['sm_gl'])
                    flags['SM_LOCAL'] = bool(result['sm_local'])
                    flags['Normal_account'] = bool(result['Normal_account'])
                    flags['Shop_account'] = bool(result['Shop_account'])
                else:
                    bitrix_connector.send_msg_error(f"Ошибка: должность {position} отдела {department} не найдена.")
        except pymysql.MySQLError as e:
            bitrix_connector.send_msg_error("Ошибка выполнения SQL-запроса:", e)

        finally:
            self.connection.close()
        return flags