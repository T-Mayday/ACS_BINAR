import pandas as pd
import psycopg2

# Параметры подключения к PostgreSQL
DB_CONFIG = {
    "dbname": "system_access",
    "user": "пользоватль",
    "password": "пароль",
    "host": "localhost",
    "port": "5432",
    "client_encoding": "UTF8"
}

# путь до файла info.xlsx
file = "путь до таблицы"
access_rights_query = "insert_access_rights.sql"  # Файл для сохранения SQL-запросов

df = pd.read_excel(file)

# Подключение к БД
conn = psycopg2.connect(**DB_CONFIG)
cursor = conn.cursor()

# Получаем ID для отделов, должностей и систем
cursor.execute("SELECT id, name FROM departments;")
departments = {name: id for id, name in cursor.fetchall()}

cursor.execute("SELECT id, name FROM positions;")
positions = {name: id for id, name in cursor.fetchall()}

cursor.execute("SELECT id, name FROM systems;")
systems = {name: id for id, name in cursor.fetchall()}

cursor.close()
conn.close()

with open(access_rights_query, "w", encoding="utf-8") as sql_file:
    sql_file.write("-- INSERT-запросы для access_rights\n\n")

    last_department_name = None  # Переменная для запоминания отдела

    for index, row in df.iterrows():
        # Если отдел не пустой, запоминаем его, иначе используем предыдущий
        if pd.notna(row["Отдел"]):
            last_department_name = str(row["Отдел"]).strip().upper()

        if not last_department_name:
            print(f" Ошибка в строке {index + 2}: Отдел не определен!")
            continue  # Пропускаем строку

        department_id = departments.get(last_department_name)
        position_name = str(row["Должность"]).strip()
        user_type = str(row["USERTYPE"]).strip()
        if user_type == 'nan':
            user_type = 'ФИО'

        position_id = positions.get(position_name)

        if department_id is None:
            # print(f"Пропущена строка {index + 2}: Отдел '{last_department_name}' не найден в БД!")
            continue  # Пропускаем строку

        if position_id is None:
            # print(f" Пропущена строка {index + 2}: Должность '{position_name}' не найдена в БД!")
            continue  # Пропускаем строку

        # Проверяем, в каких системах стоит "1"
        values = []
        for system_name, value in row.iloc[5:].items():
            if value == 1:
                system_id = systems.get(system_name.strip().upper())
                if system_id:
                    values.append(f"({department_id}, {position_id}, {system_id}, '{user_type}')")
        if values:
            sql_file.write(f"INSERT INTO access_rights (department_id, position_id, system_id, user_type) VALUES {', '.join(values)} ON CONFLICT DO NOTHING;\n")

print(f" SQL-запросы сохранены в файл {access_rights_query}")
