import pandas as pd

# путь до файла info.xlsx
file = "путь"
departments_query = "insert_departments.sql"  # Файл для сохранения SQL-запросов по отделам
positions_query = "insert_positions.sql" # Файл для сохранения SQL-запросов по должностям
# systems_query = 'insert_systems.sql' # Файл для сохранения SQL-запросов по системам

df = pd.read_excel(file)

with open(departments_query, "w", encoding="utf-8") as sql_file:
    sql_file.write("-- INSERT-запросы для departments\n\n")
    for index, row in df.iterrows():
        if pd.notna(row["Отдел"]):
            department_name = str(row["Отдел"]).strip()
            sql_file.write(f"INSERT INTO departments (name) VALUES ('{department_name}') \n")

with open(positions_query, "w", encoding="utf-8") as sql_file:
    sql_file.write("-- INSERT-запросы для positions\n\n")
    positions = []
    for index, row in df.iterrows():
        if pd.notna(row["Должность"]):
            positions_name = str(row["Должность"]).strip()
            positions.append(positions_name)

    # функция фильтрации одинаковых должностей
    def remove_duplicates(nested_list):
        unique_list = []
        for element in nested_list:
            if element not in unique_list:
                unique_list.append(element)
        return unique_list
    filtered_positions = remove_duplicates(positions)
    for name in filtered_positions:
        sql_file.write(f"INSERT INTO positions (name) VALUES ('{name}') \n")




print(f"SQL-запросы сохранены в файл {departments_query, positions_query}")
