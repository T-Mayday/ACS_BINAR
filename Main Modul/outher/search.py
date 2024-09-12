import ldap

# Подключение файла сообщения
from message.message import send_msg_error, log

# Подключение BitrixConnect
from connect.bitrixConnect import Bitrix24Connector
bitrix_connector = Bitrix24Connector()
bx24, tokens = bitrix_connector.connect()


def search_in_AD(INN, conn, base_dn):
    search_filter = f"(employeeID={INN})"
    try:
        result = conn.search_s(base_dn, ldap.SCOPE_SUBTREE, search_filter)
        return result
    except Exception as e:
        send_msg_error(f'LDAP Ошибка посик по employeeID: {str(e)} {search_filter}')
        return []


def user_verification(df_roles, df_users):

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

    for _, user_row in df_users.iterrows():
        department = user_row['Отдел']
        position = user_row['Должность']


        department_index = df_roles.index[df_roles['Отдел'] == department].tolist()

        if department_index:
            department_row = df_roles.iloc[department_index[0]]

            position_index = df_roles.index[
                (df_roles['Должность'] == position) & (df_roles.index > department_index[0])
                ].tolist()

            if position_index:

                position_row = df_roles.iloc[position_index[0]]

                flags['AD'] |= position_row['AD'] == 1
                flags['BX24'] |= position_row['BITRIX24'] == 1
                flags['ZUP'] |= position_row['1CZUP'] == 1
                flags['RTL'] |= position_row['1CRTL'] == 1
                flags['ERP'] |= position_row['1CERP'] == 1
                flags['SM_GEN'] |= position_row['SM_BINUU00'] == 1
                flags['SM_LOCAL'] |= position_row['SM_LOCAL'] == 1

                if position_row['USERTYPE'] == "ФИО":
                    flags['Normal_account'] = True

                if position_row['USERTYPE'] == "Должность+Магазин":
                    flags['Shop_account'] = True
            else:
                send_msg_error(f'Ошибка поискa по info.xlsx: должность {position} отдела {department} не найден.')
        else:
            send_msg_error(f'Ошибка поискa по info.xlsx: отдел {department} не найден.')
    return flags


# Функция для поиска сотрудника в Bitrix24
def find_jobfriend(post_job, codeBX24):
    filter_params = {
        'FILTER': {
            'WORK_POSITION': post_job,
            'UF_DEPARTMENT': codeBX24
        }
    }
    employees = bx24.call('user.get', filter_params)
    if 'result' in employees and employees['result']:
        employee_info = employees['result'][0]
        return f"{employee_info['LAST_NAME']} {employee_info['NAME']} {employee_info['SECOND_NAME']}"
    return None


def search_login(login, conn, base_dn):
    search_filter = f'(sAMAccountName={login})'
    try:
        result = conn.search_s(base_dn, ldap.SCOPE_SUBTREE, search_filter)
        return result
    except Exception as e:
        log.error(f'Ошибка поиска логина - {str(e)}')
        return []

# Функция для получения всех пользователей
def get_all_users():
    try:
        result = bx24.call("user.get")
        return result['result']
    except Exception as e:
        log.error(f"BX24. Ошибка при получении пользователей: {e}")
        return None

def search_bx(last_name,name,second_name):
    try:
        result = bx24.call("user.get", {"LAST_NAME": last_name,"NAME": name,"SECOND_NAME" : second_name})
        if result.get('result'):
            r = result.get('result')[0]
            return r.get('ID')
        if result.get('error'):
            log.error(f"BX24. Пользователь с ФИО '{last_name} {name} {second_name}' не найден.")
            return None
    except Exception as e:
        log.error(f"BX24. Ошибка при получении пользователей: {e}")
        return None