import ldap
from ldap.filter import escape_filter_chars

# Подключение файла сообщения
from message.message import send_msg_error, log

# Подключение BitrixConnect
from connect.bitrixConnect import Bitrix24Connector
bitrix_connector = Bitrix24Connector()
bx24, tokens = bitrix_connector.connect()


def search_in_AD(mail, conn, base_dn):
    search_filter = f"(mail={mail})"
    # search_filter = "(employeeID=%s)" % (ldap.filter.escape_filter_chars(INN))
    try:
        result = conn.search_s(base_dn, ldap.SCOPE_SUBTREE, search_filter)
        if result and len(result) > 0:
            user_dn, attributes = result[0]
            user_account_control = attributes.get('userAccountControl', [b''])[0]
            uac_value = int(user_account_control.decode())
            is_active = not (uac_value & 0x0002)

            pager = attributes.get('pager',[b''])[0].decode('utf-8')
            if is_active:
                return result
            else:
                result = []
                return result
        else:
            result = []
            return result
    except Exception as e:
        send_msg_error(f'LDAP Ошибка поиcка по mail: {search_filter} {str(e)} ')
        result = []
        return result

def search_pager(mail, conn, base_dn):
    search_filter = f"(mail={mail})"
    # search_filter = "(employeeID=%s)" % (ldap.filter.escape_filter_chars(INN))
    try:
        result = conn.search_s(base_dn, ldap.SCOPE_SUBTREE, search_filter)
        if result and len(result) > 0:
            user_dn, attributes = result[0]

            pager = attributes.get('pager',[b''])[0].decode('utf-8')
            if pager:
                return pager
            else:
                return []
    except Exception as e:
        send_msg_error(f'LDAP Ошибка поиcка по mail: {search_filter} {str(e)} ')
        result = []
        return result


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


# Функция для получения всех пользователей
def get_all_users():
    try:
        bx24.refresh_tokens()
        result = bx24.call("user.get")
        return result['result']
    except Exception as e:
        log.error(f"BX24. Ошибка при получении пользователей: {e}")
        return None

def search_bx(last_name,name,second_name):
    try:
        bx24.refresh_tokens()
        result = bx24.call("user.get", {"LAST_NAME": last_name,"NAME": name,"SECOND_NAME" : second_name})
        if result.get('result'):
            r = result.get('result')[0]
            return r.get('ID')
        if result.get('error'):
            send_msg_error(f"BX24. Пользователь с ФИО '{last_name} {name} {second_name}' не найден. {result.get('error')[0]}")
            return None
    except Exception as e:
        send_msg_error(f"BX24. Ошибка при получении пользователей: {last_name} {name} {second_name} {e}")
        return None

def search_email_bx(email):
    try:
        bx24.refresh_tokens()
        result = bx24.call("user.get", {"EMAIL": email})
        if result.get('result'):
            r = result.get('result')[0]
            return r.get('ID')
        if result.get('error'):
            send_msg_error(f"BX24. Пользователь с {email} не найден. {result.get('error')[0]}")
            return None
    except Exception as e:
        send_msg_error(f"BX24. Ошибка при получении пользователей: {e}")
        return None
