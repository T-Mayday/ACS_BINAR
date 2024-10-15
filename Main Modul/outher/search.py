# Подключение BitrixConnect
from connect.bitrixConnect import Bitrix24Connector
bitrix_connector = Bitrix24Connector()
bx24, tokens = bitrix_connector.connect()


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
                bitrix_connector.send_msg_error(f'Ошибка поискa по info.xlsx: должность {position} отдела {department} не найден.')
        else:
            bitrix_connector.send_msg_error(f'Ошибка поискa по info.xlsx: отдел {department} не найден.')
    return flags

