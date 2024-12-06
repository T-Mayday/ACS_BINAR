import random
import string
import secrets

# Подключение ldapConnect
from connect.ldapConnect import ActiveDirectoryConnector
connector = ActiveDirectoryConnector()

from message.message import log

# Матрица перевода
transliteration_dict = {
            'А': 'A', 'К': 'K', 'Х': 'KH',
            'Б': 'B', 'Л': 'L', 'Ц': 'TS',
            'В': 'V', 'М': 'M', 'Ч': 'CH',
            'Г': 'G', 'Н': 'N', 'Ш': 'SH',
            'Д': 'D', 'О': 'O', 'Щ': 'SHCH',
            'Е': 'E', 'Ё': 'E', 'Ж': 'ZH',
            'П': 'P', 'Ъ': '',
            'Р': 'R', 'Ы': 'Y',
            'З': 'Z', 'Т': 'T', 'Э': 'E',
            'И': 'I', 'У': 'U',
            'Й': 'Y', 'Ф': 'F',
            'С': 'S', 'Ь': '',
            'Ю': 'YU', 'Я': 'YA',
}

# Класс создания нужных атрибутов
class Person:
    def __init__(self, firstname, lastname, surname):
        self.firstname = firstname
        self.lastname = lastname
        self.surname = self.check_surname(surname)
        self.simple_login = self.create_simple_login()
        self.long_login = self.create_long_login()
        self.full_login = self.create_full_login()
        self.sm_login = self.create_sm_login()
        self.sm_login_login = self.create_sm_long_login()
        self.sm_full_login = self.create_sm_full_login()
        self.password = self.generate_password(8)
        self.full_name = self.full_name()

    def full_name(self):
        return self.lastname + ' ' + self.firstname + ' ' + self.surname

    def check_surname(self, surname):
        return surname if surname else 'Нету'

    def custom_transliterate(self, text):
        result = []
        for item in text:
            if item.upper() in transliteration_dict:
                result.append(transliteration_dict[item.upper()].upper() if item.isupper() else transliteration_dict[item.upper()].lower())
            else:
                result.append(item)
        return ''.join(result)

    def create_simple_login(self):
        firstname = self.custom_transliterate(self.firstname)
        lastname = self.custom_transliterate(self.lastname)
        if firstname and lastname:
            return f"{firstname[0].lower()}.{lastname.lower()}"
        return None

    def create_long_login(self):
        firstname = self.custom_transliterate(self.firstname)
        lastname = self.custom_transliterate(self.lastname)
        if firstname and lastname:
            return f"{firstname.lower()}.{lastname.lower()}"
        return None

    def create_full_login(self):
        firstname = self.custom_transliterate(self.firstname)
        lastname = self.custom_transliterate(self.lastname)
        surname = self.custom_transliterate(self.surname)
        if firstname and lastname and surname:
            return f"{firstname.lower()}.{surname[0].lower()}.{lastname.lower()}"
        return None

    def create_email(self, login):
        return f"{login}@{connector.getAdress()}" if login else None

    def create_sm_login(self):
        firstname = self.custom_transliterate(self.firstname)
        lastname = self.custom_transliterate(self.lastname)
        if firstname and lastname:
            return f"{firstname[0].lower()}_{lastname.lower()}"
        return None

    def create_sm_long_login(self):
        firstname = self.custom_transliterate(self.firstname)
        lastname = self.custom_transliterate(self.lastname)
        if firstname and lastname:
            return f"{firstname.lower()}_{lastname.lower()}"
        return None

    def create_sm_full_login(self):
        firstname = self.custom_transliterate(self.firstname)
        lastname = self.custom_transliterate(self.lastname)
        surname = self.custom_transliterate(self.surname)
        if firstname and lastname and surname:
            return f"{firstname.lower()}_{surname[0].lower()}_{lastname.lower()}"
        return None

    def generate_password(self, length):
        lowercase = ''.join([c for c in string.ascii_lowercase if c not in ['l', 'i']])
        uppercase = ''.join([c for c in string.ascii_uppercase if c not in ['L', 'I', 'O']])
        digits = ''.join([c for c in string.digits if c not in ['0', '1']])
        special_characters = '!@#$%^&*()-_=+'

        if length < 4:
            raise ValueError("Длина пароля должна быть как минимум 4 символа.")


        lower = secrets.choice(lowercase)
        upper = secrets.choice(uppercase)
        digit = secrets.choice(digits)
        special = secrets.choice(special_characters)

        all_characters = lowercase + uppercase + digits + special_characters
        remaining_length = length - 4
        remaining_chars = ''.join(secrets.choice(all_characters) for _ in range(remaining_length))


        password_list = list(lower + upper + digit + special + remaining_chars)
        secrets.SystemRandom().shuffle(password_list)
        password = ''.join(password_list)

        return password

    def transform_login(self, login):
        parts = login.split(".")
        firstname_translit = self.custom_transliterate(parts[0]) if parts else ""
        lastname_translit = self.custom_transliterate(parts[-1]) if len(parts) >= 2 else ""
        if len(parts) == 2:
            return f"{firstname_translit.lower()}_{lastname_translit.lower()}"
        elif len(parts) == 3:
            return f"{firstname_translit.lower()}_{parts[1].lower()}_{lastname_translit.lower()}"
        return login


# Словарь для шифрования и обратный
cipher_dict = {
    "0": "g",
    "1": "M",
    "2": "k",
    "3": "A",
    "4": "r",
    "5": "X",
    "6": "b",
    "7": "@",
    "8": "#",
    "9": "!"
}
reverse_cipher_dict = {v: k for k, v in cipher_dict.items()}


# Функция для шифрование ИНН
def encrypt_inn(inn):
    encrypted_inn = ''
    for digit in inn:
        if digit in cipher_dict:
            encrypted_inn += cipher_dict[digit]
        else:
            log.info(f"Ошибка шифрования: Символ '{digit}' присутсвует в ИНН ")
            encrypted_inn += digit
    return encrypted_inn