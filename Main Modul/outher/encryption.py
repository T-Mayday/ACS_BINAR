import hashlib

from message.message import log

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

def encrypt_inn_new(inn):
    if not isinstance(inn, str) or not inn.isdigit():
        raise ValueError("ИНН должен быть строкой, содержащей только цифры.")
    hash_object = hashlib.sha256(inn.encode('utf-8'))
    encrypted_inn = hash_object.hexdigest()[:10]
    return encrypted_inn
