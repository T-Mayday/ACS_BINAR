import logging
from datetime import datetime

# Подключение к Bitrix24
from connect.bitrixConnect import Bitrix24Connector

bitrix_connector = Bitrix24Connector()
bx24, tokens = bitrix_connector.connect()
chatID = bitrix_connector.chatID

today = datetime.now()
log = logging.getLogger("ADmain")
log.setLevel(logging.DEBUG)
fh = logging.FileHandler('.\\logs\\{:%Y-%m-%d}.log'.format(today))
formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
fh.setFormatter(formatter)
log.addHandler(fh)
log.debug('level=DEBUG')

def send_msg(msg):
    global chatID
    try:
        res = bx24.call('im.message.add', {'DIALOG_ID': chatID, 'MESSAGE': msg, 'URL_PREVIEW': 'N'})
    except Exception as e:
        log.exception("Error sending message", e)

def send_msg_error(msg):
    global chatID
    try:
        log.exception(msg)
        res = bx24.call('im.message.add', {'DIALOG_ID': chatID, 'MESSAGE': msg, 'URL_PREVIEW': 'N'})
    except Exception as e:
        log.exception(e)

