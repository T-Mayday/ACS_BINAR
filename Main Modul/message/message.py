import logging
from logging.handlers import TimedRotatingFileHandler
from datetime import datetime
# from message.message import send_msg, send_msg_error, log

# Подключение к Bitrix24
from connect.bitrixConnect import Bitrix24Connector

bitrix_connector = Bitrix24Connector()
bx24, tokens = bitrix_connector.connect()
chatID = bitrix_connector.chatID
chatadmID = bitrix_connector.chatadmID

today = datetime.now()
log = logging.getLogger("ADmain")
log.setLevel(logging.DEBUG)
#fh = logging.FileHandler('.\\logs\\{:%Y-%m-%d}.log'.format(today))
LOG_FILE = ".\\logs\\acs.log"
fh = TimedRotatingFileHandler(LOG_FILE, when='midnight')
formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
fh.setFormatter(formatter)
log.addHandler(fh)
log.debug('level=DEBUG')

def send_msg(msg):
    global chatID
    try:
        log.info(msg)
        bx24.refresh_tokens()
        res = bx24.call('im.message.add', {'DIALOG_ID': chatID, 'MESSAGE': msg, 'URL_PREVIEW': 'N'})
    except Exception as e:
        log.exception("Error sending message", e)

def send_msg_error(msg):
    global chatID
    try:
        log.exception(msg)
        bx24.refresh_tokens()
        res = bx24.call('im.message.add', {'DIALOG_ID': chatID, 'MESSAGE': msg, 'URL_PREVIEW': 'N'})
    except Exception as e:
        log.exception(e)

def send_msg_adm(msg):
    global chatID
    try:
        log.info(msg)
        bx24.refresh_tokens()
        res = bx24.call('im.message.add', {'DIALOG_ID': chatadmID, 'MESSAGE': msg, 'URL_PREVIEW': 'N'})
    except Exception as e:
        log.exception("Error sending message", e)
