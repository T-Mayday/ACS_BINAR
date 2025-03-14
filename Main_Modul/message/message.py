import logging
from logging.handlers import TimedRotatingFileHandler
from datetime import datetime


today = datetime.now()
log = logging.getLogger("binidm")
log.setLevel(logging.DEBUG)
#fh = logging.FileHandler('.\\logs\\{:%Y-%m-%d}.log'.format(today))
LOG_FILE = ".\\logs\\binidm.log"
#fh = TimedRotatingFileHandler(LOG_FILE, when='midnight')
fh = logging.StreamHandler()
formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
fh.setFormatter(formatter)
log.addHandler(fh)
log.debug('level=DEBUG')

