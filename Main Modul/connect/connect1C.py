import configparser

class Connector1C:
    def __init__(self):
        self.config = configparser.ConfigParser()
        self.config.read('connect_domain.ini')
        self.url_create = self.config.get('INFO1C', 'url_create')
        self.url_changes = self.config.get('INFO1C', 'url_changes')
        self.url_block = self.config.get('INFO1C', 'url_block')

    def getUrlCreate(self):
        return self.url_create

    def getUrlChanges(self):
        return self.url_changes

    def getUrlBlock(self):
        return self.url_block