# Default model for SAP automations, developed by Robert Aron Zimmermann, using Google AI Studio tuned prompt model
from sap_functions import SAP

default_language = 'PT'
login = open('sap_login.txt', 'r').readline().strip().split(',')
scheduled_execution = {'scheduled?': False, 'username': login[0], 'password': login[1], 'principal': '100'}
sap_window = 0

# Python Default Script
# Default model for SAP automations, developed by Robert Aron Zimmermann, using Google AI Studio tuned prompt model

# Solicitado por Karoline Luciani Fritsche
# Desenvolvido por Robert Aron Zimmermann

class Work:
    def __init__(self):
        self.sap = SAP(sap_window, scheduled_execution, default_language)


if __name__ == '__main__':
    work = Work()
