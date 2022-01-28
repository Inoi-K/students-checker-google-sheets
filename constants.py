import httplib2
import apiclient
from oauth2client.service_account import ServiceAccountCredentials

# Файл, полученный в Google Developer Console
CREDENTIALS_FILE = 'creds.json'

# Авторизуемся и получаем service — экземпляр доступа к API
credentials = ServiceAccountCredentials.from_json_keyfile_name(
    CREDENTIALS_FILE,
    ['https://www.googleapis.com/auth/spreadsheets',
     'https://www.googleapis.com/auth/drive'])
httpAuth = credentials.authorize(httplib2.Http())
service = apiclient.discovery.build('sheets', 'v4', http=httpAuth)

# ID Google Sheets документа (можно взять из его URL)
# -------------------------------WRITE---------------------------------------
spreadsheet_id_w = 'SPREADSHEET_ID'

# -------------------------------FORM----------------------------------------
spreadsheet_id_form = ['SPREADSHEET_FORM_ID',
                       'SPREADSHEET_FORM_ID',
                       'SPREADSHEET_FORM_ID',
                       'SPREADSHEET_FORM_ID',
                       'SPREADSHEET_FORM_ID']


ids = {'04': 134458922, '05': 606282152, '06': 295058043}
numbers = ['01', '02', '03', '04', '05', '06', '07', '08', '09', '10', '11', '12']
cells, deadlines, keywords, persons = {}, [], [], []
n = len(spreadsheet_id_form)

black = {
    "red": 0,
    "green": 0,
    "blue": 0,
    "alpha": 1
}
borders = {
    "top": {
        "style": "SOLID",
        "color": black
    },
    "bottom": {
        "style": "SOLID",
        "color": black
    },
    "left": {
        "style": "SOLID",
        "color": black
    },
    "right": {
        "style": "SOLID",
        "color": black
    }
}
cell_format = {
    "requests": [
        {
            "repeatCell": {
                "cell": {
                    "userEnteredFormat": {
                        "borders": {
                            "top": {
                                "style": "SOLID",
                                "color": black
                            },
                            "bottom": {
                                "style": "SOLID",
                                "color": black
                            },
                            "left": {
                                "style": "SOLID",
                                "color": black
                            },
                            "right": {
                                "style": "SOLID",
                                "color": black
                            }
                        },
                        "horizontalAlignment": "CENTER",
                        "textFormat": {
                            "bold": True
                        }
                    }
                },
                "range": {
                },
                "fields": "userEnteredFormat"
            }
        }
    ]
}
green = {
    "userEnteredFormat": {
        "backgroundColor": {
            "red": 0.72,
            "green": 0.88,
            "blue": 0.8,
            "alpha": 1
        },
        "borders": borders,
        "horizontalAlignment": "CENTER",
        "textFormat": {
            "bold": True
        }
    }
}
yellow = {
    "userEnteredFormat": {
        "backgroundColor": {
            "red": 0.99,
            "green": 0.91,
            "blue": 0.7,
            "alpha": 1
        },
        "borders": borders,
        "horizontalAlignment": "CENTER",
        "textFormat": {
            "bold": True
        }
    }
}
red_error = {
    "userEnteredFormat": {
        "backgroundColor": {
            "red": 1,
            "green": 0,
            "blue": 0,
            "alpha": 1
        }
    }
}


class Person:
    def __init__(self, point='0.0', color=None):
        self.point = point
        self.color = color

    def upd(self, point, color):
        self.point = point
        self.color = color
