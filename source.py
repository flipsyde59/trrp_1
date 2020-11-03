import httplib2
import apiclient
from oauth2client.service_account import ServiceAccountCredentials


def key_read():
    # Читаем ключи из файла
    from cryptography.fernet import Fernet
    f = open('key.bin', 'rb')
    cipher_key = f.read()
    f.close()
    fo = open('oauth.ini', 'rb')
    encrypted_text = fo.read()
    fo.close()
    cipher = Fernet(cipher_key)
    dec_text = cipher.decrypt(encrypted_text).decode("utf-8")
    name_temp_file = '1.json'
    fo = open(name_temp_file, 'w')
    fo.write(dec_text)
    fo.close()
    sac = ServiceAccountCredentials.from_json_keyfile_name(name_temp_file,
                                                           ['https://www.googleapis.com/auth/spreadsheets',
                                                            'https://www.googleapis.com/auth/drive'])
    import os
    os.remove(name_temp_file)
    return sac


def auth():
    import pickle
    import os.path
    from googleapiclient.discovery import build
    from google_auth_oauthlib.flow import InstalledAppFlow
    from google.auth.transport.requests import Request

    # If modifying these scopes, delete the file token.pickle.
    SCOPES = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']

    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'D:\\temp\\client_secret_892012293427-lqvv0dlp1n4s2buuqjlpgh55aeeao99n.apps.googleusercontent.com.json',
                SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('sheets', 'v4', credentials=creds)
    return service
#region Работа с книгой
def create_gs(name, service):
    spreadsheet = service.spreadsheets().create(body={
        'properties': {'title': name, 'locale': 'ru_RU'},
        'sheets': [{'properties': {'sheetType': 'GRID',
                                   'sheetId': 0,
                                   'title': 'Листок1',
                                   'gridProperties': {'rowCount': 100, 'columnCount': 15}}}]
    }).execute()
    return spreadsheet['spreadsheetId']


def open_access(spreadsheetId, httpAuth, mail='flipsyde59@mail.ru'):
    driveService = apiclient.discovery.build('drive', 'v3',
                                             http=httpAuth)  # Выбираем работу с Google Drive и 3 версию API
    access = driveService.permissions().create(
        fileId=spreadsheetId,
        body={'type': 'user', 'role': 'reader', 'emailAddress': mail},  # Открываем доступ на чтение
        fields='id'
    ).execute()


# Добавление листа
def add_sheet(service, spreadsheetId, title='Еще один лист'):
    results = service.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheetId,
        body=
        {
            "requests": [
                {
                    "addSheet": {
                        "properties": {
                            "title": title,
                            "gridProperties": {
                                "rowCount": 20,
                                "columnCount": 12
                            }
                        }
                    }
                }
            ]
        }).execute()


def del_sheet(service, spreadsheetId, sheet):
    results = service.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheetId,
        body=
        {
            "requests": [
                {
                    "deleteSheet": { # Deletes the requested sheet. # Deletes a sheet.
                    "sheetId": sheet['properties']['sheetId'], # The ID of the sheet to delete. If the sheet is of SheetType.DATA_SOURCE type, the associated DataSource is also deleted.
                                    },
                }
            ]
        }).execute()

# Получаем список листов, их ID и название
def get_lists(service, spreadsheetId, show=True):
    spreadsheet = service.spreadsheets().get(spreadsheetId=spreadsheetId).execute()
    sheetList = spreadsheet.get('sheets')
    if show:
        for sheet in sheetList:
            print(sheetList.index(sheet) + 1, sheet['properties']['sheetId'], sheet['properties']['title'])
    return sheetList


def cell_dev(service, spreadsheetId, sheet, cell, data):
    results = service.spreadsheets().values().batchUpdate(spreadsheetId=spreadsheetId, body={
        "valueInputOption": "USER_ENTERED",
        # Данные воспринимаются, как вводимые пользователем (считается значение формул)
        "data": [
            {"range": f"{sheet['properties']['title']}!{cell}",
             "majorDimension": "ROWS",  # Сначала заполнять строки, затем столбцы
             "values": [f'{data}']}
        ]
    }).execute()

#endregion
# region Работа с форматом
# Зададим ширину колонок. Функция batchUpdate может принимать несколько команд сразу, так что мы одним запросом
# установим ширину трех групп колонок. В первой и третьей группе одна колонка, а во второй - две.
def col_width(service, spreadsheetId, sheetId=0):
    results = service.spreadsheets().batchUpdate(spreadsheetId=spreadsheetId, body={
        "requests": [

            # Задать ширину столбца A: 20 пикселей
            {
                "updateDimensionProperties": {
                    "range": {
                        "sheetId": sheetId,
                        "dimension": "COLUMNS",  # Задаем ширину колонки
                        "startIndex": 0,  # Нумерация начинается с нуля
                        "endIndex": 1  # Со столбца номер startIndex по endIndex - 1 (endIndex не входит!)
                    },
                    "properties": {
                        "pixelSize": 20  # Ширина в пикселях
                    },
                    "fields": "pixelSize"  # Указываем, что нужно использовать параметр pixelSize
                }
            },

            # Задать ширину столбцов B и C: 150 пикселей
            {
                "updateDimensionProperties": {
                    "range": {
                        "sheetId": sheetId,
                        "dimension": "COLUMNS",
                        "startIndex": 1,
                        "endIndex": 3
                    },
                    "properties": {
                        "pixelSize": 150
                    },
                    "fields": "pixelSize"
                }
            },

            # Задать ширину столбца D: 200 пикселей
            {
                "updateDimensionProperties": {
                    "range": {
                        "sheetId": sheetId,
                        "dimension": "COLUMNS",
                        "startIndex": 3,
                        "endIndex": 4
                    },
                    "properties": {
                        "pixelSize": 200
                    },
                    "fields": "pixelSize"
                }
            }
        ]
    }).execute()


# Рисуем рамку
def border(service, spreadsheetId, sheetId=0):
    results = service.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheetId,
        body={
            "requests": [
                {'updateBorders': {'range': {'sheetId': sheetId,
                                             'startRowIndex': 1,
                                             'endRowIndex': 3,
                                             'startColumnIndex': 1,
                                             'endColumnIndex': 4},
                                   'bottom': {  # Задаем стиль для верхней границы
                                       'style': 'SOLID',  # Сплошная линия
                                       'width': 1,  # Шириной 1 пиксель
                                       'color': {'red': 0, 'green': 0, 'blue': 0, 'alpha': 1}},  # Черный цвет
                                   'top': {  # Задаем стиль для нижней границы
                                       'style': 'SOLID',
                                       'width': 1,
                                       'color': {'red': 0, 'green': 0, 'blue': 0, 'alpha': 1}},
                                   'left': {  # Задаем стиль для левой границы
                                       'style': 'SOLID',
                                       'width': 1,
                                       'color': {'red': 0, 'green': 0, 'blue': 0, 'alpha': 1}},
                                   'right': {  # Задаем стиль для правой границы
                                       'style': 'SOLID',
                                       'width': 1,
                                       'color': {'red': 0, 'green': 0, 'blue': 0, 'alpha': 1}},
                                   'innerHorizontal': {  # Задаем стиль для внутренних горизонтальных линий
                                       'style': 'SOLID',
                                       'width': 1,
                                       'color': {'red': 0, 'green': 0, 'blue': 0, 'alpha': 1}},
                                   'innerVertical': {  # Задаем стиль для внутренних вертикальных линий
                                       'style': 'SOLID',
                                       'width': 1,
                                       'color': {'red': 0, 'green': 0, 'blue': 0, 'alpha': 1}}

                                   }}
            ]
        }).execute()


# Объединяем ячейки B1:D1
def unite_cells(service, spreadsheetId, sheetId=0):
    results = service.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheetId,
        body={
            "requests": [
                {'mergeCells': {'range': {'sheetId': sheetId,
                                          'startRowIndex': 0,
                                          'endRowIndex': 1,
                                          'startColumnIndex': 1,
                                          'endColumnIndex': 4},
                                'mergeType': 'MERGE_ALL'}}
            ]
        }).execute()


# Добавляем заголовок таблицы
def header(service, spreadsheetId, sheetId=0):
    results = service.spreadsheets().values().batchUpdate(spreadsheetId=spreadsheetId, body={
        "valueInputOption": "USER_ENTERED",
        # Данные воспринимаются, как вводимые пользователем (считается значение формул)
        "data": [
            {"range": "Лист номер один!B1",
             "majorDimension": "ROWS",  # Сначала заполнять строки, затем столбцы
             "values": [["Заголовок таблицы"]
                        ]}
        ]
    }).execute()


def data_about_cells(service, spreadsheetId, range):
    ranges = ["Лист номер один!C2:C2"]  #

    results = service.spreadsheets().get(spreadsheetId=spreadsheetId,
                                         ranges=ranges, includeGridData=True).execute()
    print('Основные данные')
    print(results['properties'])
    print('\nЗначения и раскраска')
    print(results['sheets'][0]['data'][0]['rowData'])
    print('\nВысота ячейки')
    print(results['sheets'][0]['data'][0]['rowMetadata'])
    print('\nШирина ячейки')
    print(results['sheets'][0]['data'][0]['columnMetadata'])


# endregion

# чтение
def read(service, spreadsheetId, sheet):
    ranges = [f"{sheet['properties']['title']}!A1:J100"]  #

    results = service.spreadsheets().values().batchGet(spreadsheetId=spreadsheetId,
                                                       ranges=ranges,
                                                       valueRenderOption='FORMATTED_VALUE',
                                                       dateTimeRenderOption='FORMATTED_STRING').execute()
    try:
        sheet_values = results['valueRanges'][0]['values']
        for row in sheet_values:
            print('|'.join(row))
    except KeyError:
        print('Лист пуст')


serv = auth()
flag = True
credentials = key_read()  # Чтение зашифрованного файла
httpAuth = credentials.authorize(httplib2.Http())  # Авторизуемся в системе

while flag:
    service = serv # apiclient.discovery.build('sheets', 'v4', http=httpAuth)  # Выбираем работу с таблицами и 4 версию API
    create = input('Создать новую таблицу? да/нет: ')
    spreadsheetId=''
    if 'Д' in create.upper():
        name = input('Введите заголовок для таблицы: ')
        spreadsheetId = create_gs(name, service)
        print('https://docs.google.com/spreadsheets/d/' + spreadsheetId)
    elif 'Н' in create.upper():
        exist = '1LAhGFMW4ApvF_sEt6k5_8_1KhmuuWZejrbZcsCqVY0I'
        spreadsheetId = input('Вставьте id существующей таблицы: ')
        if spreadsheetId == '0':
            spreadsheetId = exist
        else:
            import requests
            if requests.head('https://docs.google.com/spreadsheets/d/' + spreadsheetId, allow_redirects=True).status_code != 200:
                print('Такого файла не существует, начнем сначала.')
                continue
    else:
        print('Вы ввели что-то не то. Попробуйте ещё раз')
        continue
    access = input('Предоставить доступ к этой таблице? да/нет: ')
    if 'Д' in access.upper():
        email_acc = input('Введите gmail: ')
        open_access(spreadsheetId, httpAuth, email_acc)
    print('список листов таблицы:\n№   id         title')
    sh_l = get_lists(service, spreadsheetId)
    num_l = int(input('Введите порядковый номер интересующего листа (0 - создать новый лист): '))
    if num_l == 0:
        name_sheet = input('Введите имя нового листа: ')
        add_sheet(service, spreadsheetId, name_sheet)
        sh_l = get_lists(service, spreadsheetId, False)
    sheet = sh_l[num_l - 1]
    print("Выбран лист:", sheet['properties']['sheetId'], sheet['properties']['title'])
    print('Данные на листе:')
    read(service, spreadsheetId, sheet)
    in_up = input('Внести/изменить данные? да/нет: ')
    if 'Д' in in_up.upper():
        cell = input('Введите индекс ячейки, в которой хотите внести изменения: ')
        data = input('Ведите данные: ')
        cell_dev(service, spreadsheetId, sheet, cell, data)
        print('Данные внесены')
        print('Данные на листе:')
        read(service, spreadsheetId, sheet)
    del_l = input('Вы хотите удалить какой-либо лист? да/нет: ')
    if 'Д' in del_l.upper():
        print('список листов таблицы:\n№   id         title')
        sh_l = get_lists(service, spreadsheetId)
        num_l = int(input('Введите порядковый номер листа который нужно удалить (0 - не удалять лист): '))
        if num_l!=0:
            del_sheet(service, spreadsheetId, sh_l[num_l-1])
    cont = input("продолжить? да/нет: ")
    if 'Н' in cont.upper():
        flag = False
        delete_f = input('Удалить файл? да/нет: ')
        if 'Д' in delete_f.upper():
            service.files().delete(fileId=spreadsheetId).execute()
            print('Файл удалён')
