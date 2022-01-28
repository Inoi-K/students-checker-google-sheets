import re
import datetime
from constants import *


def to_time(timestamp):
    timestamp = re.split(r'[/ :.]', timestamp)
    timestamp[0], timestamp[2] = timestamp[2], timestamp[0]
    for i in range(1, 6):
        if timestamp[i][0] == '0':
            timestamp[i] = timestamp[i][1]
    return datetime.datetime(*(int(i) for i in timestamp))


def init_cells():
    for num in ids:
        cells_t = service.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id_w,
            range="{}!B4:B34".format(num),
            majorDimension="COLUMNS"
        ).execute()
        cells[num] = cells_t['values']


def upd_cells(i):
    letter_for_range = chr(69 + 2 * i)
    if letter_for_range > 'Z':
        letter_for_range = 'A' + chr(65 + 2 * (i - 11))
    for num in ids:
        if len(cells[num]) > 1:
            cells[num].pop()
        cells_t = service.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id_w,
            range="{}!{}4:{}35".format(num, letter_for_range, letter_for_range),
            majorDimension="COLUMNS"
        ).execute()
        cells[num].append(*cells_t['values'])


def upd_deadlines(year, month, days):
    for i in range(len(days)):
        deadlines.append(datetime.datetime(year, month, days[i], 23, 59, 59))


def check_deadline(i):
    letter_for_range = chr(69 + 2 * i)
    if letter_for_range > 'Z':
        letter_for_range = 'A' + chr(65 + 2 * (i - 11))
    if datetime.datetime.today() > deadlines[i]:
        upd_cells(i)
        for number in ids:
            start = 0
            for index in range(len(cells[number][1])):
                if cells[number][1][index]:
                    if index - start:
                        service.spreadsheets().values().batchUpdate(
                            spreadsheetId=spreadsheet_id_w,
                            body={
                                "valueInputOption": "USER_ENTERED",
                                "data": [
                                    {
                                        "range": "{}!{}".format(number, letter_for_range + str(start + 4) + ':' + letter_for_range + str(index + 3)),
                                        "values": [
                                            [0] for i in range(index - start)
                                        ]
                                    }
                                ]
                            }
                        ).execute()
                    start = index + 1
            

def upd_kw(kws):
    for i in range(len(kws)):
        keywords.append(kws[i])


def get_form(i):
    local_index = 3 # if i < 13 else 4

    service.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet_id_form[i],
        body={
            "requests": [
                {
                    "sortRange": {
                        "range": {
                            "sheetId": get_sheetId(spreadsheet_id_form[i]),
                            "startRowIndex": 0,
                            "startColumnIndex": 0,
                        },
                        "sortSpecs": {
                            "sortOrder": "DESCENDING",
                            "dimensionIndex": local_index
                        }
                    }
                }
            ]
        }
    ).execute()

    values = service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id_form[i],
        range="A1:E99",
        majorDimension="ROWS"
    ).execute()
    if values:
        values['values'].pop(0)
    
    return values['values']


def error_color(i, req):
    if len(req["requests"]):
        service.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheet_id_form[i],
            body=req
        ).execute()
        

def name_check(name, number):
    names = re.split(r' ', name)
    if name in cells[number][0] or len(names) > 1 and names[1] + ' ' + names[0] in cells[number][0]:
        return True
    return False


def get_sheetId(spreadsheet_id):
    spreadsheet = service.spreadsheets().get(
        spreadsheetId=spreadsheet_id
    ).execute()
    return spreadsheet['sheets'][0]['properties']['sheetId']


def fill_range(sheetId, startRowIndex, endRowIndex, startColumnIndex, endColumnIndex):
    t = {
        "sheetId": sheetId,
        "startRowIndex": startRowIndex,
        "endRowIndex": endRowIndex,
        "startColumnIndex": startColumnIndex,
        "endColumnIndex": endColumnIndex
    }
    return t


def repeat_cell(color, sheetId, startRowIndex, endRowIndex, startColumnIndex, endColumnIndex):
    t = {
        "repeatCell": {
            "cell": color,
            "range": fill_range(sheetId, startRowIndex, endRowIndex, startColumnIndex, endColumnIndex),
            "fields": "userEnteredFormat"
        }
    }
    return t


def upd_group(i, number):
    letter_for_range = chr(69 + 2 * i)
    if letter_for_range > 'Z':
        letter_for_range = 'A' + chr(65 + 2 * (i - 11))
    service.spreadsheets().values().batchUpdate(
        spreadsheetId=spreadsheet_id_w,
        body={
            "valueInputOption": "USER_ENTERED",
            "data": [
                {
                    "range": "{}!{}4:{}34".format(number, letter_for_range, letter_for_range),
                    "values": [
                        [person.point.replace('.', ',')] for person in persons
                    ]
                }
            ]
        }
    ).execute()

    req_color = {"requests": []}
    for j in range(len(persons)):
        if persons[j].color is not None:
            req_color["requests"].append(repeat_cell(persons[j].color, ids[number], j + 3, j + 4, 2 * i + 4, 2 * i + 5))

    if len(req_color["requests"]):
        service.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheet_id_w,
            body=req_color
        ).execute()
        

    # req_color, prev_color, start_row, = {"requests": []}, green, 0
    # for j in range(31):
    #     if persons[j].color != prev_color and j - start_row or persons[j].point == '' or j == 30:
    #         req_color["requests"].append(
    #             repeat_cell(prev_color, ids[number], start_row + 3, j + 3, 2 * i + 4, 2 * i + 5))
    #         start_row = j
    #     prev_color = persons[j].color


def init_persons(number):
    persons.clear()
    for value in cells[number][1]:
        persons.append(Person(value))
    persons.pop()


def correct_format():
    for number in ids:
        cell_format["requests"][0]["repeatCell"]["range"] = fill_range(ids[number], 3, 34, 10, 27)
        service.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheet_id_w,
            body=cell_format
        ).execute()


def manage_person(i, person, number, req_error, error_sheet_id, row_in_form):
    name_index = 2 # if i < 13 else 3
    if (i >= 13 or person[4].lower().strip() == keywords[i]) and name_check(person[name_index].strip(), number):
        try:
            row_index = cells[number][0].index(person[name_index].strip())
        except:
            names = re.split(r' ', person[name_index].strip())
            row_index = cells[number][0].index(names[1] + ' ' + names[0])

        value = cells[number][1][row_index]
        if value == '' or value == '0' or value == '/':
            point, color, time = max(person[1].split(' / ')[0], '0.1'), green, to_time(person[0])
            # CHECK DEADLINE
            if time > deadlines[i] and value != '/':
                point, color = str(float(point) / 2), yellow
            persons[row_index].upd(point, color)
    else:
        req_error["requests"].append(repeat_cell(red_error, error_sheet_id, row_in_form, row_in_form + 1, 2, 3))
