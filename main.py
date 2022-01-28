from functions import *

init_cells()
upd_deadlines(2020, 10, [4, 6, 9, 11, 13])
upd_kw(["KEYWORDS"])

for i in range(3, n):
    upd_cells(i)
    row_in_form, prev_number = 0, '06'
    init_persons(prev_number)
    req_error, error_sheet_id = {"requests": []}, get_sheetId(spreadsheet_id_form[i])
    number_index = 3
    for person in get_form(i):
        row_in_form += 1
        number = person[number_index]
        if number == prev_number:
            manage_person(i, person, number, req_error, error_sheet_id, row_in_form)
        else:
            if number > '06':
                continue
            upd_group(i, prev_number)
            if number < '04':
                break
            init_persons(number)
            manage_person(i, person, number, req_error, error_sheet_id, row_in_form)
        prev_number = number
    error_color(i, req_error)
    check_deadline(i)

#correct_format()
