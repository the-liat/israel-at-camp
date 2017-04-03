import shutil
from collections import OrderedDict

import win32com.client as win32
import os

xl_app = None


def open_excel():
    global xl_app
    xl_app = win32.gencache.EnsureDispatch('Excel.Application')
    xl_app.DisplayAlerts = False


def close_excel():
    xl_app.Application.Quit()


def create_camp_workbook_from_template(camp_name):
    f = r'C:\Users\Liat\Google Drive\UJA Toronto Israel at Camp\Deliverables\Camp Reports\UJAToronto_Israel@Camp_CampReports_TEMPLATE LS 20170331.xlsx'
    dir_name = r'C:\Users\Liat\Google Drive\UJA Toronto Israel at Camp\Deliverables\Camp Reports\Excel sheets for reports production'
    file_name = os.path.join(dir_name, 'UJAToronto.Israel@Camp.CampReports. {}.xlsx'.format(camp_name))
    if not os.path.exists(file_name):
        shutil.copyfile(f, file_name)
    wb = xl_app.Workbooks.Open(file_name)
    return wb, file_name


def populate_exhibit(exhibit_number, tables, workbook, camp):
    func = globals()['populate_exhibit{}'.format(exhibit_number)]
    func(tables, workbook, camp)


def get_index(col, item):
    return col.index(item) if item in col else -1


def populate_excel_ex_1_3_5_7(ws, index_information, t, tables, target_columns):
    campers_vs_staff = t.get_col_by_index(1)
    campers_row_index = get_index(campers_vs_staff, 'Campers')
    staff_row_index = get_index(campers_vs_staff, 'Staff')
    stakeholders_col = []
    if campers_row_index > -1:
        stakeholders_col.append((campers_row_index + 1, target_columns[0]))
    if staff_row_index > -1:
        stakeholders_col.append((staff_row_index + 1, target_columns[1]))
    for d in index_information:
        t = tables[d['table_index']]
        for row_index, target_col in stakeholders_col:
            for col_index, target_row in d['from_col_to_target_row']:
                row = t.get_row(row_index)
                cell = "{}{}".format(target_col, target_row)
                ws.Range(cell).Value = row[col_index]


def populate_exhibit1(tables, workbook, camp):
    ws = workbook.Worksheets('All.Exhibits')
    t = tables[1]  # age
    campers_vs_staff = t.get_col_by_index(0)
    campers_row_index = get_index(campers_vs_staff, 'Campers (including CITs)')
    staff_row_index = get_index(campers_vs_staff, 'Staff')
    if campers_row_index > -1:
        row = t.get_row(campers_row_index)
        age = row[1]
        ws.Range('C4').Value = age
    if staff_row_index > -1:
        row = t.get_row(staff_row_index)
        age = row[1]
        ws.Range('D4').Value = age
    t = tables[3]
    target_columns = ['C', 'D']
    index_information = [
        dict(table_name='gender',
             from_col_to_target_row=((3, 5), (4, 6)), table_index=3),
        dict(table_name='denomination',
             from_col_to_target_row=((3, 7), (4, 8), (5, 9), (6, 10)), table_index=4),
        dict(table_name='Attended day school',
             from_col_to_target_row=((4, 11),), table_index=5),
        dict(table_name='Attended overnight camp',
             from_col_to_target_row=((4, 12),), table_index=6),
        dict(table_name='Attended supplementary school',
             from_col_to_target_row=((4, 13),), table_index=7),
        dict(table_name='Attended youth group',
             from_col_to_target_row=((4, 14),), table_index=8),
        dict(table_name='Frequency of attending religious services',
             from_col_to_target_row=((3, 15), (4, 16), (5, 17), (6, 18), (7, 19), (8, 20)), table_index=9),
        dict(table_name='Previous visits to Israel',
             from_col_to_target_row=((3, 21), (4, 22), (5, 23)), table_index=10)
    ]
    populate_excel_ex_1_3_5_7(ws, index_information, t, tables, target_columns)


def populate_exhibit3(tables, workbook, camp):
    ws = workbook.Worksheets('All.Exhibits')
    t = tables[1]
    target_columns = ['B', 'C']
    index_information = [
        dict(table_name='How often do you hear Hebrew spoken around camp?',
             from_col_to_target_row=((3, 46), (4, 45), (5, 44), (6, 43), (7, 42)), table_index=1)
    ]
    populate_excel_ex_1_3_5_7(ws, index_information, t, tables, target_columns)


def populate_exhibit5(tables, workbook, camp):
    ws = workbook.Worksheets('All.Exhibits')
    t = tables[1]
    target_columns = ['B', 'C']
    index_information = [
        dict(table_name='How often do you have Israel programming at your camp?',
             from_col_to_target_row=((3, 65), (4, 64), (5, 63), (6, 62), (7, 61), (8, 60)), table_index=1)
    ]
    populate_excel_ex_1_3_5_7(ws, index_information, t, tables, target_columns)


def populate_exhibit7(tables, workbook, camp):
    ws = workbook.Worksheets('All.Exhibits')
    t = tables[1]
    target_columns = ['B', 'C']
    index_information = [
        dict(table_name='How would you rate the Israel programming at your camp?',
             from_col_to_target_row=((3, 85), (4, 84), (5, 83), (6, 82), (7, 81)), table_index=1)
    ]
    populate_excel_ex_1_3_5_7(ws, index_information, t, tables, target_columns)


def populate_excel_ex_2_8(ws, index_information, t, tables, target_rows):
    campers_vs_staff = t.get_col_by_index(1)
    campers_row_index = get_index(campers_vs_staff, 'Campers')
    staff_row_index = get_index(campers_vs_staff, 'Staff')
    stakeholders_col = []
    if campers_row_index > -1:
        stakeholders_col.append((campers_row_index + 1, target_rows[0]))
    if staff_row_index > -1:
        stakeholders_col.append((staff_row_index + 1, target_rows[1]))
    for d in index_information:
        t = tables[d['table_index']]
        for row_index, target_row in stakeholders_col:
            for col_index, target_col in d['from_col_to_target_col']:
                row = t.get_row(row_index)
                cell = "{}{}".format(target_col, target_row)
                ws.Range(cell).Value = row[col_index]


def populate_exhibit2(tables, workbook, camp):
    ws = workbook.Worksheets('All.Exhibits')
    t = tables[1]
    target_rows = [30, 29]
    index_information = [
        dict(table_name='How many Israelis are there at your camp?',
             from_col_to_target_col=((3, 'I'), (4, 'H'), (5, 'G')), table_index=1)
    ]
    populate_excel_ex_2_8(ws, index_information, t, tables, target_rows)


def populate_exhibit8(tables, workbook, camp):
    ws = workbook.Worksheets('All.Exhibits')
    t = tables[1]
    target_rows = [93, 92]
    index_information = [
        dict(table_name='Level of Israel Engagement',
             from_col_to_target_col=((3, 'I'), (4, 'H')), table_index=1)
    ]
    populate_excel_ex_2_8(ws, index_information, t, tables, target_rows)


def identify_colmuns_ex_4_6(t):
    campers_vs_staff = t.data[0]
    campers_col_index = get_index(campers_vs_staff, 'Campers (including CITs)')
    staff_col_index = get_index(campers_vs_staff, 'Staff')
    stakeholders_col = []
    target_col = []
    if campers_col_index > -1:
        stakeholders_col.append(campers_col_index)
        target_col.append('B')
    if staff_col_index > -1:
        stakeholders_col.append(staff_col_index)
        target_col.append('C')
    return stakeholders_col, target_col

def iterate_over_two_rows_at_a_time(rows, stakeholders_col):
    output = ([], [])
    for i, row in enumerate(rows):
        # Skip odd rows (module operator % returns 0 when i is even)
        if i % 2 != 0:
            continue
        label = row[1]
        for s in stakeholders_col:
            if label == 'Counsellors and':
                label = 'Counsellors and staff'
            elif label == 'Shabbat Services':
                label = 'Shabbat Services and/or Havdallah'
            elif label == 'Dining hall (Chedar':
                label = 'Dining hall (Chedar Ochel)'
            output_row = (label, row[s])
            output[s-3].append(output_row)
    return output


def sort_by_column_descending(data):
    data = [d for d in data if d[1] > 50.5]
    res = tuple(reversed(sorted(data, key=lambda r: r[1])))
    return res


def populate_exhibit4(tables, workbook, camp):
    ws = workbook.Worksheets('All.Exhibits')
    t = tables[1]
    columns, target_cols=identify_colmuns_ex_4_6(t)
    rows = t.data[2:-1]
    output_rows = iterate_over_two_rows_at_a_time(rows, columns)
    for col in columns:
        sorted_row = sort_by_column_descending(output_rows[col-3])
        target_row = 101
        target_col = target_cols[col-3]
        for r in sorted_row:
            label = r[0]
            percent = int(round(r[1]))
            cell_value = '{} ({}%)'.format(label, percent)
            cell = "{}{}".format(target_col, target_row)
            ws.Range(cell).Value = cell_value
            target_row += 1


def populate_exhibit6(tables, workbook, camp):
    ws = workbook.Worksheets('All.Exhibits')
    t = tables[1]
    columns, target_cols=identify_colmuns_ex_4_6(t)
    rows = t.data[2:-1]
    output_rows = iterate_over_two_rows_at_a_time(rows, columns)
    for col in columns:
        sorted_row = sort_by_column_descending(output_rows[col-3])
        target_row = 114
        target_col = target_cols[col-3]
        for r in sorted_row:
            label = r[0]
            percent = int(round(r[1]))
            cell_value = '{} ({}%)'.format(label, percent)
            cell = "{}{}".format(target_col, target_row)
            ws.Range(cell).Value = cell_value
            target_row += 1

def populate_exhibit9(tables, workbook, camp):
    ws = workbook.Worksheets('All.Exhibits')
    staff = ('D3', 'C41', 'C59', 'C69', 'C80', 'C100', 'C113')
    cell_value_staff = '(n={})'.format(camp['staff'])
    for s in staff:
        ws.Range(s).Value = cell_value_staff
    campers = ('C3', 'B41', 'B59', 'B69', 'B80', 'B100', 'B113')
    cell_value_campers = '(n={})'.format(camp['campers'])
    for c in campers:
        ws.Range(c).Value = cell_value_campers



"""
get_cell_value(row_index, col_index)
get_row(row_index)
get_col_by_index(col_index)
get_col_by_name(col_name)
"""
