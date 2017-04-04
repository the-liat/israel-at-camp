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


def get_values_from_table(s, table, cols_labels_and_target_rows_for_excel, ws):
    labels_from_spss = table.get_row(0)[3:-1]  # first row with all labels that need to match with the label_cols
    row = table.get_row(s[0])  # s[0] is index of the row where the data are
    for d in cols_labels_and_target_rows_for_excel:
        match = False
        for i, spss_label in enumerate(labels_from_spss):
            if spss_label == d['label']:
                match = True
                break
        if not match:
            continue
        target_row = d['row']
        value = row[i + 3]
        write_value_in_excel_cell(ws, value, target_row, s[1])


def get_stakeholders_from_rows_n_target_cols(table, target_columns):
    campers_vs_staff = table.get_col_by_index(1)
    campers_row_index = get_index(campers_vs_staff, 'Campers')
    staff_row_index = get_index(campers_vs_staff, 'Staff')
    stakeholders_col = []
    if campers_row_index > -1:
        stakeholders_col.append((campers_row_index + 1, target_columns[0]))
    if staff_row_index > -1:
        stakeholders_col.append((staff_row_index + 1, target_columns[1]))
    return stakeholders_col


def write_value_in_excel_cell(ws, value, target_row, target_col):
    cell = "{}{}".format(target_col, target_row)
    ws.Range(cell).Value = value


def populate_excel_by_col_labels(table, target_columns, cols_labels_and_target_rows_for_excel, ws):
    stakeholders_col = get_stakeholders_from_rows_n_target_cols(table, target_columns)
    # stakeholders_col is a list of one or two tuples.
    # Each tuple has 2 items - 1) the row in the spss table to get information from (row_index)
    # and 2) the target column in excel (target_col)
    for s in stakeholders_col:  # iterate on stakeholders, s[0] = row_index, s[1] = target_column
        # i = 0  # i is the index of the table line (row)  - I think need to switch to column?
        # while i < len(table[0].data):
        get_values_from_table(s, table, cols_labels_and_target_rows_for_excel, ws)


def populate_excel_ex_1(ws, index_information, tables, target_columns):
    for d in index_information:
        t = tables[d['table_index']]
        cols_labels_and_target_rows_for_excel = d['cols_labels_and_target_rows_for_excel']
        populate_excel_by_col_labels(t, target_columns, cols_labels_and_target_rows_for_excel, ws)


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
    target_columns = ['C', 'D']
    index_information = [
        dict(table_name='gender', table_index=3,
             cols_labels_and_target_rows_for_excel=[dict(label='Male', row=5),
                                                    dict(label='Female', row=6)]),
        dict(table_name='denomination', table_index=4,
             cols_labels_and_target_rows_for_excel=[dict(label='Conservative', row=7),
                                                    dict(label='Just Jewish or Secular', row=8),
                                                    dict(label='Orthodox or Modern Orthodox', row=9),
                                                    dict(label='Reform', row=10)]),
        dict(table_name='Attended day school', table_index=5,
             cols_labels_and_target_rows_for_excel=[dict(label='Yes', row=11)]),
        dict(table_name='Attended overnight camp', table_index=6,
             cols_labels_and_target_rows_for_excel=[dict(label='Yes', row=12)]),
        dict(table_name='Attended supplementary school', table_index=7,
             cols_labels_and_target_rows_for_excel=[dict(label='Yes', row=13)]),
        dict(table_name='Attended youth group', table_index=8,
             cols_labels_and_target_rows_for_excel=[dict(label='Yes', row=14)]),
        dict(table_name='Frequency of attending religious services', table_index=9,
             cols_labels_and_target_rows_for_excel=[dict(label='Never', row=15),
                                                    dict(label='Only on the High Holidays', row=16),
                                                    dict(label='A few times a year on holidays', row=17),
                                                    dict(label='About once a month', row=18),
                                                    dict(label='About once a week', row=19),
                                                    dict(label='Almost on a daily basis', row=20)]),
        dict(table_name='Previous visits to Israel', table_index=10,
             cols_labels_and_target_rows_for_excel=[dict(label='Never', row=21),
                                                    dict(label='Once', row=22),
                                                    dict(label='More than once', row=23)])
    ]
    populate_excel_ex_1(ws, index_information, tables, target_columns)


def populate_excel_ex_3_5_7(ws, index_information, t, target_columns):
    for d in index_information:
        cols_labels_and_target_rows_for_excel = d['cols_labels_and_target_rows_for_excel']
        populate_excel_by_col_labels(t, target_columns, cols_labels_and_target_rows_for_excel, ws)


def populate_exhibit3(tables, workbook, camp):
    ws = workbook.Worksheets('All.Exhibits')
    if camp['code'] in (9, 11, 19, 20, 21):
        t = tables[2]
    else:
        t = tables[1]
    target_columns = ['B', 'C']
    index_information = [
        dict(table_name='How often do you hear Hebrew spoken around camp?',
             cols_labels_and_target_rows_for_excel=[dict(label='Not at all', row=46),
                                                    dict(label='Rarely', row=45),
                                                    dict(label='Sometimes', row=44),
                                                    dict(label='Often', row=43),
                                                    dict(label='Very often', row=42)])
    ]
    populate_excel_ex_3_5_7(ws, index_information, t, target_columns)


def populate_exhibit5(tables, workbook, camp):
    ws = workbook.Worksheets('All.Exhibits')
    if camp['code'] in (9, 11, 19, 20, 21):
        t = tables[2]
    else:
        t = tables[1]
    target_columns = ['B', 'C']
    index_information = [
        dict(table_name='How often do you have Israel programming at your camp?',
             cols_labels_and_target_rows_for_excel=[dict(label='Not at all', row=65),
                                                    dict(label='Just once during the summer', row=64),
                                                    dict(label='Every few weeks during the summer',
                                                         row=63),
                                                    dict(label='Once a week', row=62),
                                                    dict(label='A few times a week', row=61),
                                                    dict(label='Every day', row=60)])
    ]
    populate_excel_ex_3_5_7(ws, index_information, t, target_columns)


def populate_exhibit7(tables, workbook, camp):
    ws = workbook.Worksheets('All.Exhibits')
    if camp['code'] in (9, 11, 19, 20, 21):
        t = tables[2]
    else:
        t = tables[1]
    target_columns = ['B', 'C']
    index_information = [
        dict(table_name='How would you rate the Israel programming at your camp?',
             cols_labels_and_target_rows_for_excel=[dict(label='Very boring', row=85),
                                                    dict(label='Boring', row=84),
                                                    dict(label='Just okay',
                                                         row=83),
                                                    dict(label='Enjoyable', row=82),
                                                    dict(label='Very enjoyable', row=81)])
    ]
    populate_excel_ex_3_5_7(ws, index_information, t, target_columns)


def get_stakeholders_from_rows_n_target_rows(t, target_rows):
    campers_vs_staff = t.get_col_by_index(1)
    campers_row_index = get_index(campers_vs_staff, 'Campers')
    staff_row_index = get_index(campers_vs_staff, 'Staff')
    stakeholders_col = []
    if campers_row_index > -1:
        stakeholders_col.append((campers_row_index + 1, target_rows[0]))
    if staff_row_index > -1:
        stakeholders_col.append((staff_row_index + 1, target_rows[1]))
    return stakeholders_col


def get_values_from_table_2_8(s, table, cols_labels_and_target_cols_for_excel, ws):
    labels_from_spss = table.get_row(0)[3:-1]
    row = table.get_row(s[0])  # s[0] is index of the row where the data are
    for d in cols_labels_and_target_cols_for_excel:
        match = False
        for i, spss_label in enumerate(labels_from_spss):
            if spss_label == d['label']:
                match = True
                break
        if not match:
            continue
        target_col = d['col']
        value = row[i + 3]
        write_value_in_excel_cell(ws, value, s[1], target_col)

def populate_excel_by_col_labels_2_8(t, target_rows, cols_labels_and_target_cols_for_excel, ws):
    stakeholders_col = get_stakeholders_from_rows_n_target_rows(t, target_rows)
    # stakeholders_col is a list of one or two tuples.
    # Each tuple has 2 items - 1) the row in the spss table to get information from (row_index)
    # and 2) the target column in excel (target_col)
    for s in stakeholders_col:  # iterate on stakeholders, s[0] = row_index, s[1] = target_column
        # i = 0  # i is the index of the table line (row)  - I think need to switch to column?
        # while i < len(table[0].data):
        get_values_from_table_2_8(s, t, cols_labels_and_target_cols_for_excel, ws)

def populate_excel_ex_2_8(ws, index_information, t, tables, target_rows):
    stakeholders_col = get_stakeholders_from_rows_n_target_rows(t, target_rows)
    for d in index_information:
        t = tables[d['table_index']]
        for row_index, target_row in stakeholders_col:
            for col_index, target_col in d['from_col_to_target_col']:
                row = t.get_row(row_index)
                cell = "{}{}".format(target_col, target_row)
                ws.Range(cell).Value = row[col_index]


def populate_exhibit2(tables, workbook, camp):
    ws = workbook.Worksheets('All.Exhibits')
    if camp['code'] in (9, 11, 19, 20, 21):
        t = tables[2]
    else:
        t = tables[1]
    target_rows = [30, 29]
    index_information = [
        dict(table_name='How many Israelis are there at your camp?',
             cols_labels_and_target_rows_for_excel=[dict(label='Very boring', col='I'),
                                                    dict(label='Enjoyable', col='H'),
                                                    dict(label='Very enjoyable', col='G')])
    ]
    populate_excel_ex_2_8(ws, index_information, t, tables, target_rows)


def populate_exhibit8(tables, workbook, camp):
    ws = workbook.Worksheets('All.Exhibits')
    if camp['code'] in (9, 11, 19, 20, 21):
        t = tables[2]
    else:
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
            output[s - 3].append(output_row)
    return output


def sort_by_column_descending(data):
    data = [d for d in data if d[1] > 50.5]
    res = tuple(reversed(sorted(data, key=lambda r: r[1])))
    return res


def populate_exhibit4(tables, workbook, camp):
    ws = workbook.Worksheets('All.Exhibits')
    if camp['code'] in (9, 11, 19, 20, 21):
        t = tables[2]
    else:
        t = tables[1]
    columns, target_cols = identify_colmuns_ex_4_6(t)
    rows = t.data[2:-1]
    output_rows = iterate_over_two_rows_at_a_time(rows, columns)
    for col in columns:
        sorted_row = sort_by_column_descending(output_rows[col - 3])
        target_row = 101
        target_col = target_cols[col - 3]
        for r in sorted_row:
            label = r[0]
            percent = int(round(r[1]))
            cell_value = '{} ({}%)'.format(label, percent)
            cell = "{}{}".format(target_col, target_row)
            ws.Range(cell).Value = cell_value
            target_row += 1


def populate_exhibit6(tables, workbook, camp):
    ws = workbook.Worksheets('All.Exhibits')
    if camp['code'] in (9, 11, 19, 20, 21):
        t = tables[2]
    else:
        t = tables[1]
    columns, target_cols = identify_colmuns_ex_4_6(t)
    rows = t.data[2:-1]
    output_rows = iterate_over_two_rows_at_a_time(rows, columns)
    for col in columns:
        sorted_row = sort_by_column_descending(output_rows[col - 3])
        target_row = 114
        target_col = target_cols[col - 3]
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
