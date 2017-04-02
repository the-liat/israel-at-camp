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


#
# def enter_students_values(ws, t_students_own, t_students_com, school):
#     num_schools = school['testedGradeCount']
#     spss_indexes = [2, 3, 4][:num_schools]
#     tables = [
#         dict(table=t_students_own, column='C'),
#         dict(table=t_students_com, column='D')
#     ]
#     all_rows = {'5th': 4, '8th': 5, '11th': 6}
#     rows = []
#     d = school['grades']
#     sorted_grades = sorted(d.keys(), key=lambda x: int(x[:-2]))
#     for grade in sorted_grades:
#         if d[grade] == 1:
#             rows.append(all_rows[grade])
#     for d in tables:
#         table = d['table']
#         col = d['column']
#         for row, spss_value_index in zip(rows, spss_indexes):
#             value = table.get_row(2)[spss_value_index]
#             cell = "{}{}".format(col, row)
#             ws.Range(cell).Value = value


# def populate_exhibit2(table_dict, workbook, school):
#     ws = workbook.Worksheets('Exhibits 2,3,4,5')
#     t_students_own = table_dict[('own_school', 'campers')][1]
#     t_students_com = table_dict[('comparison_schools', 'campers')][1]
#     enter_students_values(ws, t_students_own, t_students_com, school)
#     t_parents_own = table_dict[('own_school', 'parents')][1]
#     row_index_po = 1 if len(t_parents_own.data) > 1 else 0
#     ws.Range("C3").Value = t_parents_own.get_row(row_index_po)[2]
#     t_parents_com = table_dict[('comparison_schools', 'parents')][1]
#     row_index_pc = 1 if len(t_parents_com.data) > 1 else 0
#     ws.Range("D3").Value = t_parents_com.get_row(row_index_pc)[2]
#     t_staff_own = table_dict[('own_school', 'staff')][1]
#     row_index_so = 1 if len(t_staff_own.data) > 1 else 0
#     ws.Range("C7").Value = t_staff_own.get_row(row_index_so)[2]
#     t_staff_com = table_dict[('comparison_schools', 'staff')][1]
#     row_index_sc = 1 if len(t_staff_com.data) > 1 else 0
#     ws.Range("D7").Value = t_staff_com.get_row(row_index_sc)[2]


def get_index(col, item):
    return col.index(item) if item in col else -1


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
    campers_vs_staff = t.get_col_by_index(1)
    campers_row_index = get_index(campers_vs_staff, 'Campers')
    staff_row_index = get_index(campers_vs_staff, 'Staff')
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
    stakeholders_col = []
    if campers_row_index > -1:
        stakeholders_col.append((campers_row_index + 1, 'C'))
    if staff_row_index > -1:
        stakeholders_col.append((staff_row_index + 1, 'D'))
    for d in index_information:
        t = tables[d['table_index']]
        for row_index, target_col in stakeholders_col:
            for col_index, target_row in d['from_col_to_target_row']:
                row = t.get_row(row_index)
                cell = "{}{}".format(target_col, target_row)
                ws.Range(cell).Value = row[col_index]


def find_cells(table):  # this should work for every table in the table dictionary
    i = 0
    x = 0
    while table.get_row(i)[1] != 'Total':
        if table.get_row(i)[1] == 'Important' or table.get_row(i)[1] == 'Very important':
            x += table.get_row(i)[4]
        i += 1
    return x  # sum of 'important' and 'very important' %


def get_values_from_table(i, table, labels_rows, y):
    # Y is the index of the value needed to be pulled from the spss table
    row = table.get_row(i)
    for d in labels_rows:
        match = True
        for col, label in d['labels'].iteritems():
            if row[col] != label:
                match = False
                break
        if not match:
            continue
        return row[y], d['row']
    return None, None


def populate_excel_by_row_labels(tables, labels_rows, ws, y):  # for exhibits 7, 9, 10, 13, 17, 18, 20
    for d in tables:
        i = 0  # i is the index of the table line
        while i < len(d['table'].data):
            value, row = get_values_from_table(i, d['table'], labels_rows, y)
            if value is None:
                i += 1
                continue
            col = d['column']
            cell = "{}{}".format(col, row)
            ws.Range(cell).Value = value
            i += 1


def get_values_exhibits_6_17(i, j, table, labels_rows, y):
    value = table.get_row(i)[y]
    row = labels_rows[j]['row']
    return value, row


def populate_exhibit6(table_dict, workbook, school):
    ws = workbook.Worksheets('Exhibits 6,7,8,9')
    # where each value in table goes in excel rows
    labels_rows = [
        dict(labels={0: 'it connects Jews around the world',
                     1: 'Agree/Strongly Agree'}, row=5),
        dict(labels={0: 'it makes one feel a part of the group when people mix Hebrew into English',
                     1: 'Agree/Strongly Agree'}, row=6),
        dict(labels={0: 'it is a part of being Jewish',
                     1: 'Agree/Strongly Agree'}, row=7),
        dict(labels={0: "it maintains the Jewish people's language",
                     1: 'Agree/Strongly Agree'}, row=8),
        dict(labels={0: 'it helps in forming a connection with Israel',
                     1: 'Agree/Strongly Agree'}, row=9),
        dict(labels={0: 'it prepares one to make Aliyah in case one wants to',
                     1: 'Agree/Strongly Agree'}, row=12),
        dict(labels={0: 'it helps when visiting Israel',
                     1: 'Agree/Strongly Agree'}, row=13),
        dict(labels={0: 'it allows one to read modern Israeli books, newspapers, websites or music lyrics',
                     1: 'Agree/Strongly Agree'}, row=14),
        dict(labels={0: 'it helps in communicating with other Jews around the world',
                     1: 'Agree/Strongly Agree'}, row=15),
        dict(labels={0: 'it helps communicate with people who only speak Hebrew',
                     1: 'Agree/Strongly Agree'}, row=16),
        dict(labels={0: 'Learning a second language contributes to brain development',
                     1: 'Agree/Strongly Agree'}, row=18)
    ]
    # defining tables to get information from
    t_parents_own = table_dict[('own_school', 'parents')][0]
    t_students_own = table_dict[('own_school', 'campers')][0]
    t_staff_own = table_dict[('own_school', 'staff')][0]
    t_parents_com = table_dict[('comparison_schools', 'parents')][0]
    t_students_com = table_dict[('comparison_schools', 'campers')][0]
    t_staff_com = table_dict[('comparison_schools', 'staff')][0]
    # where each table data goes in excel columns
    tables = [
        dict(table=t_staff_own, column='D'),
        dict(table=t_students_own, column='C'),
        dict(table=t_parents_own, column='B'),
        dict(table=t_staff_com, column='G'),
        dict(table=t_students_com, column='F'),
        dict(table=t_parents_com, column='E')
    ]
    # in each table, it will go to the line and find the appropriate
    # value for this line and put it in the right position, iterate on all lines
    y = 2  # Y is the index of the value needed to be pulled from the spss table
    for d in tables:
        i = 1  # i is the index of the table line
        j = 0
        while i < len(d['table'].data):
            value, row = get_values_exhibits_6_17(i, j, d['table'], labels_rows, y)
            col = d['column']
            cell = "{}{}".format(col, row)
            ws.Range(cell).Value = value
            i += 2
            j += 1


def populate_exhibit7(table_dict, workbook, school):  # table_dict is the dictionary for the spss tables
    ws = workbook.Worksheets('Exhibits 6,7,8,9')
    # where each value in table goes in excel rows
    labels_rows = [
        dict(labels={1: 'Not at all satisfied'}, row=26),
        dict(labels={1: 'A little bit satisfied'}, row=27),
        dict(labels={1: 'Somewhat satisfied'}, row=28),
        dict(labels={1: 'Satisfied'}, row=29),
        dict(labels={1: 'Very satisfied'}, row=30)
    ]
    # defining tables to get information from
    t_parents_own = table_dict[('own_school', 'parents')][1]
    t_staff_own = table_dict[('own_school', 'staff')][1]
    t_parents_com = table_dict[('comparison_schools', 'parents')][1]
    t_staff_com = table_dict[('comparison_schools', 'staff')][1]
    # where each table data goes in excel columns
    tables = [
        dict(table=t_staff_own, column='K'),
        dict(table=t_parents_own, column='L'),
        dict(table=t_staff_com, column='M'),
        dict(table=t_parents_com, column='N')
    ]
    # in each table, it will go to the line and find the appropriate
    # value for this line and put it in the right position, iterate on all lines
    y = 4  # Y is the index of the value needed to be pulled from the spss table
    populate_excel_by_row_labels(tables, labels_rows, ws, y)


def populate_excel_by_row_labels_ex8(tables, labels_rows, ws, y):
    for d in tables:
        i = 0  # i is the index of the table line
        while i < len(d['table'].data):
            value, row = get_values_from_table(i, d['table'], labels_rows, y)
            if value is None:
                i += 1
                continue
            col = d['column']
            cell = "{}{}".format(col, row)
            if value >= 33:
                ws.Range(cell).Value = 'X'
            i += 1


def populate_exhibit8(table_dict, workbook, school):  # table_dict is the dictionary for the spss tables
    ws = workbook.Worksheets('Exhibits 6,7,8,9')
    # where each value in table goes in excel rows
    labels_rows = [
        dict(labels={1: 'Challange-The teachers do not have expertise in second language instruction'},
             row=43),
        dict(labels={1: 'Challange-The teachers are not knowledgeable in Hebrew for everyday communication'},
             row=44),
        dict(labels={1: 'Challange-The Hebrew curriculum used is not good enough (e.g, outdated, not challenging)'},
             row=45),
        dict(labels={1: 'Challange-Hebrew for everyday communication instruction is mainly conducted in English'},
             row=46),
        dict(labels={1: 'Challange-The teachers do not care about Hebrew for everyday communication proficiency'},
             row=47),
        dict(labels={1: 'Challange-It is not a priority of the school'},
             row=48),
        dict(labels={1: 'Challange-There is not enough time devoted to Hebrew for everyday communication'},
             row=49),
        dict(labels={1: 'Challange-The diversity of Hebrew levels in the class'},
             row=50),
        dict(labels={1: 'Challange-There are too many children in the classroom'},
             row=51),
        dict(labels={
            1: 'Challange-Hebrew for everyday communication is the first class that gets canceled for an activity'},
            row=52),
    ]
    # defining tables to get information from
    t_parents_own = table_dict[('own_school', 'parents')][1]
    t_staff_own = table_dict[('own_school', 'staff')][1]
    t_parents_com = table_dict[('comparison_schools', 'parents')][1]
    t_staff_com = table_dict[('comparison_schools', 'staff')][1]
    # where each table data goes in excel columns
    tables = [
        dict(table=t_staff_own, column='C'),
        dict(table=t_parents_own, column='B'),
        dict(table=t_staff_com, column='E'),
        dict(table=t_parents_com, column='D')
    ]
    # in each table, it will go to the line and find the appropriate
    # value for this line and put it in the right position, iterate on all lines
    y = 4  # Y is the index of the value needed to be pulled from the spss table
    populate_excel_by_row_labels_ex8(tables, labels_rows, ws, y)


def populate_exhibit9(table_dict, workbook, school):  # table_dict is the dictionary for the spss tables
    ws = workbook.Worksheets('Exhibits 6,7,8,9')
    # where each value in table goes in excel rows
    labels_rows = [
        dict(labels={1: 'Ongoing professional development opportunities'}, row=57),
        dict(labels={1: 'Enough time to prepare'}, row=58),
        dict(labels={1: 'Administrative support'}, row=59),
        dict(labels={1: 'Hebrew for everyday communication assessment instrument'}, row=60),
        dict(labels={1: 'Classroom support(e.g., teaching assistant)'}, row=61),
        dict(labels={1: 'Hebrew for everyday communication curriculum'}, row=62),
        dict(labels={1: 'Resources for special needs / gifted campers'}, row=63),
        dict(labels={1: 'Text and Prayer Hebrew assessment instrument'}, row=64),
        dict(labels={1: 'Text and Prayer Hebrew curriculum'}, row=65),
        dict(labels={1: 'Pedagogical materials(books, resources)'}, row=66),
    ]
    # defining tables to get information from
    t_staff_own = table_dict[('own_school', 'staff')][0]
    t_staff_com = table_dict[('comparison_schools', 'staff')][0]
    # where each table data goes in excel columns
    tables = [
        dict(table=t_staff_own, column='B'),
        dict(table=t_staff_com, column='C')
    ]
    # in each table, it will go to the line and find the appropriate
    # value for this line and put it in the right position, iterate on all lines
    y = 2  # Y is the index of the value needed to be pulled from the spss table
    for d in tables:
        i = 3  # i is the index of the table line
        j = 0
        while i < len(d['table'].data):
            value, row = get_values_exhibits_6_17(i, j, d['table'], labels_rows, y)
            col = d['column']
            cell = "{}{}".format(col, row)
            ws.Range(cell).Value = value
            i += 4
            j += 1


def populate_exhibit10(table_dict, workbook, school):  # table_dict is the dictionary for the spss tables
    ws = workbook.Worksheets('Exhibit 10 +Comments')
    # where each value in table goes in excel rows
    labels_rows = [
        dict(labels={1: 'Much worse'}, row=4),
        dict(labels={1: 'Worse'}, row=5),
        dict(labels={1: 'About the same'}, row=6),
        dict(labels={1: 'Better'}, row=7),
        dict(labels={1: 'Much better'}, row=8)
    ]
    # defining tables to get information from
    if school['code'] in (3, 35, 5, 42, 22, 27, 36, 18, 19):
        t_parents_own = table_dict[('own_school', 'parents')][2]
    else:
        t_parents_own = table_dict[('own_school', 'parents')][1]
    t_parents_com = table_dict[('comparison_schools', 'parents')][2]
    # where each table data goes in excel columns
    tables = [
        dict(table=t_parents_own, column='M'),
        dict(table=t_parents_com, column='N')
    ]
    # in each table, it will go to the line and find the appropriate
    # value for this line and put it in the right position, iterate on all lines
    y = 4  # Y is the index of the value needed to be pulled from the spss table
    populate_excel_by_row_labels(tables, labels_rows, ws, y)
    if school['code'] in (3, 35, 5, 42, 22, 27, 36, 18, 19):
        t = table_dict[('own_school', 'parents')][3]
    else:
        t = table_dict[('own_school', 'parents')][2]
    i = 0
    row = 17  # starting row to paste comments
    while i < len(t.data):
        value = t.get_row(i)[1]
        cell = "A{}".format(row)
        ws.Range(cell).Value = value
        i += 1
        row += 1


def populate_exhibit13(table_dict, workbook, school):  # table_dict is the dictionary for the spss tables
    ws = workbook.Worksheets('Exhibits 12,13')
    # where each value in table goes in excel rows
    labels_rows = [
        dict(labels={1: 'Hate it'}, row=16),
        dict(labels={1: 'Dislike it'}, row=17),
        dict(labels={1: 'Neutral'}, row=18),
        dict(labels={1: 'Like it'}, row=19),
        dict(labels={1: 'Love it'}, row=20)
    ]
    # defining tables to get information from
    t_students_own = table_dict[('own_school', 'campers')][1]
    t_students_com = table_dict[('comparison_schools', 'campers')][1]
    # where each table data goes in excel columns
    tables = [
        dict(table=t_students_own, column='M'),
        dict(table=t_students_com, column='N')
    ]
    # in each table, it will go to the line and find the appropriate
    # value for this line and put it in the right position, iterate on all lines
    y = 4  # Y is the index of the value needed to be pulled from the spss table
    populate_excel_by_row_labels(tables, labels_rows, ws, y)


def populate_exhibit16(table_dict, workbook, school):  # same as exhibit 5
    ws = workbook.Worksheets('Exhibits 16,17,18,19,20')
    for school, stakeholder in table_dict:
        table = table_dict[(school, stakeholder)][1]
        value = find_cells(table)
        if school == 'own_school':
            col = 'M'
        else:
            col = 'N'
        if stakeholder == 'staff':
            row = '5'
        elif stakeholder == 'campers':
            row = '6'
        else:
            row = '7'
        cell = "{}{}".format(col, row)
        ws.Range(cell).Value = value


def populate_exhibit17(table_dict, workbook, school):  # similar to exhibit 6
    ws = workbook.Worksheets('Exhibits 16,17,18,19,20')
    # where each value in table goes in excel rows
    labels_rows = [
        dict(labels={0: 'it helps recognize Hebrew prayer as a part of the Jewish heritage/tradition',
                     1: 'Agree/Strongly Agree'}, row=23),
        dict(labels={0: 'it prepares one to lead prayers',
                     1: 'Agree/Strongly Agree'}, row=24),
        dict(labels={0: 'it makes one feel comfortable when at a service in Hebrew',
                     1: 'Agree/Strongly Agree'}, row=25),
        dict(labels={0: 'it strengthens the appreciation of Jewish culture and tradition',
                     1: 'Agree/Strongly Agree'}, row=26),
        dict(labels={0: 'it makes one feel a part of the synagogue',
                     1: 'Agree/Strongly Agree'}, row=27),
        dict(labels={0: 'it deepens the experience of studying Jewish text',
                     1: 'Agree/Strongly Agree'}, row=30),
        dict(labels={0: 'it helps in understanding Jewish texts in their original Hebrew',
                     1: 'Agree/Strongly Agree'}, row=31),
        dict(labels={0: 'it helps in reading out loud Jewish texts in their original Hebrew',
                     1: 'Agree/Strongly Agree'}, row=32),
        dict(labels={0: 'it helps understand the meaning of prayers',
                     1: 'Agree/Strongly Agree'}, row=33),
        dict(labels={0: 'it prepares for studying Jewish text independently',
                     1: 'Agree/Strongly Agree'}, row=34)
    ]
    # defining tables to get information from
    t_parents_own = table_dict[('own_school', 'parents')][0]
    t_students_own = table_dict[('own_school', 'campers')][0]
    t_staff_own = table_dict[('own_school', 'staff')][0]
    t_parents_com = table_dict[('comparison_schools', 'parents')][0]
    t_students_com = table_dict[('comparison_schools', 'campers')][0]
    t_staff_com = table_dict[('comparison_schools', 'staff')][0]
    # where each table data goes in excel columns
    tables = [
        dict(table=t_staff_own, column='D'),
        dict(table=t_students_own, column='C'),
        dict(table=t_parents_own, column='B'),
        dict(table=t_staff_com, column='G'),
        dict(table=t_students_com, column='F'),
        dict(table=t_parents_com, column='E')
    ]
    # in each table, it will go to the line and find the appropriate
    # value for this line and put it in the right position, iterate on all lines
    y = 2  # Y is the index of the value needed to be pulled from the spss table
    for d in tables:
        i = 1  # i is the index of the table line
        j = 0
        while i < len(d['table'].data):
            value, row = get_values_exhibits_6_17(i, j, d['table'], labels_rows, y)
            col = d['column']
            cell = "{}{}".format(col, row)
            ws.Range(cell).Value = value
            i += 2
            j += 1


def populate_exhibit18(table_dict, workbook, school):  # tsimilar to exhibit 7
    ws = workbook.Worksheets('Exhibits 16,17,18,19,20')
    # where each value in table goes in excel rows
    labels_rows = [
        dict(labels={1: 'Not at all satisfied'}, row=41),
        dict(labels={1: 'A little bit satisfied'}, row=42),
        dict(labels={1: 'Somewhat satisfied'}, row=43),
        dict(labels={1: 'Satisfied'}, row=44),
        dict(labels={1: 'Very satisfied'}, row=45)
    ]
    # defining tables to get information from
    t_parents_own = table_dict[('own_school', 'parents')][1]
    t_staff_own = table_dict[('own_school', 'staff')][1]
    t_parents_com = table_dict[('comparison_schools', 'parents')][1]
    t_staff_com = table_dict[('comparison_schools', 'staff')][1]
    # where each table data goes in excel columns
    tables = [
        dict(table=t_staff_own, column='M'),
        dict(table=t_parents_own, column='N'),
        dict(table=t_staff_com, column='O'),
        dict(table=t_parents_com, column='P')
    ]
    # in each table, it will go to the line and find the appropriate
    # value for this line and put it in the right position, iterate on all lines
    y = 4  # Y is the index of the value needed to be pulled from the spss table
    populate_excel_by_row_labels(tables, labels_rows, ws, y)


def populate_exhibit19(table_dict, workbook, school):  # similar to exhibit 8
    ws = workbook.Worksheets('Exhibits 16,17,18,19,20')
    # where each value in table goes in excel rows
    labels_rows = [
        dict(labels={1: 'Challange-The teachers do not have enough teaching experience'},
             row=59),
        dict(labels={1: 'Challange-The teachers are not knowledgeable in Hebrew for prayer or text study'},
             row=60),
        dict(labels={1: 'Challange-Hebrew for prayer and text study instruction is mainly conducted in English'},
             row=61),
        dict(labels={1: 'Challange-The teachers do not prioritize the mastery of classical text in Hebrew'},
             row=62),
        dict(labels={1: 'Challange-There are not enough Hebrew for prayer or text study teachers'},
             row=63),
        dict(
            labels={1: 'Challange-There is not enough time devoted to study Hebrew for prayer or text study in Hebrew'},
            row=64),
        dict(labels={1: 'Challange-Classical Hebrew texts are taught in translation'},
             row=65),
        dict(labels={1: 'Challange-There are too many children in the classroom'},
             row=66),
        dict(labels={1: "Challange-The instruction is mainly conducted in Hebrew (Ivrit b'Ivrit)"},
             row=67),
        dict(labels={1: 'Challange-The diversity of Hebrew levels in the class'},
             row=68),
    ]
    # defining tables to get information from
    t_parents_own = table_dict[('own_school', 'parents')][1]
    t_staff_own = table_dict[('own_school', 'staff')][1]
    t_parents_com = table_dict[('comparison_schools', 'parents')][1]
    t_staff_com = table_dict[('comparison_schools', 'staff')][1]
    # where each table data goes in excel columns
    tables = [
        dict(table=t_staff_own, column='C'),
        dict(table=t_parents_own, column='B'),
        dict(table=t_staff_com, column='E'),
        dict(table=t_parents_com, column='D')
    ]
    # in each table, it will go to the line and find the appropriate
    # value for this line and put it in the right position, iterate on all lines
    y = 4  # Y is the index of the value needed to be pulled from the spss table
    populate_excel_by_row_labels_ex8(tables, labels_rows, ws, y)


def populate_exhibit20(table_dict, workbook, school):  # similar to exhibit 13
    ws = workbook.Worksheets('Exhibits 16,17,18,19,20')
    # where each value in table goes in excel rows
    labels_rows = [
        dict(labels={1: 'Hate it'}, row=73),
        dict(labels={1: 'Dislike it'}, row=74),
        dict(labels={1: 'Neutral'}, row=75),
        dict(labels={1: 'Like it'}, row=76),
        dict(labels={1: 'Love it'}, row=77)
    ]
    # defining tables to get information from
    t_students_own = table_dict[('own_school', 'campers')][1]
    t_students_com = table_dict[('comparison_schools', 'campers')][1]
    # where each table data goes in excel columns
    tables = [
        dict(table=t_students_own, column='M'),
        dict(table=t_students_com, column='N')
    ]
    # in each table, it will go to the line and find the appropriate
    # value for this line and put it in the right position, iterate on all lines
    y = 4  # Y is the index of the value needed to be pulled from the spss table
    populate_excel_by_row_labels(tables, labels_rows, ws, y)


def get_value_ex11(table, spss_index):
    i = 0
    value = 0
    while i < len(table.data):
        if table.get_row(i)[1] == 'Agree' or table.get_row(i)[1] == 'Strongly agree':
            v = table.get_row(i)[spss_index]
            v = 0 if isinstance(v, str) else v
            value += v  # sum of 'Agree' and 'SA' %
        i += 1
    return value


def populate_exhibit11(table_dict, workbook, school):
    ws = workbook.Worksheets('Exhibit 11, three options')
    # where each value in table goes in excel rows
    # number of grades: own_school row , comparison_schools row
    all_rows = {3: (5, 6), 2: (22, 23), 1: (39, 40)}
    # where each table data goes in excel columns
    # a dictionary that maps table number (from spss) to relevant columns (in xl)
    columns = {
        1: ('Q', 'S', 'U'),  # The teaching of Hebrew is fun and interesting
        2: ('R', 'T', 'V')  # I like the learning materials in my Hebrew language classes
    }
    grade_names = school['grades']  # dict with grade names as keys and 0/1 as values
    num_grades = school['testedGradeCount']  # number of grades in this school
    if num_grades == 2:
        if grade_names['5th'] == 1:
            c1, c2 = 'Grade 5', 'Grade 8'
        else:
            c1, c2 = 'Grade 8', 'Grade 11'
        ws.Range("Q20").Value = c1
        ws.Range("S20").Value = c2
    spss_column_index_list = [2, 3, 4]
    indexes = spss_column_index_list[:num_grades]
    rows = all_rows[num_grades]
    for table_num, col_list in columns.iteritems():
        tables = {0: table_dict[('own_school', 'campers')][table_num],
                  1: table_dict[('comparison_schools', 'campers')][table_num]}
        for xl_row_index, table in tables.iteritems():
            xl_col_index = 0
            for spss_index in indexes:
                value = get_value_ex11(table, spss_index)
                cell = "{}{}".format(col_list[xl_col_index], rows[xl_row_index])
                ws.Range(cell).Value = value
                xl_col_index += 1


def get_values_from_table_ex12(i, table, labels_and_cols, y):
    # Y is the index of the value needed to be pulled from the spss table
    row = table.get_row(i)
    for d in labels_and_cols:
        match = True
        for col, label in d['labels'].iteritems():
            if row[col] != label:
                match = False
                break
        if not match:
            continue
        return row[y], d['column']
    return None, None


def populate_excel_by_col_labels(tables, labels_and_cols, ws, y):
    for d in tables:
        i = 0  # i is the index of the table line
        while i < len(d['table'].data):
            value, col = get_values_from_table_ex12(i, d['table'], labels_and_cols, y)
            if value is None:
                i += 1
                continue
            row = d['row']
            cell = "{}{}".format(col, row)
            ws.Range(cell).Value = value
            i += 1


def populate_exhibit12(table_dict, workbook, school):
    ws = workbook.Worksheets('Exhibits 12,13')
    # where each value in table goes in excel columns
    labels_and_cols = [
        dict(labels={1: 'Much worse'}, column='M'),
        dict(labels={1: 'Worse'}, column='N'),
        dict(labels={1: 'About the same'}, column='O'),
        dict(labels={1: 'Better'}, column='P'),
        dict(labels={1: 'Much better'}, column='Q')
    ]
    # where each table data goes in excel rows
    tables_and_rows = [
        # Compared to other topics, how would you rate Hebrew instruction?
        dict(table=table_dict[('own_school', 'campers')][1], row=4),
        dict(table=table_dict[('comparison_schools', 'campers')][1], row=5),
        # Compared to other second-language classes, how are you doing in Hebrew?
        dict(table=table_dict[('own_school', 'campers')][2], row=6),
        dict(table=table_dict[('comparison_schools', 'campers')][2], row=7)
    ]
    # in each table, it will go to the line and find the appropriate
    # value for this line and put it in the right position, iterate on all lines
    y = 4  # Y is the index of the value needed to be pulled from the spss table
    populate_excel_by_col_labels(tables_and_rows, labels_and_cols, ws, y)


def change_labels_if_only_2_grades(num_grades, grade_names, ws):
    if num_grades == 2:
        if grade_names['5th'] == 1:
            c1, c2 = 'Students: Grade 5', 'Students: Grade 8'
        else:
            c1, c2 = 'Students: Grade 8', 'Students: Grade 11'
            ws.Range("A18").Value = c1
            ws.Range("A19").Value = c2


def define_dictionaries_ex14(table_dict, num_grades, grade_names):
    all_rows = {
        3: dict(parents=[6], staff=[7], students=[8, 9, 10]),
        2: dict(parents=[16], staff=[17], students=[18, 19]),
        1: dict(parents=[25], staff=[26], students=[27])
    }
    # where each table data goes in excel columns
    # a dictionary that maps table number (from spss) to relevant columns (in xl)
    # for campers need to add +2 to the key (indexes are
    columns = dict(
        parents={0: ('B', 'F'),  # Reading, own school and comparison
                 1: ('C', 'G'),  # Writing, own school and comparison
                 2: ('D', 'H'),  # speaking, own school and comparison
                 3: ('E', 'I')  # understanding, own school and comparison
                 },
        staff={0: ('B', 'F'),  # Reading, own school and comparison
               1: ('C', 'G'),  # Writing, own school and comparison
               2: ('D', 'H'),  # speaking, own school and comparison
               3: ('E', 'I')  # understanding, own school and comparison
               },
        students={2: ('B', 'F'),  # Reading, own school and comparison
                  3: ('C', 'G'),  # Writing, own school and comparison
                  4: ('D', 'H'),  # speaking, own school and comparison
                  5: ('E', 'I')  # understanding, own school and comparison
                  })
    # the key (0/1) is the specific spss table
    tables_d = dict(
        parents={0: table_dict[('own_school', 'parents')][0],
                 1: table_dict[('comparison_schools', 'parents')][0]},
        staff={0: table_dict[('own_school', 'staff')][0],
               1: table_dict[('comparison_schools', 'staff')][0]},
        students={0: table_dict[('own_school', 'campers')][0],
                  1: table_dict[('comparison_schools', 'campers')][0]}
    )
    indexes = get_indexes_by_grades(grade_names)
    spss_indexes = dict(parents=[1], staff=[1], students=indexes)
    rows = all_rows[num_grades]
    return rows, columns, tables_d, spss_indexes


def get_indexes_by_grades(grades):
    """
    :param grades: dict of grades (e.g. {'5th': 0, '8th': 1, '11th: 1})
    :return:
    """
    indexes = {'5th': 1, '8th': 2, '11th': 3}
    result = sorted(indexes[k] for k, v in grades.iteritems() if v == 1)
    return result


def enter_value_in_xl_cell(ws, table, spss_line, spss_value_index, row, col):
    value = table.get_row(spss_line)[spss_value_index]
    cell = "{}{}".format(col, row)
    ws.Range(cell).Value = value


def iterate_over_tables(ws, tables, index_col_s, row_s, spss_indexes_s):
    for col_index, table in tables.iteritems():
        for spss_line, col in index_col_s.iteritems():
            iterate_over_indexes(ws, table, col[col_index], spss_line, row_s, spss_indexes_s)


def iterate_over_indexes(ws, table, col, spss_line, row_s, spss_indexes_s):
    for row, spss_value_index in zip(row_s, spss_indexes_s):
        enter_value_in_xl_cell(ws, table, spss_line, spss_value_index, row, col)


def populate_exhibit14(table_dict, workbook, school):
    ws = workbook.Worksheets('Exhibit 14, three options')
    grade_names = school['grades']  # dict with grade names as keys and 0/1 as values
    num_grades = school['testedGradeCount']  # number of grades in this school
    change_labels_if_only_2_grades(num_grades, grade_names, ws)
    rows, columns, tables_d, spss_indexes = define_dictionaries_ex14(table_dict, num_grades, grade_names)
    stakeholders = ('parents', 'staff', 'campers')
    for s in stakeholders:
        index_col_s = columns[s]  # dictionary of spss indexes and corresponding xl columns
        row_s = rows[s]  # list of rows in xl
        spss_indexes_s = spss_indexes[s]  # list of indexes in spss
        tables = tables_d[s]  # dict of two tables - own and comparison
        iterate_over_tables(ws, tables, index_col_s, row_s, spss_indexes_s)


def define_dictionaries_ex21(table_dict, num_grades, grade_names):
    all_rows = {
        3: dict(parents=[6], staff=[7], students=[8, 9, 10]),
        2: dict(parents=[16], staff=[17], students=[18, 19]),
        1: dict(parents=[25], staff=[26], students=[27])
    }
    # where each table data goes in excel columns
    # a dictionary that maps table number (from spss) to relevant columns (in xl)
    # for campers need to add +2 to the key (indexes are
    columns = dict(
        parents={0: ('B', 'D'),  # Reading, own school and comparison
                 1: ('C', 'E'),  # understanding, own school and comparison
                 },
        staff={0: ('B', 'D'),  # Reading, own school and comparison
               1: ('C', 'E'),  # understanding, own school and comparison
               },
        students={2: ('B', 'D'),  # Reading, own school and comparison
                  3: ('C', 'E'),  # understanding, own school and comparison
                  })
    # the key (0/1) is the specific spss table
    tables_d = dict(
        parents={0: table_dict[('own_school', 'parents')][0],
                 1: table_dict[('comparison_schools', 'parents')][0]},
        staff={0: table_dict[('own_school', 'staff')][0],
               1: table_dict[('comparison_schools', 'staff')][0]},
        students={0: table_dict[('own_school', 'campers')][0],
                  1: table_dict[('comparison_schools', 'campers')][0]}
    )
    indexes = get_indexes_by_grades(grade_names)
    spss_indexes = dict(parents=[1], staff=[1], students=indexes)
    rows = all_rows[num_grades]
    return rows, columns, tables_d, spss_indexes


def populate_exhibit21(table_dict, workbook, school):
    ws = workbook.Worksheets('Exhibit 21, three options')
    grade_names = school['grades']  # dict with grade names as keys and 0/1 as values
    num_grades = school['testedGradeCount']  # number of grades in this school
    change_labels_if_only_2_grades(num_grades, grade_names, ws)
    rows, columns, tables_d, spss_indexes = define_dictionaries_ex21(table_dict, num_grades, grade_names)
    stakeholders = ('parents', 'staff', 'campers')
    for s in stakeholders:
        index_col_s = columns[s]  # dictionary of spss indexes and corresponding xl columns
        row_s = rows[s]  # list of rows in xl
        spss_indexes_s = spss_indexes[s]  # list of indexes in spss
        tables = tables_d[s]  # dict of two tables - own and comparison
        iterate_over_tables(ws, tables, index_col_s, row_s, spss_indexes_s)


def change_labels_if_only_2_grades_ex15(num_grades, grade_names, ws):
    if num_grades == 2:
        if grade_names['5th'] == 1:
            c1, c2 = 'Students: Grade 5', 'Students: Grade 8'
        else:
            c1, c2 = 'Students: Grade 8', 'Students: Grade 11'
            ws.Range("A29").Value = c1
            ws.Range("A35").Value = c2


def change_labels_if_only_2_grades_ex22(num_grades, grade_names, ws):
    if num_grades == 2:
        if grade_names['5th'] == 1:
            c1, c2 = 'Students: Grade 5', 'Students: Grade 8'
        else:
            c1, c2 = 'Students: Grade 8', 'Students: Grade 11'
            ws.Range("A25").Value = c1
            ws.Range("A30").Value = c2


def iterate_over_spss_indexes(ws, rows, tables_d, spss_indexes, row_jump):
    j = 0
    for indx in spss_indexes:
        iterate_over_tables_ex15(ws, rows, tables_d, j, indx)
        j += row_jump


def iterate_over_tables_ex15(ws, rows, tables_d, j, indx):
    # the following gets own or comp table with associated xl col and spss indexes (dictionary)
    for table_col_spss_indexes in tables_d:  # pick own or comparison table
        table = table_col_spss_indexes['table']
        xl_column_spss_indexes_d = table_col_spss_indexes['xl_column_spss_indexes']
        for col, line_indexes in xl_column_spss_indexes_d.iteritems():  # get list of col
            iterate_over_line_indexes(ws, table, rows, col, line_indexes, j, indx)


def create_spss_index_list(grade_names):
    # pick the index based on which grades
    spss_indexes = []
    if grade_names['5th'] == 1:
        spss_indexes.append(2)
    if grade_names['8th'] == 1:
        spss_indexes.append(3)
    if grade_names['11th'] == 1:
        spss_indexes.append(4)
    return spss_indexes


def iterate_over_line_indexes(ws, table, rows, col, line_indexes, j, indx):
    # focuse on one column in xl at a time, iterate over indexes to get values
    # match the values with the rows
    for r, spss_line in zip(rows, line_indexes):
        row = r['row'] + j  # row in excel
        enter_value_in_xl_cell(ws, table, spss_line, indx, row, col)


def define_dictionaries_ex15(table_dict, num_grades, grade_names):
    # where each value in table goes in excel rows
    all_labels_and_rows = {
        3: [
            dict(label='Chat with people in Hebrew', row=6),  # for the next grade jump all rows by 6
            dict(label='Speak Hebrew when called on to do so in class', row=7),
            dict(label='Understand Israeli songs', row=8),
            dict(label='Understand Israeli news or literature', row=9),
            dict(label='Understand social media posts in Hebrew', row=10),
            dict(label='Understand what my teacher(s) says in Hebrew', row=11)
        ],
        2: [
            dict(label='Chat with people in Hebrew', row=29),  # for the next grade jump all rows by 6
            dict(label='Speak Hebrew when called on to do so in class', row=30),
            dict(label='Understand Israeli songs', row=31),
            dict(label='Understand Israeli news or literature', row=32),
            dict(label='Understand social media posts in Hebrew', row=33),
            dict(label='Understand what my teacher(s) says in Hebrew', row=34)
        ],
        1: [
            dict(label='Chat with people in Hebrew', row=47),
            dict(label='Speak Hebrew when called on to do so in class', row=48),
            dict(label='Understand Israeli songs', row=49),
            dict(label='Understand Israeli news or literature', row=50),
            dict(label='Understand social media posts in Hebrew', row=51),
            dict(label='Understand what my teacher(s) says in Hebrew', row=52)
        ]
    }
    rows = all_labels_and_rows[num_grades]
    # defining tables to get information from and where each table data goes in excel columns
    tables_d = [
        dict(table=table_dict[('own_school', 'campers')][0],
             xl_column_spss_indexes=dict(
                 C=[2, 6, 10, 14, 18, 22],
                 D=[3, 7, 11, 15, 19, 23],
                 E=[4, 8, 12, 16, 20, 24],
                 F=[5, 9, 13, 17, 21, 25])),
        dict(table=table_dict[('comparison_schools', 'campers')][0],
             xl_column_spss_indexes=dict(
                 G=[2, 6, 10, 14, 18, 22],
                 H=[3, 7, 11, 15, 19, 23],
                 I=[4, 8, 12, 16, 20, 24],
                 J=[5, 9, 13, 17, 21, 25]))
    ]
    return rows, tables_d


def populate_exhibit15(table_dict, workbook, school):
    ws = workbook.Worksheets('Exhibit 15, three options')
    grade_names = school['grades']  # dict with grade names as keys and 0/1 as values
    spss_indexes = create_spss_index_list(
        grade_names)  # returns a list of indx- where to get info by grade level in each spsstable line
    num_grades = school['testedGradeCount']  # number of grades in this school
    change_labels_if_only_2_grades_ex15(num_grades, grade_names, ws)
    rows, tables_d = define_dictionaries_ex15(table_dict, num_grades, grade_names)
    row_jump = 6
    iterate_over_spss_indexes(ws, rows, tables_d, spss_indexes, row_jump)


def define_dictionaries_ex22(table_dict, num_grades, grade_names):
    # where each value in table goes in excel rows
    all_labels_and_rows = {
        3: [
            dict(label='Read unfamiliar siddur text', row=6),  # for the next grade jump all rows by 5
            dict(label='Understand unfamiliar siddur text', row=7),
            dict(label='Lead prayer', row=8),
            dict(label='Chant from the Torah', row=9),
            dict(label='Learn Jewish text independently', row=10)
        ],
        2: [
            dict(label='Read unfamiliar siddur text', row=25),  # for the next grade jump all rows by 5
            dict(label='Understand unfamiliar siddur text', row=26),
            dict(label='Lead prayer', row=27),
            dict(label='Chant from the Torah', row=28),
            dict(label='Learn Jewish text independently', row=29)
        ],
        1: [
            dict(label='Read unfamiliar siddur text', row=39),
            dict(label='Understand unfamiliar siddur text', row=40),
            dict(label='Lead prayer', row=41),
            dict(label='Chant from the Torah', row=42),
            dict(label='Learn Jewish text independently', row=43)
        ]
    }
    rows = all_labels_and_rows[num_grades]
    # defining tables to get information from and where each table data goes in excel columns
    tables_d = [
        dict(table=table_dict[('own_school', 'campers')][0],
             xl_column_spss_indexes=dict(
                 C=[2, 6, 10, 14, 18],
                 D=[3, 7, 11, 15, 19],
                 E=[4, 8, 12, 16, 20],
                 F=[5, 9, 13, 17, 21])),
        dict(table=table_dict[('comparison_schools', 'campers')][0],
             xl_column_spss_indexes=dict(
                 G=[2, 6, 10, 14, 18],
                 H=[3, 7, 11, 15, 19],
                 I=[4, 8, 12, 16, 20],
                 J=[5, 9, 13, 17, 21]))
    ]
    return rows, tables_d


def populate_exhibit22(table_dict, workbook, school):
    ws = workbook.Worksheets('Exhibit 22, three options')
    grade_names = school['grades']  # dict with grade names as keys and 0/1 as values
    spss_indexes = create_spss_index_list(
        grade_names)  # returns a list of indx- where to get info by grade level in each spsstable line
    num_grades = school['testedGradeCount']  # number of grades in this school
    change_labels_if_only_2_grades_ex22(num_grades, grade_names, ws)
    rows, tables_d = define_dictionaries_ex22(table_dict, num_grades, grade_names)
    row_jump = 5
    iterate_over_spss_indexes(ws, rows, tables_d, spss_indexes, row_jump)


"""
get_cell_value(row_index, col_index)
get_row(row_index)
get_col_by_index(col_index)
get_col_by_name(col_name)
"""
