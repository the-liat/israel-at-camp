import spssaux

from part_b_migrating_results_to_excel import open_excel, close_excel, create_camp_workbook_from_template, \
    populate_exhibit
from spss_analyses import run
from all_exhibits_syntax import exhibit_syntax
from camp_list import camps
from spss_output_parser import parse_output

spss_filename = r'C:\Users\Liat\Google Drive\UJA Toronto Israel at Camp\Databases and Data Files\Israel at Camp for report 3.9.sav'


def selection_of_camp(camp):
    """ Build the spss command for selection the specific camp
    populaate placehoolder in template with camp code
    :param camp:
    :return:
    """
    template = """USE ALL.
                  COMPUTE filter_$=(Camp_name= {0}).
                  VARIABLE LABELS filter_$ 'Camp_name= {0} (FILTER)'.
                  VALUE LABELS filter_$ 0 'Not Selected' 1 'Selected'.
                  FORMATS filter_$ (f1.0).
                  FILTER BY filter_$.
                  EXECUTE.
                  """
    selection = template.format(camp['code'])
    return selection


from all_exhibits_syntax import exhibit_syntax


def run_spss_syntax_per_exhibit(exhibit_number, spss_filename, selection):
    spssaux.OpenDataFile(spss_filename)
    print '--- exhibit: {}'.format(exhibit_number)
    commands = exhibit_syntax[exhibit_number]
    cmd_list = commands.split('\n')
    run(selection)
    run(cmd_list)
    out = run([commands])
    lines = out.split('\r\n')
    tables = parse_output(lines, tables_only=True)
    return tables


def run_analyses(camp):
    wb, file_name = create_camp_workbook_from_template(camp['name'])
    for exhibit_number, commands in exhibit_syntax.iteritems():
        selection = selection_of_camp(camp)
        tables = run_spss_syntax_per_exhibit(exhibit_number, spss_filename, selection)
        populate_exhibit(exhibit_number, tables, wb, camp)
    wb.SaveAs(file_name)


def main():
    open_excel()
    for camp in camps:
        run_analyses(camp)
    close_excel()


if __name__ == '__main__':
    main()
