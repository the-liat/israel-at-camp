import spss
import spssaux
import sys
from cStringIO import StringIO
from table import Table

def run(cmd_list):
    orig_stdout = sys.stdout
    try:
        sys.stdout = StringIO()
        spss.Submit(cmd_list)
        result = sys.stdout.getvalue()
        sys.stdout.close()
        return result
    finally:
        sys.stdout = orig_stdout


def parse_output(lines, tables_only=False):
    """Take list of lines and replace lines that represent tables with table objects

    """
    result = []
    table_lines = []
    for line in lines:
        if line.startswith('|'):
            table_lines.append(line)
        elif line == '' and table_lines:
            table = Table(table_lines)
            table_lines = []
            result.append(table)
        else:
            result.append(line)

    if tables_only:
        result = [t for t in result if isinstance(t, Table)]
    return result

def main():
    spssaux.OpenDataFile('f.sav')
    assert spss.IsBackendReady()

    cmds = [
        'FREQUENCIES VARIABLES=Finished /ORDER=ANALYSIS.'
    ]

    out = run(cmds)
    lines = out.split('\r\n')
    output = parse_output(lines)
    for line in output:
        print (line)
    print ('Done.')


if __name__ == '__main__':
    main()
