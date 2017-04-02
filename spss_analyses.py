import spss
import spssaux
import sys
from cStringIO import StringIO

def run(cmd_list):
    orig_stdout = sys.stdout
    try:
        sys.stdout = StringIO()
        spss.Submit(cmd_list)
        result = sys.stdout.getvalue()
        sys.stdout.close()
        return result
    except Exception as e:
        print e
    finally:
        sys.stdout = orig_stdout


def main():
    spssaux.OpenDataFile(r'C:\Users\Liat\Google Drive\102-04  ACF Hebrew in Jewish Day Schools\Data Bases and Data Files\Survey Responses\Student Responses\Researching_Hebrew_in_Day_Schools_Student_Survey__V22 LS 20161226 USE THIS.sav')
    assert spss.IsBackendReady()

    cmds = [
        """CTABLES
  /VLABELS VARIABLES=
    A_D_important_hebrew_connects_jewsYes_No
    A_D_important_hebrew_part_group_mix_hebrewYes_No
    A_D_important_hebrew_part_being_jewishYes_No
    A_D_important_hebrew_maintains_jewish_languageYes_No
    A_D_important_hebrew_connect_israelYes_No
    A_D_important_hebrew_prepares_aliyaYes_No
    A_D_important_hebrew_helps_visit_israelYes_No
    A_D_important_hebrew_read_modern_israel_booksYes_No
    A_D_important_hebrew_communicate_jews_worldYes_No
    A_D_important_hebrew_communicate_hebrew_speakersYes_No
    A_D_important_hebrew_learn_2ndlanguageYes_No
    DISPLAY=LABEL
  /TABLE
    A_D_important_hebrew_connects_jewsYes_No [C][COLPCT.COUNT PCT40.1] +
    A_D_important_hebrew_part_group_mix_hebrewYes_No [C][COLPCT.COUNT PCT40.1] +
    A_D_important_hebrew_part_being_jewishYes_No [C][COLPCT.COUNT PCT40.1] +
    A_D_important_hebrew_maintains_jewish_languageYes_No [C][COLPCT.COUNT PCT40.1] +
    A_D_important_hebrew_connect_israelYes_No  [C][COLPCT.COUNT PCT40.1] +
    A_D_important_hebrew_prepares_aliyaYes_No [C][COLPCT.COUNT PCT40.1] +
    A_D_important_hebrew_helps_visit_israelYes_No [C][COLPCT.COUNT PCT40.1] +
    A_D_important_hebrew_read_modern_israel_booksYes_No [C][COLPCT.COUNT PCT40.1] +
    A_D_important_hebrew_communicate_jews_worldYes_No   [C][COLPCT.COUNT PCT40.1] +
    A_D_important_hebrew_communicate_hebrew_speakersYes_No [C][COLPCT.COUNT PCT40.1] +
    A_D_important_hebrew_learn_2ndlanguageYes_No [C][COLPCT.COUNT PCT40.1]
  /CATEGORIES VARIABLES= A_D_important_hebrew_connects_jewsYes_No [0, 1, OTHERNM] EMPTY=INCLUDE
  /CATEGORIES VARIABLES= A_D_important_hebrew_part_group_mix_hebrewYes_No [0, 1, OTHERNM]
    EMPTY=INCLUDE
  /CATEGORIES VARIABLES=A_D_important_hebrew_part_being_jewishYes_No [0, 1, OTHERNM]
    EMPTY=INCLUDE
  /CATEGORIES VARIABLES= A_D_important_hebrew_maintains_jewish_languageYes_No  [0, 1, OTHERNM] EMPTY=INCLUDE
  /CATEGORIES VARIABLES=A_D_important_hebrew_connect_israelYes_No [0, 1, OTHERNM] EMPTY=INCLUDE
  /CATEGORIES VARIABLES= A_D_important_hebrew_prepares_aliyaYes_No [0, 1, OTHERNM] EMPTY=INCLUDE
  /CATEGORIES VARIABLES= A_D_important_hebrew_helps_visit_israelYes_No [0, 1, OTHERNM] EMPTY=INCLUDE
  /CATEGORIES VARIABLES= A_D_important_hebrew_read_modern_israel_booksYes_No [0, 1, OTHERNM]
    EMPTY=INCLUDE
 /CATEGORIES VARIABLES=A_D_important_hebrew_communicate_jews_worldYes_No  [0, 1, OTHERNM]
    EMPTY=INCLUDE
  /CATEGORIES VARIABLES= A_D_important_hebrew_communicate_hebrew_speakersYes_No [0, 1, OTHERNM] EMPTY=INCLUDE
  /CATEGORIES VARIABLES= A_D_important_hebrew_learn_2ndlanguageYes_No [0, 1, OTHERNM]
    EMPTY=INCLUDE
 /TITLES
   TITLE='*Exhibit 6: Why is Heb for comm important.'."""
    ]

    out = run(cmds)
    lines = out.split('\r\n')
    for line in lines:
        print line
    print 'Done.'


if __name__ == '__main__':
    main()

#def spss_syntax():
    """
    """
