import unittest

from part_a_spss_analyses import build_selection_for_camp, build_grade_levels, \
    build_selection_for_comparison_schools


class SelectionTest(unittest.TestCase):
    def setUp(self):
        pass

    def test_build_selection_for_own_school(self):
        expected = """USE ALL.
                  COMPUTE filter_$=(School_Name = 4).
                  VARIABLE LABELS filter_$ 'School_Name = 4 (FILTER)'.
                  VALUE LABELS filter_$ 0 'Not Selected' 1 'Selected'.
                  FORMATS filter_$ (f1.0).
                  FILTER BY filter_$.
                  EXECUTE."""

        result = build_selection_for_camp(dict(code=4))
        self.assertEqual(expected, result)

    def test_build_grade_levels(self):
        grades = {'5th': 1, '8th': 1, '11th': 1}
        expected = 'grade_level=1 or grade_level=2 or grade_level=3'
        result = build_grade_levels(grades)
        self.assertEqual(expected, result)

    def test_get_line_indexes_ex11(self):
        num_grades = 3
        i = 0
        indexes = []
        while i < num_grades:
            indexes.append(i + 2)
        return indexes

    def test_build_selection_for_comparison_schools(self):
        stakeholder_name = 'staff'
        school = {'code': 4, 'sector': 1, 'grades': {'5th': 1, '8th': 1, '11th': 1}}
        expected = """USE ALL.
        COMPUTE filter_$=(School_Name <> 4 and School_Denomination_3_Groups=1
        and (grade_level=1 or grade_level=2 or grade_level=3)).
        VARIABLE LABELS filter_$ 'School_Name <> 4 and School_Denomination_3_Groups=1
        and (grade_level=1 or grade_level=2 or grade_level=3)(FILTER)'.
        VALUE LABELS filter_$ 0 'Not Selected' 1 'Selected'.
        FORMATS filter_$ (f1.0).
        FILTER BY filter_$.
        EXECUTE.""".strip()
        result = build_selection_for_comparison_schools(school, stakeholder_name)
        self.assertEqual(expected, result.strip())
