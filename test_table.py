import unittest

from table import Table


class TableTest(unittest.TestCase):
    def setUp(self):
        test_output = """
            |_____|______________|_________|_______|_____________|__________________|
            |     |              |A        |B      |C            |D                 |
            |_____|______________|_________|_______|_____________|__________________|
            |Merge|Something     |146      |4.6    |4.6          |4.6               |
            |     |______________|_________|_______|_____________|__________________|
            |     |Something Else|3060     |95.4   |95.4         |100.0             |
            |     |With Wrapping |         |       |             |                  |
            |_____|______________|_________|_______|_____________|__________________|
            |Total               |3206     |100.0  |100.0        |                  |
            |_____|______________|_________|_______|_____________|__________________|
        """

        test_lines = [x.strip() for x in test_output.split('\n')]
        self.test_lines = [x for x in test_lines if x]

    def test_get_row_type(self):
        row = self.test_lines[0].split('|')[1:-1]
        row_type = Table.get_row_type(row, 6)
        self.assertEqual('separator', row_type)

        row = self.test_lines[3].split('|')[1:-1]
        row_type = Table.get_row_type(row, 6)
        self.assertEqual('data', row_type)

        row = self.test_lines[8].split('|')[1:-1]
        row_type = Table.get_row_type(row, 6)
        self.assertEqual('ignore', row_type)

    def test_process_data(self):
        data = Table.process_data(self.test_lines, 6)
        expected = (
            ['Merge', 'Something', 146, 4.6, 4.6, 4.6],
            ['Merge', 'Something Else With Wrapping', 3060, 95.4, 95.4, 100.0]
        )
        self.assertEqual(2, len(data))
        self.assertItemsEqual(expected, data)

    def test_construct_table(self):
        t = Table(self.test_lines)
        expected = ['', '', 'A', 'B', 'C', 'D']
        self.assertItemsEqual(t.headers, expected)

        expected = (
            ['Merge', 'Something', 146, 4.6, 4.6, 4.6],
            ['Merge', 'Something Else With Wrapping', 3060, 95.4, 95.4, 100.0]
        )
        self.assertItemsEqual(t.data, expected)

    def test_get_cell_value(self):
        t = Table(self.test_lines)
        v = t.get_cell_value(0, 0)
        self.assertEqual('Merge', v)

        v = t.get_cell_value(0, 2)
        self.assertEqual(146, v)

        v = t.get_cell_value(1, 0)
        self.assertEqual('Merge', v)

        v = t.get_cell_value(1, 1)
        self.assertEqual('Something Else With Wrapping', v)

        v = t.get_cell_value(1, 5)
        self.assertEqual(100.0, v)

    def test_get_row(self):
        t = Table(self.test_lines)

        expected = ['Merge', 'Something', 146, 4.6, 4.6, 4.6]
        self.assertEqual(expected, t.get_row(0))

        expected = ['Merge', 'Something Else With Wrapping', 3060, 95.4, 95.4, 100.0]
        self.assertEqual(expected, t.get_row(1))


    def test_get_col_by_index(self):
        t = Table(self.test_lines)
        expected = ['Merge', 'Merge']
        self.assertEqual(expected, t.get_col_by_index(0))

        expected = [146, 3060]
        self.assertEqual(expected, t.get_col_by_index(2))

        expected = [4.6, 100.0]
        self.assertEqual(expected, t.get_col_by_index(5))

    def test_get_col_by_name(self):
        t = Table(self.test_lines)
        expected = ['Merge', 'Merge']
        self.assertEqual(expected, t.get_col_by_name(''))

        expected = [146, 3060]
        self.assertEqual(expected, t.get_col_by_name('A'))

        expected = [4.6, 100.0]
        self.assertEqual(expected, t.get_col_by_name('D'))

    def test_repr(self):
        t = Table(self.test_lines)
        expected = """
            [['', '', 'A', 'B', 'C', 'D'],
             ['Merge', 'Something', 146, 4.6, 4.6, 4.6],
             ['Merge', 'Something Else With Wrapping', 3060, 95.4, 95.4, 100.0]]
        """.strip().replace(' ', '')
        actual = repr(t).strip().replace(' ', '')
        self.assertEqual(expected, actual)
