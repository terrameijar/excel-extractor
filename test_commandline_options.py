import unittest
import retrieve_data # this is where the program module goes

class TestCommandlineOptions(unittest.TestCase):
    def test_no_input(self):
        self.assertRaises(SystemExit, retrieve_data.parse_args, [])

    def test_incorrect_argument(self):
        self.assertRaises(SystemExit, retrieve_data.parse_args, ['5'])

    def test_7_days(self):
        self.assertEqual(7, retrieve_data.parse_args(['7']))

    def test_15_days(self):
        self.assertEqual(15, retrieve_data.parse_args(['15']))

    def test_30_days(self):
        self.assertEqual(30, retrieve_data.parse_args(['30']))

if __name__ == '__main__':
    unittest.main()
