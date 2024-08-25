import sys
import unittest

from docrobot.guimain import MainWindow

sys.path.insert(0, '../')

class MyTestCase(unittest.TestCase):
    def setUp(self):
        self.window = MainWindow()

    def tearDown(self):
        pass

    def test_something(self):
        self.window.prepare_rdsummary_table()
        self.assertEqual(True, False)  # add assertion here


if __name__ == '__main__':
    unittest.main()
