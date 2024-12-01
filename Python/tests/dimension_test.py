import sys
import unittest

sys.path.append('../src')
from PyGDT.Dimension import (SizeDimension,
                   OutsideDimension,
                   InsideDimension)

class TestSizeDimension(unittest.TestCase):
    
    def test_init_symmetric_tolerance(self):
        d = SizeDimension(1, .01, modifier="MMC")
        self.assertEqual(d.modifier, "MMC")
        self.assertEqual(d.nominal, 1)
        self.assertEqual(d.plus, .01)
        self.assertEqual(d.minus, .01)
    
    def test_init_asymmetric_tolerance(self):
        d = SizeDimension(1, .01, .02)
        self.assertEqual(d.modifier, "RFS")
        self.assertEqual(d.nominal, 1)
        self.assertEqual(d.plus, .01)
        self.assertEqual(d.minus, .02)

if __name__ == "__main__":
    unittest.main()