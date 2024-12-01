import sys
import unittest

sys.path.append('../src')
from PyGDT.unit_converter import UnitConverter

unit_info_file = "unit_definition.yaml"

uc = UnitConverter(unit_info_file)

print(uc.available_units)
print(uc.convert(100, "R", "K"))