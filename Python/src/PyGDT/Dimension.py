
class SizeDimension:
    
    def __init__(self, nominal, symmetric_plus=0, asymmetric_minus=None, modifier="RFS"):
        self.modifier = modifier
        self.nominal = nominal
        self.plus = symmetric_plus
        if asymmetric_minus is None:
            self.minus = self.plus
        else:
            self.minus = asymmetric_minus

class OutsideDimension(SizeDimension):
    
    def __init__(self, nominal, plus, minus, modifier="RFS"):
        super().__init__(nominal, plus, minus, modifier="RFS")
    
    def _set_MMC(self):
        self.MMC = self.nominal + self.plus
    
    def _set_LMC(self):
        self.LMC = self.nominal - self.minus

class InsideDimension(SizeDimension):
    
    def __init__(self, nominal, plus, minus, modifier="RFS"):
        super().__init__(nominal, plus, minus, modifier="RFS")
    
    def _set_MMC(self):
        self.MMC = self.nominal - self.minus
    
    def _set_LMC(self):
        self.LMC = self.nominal + self.plus

class PositionalTolerance:
    
    def __init__(self, value):
        pass