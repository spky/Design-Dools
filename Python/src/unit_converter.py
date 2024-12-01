import yaml

class Unit:
    
    def __init__(self, name, type_, conversion_factor):
        self.name = name
        self.type_ = type_
        self.conversion_factor = conversion_factor

class UnitConverter:
    
    def __init__(self, unit_definition_filepath, precision_places=6):
        self.precision_places = 6
        self.filepath = unit_definition_filepath
        with open(unit_definition_filepath, "r", encoding="utf8") as file:
            try:
                self.unit_info = yaml.safe_load(file)
            except yaml.YAMLError as exc:
                print(exc)
        
        # Check if there are any duplicate names in the yaml file
        duplicates = self._duplicates(self.available_units)
        if duplicates:
            print("WARNING:")
            print(self.filepath + " has duplicate unit names:")
            print(duplicates)
        self.initialize_units()
        
    @property
    def available_units(self):
        units = []
        for unit_type in self.unit_info:
            for unit in self.unit_info[unit_type]:
                units.append(unit)
        return units
    
    
    def initialize_units(self):
        self.units = {}
        for unit_type in self.unit_info:
            unit_info_dict = self.unit_info[unit_type]
            for unit in unit_info_dict:
                self.units[unit] = Unit(unit, unit_type, unit_info_dict[unit])
    
    def convert(self, value, from_unit, to_unit):
        f_unit = self.units[from_unit]
        t_unit = self.units[to_unit]
        
        if f_unit.type_ == t_unit.type_:
            fcf = f_unit.conversion_factor
            tcf = t_unit.conversion_factor
            
            # Convert using an equation or by using a constant depending 
            # on the unit type/combo
            if isinstance(fcf,dict):
                default_from_value = eval(fcf["to_default"], 
                                          None,
                                          {"X": value})
            else:
                default_from_value = value * fcf
            if isinstance(tcf, dict):
                converted_value = eval(tcf["from_default"], 
                                       None,
                                       {"X": default_from_value})
            else:
                converted_value = default_from_value/tcf
            
            return round(converted_value, self.precision_places)
        else:
            print("Cannot convert " + f_unit.type_ + " to " + t_unit.type_)
            return None
    
    
    def _duplicates(self,list_):
        if len(list_) != len(set(list_)):
            
            seen = set()
            dupes = []
            for unit in list_:
                if unit in seen:
                    dupes.append(unit)
                else:
                    seen.add(unit)
            return dupes
        else:
            return None