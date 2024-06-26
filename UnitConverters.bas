Attribute VB_Name = "UnitConverters"
Option Explicit

Public Function bg_UnitType(unit As String) As String
    
    Select Case unit
    Case "kg/s", "kg/min", "kg/hr", "g/s", "g/min", "g/hr", "lbm/s", "lbm/min", "lbm/hr"
        bg_UnitType = "Mass Flow"
    Case "ccs", "m3/s", "m3/min", "m3/hr", "cm3/s", "cm3/min", "cm3/hr", "L/s", "lps", "L/min", "lpm", "L/hr", "in3/s", "in3/min", "in3/hr", "ft3/s", "ft3/min", "ft3/hr", "gal/s", "gps", "gal/min", "gpm", "gal/hr"
        bg_UnitType = "Volume Flow"
    Case "kg/m3", "kg/cm3", "kg/L", "g/m3", "g/cm3", "g/L", "lbm/in3", "lbm/ft3", "lbm/gal"
        bg_UnitType = "Density"
    Case "scms", "scmm", "scmh", "slps", "slpm", "slph", "sccs", "sccm", "scch", "scfs", "scfm", "scfh", "scis", "scim", "scih"
        bg_UnitType = "Standard Volume Flow"
    Case "s", "min", "hr"
        bg_UnitType = "Time"
    Case "km", "m", "cm", "mm", "in", "ft", "yd", "mi"
        bg_UnitType = "Length"
    Case "km/s", "m/s", "cm/s", "mm/s", "in/s", "ft/s", "yd/s", "mi/s", "km/min", "m/min", "cm/min", "mm/min", "in/min", "ft/min", "yd/min", "mi/min", "km/hr", "m/hr", "cm/hr", "mm/hr", "in/hr", "ft/hr", "yd/hr", "mi/hr", "mph"
        bg_UnitType = "Speed"
    Case "F", "�F", "C", "�C", "K", "R", "Rank", "Rankine"
        bg_UnitType = "Temperature"
    Case "psia", "psig", "psid", "psi", "ksi", "lbf/in2", "lb/in2", "mmHg", "atm", "Pa", "kPa", "MPa", "GPa", "torr", "Torr", "bar"
        bg_UnitType = "Pressure"
    Case "km2", "m2", "cm2", "mm2", "in2", "ft2", "yd2", "mi2", "km^2", "m^2", "cm^2", "mm^2", "in^2", "ft^2", "yd^2", "mi^2"
        bg_UnitType = "Area"
    Case "km3", "m3", "cm3", "mm3", "in3", "ft3", "yd3", "mi3", "ml", "mL", "L", "gal", "km^3", "m^3", "cm^3", "mm^3", "in^3", "ft^3", "yd^3", "mi^3"
        bg_UnitType = "Volume"
    Case "km4", "m4", "cm4", "mm4", "in4", "ft4", "yd4", "mi4", "km^4", "m^4", "cm^4", "mm^4", "in^4", "ft^4", "yd^4", "mi^4"
        bg_UnitType = "Moment of Inertia"
    Case "lbm", "ozm", "kg", "g", "mg", "grams"
        bg_UnitType = "Mass"
    Case "kN", "N", "mN", "Newton", "lbf", "ozf"
        bg_UnitType = "Force"
    Case "N-m", "mN-m", "in-lb", "ft-lb", "in-oz"
        bg_UnitType = "Torque"
    Case Else
        bg_UnitType = "Unit Error, " & unit & " not recognized"
    End Select

End Function



Public Function bg_MassFlowConvert(Value As Double, from_unit As String, to_unit As String) As Double
    'converts mass flow units between each other
    
    'default unit is going to be kg/s
    Dim default_unit_value As Double
    
    'convert everything to the default unit
    Select Case from_unit
    Case "kg/s", "kg/sec"
    'Case InStr(KGPS, from_unit) > 0
        default_unit_value = Value
    Case "kg/min"
        default_unit_value = Value / 60
    Case "kg/hr"
        default_unit_value = Value / 3600
    Case "g/s", "g/sec"
        default_unit_value = Value / 1000
    Case "g/min"
        default_unit_value = Value / 1000 / 60
    Case "g/hr"
        default_unit_value = Value / 1000 / 3600
    Case "lbm/s", "lbm/sec"
        default_unit_value = Value * 0.453592
    Case "lbm/min"
        default_unit_value = Value * 0.453592 / 60
    Case "lbm/hr"
        default_unit_value = Value * 0.453592 / 3600
    End Select
    
    Dim new_unit_value As Double
    'convert from the default unit to the output unit
    Select Case to_unit
    Case "kg/s"
        new_unit_value = default_unit_value
    Case "kg/min"
        new_unit_value = default_unit_value * 60
    Case "kg/hr"
        new_unit_value = default_unit_value * 3600
    Case "g/s"
        new_unit_value = default_unit_value * 1000
    Case "g/min"
        new_unit_value = default_unit_value * 1000 * 60
    Case "g/hr"
        new_unit_value = default_unit_value * 1000 * 3600
    Case "lbm/s"
        new_unit_value = default_unit_value / 0.453592
    Case "lbm/min"
        new_unit_value = default_unit_value / 0.453592 * 60
    Case "lbm/hr"
        new_unit_value = default_unit_value / 0.453592 * 3600
    End Select
    
    bg_MassFlowConvert = new_unit_value

End Function

Public Function bg_VolumeFlowConvert(Value As Double, from_unit As String, to_unit As String) As Double
    'converts volume flow units between each other
    
    'default unit is going to be m3/s
    Dim default_unit_value As Double
    
    'TO-DO - finish changing all these to the right volume conversion factors
    'convert everything to the default unit
    Select Case from_unit
    Case "m3/s"
        default_unit_value = Value
    Case "m3/min"
        default_unit_value = Value / 60
    Case "m3/hr"
        default_unit_value = Value / 3600
    Case "cm3/s", "ccs"
        default_unit_value = Value / 100 ^ 3
    Case "cm3/min", "ccm"
        default_unit_value = Value / 100 ^ 3 / 60
    Case "cm3/hr"
        default_unit_value = Value / 100 ^ 3 / 3600
    Case "L/s", "lps"
        default_unit_value = Value / 1000
    Case "L/min", "lpm"
        default_unit_value = Value / 1000 / 60
    Case "L/hr"
        default_unit_value = Value / 1000 / 3600
    Case "in3/s"
        default_unit_value = Value * 0.000016387
    Case "in3/min"
        default_unit_value = Value * 0.000016387 / 60
    Case "in3/hr"
        default_unit_value = Value * 0.000016387 / 3600
    Case "ft3/s"
        default_unit_value = Value * 0.0283168
    Case "ft3/min"
        default_unit_value = Value * 0.0283168 / 60
    Case "ft3/hr"
        default_unit_value = Value * 0.0283168 / 3600
    Case "gal/s", "gps"
        default_unit_value = Value / 264.172
    Case "gal/min", "gpm"
        default_unit_value = Value / 264.172 / 60
    Case "gal/hr"
        default_unit_value = Value / 264.172 / 3600
    End Select
    
    Dim new_unit_value As Double
    'convert from the default unit to the output unit
    Select Case to_unit
    Case "m3/s"
        new_unit_value = default_unit_value
    Case "m3/m", "m3/min"
        new_unit_value = default_unit_value * 60
    Case "m3/hr"
        new_unit_value = default_unit_value * 3600
    Case "cm3/s", "ccs"
        new_unit_value = default_unit_value * 100 ^ 3
    Case "cm3/min", "ccm"
        new_unit_value = default_unit_value * 100 ^ 3 * 60
    Case "cm3/hr"
        new_unit_value = default_unit_value * 100 ^ 3 * 3600
    Case "L/s", "lps"
        new_unit_value = default_unit_value * 1000
    Case "L/min", "lpm"
        new_unit_value = default_unit_value * 1000 * 60
    Case "L/hr"
        new_unit_value = default_unit_value * 1000 * 3600
    Case "in3/s"
        new_unit_value = default_unit_value / 0.000016387
    Case "in3/min"
        new_unit_value = default_unit_value / 0.000016387 * 60
    Case "in3/hr"
        new_unit_value = default_unit_value / 0.000016387 * 3600
    Case "ft3/s"
        new_unit_value = default_unit_value / 0.0283168
    Case "ft3/min"
        new_unit_value = default_unit_value / 0.0283168 * 60
    Case "ft3/hr"
        new_unit_value = default_unit_value / 0.0283168 * 3600
    Case "gal/s", "gps"
        new_unit_value = default_unit_value * 264.172
    Case "gal/min", "gpm"
        new_unit_value = default_unit_value * 264.172 * 60
    Case "gal/hr"
        new_unit_value = default_unit_value * 264.172 * 3600
    End Select
    
    bg_VolumeFlowConvert = new_unit_value

End Function

Public Function bg_DensityConvert(Value As Double, from_unit As String, to_unit As String) As Double
    'converts mass flow units between each other
    
    'default unit is going to be kg/m3
    Dim default_unit_value As Double
    
    'convert everything to the default unit
    Select Case from_unit
    Case "kg/m3"
        default_unit_value = Value
    Case "kg/cm3"
        default_unit_value = Value * 100 ^ 3
    Case "kg/L"
        default_unit_value = Value * 10 ^ 3
    Case "g/m3"
        default_unit_value = Value / 1000
    Case "g/cm3"
        default_unit_value = Value * 100 ^ 3 / 1000
    Case "g/L"
        default_unit_value = Value * 10 ^ 3 / 1000
    Case "lbm/in3"
        default_unit_value = Value * 27679.9
    Case "lbm/ft3"
        default_unit_value = Value * 16.018
    Case "lbm/gal"
        default_unit_value = Value * 119.826
    End Select
    
    Dim new_unit_value As Double
    Select Case to_unit
    Case "kg/m3"
        new_unit_value = Value
    Case "kg/cm3"
        new_unit_value = Value / 100 ^ 3
    Case "kg/L"
        new_unit_value = Value / 10 ^ 3
    Case "g/m3"
        new_unit_value = Value * 1000
    Case "g/cm3"
        new_unit_value = Value / 100 ^ 3 * 1000
    Case "g/L"
        new_unit_value = Value / 10 ^ 3 * 1000
    Case "lbm/in3"
        new_unit_value = Value / 27679.9
    Case "lbm/ft3"
        new_unit_value = Value / 16.018
    Case "lbm/gal"
        new_unit_value = Value / 119.826
    End Select
    
    bg_DensityConvert = new_unit_value

End Function

Public Function bg_MassFlowToStandardFlow(Value As Double, from_unit As String, to_unit As String, gas_name As String, Optional Standard As String = "IUPAC_STP", Optional P_init_Pa As Double = 0, Optional T_init_K As Double = 0) As Double
    'Convert whatever mass flow unit to standard volumetric flow based on some standard
    'Default mass flow unit is kg/s
    Dim default_unit_value As Double, std_Density As Double, default_std_unit_value As Double, default_std_unit As String
    default_unit_value = bg_MassFlowConvert(Value, from_unit, "kg/s")
    
    std_Density = fluidStandardDensity(gas_name, Standard, "kg/m3", P_init_Pa, T_init_K) ' kg/m3
    default_std_unit = "m3/s"
    default_std_unit_value = default_unit_value / std_Density 'standard m3/s
    
    Dim new_unit_value As Double
    'convert from the default unit to the output unit
    Select Case to_unit
    Case "scms" 'Standard Cubic Meter per Second
        new_unit_value = bg_VolumeFlowConvert(default_std_unit_value, default_std_unit, "m3/s")
    Case "scmm" 'Standard Cubic Meter per Minute
        new_unit_value = bg_VolumeFlowConvert(default_std_unit_value, default_std_unit, "m3/min")
    Case "scmh" 'Standard Cubic Meter per Hour
        new_unit_value = bg_VolumeFlowConvert(default_std_unit_value, default_std_unit, "m3/hr")
    Case "slps", "sls" 'Standard Liter per Second
        new_unit_value = bg_VolumeFlowConvert(default_std_unit_value, default_std_unit, "L/s")
    Case "slpm", "slm" 'Standard Liter per Minute
        new_unit_value = bg_VolumeFlowConvert(default_std_unit_value, default_std_unit, "L/min")
    Case "slph", "slh" 'Standard Liter per Hour
        new_unit_value = bg_VolumeFlowConvert(default_std_unit_value, default_std_unit, "L/hr")
    Case "sccs" 'Standard Cubic Centimeter per Second
        new_unit_value = bg_VolumeFlowConvert(default_std_unit_value, default_std_unit, "cm3/s")
    Case "sccm" 'Standard Cubic Centimeter per Minute
        new_unit_value = bg_VolumeFlowConvert(default_std_unit_value, default_std_unit, "cm3/min")
    Case "scch" 'Standard Cubic Centimeter per Hour
        new_unit_value = bg_VolumeFlowConvert(default_std_unit_value, default_std_unit, "cm3/hr")
    Case "scfs" 'Standard Cubic Feet per Second
        new_unit_value = bg_VolumeFlowConvert(default_std_unit_value, default_std_unit, "ft3/s")
    Case "scfm" 'Standard Cubic Feet per Minute
        new_unit_value = bg_VolumeFlowConvert(default_std_unit_value, default_std_unit, "ft3/min")
    Case "scfh" 'Standard Cubic Feet per Hour
        new_unit_value = bg_VolumeFlowConvert(default_std_unit_value, default_std_unit, "in3/hr")
    Case "scis" 'Standard Cubic Inches per Second
        new_unit_value = bg_VolumeFlowConvert(default_std_unit_value, default_std_unit, "in3/s")
    Case "scim" 'Standard Cubic Inches per Minute
        new_unit_value = bg_VolumeFlowConvert(default_std_unit_value, default_std_unit, "in3/min")
    Case "scih" 'Standard Cubic Inches per Hour
        new_unit_value = bg_VolumeFlowConvert(default_std_unit_value, default_std_unit, "in3/hr")
    End Select
    
    bg_MassFlowToStandardFlow = new_unit_value

End Function

Public Function bg_StandardFlowConvert(Value As Double, from_unit As String, to_unit As String) As Double
    Dim default_std_unit As String, new_unit_value As Double, default_std_unit_value As Double
    default_std_unit = "m3/s"
    
    
    Select Case from_unit
    Case "scms" 'Standard Cubic Meter per Second
        default_std_unit_value = bg_VolumeFlowConvert(Value, "m3/s", default_std_unit)
    Case "scmm" 'Standard Cubic Meter per Minute
        default_std_unit_value = bg_VolumeFlowConvert(Value, "m3/min", default_std_unit)
    Case "scmh" 'Standard Cubic Meter per Hour
        default_std_unit_value = bg_VolumeFlowConvert(Value, "m3/hr", default_std_unit)
    Case "slps", "sls" 'Standard Liter per Second
        default_std_unit_value = bg_VolumeFlowConvert(Value, "L/s", default_std_unit)
    Case "slpm", "slm" 'Standard Liter per Minute
        default_std_unit_value = bg_VolumeFlowConvert(Value, "L/min", default_std_unit)
    Case "slph", "slh" 'Standard Liter per Hour
        default_std_unit_value = bg_VolumeFlowConvert(Value, "L/hr", default_std_unit)
    Case "sccs" 'Standard Cubic Centimeter per Second
        default_std_unit_value = bg_VolumeFlowConvert(Value, "cm3/s", default_std_unit)
    Case "sccm" 'Standard Cubic Centimeter per Minute
        default_std_unit_value = bg_VolumeFlowConvert(Value, "cm3/min", default_std_unit)
    Case "scch" 'Standard Cubic Centimeter per Hour
        default_std_unit_value = bg_VolumeFlowConvert(Value, "cm3/hr", default_std_unit)
    Case "scfs" 'Standard Cubic Feet per Second
        default_std_unit_value = bg_VolumeFlowConvert(Value, "ft3/s", default_std_unit)
    Case "scfm" 'Standard Cubic Feet per Minute
        default_std_unit_value = bg_VolumeFlowConvert(Value, "ft3/min", default_std_unit)
    Case "scfh" 'Standard Cubic Feet per Hour
        default_std_unit_value = bg_VolumeFlowConvert(Value, "in3/hr", default_std_unit)
    Case "scis" 'Standard Cubic Inches per Second
        default_std_unit_value = bg_VolumeFlowConvert(Value, "in3/s", default_std_unit)
    Case "scim" 'Standard Cubic Inches per Minute
        default_std_unit_value = bg_VolumeFlowConvert(Value, "in3/min", default_std_unit)
    Case "scih" 'Standard Cubic Inches per Hour
        default_std_unit_value = bg_VolumeFlowConvert(Value, "in3/hr", default_std_unit)
    End Select
    
    Select Case to_unit
    Case "scms" 'Standard Cubic Meter per Second
        new_unit_value = bg_VolumeFlowConvert(default_std_unit_value, default_std_unit, "m3/s")
    Case "scmm" 'Standard Cubic Meter per Minute
        new_unit_value = bg_VolumeFlowConvert(default_std_unit_value, default_std_unit, "m3/min")
    Case "scmh" 'Standard Cubic Meter per Hour
        new_unit_value = bg_VolumeFlowConvert(default_std_unit_value, default_std_unit, "m3/hr")
    Case "slps", "sls" 'Standard Liter per Second
        new_unit_value = bg_VolumeFlowConvert(default_std_unit_value, default_std_unit, "L/s")
    Case "slpm", "slm" 'Standard Liter per Minute
        new_unit_value = bg_VolumeFlowConvert(default_std_unit_value, default_std_unit, "L/min")
    Case "slph", "slh" 'Standard Liter per Hour
        new_unit_value = bg_VolumeFlowConvert(default_std_unit_value, default_std_unit, "L/hr")
    Case "sccs" 'Standard Cubic Centimeter per Second
        new_unit_value = bg_VolumeFlowConvert(default_std_unit_value, default_std_unit, "cm3/s")
    Case "sccm" 'Standard Cubic Centimeter per Minute
        new_unit_value = bg_VolumeFlowConvert(default_std_unit_value, default_std_unit, "cm3/min")
    Case "scch" 'Standard Cubic Centimeter per Hour
        new_unit_value = bg_VolumeFlowConvert(default_std_unit_value, default_std_unit, "cm3/hr")
    Case "scfs" 'Standard Cubic Feet per Second
        new_unit_value = bg_VolumeFlowConvert(default_std_unit_value, default_std_unit, "ft3/s")
    Case "scfm" 'Standard Cubic Feet per Minute
        new_unit_value = bg_VolumeFlowConvert(default_std_unit_value, default_std_unit, "ft3/min")
    Case "scfh" 'Standard Cubic Feet per Hour
        new_unit_value = bg_VolumeFlowConvert(default_std_unit_value, default_std_unit, "in3/hr")
    Case "scis" 'Standard Cubic Inches per Second
        new_unit_value = bg_VolumeFlowConvert(default_std_unit_value, default_std_unit, "in3/s")
    Case "scim" 'Standard Cubic Inches per Minute
        new_unit_value = bg_VolumeFlowConvert(default_std_unit_value, default_std_unit, "in3/min")
    Case "scih" 'Standard Cubic Inches per Hour
        new_unit_value = bg_VolumeFlowConvert(default_std_unit_value, default_std_unit, "in3/hr")
    End Select
    bg_StandardFlowConvert = new_unit_value

End Function

Public Function bg_StandardFlowToMassFlow(Value As Double, from_unit As String, to_unit As String, gas_name As String, Optional Standard As String = "IUPAC_STP", Optional P_init_Pa As Double = 0, Optional T_init_K As Double = 0) As Double
    'Output the equivalent mass flow of a standard and gas volumetric flow in a given unit
    Dim default_unit_value As Double, std_Density As Double, default_std_unit_value As Double, default_std_unit As String
    default_std_unit = "scms"
    
    default_std_unit_value = bg_StandardFlowConvert(Value, from_unit, default_std_unit)
    
    std_Density = fluidStandardDensity(gas_name, Standard, "kg/m3", P_init_Pa, T_init_K) 'kg/m3
    default_unit_value = default_std_unit_value * std_Density
    bg_StandardFlowToMassFlow = bg_MassFlowConvert(default_unit_value, "kg/s", to_unit)

End Function

Public Function bg_LengthConvert(Value As Double, from_unit As String, to_unit As String) As Double
    'converts between length units using excel's convert function
    'TO-DO protect for length units
    bg_LengthConvert = Application.WorksheetFunction.Convert(Value, from_unit, to_unit)
End Function

Public Function bg_SpeedConvert(Value As Double, from_unit As String, to_unit As String) As Double
    'convert whatever the from_unit is to /s, then convert it to the to_unit using bg_LengthConvert
    
    'default value is m/s
    Dim default_time_unit_value As Double
    Dim default_unit_value As Double
    
    Dim to_time_unit_value As Double
    Dim to_unit_length_unit As String
    Dim from_unit_length_unit As String
    
    'convert time unit to /s
    Select Case from_unit
    Case "km/s", "m/s", "cm/s", "mm/s", "in/s", "ft/s", "yd/s", "mi/s"
        default_time_unit_value = Value
    Case "km/min", "m/min", "cm/min", "mm/min", "in/min", "ft/min", "yd/min", "mi/min"
        default_time_unit_value = Value * 60
    Case "km/hr", "m/hr", "cm/hr", "mm/hr", "in/hr", "ft/hr", "yd/hr", "mi/hr", "mph"
        default_time_unit_value = Value * 3600
    End Select
    
    'convert time unit to the to_unit
    Select Case to_unit
    Case "km/s", "m/s", "cm/s", "mm/s", "in/s", "ft/s", "yd/s", "mi/s"
        to_time_unit_value = default_time_unit_value
    Case "km/min", "m/min", "cm/min", "mm/min", "in/min", "ft/min", "yd/min", "mi/min"
        to_time_unit_value = default_time_unit_value * 60
    Case "km/hr", "m/hr", "cm/hr", "mm/hr", "in/hr", "ft/hr", "yd/hr", "mi/hr", "mph"
        to_time_unit_value = default_time_unit_value * 3600
    End Select
    
    'figure out to_unit's length unit
    Select Case to_unit
    Case "km/s", "km/min", "km/hr"
        to_unit_length_unit = "km"
    Case "m/s", "m/min", "m/hr"
        to_unit_length_unit = "m"
    Case "cm/s", "cm/min", "cm/hr"
        to_unit_length_unit = "cm"
    Case "mm/s", "mm/min", "mm/hr"
        to_unit_length_unit = "mm"
    Case "in/s", "in/min", "in/hr"
        to_unit_length_unit = "in"
    Case "ft/s", "ft/min", "ft/hr"
        to_unit_length_unit = "ft"
    Case "yd/s", "yd/min", "yd/hr"
        to_unit_length_unit = "yd"
    Case "mi/s", "mi/min", "mi/hr", "mph"
        to_unit_length_unit = "mi"
    End Select
    
    'figure out from_unit's length unit
    Select Case from_unit
    Case "km/s", "km/min", "km/hr"
        from_unit_length_unit = "km"
    Case "m/s", "m/min", "m/hr"
        from_unit_length_unit = "m"
    Case "cm/s", "cm/min", "cm/hr"
        from_unit_length_unit = "cm"
    Case "mm/s", "mm/min", "mm/hr"
        from_unit_length_unit = "mm"
    Case "in/s", "in/min", "in/hr"
        from_unit_length_unit = "in"
    Case "ft/s", "ft/min", "ft/hr"
        from_unit_length_unit = "ft"
    Case "yd/s", "yd/min", "yd/hr"
        from_unit_length_unit = "yd"
    Case "mi/s", "mi/min", "mi/hr", "mph"
        from_unit_length_unit = "mi"
    End Select
    
    'Final convert
    bg_SpeedConvert = bg_LengthConvert(to_time_unit_value, from_unit_length_unit, to_unit_length_unit)
End Function

Public Function bg_TemperatureConvert(Value As Double, from_unit As String, to_unit As String) As Double
    'figures out the temperature unit to input into excel's default convert function
    
    'TO-DO: protect for temperature units
    Dim from_unit_unit As String, to_unit_unit As String
    
    
    'figure out from unit's unit
    Select Case from_unit
        Case "F", "�F"
            from_unit_unit = "F"
        Case "C", "�C"
            from_unit_unit = "C"
        Case "K"
            from_unit_unit = "K"
        Case "R", "Rank", "Rankine"
            from_unit_unit = "Rank"
    End Select
    
    'figure out to unit's unit
    Select Case to_unit
        Case "F", "�F"
            to_unit_unit = "F"
        Case "C", "�C"
            to_unit_unit = "C"
        Case "K"
            to_unit_unit = "K"
        Case "R", "Rank", "Rankine"
            to_unit_unit = "Rank"
    End Select
    
    'Final convert
    bg_TemperatureConvert = Application.WorksheetFunction.Convert(Value, from_unit_unit, to_unit_unit)

End Function

Public Function bg_AreaConvert(Value As Double, from_unit As String, to_unit As String) As Double
    'figures out the Area unit to input into excel's default convert function
    
    'TO-DO: protect for area units
    Dim from_unit_unit As String, to_unit_unit As String
    
    
    'figure out from unit's unit
    Select Case from_unit
        Case "km2", "km^2"
            from_unit_unit = "km^2"
        Case "m2", "m^2"
            from_unit_unit = "m^2"
        Case "cm2", "cm^2"
            from_unit_unit = "cm^2"
        Case "mm2", "mm^2"
            from_unit_unit = "mm^2"
        Case "in2", "in^2"
            from_unit_unit = "in^2"
        Case "ft2", "ft^2"
            from_unit_unit = "ft^2"
        Case "yd2", "yd^2"
            from_unit_unit = "yd^2"
        Case "mi2", "mi^2"
            from_unit_unit = "mi^2"
    End Select
    
    'figure out to unit's unit
    Select Case to_unit
        Case "km2", "km^2"
            to_unit_unit = "km^2"
        Case "m2", "m^2"
            to_unit_unit = "m^2"
        Case "cm2", "cm^2"
            to_unit_unit = "cm^2"
        Case "mm2", "mm^2"
            to_unit_unit = "mm^2"
        Case "in2", "in^2"
            to_unit_unit = "in^2"
        Case "ft2", "ft^2"
            to_unit_unit = "ft^2"
        Case "yd2", "yd^2"
            to_unit_unit = "yd^2"
        Case "mi2", "mi^2"
            to_unit_unit = "mi^2"
    End Select
    
    'Final convert
    bg_AreaConvert = Application.WorksheetFunction.Convert(Value, from_unit_unit, to_unit_unit)

End Function

Public Function bg_VolumeConvert(Value As Double, from_unit As String, to_unit As String) As Double
    'figures out the Area unit to input into excel's default convert function
    
    'TO-DO: protect for area units
    Dim from_unit_unit As String, to_unit_unit As String
    
    
    'figure out from unit's unit
    Select Case from_unit
        Case "km3", "km^3"
            from_unit_unit = "km^3"
        Case "m3", "m^3"
            from_unit_unit = "m^3"
        Case "cm3", "cm^3"
            from_unit_unit = "cm^3"
        Case "mm3", "mm^3"
            from_unit_unit = "mm^3"
        Case "in3", "in^3"
            from_unit_unit = "in^3"
        Case "ft3", "ft^3"
            from_unit_unit = "ft^3"
        Case "yd3", "yd^3"
            from_unit_unit = "yd^3"
        Case "mi3", "mi^3"
            from_unit_unit = "mi^3"
        Case Else
            from_unit_unit = from_unit
    End Select
    
    'figure out to unit's unit
    Select Case to_unit
        Case "km3", "km^3"
            to_unit_unit = "km^3"
        Case "m3", "m^3"
            to_unit_unit = "m^3"
        Case "cm3", "cm^3"
            to_unit_unit = "cm^3"
        Case "mm3", "mm^3"
            to_unit_unit = "mm^3"
        Case "in3", "in^3"
            to_unit_unit = "in^3"
        Case "ft3", "ft^3"
            to_unit_unit = "ft^3"
        Case "yd3", "yd^3"
            to_unit_unit = "yd^3"
        Case "mi3", "mi^3"
            to_unit_unit = "mi^3"
        Case Else
            to_unit_unit = to_unit
    End Select
    
    'Final convert
    bg_VolumeConvert = Application.WorksheetFunction.Convert(Value, from_unit_unit, to_unit_unit)

End Function

Public Function bg_MomentOfInertiaConvert(Value As Double, from_unit As String, to_unit As String) As Double
    'figures out the Area unit to input into excel's default convert function
    
    'TO-DO: protect for area units
    Dim from_unit_unit As String, to_unit_unit As String
    
    
    'figure out from unit's unit
    Select Case from_unit
        Case "km4", "km^4"
            from_unit_unit = "km^2"
        Case "m4", "m^4"
            from_unit_unit = "m^2"
        Case "cm4", "cm^4"
            from_unit_unit = "cm^2"
        Case "mm4", "mm^4"
            from_unit_unit = "mm^2"
        Case "in4", "in^4"
            from_unit_unit = "in^2"
        Case "ft4", "ft^4"
            from_unit_unit = "ft^2"
        Case "yd4", "yd^4"
            from_unit_unit = "yd^2"
        Case "mi4", "mi^4"
            from_unit_unit = "mi^2"
    End Select
    
    'figure out to unit's unit
    Select Case to_unit
        Case "km4", "km^4"
            to_unit_unit = "km^2"
        Case "m4", "m^4"
            to_unit_unit = "m^2"
        Case "cm4", "cm^4"
            to_unit_unit = "cm^2"
        Case "mm4", "mm^4"
            to_unit_unit = "mm^2"
        Case "in4", "in^4"
            to_unit_unit = "in^2"
        Case "ft4", "ft^4"
            to_unit_unit = "ft^2"
        Case "yd4", "yd^4"
            to_unit_unit = "yd^2"
        Case "mi4", "mi^4"
            to_unit_unit = "mi^2"
    End Select
    
    'Final convert
    bg_MomentOfInertiaConvert = Application.WorksheetFunction.Convert(Application.WorksheetFunction.Convert(Value, from_unit_unit, to_unit_unit), from_unit_unit, to_unit_unit)

End Function


Public Function bg_PressureConvert(Value As Double, from_unit As String, to_unit As String) As Double
    'figures out the pressure unit and converts it either with excel's default convert function or otherwise
    'TO-DO protect for pressure units
    
    Dim from_unit_unit As String, deciphered_value As Double
    Dim to_unit_unit As String
    'decipher weird units and convert from_unit
    Select Case from_unit
        Case "psia", "psid", "psi", "lbf/in2", "lb/in2"
            from_unit_unit = "psi"
            deciphered_value = Value
        Case "ksi"
            deciphered_value = Value * 1000
            from_unit_unit = "psi"
        Case "Pa"
            deciphered_value = Value
            from_unit_unit = "Pa"
        Case "kPa"
            deciphered_value = Value * 1000
            from_unit_unit = "Pa"
        Case "MPa"
            deciphered_value = Value * 1000000
            from_unit_unit = "Pa"
        Case "GPa"
            deciphered_value = Value * 1000000000
            from_unit_unit = "Pa"
        Case "bar"
            deciphered_value = Value * 100000
            from_unit_unit = "Pa"
        Case "Torr", "torr"
            deciphered_value = Value
            from_unit_unit = "Torr"
        Case Else
            from_unit_unit = from_unit
            deciphered_value = Value
    End Select
    
    Select Case to_unit
        Case "psia", "psid", "psi", "lbf/in2", "lb/in2"
            to_unit_unit = "psi"
            bg_PressureConvert = Application.WorksheetFunction.Convert(deciphered_value, from_unit_unit, to_unit_unit)
        Case "ksi"
            to_unit_unit = "psi"
            bg_PressureConvert = Application.WorksheetFunction.Convert(deciphered_value, from_unit_unit, to_unit_unit) / 1000
        Case "Pa"
            to_unit_unit = "Pa"
            bg_PressureConvert = Application.WorksheetFunction.Convert(deciphered_value, from_unit_unit, to_unit_unit)
        Case "kPa"
            to_unit_unit = "Pa"
            bg_PressureConvert = Application.WorksheetFunction.Convert(deciphered_value, from_unit_unit, to_unit_unit) / 1000
        Case "MPa"
            to_unit_unit = "Pa"
            bg_PressureConvert = Application.WorksheetFunction.Convert(deciphered_value, from_unit_unit, to_unit_unit) / 1000000
        Case "GPa"
            to_unit_unit = "Pa"
            bg_PressureConvert = Application.WorksheetFunction.Convert(deciphered_value, from_unit_unit, to_unit_unit) / 1000000000
        Case "bar"
            to_unit_unit = "Pa"
            bg_PressureConvert = Application.WorksheetFunction.Convert(deciphered_value, from_unit_unit, to_unit_unit) / 100000
        Case "Torr", "torr"
            to_unit_unit = "Torr"
            bg_PressureConvert = Application.WorksheetFunction.Convert(deciphered_value, from_unit_unit, to_unit_unit)
        Case Else
            to_unit_unit = to_unit
            bg_PressureConvert = Application.WorksheetFunction.Convert(deciphered_value, from_unit_unit, to_unit_unit)
    End Select
    
End Function

Public Function bg_TimeConvert(Value As Double, from_unit As String, to_unit As String) As Double
    'converts between time units using excel's convert function
    'TO-DO protect for time units
    bg_TimeConvert = Application.WorksheetFunction.Convert(Value, from_unit, to_unit)
End Function

Public Function bg_MassConvert(Value As Double, from_unit As String, to_unit As String) As Double
    'converts between mass units using excel's convert function or other options if needed
    'converts the from_unit value to grams and then to the to_unit
    
    Dim default_unit As String, deciphered_value As Double, output As Double, default_unit_value As Double
    Dim from_unit_unit As String, to_unit_unit As String
    
    default_unit = "g"
    
    
    Select Case from_unit
        Case "g", "kg", "lbm", "ozm"
            deciphered_value = Value
            from_unit_unit = from_unit
        Case "grams"
            deciphered_value = Value
            from_unit_unit = "g"
        Case "mg"
            deciphered_value = Value / 1000
            from_unit_unit = "g"
    End Select
    
    default_unit_value = Application.WorksheetFunction.Convert(deciphered_value, from_unit_unit, default_unit)
    
    Select Case to_unit
        Case "g", "kg", "lbm", "ozm"
            to_unit_unit = to_unit
            output = Application.WorksheetFunction.Convert(default_unit_value, default_unit, to_unit_unit)
        Case "grams"
            to_unit_unit = "g"
            output = Application.WorksheetFunction.Convert(default_unit_value, default_unit, to_unit_unit)
        Case "mg"
            to_unit_unit = "g"
            output = Application.WorksheetFunction.Convert(default_unit_value, default_unit, to_unit_unit) * 1000
    End Select
    
    bg_MassConvert = output
    
End Function


Public Function bg_ForceConvert(Value As Double, from_unit As String, to_unit As String) As Double
    'converts between force units using excel's convert function or other options if needed
    'converts the from_unit value to Newtons and then to the to_unit
    
    Dim default_unit As String, deciphered_value As Double, output As Double, default_unit_value As Double
    Dim from_unit_unit As String, to_unit_unit As String
    default_unit = "N"
    
    
    Select Case from_unit
        Case "N", "lbf"
            deciphered_value = Value
            from_unit_unit = from_unit
        Case "Newton"
            deciphered_value = Value
            from_unit_unit = "N"
        Case "kN"
            deciphered_value = Value * 1000
            from_unit_unit = "N"
        Case "mN"
            deciphered_value = Value / 1000
            from_unit_unit = "N"
        Case "ozf"
            deciphered_value = Value / 16
            from_unit_unit = "lbf"
    End Select
    
    default_unit_value = Application.WorksheetFunction.Convert(deciphered_value, from_unit_unit, default_unit)
    
    Select Case to_unit
        Case "N", "Newton"
            to_unit_unit = "N"
            output = Application.WorksheetFunction.Convert(default_unit_value, default_unit, to_unit_unit)
        Case "kN"
            to_unit_unit = "N"
            output = Application.WorksheetFunction.Convert(default_unit_value, default_unit, to_unit_unit) / 1000
        Case "mN"
            to_unit_unit = "N"
            output = Application.WorksheetFunction.Convert(default_unit_value, default_unit, to_unit_unit) * 1000
        Case "lbf"
            to_unit_unit = "lbf"
            output = Application.WorksheetFunction.Convert(default_unit_value, default_unit, to_unit_unit)
        Case "ozf"
            to_unit_unit = "lbf"
            output = Application.WorksheetFunction.Convert(default_unit_value, default_unit, to_unit_unit) * 16
    End Select
    
    bg_ForceConvert = output
End Function


Public Function bg_TorqueConvert(Value As Double, from_unit As String, to_unit As String) As Double
    'converts between torque units using other functions Case "N-m", "mN-m", "in-lb", "ft-lb", "in-oz"
    'converts between force units using excel's convert function or other options if needed
    'converts the from_unit value to Newtons and then to the to_unit
    
    Dim default_unit As String, deciphered_value As Double, output As Double, default_unit_value As Double
    Dim from_unit_unit As String, to_unit_unit As String
    default_unit = "N-m"
    
    
    Select Case from_unit
        Case "N-m"
            default_unit_value = Value
        Case "mN-m"
            default_unit_value = bg_ForceConvert(Value, "mN", "N")
        Case "in-lb"
            default_unit_value = bg_LengthConvert(bg_ForceConvert(Value, "lbf", "N"), "in", "m")
        Case "ft-lb"
            default_unit_value = bg_LengthConvert(bg_ForceConvert(Value, "lbf", "N"), "ft", "m")
        Case "in-oz"
            default_unit_value = bg_LengthConvert(bg_ForceConvert(Value, "ozf", "N"), "in", "m")
    End Select
    
    from_unit_unit = default_unit
    
    Select Case to_unit
        Case "N-m"
            output = default_unit_value
        Case "mN-m"
            output = bg_ForceConvert(default_unit_value, "N", "mN")
        Case "in-lb"
            output = bg_LengthConvert(bg_ForceConvert(default_unit_value, "N", "lbf"), "m", "in")
        Case "ft-lb"
            output = bg_LengthConvert(bg_ForceConvert(default_unit_value, "N", "lbf"), "m", "ft")
        Case "in-oz"
            output = bg_LengthConvert(bg_ForceConvert(default_unit_value, "N", "ozf"), "m", "in")
    End Select
    
    bg_TorqueConvert = output
    
End Function

