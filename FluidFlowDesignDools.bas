Attribute VB_Name = "FluidUnits"
Option Explicit

Public Function fluidUnitType(unit As String) As String

Select Case unit
Case "kg/s", "kg/min", "kg/hr", "g/s", "g/min", "g/hr", "lbm/s", "lbm/min", "lbm/hr"
    fluidUnitType = "Mass Flow"
Case "ccs", "m3/s", "m3/min", "m3/hr", "cm3/s", "cm3/min", "cm3/hr", "L/s", "lps", "L/min", "lpm", "L/hr", "in3/s", "in3/min", "in3/hr", "ft3/s", "ft3/min", "ft3/hr", "gal/s", "gps", "gal/min", "gpm", "gal/hr"
    fluidUnitType = "Volume Flow"
Case "kg/m3", "kg/cm3", "kg/L", "g/m3", "g/cm3", "g/L", "lbm/in3", "lbm/ft3", "lbm/gal"
    fluidUnitType = "Density"
Case "scms", "scmm", "scmh", "slps", "slpm", "slph", "sccs", "sccm", "scch", "scfs", "scfm", "scfh", "scis", "scim", "scih"
    fluidUnitType = "Standard Volume Flow"
Case "s", "min", "hr"
    fluidUnitType = "Time"
Case "km", "m", "cm", "mm", "in", "ft", "yd", "mi"
    fluidUnitType = "Length"
Case "km/s", "m/s", "cm/s", "mm/s", "in/s", "ft/s", "yd/s", "mi/s", "km/min", "m/min", "cm/min", "mm/min", "in/min", "ft/min", "yd/min", "mi/min", "km/hr", "m/hr", "cm/hr", "mm/hr", "in/hr", "ft/hr", "yd/hr", "mi/hr", "mph"
    fluidUnitType = "Speed"
Case "F", "°F", "C", "°C", "K", "R", "Rank", "Rankine"
    fluidUnitType = "Temperature"
Case "psia", "psig", "psid", "psi", "ksi", "lbf/in2", "lb/in2", "mmHg", "atm", "Pa", "kPa", "MPa", "GPa", "torr", "Torr", "bar"
    fluidUnitType = "Pressure"
Case Else
    fluidUnitType = "Unit Error, " & unit & " not recognized"
End Select

End Function

Public Function fluidMassFlowConvert(value As Double, from_unit As String, to_unit As String) As Double
'converts mass flow units between each other

'default unit is going to be kg/s
Dim default_unit_value As Double

'convert everything to the default unit
Select Case from_unit
Case "kg/s", "kg/sec"
'Case InStr(KGPS, from_unit) > 0
    default_unit_value = value
Case "kg/min"
    default_unit_value = value / 60
Case "kg/hr"
    default_unit_value = value / 3600
Case "g/s", "g/sec"
    default_unit_value = value / 1000
Case "g/min"
    default_unit_value = value / 1000 / 60
Case "g/hr"
    default_unit_value = value / 1000 / 3600
Case "lbm/s", "lbm/sec"
    default_unit_value = value * 0.453592
Case "lbm/min"
    default_unit_value = value * 0.453592 / 60
Case "lbm/hr"
    default_unit_value = value * 0.453592 / 3600
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

fluidMassFlowConvert = new_unit_value

End Function

Public Function fluidVolumeFlowConvert(value As Double, from_unit As String, to_unit As String) As Double
'converts volume flow units between each other

'default unit is going to be m3/s
Dim default_unit_value As Double

'TO-DO - finish changing all these to the right volume conversion factors
'convert everything to the default unit
Select Case from_unit
Case "m3/s"
    default_unit_value = value
Case "m3/min"
    default_unit_value = value / 60
Case "m3/hr"
    default_unit_value = value / 3600
Case "cm3/s", "ccs"
    default_unit_value = value / 100 ^ 3
Case "cm3/min", "ccm"
    default_unit_value = value / 100 ^ 3 / 60
Case "cm3/hr"
    default_unit_value = value / 100 ^ 3 / 3600
Case "L/s", "lps"
    default_unit_value = value / 1000
Case "L/min", "lpm"
    default_unit_value = value / 1000 / 60
Case "L/hr"
    default_unit_value = value / 1000 / 3600
Case "in3/s"
    default_unit_value = value * 0.000016387
Case "in3/min"
    default_unit_value = value * 0.000016387 / 60
Case "in3/hr"
    default_unit_value = value * 0.000016387 / 3600
Case "ft3/s"
    default_unit_value = value * 0.0283168
Case "ft3/min"
    default_unit_value = value * 0.0283168 / 60
Case "ft3/hr"
    default_unit_value = value * 0.0283168 / 3600
Case "gal/s", "gps"
    default_unit_value = value / 264.172
Case "gal/min", "gpm"
    default_unit_value = value / 264.172 / 60
Case "gal/hr"
    default_unit_value = value / 264.172 / 3600
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

fluidVolumeFlowConvert = new_unit_value

End Function

Public Function fluidDensityConvert(value As Double, from_unit As String, to_unit As String) As Double
'converts mass flow units between each other

'default unit is going to be kg/m3
Dim default_unit_value As Double

'convert everything to the default unit
Select Case from_unit
Case "kg/m3"
    default_unit_value = value
Case "kg/cm3"
    default_unit_value = value * 100 ^ 3
Case "kg/L"
    default_unit_value = value * 10 ^ 3
Case "g/m3"
    default_unit_value = value / 1000
Case "g/cm3"
    default_unit_value = value * 100 ^ 3 / 1000
Case "g/L"
    default_unit_value = value * 10 ^ 3 / 1000
Case "lbm/in3"
    default_unit_value = value * 27679.9
Case "lbm/ft3"
    default_unit_value = value * 16.018
Case "lbm/gal"
    default_unit_value = value * 119.826
End Select

Dim new_unit_value As Double
Select Case to_unit
Case "kg/m3"
    new_unit_value = value
Case "kg/cm3"
    new_unit_value = value / 100 ^ 3
Case "kg/L"
    new_unit_value = value / 10 ^ 3
Case "g/m3"
    new_unit_value = value * 1000
Case "g/cm3"
    new_unit_value = value / 100 ^ 3 * 1000
Case "g/L"
    new_unit_value = value / 10 ^ 3 * 1000
Case "lbm/in3"
    new_unit_value = value / 27679.9
Case "lbm/ft3"
    new_unit_value = value / 16.018
Case "lbm/gal"
    new_unit_value = value / 119.826
End Select

fluidDensityConvert = new_unit_value

End Function


Public Function fluidMolecularWeight(Gas As String) As Double
'Returns molecular weight in kg/mol for a given gas.
'Each number is divided by 1000 to convert it from g/mol to kg/mol
fluidMolecularWeight = Props1SI("M", Gas)
'Select Case Gas
'Case "air", "Air"
'    fluidMolecularWeight = 28.96 / 1000
'Case "O2", "Oxygen", "oxygen"
'    fluidMolecularWeight = 31.999 / 1000
'Case "N2", "Nitrogen", "nitrogen"
'    fluidMolecularWeight = 28.0134 / 1000
'Case "He", "Helium", "helium"
'    fluidMolecularWeight = 4.002602 / 1000
'End Select

End Function


Public Function fluidStandardConditions(Standard As String, output As String, Optional Metric As Boolean = True) As Double
'Outputs pressure or temperature of a standard condition set in either metric or customary units
Dim Temperature As Double 'in Kelvin by default
Dim Pressure As Double 'in Pascals by default

Select Case Standard
Case "IUPAC_STP"
    Temperature = 273.15
    Pressure = 100000
Case "Pre1982IUPAC_STP"
    Temperature = 273.15
    Pressure = 101325
Case "NTP"
    Temperature = 293.15
    Pressure = 100325
Case "IUPAC_SATP"
    Temperature = 298.15
    Pressure = 100000
Case "EPA"
    Temperature = 298.15
    Pressure = 101325
Case "ISO 2533", "ISO 13443", "ISO 7504"
    Temperature = 288.15
    Pressure = 101325
End Select

If Not Metric Then
    Temperature = Temperature * 9 / 5 'Convert to Rankine
    Pressure = Pressure * 0.000145038 'Convert to psia
End If

Select Case output
Case "T"
    fluidStandardConditions = Temperature
Case "P"
    fluidStandardConditions = Pressure
End Select

End Function

Public Function fluidUniversalGasConstant() As Double
'Outputs the universal gas constant R in J/(K*mol), aka m3-Pa/(K-mol)
fluidUniversalGasConstant = 8.3144598
End Function

Public Function fluidGasConstant(gas_name As String) As Double
'Outputs the specific gas constant of the specified fluid
Dim R_uni As Double
R_uni = fluidUniversalGasConstant()
fluidGasConstant = R_uni / fluidMolecularWeight(gas_name)
End Function

Public Function fluidCompressibility(Optional P_init_Pa As Double = 0, Optional T_init_K As Double, Optional gas_name As String) As Double
    'temporary function for outputting Z for use in standard fluid flow calcs
    If P_init_Pa = 0 Or T_init_K = 0 Then
        fluidCompressibility = 1
    Else
        fluidCompressibility = PropsSI("Z", "P", P_init_Pa, "T", T_init_K, gas_name)
    End If
End Function

Public Function fluidStandardDensity(gas_name As String, Standard As String, Optional unit As String = "kg/m3", Optional P_init_Pa As Double, Optional T_init_K As Double) As Double
    'Output the given standard and gas density in the given unit
    Dim std_P As Double, std_T As Double, R_gas As Double, Z As Double
    std_P = fluidStandardConditions(Standard, "P", True) 'Pa
    std_T = fluidStandardConditions(Standard, "T", True) 'K
    R_gas = fluidGasConstant(gas_name)
    Z = fluidCompressibility(P_init_Pa, T_init_K, gas_name) 'unitless
    
    'check if the unit is actually a density unit
    If fluidUnitType(unit) = "Density" Then
        fluidStandardDensity = fluidDensityConvert(std_P / (Z * R_gas * std_T), "kg/m3", unit) ' kg/m3
    End If
End Function

Public Function fluidMassFlowToStandardFlow(value As Double, from_unit As String, to_unit As String, gas_name As String, Optional Standard As String = "IUPAC_STP", Optional P_init_Pa As Double = 0, Optional T_init_K As Double = 0) As Double
'Convert whatever mass flow unit to standard volumetric flow based on some standard
'Default mass flow unit is kg/s
Dim default_unit_value As Double, std_Density As Double, default_std_unit_value As Double, default_std_unit As String
default_unit_value = fluidMassFlowConvert(value, from_unit, "kg/s")

std_Density = fluidStandardDensity(gas_name, Standard, "kg/m3", P_init_Pa, T_init_K) ' kg/m3
default_std_unit = "m3/s"
default_std_unit_value = default_unit_value / std_Density 'standard m3/s

Dim new_unit_value As Double
'convert from the default unit to the output unit
Select Case to_unit
Case "scms" 'Standard Cubic Meter per Second
    new_unit_value = fluidVolumeFlowConvert(default_std_unit_value, default_std_unit, "m3/s")
Case "scmm" 'Standard Cubic Meter per Minute
    new_unit_value = fluidVolumeFlowConvert(default_std_unit_value, default_std_unit, "m3/min")
Case "scmh" 'Standard Cubic Meter per Hour
    new_unit_value = fluidVolumeFlowConvert(default_std_unit_value, default_std_unit, "m3/hr")
Case "slps", "sls" 'Standard Liter per Second
    new_unit_value = fluidVolumeFlowConvert(default_std_unit_value, default_std_unit, "L/s")
Case "slpm", "slm" 'Standard Liter per Minute
    new_unit_value = fluidVolumeFlowConvert(default_std_unit_value, default_std_unit, "L/min")
Case "slph", "slh" 'Standard Liter per Hour
    new_unit_value = fluidVolumeFlowConvert(default_std_unit_value, default_std_unit, "L/hr")
Case "sccs" 'Standard Cubic Centimeter per Second
    new_unit_value = fluidVolumeFlowConvert(default_std_unit_value, default_std_unit, "cm3/s")
Case "sccm" 'Standard Cubic Centimeter per Minute
    new_unit_value = fluidVolumeFlowConvert(default_std_unit_value, default_std_unit, "cm3/min")
Case "scch" 'Standard Cubic Centimeter per Hour
    new_unit_value = fluidVolumeFlowConvert(default_std_unit_value, default_std_unit, "cm3/hr")
Case "scfs" 'Standard Cubic Feet per Second
    new_unit_value = fluidVolumeFlowConvert(default_std_unit_value, default_std_unit, "ft3/s")
Case "scfm" 'Standard Cubic Feet per Minute
    new_unit_value = fluidVolumeFlowConvert(default_std_unit_value, default_std_unit, "ft3/min")
Case "scfh" 'Standard Cubic Feet per Hour
    new_unit_value = fluidVolumeFlowConvert(default_std_unit_value, default_std_unit, "in3/hr")
Case "scis" 'Standard Cubic Inches per Second
    new_unit_value = fluidVolumeFlowConvert(default_std_unit_value, default_std_unit, "in3/s")
Case "scim" 'Standard Cubic Inches per Minute
    new_unit_value = fluidVolumeFlowConvert(default_std_unit_value, default_std_unit, "in3/min")
Case "scih" 'Standard Cubic Inches per Hour
    new_unit_value = fluidVolumeFlowConvert(default_std_unit_value, default_std_unit, "in3/hr")
End Select

fluidMassFlowToStandardFlow = new_unit_value

End Function

Public Function fluidStandardFlowConvert(value As Double, from_unit As String, to_unit As String) As Double
Dim default_std_unit As String, new_unit_value As Double, default_std_unit_value As Double
default_std_unit = "m3/s"


Select Case from_unit
Case "scms" 'Standard Cubic Meter per Second
    default_std_unit_value = fluidVolumeFlowConvert(value, "m3/s", default_std_unit)
Case "scmm" 'Standard Cubic Meter per Minute
    default_std_unit_value = fluidVolumeFlowConvert(value, "m3/min", default_std_unit)
Case "scmh" 'Standard Cubic Meter per Hour
    default_std_unit_value = fluidVolumeFlowConvert(value, "m3/hr", default_std_unit)
Case "slps", "sls" 'Standard Liter per Second
    default_std_unit_value = fluidVolumeFlowConvert(value, "L/s", default_std_unit)
Case "slpm", "slm" 'Standard Liter per Minute
    default_std_unit_value = fluidVolumeFlowConvert(value, "L/min", default_std_unit)
Case "slph", "slh" 'Standard Liter per Hour
    default_std_unit_value = fluidVolumeFlowConvert(value, "L/hr", default_std_unit)
Case "sccs" 'Standard Cubic Centimeter per Second
    default_std_unit_value = fluidVolumeFlowConvert(value, "cm3/s", default_std_unit)
Case "sccm" 'Standard Cubic Centimeter per Minute
    default_std_unit_value = fluidVolumeFlowConvert(value, "cm3/min", default_std_unit)
Case "scch" 'Standard Cubic Centimeter per Hour
    default_std_unit_value = fluidVolumeFlowConvert(value, "cm3/hr", default_std_unit)
Case "scfs" 'Standard Cubic Feet per Second
    default_std_unit_value = fluidVolumeFlowConvert(value, "ft3/s", default_std_unit)
Case "scfm" 'Standard Cubic Feet per Minute
    default_std_unit_value = fluidVolumeFlowConvert(value, "ft3/min", default_std_unit)
Case "scfh" 'Standard Cubic Feet per Hour
    default_std_unit_value = fluidVolumeFlowConvert(value, "in3/hr", default_std_unit)
Case "scis" 'Standard Cubic Inches per Second
    default_std_unit_value = fluidVolumeFlowConvert(value, "in3/s", default_std_unit)
Case "scim" 'Standard Cubic Inches per Minute
    default_std_unit_value = fluidVolumeFlowConvert(value, "in3/min", default_std_unit)
Case "scih" 'Standard Cubic Inches per Hour
    default_std_unit_value = fluidVolumeFlowConvert(value, "in3/hr", default_std_unit)
End Select

Select Case to_unit
Case "scms" 'Standard Cubic Meter per Second
    new_unit_value = fluidVolumeFlowConvert(default_std_unit_value, default_std_unit, "m3/s")
Case "scmm" 'Standard Cubic Meter per Minute
    new_unit_value = fluidVolumeFlowConvert(default_std_unit_value, default_std_unit, "m3/min")
Case "scmh" 'Standard Cubic Meter per Hour
    new_unit_value = fluidVolumeFlowConvert(default_std_unit_value, default_std_unit, "m3/hr")
Case "slps", "sls" 'Standard Liter per Second
    new_unit_value = fluidVolumeFlowConvert(default_std_unit_value, default_std_unit, "L/s")
Case "slpm", "slm" 'Standard Liter per Minute
    new_unit_value = fluidVolumeFlowConvert(default_std_unit_value, default_std_unit, "L/min")
Case "slph", "slh" 'Standard Liter per Hour
    new_unit_value = fluidVolumeFlowConvert(default_std_unit_value, default_std_unit, "L/hr")
Case "sccs" 'Standard Cubic Centimeter per Second
    new_unit_value = fluidVolumeFlowConvert(default_std_unit_value, default_std_unit, "cm3/s")
Case "sccm" 'Standard Cubic Centimeter per Minute
    new_unit_value = fluidVolumeFlowConvert(default_std_unit_value, default_std_unit, "cm3/min")
Case "scch" 'Standard Cubic Centimeter per Hour
    new_unit_value = fluidVolumeFlowConvert(default_std_unit_value, default_std_unit, "cm3/hr")
Case "scfs" 'Standard Cubic Feet per Second
    new_unit_value = fluidVolumeFlowConvert(default_std_unit_value, default_std_unit, "ft3/s")
Case "scfm" 'Standard Cubic Feet per Minute
    new_unit_value = fluidVolumeFlowConvert(default_std_unit_value, default_std_unit, "ft3/min")
Case "scfh" 'Standard Cubic Feet per Hour
    new_unit_value = fluidVolumeFlowConvert(default_std_unit_value, default_std_unit, "in3/hr")
Case "scis" 'Standard Cubic Inches per Second
    new_unit_value = fluidVolumeFlowConvert(default_std_unit_value, default_std_unit, "in3/s")
Case "scim" 'Standard Cubic Inches per Minute
    new_unit_value = fluidVolumeFlowConvert(default_std_unit_value, default_std_unit, "in3/min")
Case "scih" 'Standard Cubic Inches per Hour
    new_unit_value = fluidVolumeFlowConvert(default_std_unit_value, default_std_unit, "in3/hr")
End Select
fluidStandardFlowConvert = new_unit_value

End Function

Public Function fluidStandardFlowToMassFlow(value As Double, from_unit As String, to_unit As String, gas_name As String, Optional Standard As String = "IUPAC_STP", Optional P_init_Pa As Double = 0, Optional T_init_K As Double = 0) As Double
'Output the equivalent mass flow of a standard and gas volumetric flow in a given unit
Dim default_unit_value As Double, std_Density As Double, default_std_unit_value As Double, default_std_unit As String
default_std_unit = "scms"

default_std_unit_value = fluidStandardFlowConvert(value, from_unit, default_std_unit)

std_Density = fluidStandardDensity(gas_name, Standard, "kg/m3", P_init_Pa, T_init_K) 'kg/m3
default_unit_value = default_std_unit_value * std_Density
fluidStandardFlowToMassFlow = fluidMassFlowConvert(default_unit_value, "kg/s", to_unit)

End Function

Public Function fluidLengthConvert(value As Double, from_unit As String, to_unit As String) As Double
    'converts between length units using excel's convert function
    'TO-DO protect for length units
    fluidLengthConvert = Application.WorksheetFunction.Convert(value, from_unit, to_unit)
End Function

Public Function fluidTimeConvert(value As Double, from_unit As String, to_unit As String) As Double
    'converts between time units using excel's convert function
    'TO-DO protect for time units
    fluidTimeConvert = Application.WorksheetFunction.Convert(value, from_unit, to_unit)
End Function

Public Function fluidSpeedConvert(value As Double, from_unit As String, to_unit As String) As Double
    'convert whatever the from_unit is to /s, then convert it to the to_unit using fluidLengthConvert
    
    'default value is m/s
    Dim default_time_unit_value As Double
    Dim default_unit_value As Double
    
    Dim to_time_unit_value As Double
    Dim to_unit_length_unit As String
    Dim from_unit_length_unit As String
    
    'convert time unit to /s
    Select Case from_unit
    Case "km/s", "m/s", "cm/s", "mm/s", "in/s", "ft/s", "yd/s", "mi/s"
        default_time_unit_value = value
    Case "km/min", "m/min", "cm/min", "mm/min", "in/min", "ft/min", "yd/min", "mi/min"
        default_time_unit_value = value * 60
    Case "km/hr", "m/hr", "cm/hr", "mm/hr", "in/hr", "ft/hr", "yd/hr", "mi/hr", "mph"
        default_time_unit_value = value * 3600
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
    fluidSpeedConvert = fluidLengthConvert(to_time_unit_value, from_unit_length_unit, to_unit_length_unit)
End Function

Public Function fluidTemperatureConvert(value As Double, from_unit As String, to_unit As String) As Double
    'figures out the temperature unit to input into excel's default convert function
    
    'TO-DO: protect for temperature units
    Dim from_unit_unit As String, to_unit_unit As String
    
    
    'figure out from unit's unit
    Select Case from_unit
        Case "F", "°F"
            from_unit_unit = "F"
        Case "C", "°C"
            from_unit_unit = "C"
        Case "K"
            from_unit_unit = "K"
        Case "R", "Rank", "Rankine"
            from_unit_unit = "Rank"
    End Select
    
    'figure out to unit's unit
    Select Case to_unit
        Case "F", "°F"
            to_unit_unit = "F"
        Case "C", "°C"
            to_unit_unit = "C"
        Case "K"
            to_unit_unit = "K"
        Case "R", "Rank", "Rankine"
            to_unit_unit = "Rank"
    End Select
    
    'Final convert
    fluidTemperatureConvert = Application.WorksheetFunction.Convert(value, from_unit_unit, to_unit_unit)

End Function

Public Function fluidPressureConvert(value As Double, from_unit As String, to_unit As String) As Double
    'figures out the pressure unit and converts it either with excel's default convert function or otherwise
    'TO-DO protect for pressure units
    
    Dim from_unit_unit As String, deciphered_value As Double
    Dim to_unit_unit As String
    'decipher weird units and convert from_unit
    Select Case from_unit
        Case "psia", "psid", "psi", "lbf/in2", "lb/in2"
            from_unit_unit = "psi"
            deciphered_value = value
        Case "ksi"
            deciphered_value = value * 1000
            from_unit_unit = "psi"
        Case "Pa"
            deciphered_value = value
            from_unit_unit = "Pa"
        Case "kPa"
            deciphered_value = value * 1000
            from_unit_unit = "Pa"
        Case "MPa"
            deciphered_value = value * 1000000
            from_unit_unit = "Pa"
        Case "GPa"
            deciphered_value = value * 1000000000
            from_unit_unit = "Pa"
        Case "bar"
            deciphered_value = value * 100000
            from_unit_unit = "Pa"
        Case "Torr", "torr"
            deciphered_value = value
            from_unit_unit = "Torr"
        Case Else
            from_unit_unit = from_unit
            deciphered_value = value
    End Select
    
    Select Case to_unit
        Case "psia", "psid", "psi", "lbf/in2", "lb/in2"
            to_unit_unit = "psi"
            fluidPressureConvert = Application.WorksheetFunction.Convert(deciphered_value, from_unit_unit, to_unit_unit)
        Case "ksi"
            to_unit_unit = "psi"
            fluidPressureConvert = Application.WorksheetFunction.Convert(deciphered_value, from_unit_unit, to_unit_unit) / 1000
        Case "Pa"
            to_unit_unit = "Pa"
            fluidPressureConvert = Application.WorksheetFunction.Convert(deciphered_value, from_unit_unit, to_unit_unit)
        Case "kPa"
            to_unit_unit = "Pa"
            fluidPressureConvert = Application.WorksheetFunction.Convert(deciphered_value, from_unit_unit, to_unit_unit) / 1000
        Case "MPa"
            to_unit_unit = "Pa"
            fluidPressureConvert = Application.WorksheetFunction.Convert(deciphered_value, from_unit_unit, to_unit_unit) / 1000000
        Case "GPa"
            to_unit_unit = "Pa"
            fluidPressureConvert = Application.WorksheetFunction.Convert(deciphered_value, from_unit_unit, to_unit_unit) / 1000000000
        Case "bar"
            to_unit_unit = "Pa"
            fluidPressureConvert = Application.WorksheetFunction.Convert(deciphered_value, from_unit_unit, to_unit_unit) / 100000
        Case "Torr", "torr"
            to_unit_unit = "Torr"
            fluidPressureConvert = Application.WorksheetFunction.Convert(deciphered_value, from_unit_unit, to_unit_unit)
        Case Else
            to_unit_unit = to_unit
            fluidPressureConvert = Application.WorksheetFunction.Convert(deciphered_value, from_unit_unit, to_unit_unit)
    End Select
    
End Function

Public Function fluidConvert(value As Double, from_unit As String, to_unit As String, Optional gas_name As String = "None", Optional Standard As String = "IUPAC_STP", Optional P_init_Pa As Double = 0, Optional T_init_K As Double = 0)
'Combines all the previous function capabilities into a general convert function
Dim from_unit_type As String, to_unit_type As String

from_unit_type = fluidUnitType(from_unit)
to_unit_type = fluidUnitType(to_unit)

Select Case from_unit_type
Case "Mass Flow"
    Select Case to_unit_type
        Case "Mass Flow"
            fluidConvert = fluidMassFlowConvert(value, from_unit, to_unit)
        Case "Standard Volume Flow"
            fluidConvert = fluidMassFlowToStandardFlow(value, from_unit, to_unit, gas_name, Standard, P_init_Pa, T_init_K)
    End Select
Case "Volume Flow"
    Select Case to_unit_type
        Case "Volume Flow"
            fluidConvert = fluidVolumeFlowConvert(value, from_unit, to_unit)
    End Select
Case "Density"
    Select Case to_unit_type
        Case "Density"
            fluidConvert = fluidDensityConvert(value, from_unit, to_unit)
    End Select
Case "Standard Volume Flow"
    Select Case to_unit_type
        Case "Standard Volume Flow"
            fluidConvert = fluidStandardFlowConvert(value, from_unit, to_unit)
        Case "Mass Flow"
            fluidConvert = fluidStandardFlowToMassFlow(value, from_unit, to_unit, gas_name, Standard, P_init_Pa, T_init_K)
    End Select
Case "Time"
    Select Case to_unit_type
        Case "Time"
            fluidConvert = fluidTimeConvert(value, from_unit, to_unit)
    End Select
Case "Length"
    Select Case to_unit_type
        Case "Length"
            fluidConvert = fluidLengthConvert(value, from_unit, to_unit)
    End Select
Case "Speed"
    Select Case to_unit_type
        Case "Speed"
            fluidConvert = fluidSpeedConvert(value, from_unit, to_unit)
    End Select
Case "Temperature"
    Select Case to_unit_type
        Case "Temperature"
            fluidConvert = fluidTemperatureConvert(value, from_unit, to_unit)
    End Select
Case "Pressure"
    Select Case to_unit_type
        Case "Pressure"
            fluidConvert = fluidPressureConvert(value, from_unit, to_unit)
    End Select
End Select

If fluidConvert = Empty Then
    fluidConvert = "Cannot convert " & from_unit_type & " to " & to_unit_type
End If

End Function


