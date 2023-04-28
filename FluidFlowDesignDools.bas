Attribute VB_Name = "Module1"
Public Function fluidMassFlowConvert(value As Double, from_unit As String, to_unit As String) As Double
'converts mass flow units between each other

'default unit is going to be kg/s
Dim default_unit_value As Double

'convert everything to the default unit
Select Case from_unit
Case "kg/s"
    default_unit_value = value
Case "kg/m", "kg/min"
    default_unit_value = value / 60
Case "kg/hr"
    default_unit_value = value / 3600
Case "g/s"
    default_unit_value = value / 1000
Case "g/m", "g/min"
    default_unit_value = value / 1000 / 60
Case "g/hr"
    default_unit_value = value / 1000 / 3600
Case "lbm/s"
    default_unit_value = value * 0.453592
Case "lbm/m", "lbm/min"
    default_unit_value = value * 0.453592 / 60
Case "lbm/hr"
    default_unit_value = value * 0.453592 / 3600
End Select

Dim new_unit_value As Double
'convert from the default unit to the output unit
Select Case to_unit
Case "kg/s"
    new_unit_value = default_unit_value
Case "kg/m", "kg/min"
    new_unit_value = default_unit_value * 60
Case "kg/hr"
    new_unit_value = default_unit_value * 3600
Case "g/s"
    new_unit_value = default_unit_value * 1000
Case "g/m", "g/min"
    new_unit_value = default_unit_value * 1000 * 60
Case "g/hr"
    new_unit_value = default_unit_value * 1000 * 3600
Case "lbm/s"
    new_unit_value = default_unit_value / 0.453592
Case "lbm/m", "lbm/min"
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
Case "m3/m", "m3/min"
    default_unit_value = value / 60
Case "m3/hr"
    default_unit_value = value / 3600
Case "cm3/s", "ccs"
    default_unit_value = value / 100 ^ 3
Case "cm3/m", "cm3/min", "ccm"
    default_unit_value = value / 100 ^ 3 / 60
Case "cm3/hr"
    default_unit_value = value / 100 ^ 3 / 3600
Case "L/s"
    default_unit_value = value / 1000
Case "L/m", "L/min"
    default_unit_value = value / 1000 / 60
Case "L/hr"
    default_unit_value = value / 1000 / 3600
Case "in3/s"
    default_unit_value = value * 0.000016387
Case "in3/m", "in3/min"
    default_unit_value = value * 0.000016387 / 60
Case "in3/hr"
    default_unit_value = value * 0.000016387 / 3600
Case "ft3/s"
    default_unit_value = value * 0.0283168
Case "ft3/m", "ft3/min"
    default_unit_value = value * 0.0283168 / 60
Case "ft3/hr"
    default_unit_value = value * 0.0283168 / 3600
Case "gal/s"
    default_unit_value = value / 264.172
Case "gal/m", "gal/min", "gpm"
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
Case "cm3/m", "cm3/min", "ccm"
    new_unit_value = default_unit_value * 100 ^ 3 * 60
Case "cm3/hr"
    new_unit_value = default_unit_value * 100 ^ 3 * 3600
Case "L/s"
    new_unit_value = default_unit_value * 1000
Case "L/m", "L/min"
    new_unit_value = default_unit_value * 1000 * 60
Case "L/hr"
    new_unit_value = default_unit_value * 1000 * 3600
Case "in3/s"
    new_unit_value = default_unit_value / 0.000016387
Case "in3/m", "in3/min"
    new_unit_value = default_unit_value / 0.000016387 * 60
Case "in3/hr"
    new_unit_value = default_unit_value / 0.000016387 * 3600
Case "ft3/s"
    new_unit_value = default_unit_value / 0.0283168
Case "ft3/m", "ft3/min"
    new_unit_value = default_unit_value / 0.0283168 * 60
Case "ft3/hr"
    new_unit_value = default_unit_value / 0.0283168 * 3600
Case "gal/s"
    new_unit_value = default_unit_value * 264.172
Case "gal/m", "gal/min", "gpm"
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

Select Case Gas
Case "air"
    fluidMolecularWeight = 28.96 / 1000
Case "O2"
    fluidMolecularWeight = 31.999 / 1000
Case "N2"
    fluidMolecularWeight = 28.0134 / 1000
Case "He"
    fluidMolecularWeight = 4.002602 / 1000
End Select

End Function


Public Function fluidStandardConditions(Standard As String, Output As String, Optional Metric As Boolean = True) As Double
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
End Select

If Not Metric Then
    Temperature = Temperature * 9 / 5 'Convert to Rankine
    Pressure = Pressure * 0.000145038 'Convert to psia
End If

Select Case Output
Case "T"
    fluidStandardConditions = Temperature
Case "P"
    fluidStandardConditions = Pressure
End Select

End Function

Public Function fluidUniversalGasConstant() As Double
'Outputs the universal gas constant R in J/(K*mol)
fluidUniversalGasConstant = 8.3144598
End Function
