Attribute VB_Name = "FluidUnits"
Option Explicit

Public Function fluidMolecularWeight(Gas As String) As Double
    'Returns molecular weight in kg/mol for a given gas.
    fluidMolecularWeight = Props1SI("M", Gas)
    
    'Uncomment below code to be able to use this function without Coolprop
    'Each number is divided by 1000 to convert it from g/mol to kg/mol
    
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

Public Function bg_StandardConditions(Standard As String, output As String, Optional Metric As Boolean = True) As Double
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
        bg_StandardConditions = Temperature
    Case "P"
        bg_StandardConditions = Pressure
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
    std_P = bg_StandardConditions(Standard, "P", True) 'Pa
    std_T = bg_StandardConditions(Standard, "T", True) 'K
    R_gas = fluidGasConstant(gas_name)
    Z = fluidCompressibility(P_init_Pa, T_init_K, gas_name) 'unitless
    
    'check if the unit is actually a density unit
    If bg_UnitType(unit) = "Density" Then
        fluidStandardDensity = bg_DensityConvert(std_P / (Z * R_gas * std_T), "kg/m3", unit) ' kg/m3
    End If
End Function



Public Function fluidConvert(Value As Double, from_unit As String, to_unit As String, Optional gas_name As String = "None", Optional Standard As String = "IUPAC_STP", Optional P_init_Pa As Double = 0, Optional T_init_K As Double = 0)
    'Combines all the previous function capabilities into a general convert function
    Dim from_unit_type As String, to_unit_type As String
    
    from_unit_type = bg_UnitType(from_unit)
    to_unit_type = bg_UnitType(to_unit)
    
    Select Case from_unit_type
    Case "Mass Flow"
        Select Case to_unit_type
            Case "Mass Flow"
                fluidConvert = bg_MassFlowConvert(Value, from_unit, to_unit)
            Case "Standard Volume Flow"
                fluidConvert = bg_MassFlowToStandardFlow(Value, from_unit, to_unit, gas_name, Standard, P_init_Pa, T_init_K)
        End Select
    Case "Volume Flow"
        Select Case to_unit_type
            Case "Volume Flow"
                fluidConvert = bg_VolumeFlowConvert(Value, from_unit, to_unit)
        End Select
    Case "Density"
        Select Case to_unit_type
            Case "Density"
                fluidConvert = bg_DensityConvert(Value, from_unit, to_unit)
        End Select
    Case "Standard Volume Flow"
        Select Case to_unit_type
            Case "Standard Volume Flow"
                fluidConvert = bg_StandardFlowConvert(Value, from_unit, to_unit)
            Case "Mass Flow"
                fluidConvert = bg_StandardFlowToMassFlow(Value, from_unit, to_unit, gas_name, Standard, P_init_Pa, T_init_K)
        End Select
    Case "Time"
        Select Case to_unit_type
            Case "Time"
                fluidConvert = bg_TimeConvert(Value, from_unit, to_unit)
        End Select
    Case "Length"
        Select Case to_unit_type
            Case "Length"
                fluidConvert = bg_LengthConvert(Value, from_unit, to_unit)
        End Select
    Case "Speed"
        Select Case to_unit_type
            Case "Speed"
                fluidConvert = bg_SpeedConvert(Value, from_unit, to_unit)
        End Select
    Case "Temperature"
        Select Case to_unit_type
            Case "Temperature"
                fluidConvert = bg_TemperatureConvert(Value, from_unit, to_unit)
        End Select
    Case "Pressure"
        Select Case to_unit_type
            Case "Pressure"
                fluidConvert = bg_PressureConvert(Value, from_unit, to_unit)
        End Select
    Case "Area"
        Select Case to_unit_type
            Case "Area"
                fluidConvert = bg_AreaConvert(Value, from_unit, to_unit)
        End Select
    Case "Volume"
        Select Case to_unit_type
            Case "Volume"
                fluidConvert = bg_VolumeConvert(Value, from_unit, to_unit)
        End Select
    Case "Moment of Inertia"
        Select Case to_unit_type
            Case "Moment of Inertia"
                fluidConvert = bg_MomentOfInertiaConvert(Value, from_unit, to_unit)
        End Select
    Case "Mass"
        Select Case to_unit_type
            Case "Mass"
                fluidConvert = bg_MassConvert(Value, from_unit, to_unit)
        End Select
    Case "Force"
        Select Case to_unit_type
            Case "Force"
                fluidConvert = bg_ForceConvert(Value, from_unit, to_unit)
        End Select
    Case "Torque"
        Select Case to_unit_type
            Case "Torque"
                fluidConvert = bg_TorqueConvert(Value, from_unit, to_unit)
        End Select
    End Select
    
    If fluidConvert = Empty Then
        fluidConvert = "Cannot convert " & from_unit_type & " to " & to_unit_type
    End If

End Function

Public Function fluidRatioOfSpecificHeats(Gas As String, Pressure_Pa As Double, Temperature_K As Double) As Double
    'Calculates the Ratio of Specific Heats of the given gas using pressure in Pa and temperature in K using CoolProp
    Dim Cp As Double, Cv As Double
    
    Cp = PropsSI("CPMASS", "P", Pressure_Pa, "T", Temperature_K, Gas)
    Cv = PropsSI("CVMASS", "P", Pressure_Pa, "T", Temperature_K, Gas)
    fluidRatioOfSpecificHeats = Cp / Cv

End Function

Public Function fluidDensity(Gas As String, Pressure_Pa As Double, Temperature_K As Double) As Double
    'Calculates the Density of the given gas using pressure in Pa and temperature in K using CoolProp
    
    fluidDensity = PropsSI("D", "P", Pressure_Pa, "T", Temperature_K, Gas)

End Function


Public Function fluidCircleArea(Diameter As Double) As Double
    'Calculates the area of a circle given its diameter
    fluidCircleArea = Application.WorksheetFunction.Pi() / 4 * Diameter ^ 2
End Function

Public Function fluidCdA(Cd As Double, Area As Double, P1 As Double, P2 As Double, T1 As Double, Gas As String, Optional output As String = "Mdot") As Variant
    'Calculates mdot (kg/s) of an orifice given upstream pressure and temperature P1 (Pa) & T1 (K), downstream pressure P2 (Pa), Discharge coefficient Cd, Area A (m2), and the gas name as a string
    'Can also output other things besides Mdot if the optional output is given.  Orifice Flow Speed (m/s), Choked Check (string)
    
    'Define easier constants
    Dim P2dP1 As Double
    
    
    'Define Gas property constants.  G1 = Gamma = Ratio of Specific Heats, D1 = Upstream Density, P2dP1_choked = Choked Flow Pressure Ratio
    Dim G1 As Double, D1 As Double, P2dP1_choked As Double
    
    'Define common gamma constants here
    Dim C1 As Double, C2 As Double, C3 As Double, C4 As Double, C5 As Double
    
    G1 = fluidRatioOfSpecificHeats(Gas, P1, T1)
    D1 = fluidDensity(Gas, P1, T1)
    
    C1 = 2 / (G1 + 1)
    C2 = G1 / (G1 - 1)
    C3 = 2 / G1
    C4 = (G1 + 1) / G1
    C5 = (G1 + 1) / (G1 - 1)
    
    P2dP1_choked = C1 ^ C2
    
    P2dP1 = P2 / P1
    
    
    Select Case output
        Case "Mdot"
            If P2dP1 > P2dP1_choked Then
                'Unchoked
                fluidCdA = Cd * Area * (2 * D1 * P1 * C2 * (P2dP1 ^ C3 - P2dP1 ^ C4)) ^ 0.5
            Else
                'choked
                fluidCdA = Cd * Area * (G1 * D1 * P1 * C1 ^ C5) ^ 0.5
            End If
        Case "Orifice Flow Speed"
            If P2dP1 > P2dP1_choked Then
                'Unchoked
                fluidCdA = (Cd * Area * (2 * D1 * P1 * C2 * (P2dP1 ^ C3 - P2dP1 ^ C4)) ^ 0.5) / (D1 * Area)
            Else
                'choked
                fluidCdA = (Cd * Area * (G1 * D1 * P1 * C1 ^ C5) ^ 0.5) / (D1 * Area)
            End If
        Case "Choked Check"
            If P2dP1 > P2dP1_choked Then
                'Unchoked
                fluidCdA = "Unchoked"
            Else
                'choked
                fluidCdA = "Choked"
            End If
    End Select

End Function


