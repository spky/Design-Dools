Attribute VB_Name = "GDT"
Option Explicit

Public Function GDT_Position(Nominal_Dimension As Double, Plus_Tolerance As Double, Minus_Tolerance As Double, Positional_Tolerance As Double, Dimension_Type As String, Output As String) As Double

    Dim MMC As Double, LMC As Double, VC As Double, RC As Double, Bonus As Double
    
    Bonus = Plus_Tolerance + Minus_Tolerance
    
    'Determine MMC and LMC
    Select Case Dimension_Type
        Case "RFS Hole", "MMC Hole", "LMC Hole", "RFS Slot", "MMC Slot", "LMC Slot"
            MMC = Nominal_Dimension - Minus_Tolerance
            LMC = Nominal_Dimension + Plus_Tolerance
        Case "RFS Shaft", "MMC Shaft", "LMC Shaft", "RFS Width", "MMC Width", "LMC Width"
            LMC = Nominal_Dimension - Minus_Tolerance
            MMC = Nominal_Dimension + Plus_Tolerance
    End Select
    
    'Determine VC and RC
    Select Case Dimension_Type
        Case "RFS Hole", "RFS Slot"
            VC = MMC - Positional_Tolerance
            RC = LMC + Positional_Tolerance
        Case "MMC Hole", "MMC Slot"
            VC = MMC - Positional_Tolerance
            RC = LMC + Positional_Tolerance + Bonus
        Case "LMC Hole", "LMC Slot"
            VC = LMC + Positional_Tolerance
            RC = MMC - Positional_Tolerance - Bonus
        Case "RFS Shaft", "RFS Width"
            VC = MMC + Positional_Tolerance
            RC = LMC - Positional_Tolerance
        Case "MMC Shaft", "MMC Width"
            VC = MMC + Positional_Tolerance
            RC = LMC - Positional_Tolerance - Bonus
        Case "LMC Shaft", "LMC Width"
            VC = LMC - Positional_Tolerance
            RC = MMC + Positional_Tolerance + Bonus
    End Select
    
    'Return the requested output
    Select Case Output
        Case "MMC"
            GDT_Position = MMC
        Case "LMC"
            GDT_Position = LMC
        Case "VC"
            GDT_Position = VC
        Case "RC"
            GDT_Position = RC
    End Select
End Function
