Attribute VB_Name = "GDT"
Option Explicit
Public Const ERROR_MORE_THAN_ONE_ROW As Long = 514
Public Const ERROR_INPUT_NOT_5_ELEMENTS As Long = 515


Public Function GDT_Position(Nominal_Dimension As Double, Plus_Tolerance As Double, Minus_Tolerance As Double, Positional_Tolerance As Double, Dimension_Type As String, output As String) As Double

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
    Select Case output
        Case "MMC"
            GDT_Position = MMC
        Case "LMC"
            GDT_Position = LMC
        Case "VC"
            GDT_Position = VC
        Case "RC"
            GDT_Position = RC
        Case "Max"
            GDT_Position = Nominal_Dimension + Plus_Tolerance
        Case "Min"
            GDT_Position = Nominal_Dimension - Minus_Tolerance
    End Select
End Function

Public Function GDT_Difference(BigDim_Nom_Plus_Min_Pos_Type As Range, SmallDim_Nom_Plus_Min_Pos_Type As Range) As Double
    Dim Bvector() As Variant, Svector As Variant
    Dim B_Nom As Double, B_PTol As Double, B_NTol As Double, B_PosTol As Double, B_Type As String
    Dim S_Nom As Double, S_PTol As Double, S_NTol As Double, S_PosTol As Double, S_Type As String
    
    Dim B_Dim_Type As String
    
    On Error GoTo eh
    If BigDim_Nom_Plus_Min_Pos_Type.Rows.Count > 1 Or SmallDim_Nom_Plus_Min_Pos_Type.Rows.Count > 1 Then
        Err.Raise ERROR_MORE_THAN_ONE_ROW, "GDT_Difference", "Each input range can only have one row"
    ElseIf BigDim_Nom_Plus_Min_Pos_Type.Columns.Count <> 5 Or SmallDim_Nom_Plus_Min_Pos_Type.Columns.Count <> 5 Then
        Err.Raise ERROR_INPUT_NOT_5_ELEMENTS, "GDT_Difference", "Each input range has to have 5 elements"
    End If
    
    Bvector = BigDim_Nom_Plus_Min_Pos_Type.Value
    B_Nom = Bvector(1, 1)
    B_PTol = Bvector(1, 2)
    B_NTol = Bvector(1, 3)
    B_PosTol = Bvector(1, 4)
    B_Type = Bvector(1, 5)
    
    Svector = SmallDim_Nom_Plus_Min_Pos_Type.Value
    S_Nom = Svector(1, 1)
    S_PTol = Svector(1, 2)
    S_NTol = Svector(1, 3)
    S_PosTol = Svector(1, 4)
    S_Type = Svector(1, 5)
    
    Select Case B_Type
    Case "RFS Hole", "MMC Hole", "LMC Hole"
        B_Dim_Type = "Hole"
    End Select
        
    
    GDT_Difference = 1
Exit Function
eh:
    'Return -1 if the function errored out and send the error message to debug
    GDT_Difference = Err.Number
    Debug.Print Err.Source & ": " & Err.Description
End Function
