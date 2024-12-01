Attribute VB_Name = "Bolt"
Option Explicit
Public Const ASMEB11_2003_SHEET As String = "ASME B1.1-2003 Dimensions"
Public Const NAS1352_SHEET As String = "NAS1352 Dimensions"


Public Const ERROR_PLATE_NUMBER_MISMATCH As Long = vbObjectError + 513
Public Const ERROR_MORE_THAN_ONE_COLUMN As Long = vbObjectError + 514
Public Const NAS1352rangestring As String = "A22:W172"



Public Function Bolt_TensileStressArea(Major_Diameter As Double, ThreadsPerInch As Double) As Double
'Calculate the minimum cross sectional area of a bolt using eq 4 from NASA Technical Memorandum 106943 Preloaded Joint Analysis Methodology for Space Flight Systems
'Major diameter is external thread major diameter in inches
'Threads Per Inch is the number of threads per inch

Dim Pitch As Double
Pitch = 1 / ThreadsPerInch

Bolt_TensileStressArea = Application.WorksheetFunction.Pi() / 4 * (Major_Diameter - 0.9743 * Pitch) ^ 2

End Function

Public Function Bolt_ExternalThreadShearArea(InternalThreadMinorDiameter As Double, ThreadEngagementLength As Double) As Double
    'Evaluates Equation 63 of NASA Techical Memorandum 106943.
    'This equation can be used for bolt shear area and for external insert thread area.  It may also be good to account for the locking thread loss on a locking insert, if applicable.
    Bolt_ExternalThreadShearArea = 5 * Application.WorksheetFunction.Pi() * ThreadEngagementLength * InternalThreadMinorDiameter / 8
End Function

Public Function Bolt_InternalThreadShearArea(ExternalThreadMajorDiameter As Double, ThreadEngagementLength As Double) As Double
    'Evaluates Equation 76 of NASA Techical Memorandum 106943.
    'This equation can be used for insert internal thread shear area and for parent material shear area.
    Bolt_InternalThreadShearArea = 3 * Application.WorksheetFunction.Pi() * ThreadEngagementLength * ExternalThreadMajorDiameter / 4
End Function

Public Function Bolt_UltimateShearLoad(UltimateShearStrength As Double, ShearArea As Double) As Double
    'Evaluates Equation 64 of NASA Techical Memorandum 106943.
    Bolt_UltimateShearLoad = UltimateShearStrength * ShearArea
End Function

Public Function Bolt_NutFactor(BoltType As String) As Double
    Select Case BoltType
        Case "Minimum Lubricated"
            Bolt_NutFactor = 0.11
        Case "Maximum Lubricated"
            Bolt_NutFactor = 0.15
        Case "Unlubricated"
            Bolt_NutFactor = 0.2
        Case "Calculated"
            'not implemented yet
            Bolt_NutFactor = -1
    End Select
End Function

Public Function Bolt_MaxFastenerAxialLoad(MaxInitialPreload As Double, ExternalBoltLoad As Double, SafetyFactor As Double) As Double
    'Evaluates Equation 17 of NASA Techical Memorandum 106943.
    Bolt_MaxFastenerAxialLoad = MaxInitialPreload + SafetyFactor * ExternalBoltLoad

End Function

Public Function Bolt_TensileForceAllowable(UltimateTensileStrength As Double, CrossSectionalArea As Double) As Double
    'Calculates the Maximum Allowable Tensile Load
    Bolt_TensileForceAllowable = UltimateTensileStrength * CrossSectionalArea

End Function

Public Function Bolt_ThreadShearForceAllowable(UltimateShearStrength As Double, ThreadShearArea As Double) As Double
    'Calculates the Maximum Allowable Tensile Load
    Bolt_ThreadShearForceAllowable = UltimateShearStrength * ThreadShearArea

End Function


Public Function Bolt_TensionOnlyMarginOfSafety(DesignBoltLoad As Double, TensileForceAllowable As Double) As Double
    'Evaluates Equation 53 of NASA Techical Memorandum 106943.
    Bolt_TensionOnlyMarginOfSafety = TensileForceAllowable / DesignBoltLoad - 1
End Function

Public Function Bolt_ThreadShearMarginOfSafety(DesignBoltLoad As Double, ThreadShearForceAllowable As Double) As Double
    'Evaluates Equation 65 of NASA Techical Memorandum 106943.
    Bolt_ThreadShearMarginOfSafety = ThreadShearForceAllowable / DesignBoltLoad - 1
End Function


Function Read_ASME_B11_2003() As Dictionary
    'Function for reading the ASME B1.1-2003 Standard into a VBA dictionary for easy use elsewhere
    
    'Threadsheet is the sheet with all the information about threads
    Dim ThreadSheet As Worksheet
    'lastRow is the last row of the thread sheet. lastColumn is the last column of the same sheet.  i is just an iterator
    Dim lastRow As Long, lastColumn As Long, i As Long
    'ThreadData is the raw data from all of the threads in a 2D array
    Dim ThreadData() As Variant
    'Thread Dictionary is the dictionary that will be filled and output
    Dim ThreadDictionary As New Scripting.Dictionary
    'Thread row is the 1D array with just one thread's info
    Dim ThreadRow() As Variant
    'Check is for the GetRow function to check that it succeeded
    Dim Check As Boolean
    'oThread is an object with all the thread data in it
    Dim oThread As Thread
    
    'Read worksheet given in input, assuming the first row is just headers
    'ASMEB11_2003_SHEET is a global variable that needs to be input by the used at the top if changed
    Set ThreadSheet = Worksheets(ASMEB11_2003_SHEET)
    lastRow = ThreadSheet.Cells(Rows.Count, 1).End(xlUp).row
    lastColumn = ThreadSheet.Cells(1, Columns.Count).End(xlToLeft).column
    ThreadData = Range(ThreadSheet.Cells(2, 1), ThreadSheet.Cells(lastRow, lastColumn)).Value2
    
    'Iterate through the thread data array's rows and assign them to keys in the dictionary
    'The sheet has headers so we need to subtract 1 to not go overflow
    For i = 1 To (lastRow - 1)
        Set oThread = New Thread
        Check = GetRow(ThreadData, ThreadRow, i)
        oThread.InitiateThread ThreadRow
        Set ThreadDictionary(oThread.Name) = oThread
    Next i
    
    'Output the dictionary
    Set Read_ASME_B11_2003 = ThreadDictionary
    
End Function

Public Function Read_NAS1352(PartNo As String, Output As String) As Variant
    Dim NAS1352values() As Variant, ColumnNames() As Variant, AvailablePartNos() As Variant
    Dim i As Long, Success As Boolean
    Dim NAS1352partno As Variant, outputheader As Variant
    Dim row As Long, column As Long
    
    Dim NAS1352Sheet As Worksheet
    Set NAS1352Sheet = Worksheets(NAS1352_SHEET)
    
    NAS1352values = NAS1352Sheet.Range(NAS1352rangestring).Value
    Success = GetRow(NAS1352values, ColumnNames, 1)
    Success = GetColumn(NAS1352values, AvailablePartNos, 1)
    
    row = 0
    For Each NAS1352partno In AvailablePartNos
        row = row + 1
        If NAS1352partno = PartNo Then
            Exit For
        End If
    Next
    
    column = 0
    For Each outputheader In ColumnNames
        column = column + 1
        If outputheader = Output Then
            Exit For
        End If
    Next
    
    Read_NAS1352 = NAS1352values(row, column)

End Function

Public Function Bolt_ThreadDimensions(ThreadName As String, OutputValue As String) As Variant
    'Output Value Options for Copying (uncomment block, then copy, then comment again if you don't want to have to take away the 's)
    'Allowance
    'External Thread Class
    'External Major Diameter Max
    'External Major Diameter Min
    'External Major Diameter ASME Note 3
    'External Pitch Diameter Max
    'External Pitch Diameter Min
    'External Pitch Diameter Tolerance
    'External UNR Minor Diameter Max Reference
    'Internal Thread Class
    'Internal Thread Major Diameter Min
    'Internal Thread Minor Diameter Max
    'Internal Thread Minor Diameter Min
    'Internal Thread Pitch Diameter Max
    'Internal Thread Pitch Diameter Min
    'Internal Thread Pitch Diameter Tolerance

    Dim ThreadDictionary As Scripting.Dictionary
    Dim OutputThread As Thread
    
    Set ThreadDictionary = Read_ASME_B11_2003()
    
    Set OutputThread = ThreadDictionary(ThreadName)
    
    Select Case OutputValue
        Case "Allowance"
            Bolt_ThreadDimensions = OutputThread.Allowance
        Case "External Major Diameter Max"
            Bolt_ThreadDimensions = OutputThread.ExtMajorDMax
        Case "External Major Diameter Min"
            Bolt_ThreadDimensions = OutputThread.ExtMajorDMin
        Case "External Major Diameter ASME Note 3"
            Bolt_ThreadDimensions = OutputThread.ExtMajorDMinNote3
        Case "External Pitch Diameter Max"
            Bolt_ThreadDimensions = OutputThread.ExtPitchDMax
        Case "External Pitch Diameter Min"
            Bolt_ThreadDimensions = OutputThread.ExtPitchDmin
        Case "External Pitch Diameter Tolerance"
            Bolt_ThreadDimensions = OutputThread.ExtPitchDTol
        Case "External Thread Class"
            Bolt_ThreadDimensions = OutputThread.ExtThreadClass
        Case "External UNR Minor Diameter Max Reference"
            Bolt_ThreadDimensions = OutputThread.ExtUNRMinorDMaxRef
        Case "Internal Thread Major Diameter Min"
            Bolt_ThreadDimensions = OutputThread.IntMajorDMin
        Case "Internal Thread Minor Diameter Max"
            Bolt_ThreadDimensions = OutputThread.IntMinorDMax
        Case "Internal Thread Minor Diameter Min"
            Bolt_ThreadDimensions = OutputThread.IntMinorDMin
        Case "Internal Thread Pitch Diameter Max"
            Bolt_ThreadDimensions = OutputThread.IntPitchDMax
        Case "Internal Thread Pitch Diameter Min"
            Bolt_ThreadDimensions = OutputThread.IntPitchDmin
        Case "Internal Thread Pitch Diameter Tolerance"
            Bolt_ThreadDimensions = OutputThread.IntPitchDTol
        Case "Internal Thread Class"
            Bolt_ThreadDimensions = OutputThread.IntThreadClass
        Case "External Thread Minor Diameter Min"
            Bolt_ThreadDimensions = OutputThread.ExtMinorThreadDMin
        Case "External Thread Minor Diameter Max"
            Bolt_ThreadDimensions = OutputThread.ExtMinorThreadDMax
        Case "Pitch"
            Bolt_ThreadDimensions = OutputThread.Pitch
    End Select
    
End Function

Public Function Bolt_CircularRodAxialStiffness(Elastic_Modulus As Double, Cross_Sectional_Area As Double, Length As Double) As Double
'Calculates the stiffness in force/length of a circular rod in tension
Bolt_CircularRodAxialStiffness = Elastic_Modulus * Cross_Sectional_Area / Length
End Function

Public Function Bolt_NASA_EffectiveCounterSunkHeadDiameter(CounterSunkHeadDiameter As Double, BoltDiameter As Double) As Double
    'Evaluates Equation 39 of NASA Techical Memorandum 106943.
    Bolt_NASA_EffectiveCounterSunkHeadDiameter = (CounterSunkHeadDiameter + BoltDiameter) / 2
End Function

Public Function Bolt_NASA_BoltStiffness(BoltDiameter As Double, ThreadsPerInch As Double, GripLength As Double, BoltElasticModulus As Double) As Double
    'Calculates the Bolt's stiffness based upon the minor thread area and grip length
    Dim MinorThreadArea As Double
    MinorThreadArea = Bolt_TensileStressArea(BoltDiameter, ThreadsPerInch)
    Bolt_NASA_BoltStiffness = Bolt_CircularRodAxialStiffness(BoltElasticModulus, MinorThreadArea, GripLength)
    
End Function

Public Function Bolt_NASA_StiffnessFactor(BoltStiffness As Double, JointStiffness As Double) As Double
    'Evaluates Equation 29 of NASA Techical Memorandum 106943.
    Bolt_NASA_StiffnessFactor = BoltStiffness / (BoltStiffness + JointStiffness)
End Function

Public Function Bolt_InitialPreload(BoltDiameter As Double, NutFactor As Double, Torque As Double, TorqueUncertainty As Double, MaxOrMin As String) As Double
    'Evaluates Equation 1 of NASA Techical Memorandum 106943.
    Dim T As Double, K As Double, D As Double, u As Double
    
    T = Torque
    K = NutFactor
    D = BoltDiameter
    u = TorqueUncertainty
    
    Select Case MaxOrMin
        Case "Max", "max"
            Bolt_InitialPreload = T * (1 + u) / (K * D)
        Case "Min", "min"
            Bolt_InitialPreload = T * (1 - u) / (K * D)
        Case "Nominal", "nominal"
            Bolt_InitialPreload = T / (K * D)
    End Select

End Function

Public Function Bolt_BoltExternalLoadPreloadDelta(n As Double, JointFactor As Double, ExternalLoad As Double) As Double
    'Evaluates Equation 30 of NASA Techical Memorandum 106943.
    Bolt_BoltExternalLoadPreloadDelta = n * JointFactor * ExternalLoad
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''NASA CONFIG 1''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Function Bolt_NASA_Config1_L(PlateThicknesses As Range) As Double
    'Sums the thicknesses of the bolts according to equation 31 of NASA Technical Memorandum 106943
    'Config 1 is a through bolt using a washer with nut
    On Error GoTo eh
    If PlateThicknesses.Columns.Count > 1 Then
        Err.Raise ERROR_MORE_THAN_ONE_COLUMN, "Bolt_NASA_Config1_L", "Each input range can only have one column"
    End If
    
    Bolt_NASA_Config1_L = Application.WorksheetFunction.Sum(PlateThicknesses)
    
Exit Function
eh:
    'Return -1 if the function errored out and send the error message to debug
    Bolt_NASA_Config1_L = -1
    Debug.Print Err.Source & ": " & Err.Description
End Function

Public Function Bolt_NASA_Config1_Ej(PlateThicknesses As Range, PlateElasticModuli As Range, L As Double) As Double
    'Evaluates Equation 34 of NASA Techical Memorandum 106943
    Dim denominator As Double
    Dim i As Integer
    Dim PlateTs() As Variant, PlateEs() As Variant
    
    PlateTs = PlateThicknesses.Value
    PlateEs = PlateElasticModuli.Value
    
    'Input Error Handling
    On Error GoTo eh
    If PlateThicknesses.Columns.Count > 1 Or PlateElasticModuli.Columns.Count > 1 Then
        Err.Raise ERROR_MORE_THAN_ONE_COLUMN, "Bolt_NASA_Config1_Ej", "Each input range can only have one column"
    ElseIf PlateThicknesses.Count <> PlateElasticModuli.Count Then
        Err.Raise ERROR_PLATE_NUMBER_MISMATCH, "Bolt_NASA_Config1_Ej", "The plate property ranges must have the same number of elements"
    End If
    
    denominator = 0
    For i = LBound(PlateTs) To UBound(PlateTs)
        denominator = denominator + PlateTs(i, 1) / PlateEs(i, 1)
    Next i
    
    Bolt_NASA_Config1_Ej = L / denominator
Exit Function
eh:
    'Return -1 if the function errored out and send the error message to debug
    Bolt_NASA_Config1_Ej = -1
    Debug.Print Err.Source & ": " & Err.Description
End Function

Public Function Bolt_NASA_Config1_Kj(JointElasticModulus As Double, Diameter As Double, GripLength As Double) As Double
    'Evaluates Equation 33 of NASA Techical Memorandum 106943.  Diameter in the TM is supposed to be bolt shank diameter, but it seems like it should possibly be the hole diameter
    Dim Ej As Double, D As Double, L As Double, denominator As Double
    
    Ej = JointElasticModulus
    D = Diameter
    L = GripLength
    
    denominator = 2 * Log(5 * (L + 0.5 * D) / (L + 2.5 * D))
    
    'default vba log function base is e, so it's a natural log
    Bolt_NASA_Config1_Kj = Application.WorksheetFunction.Pi() * Ej * D / denominator
End Function

Public Function Bolt_NASA_Config1_n(PlateThicknesses As Range) As Double
    'Evaluates Equation 35 of NASA Techical Memorandum 106943.
    Dim L As Double, numerator As Double
    Dim i As Integer
    Dim PlateTs() As Variant
    L = Bolt_NASA_Config1_L(PlateThicknesses)
    PlateTs = PlateThicknesses.Value
    
    'Split out the first and last ones since they have to be divided by 2
    numerator = PlateTs(1, 1) / 2 + PlateTs(UBound(PlateTs), 1) / 2
    'Go through the rest
    For i = LBound(PlateTs) + 1 To UBound(PlateTs) - 1
        numerator = numerator + PlateTs(i, 1)
    Next i
    Bolt_NASA_Config1_n = numerator / L
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''NASA CONFIG 2''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Function Bolt_NASA_Config2_L(PlateThicknesses As Range, CounterSunkHeadHeight As Double) As Double
    'Sums the thicknesses of the bolts according to equation 36 of NASA Technical Memorandum 106943
    'Config 2 is a countersunk through bolt with a nut
    Dim i As Integer
    Dim L As Double
    
    'Error Handling
    On Error GoTo eh
    If PlateThicknesses.Columns.Count > 1 Then
        Err.Raise ERROR_MORE_THAN_ONE_COLUMN, "Bolt_NASA_Config2_L", "Each input range can only have one column"
    End If
    
    Dim PlateTs() As Variant
    PlateTs = PlateThicknesses.Value
    
    'Subtract Countersunk Head Height
    L = -CounterSunkHeadHeight / 2
    
    For i = LBound(PlateTs) To UBound(PlateTs)
        L = L + PlateTs(i, 1)
    Next i
    Bolt_NASA_Config2_L = L
Exit Function
eh:
    'Return -1 if the function errored out and send the error message to debug
    Bolt_NASA_Config2_L = -1
    Debug.Print Err.Source & ": " & Err.Description
End Function

Public Function Bolt_NASA_Config2_Ej(PlateThicknesses As Range, PlateElasticModuli As Range, L As Double, CounterSunkHeadHeight As Double) As Double
    'Evaluates Equation 38 of NASA Techical Memorandum 106943
    Dim denominator As Double
    Dim i As Integer
    Dim PlateTs() As Variant, PlateEs() As Variant
    
    PlateTs = PlateThicknesses.Value
    PlateEs = PlateElasticModuli.Value
    
    'Input Error Handling
    On Error GoTo eh
    If PlateThicknesses.Columns.Count > 1 Or PlateElasticModuli.Columns.Count > 1 Then
        Err.Raise ERROR_MORE_THAN_ONE_COLUMN, "Bolt_NASA_Config1_Ej", "Each input range can only have one column"
    ElseIf PlateThicknesses.Count <> PlateElasticModuli.Count Then
        Err.Raise ERROR_PLATE_NUMBER_MISMATCH, "Bolt_NASA_Config1_Ej", "The plate property ranges must have the same number of elements"
    End If
    
    'Add in the first element that includes the negative countersunk head height
    denominator = (PlateTs(1, 1) - CounterSunkHeadHeight / 2) / PlateEs(1, 1)
    For i = LBound(PlateTs) + 1 To UBound(PlateTs)
        denominator = denominator + PlateTs(i, 1) / PlateEs(i, 1)
    Next i
    
    Bolt_NASA_Config2_Ej = L / denominator
Exit Function
eh:
    'Return -1 if the function errored out and send the error message to debug
    Bolt_NASA_Config2_Ej = -1
    Debug.Print Err.Source & ": " & Err.Description
End Function

Public Function Bolt_NASA_Config2_Kj(JointElasticModulus As Double, Diameter As Double, GripLength As Double, EffectiveCounterSunkHeadDiameter As Double) As Double
    'Evaluates Equation 38 of NASA Techical Memorandum 106943.  Diameter in the TM is supposed to be bolt shank diameter, but it seems like it should possibly be the hole diameter
    Dim Ej As Double, D As Double, L As Double, dw As Double, denominator As Double
    
    Ej = JointElasticModulus
    D = Diameter
    dw = EffectiveCounterSunkHeadDiameter
    L = GripLength
    
    'default vba log function base is e, so it's a natural log
    denominator = Log((L + dw - D) * (dw + D) * (L + 0.5 * D) / ((L + dw + D) * (dw - D) * (L + 2.5 * D)))
    
    Bolt_NASA_Config2_Kj = Application.WorksheetFunction.Pi() * Ej * D / denominator
End Function

Public Function Bolt_NASA_Config2_n(PlateThicknesses As Range, CounterSunkHeadHeight As Double) As Double
    'Evaluates Equation 41 of NASA Techical Memorandum 106943.
    Dim denominator As Double, numerator As Double
    Dim i As Integer
    Dim PlateTs() As Variant
    denominator = Bolt_NASA_Config1_L(PlateThicknesses) 'config 1 just sums it with error checking so it's convenient here
    PlateTs = PlateThicknesses.Value
    
    'Split out the first and last ones since they have to be divided by 2
    numerator = PlateTs(1, 1) - CounterSunkHeadHeight / 2 + PlateTs(UBound(PlateTs), 1) / 2
    'Go through the rest
    For i = LBound(PlateTs) + 1 To UBound(PlateTs) - 1
        numerator = numerator + PlateTs(i, 1)
    Next i
    Bolt_NASA_Config2_n = numerator / denominator
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''NASA CONFIG 3''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Bolt_NASA_Config3_L(PlateThicknesses As Range, InsertLength As Double) As Double
    'Evaluates equation 42 of NASA Technical Memorandum 106943
    'Config 3 is a hex/socket head bolt with an insert
    Dim i As Integer
    Dim L As Double
    
    'Error Handling
    On Error GoTo eh
    If PlateThicknesses.Columns.Count > 1 Then
        Err.Raise ERROR_MORE_THAN_ONE_COLUMN, "Bolt_NASA_Config3_L", "Each input range can only have one column"
    End If
    
    Dim PlateTs() As Variant
    PlateTs = PlateThicknesses.Value
    
    'Subtract Countersunk Head Height
    L = -InsertLength / 2
    
    For i = LBound(PlateTs) To UBound(PlateTs)
        L = L + PlateTs(i, 1)
    Next i
    Bolt_NASA_Config3_L = L
Exit Function
eh:
    'Return -1 if the function errored out and send the error message to debug
    Bolt_NASA_Config3_L = -1
    Debug.Print Err.Source & ": " & Err.Description
End Function

Public Function Bolt_NASA_Config3_Ej(PlateThicknesses As Range, PlateElasticModuli As Range, L As Double, InsertLength As Double) As Double
    'TO-DO: need to take into account that the insert can be placed at the top of the plate
    'Evaluates Equation 45 of NASA Techical Memorandum 106943
    Dim denominator As Double
    Dim i As Integer
    Dim PlateTs() As Variant, PlateEs() As Variant
    
    PlateTs = PlateThicknesses.Value
    PlateEs = PlateElasticModuli.Value
    
    'Input Error Handling
    On Error GoTo eh
    If PlateThicknesses.Columns.Count > 1 Or PlateElasticModuli.Columns.Count > 1 Then
        Err.Raise ERROR_MORE_THAN_ONE_COLUMN, "Bolt_NASA_Config3_Ej", "Each input range can only have one column"
    ElseIf PlateThicknesses.Count <> PlateElasticModuli.Count Then
        Err.Raise ERROR_PLATE_NUMBER_MISMATCH, "Bolt_NASA_Config3_Ej", "The plate property ranges must have the same number of elements"
    End If
    
    denominator = (PlateTs(UBound(PlateTs), 1) - InsertLength / 2) / PlateEs(UBound(PlateEs), 1)
    For i = LBound(PlateTs) To UBound(PlateTs) - 1
        denominator = denominator + PlateTs(i, 1) / PlateEs(i, 1)
    Next i
    
    Bolt_NASA_Config3_Ej = L / denominator
Exit Function
eh:
    'Return -1 if the function errored out and send the error message to debug
    Bolt_NASA_Config3_Ej = -1
    Debug.Print Err.Source & ": " & Err.Description
End Function

Public Function Bolt_NASA_Config3_Kj(JointElasticModulus As Double, Diameter As Double, GripLength As Double) As Double
    'Evaluates Equation 44 of NASA Techical Memorandum 106943, which is the same as config 1's
    Dim Ej As Double, D As Double, L As Double, denominator As Double
    
    Ej = JointElasticModulus
    D = Diameter
    L = GripLength
    
    denominator = 2 * Log(5 * (L + 0.5 * D) / (L + 2.5 * D))
    
    'default vba log function base is e, so it's a natural log
    Bolt_NASA_Config3_Kj = Application.WorksheetFunction.Pi() * Ej * D / denominator
End Function

Public Function Bolt_NASA_Config3_n(PlateThicknesses As Range, InsertLength As Double) As Double
    'Evaluates Equation 46 of NASA Techical Memorandum 106943.
    Dim denominator As Double, numerator As Double
    Dim i As Integer
    Dim PlateTs() As Variant
    denominator = Bolt_NASA_Config1_L(PlateThicknesses) 'config 1 just sums it with error checking so it's convenient here
    PlateTs = PlateThicknesses.Value
    
    'Split out the first and the insert since they have to be divided by 2
    numerator = PlateTs(1, 1) / 2 - InsertLength / 2
    'Go through the rest
    For i = LBound(PlateTs) + 1 To UBound(PlateTs)
        numerator = numerator + PlateTs(i, 1)
    Next i
    Bolt_NASA_Config3_n = numerator / denominator
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''NASA CONFIG 4''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Bolt_NASA_Config4_L(PlateThicknesses As Range, CounterSunkHeadHeight As Double, InsertLength As Double) As Double
    'Evaluates equation 47 of NASA Technical Memorandum 106943
    'Config 4 is a countersunk head bolt with an insert
    Dim i As Integer
    Dim L As Double
    
    'Error Handling
    On Error GoTo eh
    If PlateThicknesses.Columns.Count > 1 Then
        Err.Raise ERROR_MORE_THAN_ONE_COLUMN, "Bolt_NASA_Config4_L", "Each input range can only have one column"
    End If
    
    Dim PlateTs() As Variant
    PlateTs = PlateThicknesses.Value
    
    'Subtract Countersunk Head Height
    L = -CounterSunkHeadHeight / 2 - InsertLength / 2
    
    For i = LBound(PlateTs) To UBound(PlateTs)
        L = L + PlateTs(i, 1)
    Next i
    Bolt_NASA_Config4_L = L
Exit Function
eh:
    'Return -1 if the function errored out and send the error message to debug
    Bolt_NASA_Config4_L = -1
    Debug.Print Err.Source & ": " & Err.Description
End Function

Public Function Bolt_NASA_Config4_Ej(PlateThicknesses As Range, PlateElasticModuli As Range, L As Double, CounterSunkHeadHeight As Double, InsertLength As Double) As Double
    'TO-DO: need to take into account that the insert can be placed at the top of the plate
    'Evaluates Equation 51 of NASA Techical Memorandum 106943
    Dim denominator As Double
    Dim i As Integer
    Dim PlateTs() As Variant, PlateEs() As Variant
    
    PlateTs = PlateThicknesses.Value
    PlateEs = PlateElasticModuli.Value
    
    'Input Error Handling
    On Error GoTo eh
    If PlateThicknesses.Columns.Count > 1 Or PlateElasticModuli.Columns.Count > 1 Then
        Err.Raise ERROR_MORE_THAN_ONE_COLUMN, "Bolt_NASA_Config4_Ej", "Each input range can only have one column"
    ElseIf PlateThicknesses.Count <> PlateElasticModuli.Count Then
        Err.Raise ERROR_PLATE_NUMBER_MISMATCH, "Bolt_NASA_Config4_Ej", "The plate property ranges must have the same number of elements"
    End If
    
    denominator = (PlateTs(1, 1) - CounterSunkHeadHeight / 2) / PlateEs(1, 1) + (PlateTs(UBound(PlateTs), 1) - InsertLength / 2) / PlateEs(UBound(PlateEs), 1)
    For i = LBound(PlateTs) + 1 To UBound(PlateTs) - 1
        denominator = denominator + PlateTs(i, 1) / PlateEs(i, 1)
    Next i
    
    Bolt_NASA_Config4_Ej = L / denominator
Exit Function
eh:
    'Return -1 if the function errored out and send the error message to debug
    Bolt_NASA_Config4_Ej = -1
    Debug.Print Err.Source & ": " & Err.Description
End Function

Public Function Bolt_NASA_Config4_Kj(JointElasticModulus As Double, Diameter As Double, GripLength As Double, EffectiveCounterSunkHeadDiameter As Double) As Double
    'Evaluates Equation 49 of NASA Techical Memorandum 106943.  Diameter in the TM is supposed to be bolt shank diameter, but it seems like it should possibly be the hole diameter
    Dim Ej As Double, D As Double, L As Double, dw As Double, denominator As Double
    
    Ej = JointElasticModulus
    D = Diameter
    dw = EffectiveCounterSunkHeadDiameter
    L = GripLength
    
    'default vba log function base is e, so it's a natural log
    denominator = Log((L + dw - D) * (dw + D) / ((L + dw + D) * (dw - D)))
    
    Bolt_NASA_Config4_Kj = Application.WorksheetFunction.Pi() * Ej * D / denominator
End Function

Public Function Bolt_NASA_Config4_n(PlateThicknesses As Range, CounterSunkHeadHeight As Double, InsertLength As Double) As Double
    'Evaluates Equation 52 of NASA Techical Memorandum 106943.
    Dim denominator As Double, numerator As Double
    Dim i As Integer
    Dim PlateTs() As Variant
    denominator = Bolt_NASA_Config1_L(PlateThicknesses) 'config 1 just sums it with error checking so it's convenient here
    PlateTs = PlateThicknesses.Value
    
    'Split out the first and the insert since they have to be divided by 2
    numerator = -CounterSunkHeadHeight / 2 - InsertLength / 2
    'Go through the rest
    For i = LBound(PlateTs) To UBound(PlateTs)
        numerator = numerator + PlateTs(i, 1)
    Next i
    Bolt_NASA_Config4_n = numerator / denominator
End Function





