Attribute VB_Name = "Utils"
Option Explicit


Public Function Read_Table(TableSheetName As String, TableRangeString As String, DesiredRow As String, DesiredColumn As String) As Variant
    Dim TableValues() As Variant, ColumnNames() As Variant, RowNames() As Variant
    Dim i As Long, Success As Boolean
    Dim RowName As Variant, ColumnName As Variant
    Dim DesiredRowNum As Long, DesiredColNum As Long
    Dim row As Long, column As Long
    
    Dim TableSheet As Worksheet
    Set TableSheet = Worksheets(TableSheetName)
    
    TableValues = TableSheet.Range(TableRangeString).Value
    Success = GetRow(TableValues, ColumnNames, 1)
    Success = GetColumn(TableValues, RowNames, 1)
    
    row = 0
    For Each RowName In RowNames
        row = row + 1
        If RowName = DesiredRow Then
            DesiredRowNum = row
            Exit For
        End If
    Next
    
    column = 0
    For Each ColumnName In ColumnNames
        column = column + 1
        If ColumnName = DesiredColumn Then
            DesiredColNum = column
            Exit For
        End If
    Next
    
    Read_Table = TableValues(DesiredRowNum, DesiredColNum)

End Function
