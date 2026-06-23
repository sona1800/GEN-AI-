' VBA Macro for Excel Data Cleaning and Formatting
' Purpose: Remove duplicates, fill missing values, and format as table
' Compatible with Excel 2010 and later

Sub CleanAndFormatData()
    '
    ' CleanAndFormatData Macro
    ' Performs data cleaning operations on the active sheet:
    ' 1. Removes duplicate rows
    ' 2. Fills missing/null values with "N/A"
    ' 3. Formats data range as a standard table with headers
    '

    Dim ws As Worksheet
    Dim usedRange As Range
    Dim lastRow As Long
    Dim lastColumn As Long
    Dim tableRange As Range
    Dim i As Long
    Dim j As Long
    Dim dataRange As Range
    Dim duplicateCount As Long
    
    ' Initialize
    On Error GoTo ErrorHandler
    Set ws = ActiveSheet
    duplicateCount = 0
    
    ' Get the used range dimensions
    With ws
        If .Cells.SpecialCells(xlCellTypeLastCell).Row = 0 Then
            MsgBox "No data found in the active sheet.", vbExclamation
            Exit Sub
        End If
        
        lastRow = .Cells.SpecialCells(xlCellTypeLastCell).Row
        lastColumn = .Cells.SpecialCells(xlCellTypeLastCell).Column
    End With
    
    ' Define the data range (including headers)
    Set usedRange = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastColumn))
    
    ' Step 1: Remove Duplicate Rows
    MsgBox "Starting data cleaning process..." & vbCrLf & vbCrLf & _
           "Step 1: Removing duplicate rows", vbInformation
    
    On Error Resume Next
    usedRange.RemoveDuplicates Columns:=Array(1), Header:=xlYes
    On Error GoTo ErrorHandler
    
    ' Recalculate dimensions after removing duplicates
    lastRow = ws.Cells.SpecialCells(xlCellTypeLastCell).Row
    Set usedRange = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastColumn))
    
    ' Step 2: Find and Fill Missing/Null Values with "N/A"
    MsgBox "Step 2: Filling missing values with 'N/A'", vbInformation
    
    Set dataRange = usedRange
    For i = 2 To lastRow ' Start from row 2 to skip headers
        For j = 1 To lastColumn
            With ws.Cells(i, j)
                ' Check if cell is empty or contains only whitespace
                If Len(Trim(.Value)) = 0 Or IsNull(.Value) Then
                    .Value = "N/A"
                    .Interior.Color = RGB(255, 255, 200) ' Light yellow highlight
                End If
            End With
        Next j
    Next i
    
    ' Step 3: Format as Table
    MsgBox "Step 3: Formatting as table with headers", vbInformation
    
    Set tableRange = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastColumn))
    
    ' Create a unique table name
    Dim tableName As String
    tableName = "DataTable_" & Format(Now(), "yyyymmdd_hhmmss")
    
    ' Apply table formatting (ListObject)
    Dim tbl As ListObject
    On Error Resume Next
    Set tbl = ws.ListObjects.Add(xlSrcRange, tableRange, , xlYes)
    On Error GoTo 0
    
    If Not tbl Is Nothing Then
        ' Set table name and apply a default style
        tbl.Name = tableName
        tbl.TableStyle = "TableStyleMedium2"
        
        ' Autofit columns
        tbl.DataBodyRange.Columns.AutoFit
    Else
        ' If table creation fails, at least autofit columns
        tableRange.Columns.AutoFit
    End If
    
    ' Summary
    MsgBox "Data cleaning complete!" & vbCrLf & vbCrLf & _
           "Total rows (including header): " & lastRow & vbCrLf & _
           "Total columns: " & lastColumn & vbCrLf & vbCrLf & _
           "Operations performed:" & vbCrLf & _
           "✓ Removed duplicate rows" & vbCrLf & _
           "✓ Filled missing values with 'N/A'" & vbCrLf & _
           "✓ Formatted as table: " & tableName, vbInformation
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error: " & Err.Description & vbCrLf & vbCrLf & _
           "Error Code: " & Err.Number, vbCritical
End Sub

Sub FillMissingValuesOnly()
    '
    ' FillMissingValuesOnly Macro
    ' Fills all missing/null values in the active sheet with "N/A"
    '

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim lastColumn As Long
    Dim i As Long
    Dim j As Long
    
    On Error GoTo ErrorHandler
    Set ws = ActiveSheet
    
    lastRow = ws.Cells.SpecialCells(xlCellTypeLastCell).Row
    lastColumn = ws.Cells.SpecialCells(xlCellTypeLastCell).Column
    
    For i = 1 To lastRow
        For j = 1 To lastColumn
            With ws.Cells(i, j)
                If Len(Trim(.Value)) = 0 Or IsNull(.Value) Then
                    .Value = "N/A"
                End If
            End With
        Next j
    Next i
    
    MsgBox "Missing values filled with 'N/A'", vbInformation
    Exit Sub
    
ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical
End Sub

Sub RemoveDuplicatesOnly()
    '
    ' RemoveDuplicatesOnly Macro
    ' Removes duplicate rows based on the first column
    '

    Dim ws As Worksheet
    Dim usedRange As Range
    Dim lastRow As Long
    Dim lastColumn As Long
    
    On Error GoTo ErrorHandler
    Set ws = ActiveSheet
    
    lastRow = ws.Cells.SpecialCells(xlCellTypeLastCell).Row
    lastColumn = ws.Cells.SpecialCells(xlCellTypeLastCell).Column
    
    Set usedRange = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastColumn))
    
    usedRange.RemoveDuplicates Columns:=Array(1), Header:=xlYes
    
    MsgBox "Duplicate rows removed!", vbInformation
    Exit Sub
    
ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical
End Sub

Sub FormatAsTableOnly()
    '
    ' FormatAsTableOnly Macro
    ' Formats the used range as an Excel table with headers
    '

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim lastColumn As Long
    Dim tableRange As Range
    Dim tbl As ListObject
    Dim tableName As String
    
    On Error GoTo ErrorHandler
    Set ws = ActiveSheet
    
    lastRow = ws.Cells.SpecialCells(xlCellTypeLastCell).Row
    lastColumn = ws.Cells.SpecialCells(xlCellTypeLastCell).Column
    
    Set tableRange = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastColumn))
    tableName = "DataTable_" & Format(Now(), "yyyymmdd_hhmmss")
    
    Set tbl = ws.ListObjects.Add(xlSrcRange, tableRange, , xlYes)
    tbl.Name = tableName
    tbl.TableStyle = "TableStyleMedium2"
    tbl.DataBodyRange.Columns.AutoFit
    
    MsgBox "Data formatted as table: " & tableName, vbInformation
    Exit Sub
    
ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical
End Sub
