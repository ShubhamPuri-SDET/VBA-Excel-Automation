Attribute VB_Name = "Module4"
Sub D_Deleteduplicatepairs_AltMismatchTab()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim expectedImageCol As Long, expectedAltCol As Long, actualAltCol As Long
    Dim pairDict As Object
    Dim compositeKey As String
    Dim i As Long
    Dim rowsToDelete As Collection
    Dim rowNum As Long
    Dim colValueCount As Integer
    Dim rowValueCount As Integer

    ' Set the worksheet
    Set ws = ActiveWorkbook.Worksheets("Alt mismatch")
    
    ' Identify column indices
    On Error Resume Next
    expectedImageCol = Application.WorksheetFunction.Match("Expected Image Name", ws.Rows(1), 0)
    expectedAltCol = Application.WorksheetFunction.Match("Expected Alt Tag Name", ws.Rows(1), 0)
    actualAltCol = Application.WorksheetFunction.Match("Actual Alt Tag Name", ws.Rows(1), 0)
    On Error GoTo 0
    
    If expectedImageCol = 0 Or expectedAltCol = 0 Or actualAltCol = 0 Then
        MsgBox "One or more required columns are missing. Please check the column headers.", vbExclamation
        Exit Sub
    End If

    ' Find the last row
    lastRow = ws.Cells(ws.Rows.Count, expectedImageCol).End(xlUp).Row

    ' Initialize dictionary and collection
    Set pairDict = CreateObject("Scripting.Dictionary")
    Set rowsToDelete = New Collection

    ' Build dictionary and identify rows to delete
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    For i = 2 To lastRow
        ' Create composite key
        compositeKey = ws.Cells(i, expectedImageCol).Value & "|" & ws.Cells(i, actualAltCol).Value
        
        If pairDict.exists(compositeKey) Then
            ' Compare the number of non-empty cells
            rowNum = pairDict(compositeKey)
            colValueCount = CountNonEmptyCells(ws, i)
            rowValueCount = CountNonEmptyCells(ws, rowNum)

            If colValueCount > rowValueCount Then
                rowsToDelete.Add rowNum ' Mark the earlier row for deletion
                pairDict(compositeKey) = i ' Update to the current row
            Else
                rowsToDelete.Add i ' Mark the current row for deletion
            End If
        Else
            pairDict.Add compositeKey, i ' Add the composite key and row number
        End If
    Next i

    ' Delete rows in bulk to reduce overhead
    DeleteRows ws, rowsToDelete

    ' Restore application settings
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

    ' Completion message
    MsgBox "Removed duplicates successfully!", vbInformation
End Sub

' Function to count non-empty cells in a row
Function CountNonEmptyCells(ws As Worksheet, rowNum As Long) As Integer
    Dim lastCol As Long, colCount As Integer, j As Long
    lastCol = ws.Cells(rowNum, ws.Columns.Count).End(xlToLeft).Column
    For j = 1 To lastCol
        If Len(ws.Cells(rowNum, j).Value) > 0 Then
            colCount = colCount + 1
        End If
    Next j
    CountNonEmptyCells = colCount
End Function

' Subroutine to delete rows in bulk
Sub DeleteRows(ws As Worksheet, rowsToDelete As Collection)
    Dim i As Long, rng As Range
    If rowsToDelete.Count = 0 Then Exit Sub

    ' Combine rows into a single range
    For i = 1 To rowsToDelete.Count
        If rng Is Nothing Then
            Set rng = ws.Rows(rowsToDelete(i))
        Else
            Set rng = Union(rng, ws.Rows(rowsToDelete(i)))
        End If
    Next i

    ' Delete all rows at once
    If Not rng Is Nothing Then rng.Delete
End Sub

