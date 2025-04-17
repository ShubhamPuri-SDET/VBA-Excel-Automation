Attribute VB_Name = "Module29"
Sub A_newreq_DataMismatch_Validation(Optional ByVal dummy As Integer = 0)

    Dim wsSource As Worksheet
    Dim wbReport As Workbook
    Dim wsReport As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim cell As Range
    Dim reportRow As Long
    Dim headerName As String
    Dim adjacentValue As String
    Dim col As Long
    Dim reportFileName As String
    Dim wb As Workbook
    Dim folderPath As String
    Dim sheetName As String
    
    ' Set the folder path and create the folder if it doesn't exist
    folderPath = "C:\Users\shubham.puri\Genentech\SEM Automation Scripts\SEM Validation Report"
    If Dir(folderPath, vbDirectory) = "" Then
        MkDir folderPath
    End If

    ' Report file name
    reportFileName = "SEM Validation Report " & Format(Now(), "yyyy-mm-dd hh-mm") & ".xlsx"

    ' Check if the report workbook is already open
    On Error Resume Next
    Set wbReport = Nothing
    For Each wb In Workbooks
        If wb.Name Like "SEM Validation Report *.xlsx" Then
            Set wbReport = wb
            Exit For
        End If
    Next wb
    On Error GoTo 0

    ' If the report workbook does not exist, create it
    If wbReport Is Nothing Then
        Set wbReport = Workbooks.Add
        wbReport.SaveAs Filename:=folderPath & "\" & reportFileName
    End If

    ' Create or set the report sheet
    On Error Resume Next
    Set wsReport = wbReport.Sheets("Validation Report")
    On Error GoTo 0

    If wsReport Is Nothing Then
        Set wsReport = wbReport.Sheets.Add
        wsReport.Name = "Validation Report"
        
        ' Add headers to the report
        With wsReport
            .Cells(1, 1).Value = "Sheet Name"
            .Cells(1, 2).Value = "Column Header"
            .Cells(1, 3).Value = "Cell Reference"
            .Cells(1, 4).Value = "Expected Result"
            .Cells(1, 5).Value = "Actual Result"
            
            ' Make Sheet Name header bold and yellow
            .Cells(1, 1).Font.Bold = True
            .Cells(1, 1).Interior.Color = RGB(255, 255, 0)
        End With
    End If

    reportRow = wsReport.Cells(wsReport.Rows.Count, 1).End(xlUp).Row + 1

    ' Loop through each sheet in the active workbook
    For Each wsSource In ActiveWorkbook.Sheets
        sheetName = wsSource.Name

        ' Find the last row and column with data in the source sheet
        lastRow = wsSource.Cells(wsSource.Rows.Count, 1).End(xlUp).Row
        lastCol = wsSource.Cells(1, wsSource.Columns.Count).End(xlToLeft).Column

        ' Loop through each column in the source sheet
        For col = 1 To lastCol
            headerName = wsSource.Cells(1, col).Value

            ' Loop through each cell in the column, starting from the second row
            For Each cell In wsSource.Range(wsSource.Cells(2, col), wsSource.Cells(lastRow, col))
                ' Check if the cell contains an error value
                If IsError(cell.Value) Then
                    ' Check if the error is specifically a #N/A error
                    If cell.Value = CVErr(xlErrNA) Then
                        If col > 1 Then
                            adjacentValue = cell.Offset(0, -1).Text
                        Else
                            adjacentValue = "N/A"
                        End If

                        ' Document the mismatch in the report sheet
                        wsReport.Cells(reportRow, 1).Value = sheetName
                        wsReport.Cells(reportRow, 2).Value = headerName
                        wsReport.Cells(reportRow, 3).Value = cell.Address(False, False)
                        wsReport.Cells(reportRow, 4).Value = adjacentValue
                        wsReport.Cells(reportRow, 5).Value = "#N/A"
                        
                        reportRow = reportRow + 1
                    End If
                ' Alternatively, check if the cell contains the text "#N/A"
                ElseIf cell.Text = "#N/A" Then
                    If col > 1 Then
                        adjacentValue = cell.Offset(0, -1).Text
                    Else
                        adjacentValue = "N/A"
                    End If

                    ' Document the mismatch in the report sheet
                    wsReport.Cells(reportRow, 1).Value = sheetName
                    wsReport.Cells(reportRow, 2).Value = headerName
                    wsReport.Cells(reportRow, 3).Value = cell.Address(False, False)
                    wsReport.Cells(reportRow, 4).Value = adjacentValue
                    wsReport.Cells(reportRow, 5).Value = "#N/A"

                    reportRow = reportRow + 1
                End If
            Next cell
        Next col
    Next wsSource

    ' Save the report workbook
    wbReport.Save

    ' Inform the user that the report is ready
    MsgBox "Mismatches documented in the '" & wbReport.Name & "' workbook under the 'Validation Report' sheet.", vbInformation
End Sub


