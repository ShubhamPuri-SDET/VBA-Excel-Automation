Attribute VB_Name = "Module2"
Sub B_Generate_PassingData()

    Dim ws As Worksheet
    Dim outputWs As Worksheet
    Dim lastRow As Long
    Dim imageCol As Long, altTagCol As Long, imageResultCol As Long, altResultCol As Long, websiteCol As Long
    Dim outputRow As Long
    Dim i As Long, j As Long
    Dim printedImages As Object
    Dim wsImageNotPresentOnSite As Worksheet
    Dim wsAltMismatch As Worksheet
    Dim rowMatch As Boolean

    ' Initialize dictionary to track unique values
    Set printedImages = CreateObject("Scripting.Dictionary")

  ' Set active worksheet
    Set ws = ActiveWorkbook.Worksheets("GeneralScenarios")

    ' Find the columns based on headers
    imageCol = Application.WorksheetFunction.Match("image-tosearch", ws.Rows(1), 0)
    altTagCol = Application.WorksheetFunction.Match("alttag-tosearch", ws.Rows(1), 0)
    actualAltTagCol = Application.WorksheetFunction.Match("alttag-tosearch", ws.Rows(1), 0)
    imageResultCol = Application.WorksheetFunction.Match("imageSearchResult", ws.Rows(1), 0)
    altResultCol = Application.WorksheetFunction.Match("altTagSearchResult", ws.Rows(1), 0)
    
    websiteCol = Application.WorksheetFunction.Match("Website", ws.Rows(1), 0) ' Find the Website column
    


    ' Determine the last row in the active sheet
    lastRow = ws.Cells(ws.Rows.Count, imageCol).End(xlUp).Row

    ' Add a new sheet for output
    On Error Resume Next
    Set outputWs = Worksheets("PASSING OUTPUT")
    If outputWs Is Nothing Then
        Set outputWs = Worksheets.Add
        outputWs.Name = "PASSING OUTPUT"
    End If
    On Error GoTo 0

    ' Set headers in the new output sheet
    outputWs.Cells(1, 1).Value = "Expected Image Name"
    outputWs.Cells(1, 3).Value = "Expected Alt Tag Name"
    outputWs.Cells(1, 2).Value = "Actual Alt Tag Name"
    outputWs.Cells(1, 4).Value = "Status"
    outputWs.Cells(1, 5).Value = "Cell Location"

    ' Make headers bold and highlight in yellow
    With outputWs.Rows(1)
        .Font.Bold = True
        .Interior.Color = RGB(255, 255, 0) ' Yellow highlight
    End With

    ' Start output row at 2 (below the header)
    outputRow = 2

    ' Loop through the rows and check for specific conditions
    For i = 2 To lastRow
        ' Scenario 5: Pass in both 'imageSearchResult' and 'altTagSearchResult'
        If ws.Cells(i, imageResultCol).Value = "PASS" And ws.Cells(i, altResultCol).Value = "PASS" Then
            If Not printedImages.exists(ws.Cells(i, imageCol).Value) Then
                ' Copy the values and add status "Image name alt tag report not present"
                outputWs.Cells(outputRow, 1).Value = ws.Cells(i, imageCol).Value
                outputWs.Cells(outputRow, 3).Value = ws.Cells(i, altTagCol).Value
                outputWs.Cells(outputRow, 2).Value = ws.Cells(i, actualAltTagCol).Value
                outputWs.Cells(outputRow, 4).Value = "Image name and Alt tag present on the Site"
                outputWs.Cells(outputRow, 5).Value = ws.Cells(i, imageCol).Address

                ' Add this value to the dictionary to avoid future duplicates
                printedImages.Add ws.Cells(i, imageCol).Value, True

                outputRow = outputRow + 1
            End If

        ' Scenario 6: PASS in 'imageSearchResult' and 'Not Applicable' in 'altTagSearchResult'
        ElseIf ws.Cells(i, imageResultCol).Value = "PASS" And ws.Cells(i, altResultCol).Value = "Not Applicable" Then
            If Not printedImages.exists(ws.Cells(i, imageCol).Value) Then
                outputWs.Cells(outputRow, 1).Value = ws.Cells(i, imageCol).Value
                outputWs.Cells(outputRow, 3).Value = ws.Cells(i, altTagCol).Value
                
                outputWs.Cells(outputRow, 2).Value = ws.Cells(i, actualAltTagCol).Value
                
                outputWs.Cells(outputRow, 4).Value = "This is a Background image/Background Image Present on Site"
                outputWs.Cells(outputRow, 5).Value = ws.Cells(i, imageCol).Address

                printedImages.Add ws.Cells(i, imageCol).Value, True
                outputRow = outputRow + 1
            End If
        End If
    Next i

    ' Process ImageNotPresentOnSite sheet
    Set wsImageNotPresentOnSite = ActiveWorkbook.Sheets("ImageNotPresentOnSite")
    Set wsAltMismatch = ActiveWorkbook.Sheets("Alt mismatch")
    With wsImageNotPresentOnSite
        lastRow = .Cells(.Rows.Count, "A").End(xlUp).Row

        ' Set headers for VLOOKUP columns
        .Cells(1, 6).Value = "Vlookup Passing Output - Expected Image Name"
        .Cells(1, 7).Value = "Vlookup Passing Output - Expected Alt Tag"
        
        ' Format the headers
        With .Range("F1:G1")
            .Font.Bold = True
            .Interior.Color = RGB(255, 255, 0) ' Yellow highlight
        End With

        ' Insert VLOOKUP formulas
        .Range("F2:F" & lastRow).Formula = "=IFERROR(VLOOKUP(A2,'PASSING OUTPUT'!$A:$A,1,FALSE),"""")"
        .Range("G2:G" & lastRow).Formula = "=IFERROR(VLOOKUP(C2,'PASSING OUTPUT'!$C:$C,1,FALSE),"""")"

        ' Loop through rows from bottom to top to compare and delete if conditions match
        For i = lastRow To 2 Step -1
            If .Cells(i, 6).Value = .Cells(i, 1).Value And .Cells(i, 7).Value = .Cells(i, 2).Value Then
                .Rows(i).Delete
            End If
        Next i

        ' Hide VLOOKUP columns
        .Columns("F:G").EntireColumn.Hidden = True
    End With

    ' Process Alt mismatch sheet
    With wsAltMismatch
        lastRow = .Cells(.Rows.Count, "A").End(xlUp).Row

        ' Set headers for VLOOKUP columns
        .Cells(1, 7).Value = "Vlookup Passing Output - Expected Image Name"
        .Cells(1, 8).Value = "Vlookup Passing Output - Expected Alt Tag"

        ' Format the headers
        With .Range("G1:H1")
            .Font.Bold = True
            .Interior.Color = RGB(255, 255, 0) ' Yellow highlight
        End With

        ' Insert VLOOKUP formulas
        .Range("G2:G" & lastRow).Formula = "=IFERROR(VLOOKUP(A2,'PASSING OUTPUT'!$A:$A,1,FALSE),"""")"
        .Range("H2:H" & lastRow).Formula = "=IFERROR(VLOOKUP(C2,'PASSING OUTPUT'!$C:$C,1,FALSE),"""")"

        ' Loop through rows from bottom to top to compare and delete if conditions match
        For i = lastRow To 2 Step -1
            If .Cells(i, 7).Value = .Cells(i, 1).Value And .Cells(i, 8).Value = .Cells(i, 2).Value Then
                .Rows(i).Delete
            End If
        Next i

        ' Hide VLOOKUP columns
        .Columns("G:H").EntireColumn.Hidden = True
    End With
    
    outputWs.Visible = xlSheetHidden

    MsgBox "Macros Executed Successfully, and the Passing result sheet is hidden", vbInformation

End Sub




