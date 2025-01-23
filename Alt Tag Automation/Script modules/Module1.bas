Attribute VB_Name = "Module1"
Sub A_Generate_FailingOutput()
    Dim ws As Worksheet
    Dim outputWs1 As Worksheet, outputWs2 As Worksheet
    Dim lastRow As Long
    Dim imageCol As Long, altTagCol As Long, imageResultCol As Long, altResultCol As Long, websiteCol As Long, actualAltTagCol As Long
    Dim outputRow1 As Long, outputRow2 As Long
    Dim printedImages1 As Object, printedImages2 As Object
    Dim compositeKey As String

    
    Module3.D_removeextension
    ' Initialize the dictionaries to track unique values
    Set printedImages1 = CreateObject("Scripting.Dictionary")
    Set printedImages2 = CreateObject("Scripting.Dictionary")

    ' Set active worksheet
    Set ws = ActiveWorkbook.Worksheets("GeneralScenarios")
      
    ' Find the columns based on headers
    imageCol = Application.WorksheetFunction.Match("image-tosearch", ws.Rows(1), 0)
    altTagCol = Application.WorksheetFunction.Match("alttag-tosearch", ws.Rows(1), 0)
    imageResultCol = Application.WorksheetFunction.Match("imageSearchResult", ws.Rows(1), 0)
    altResultCol = Application.WorksheetFunction.Match("altTagSearchResult", ws.Rows(1), 0)
    websiteCol = Application.WorksheetFunction.Match("Website", ws.Rows(1), 0)
    actualAltTagCol = Application.WorksheetFunction.Match("alt tag in image result", ws.Rows(1), 0)

    lastRow = ws.Cells(ws.Rows.Count, imageCol).End(xlUp).Row

    ' Add a new sheet for output1 ("ImageNotPresentOnSite")
    On Error Resume Next
    Set outputWs1 = Worksheets("ImageNotPresentOnSite")
    If outputWs1 Is Nothing Then
        Set outputWs1 = Worksheets.Add
        outputWs1.Name = "ImageNotPresentOnSite"
    End If
    On Error GoTo 0

    ' Add a new sheet for output2 ("Alt mismatch")
    On Error Resume Next
    Set outputWs2 = Worksheets("Alt mismatch")
    If outputWs2 Is Nothing Then
        Set outputWs2 = Worksheets.Add
        outputWs2.Name = "Alt mismatch"
    End If
    On Error GoTo 0

    ' Set headers in the "ImageNotPresentOnSite" sheet
    outputWs1.Cells(1, 1).Value = "Expected Image Name"
    outputWs1.Cells(1, 3).Value = "Expected Alt Tag Name"
    outputWs1.Cells(1, 2).Value = "Actual Alt Tag Name"
    outputWs1.Cells(1, 4).Value = "Status"
    outputWs1.Cells(1, 5).Value = "Cell Location"

    ' Set headers in the "Alt mismatch" sheet
    outputWs2.Cells(1, 1).Value = "Expected Image Name"
    outputWs2.Cells(1, 3).Value = "Expected Alt Tag Name"
    outputWs2.Cells(1, 2).Value = "Actual Alt Tag Name"
    outputWs2.Cells(1, 4).Value = "Status"
    outputWs2.Cells(1, 5).Value = "Cell Location"
    outputWs2.Cells(1, 6).Value = "URLs"

    With outputWs1.Rows(1)
        .Font.Bold = True
        .Interior.Color = RGB(255, 255, 0) ' Yellow highlight
    End With
    With outputWs2.Rows(1)
        .Font.Bold = True
        .Interior.Color = RGB(255, 255, 0)
    End With

    ' Start output rows at 2 (below the header)
    outputRow1 = 2
    outputRow2 = 2

    ' Loop through the rows and check for specific conditions
    For i = 2 To lastRow
        ' Create a composite key for duplicate checking
        compositeKey = ws.Cells(i, imageCol).Value & "|" & ws.Cells(i, altTagCol).Value & "|" & ws.Cells(i, actualAltTagCol).Value

        Select Case True
            ' Scenario 3: 'imageSearchResult' PASS and 'altTagSearchResult' FAIL
            Case ws.Cells(i, imageResultCol).Value = "PASS" And ws.Cells(i, altResultCol).Value = "FAIL"
                If Not printedImages2.exists(compositeKey) Then
                    outputWs2.Cells(outputRow2, 1).Value = ws.Cells(i, imageCol).Value
                    outputWs2.Cells(outputRow2, 3).Value = ws.Cells(i, altTagCol).Value
                    outputWs2.Cells(outputRow2, 2).Value = ws.Cells(i, actualAltTagCol).Value
                    outputWs2.Cells(outputRow2, 4).Value = "Alt tag mismatched"
                    outputWs2.Cells(outputRow2, 5).Value = ws.Cells(i, imageCol).Address
                    outputWs2.Cells(outputRow2, 6).Value = ws.Cells(i, websiteCol).Value
                    printedImages2.Add compositeKey, True
                    outputRow2 = outputRow2 + 1
                End If

            ' Scenario 2 and Scenario 1: Cross-check with Alt mismatch
            Case (ws.Cells(i, imageResultCol).Value = "FAIL" And ws.Cells(i, altResultCol).Value = "Not Applicable") Or _
                 (ws.Cells(i, imageResultCol).Value = "FAIL" And ws.Cells(i, altResultCol).Value = "FAIL")
                ' Skip if the combination is already in Alt mismatch
                If Not printedImages2.exists(compositeKey) And Not printedImages1.exists(compositeKey) Then
                    outputWs1.Cells(outputRow1, 1).Value = ws.Cells(i, imageCol).Value
                    outputWs1.Cells(outputRow1, 3).Value = ws.Cells(i, altTagCol).Value
                    outputWs1.Cells(outputRow1, 2).Value = ws.Cells(i, actualAltTagCol).Value
                    outputWs1.Cells(outputRow1, 4).Value = IIf(ws.Cells(i, altResultCol).Value = "FAIL", _
                        "Image name and Alt tag not present on the Site", _
                        "Background Image not Present on Site")
                    outputWs1.Cells(outputRow1, 5).Value = ws.Cells(i, imageCol).Address
                    printedImages1.Add compositeKey, True
                    outputRow1 = outputRow1 + 1
                End If
        End Select
    Next i
    
    ' Remove duplicate rows from "ImageNotPresentOnSite" if they match pairs in "Alt mismatch"
    lastRowOutput1 = outputWs1.Cells(outputWs1.Rows.Count, 1).End(xlUp).Row
    lastRowOutput2 = outputWs2.Cells(outputWs2.Rows.Count, 1).End(xlUp).Row

    For i = lastRowOutput1 To 2 Step -1 ' Start from bottom to avoid shifting rows
        For j = 2 To lastRowOutput2
            If outputWs1.Cells(i, 1).Value = outputWs2.Cells(j, 1).Value And _
               outputWs1.Cells(i, 2).Value = outputWs2.Cells(j, 2).Value Then
                outputWs1.Rows(i).Delete
                Exit For
            End If
        Next j
    Next i

    MsgBox "Macro executed successfully! Data moved to 'ImageNotPresentOnSite' and 'Alt mismatch' sheets."
End Sub




