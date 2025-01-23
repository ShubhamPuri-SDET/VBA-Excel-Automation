Attribute VB_Name = "Module5"
Sub C_Scenario3_AltImgdata()
    Dim ws As Worksheet
    Dim outputWs As Worksheet
    Dim altMismatchWs As Worksheet
    Dim wsImageNotPresentOnSite As Worksheet
    Dim lastRow As Long, mismatchRow As Long
    Dim websiteCol As Long, altTagCol As Long, imgTagCol As Long
    Dim altPresenceCol As Long, imgPresenceCol As Long
    Dim outputRow As Long
    Dim uniqueEntries As Object
    Dim compositeKey As String
    
    Set uniqueEntries = CreateObject("Scripting.Dictionary")
    Set ws = ActiveWorkbook.Worksheets("Scenario3")
    
    
    websiteCol = Application.WorksheetFunction.Match("Website", ws.Rows(1), 0)
    altTagCol = Application.WorksheetFunction.Match("AltTags", ws.Rows(1), 0)
    imgTagCol = Application.WorksheetFunction.Match("Img Tag", ws.Rows(1), 0)
    altPresenceCol = Application.WorksheetFunction.Match("Alt Tag Presence in Excel Sheet?", ws.Rows(1), 0)
    imgPresenceCol = Application.WorksheetFunction.Match("Img Tag Presence in Excel Sheet?", ws.Rows(1), 0)
    
    
    lastRow = ws.Cells(ws.Rows.Count, websiteCol).End(xlUp).Row
 
    On Error Resume Next
    Set outputWs = Worksheets("ImageNameAltTagNotDocumented")
    If outputWs Is Nothing Then
        Set outputWs = Worksheets.Add
        outputWs.Name = "ImageNameAltTagNotDocumented"
    Else
        outputWs.Cells.Clear
    End If
    On Error GoTo 0
    
    
    outputWs.Cells(1, 1).Value = "Img Tag"
   ' outputWs.Cells(1, 2).Value = "Expected Alt Tag"
     outputWs.Cells(1, 2).Value = "Alt Tag"
     outputWs.Cells(1, 3).Value = "Website"
    
    outputWs.Cells(1, 4).Value = "Status"
  

    With outputWs.Rows(1)
        .Font.Bold = True
        .Interior.Color = RGB(255, 255, 0)
    End With
    
   
    outputRow = 2
    
        ' Loop through the data and evaluate scenarios
        For i = 2 To lastRow
            compositeKey = ws.Cells(i, altTagCol).Value & "|" & ws.Cells(i, imgTagCol).Value
            
            ' Scenario 1: Both Alt Tag and Img Tag are "Not Present"
            If ws.Cells(i, altPresenceCol).Value = "Not Present" And ws.Cells(i, imgPresenceCol).Value = "Not Present" Then
                If Not uniqueEntries.exists(compositeKey) Then
                    outputWs.Cells(outputRow, 3).Value = ws.Cells(i, websiteCol).Value
                    outputWs.Cells(outputRow, 2).Value = ws.Cells(i, altTagCol).Value
                    outputWs.Cells(outputRow, 1).Value = ws.Cells(i, imgTagCol).Value
                   ' outputWs.Cells(outputRow, 2).Value = ws.Cells(i).Value
                    outputWs.Cells(outputRow, 4).Value = "ImgName and AltTag Data present on site but not documented"
                    uniqueEntries.Add compositeKey, True
                    outputRow = outputRow + 1
                End If
            End If
            
            ' Scenario 2: Alt Tag is "Not Present" and Img Tag is "Present"
            If ws.Cells(i, altPresenceCol).Value = "Not Present" And ws.Cells(i, imgPresenceCol).Value = "Present" Then
                If Not uniqueEntries.exists(compositeKey) Then
                    outputWs.Cells(outputRow, 3).Value = ws.Cells(i, websiteCol).Value
                    outputWs.Cells(outputRow, 2).Value = ws.Cells(i, altTagCol).Value
                    outputWs.Cells(outputRow, 1).Value = ws.Cells(i, imgTagCol).Value
                    outputWs.Cells(outputRow, 4).Value = "AltTag Mismatched"
                    uniqueEntries.Add compositeKey, True
                    outputRow = outputRow + 1
                End If
            End If
        Next i
    

' Process Alt mismatch sheet
On Error Resume Next
Set wsAltMismatch = ActiveWorkbook.Sheets("Alt mismatch")
If wsAltMismatch Is Nothing Then
    Set wsAltMismatch = Worksheets.Add
    wsAltMismatch.Name = "Alt mismatch"
End If
On Error GoTo 0

' Insert VLOOKUP formulas for Expected Image Name and Alt Tag
With outputWs
    lastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
    .Cells(1, 6).Value = "Vlookup Passing Output - Expected Image Name"
    .Cells(1, 7).Value = "Vlookup Passing Output - Expected Alt Tag"

    With .Range("E1:F1")
        .Font.Bold = True
        .Interior.Color = RGB(255, 255, 0)
    End With

    ' Insert VLOOKUP formulas
    .Range("F2:F" & lastRow).Formula = "=IFERROR(VLOOKUP(A2,'Alt mismatch'!A:C,1,FALSE),"""")"
    .Range("G2:G" & lastRow).Formula = "=IFERROR(VLOOKUP(B2,'Alt mismatch'!A:C,3,FALSE),"""")"

    ' Loop through rows from bottom to top to compare and delete if conditions match
    For i = lastRow To 2 Step -1
        If .Cells(i, 6).Value = .Cells(i, 1).Value And .Cells(i, 7).Value = .Cells(i, 2).Value Then
            .Rows(i).Delete
        End If
    Next i

    ' Optionally clear VLOOKUP helper columns if not needed
    .Columns("F:G").ClearContents
End With

' Move "AltTag Mismatched" rows to the Alt mismatch sheet
With outputWs
    lastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
    mismatchRow = wsAltMismatch.Cells(wsAltMismatch.Rows.Count, "A").End(xlUp).Row + 1

    ' Copy rows with "AltTag Mismatched" status
    For i = lastRow To 2 Step -1
        If .Cells(i, 4).Value = "AltTag Mismatched" Then
            wsAltMismatch.Rows(mismatchRow).Value = .Rows(i).Value
            mismatchRow = mismatchRow + 1
            .Rows(i).Delete
        End If
    Next i
End With

' Highlight the rows in the Alt mismatch sheet
With wsAltMismatch
    .Rows("1:1").Font.Bold = True
    .Rows("1:1").Interior.Color = RGB(255, 255, 0) ' Yellow header
End With

MsgBox "Macro executed successfully! Data processed, cleaned up, and mismatches moved to 'Alt mismatch' sheet.", vbInformation
End Sub


