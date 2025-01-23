Attribute VB_Name = "Module3"
Sub D_removeextension(Optional ByVal dummy As Integer = 0)
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim expectedImageCol As Long, expectedAltCol As Long, actualAltCol As Long
    Dim pairDict As Object
    Dim compositeKey As String
    Dim i As Long
    Dim fileExtension As String
    Dim imageName As String

    Set ws = ActiveWorkbook.Worksheets("GeneralScenarios")
    expectedImageCol = Application.WorksheetFunction.Match("image-tosearch", ws.Rows(1), 0)
    
'    expectedAltCol = Application.WorksheetFunction.Match("Expected Alt Tag Name", ws.Rows(1), 0)
'    actualAltCol = Application.WorksheetFunction.Match("Actual Alt Tag Name", ws.Rows(1), 0)

  
    lastRow = ws.Cells(ws.Rows.Count, expectedImageCol).End(xlUp).Row

    Set pairDict = CreateObject("Scripting.Dictionary")

    For i = 2 To lastRow
  
        imageName = ws.Cells(i, expectedImageCol).Value

        If Len(imageName) > 4 Then
            fileExtension = Right(imageName, 4)
            If fileExtension = ".png" Or fileExtension = ".jpg" Or fileExtension = ".svg" Then
                imageName = Left(imageName, Len(imageName) - 4)
                ws.Cells(i, expectedImageCol).Value = imageName
            End If
        End If
        
''        ' Populate the dictionary with Expected Image Name and Expected Alt Tag Name pairs
''        compositeKey = ws.Cells(i, expectedImageCol).Value & "|" & ws.Cells(i, expectedAltCol).Value
''        If Not pairDict.exists(compositeKey) Then
''            pairDict.Add compositeKey, True
''        End If
''    Next i
''
''    ' Loop from the bottom to the top to delete matching rows
''    For i = lastRow To 2 Step -1
''        compositeKey = ws.Cells(i, expectedImageCol).Value & "|" & ws.Cells(i, actualAltCol).Value
''        If pairDict.exists(compositeKey) Then
''            ws.Rows(i).Delete
''        End If
    Next i

    MsgBox "script removed the .extension from image name in General Scenario tab successfully!", vbInformation
End Sub


