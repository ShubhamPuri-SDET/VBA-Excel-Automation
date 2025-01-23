Attribute VB_Name = "Module1"
Sub LinkMappingFailing()

    Dim ws As Worksheet
    Dim outputWs As Worksheet
    Dim lastRow As Long, outputRow As Long
    Dim websiteCol As Long, linkPresenceCol As Long, actualLinkCol As Long
    Dim statusCol As Long
    Dim linkdict As Object
    Dim i As Long
    Dim statusText As String

    Set linkdict = CreateObject("Scripting.Dictionary")
    Set ws = ActiveWorkbook.ActiveSheet
    websiteCol = Application.WorksheetFunction.Match("Website to search", ws.Rows(1), 0)
    linkPresenceCol = Application.WorksheetFunction.Match("3rd Party/Whitelist Presence", ws.Rows(1), 0)
    actualLinkCol = Application.WorksheetFunction.Match("Actual Link Present", ws.Rows(1), 0)
    
    lastRow = ws.Cells(ws.Rows.Count, websiteCol).End(xlUp).Row

    On Error Resume Next
    Set outputWs = Worksheets("LinksmappingFailing")
    If outputWs Is Nothing Then
        Set outputWs = Worksheets.Add
        outputWs.Name = "LinksmappingFailing"
    Else
        outputWs.Cells.Clear
    End If
    On Error GoTo 0

    outputWs.Cells(1, 1).Value = "Website to search"
    outputWs.Cells(1, 2).Value = "3rd Party/Whitelist Presence"
    outputWs.Cells(1, 3).Value = "Actual Link Present"
    outputWs.Cells(1, 4).Value = "Status"
 
    With outputWs.Rows(1)
        .Font.Bold = True
        .Interior.Color = RGB(0, 0, 128)
        .Font.Color = RGB(255, 255, 255)
    End With
 
    outputRow = 2
    For i = 2 To lastRow
   
        If ws.Cells(i, linkPresenceCol).Value = "Third Party Link Not Present" Or ws.Cells(i, linkPresenceCol).Value = "WhiteList Link Not Present" Then
            
            If ws.Cells(i, linkPresenceCol).Value = "Third Party Link Not Present" Then
                statusText = "3rd Party link is not documented in URL matrix"
            ElseIf ws.Cells(i, linkPresenceCol).Value = "WhiteList Link Not Present" Then
                statusText = "Whitelist link is not documented in the URL matrix"
            End If
            
            
    
            If Not linkdict.exists(ws.Cells(i, actualLinkCol).Value) Then
                linkdict.Add ws.Cells(i, actualLinkCol).Value, True
  
                outputWs.Cells(outputRow, 1).Value = ws.Cells(i, websiteCol).Value
                outputWs.Cells(outputRow, 2).Value = ws.Cells(i, linkPresenceCol).Value
                outputWs.Cells(outputRow, 3).Value = ws.Cells(i, actualLinkCol).Value
                outputWs.Cells(outputRow, 4).Value = statusText
                outputRow = outputRow + 1
                
                 
            End If
            
            
        End If
    Next i
    
    MsgBox "Macro executed successfully! check the 'LinksmappingFailing' sheet for output data.", vbInformation

End Sub



