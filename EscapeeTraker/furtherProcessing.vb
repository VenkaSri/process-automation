
' This module evaluates each shipment to determine if it requires further processing.
' It checks various shipment-related parameters to update the status and cell color based on specific conditions.

Sub FurtherProcessing()
    
    Dim ws As Worksheet
    Set ws = Workbooks("EscapeeTracker.xlsm").Sheets("Sheet1")

    Dim currentCell As Range
    Set currentCell = ws.Range("A4")

    Do Until IsEmpty(currentCell.Value)
        ' Read relevant data from the current row
        Dim shelfLocation As String
        Dim userHold As String
        Dim brokerRelease As String
        Dim clearanceRequired As String
        Dim releaseType As String

        shelfLocation = ws.Cells(currentCell.Row, 4).Value
        userHold = ws.Cells(currentCell.Row, 6).Value
        brokerRelease = ws.Cells(currentCell.Row, 5).Value
        clearanceRequired = ws.Cells(currentCell.Row, 7).Value
        releaseType = ws.Cells(currentCell.Row, 8).Value
        
        ' Skip processing for cells with specific background colors
        If currentCell.Interior.Color = RGB(245, 245, 245) Or currentCell.Interior.Color = RGB(254, 254, 254) Then
            GoTo NextIteration
        End If
        
        ' Determine and set the status based on shipment parameters
        If shelfLocation <> "" And userHold <> "E" Then
            ws.Cells(currentCell.Row, 3).Value = "Shelved"
            currentCell.Interior.Color = RGB(245, 245, 245)
        ElseIf brokerRelease = "Yes" And UserHoldInList(userHold) And clearanceRequired = "Yes" And releaseType <> "" Then
            ws.Cells(currentCell.Row, 3).Value = "Released"
            currentCell.Interior.Color = RGB(245, 245, 245)
        ElseIf (brokerRelease = "No" Or brokerRelease = "Yes") And UserHoldInList(userHold) And clearanceRequired = "No" Then
            ws.Cells(currentCell.Row, 3).Value = "LVS Shipment, ok to move."
            currentCell.Interior.Color = RGB(245, 245, 245)
        End If

NextIteration:
        ' Move to the next cell
        Set currentCell = currentCell.Offset(1, 0)
    Loop
End Sub

Function UserHoldInList(userHold As String) As Boolean
    ' Helper function to check if userHold is one of several specific codes
    UserHoldInList = (userHold = "R" Or userHold = "O" Or userHold = "I" Or userHold = "G" Or userHold = "M" Or userHold = "S")
End Function
