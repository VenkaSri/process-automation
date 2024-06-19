' This module contains subs that are responsible for accuratlely displaying the current station/status code/scan type seen on FTRACk. 

Sub FindEscapees()
    Dim ws As Worksheet, wsScans As Worksheet
    Set wb = Workbooks("EscapeeTracker.xlsm")
    Set ws = wb.Sheets("Sheet1")
    Set wsScans = wb.Sheets("Scans")
    Dim currentCell As Range
    Set currentCell = ws.Range("A4")
    Dim shipmentID As String
    Dim trackingID As String
    Dim currentLocation As String
    Dim status As String
    
    Do Until IsEmpty(currentCell.Value)
    If currentCell.Interior.Color <> RGB(245, 245, 245) Then

        shipmentID = CleanText(currentCell.Value)
        trackingID = CleanText(currentCell.Offset(0, 1).Value)
        status = currentCell.Offset(0, 2).Value
        
        currentLocation = GetCurrentLocation(shipmentID, trackingID, wsScans, currentCell)
    End If
        Set currentCell = currentCell.Offset(1, 0)
    Loop
End Sub

Function GetDelDate(shipmentID As String, trackingID As String, wsScans As Worksheet, currentCell As Range) As String
    Dim lastRow As Long
    Dim i As Long
    Dim currentLocation As String
    Dim foundDelivery As Boolean
    foundDelivery = False
    
    lastRow = wsScans.cells(wsScans.rows.Count, 1).End(xlUp).row
    
    For i = lastRow To 1 Step -1
        If CleanText(wsScans.cells(i, 1).Value) = shipmentID And CleanText(wsScans.cells(i, 2).Value) = trackingID Then
            If wsScans.cells(i, 3).Value = "Delivery" Then

                currentCell.Offset(0, 10).Value = wsScans.cells(i, 4).Value
                    
            End If
        End If
    Next i
    

    GetDelDate = currentLocation
End Function


Sub FindDelivery()
    Dim ws As Worksheet, wsScans As Worksheet
    Set wb = Workbooks("EscapeeTracker.xlsm")
    Set ws = wb.Sheets("Sheet1")
    Set wsScans = wb.Sheets("Scans")
    Dim currentCell As Range
    Set currentCell = ws.Range("A4")
    Dim shipmentID As String
    Dim trackingID As String
    Dim currentLocation As String
    Dim status As String
    
    Do Until IsEmpty(currentCell.Value)
        If currentCell.Interior.Color <> RGB(245, 245, 245) Then
            shipmentID = CleanText(currentCell.Value)
            trackingID = CleanText(currentCell.Offset(0, 1).Value)
            status = currentCell.Offset(0, 2).Value
        
            currentLocation = GetDelDate(shipmentID, trackingID, wsScans, currentCell)
            

        End If
        Set currentCell = currentCell.Offset(1, 0)
    Loop
End Sub
