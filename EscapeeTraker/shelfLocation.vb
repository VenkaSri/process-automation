' This checks the current shelf location of a package based on shipment number and tracking ID.

Sub CheckShelfLocation(Host)
    Dim ws As Worksheet
    Set wb = Workbooks("GenericTracker.xlsm") ' Generalized workbook name
    Set ws = wb.Sheets("MainSheet") ' Generalized worksheet name
    
    Dim currentCell As Range
    Set currentCell = ws.Range("A4")
    
    Do Until IsEmpty(currentCell.Value)

        Dim shipmentNum As String
        Dim trackingID As String
        shipmentNum = CleanText(currentCell.Value)
        trackingID = CleanText(currentCell.Offset(0, 1).Value)
        
        Dim screenHeader As String
        Host.WaitReady 10, 0
        Host.ReadScreen screenHeader, 38, 2, 23
        screenHeader = CleanText(screenHeader)
        
        If screenHeader <> "Shelf Aging Screen" Then
            MsgBox "You're not on the correct screen", vbCritical, "Error"
            Exit Sub
        End If
        
        Host.SetCursor 4, 16
        Host.WaitReady 10, 0
        ClearField Host
        Host.SendKey shipmentNum
        Host.WaitReady 10, 0
        Host.SendKey "<Enter>"
        Host.WaitReady 10, 0
        
        Dim found As Boolean
        found = False
        Dim maxScrolls As Integer
        maxScrolls = 10
        
        For scrollIndex = 1 To maxScrolls
            Dim rowOffset As Integer
            rowOffset = 8

            Do
                Dim screenShipmentNum As String
                Host.ReadScreen screenShipmentNum, 10, rowOffset, 9
                screenShipmentNum = CleanText(screenShipmentNum)
                
                If screenShipmentNum = shipmentNum Then
                    Dim screenTrackingID As String
                    Host.ReadScreen screenTrackingID, 18, rowOffset, 20
                    screenTrackingID = CleanText(screenTrackingID)
                    
                    If screenTrackingID = trackingID Then
                        Dim shelfLocation As String
                        Host.ReadScreen shelfLocation, 2, rowOffset, 57
                        shelfLocation = CleanText(shelfLocation)
                        currentCell.Offset(0, 3).Value = shelfLocation
                        found = True
                        Exit For
                    End If
                Else
                 Exit For
                End If
                
                rowOffset = rowOffset + 1
            Loop Until rowOffset > 19
            
            If found Then Exit For
            
            Host.SendKey "<RollDown>"
            Host.WaitReady 10, 0
        Next scrollIndex
        
        Set currentCell = currentCell.Offset(1, 0)
        Host.WaitReady 10, 0
    Loop
End Sub

Function CleanText(text As String) As String
    text = Trim(text)
    text = Replace(text, Chr(160), " ")
    text = WorksheetFunction.Trim(text)
    CleanText = text
End Function

Sub ClearField(Host)
    Host.SendKey "<FieldPlus>"
    Host.SendKey "<Enter>"
End Sub
