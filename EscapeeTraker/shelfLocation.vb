' Checking for the current Shelf Location of a package

Sub CheckAging(Host)
    Dim ws As Worksheet
    Set wb = Workbooks("EscapeeTracker.xlsm")
    Set ws = wb.Sheets("Sheet1")
    
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
        
        If screenHeader <> "Destination Gateway Shelf Aging Screen" Then
            MsgBox "You're not on the Aging screen", vbCritical, "Error"
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
                        Dim agingDetail As String
                        Host.ReadScreen agingDetail, 2, rowOffset, 57
                        agingDetail = CleanText(agingDetail)
                        currentCell.Offset(0, 3).Value = agingDetail
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


