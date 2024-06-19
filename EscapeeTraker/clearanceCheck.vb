' Checks weather clearance is required for shipments

Sub ClearanceReq(Host)
    Dim ws As Worksheet
    Set wb = Workbooks("EscapeeTracker.xlsm")
    Set ws = wb.Sheets("Sheet1")
    
    Dim currentCell As Range
    Set currentCell = ws.Range("A4")

    Do Until IsEmpty(currentCell.Value)
        Dim cellContent As String
        cellContent = Trim(currentCell.Value)
        InteractWithHost Host, cellContent, currentCell
        
        
        
Skip:
        Set currentCell = currentCell.Offset(1, 0)
    Loop
End Sub

Sub InteractWithHost(Host, ByVal cellContent As String, currentCell As Range)
    Dim BufV As String
        Host.WaitReady 10, 0
        Host.ReadScreen BufV, 9, 2, 72
        BufV = Trim(BufV)
        
        If BufV <> "INTD010A" Then
            MsgBox "Not on the Entry Menu", vbCritical, "Error"
            Exit Sub
        End If

        Host.SetCursor 13, 45
        Host.SendKeys cellContent
        Host.SendKey "<FieldPlus>"
        Host.WaitReady 10, 0
        Host.SetCursor 14, 38
        Dim trimmedValue As String
        trimmedValue = Replace(currentCell.Offset(0, 1).Value, Chr(160), "")
        trimmedValue = Trim(trimmedValue)
        Host.SendKeys trimmedValue
        Host.SendKey "<FieldPlus>"
        Host.WaitReady 10, 0
        Host.SendKeys "<Enter>"
        Host.WaitReady 10, 0
        
        Dim anotherShipment As String
        Dim anotherShipmentNum As String
        Host.ReadScreen anotherShipment, 32, 24, 2

        If anotherShipment = "Track ID is on another shipment:" Then
            Host.ReadScreen anotherShipmentNum, 10, 24, 40
            anotherShipmentNum = Trim(anotherShipmentNum)
            currentCell.Value = anotherShipmentNum
            Call InteractWithHost(Host, anotherShipmentNum, currentCell)
        End If

        Dim ShipmentProcess As String
        Host.ReadScreen ShipmentProcess, 19, 8, 31
        Dim ShipmentNumChange As String
        Host.ReadScreen ShipmentNumChange, 9, 8, 48
        Dim ShipmentCannot As String
        Host.ReadScreen ShipmentCannot, 9, 10, 49
        Dim isShipmentProcessed As Boolean
        isShipmentProcessed = False
        
        If ShipmentNumChange = "cannot be" Then
            currentCell.Offset(0, 2) = "Shipment number has been changed."
            currentCell.Interior.Color = RGB(221, 221, 221)
            Host.WaitReady 10, 0
            Host.SendKey "<PF12>"
            Host.WaitReady 10, 0
            Exit Sub
        End If
        
        
        If ShipmentCannot = "cannot be" Then
            currentCell.Offset(0, 2) = "RTS Shipment"
            currentCell.Interior.Color = RGB(221, 221, 221)
            Host.WaitReady 10, 0
            Host.SendKey "<PF12>"
            Host.WaitReady 10, 0
            Exit Sub
        End If
        

        If Not isShipmentProcessed Then
            Dim Buf As String
            Do
                Host.ReadScreen Buf, 9, 2, 72
                Buf = Trim(Buf)
                If Buf = "INTD010A" Then
                    Dim attemptCount As Integer
                    attemptCount = 0
                    Do While attemptCount < 5 And Buf = "INTD010A"
                        Host.SendKeys "<Enter>"
                        Host.WaitReady 10, 0
                        Host.ReadScreen Buf, 9, 2, 72
                        Buf = Trim(Buf)
                        attemptCount = attemptCount + 1
                    Loop
                End If
                If Buf = "INTD010D" Then
                    Host.ReadScreen Bkr, 1, 21, 30
                    If Bkr <> "Y" Then
                        currentCell.Offset(0, 6).Value = "No"
                    Else
                        currentCell.Offset(0, 6).Value = "Yes"
                    End If
                    Exit Do
                Else
                    Host.SendKeys "<Enter>"
                    Host.WaitReady 10, 0
                End If
            Loop
        End If
        
        NavigateBack Host
End Sub

Sub NavigateBack(Host)
    Dim Buf As String
    Do
        Host.ReadScreen Buf, 9, 2, 72
        Buf = Trim(Buf)
        
        If Buf = "INTD010A" Then
            Exit Do
        ElseIf Buf = "INTD10C3" Then
            Host.SendKeys "<F11>"
            Host.WaitReady 10, 0
        Else
            Host.SendKeys "<F12>"
            Host.WaitReady 10, 0
        End If
    Loop
End Sub
