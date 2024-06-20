' Provides the updated broker name for the shipment from a mainframe system

Sub GetBroker(Host)
    Dim ws As Worksheet
    Set wb = Workbooks("GenericTracker.xlsm") ' Name changed to a generic tracker
    Set ws = wb.Sheets("MainSheet") ' Changed to a more generic sheet name
    Dim currentCell As Range
    Set currentCell = ws.Range("A4") 
    Dim screenHeader As String
    Host.WaitReady 10, 0
    Host.ReadScreen screenHeader, 27, 2, 28
    screenHeader = CleanText(screenHeader)
    If screenHeader <> "CUSTOMS BROKER CHANGE ENTRY" Then
        MsgBox "You're not on the correct screen", vbCritical, "Error"
        Exit Sub
    End If

    Do Until IsEmpty(currentCell.Value)
        Host.SetCursor 9, 40
        Host.WaitReady 10, 0
        Host.SendKey currentCell.Value
        Host.SendKey "<FieldPlus>"
        Host.WaitReady 10, 0
        Host.SendKey "<Enter>"
        Host.WaitReady 10, 0
        Dim shipmentNotOnFile As String
        Host.ReadScreen shipmentNotOnFile, 28, 23, 2
        Dim screenShipmentNum As String
        Host.ReadScreen screenShipmentNum, 10, 5, 48
        screenShipmentNum = CleanText(screenShipmentNum)
        If screenShipmentNum = currentCell.Value Then
            Dim BrokerName As String
            Host.ReadScreen BrokerName, 40, 10, 30
            BrokerName = CleanText(BrokerName)
 
            currentCell.Offset(0, 8).Value = BrokerName
        ElseIf shipmentNotOnFile = "Shipment Number not on file." Then
            currentCell.Offset(0, 8).Value = "Shipment Number not on file."
            GoTo Skip
        End If
        Host.SendKey "<PF3>"
        Host.WaitReady 10, 0
Skip:
        Set currentCell = currentCell.Offset(1, 0)
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
