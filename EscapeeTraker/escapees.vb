' Contains the logic for determining the escapees 

Sub TrackShipments()
    Dim ws As Worksheet, wsStatus As Worksheet, wsScanTypes As Worksheet
    Dim currentCell As Range
    Set ws = ThisWorkbook.Sheets("Sheet1")
    Set wsStatus = ThisWorkbook.Sheets("Status")
    Set wsScanTypes = ThisWorkbook.Sheets("ScanTypes")
    Set currentCell = ws.Range("A4")

    Do While Not IsEmpty(currentCell.Value)
        If currentCell.Interior.Color <> RGB(245, 245, 245) Then
            ProcessCell currentCell, wsStatus, wsScanTypes
        End If
        Set currentCell = currentCell.Offset(1, 0)
    Loop
End Sub

Sub ProcessCell(ByRef currentCell As Range, ByVal wsStatus As Worksheet, ByVal wsScanTypes As Worksheet)
    Dim statusDescription As String, scanTypeDescription As String, description As String
    Dim scanType As String, statusCode As String, formattedDate As String
    
    scanType = Trim(currentCell.Offset(0, 12).Value)
    statusCode = Trim(currentCell.Offset(0, 13).Value)
    statusDescription = GetDescription(wsStatus, ExtractStatusCode(statusCode), "Status")
    scanTypeDescription = GetDescription(wsScanTypes, scanType, "ScanType")
    description = ConstructDescription(statusDescription, scanTypeDescription)

    EvaluateShipment currentCell, description
End Sub

Function GetDescription(ws As Worksheet, lookupValue As String, type As String) As String
    Dim rng As Range
    Set rng = ws.Range("A1:A1000").Find(What:=lookupValue, LookIn:=xlValues, LookAt:=xlWhole)
    If Not rng Is Nothing Then
        GetDescription = rng.Offset(0, 1).Value
    Else
        GetDescription = type & " not found"
    End If
End Function

Function ExtractStatusCode(fullStatus As String) As String
    Dim pos As Integer
    pos = InStr(fullStatus, " ")
    If pos > 0 Then
        ExtractStatusCode = Trim(Left(fullStatus, pos - 1))
    Else
        ExtractStatusCode = Trim(fullStatus)
    End If
End Function

Function ConstructDescription(statusDesc As String, scanTypeDesc As String) As String
    If statusDesc <> "" And scanTypeDesc <> "" Then
        ConstructDescription = statusDesc & ", " & scanTypeDesc
    ElseIf statusDesc <> "" Then
        ConstructDescription = statusDesc
    ElseIf scanTypeDesc <> "" Then
        ConstructDescription = scanTypeDesc
    Else
        ConstructDescription = "Details not available"
    End If
End Function

Sub EvaluateShipment(ByRef cell As Range, ByVal description As String)
    Dim shelfLocation As String, deliveredDate As String, brokerRelease As String
    Dim clearanceRequired As String, hub As String, userHold As String
    Dim transactionNumber As String, currentStatus As String
    
    shelfLocation = cell.Offset(0, 3).Value
    deliveredDate = cell.Offset(0, 10).Value
    brokerRelease = cell.Offset(0, 4).Value
    clearanceRequired = cell.Offset(0, 6).Value
    hub = cell.Offset(0, 11).Value
    userHold = cell.Offset(0, 5).Value
    transactionNumber = cell.Offset(0, 7).Value
    currentStatus = cell.Offset(0, 2).Value

    If EvaluateDeliveryStatus(cell, deliveredDate, clearanceRequired, description) Then Exit Sub
    If EvaluateDashboardStatus(cell, currentStatus, hub, clearanceRequired, userHold, description) Then Exit Sub
    EvaluateDefaultStatus cell, brokerRelease, clearanceRequired, userHold, description
End Sub

Function EvaluateDeliveryStatus(ByRef cell As Range, ByVal deliveredDate As String, ByVal clearanceRequired As String, ByVal description As String) As Boolean
    If deliveredDate <> "" Then
        If clearanceRequired = "No" Then
            cell.Offset(0, 2).Value = "LVS Shipment, ok to move"
        Else
            formattedDate = ExtractAndFormatDate(deliveredDate)
            cell.Resize(1, 14).Interior.Color = RGB(255, 100, 100)
            cell.Offset(0, 2).Value = "Delivered " & formattedDate & ". Cannot be recalled."
        End If
        Return True
    End If
    Return False
End Function

Function EvaluateDashboardStatus(ByRef cell As Range, ByVal currentStatus As String, ByVal hub As String, ByVal clearanceRequired As String, ByVal userHold As String, ByVal description As String) As Boolean
    If currentStatus = "Shipment is not on Dashboard" And hub <> "6124" Then
        If clearanceRequired = "Yes" Then
            cell.Offset(0, 2).Value = "At " & CleanText(hub) & " (" & description & "); This shipment does not have a valid release."
            cell.Resize(1, 14).Interior.Color = RGB(255, 100, 100)
        Else
            cell.Offset(0, 2).Value = "At " & CleanText(hub) & "; LVS Shipment, ok to move"
        End If
        Return True
    End If
    Return False
End Function

Function EvaluateDefaultStatus(ByRef cell As Range, ByVal brokerRelease As String, ByVal clearanceRequired As String, ByVal userHold As String, ByVal description As String)
    If brokerRelease = "No" And clearanceRequired = "Yes" And userHold = "E" Then
        cell.Offset(0, 2).Value = "Escaped, was to be inspected. Currently at station/hub " & CleanText(hub)
        cell.Resize(1, 14).Interior.Color = RGB(255, 100, 100)
    Else
        cell.Offset(0, 2).Value = "At " & CleanText(hub) & " (" & description & "); This shipment does not have a valid release."
        cell.Resize(1, 14).Interior.Color = RGB(255, 100, 100)
    End If
End Function

Function ExtractAndFormatDate(dateStr As String) As String
    Dim dateOnly As String
    dateOnly = Trim(Split(dateStr, " ")(0))
    ExtractAndFormatDate = Format(DateValue(dateOnly), "dddd, dd, mmm, yyyy")
End Function

Function CleanText(text As String) As String
    CleanText = Trim(Replace(text, Chr(160), " "))
End Function
