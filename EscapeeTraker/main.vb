' This module is designed to automate various tasks associated with managing and tracking shipments.
' It connects to a BlueZone host terminal to perform tasks such as logging in, navigating through the host system,
' and retrieving information related to clearance requirements and shipment details.
' After performing operations on the host, the module processes data within an Excel workbook, including clearing formats,
' executing predefined macros like PITT3Main, FurtherProcessing, FTRACK, and TrackShipments to update and manage shipment data effectively.
' The module handles errors gracefully by alerting the user to any connection issues or operational errors during the execution.



Sub Main()
    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Dim Host As Object
    Dim ResultCode As Long

    Set ws = Workbooks("EscapeeTracker.xlsm").Sheets("Sheet1")
    ws.Range("A3:M1000").ClearFormats

    Set Host = InitializeHost("OPS2.zad")
    If Host Is Nothing Then Exit Sub

    LoginToHost Host
    NavigateToNorthboundMenu Host
    ProcessClearanceRequirements Host
    ReturnToDashboard Host

    CloseHostSession Host


    PITT3Main
    FurtherProcessing
    FTRACK
    TrackShipments

    Exit Sub

ErrHandler:
    MsgBox "An error occurred: " & Err.Description, vbCritical
End Sub

Function InitializeHost(ByVal zadFile As String) As Object
    Dim Host As Object
    Dim ResultCode As Long
    Set Host = CreateObject("BZWhll.WhllObj")
    ResultCode = Host.OpenSession(1, 5, zadFile, 30, 1)
    If ResultCode <> 0 Then
        MsgBox "Error connecting to host!", vbCritical
        Set InitializeHost = Nothing
        Exit Function
    End If

    ResultCode = Host.Connect("!")
    If ResultCode <> 0 Then
        MsgBox "Error connecting to session", vbCritical
        Set InitializeHost = Nothing
        Exit Function
    End If

    Set InitializeHost = Host
End Function

Sub LoginToHost(ByVal Host As Object)
    With Host
        .SendKey <REDACTED>
        .SendKey <REDACTED>
        For i = 1 To 4
            .SendKey "<Enter>"
            .WaitReady 10, 0
        Next i
    End With
End Sub

Sub NavigateToNorthboundMenu(ByVal Host As Object)
    With Host
        .SendKey "03<Enter>27<Enter>5<Enter>1<Enter>"
        .WaitReady 10, 0
    End With
End Sub

Sub ProcessClearanceRequirements(ByVal Host As Object)
    With Host
        .SendKey "<PF3>"
        .WaitReady 10, 0
        .SendKeys "5<Enter>3<Enter>"
        .WaitReady 10, 0
        .SendKey "<PF3>"
        .WaitReady 10, 0
    End With
End Sub

Sub ReturnToDashboard(ByVal Host As Object)
    With Host
        .SetCursor 21, 29
        .SendKey "1<Enter>7564843"
        .SetCursor 12, 49
        .SendKey "1234<Enter><PF13>"
        .WaitReady 10, 0
    End With
End Sub

Sub CloseHostSession(ByVal Host As Object)
    Host.CloseSession 1, 5
End Sub
