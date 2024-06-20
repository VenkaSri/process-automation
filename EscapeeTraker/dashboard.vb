' Retrieving highly relevant information such release, user hold flag, trans number, etc.

Sub GetTrans(Host)
    Dim ws As Worksheet
    Set wb = Workbooks("GenericName.xlsm")
    Set ws = wb.Sheets("Sheet1")
    
    Dim currentCell As Range
    Set currentCell = ws.Range("A4")

    Do Until IsEmpty(currentCell.Value)
        Dim cellContent As String
        cellContent = Trim(currentCell.Value)
        
        
        Dim BufV As String
        Host.WaitReady 10, 0
        Host.ReadScreen BufV, 21, 2, 29
        BufV = Trim(BufV)
        
        If BufV <> "SCREEN_HEADER" Then
            MsgBox "You're not on Dashboard", vbCritical, "Error"
            Exit Sub
        End If
        Dim status As String
        status = Trim(currentCell.Offset(0, 2).Value)
        
        
        If currentCell.Interior.Color = RGB(221, 221, 221) Then
            GoTo Skip
        End If
        
        
        Host.SetCursor 6, 33
        Host.SendKey currentCell.Value
        Host.SendKey "<FieldPlus>"
        Host.SendKeys "<Enter>"
        Host.WaitReady 10, 0
        Host.ReadScreen Buf, 10, 9, 10
        Host.ReadScreen BrkRel, 3, 9, 21
        Host.ReadScreen UsrHold, 1, 9, 78

        Host.ReadScreen RelType, 24, 9, 25
        Dim RType As String
        RType = Trim(RelType)
        If Trim(Buf) = cellContent Then
            currentCell.Offset(0, 4).Value = BrkRel
            currentCell.Offset(0, 5).Value = UsrHold
            If BrkRel = "Yes" Then
                    If RType = "SOMETHING" Then
                        Released Host, currentCell
                    ElseIf RType = "SOMETHING" Then
                        currentCell.Offset(0, 7).Value = "SOMETHING"
                    ElseIf InStr(RType, "RNS Doc") > 0 Then
                        Host.ReadScreen TransNumRNS, 14, 9, 33
                        currentCell.Offset(0, 7).Value = "SOMETHING" & TransNumRNS
                    End If
            Else
                NoReleaseTrans Host, currentCell
            End If
        Else
            currentCell.Offset(0, 2).Value = "Shipment is not on Dashboard"
            CallAPI currentCell
        End If
Skip:
        Set currentCell = currentCell.Offset(1, 0)
        Host.WaitReady 10, 0
    Loop
End Sub

Sub Released(Host, currentCell As Range)
    Host.SetCursor 9, 2
    Host.SendKeys "<Enter>"
    Host.WaitReady 10, 0

    Dim Buf As String
    Dim Cmmt As String
    Dim Trans As String
    Dim found As Boolean
    Dim startRow As Integer
    startRow = 6

    Do
        found = False

        For i = startRow To 20
            Host.ReadScreen Buf, 10, i, 13
            If Trim(Buf) = "SOMETHING" Or Trim(Buf) = "SOMETHING" Then
                Host.ReadScreen Cmmt, 17, i, 24
                If Trim(Cmmt) = "SOMETHING" Or Trim(Cmmt) = "SOMETHING" Then
                    Host.ReadScreen Trans, 14, i, 65
                    currentCell.Offset(0, 7).Value = "Transaction #" & Trans
                    Host.SendKey "<PF12>"
                    Exit Do
                ElseIf Cmmt = "SOMETHING" Then
                    Host.ReadScreen usrRel, 17, i, 65
                    currentCell.Offset(0, 7).Value = "SOMETHING"
                    Host.SendKey "<PF12>"
                    Exit Do
                Else
                    startRow = i + 1
                    found = True
                    Exit For
                End If
            End If
        Next i
        If Not found Then
            CallAPI currentCell
            Host.SendKey "<PF12>"
            Exit Do
        End If
    Loop While found
End Sub

Sub NoReleaseTrans(Host, currentCell As Range)
    Host.SetCursor 9, 2
    Host.SendKeys "<Enter>"
    Host.WaitReady 10, 0

    Dim Buf As String
    Dim Cmmt As String
    Dim Trans As String
    Dim found As Boolean
    Dim startRow As Integer
    startRow = 6


    Do
        found = False

        For i = startRow To 20
            Host.ReadScreen Buf, 6, i, 13
            If InStr(Buf, "SOMETHING") > 0 Then
                Host.ReadScreen Cmmt, 8, i, 24
                If Cmmt = "SOMETHING" Or Cmmt = "SOMETHING" Then
                    Host.ReadScreen Trans, 14, i, 65
                    currentCell.Offset(0, 7).Value = "SOMETHING" & Trans
                    Host.SendKey "<PF12>"
                    Exit Do
                Else
            
                    startRow = i + 1
                    found = True
                    Exit For
                End If
            End If
        Next i
        If Not found Then
            Host.SendKey "<PF12>"
            Exit Do
        End If
    Loop While found
End Sub

