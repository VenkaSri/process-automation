' This module is for calling zipments API to check for release

Sub CallAPI(currentCell As Range)
    Dim xhr As Object
    Dim url As String
    Dim apiKey As String
    Dim parsNumber As String
    Dim response As String
    Dim Json As Object
    Dim currentStatusType As Variant
    Dim transNumber As Variant
    Dim relCode As Variant
    Dim portCode As Variant
    Dim statusStr As Variant
    

    Set xhr = CreateObject("MSXML2.XMLHTTP")
    apiKey = <REDACTED>
    parsNumber = "2314S4" & currentCell.Value
    url = "https://api.zipments.io/pars/" & parsNumber

    xhr.Open "GET", url, False
    xhr.setRequestHeader "Authorization", apiKey
    xhr.Send

    If xhr.status = 200 Then
        response = xhr.responseText

        On Error Resume Next
        Set Json = JsonConverter.ParseJson(response)
        If Err.Number <> 0 Then
            MsgBox "Error parsing JSON: " & Err.description
            On Error GoTo 0
            Set currentCell = currentCell.Offset(1, 0)
            Exit Sub
        End If
        On Error GoTo 0

        currentStatusType = ""

        If Not Json Is Nothing Then
            If Json.Exists("data") Then
                If Json("data").Exists("parsData") Then
                    If Json("data")("parsData").Exists("currentStatusType") Then
                        currentStatusType = Json("data")("parsData")("currentStatusType")
                    End If
                    
                    If Json("data")("parsData").Exists("transactionNumber") Then
                        transNumber = Json("data")("parsData")("transactionNumber")
                    End If
                    
                    If Json("data")("parsData").Exists("releaseCode") Then
                        relCode = Json("data")("parsData")("releaseCode")
                    End If
                    If Json("data")("parsData").Exists("portCode") Then
                        portCode = Json("data")("parsData")("portCode")
                    End If
                    If Json("data")("parsData").Exists("statusString") Then
                        statusStr = Json("data")("parsData")("portCode")
                    End If
                Else
                    MsgBox "parsData key not found"
                End If
            Else
                MsgBox "data key not found"
            End If
        Else
            MsgBox "JSON response is empty"
        End If
        If currentStatusType = "released" Then
            currentCell.Offset(0, 2).Value = "Released (PARS TRACKER)"
            currentCell.Offset(0, 7).Value = transNumber & "At port: " & portCode
            currentCell.Interior.Color = RGB(245, 245, 245)
        End If
    Else
        MsgBox "Error: " & xhr.status & " - " & xhr.statusText
    End If

    Set xhr = Nothing
End Sub


