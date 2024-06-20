' This module is for calling a generic API to check for release status

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
    apiKey = "Your-API-Key-Here" 
    parsNumber = "GenericPrefix" & currentCell.Value 
    url = "https://api.genericdomain.com/status/" & parsNumber 

    xhr.Open "GET", url, False
    xhr.setRequestHeader "Authorization", "Bearer " & apiKey
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
                If Json("data").Exists("statusData") Then
                    If Json("data")("statusData").Exists("currentStatusType") Then
                        currentStatusType = Json("data")("statusData")("currentStatusType")
                    End If
                    
                    If Json("data")("statusData").Exists("transactionNumber") Then
                        transNumber = Json("data")("statusData")("transactionNumber")
                    End If
                    
                    If Json("data")("statusData").Exists("releaseCode") Then
                        relCode = Json("data")("statusData")("releaseCode")
                    End If
                    
                    If Json("data")("statusData").Exists("portCode") Then
                        portCode = Json("data")("statusData")("portCode")
                    End If
                    
                    If Json("data")("statusData").Exists("statusString") Then
                        statusStr = Json("data")("statusData")("statusString")
                    End If
                Else
                    MsgBox "Status data key not found"
                End If
            Else
                MsgBox "Data key not found"
            End If
        Else
            MsgBox "JSON response is empty"
        End If

        If currentStatusType = "released" Then
            currentCell.Offset(0, 2).Value = "Released (Status Updated)"
            currentCell.Offset(0, 7).Value = transNumber & " at Port: " & portCode
            currentCell.Interior.Color = RGB(245, 245, 245)
        End If
    Else
        MsgBox "Error: " & xhr.status & " - " & xhr.statusText
    End If

    Set xhr = Nothing
End Sub
