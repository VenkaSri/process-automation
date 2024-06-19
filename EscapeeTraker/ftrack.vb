' This finds the up to date location of the package, using FTRACK web application via Selenium

Sub FTRACK()
    On Error GoTo ErrHandler
    
    Dim driver As New WebDriver
    Dim wb As Workbook
    Set wb = Workbooks("EscapeeTracker.xlsm")
    Dim ws As Worksheet
    Set ws = wb.Sheets("Sheet1")
    Dim wsScans As Worksheet
    Set wsScans = wb.Sheets("Scans")
    Dim currentCell As Range
    Set currentCell = ws.Range("A4")
    Dim nextRow As Integer
    nextRow = 2

    'driver.AddArgument "--headless"
    'driver.AddArgument "--disable-gpu"
    'driver.AddArgument "--window-size=1920,1080"
    driver.Start "chrome", ""
    driver.Get "http://ftrack.gss.ground.fedex.com:8085/cgi-bin/PKG621CL?IIWEBI="

    frmProgress.Show vbModeless
    
    Do While Not IsEmpty(currentCell.Value)
        If currentCell.Interior.Color <> RGB(245, 245, 245) Then
            frmProgress.lblRow.Caption = "Processing Row: " & currentCell.row
            ProcessCell driver, currentCell, wsScans, nextRow
            driver.ExecuteScript "window.history.back();"
        End If
        Set currentCell = currentCell.Offset(1, 0)
    Loop

    GoTo CleanUp

ErrHandler:
    MsgBox "An error occurred: " & Err.description
    Resume CleanUp

CleanUp:
    Unload frmProgress
    driver.Quit
    Set driver = Nothing
    Exit Sub
End Sub



Function CleanText(text As String) As String
    text = Trim(text)
    text = Replace(text, Chr(160), " ")
    text = WorksheetFunction.Trim(text)
    CleanText = text
End Function


Sub ProcessCell(driver As WebDriver, currentCell As Range, wsScans As Worksheet, ByRef nextRow As Integer)
    On Error GoTo ErrHandler
    

    FillForm driver, currentCell.Offset(0, 1).Value
    

    If Not WaitForForm(driver) Then
        MsgBox "Form not loaded properly."
        Exit Sub
    End If
    

    ProcessTables driver, currentCell, wsScans, nextRow

    Exit Sub

ErrHandler:
    MsgBox "An error occurred in ProcessCell: " & Err.description
    Resume Next
End Sub

Sub FillForm(driver As WebDriver, inputValue As String)
    On Error GoTo ErrHandler
    
    Dim inputField As Object
    Set inputField = driver.FindElementByName("iitraks", timeout:=1000)
    inputField.Clear
    inputField.SendKeys inputValue
    driver.FindElementByName("submit", timeout:=1000).Click
    
    Exit Sub

ErrHandler:
    MsgBox "An error occurred in FillForm: " & Err.description
    Resume Next
End Sub
Function WaitForForm(driver As WebDriver) As Boolean
    On Error GoTo ErrHandler
    
    Dim endTime As Date
    endTime = DateAdd("s", 10, Now)
    
    While Now < endTime
        On Error Resume Next
        Dim form As Object
        Set form = driver.FindElementByName("form1", timeout:=1000)
        If Not form Is Nothing Then
            WaitForForm = True
            Exit Function
        End If
        On Error GoTo 0
        DoEvents '
    Wend
    
    WaitForForm = False
    Exit Function

ErrHandler:
    MsgBox "An error occurred in WaitForForm: " & Err.description
    WaitForForm = False
End Function

Sub ProcessTables(driver As WebDriver, currentCell As Range, wsScans As Worksheet, ByRef nextRow As Integer)
    On Error GoTo ErrHandler
    
    Dim tables As Object
    Set tables = driver.FindElementByName("form1").FindElementsByTag("table")
    
    If tables.Count >= 2 Then
        
        If Not VerifyFirstTable(tables.Item(1), currentCell.Offset(0, 1).Value) Then
            Exit Sub
        End If
        
        
        ProcessSecondTable tables.Item(2), currentCell, wsScans, nextRow
    Else
        MsgBox "Expected number of tables not found."
    End If
    
    Exit Sub

ErrHandler:
    MsgBox "An error occurred in ProcessTables: " & Err.description
    Resume Next
End Sub
Function VerifyFirstTable(firstTable As Object, expectedTrackId As String) As Boolean
    On Error GoTo ErrHandler
    
    Dim firstTableRows As Object
    Set firstTableRows = firstTable.FindElementsByTag("tr")
    
    If firstTableRows.Count > 0 Then
        Dim firstRow As Object
        Set firstRow = firstTableRows.Item(1)
        Dim tds As Object
        Set tds = firstRow.FindElementsByTag("td")
        
        If tds.Count > 0 Then
            Dim trackId As String
            trackId = Trim(tds.Item(1).text)
            trackId = Replace(trackId, "Track Id:", "")
            trackId = Trim(Replace(trackId, "&nbsp;", ""))
            
            
            VerifyFirstTable = (trackId = CleanText(expectedTrackId))
        Else
            MsgBox "No td elements found in the first row."
            VerifyFirstTable = False
        End If
    Else
        MsgBox "No rows found in the first table."
        VerifyFirstTable = False
    End If
    
    Exit Function

ErrHandler:
    MsgBox "An error occurred in VerifyFirstTable: " & Err.description
    VerifyFirstTable = False
End Function


Sub ProcessSecondTable(secondTable As Object, currentCell As Range, wsScans As Worksheet, ByRef nextRow As Integer)
    On Error GoTo ErrHandler
    
    Dim rows As Object
    Set rows = secondTable.FindElementsByTag("tr")
    Dim locationColumn As Integer
    locationColumn = 4
    Dim collectScans As Boolean
    collectScans = False
    Dim found6124TOROA As Boolean
    found6124TOROA = False
    Dim lastRow As Object
    Dim statusColumn As Integer, scanTypeColumn As Integer
    scanTypeColumn = 6
    statusColumn = 7
    Dim row As Object, cells As Object
    For Each row In rows
        Set cells = row.FindElementsByTag("td")
        If cells.Count >= statusColumn Then
            If Not collectScans And cells.Item(locationColumn).text = "6124-TORO-A" Then
                collectScans = True
                found6124TOROA = True
            End If

            If collectScans Then
      
                ProcessRowData cells, currentCell, wsScans, nextRow, locationColumn, statusColumn, scanTypeColumn
            End If
        End If
        Set lastRow = cells
    Next row

   
    If Not found6124TOROA And Not lastRow Is Nothing Then
        ProcessRowData lastRow, currentCell, wsScans, nextRow, locationColumn, statusColumn, scanTypeColumn
    End If

    Exit Sub

ErrHandler:
    MsgBox "An error occurred in ProcessSecondTable: " & Err.description
    Resume Next
End Sub

Sub ProcessRowData(cells As Object, currentCell As Range, wsScans As Worksheet, ByRef nextRow As Integer, locationColumn As Integer, statusColumn As Integer, scanTypeColumn As Integer)
    Dim locationText As String
    locationText = ExtractNumbers(cells.Item(locationColumn).text)
    
    Dim scanTypeText As String
    scanTypeText = IIf(cells.Item(scanTypeColumn).text = "-", "", cells.Item(scanTypeColumn).text)

    Dim statusCodeText As String
    Dim statusTitle As String
    If Not IsEmpty(cells.Item(statusColumn).Attribute("title")) Then
        statusTitle = cells.Item(statusColumn).Attribute("title")
        If statusTitle <> "" Then
            statusCodeText = CleanStatusCode(cells.Item(statusColumn).text) & " (" & statusTitle & ")"
        Else
            statusCodeText = CleanStatusCode(cells.Item(statusColumn).text)
        End If
    Else
        statusCodeText = CleanStatusCode(cells.Item(statusColumn).text)
    End If
    
    wsScans.cells(nextRow, 1).Value = currentCell.Value
    wsScans.cells(nextRow, 2).Value = currentCell.Offset(0, 1).Value
    wsScans.cells(nextRow, 3).Value = cells.Item(1).text
    wsScans.cells(nextRow, 4).Value = cells.Item(2).text
    wsScans.cells(nextRow, 5).Value = locationText
    wsScans.cells(nextRow, 6).Value = scanTypeText
    wsScans.cells(nextRow, 7).Value = statusCodeText
    nextRow = nextRow + 1
End Sub

Function ExtractNumbers(str As String) As String
    Dim output As String
    Dim i As Integer

    
    If Not Mid(str, 1, 1) Like "[0-9]" Then
        ExtractNumbers = str
        Exit Function
    End If

   
    For i = 1 To Len(str)
        If Mid(str, i, 1) Like "[0-9]" Then
            output = output & Mid(str, i, 1)
        End If
    Next i

    ExtractNumbers = output
End Function


Function CleanStatusCode(statusCode As String) As String
    Dim numericValue As String
    Dim i As Integer
    For i = 1 To Len(statusCode)
        If Mid(statusCode, i, 1) Like "[0-9]" Then
            numericValue = numericValue & Mid(statusCode, i, 1)
        Else
            Exit For
        End If
    Next i
    CleanStatusCode = numericValue
End Function