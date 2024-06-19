' Clearing exisiting contents from all sheets

Sub ClearAll()
    Set wb = Workbooks("EscapeeTracker.xlsm")
    Dim ws As Worksheet
    Dim wsScans As Worksheet
    Set ws = wb.Sheets("Sheet1")
    Set wsScans = wb.Sheets("scans")
    Dim cell As Range
    

    With ws.Range("A4:N1000")
        .ClearContents
        .ClearFormats
        
        
        For Each cell In .cells
            If Not cell.Comment Is Nothing Then
                cell.Comment.Delete
            End If
        Next cell
    End With
    
    With wsScans.Range("A2:G8000")
        .ClearContents
        .ClearFormats
        
        
        For Each cell In .cells
            If Not cell.Comment Is Nothing Then
                cell.Comment.Delete
            End If
        Next cell
    End With
End Sub
