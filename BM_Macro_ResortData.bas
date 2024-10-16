Sub resort_data()

' --------------------------------------
' DESCRIPTION: Resort data in each sheet
' alphabetically and chronologically
' --------------------------------------

' loop through each sheet
For Each ws In Worksheets
    
    ' skip Instructions sheet
    If ws.Name <> "Instructions" Then
    
        ' get last row
        LastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        
        ' clear any existing sorting
        ws.Sort.SortFields.Clear
        
        ' create sort condtions to organize data by ticker and then date
        ws.Sort.SortFields.Add Key:=ws.Range("A2:A" & LastRow), Order:=xlAscending
        ws.Sort.SortFields.Add Key:=ws.Range("B2:B" & LastRow), Order:=xlAscending
        
        ' sort entire data range
        With ws.Sort
            .SetRange ws.Range("A1:G" & LastRow)  ' entire data range
            .Header = xlYes  ' first row has headers
            .MatchCase = False  ' not case sensitive
            .Orientation = xlTopToBottom  ' sort top to bottom
            .Apply  ' applying previously defined sort conditions
        End With
    
    End If

Next ws


End Sub
