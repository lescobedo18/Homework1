Sub MultYearStock()

'loop through all sheets
For Each ws In Worksheets
    
    'create variable for worksheets
    Dim worksheet_name As String
    
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    worksheet_name = ws.Name
    
    'MsgBox WorksheetName
    
    'variable fot ticker
    Dim ticker As String
    
    'variable for total stock volume
    Dim total_volume As LongLong
    total_volume = 0
    
    'variable for opening price
    Dim opening_price As Double
    opening_price = 0
    
    'Variable for closing price
    Dim closing_price As Double
    closing_price = 0
    
    'variable for yearly change
    Dim yearly_change As Double
    yearly_change = 0
    
    'varible for percent chnage
    Dim percent_change As Double
    percent_change = 0
    
    'place to keep data
    Dim summary As Integer
    summary = 2
    
    'set intial value of opening stock
    opening_price = ws.Cells(2, 3).Value
    
For i = 2 To LastRow
    
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
            ticker = ws.Cells(i, 1).Value
            
            'calculate
            closing_price = ws.Cells(i, 6).Value
            yearly_change = closing_price - opening_price
        
        
            If opening_price <> 0 Then
                percent_change = (yearly_change / opening_price) * 100
        
            End If
            
    'print yearly chnage to summary
    ws.Range("J" & summary).Value = yearly_change
    
    'print percent change to summary
    ws.Range("K" & summary).Value = (CStr(percent_change) & "%")
            
    'add to stock volume
    total_volume = total_volume + ws.Cells(i, 7).Value
    
    'print ticker to summary
    ws.Range("I" & summary).Value = ticker
    
    'print total volume to summary
    
    ws.Range("L" & summary).Value = total_volume
    
    'add one to summary
    summary = summary + 1
    
    'next opening price
    opening_price = ws.Cells(i + 1, 3).Value
    
    'reset
    total_volume = 0
    percent_change = 0
    
    'color fill
        If (yearly_change > 0) Then
            ws.Range("J" & summary).Interior.ColorIndex = 4
            
        ElseIf (yearly_change <= 0) Then
            ws.Range("J" & summary).Interior.ColorIndex = 3
            
        End If
    
    Else
 
    total_volume = total_volume + ws.Cells(i, 7).Value
 
    End If
        Next i
     
     
Next ws

    
End Sub
