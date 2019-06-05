Attribute VB_Name = "Module1"
Sub Easy()

    ' for all sheets
    
Dim WS As Worksheet
    For Each WS In ActiveWorkbook.Worksheets
    WS.Activate
    
        ' To find Last Row
        LastRow = WS.Cells(Rows.Count, 1).End(xlUp).Row

        ' Title for Ticker and Total Vol
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Total Stock Volume"
        
        'Variable and Counters
        Dim Vol As Double
        Vol = 0
        
        Dim Name As String
        
        Dim Row As Double
        
        Row = 2
        
        Dim Column As Integer
        
        Column = 1
        
        Dim i As Long
        
         ' Loop through ticker
        
        For i = 2 To LastRow
        
         ' to check if ticker symbol is different,
            If Cells(i + 1, Column).Value <> Cells(i, Column).Value Then
            
                ' store in Ticker column in column I
                Name = Cells(i, Column).Value
                Cells(Row, Column + 8).Value = Name
                
                ' Add Total Volume in column J
                Vol = Vol + Cells(i, Column + 6).Value
                Cells(Row, Column + 9).Value = Vol
                
                ' Add one to the summary table row to move to next row
                Row = Row + 1
            
                ' reset the Volume Total
                Vol = 0
            'if cells are the same ticker
            Else
                Vol = Vol + Cells(i, Column + 6).Value
            End If
        Next i
        
        
    Next WS
        
End Sub








