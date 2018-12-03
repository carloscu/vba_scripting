
'Steps:
'
'
'Create a script that will loop through each year of stock data and grab the total amount of volume each stock had over the year.

Sub Multiple_year_stock_data_easy()

    ' Initiate variables
    Dim ticker As String
    Dim tally_stock As Double
    Dim counter As Integer
    
    tally_stock = 0
    counter = 2
    
    'Identify Last Row
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row

    ' Loop thru worksheet identify the ticker stock and tally volume
    For i = 2 To LastRow
        
        'check for ticker symbol and tally volume
        If Cells(i + 1, 1).Value <> Cells(i, 1) Then '= ticker Then
                    
            'Set ticker & tally stock volume
            ticker = Cells(i, 1).Value
            tally_stock = tally_stock + Cells(i, 7).Value
             
            'Print the ticker name and tally amount
            Range("J" & counter).Value = ticker
            Range("K" & counter).Value = tally_stock
            
            'Add one to summary table
            counter = counter + 1
            
            'Reset tally amount
            tally_stock = 0
        
        Else
            tally_stock = tally_stock + Cells(i, 7).Value
            
        End If
  
    Next i
    
End Sub
