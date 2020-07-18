Attribute VB_Name = "Module1"
Sub tickerdata()

'define variables
Dim ticker As String
Dim stock_vol As Long
Dim yrclose As Double
'not defining the year_open yet
Dim yearly_change As Double
Dim variation As Double
Dim i As Long

Dim Summary_Table_Row As Integer

Dim lastrow As Long

lastrow = ActiveSheet.UsedRange.Rows.Count

'create the column headers
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"

    'which row our summaries will be placed for above columns
    Summary_Table_Row = 2

    'the loop checks each iteration until the last row
    
        For i = 2 To lastrow
        
        'we need to capture the price of the ticker if it is the first of its year
        Dim firstprice As Boolean
        
            If firstprice = False Then 'false is the default boolean value, so this statement is true
            
                Dim yropen As Double
                
                yropen = Cells(i, 3).Value
                
                firstprice = True 'we have captured the opening price of the year for the ticker
            
            End If
            
            'now we can check if we are in the same ticker value
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            
            'this should happen when the cell values are finally different / capture all the values
            ticker = Cells(i, 1).Value
            
            stock_vol = Cells(i, 7).Value
            
            yrclose = Cells(i, 6).Value
            
            yrchange = yrclose - yropen
            
                If yropen <> 0 Then 'this prevents dividing by zero which will result in overflow error 6
            
                     variation = (yrclose - yropen) / yrclose
            
                Else
                
                    variation = 0
                    yrchange = 0
                
                End If
            

            'insert values into the summary
            Cells(Summary_Table_Row, 9).Value = ticker
            Cells(Summary_Table_Row, 10).Value = yrchange
            Cells(Summary_Table_Row, 11).Value = variation
            Cells(Summary_Table_Row, 12).Value = stock_vol
            Summary_Table_Row = Summary_Table_Row + 1 'sets the stage for the next set of data into row 3

            vol = 0 'resets vol for the next ticker
             
            firstprice = False 'allows the next 'first' open price of the loop to be captured
        
        End If

'finish i iteration of the loop
    Next i
    


Range("K:K").NumberFormat = "0.0%" 'aesthetic preference


    'format columns colors
    Dim colJ As Range
    
    Set colJ = Range("J2", Range("J2").End(xlDown)) 'from J2 to the last cell entry
    
    
    For Each Cell In colJ
        
        If Cell.Value > 0 Then
            Cell.Interior.ColorIndex = 50
            Cell.Font.ColorIndex = 2
        ElseIf Cell.Value < 0 Then
            Cell.Interior.ColorIndex = 30
            Cell.Font.ColorIndex = 2
        Else
            Cell.Interior.ColorIndex = xlNone 'no color
        End If
        
    Next

End Sub
