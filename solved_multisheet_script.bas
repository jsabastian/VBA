Attribute VB_Name = "Module6"
Option Explicit




Sub tickerdata_all_ws()

    
    'define variables
        Dim ws As Worksheet
        Dim ticker As String
        Dim stock_vol As Long
        Dim yrclose As Double
        Dim yrchange As Double
        Dim yrvar As Double
        Dim i As Long
        Dim sumrow As Integer
        Dim lastrow As Long

        lastrow = ActiveSheet.UsedRange.Rows.Count
    For Each ws In Worksheets
        'create the column headers
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"

        'which row our summaries will be placed for above columns
        sumrow = 2

            'the loop checks each iteration until the last row
            
        For i = 2 To lastrow
                
                'we need to capture the price of the ticker if it is the first of its year
            Dim firstprice As Boolean
                
            If firstprice = False Then 'false is the default boolean value, so this statement is true
                    
                Dim yropen As Double
                        
                    yropen = ws.Cells(i, 3).Value
                        
                    firstprice = True 'we have captured the opening price of the year for the ticker
                    
            End If
                    
                    'now we can check if we are in the same ticker value
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                    
                    'this should happen when the cell values are finally different / capture all the values
                    ticker = ws.Cells(i, 1).Value
                    
                    stock_vol = ws.Cells(i, 7).Value
                    
                    yrclose = ws.Cells(i, 6).Value
                    
                    yrchange = yrclose - yropen
                    
                If yropen <> 0 Then 'this prevents dividing by zero which will result in overflow error 6
                    
                            yrvar = (yrclose - yropen) / yrclose
                    
                Else
                        
                            yrvar = 0
                            yrchange = 0
                        
                End If
                    

                'insert values into the summary
                ws.Cells(sumrow, 9).Value = ticker
                ws.Cells(sumrow, 10).Value = yrchange
                ws.Cells(sumrow, 11).Value = yrvar
                ws.Cells(sumrow, 12).Value = stock_vol
                sumrow = sumrow + 1 'sets the stage for the next set of data into row 3

                stock_vol = 0 'resets vol for the next ticker
                    
                firstprice = False 'allows the next 'first' open price of the loop to be captured
                
            End If

        
        Next i  'finish i iteration of the loop
            


        ws.Range("K:K").NumberFormat = "0.0%" 'aesthetic preference


        'format columns colors
        Dim colJ As Range
        Dim Cell As Range

        Set colJ = ws.Range("J2", ws.Range("J2").End(xlDown)) 'from J2 to the last cell entry
            
        For Each Cell In colJ
                
            If Cell.Value > 0 Then
                    Cell.Interior.ColorIndex = 50
                    Cell.Font.ColorIndex = 2
            ElseIf Cell.Value < 0 Then
                    Cell.Interior.ColorIndex = 30
                    Cell.Font.ColorIndex = 2
            Else
                    Cell.Interior.ColorIndex = xlNone 'this really serves no purpose, just to show elseif use
            End If
                
        Next
    Next ws
End Sub
