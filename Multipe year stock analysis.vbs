Sub WallStreet():

    Dim LastRow As Long
    Dim YearlyRange As Range
    Dim Condition1, Condition2 As FormatCondition
    

    'Creating Loop to go through each worksheet in workbook
    For Each ws In Worksheets

       'Delete all previous data in result table
        Worksheets(ws.Name).Columns("I:L").ClearContents
        
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'Creating Header row for result table
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
        Dim TotalStock As LongLong
        Dim ResultIndex As Integer
        ResultIndex = 2
        Dim CloseValue As Double
        Dim OpenValue As Double
        Dim YearlyChange As Double
        Dim PercentChange As String
        
        'Stating first open Value for each sheet
        If ws.Name = 2016 Then
            OpenValue = 41.81
        ElseIf ws.Name = 2015 Then
            OpenValue = 40.94
        ElseIf ws.Name = 2014 Then
            OpenValue = 57.19
        End If
        
        'Creating Loop to go through base table
        For i = 2 To LastRow
            
            'Check if the ticker of one row is as same as the next one or not
            If ws.Cells(i, 1).Value = ws.Cells(i + 1, 1).Value Then
                TotalStock = TotalStock + ws.Cells(i, 7).Value
                ws.Cells(ResultIndex, 9).Value = ws.Cells(i, 1).Value
                ws.Cells(ResultIndex, 12).Value = TotalStock
                
            
             ElseIf ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                 TotalStock = TotalStock + ws.Cells(i, 7).Value
                 CloseValue = ws.Cells(i, 6).Value
                 YearlyChange = CloseValue - OpenValue
                    If OpenValue = 0 Then
                        PercentChange = 0
                    Else
                        PercentChange = FormatPercent((YearlyChange / OpenValue), 2)
                    End If
                 ws.Cells(ResultIndex, 12).Value = TotalStock
                 ws.Cells(ResultIndex, 10).Value = YearlyChange
                 ws.Cells(ResultIndex, 11).Value = PercentChange
                 
                'Make positive yearly change, green and negative yearly change, red
                Set YearlyRange = ws.Range("J2: J5")
                YearlyRange.FormatConditions.Delete
                Set Condition1 = YearlyRange.FormatConditions.Add(xlCellValue, xlGreaterEqual, "=0")
                Set Condition2 = YearlyRange.FormatConditions.Add(xlCellValue, xlLess, "=0")
                With Condition1.Interior.Color = vbGreen
                End With
                With Condition2.Interior.Color = vbRed
                End With
                    
                    If YearlyChange >= 0 Then
                        ws.Cells(ResultIndex, 10).Interior.Color = vbGreen
                    Else
                        ws.Cells(ResultIndex, 10).Interior.Color = vbRed
                    End If

                 
                 'Reseting Total Stock and set next Index and open value for next unique ticker
                 ResultIndex = ResultIndex + 1
                 TotalStock = 0
                 OpenValue = ws.Cells(i + 1, 3).Value
             End If
                 
            

        Next i
    
    Next ws
    
End Sub

'Challenge
Sub WallStreet2():
    Dim LastRow As Long
    
      
    'Creating Loop to go through each worksheet in workbook
    For Each ws In Worksheets
        
        'Delete all previous data in result 2 table
        Worksheets(ws.Name).Columns("O:Q").ClearContents
        LastRow = ws.Cells(Rows.Count, 10).End(xlUp).Row
        
        'Creating Header row for result 2 table
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        
        ws.Range("Q2") = 0
        ws.Range("Q3") = 0
        
        
        'Creating loop to go through each row of the result 1 table
        For i = 2 To LastRow
            
            'finding Maximum increse and decrese in percent change
            If ws.Cells(i, 11).Value > ws.Cells(i + 1, 11).Value And ws.Cells(i, 11).Value > ws.Cells(2, 17).Value Then
            
                ws.Cells(2, 17).Value = FormatPercent((ws.Cells(i, 11).Value), 2)
                ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
                
            ElseIf ws.Cells(i, 11).Value < ws.Cells(i + 1, 11).Value And ws.Cells(i, 11).Value < ws.Cells(3, 17).Value Then
                ws.Cells(3, 17).Value = FormatPercent((ws.Cells(i, 11).Value), 2)
                ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
            End If
                
            If ws.Cells(i, 12).Value > ws.Cells(i + 1, 12).Value And ws.Cells(i, 12).Value > ws.Cells(4, 17).Value Then
                ws.Cells(4, 17).Value = ws.Cells(i, 12).Value
                ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
            End If
                
                
            
        Next i
        
    
    Next ws

End Sub

