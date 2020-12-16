Attribute VB_Name = "Module1"
Sub stock_analysis()
' Declaring variables
    Dim open_value As Double
    Dim close_value As Double
    Dim count As Integer
    Dim volume As Double
    'Dim great As Double
    'Dim low As Double
    Dim max As Double
    Dim min As Double
    Dim tmax As String
    Dim tmin As String
    Dim vol_max As Double
    Dim vmax As String
    Dim first_val(100000) As Double
    Dim ws As Worksheet
    max = -1000000
    min = 10000000
    vol_max = 0
    
    
For Each ws In ThisWorkbook.Worksheets

    

            max = -1000
            min = 10000
            vol_max = 0

    
        ' copying the ticker into new row and removing duplicates
        lrow1 = ws.Cells(Rows.count, 1).End(xlUp).Row
               
        
    With ws.Sort
    'sorting data first on the basis of ticker name and then dates
    
    
        .SortFields.Add Key:=ws.Range("A1"), Order:=xlAscending
        .SortFields.Add Key:=ws.Range("B1"), Order:=xlAscending
        .SetRange ws.Range("A1:G" & lrow1)
        .Header = xlYes
        .Apply
    
    End With
        
         'headers and formatting
            ws.Cells(1, 9).Value = "Ticker"
            ws.Cells(1, 10).Value = "Yearly Change"
            ws.Cells(1, 11).Value = "Percentage Change"
            ws.Cells(1, 12).Value = "Stock Volume"
            
            ws.Cells(1, 16).Value = "Ticker"
            ws.Cells(1, 17).Value = "Value"
            ws.Range("K:K").NumberFormat = "0.00%"
            ws.Range("L:L").NumberFormat = "#,##.0"
            ws.Range("Q2").NumberFormat = "0.00%"
            ws.Range("Q3").NumberFormat = "0.00%"
            ws.Range("Q4").NumberFormat = "#,##.0"
        ' analysing each stock one by one
        
    rcount = 2
            For i = 2 To lrow1
               ticker = ws.Cells(i, 1).Value
                
                
                
        
                    If ticker <> ws.Cells((i + 1), 1).Value Then
                        open_value = first_val(0)
                        close_value = ws.Cells(i, 6).Value
                        yearly = close_value - open_value
                        ws.Range("I" & rcount).Value = ticker
                        ws.Range("J" & rcount).Value = yearly
                        ws.Range("L" & rcount).Value = volume
                        
                        
                        If first_val(0) <> 0 Then
                            ws.Range("K" & rcount).Value = yearly / first_val(0) 'percentage change
                             
                        Else
                            ws.Range("K" & rcount).Value = "Error"
    
                        End If
                        rcount = rcount + 1
                        count = 0
                        volume = 0
                        
                    
                    Else
                    volume = volume + ws.Cells(i, 7).Value
                    first_val(count) = ws.Cells(i, 3).Value
                    count = count + 1
                    
                    
                    End If
                    
                Next i
            
                    
                        
                  'conditional formatting
            lrow2 = ws.Cells(Rows.count, 9).End(xlUp).Row
            For i = 2 To lrow2
            If Cells(i, 10) <> "" Then
            
                If ws.Cells(i, 10).Value >= 0 Then
                   ws.Cells(i, 10).Interior.Color = RGB(0, 255, 0)
                Else
                   ws.Cells(i, 10).Interior.Color = RGB(255, 0, 0)
                End If
            End If
            
            If ws.Cells(i, 11).Value <> "Error" Then
            
            
            
    
                If ws.Cells(i, 11).Value > max Then 'finding the max % increase, minimum % increase and max volume with corresponding tickers
                     max = ws.Cells(i, 11).Value
                     tmax = ws.Cells(i, 9).Value
                 End If
                 If ws.Cells(i, 11).Value < min Then
                     min = ws.Cells(i, 11).Value
                     tmin = ws.Cells(i, 9).Value
                 End If
                 If ws.Cells(i, 12).Value > vol_max Then
                     vol_max = ws.Cells(i, 12).Value
                     vmax = ws.Cells(i, 9).Value
                 End If
            End If
            
            Next i
    
    
    
            ws.Cells(2, 15).Value = "Greatest % increase"
            ws.Cells(2, 16).Value = tmax
            ws.Cells(2, 17).Value = max
    
            ws.Cells(3, 15).Value = "Greatest % Decrease"
            ws.Cells(3, 16).Value = tmin
            ws.Cells(3, 17).Value = min
    
            ws.Cells(4, 15).Value = "Greatest Total Volume"
            ws.Cells(4, 16).Value = vmax
            ws.Cells(4, 17).Value = vol_max
    
    ws.Columns("A:Q").AutoFit

 Next ws


End Sub





