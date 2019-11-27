Sub StockResults()
   'Declare the Workbook Variables for Loop
    Dim ws As Integer
    Dim ws_num As Integer
    Dim starting_ws As Worksheet
    Dim Summary_Table As Integer
    Dim Ticker As String
    Dim Yearly As Double
    Dim Volume As Double
    Dim Max As Double
    Dim Min As Double
    Dim Vol As Double
    Dim LastRow As Long
    Dim year_open As Double
    Dim year_close As Double
        Set starting_ws = ActiveSheet
        ws_num = ThisWorkbook.Worksheets.Count
        Summary_Table = 2

    'Label the data table
    Cells(1, "I").Value = "Ticker"
    Cells(1, "J").Value = "Yearly Change"
    Cells(1, "K").Value = "Percent Change"
    Cells(1, "L").Value = "Total Stock Volume"
    Cells(2, "O").Value = "Greatest % Increase"
    Cells(3, "O").Value = "Greatest % Decrease"
    Cells(4, "O").Value = "Greatest Total Volume"
    Cells(1, "P").Value = "Ticker"
    Cells(1, "Q").Value = "Value"
    
   For ws = 1 To ws_num
        ThisWorkbook.Worksheets(ws).Activate
            
            'Identify where to find the Last Row
            LastRow = Cells(Rows.Count, "A").End(xlUp).Row
        
        For i = 2 To LastRow:
            
            If Cells(i - 1, "A").Value <> Cells(i, "A").Value Then
                year_open = Cells(i, "C").Value
                Ticker = Cells(i, 1).Value
                    
            End If
            
            If Cells(i + 1, "A").Value <> Cells(i, "A").Value Then
                
                year_close = Cells(i, "F").Value
  
                    If year_open = 0 Then
                    Range("K" & Summary_Table).Value = 0
                
                    Else
                    Range("K" & Summary_Table).Value = (year_close - year_open) / year_open
                    End If
            
            Range("K2:K" & LastRow).NumberFormat = "0.00%"
            
            Range("J" & Summary_Table).Value = year_close - year_open
          
            Volume = Volume + Cells(i, 7).Value
                
            Range("I" & Summary_Table).Value = Ticker
                
            Range("L" & Summary_Table).Value = Volume
            
            If Range("J" & Summary_Table).Value >= 0 Then
                Range("J" & Summary_Table).Interior.ColorIndex = 4
            Else:
            Range("J" & Summary_Table).Interior.ColorIndex = 3
            End If
            Range("J2:J" & LastRow).NumberFormat = "0.00"
                
            Summary_Table = Summary_Table + 1

            Volume = 0
        
        Else
             Volume = Volume + Cells(i, 7).Value
    End If
    Next i
  
        Max = WorksheetFunction.Max(Range("K:K").Value)
        Cells(2, "Q").Value = Max
        
        Min = WorksheetFunction.Min(Range("K:K").Value)
        Cells(3, "Q").Value = Min
        
        Range("Q2:Q3").NumberFormat = "0.00%"
        
        Vol = WorksheetFunction.Max(Range("L:L").Value)
        Cells(4, "Q").Value = Vol
        
        
        For i = 2 To LastRow
            
            If Cells(i, "K").Value = Max Then
                Cells(2, "P").Value = Cells(i, "I").Value
            End If
            
            If Cells(i, "K").Value = Min Then
                Cells(3, "P").Value = Cells(i, "I").Value
            End If
            
            If Cells(i, "L").Value = Vol Then
                Cells(4, "P").Value = Cells(i, "I").Value
            End If
    Next i

Next ws
    starting_ws.Activate
    
End Sub

