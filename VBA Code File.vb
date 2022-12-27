Sub Stock_Analyzer()

'Declare Variables & Set Initial Values
Dim Current_Ticker As String
Dim Opening_Price As Double
Dim Closing_Price As Double
Dim Total_Volume As Double
Dim Greatest_Increase As Double
Dim Greatest_Decrease As Double
Dim Max_Volume As Double
Dim Greatest_Increase_Ticker As String
Dim Greatest_Decrease_Ticker As String
Dim Max_Volume_Ticker As String
Dim LastRow As Long
Dim LastEntry As Long

'Loop Through Individual Sheets & Stocks & Print Data To Appropriate Cells
For Each ws In Worksheets
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    LastEntry = ws.Cells(Rows.Count, 9).End(xlUp).Row
    Opening_Price = ws.Cells(2, 3).Value
    Greatest_Increase_Ticker = 0
    Greatest_Decrease_Ticker = 0
    Max_Volume_Ticker = 0
    Greatest_Increase = 0
    Greatest_Decrease = 0
    Max_Volume = 0

    For i = 2 To LastRow
        Total_Volume = Total_Volume + ws.Cells(i, 7).Value
        
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
           LastEntry = ws.Cells(Rows.Count, 9).End(xlUp).Row
           
           Current_Ticker = ws.Cells(i, 1).Value
           Closing_Price = ws.Cells(i, 6).Value
           ws.Cells(LastEntry + 1, 9).Value = Current_Ticker
           ws.Cells(LastEntry + 1, 10).Value = Closing_Price - Opening_Price
           ws.Cells(LastEntry + 1, 11).Value = (Closing_Price / Opening_Price) - 1
           ws.Cells(LastEntry + 1, 12).Value = Total_Volume
           
           Opening_Price = ws.Cells(i + 1, 3).Value
           Total_Volume = 0
        End If
    Next i
    
    'Set Cell Colors
    For n = 2 To LastEntry + 1
        If ws.Cells(n, 10).Value > 0 Then
            ws.Cells(n, 10).Interior.ColorIndex = 4
        Else:
            ws.Cells(n, 10).Interior.ColorIndex = 3
        End If
    Next n
    
    'Find Greatest Outliers
    For Z = 2 To LastEntry + 1
        If ws.Cells(Z, 11).Value > Greatest_Increase Then
            Greatest_Increase_Ticker = ws.Cells(Z, 9).Value
            Greatest_Increase = ws.Cells(Z, 11).Value
        ElseIf ws.Cells(Z, 11).Value < Greatest_Decrease Then
            Greatest_Decrease_Ticker = ws.Cells(Z, 9).Value
            Greatest_Decrease = ws.Cells(Z, 11).Value
        End If
        
        If ws.Cells(Z, 12).Value > Max_Volume Then
            Max_Volume_Ticker = ws.Cells(Z, 9).Value
            Max_Volume = ws.Cells(Z, 12).Value
        End If
    Next Z
    
    ws.Range("P2").Value = Greatest_Increase_Ticker
    ws.Range("P3").Value = Greatest_Decrease_Ticker
    ws.Range("P4").Value = Max_Volume_Ticker
    ws.Range("Q2").Value = Greatest_Increase
    ws.Range("Q3").Value = Greatest_Decrease
    ws.Range("Q4").Value = Max_Volume
    
Next ws


End Sub