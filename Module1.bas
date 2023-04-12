Attribute VB_Name = "Module1"

Sub stocksummary()


Dim ws As Worksheet
Dim starting_ws As Worksheet
For Each ws In ThisWorkbook.Worksheets
    ws.Activate
    
Dim Ticker_Symbol As String
Dim Total_Volume As Double
Total_Volume = 0
Dim Last_Row As Double
Last_Row = Range("A1").End(xlDown).Row
Dim First_Open As Double
First_Open = Cells(2, 3).Value
Dim Last_Close As Double



'Summary Table Headers and row initialization
Dim Summary_Row As Integer
Summary_Row = 2
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"

'looping through the data

For i = 2 To Last_Row
    
    'if different, print and reset
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        Ticker_Symbol = Cells(i, 1).Value
        Total_Volume = Total_Volume + Cells(i, 7).Value
        Last_Close = Cells(i, 6).Value
        
        'Write to Summary table part
        Range("I" & Summary_Row).Value = Ticker_Symbol
        Range("L" & Summary_Row).Value = Total_Volume
        Range("J" & Summary_Row).Value = Last_Close - First_Open
        Percent_Change = (Last_Close - First_Open) / First_Open
        Percent_Change = Format(Percent_Change, "0.00%")
        Range("K" & Summary_Row).Value = Percent_Change
        
        'conditional coloring
         If Range("J" & Summary_Row).Value < 0 Then
            Cells(Summary_Row, 10).Resize(, 2).Interior.ColorIndex = 3
            ElseIf Range("J" & Summary_Row).Value > 0 Then
            Cells(Summary_Row, 10).Resize(, 2).Interior.ColorIndex = 4
         End If
         
        'reset First_Open to correspond with next ticker if not blank
        If Cells(i + 1, 3) <> 0 Then
        First_Open = Cells(i + 1, 3).Value
        End If
        Summary_Row = Summary_Row + 1
        Total_Volume = 0
        
    Else
        'if same, keep totaling
        Total_Volume = Total_Volume + Cells(i, 7).Value
    

End If
Next i



'Other little summary table
Cells(2, 15).Value = "Greatest % Increase"
Cells(3, 15).Value = "Greatest % Decrease"
Cells(4, 15).Value = "Greatest Total Volume"
Cells(1, 16).Value = "Ticker"
Cells(1, 17).Value = "Value"
Dim TickerInc As String
Dim TickerDec As String
Dim TickerVol As String


Greatest_Increase = Cells(2, 11).Value
Greatest_Decrease = Cells(2, 11).Value
Greatest_Volume = Cells(2, 12).Value
    
For j = 2 To Summary_Row
    
    
    If Cells(j, 11).Value > Greatest_Increase Then
        Greatest_Increase = Cells(j, 11).Value
        TickerInc = Cells(j, 9).Value
    End If
    
    If Cells(j, 11).Value < Greatest_Decrease Then
        Greatest_Decrease = Cells(j, 11).Value
        TickerDec = Cells(j, 9).Value
    End If
    
    If Cells(j, 12) > Greatest_Volume Then
        Greatest_Volume = Cells(j, 12).Value
        TickerVol = Cells(j, 9).Value
    End If
    
Next j
 Greatest_Increase = Format(Greatest_Increase, "0.00%")
 Greatest_Decrease = Format(Greatest_Decrease, "0.00%")
Cells(2, 17).Value = Greatest_Increase
Cells(3, 17).Value = Greatest_Decrease
Cells(4, 17).Value = Greatest_Volume
Cells(2, 16).Value = TickerInc
Cells(3, 16).Value = TickerDec
Cells(4, 16).Value = TickerVol

Next ws

End Sub

