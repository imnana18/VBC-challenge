Dim Ticker As String
Dim Total_Stock_Vol As Double
Dim Counter As Long
Dim Open_Value As Double
Dim Close_Value As Double
Dim Result_Counter As Long
Dim lRow As Long

Dim Greatest_Increase As Double
Dim Greatest_Decrease As Double
Dim Greatest_Total_Volumn As Double
Sub ProcessAllSheets()
    Dim ws As Worksheet
    
    ' Loop through all worksheets in the workbook
    For Each ws In ThisWorkbook.Worksheets
        ' Activate the current worksheet
        ws.Activate
        ' Call your Stocks subroutine to process the data on the current sheet
        Stocks
    Next ws
End Sub

Sub Stocks()
    Greatest_Increase = 0
    Greatest_Decrease = 0
    Greatest_Total_Volumn = 0
    
    ' Init the header
    Call initHeader

    ' Find the last used row in column A
    lRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    ' Initialize the values
    Total_Stock_Vol = 0
    Result_Counter = 2
    Counter = 2
    Ticker = Cells(Counter, 1).Value
    Open_Value = Cells(Counter, 3).Value

    For Counter = 2 To lRow
        ' Found a new Ticker or Last date of the Ticker
        If IsEmpty(Cells(Counter, 1).Value) Or Cells(Counter, 1).Value <> Ticker Then
            Call update
            
            ' Finally, update the Open_Value and reset Total_Stock_Vol
            Open_Value = Cells(Counter, 3).Value
            Total_Stock_Vol = 0
            
            Result_Counter = Result_Counter + 1
        End If
        
        ' Update Total_Stock_Vol
        Total_Stock_Vol = Total_Stock_Vol + Cells(Counter, 7).Value
        Ticker = Cells(Counter, 1).Value
    Next Counter
    
    ' Update the last one
    Call update
        
End Sub


Sub update():


    Close_Value = Cells(Counter - 1, 6).Value
        
    ' Store the Ticker
    Cells(Result_Counter, 9).Value = Ticker

    ' Calculate the Yearly change and Percent change and Total stock Volume and store them
    Dim YearlyChange As Double
    Dim YearlyChangePercent As Double

    YearlyChange = Close_Value - Open_Value
    Cells(Result_Counter, 10).Value = Round(YearlyChange, 2)
    If YearlyChange >= 0 Then
        Cells(Result_Counter, 10).Interior.Color = RGB(0, 255, 0) ' Green for positive change
    Else
        Cells(Result_Counter, 10).Interior.Color = RGB(255, 0, 0) ' Red for negative change
    End If
    
    ' Handle the case where Open_Value is zero or too close to zero
    If Abs(Open_Value) < 1E-06 Then
        ' You can choose how to handle this situation, for example, by setting Cells(Result_Counter, 11) to 0
        YearlyChangePercent = 0
    Else
        YearlyChangePercent = YearlyChange / Open_Value
    End If
    Cells(Result_Counter, 11).Value = YearlyChangePercent
    
    'Update the Greatest_Increase
    If YearlyChangePercent > Greatest_Increase Then
        Range("P2").Value = Ticker
        Range("Q2").Value = YearlyChangePercent
        Range("Q2").NumberFormat = "0.00%"
        Greatest_Increase = YearlyChangePercent
    End If

    'Update the Greatest_Decrease
    If YearlyChangePercent < Greatest_Decrease Then
        Range("P3").Value = Ticker
        Range("Q3").Value = YearlyChangePercent
        Range("Q3").NumberFormat = "0.00%"
        Greatest_Decrease = YearlyChangePercent
    End If
    
    Cells(Result_Counter, 11).NumberFormat = "0.00%"
    Cells(Result_Counter, 12).Value = Total_Stock_Vol


    'Update the Greatest_Total_Volumn
    If Total_Stock_Vol > Greatest_Total_Volumn Then
        Range("P4").Value = Ticker
        Range("Q4").Value = Total_Stock_Vol
        
        Greatest_Total_Volumn = Total_Stock_Vol
    End If
End Sub

Sub initHeader()
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
    
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volumn"
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
End Sub