Attribute VB_Name = "Module3"
Sub testing()

' Set an initial variable for holding the ticker name
 Dim Ticker_Name As String

' Set an initial variable for holding the total per ticker name
Dim Year_Total As Double
Year_Total = 0

' Set an initial variable for holding the yearly percent change per ticker name
Dim Year_Percent As Double
Year_Percent = 0

' Set an initial variable for holding the total stock volume per ticker name
Dim Total_Volume As Double
Total_Volume = 0


' Keep track of the location for each ticker name in the summary table
Dim Summary_Year_Row As Integer
Summary_Year_Row = 2

' Loop through all ticker values
For i = 2 To Range("L2").End(xlDown).Row

' Check if we are still within the same ticker name, if it is not...
If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

' Set the ticker name
Ticker_Name = Cells(i, 1).Value

' Add to the ticker Total
Year_Total = Year_Total + Cells(i, 6).Value - Cells(i, 3).Value

' Add to the percentage change
Year_Percent = (Year_Total + Cells(i, 6).Value - Cells(i, 3).Value) / Cells(i, 3).Value

' Add to the Volume Total
Total_Volume = Total_Volume + Cells(i, 7).Value

' Print the ticker name in the Summary Table
Range("J" & Summary_Year_Row).Value = Ticker_Name

' Print the total amount per year per ticker name to the Summary Table
Range("K" & Summary_Year_Row).Value = Year_Total

' Print the percent change per year per ticker name to the Summary Table
Range("L" & Summary_Year_Row).Value = Year_Percent

' Print the total volume per year per ticker name to the Summary Table
Range("M" & Summary_Year_Row).Value = Total_Volume


' Add one to the summary table row
Summary_Year_Row = Summary_Year_Row + 1
   
' Reset the Ticker Total
Year_Total = 0

' Reset the Ticker Total
Total_Volume = 0

' If the cell immediately following a row is the same ticker name...
Else

' Add to the Year Total
Year_Total = Year_Total + Cells(i, 6).Value - Cells(i, 3).Value

' Add to the percentage change
Year_Percent = (Year_Total + Cells(i, 6).Value - Cells(i, 3).Value) / Cells(i, 3).Value

' Add to the Volume Total
Total_Volume = Total_Volume + Cells(i, 7).Value



    End If

    Next i

' Set an initial variable for holding the Percent Change
Dim Percent_Change As Range
Set Percent_Change = Range("L2", Range("L2").End(xlDown))

Range("L2", Range("L2").End(xlDown)).NumberFormat = "0.00%"
For Each Cell In Percent_Change

If Cell.Value >= 0 Then
Cell.Interior.ColorIndex = 4

      Else

        Cell.Interior.ColorIndex = 3

End If

Next


End Sub

