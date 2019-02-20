Sub calls()

Dim ticker As String
Dim Startprice As Double
Startprice = 0
Dim endprice As Double
endprice = 0
Dim totalvolume As Double
totalvolume = 0
Dim abschg As Double
abschg = 0
Dim abschg2 As Double
abschg2 = 0
Dim pctchg As Double
pctchg = 0
Dim rowcount As Double
rowcount = 0

finalrow = Cells(Rows.Count, "b").End(xlUp).Row

Dim summary_table_row As Integer
summary_table_row = 2


For i = 2 To finalrow

    If Cells(i - 1, 1).Value <> Cells(i, 1).Value Then
        abschg = Cells(i, 3).Value

    End If
    
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
        ticker = Cells(i, 1).Value
        totalvolume = totalvolume + Cells(i, 7).Value
        abschg2 = Cells(i, 6).Value
            
     Range("L" & summary_table_row).Value = ticker
     Range("M" & summary_table_row).Value = abschg2 - abschg
     Range("N" & summary_table_row).Value = (abschg2 / abschg) - 1
     Range("O" & summary_table_row).Value = totalvolume
          
     summary_table_row = summary_table_row + 1
     
     totalvolume = 0
  
   Else
   
    totalvolume = totalvolume + Cells(i, 7).Value
          
End If

Next i

finalrow2 = Cells(Rows.Count, "l").End(xlUp).Row
Dim rng As Range
Dim rng2 As Range

For j = 2 To finalrow2

Cells(1, 12).Value = "Ticker"
Cells(1, 13).Value = "Yearly Change"
Cells(1, 14).Value = "Percent Change"
Cells(1, 15).Value = "Total Stock Volume"


Columns("N").NumberFormat = "0.00%"


If Cells(j, 14).Value > 0 Then
    Cells(j, 13).Interior.ColorIndex = 4
    End If
If Cells(j, 14).Value < 0 Then
    Cells(j, 13).Interior.ColorIndex = 3
    End If

Set rng = Range(Cells(2, 14), Cells(finalrow2, 14))
Set rng2 = Range(Cells(2, 15), Cells(finalrow2, 15))


Cells(1, 19).Value = "Ticker"
Cells(1, 20).Value = "Value"
Cells(2, 18).Value = "Greatest % Increase"
Cells(3, 18).Value = "Greatest % Decrease"
Cells(4, 18).Value = "Greatest Total Volume"

Cells(2, 20).Value = Application.WorksheetFunction.Max(rng)
Cells(3, 20).Value = Application.WorksheetFunction.Min(rng)
Cells(4, 20).Value = Application.WorksheetFunction.Max(rng2)


If Cells(j, 14) = Application.WorksheetFunction.Max(rng) Then
    Cells(2, 19).Value = Cells(j, 12)
End If
    
If Cells(j, 14) = Application.WorksheetFunction.Min(rng) Then
    Cells(3, 19).Value = Cells(j, 12)
End If

If Cells(j, 15) = Application.WorksheetFunction.Max(rng2) Then
    Cells(4, 19).Value = Cells(j, 12)
End If

Cells(2, 20).NumberFormat = "0.00%"
Cells(3, 20).NumberFormat = "0.00%"

Next j

End Sub
