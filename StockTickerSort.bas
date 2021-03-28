Attribute VB_Name = "Module2"
Sub Stock_Ticker_Sort()

Dim ticker As String

Dim Summary As Integer

Dim volume As Double

volume = 0

Summary = 2

Cells(1, 10).Value = "Ticker"

Cells(1, 11).Value = "Yearly Change"

Cells(1, 12).Value = "Percent Change"

Cells(1, 13).Value = "Total Stock Volume"

lastrow = Cells(Rows.Count, 1).End(xlUp).Row

open_price = Cells(2, 3).Value

    For i = 2 To lastrow

        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

        ticker = Cells(i, 1).Value
        
        volume = volume + Cells(i, 7).Value
        
        Range("M" & Summary).Value = volume

        Range("J" & Summary).Value = ticker

        close_price = Cells(i, 6).Value

        delta_change = close_price - open_price

        Range("K" & Summary).Value = delta_change

            If open_price = 0 Then

            Percent_change = 0

            Else

            Percent_change = (delta_change / open_price)

            End If
        
        open_price = Cells(i + 1, 3).Value
        
        Range("L" & Summary).Value = Percent_change

        Range("L" & Summary).NumberFormat = "0.00%"
        
        Summary = Summary + 1
        
        volume = 0
        
        Else
        
        volume = volume + Cells(i, 7).Value
        
        End If

If Cells(i, 11).Value < 0 Then
 
    Cells(i, 11).Interior.ColorIndex = 3
    
    Else
    
    Cells(i, 11).Interior.ColorIndex = 4
    
End If



Next i

End Sub
