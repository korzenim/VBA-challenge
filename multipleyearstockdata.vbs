Sub challenge2()
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
    Cells(1, 16).Value = "Ticker"
    Cells(1, 17).Value = "Value"
    Cells(2, 15).Value = "Greatest % Increase"
    Cells(3, 15).Value = "Greatest % Decrease"
    Cells(4, 15).Value = "Greatest Total Volume"
    
    For i = 2 To 800000
        If i = 2 Then
            ticker = Cells(2, 1).Value
            tickcount = 2
            volume = Cells(2, 7).Value
            tickopen = Cells(2, 3).Value
            tickclose = 0
            yearchange = 0
            perchange = 0
            Cells(2, 9).Value = ticker
        Else
            If (Cells(i, 1).Value = ticker) Then
                volume = volume + Cells(i, 7).Value
            Else:
                tickclose = Cells(i - 1, 6).Value
                yearchange = tickclose - tickopen
                perchange = FormatPercent((tickclose - tickopen) / tickopen)
                Cells(tickcount, 10) = yearchange
                Cells(tickcount, 11) = perchange
                tickopen = Cells(i, 3).Value
                volume = Cells(i, 7).Value
                tickcount = tickcount + 1
                ticker = Cells(i, 1).Value
                Cells(tickcount, 9).Value = ticker
            End If
        End If
    Next i

    greatinc = 0
    greatdec = 0
    greattot = 1
    For j = 2 To 4000
        If Cells(j, 10).Value < 0 Then
            Cells(j, 10).Interior.ColorIndex = 3
        ElseIf Cells(j, 10).Value > 0 Then
            Cells(j, 10).Interior.ColorIndex = 4
        Else
            Cells(j, 10).Interior.ColorIndex = 8
        End If
        If Cells(j, 11).Value > greatinc Then
            Cells(2, 16).Value = Cells(j, 9).Value
            Cells(2, 17).Value = FormatPercent(Cells(j, 11).Value)
            greatinc = Cells(j, 11).Value
        ElseIf Cells(j, 11).Value < greatdec Then
            Cells(3, 16).Value = Cells(j, 9).Value
            Cells(3, 17).Value = FormatPercent(Cells(j, 11).Value)
            greatdec = Cells(j, 11).Value
        End If
    Next j
    For k = 2 To 4000
        If Cells(k, 12) > greattot Then
            greattot = Cells(k, 12)
            Cells(4, 16) = Cells(k, 9)
            Cells(4, 17) = greattot
        End If
    Next k
End Sub
