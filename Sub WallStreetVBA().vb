Sub WallStreetVBA()

Dim ws As Integer

ws = Application.Worksheets.Count

For i = 1 To ws
    Worksheets(i).Activate

    Dim LastRow As Long
    Dim TikcerNumber As Long
    Range("J1").Value = "Ticker"
    Range("K1").Value = "Yearly Change"
    Range("L1").Value = "% Change"
    Range("M1").Value = "Total Volume"
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Volume"
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"

    'Set TickerNumber to 2 to start in the second row
    Tickernumber = 1

    LastRowLeft = Cells(Rows.Count, 1).End(xlUp).Row

    For y = 2 To LastRowLeft

        'Below says "If the ticker has changed from the row above then"
        If Cells(y, 1).Value <> Cells(y - 1, 1).Value Then
        'Resets yearly volume counter
        YearlyVolume = 0
        'Sets the YearOpen for the new ticker
        YearOpen = Cells(y, 3).Value
        'The below adds the ticker to the right of the data
        Cells(Tickernumber + 1, 10).Value = Cells(y, 1).Value
        'The below tells our Sub to move down a cell to place the next name
        Tickernumber = Tickernumber + 1
        End If

        'Declares the daily volume
        DailyVolume = Cells(y, 7).Value
        'Adds to the yearly volume running total
        YearlyVolume = YearlyVolume + DailyVolume

        'Says "If the ticker is different from the one below then"
        'I had to add the and >0 as a stock lacking data for a year was causing a divide by 0 error
        If (Cells(y, 1).Value <> Cells(y + 1, 1).Value) And (YearOpen > 0) Then
        'Prints the yearly volume
        Cells(Tickernumber, 13).Value = YearlyVolume
        'Sets the yearly close
        YearClose = Cells(y, 6).Value
        'Determines and prints the yearly change in price
        Cells(Tickernumber, 11).Value = (YearClose - YearOpen)
        'Determines and prints the yearly change in %
        Cells(Tickernumber, 12).Value = ((YearClose / YearOpen) - 1)
        Cells(y, 8).Value = Cells(y, 8).Value * 100
        End If

        'Color the yearly value & % change green for gains, red for losses, grey for no change
        If Cells(Tickernumber, 11).Value > 0 Then
        Cells(Tickernumber, 11).Interior.ColorIndex = 4
        Cells(Tickernumber, 12).Interior.ColorIndex = 4
        ElseIf Cells(Tickernumber, 11).Value < 0 Then
        Cells(Tickernumber, 11).Interior.ColorIndex = 3
        Cells(Tickernumber, 12).Interior.ColorIndex = 3
        Else
        Cells(Tickernumber, 11).Interior.ColorIndex = 15
        Cells(Tickernumber, 12).Interior.ColorIndex = 15
        End If

    Next y

    LastRowRight = Cells(Rows.Count, 10).End(xlUp).Row
    GreatestIncrease = 0
    GreatestIncreaseTicker = "null"
    GreatestDecrease = 0
    GreatestDecreaseTicker = "null"
    GreatestVolume = 0
    GreatestVolumeTicker = "null"

    For y = 2 To LastRowRight

        If Cells(y, 12).Value > GreatestIncrease Then
        GreatestIncrease = Cells(y, 12).Value
        GreatestIncreaseTicker = Cells(y, 10).Value
        ElseIf Cells(y, 12).Value < GreatestDecrease Then
        GreatestDecrease = Cells(y, 12).Value
        GreatestDecreaseTicker = Cells(y, 10).Value
        End If

        If Cells(y, 13).Value > GreatestVolume Then
        GreatestVolume = Cells(y, 13).Value
        GreatestVolumeTicker = Cells(y, 10).Value
        End If
    Next y

    Range("P2").Value = GreatestIncreaseTicker
    Range("P3").Value = GreatestDecreaseTicker
    Range("P4").Value = GreatestVolumeTicker
    Range("Q2").Value = GreatestIncrease
    Range("Q3").Value = GreatestDecrease
    Range("Q4").Value = GreatestVolume
Next i

End Sub