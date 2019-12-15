Sub StockData()
    
    'Assign Variables to worksheet
    Dim ws As Worksheet

    'Loop through each worksheet
    For Each ws In Worksheets
        ws.Activate

    'Assign Variables
    Dim deltaA As Double
    Dim nextDayTicker As String
    Dim currentTicker As String
    Dim currentOpeningValue As Double
    Dim nextOpeningValue As Double
    Dim summaryRow As Double
    Dim percentChange As Double
    Dim stockVolume As Double
    Dim currStockVolume As Double
    Dim lastRow As Long

    'Initialize Summary Table Headers
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"

    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
     'Initialize first summary row in order to populate the correct cell with summary data
    summaryRow = 2

    'Initialize Current Ticker
    currentTicker = Cells(2, 1).Value

    'Initialize the current opening value for the ticker
    currentOpeningValue = Cells(2, 3).Value

    'Initialize stock volume
    stockVolume = 0


    'Loop through rows
    For i = 1 To lastRow
        
        'Set next day ticker, the next stock value, and next opening value in order to compare within the for loop
        nextDayTicker = Cells(i + 1, 1).Value
        nextOpeningValue = Cells(i + 1, 3).Value
        currStockVolume = Cells(i + 1, 7).Value

        'Begin for loop with look ahead to begin adding up the total stock volume
        If nextDayTicker = currentTicker Then

            'Add the current stock value in the loop to the to
            stockVolume = stockVolume + currStockVolume

            'carry the stock volume into the next loop
            currStockVolume = stockVolume

        'Check Look Ahead for next tday ticker
        ElseIf nextDayTicker <> currentTicker Then
            
            'calculate the yearly change
            deltaA = Cells(i, 6).Value - currentOpeningValue

            'Calculate the percent change, taking 0 values into account
            If currentOpeningValue <> 0 Then
                percentChange = deltaA / currentOpeningValue

                'Populate cells with percent change
                Cells(summaryRow, 11).Value = percentChange
             End If

            Cells(summaryRow, 12).Value = stockVolume

            'format cells as percentage
            Cells(summaryRow, 11).NumberFormat = "0.00%"


            'Populate cells with yearly change
            Cells(summaryRow, 10).Value = deltaA

                'Use conditional formatting in a nested if statement to make positive changes green, and negative changes red
                If Cells(summaryRow, 10).Value > 0 Then
        
                Cells(summaryRow, 10).Interior.ColorIndex = 4
            
                Else
                
                Cells(summaryRow, 10).Interior.ColorIndex = 3
            
                End If

            'Populate cells with Ticker
            Cells(summaryRow, 9).Value = currentTicker

            'reset current ticker, current opening value, and summary row for the next iteration of the loop
            currentTicker = nextDayTicker
            summaryRow = summaryRow + 1
            currentOpeningValue = nextOpeningValue
            stockVolume = 0

        End If

    Next i

Next

'hard solution

    'Initialize headers
    'Cells(2, 14).Value = "Greatest % Increase"
    'Cells(3, 14).Value = "Greatest % Descrease"
    'Cells(4, 14).Value = "Greatest Total Volume"
    'Cells(1, 15).Value = "Ticker"
    'Cells(1, 16).Value = "Value"

    'Dim range As range
    'Dim dblMax As Double


    'Cells(2, 16).Value = dblMax

    'For j = 2 To lastRow
            'Set Rng = Sheet1.range("")
            'dblMax = Application.WorksheetFunction.Max(Rng)

        
        'Cells(j, 10).Value = WorksheetFunction.Min(Cells(j, 10))
        
        'Cells(j, 11).Value = WorksheetFunction.Max(Cells(j, 11)
    'Next j
        
End Sub






