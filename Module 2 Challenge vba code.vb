Sub VBA_Challenge()
Dim ws As Worksheet


For Each ws In ThisWorkbook.Worksheets
    Dim ticker(0 To 3000) As String 'Create a ticker array to save each unique ticker
    ticker(0) = ws.Cells(2, 1) 'Intialise first variable of ticker because idk how to do it all in a for loop
    Dim yearly_change(0 To 3000) As Double 'initialise yearly change in array
    yearly_change(0) = ws.Cells(2, 3) ' initialise first element in the array
    Dim percent_change(0 To 3000) As Double
    Dim total_volume(0 To 3000) As Double
    
    
    Dim counter As Integer 'create a counter to keep track of array position
    counter = 0
    Dim open_price As Double 'Create variable to hold open price
    open_price = 0
    
    LastRow = ActiveSheet.Range("A1").CurrentRegion.Rows.Count
    For i = 2 To LastRow 'For loop to go through entire dataset
        
        If ws.Cells(i + 1, 1) <> ws.Cells(i, 1) Then 'When ticker names are different, do this
            total_volume(counter) = total_volume(counter) + ws.Cells(i, 7)
            open_price = yearly_change(counter) 'Save open price to use later for percent change
            yearly_change(counter) = ws.Cells(i, 6) - yearly_change(counter) 'Calculate Yearly change of price by subtracting last closing price from first opening price
            percent_change(counter) = yearly_change(counter) / open_price 'Calculate percent change and save in array
            
            
            counter = counter + 1 'update the counter to move to next position in array
            ticker(counter) = ws.Cells(i + 1, 1) 'Add the new unique ticker name into the array
            yearly_change(counter) = ws.Cells(i + 1, 3) 'Add the opening for new unique ticker
            
        Else
            total_volume(counter) = total_volume(counter) + ws.Cells(i, 7)
        End If
    Next i

ws.Cells(1, 9) = "Ticker"
ws.Cells(1, 10) = "Yearly Change"
ws.Cells(1, 11) = "Percent Change"
ws.Cells(1, 12) = "Total Volume"

    For i = 0 To UBound(ticker)
        ws.Cells(i + 2, 9) = ticker(i)
        ws.Cells(i + 2, 10) = yearly_change(i)
        ws.Cells(i + 2, 11) = percent_change(i)
        ws.Cells(i + 2, 12) = total_volume(i)
        
        If ws.Cells(i + 2, 10) > 0 Then
            ws.Cells(i + 2, 10).Interior.ColorIndex = 4
            ws.Cells(i + 2, 11).Interior.ColorIndex = 4
        Else
            ws.Cells(i + 2, 10).Interior.ColorIndex = 3
            ws.Cells(i + 2, 11).Interior.ColorIndex = 3
        End If
    Next i

'Bonus Part
Dim greatest_increase As Double
Dim greatest_decrease As Double
Dim greatest_volume As Double
greatest_increase = 0
greatest_decrease = 0
greatest_volume = 0

Dim ticker_increase As String
Dim ticker_decrease As String
Dim ticker_volume As String

    For i = 2 To 3001
        If ws.Cells(i, 11) > greatest_increase Then
            greatest_increase = ws.Cells(i, 11)
            ticker_increase = ws.Cells(i, 9)
        ElseIf ws.Cells(i, 11) < greatest_decrease Then
            greatest_decrease = ws.Cells(i, 11)
            ticker_decrease = ws.Cells(i, 9)
        End If
        
        If ws.Cells(i, 12) > greatest_volume Then
            greatest_volume = ws.Cells(i, 12)
            ticker_volume = ws.Cells(i, 9)
        End If
    Next i
    
ws.Cells(1, 15) = "Ticker"
ws.Cells(1, 16) = "Value"
ws.Cells(2, 14) = "Greatest % Increase"
ws.Cells(3, 14) = "Greatest % Decrease"
ws.Cells(4, 14) = "Greatest Total Volume"

ws.Cells(2, 15) = ticker_increase
ws.Cells(2, 16) = greatest_increase
ws.Cells(3, 15) = ticker_decrease
ws.Cells(3, 16) = greatest_decrease
ws.Cells(4, 15) = ticker_volume
ws.Cells(4, 16) = greatest_volume

Next ws
End Sub




End Sub
