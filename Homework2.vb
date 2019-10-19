    Sub VBAofWS()
    'Looping through all worksheets in workbook
    For Each ws In Worksheets

    'declaring variables for calculations below
    Dim StockName As String
    Dim StockTotal As LongLong
    Dim StockStart As Double
    Dim StockEnd As Double
    Dim YearlyChange As Double
    Dim Table_Row As Integer
    Dim PercentChange As Double

    'setting original variable values
    StockTotal = 0
    'setting starting stock price as first row of actual data
    StockStart = ws.Cells(2, 3).Value
    StockEnd = 0
    YearlyChange = 0
    Table_Row = 2
    PercentChange = 0

    'creating new header row for information of stocks
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"

    'setting lastRow variable to move to end of row
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    'since we have a header row, we need to start at row 2
    For i = 2 To LastRow

    'if cell above does not match cell below
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        'Add stock ticker to summary list
        StockName = ws.Cells(i, 1).Value
        'setting stock end value for percent change calculation
        StockEnd = ws.Cells(i, 3).Value
        'calculating yearly change
        YearlyChange = StockEnd - StockStart
        if(StockStart=0) then 
        PercentChange = 0
        else
        'calculating percent change from beginning to end
        PercentChange = YearlyChange / StockStart
        end if
        'Add Stock name to summary table
        ws.Range("I" & Table_Row).Value = StockName
        'Add Calcuated yearly change value
        ws.Range("J" & Table_Row).Value = YearlyChange
        'if Change is postivie, highlight green, else, highlight red
        If (YearlyChange > 0) Then
        ws.Range("J" & Table_Row).Interior.ColorIndex = 4
        Else
        ws.Range("J" & Table_Row).Interior.ColorIndex = 3
        End If
        'Adding Percentage change to summary table
        ws.Range("K" & Table_Row).Value = PercentChange
        'formatting range as Percentage
        ws.Range("K" & Table_Row).NumberFormat = "0.00%"
        'Adding total value of stocks which is being calculated below
        ws.Range("L" & Table_Row).Value = StockTotal / 1000
        'Continues loop to next row
        Table_Row = Table_Row + 1
    'Sets new starting stock value for next stock in loop
    StockStart = ws.Cells(i + 1, 3).Value
    'clearing out variables for use in future loop iterations
    StockEnd = 0
    StockName = 0
    YearlyChange = 0
    PercentChange = 0
    StockTotal = 0
    'if the the cell above and below have same stock value
    Else
    'Calcuating total volume of stocks for year
    StockTotal = StockTotal + ws.Cells(i, 7).Value
    End If
    'run through next loop
    Next i
    'continue to next worksheet
    Next ws

    End Sub


