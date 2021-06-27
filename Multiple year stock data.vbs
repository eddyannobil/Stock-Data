VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub Stock_market()

    'Declare and set worksheet
    Dim ws As Worksheet

    'Loop through all stocks for one year
    For Each ws In Worksheets

        'Create column headings for summary table
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"

        'Create column headings for analysis table
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"

        'Define Ticker variable
        Dim Ticker As String

        'Define Ticker Volume variable
        Dim Ticker_volume As Long
        Ticker_volume = 0

        'Define last row of worksheet
        Dim Lastrow As Long
        Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        'Define last row of summary table
        Dim lastSummaryRow As Long
        lastSummaryRow = ws.Cells(Rows.Count, 9).End(xlUp).Row

        'Define and initialize the summary table row
        Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2

        'Define and initialize variable to hold open price
        Dim open_price As Double
        open_price = 0
        
        'Define and initialize variable to hold close price
        Dim close_price As Double
        close_price = 0

        'Define and initialize variable to hold price change
        Dim yearly_price_change As Double
        yearly_price_change = 0

        'Define and initialize variable to hold price change %
        Dim price_change_percent As Double
        price_change_percent = 0

        'Define and initialize variable to hold Greatest increase
        Dim Greatest_Increase As Double
        Greatest_Increase = 0

        'Define and initialize variable to hold Greatest decrease
        Dim Greatest_Decrease As Double
        Greatest_Decrease = 0
        
        'Define and initialize variable to hold Greatest volume
        Dim Greatest_Volume As Double
        Greatest_Volume = 0

        'Do loop of current worksheet to Lastrow
        For i = 2 To Lastrow

            'Set the yearly open price for each ticker
            If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then

                open_price = ws.Cells(i, 3).Value

            End If

            'Build the summary table
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                'Set Ticker to be added to the Summary Table
                Ticker = ws.Cells(i, 1).Value
                ws.Range("I" & Summary_Table_Row).Value = Ticker
                
                'Compute the total volume for a given ticker and add to summary table
                Ticker_volume = Ticker_volume + ws.Cells(i, 7).Value
                ws.Range("L" & Summary_Table_Row).Value = Ticker_volume

                'Set Close Price
                close_price = ws.Cells(i, 6).Value
                    
                'Calculate and print yearly change in price
                yearly_price_change = close_price - open_price
                ws.Range("J" & Summary_Table_Row).Value = yearly_price_change
                
                'Calculate the price percent change and add to summary table as a percentage
                If open_price <> 0 Then
                    price_change_percent = (yearly_price_change / open_price)
                    ws.Range("K" & Summary_Table_Row).Value = price_change_percent
                    ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
                End If
                
                If ws.Range("J" & Summary_Table_Row) < 0 Then
                    ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                Else
                    ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                End If

                'Increment summary table row count
                Summary_Table_Row = Summary_Table_Row + 1

                'Reset open price, close price, price change, price change percent, ticker volume
                open_price = 0
                close_price = 0
                year_change = 0
                price_change_percent = 0
                Ticker_volume = 0
            End If
        Next i
        'Loop to search through summary table
        For j = 2 To lastSummaryRow

            'Greatest% Increase
            If ws.Cells(j, 11).Value > Greatest_Increase Then
                Greatest_Increase = ws.Cells(j, 11).Value
                ws.Range("Q2").Value = Greatest_Increase
                ws.Range("P2").Value = ws.Cells(j, 9).Value
            End If

            'Greatest% Decrease
            If ws.Cells(j, 11).Value < Greatest_Decrease Then
                Greatest_Decrease = ws.Cells(j, 11).Value
                ws.Range("Q3").Value = Greatest_Decrease
                ws.Range("P3").Value = ws.Cells(j, 9).Value
            End If
                    
            'Greatest Total Volume
            If ws.Cells(j, 12).Value > Greatest_Volume Then
                Greatest_Volume = ws.Cells(j, 12).Value
                ws.Range("Q4").Value = Greatest_Volume
                ws.Range("P4").Value = ws.Cells(j, 9).Value
            End If
        Next j

        'Format some cells in Analysis table as percentages
        ws.Range("Q2").NumberFormat = "0.00%"
        ws.Range("Q3").NumberFormat = "0.00%"

    Next ws
End Sub

