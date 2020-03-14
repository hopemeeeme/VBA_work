
Sub Stock_Stats():
Dim ticker As String
Dim num_tickers As Integer
Dim lastrow As Long
Dim open_p As Double
Dim close_p As Double
Dim year_change As Double
Dim percent_change As Double
Dim totalV As Double
Dim greatest_percent_incr As Double
Dim greatest_percent_incrticker As String
Dim greatest_percent_decr As Double
Dim greatest_percent_decrticker As String
Dim greatest_totalV As Double
Dim greatest_totalVticker As String

For Each ws In Worksheets

    ws.Activate
    lastrow = ws.Cells(Rows.Count, "A").End(xlUp).Row

    ' Add header columns for each worksheet
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Volume"

    ' Initialize variables for each worksheet.
    num_tickers = 0
    ticker = ""
    year_change = 0
    open_p = 0
    percent_change = 0
    totalV = 0

    ' Skipping the header row to loop thru
    For i = 2 To lastrow

        ticker = Cells(i, 1).Value
        If open_p = 0 Then
            open_p = Cells(i, 3).Value
        End If

        ' Add total volume values
        totalV = totalV + Cells(i, 7).Value

        If Cells(i + 1, 1).Value <> ticker Then
            num_tickers = num_tickers + 1
            Cells(num_tickers + 1, 9) = ticker
            
            close_p = Cells(i, 6)

            'year change value
            year_change = close_p - open_p

            ' Add yearly change value to each cell in each worksheet.
            Cells(num_tickers + 1, 10).Value = year_change

            'year change is >0, shade cell green.
            If year_change > 0 Then
                Cells(num_tickers + 1, 10).Interior.ColorIndex = 4
            'year change <0, shade cell red.
            ElseIf year_change < 0 Then
                Cells(num_tickers + 1, 10).Interior.ColorIndex = 3
            End If

            'calc percent change.
            If open_p = 0 Then
                percent_change = 0
            Else
                percent_change = (year_change / open_p)
            End If

            ' Format the percent_change to %
            Cells(num_tickers + 1, 11).Value = Format(percent_change, "Percent")
            
            open_price = 0
            Cells(num_tickers + 1, 12).Value = totalV
            totalV = 0
            
        End If

    Next i

    'display greatest percent increase, greatest percent decrease, and greatest total volume for each year.
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"

    ' Get the last row
    lastrow = ws.Cells(Rows.Count, "I").End(xlUp).Row

    ' Initialize variables and set values of variables initially to the first row in the list.
    greatest_percent_incr = Cells(2, 11).Value
    greatest_percent_incrticker = Cells(2, 9).Value
    greatest_percent_decr = Cells(2, 11).Value
    greatest_percent_decrticker = Cells(2, 9).Value
    greatest_totalV = Cells(2, 12).Value
    greatest_totalVticker = Cells(2, 9).Value


    ' skipping the header row, loop thru tickers
    For i = 2 To lastrow

        ' Find the ticker with the greatest percent increase.
        If Cells(i, 11).Value > greatest_percent_incr Then
            greatest_percent_incr = Cells(i, 11).Value
            greatest_percent_incrticker = Cells(i, 9).Value
        End If

        ' Find the ticker with the greatest percent decrease.
        If Cells(i, 11).Value < greatest_percent_decr Then
            greatest_percent_decr = Cells(i, 11).Value
            greatest_percent_decrticker = Cells(i, 9).Value
        End If

        ' Find the ticker with the greatest stock volume.
        If Cells(i, 12).Value > greatest_totalV Then
            greatest_totalV = Cells(i, 12).Value
            greatest_totalVticker = Cells(i, 9).Value
        End If

    Next i

    'values for greatest percent increase, decrease, and stock volume to each worksheet.
    Range("P2").Value = Format(greatest_percent_incrticker, "Percent")
    Range("Q2").Value = Format(greatest_percent_incr, "Percent")
    Range("P3").Value = Format(greatest_percent_decrticker, "Percent")
    Range("Q3").Value = Format(greatest_percent_decr, "Percent")
    Range("P4").Value = greatest_totalVticker
    Range("Q4").Value = greatest_totalV

Next ws


End Sub
