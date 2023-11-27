Attribute VB_Name = "Module1"
'Caitlyn Epley
'Module 2 Challenge

Sub vba_challenge():

'Run script on each sheet in workbook
Dim ws As Worksheet
For Each ws In ThisWorkbook.Sheets
With ws

'Define Variables
Dim ticker As String
ticker = .Cells(2, 1)
Dim open_price As Double
open_price = .Cells(2, 3)
Dim close_price As Double
close_price = 0
Dim total_stock As Double
total_stock = .Cells(2, 7)

'Row count for yearly statistics
Dim row_count As Integer
row_count = 2

'number of rows in first column
EndRow = .Cells(Rows.count, 1).End(xlUp).Row

'create headers in file
.Cells(1, 9) = "Ticker"
.Cells(1, 10) = "Yearly Change"
.Cells(1, 11) = "Percent Change"
.Cells(1, 12) = "Total Stock Volume"

.Cells(1, 16) = "Ticker"
.Cells(1, 17) = "Value"
.Cells(2, 15) = "Greatest % Increase"
.Cells(3, 15) = "Greatest % Decrease"
.Cells(4, 15) = "Greatest Total Volume"

'loop through all data
For i = 3 To EndRow
    'if value is from the same ticker, get info from it
    If (ticker = .Cells(i, 1)) Then
        close_price = .Cells(i, 6)
        total_stock = total_stock + .Cells(i, 7)
    'if value is not from same ticker, assign values and reset variables
    Else
        .Cells(row_count, 9) = ticker
        .Cells(row_count, 10) = close_price - open_price
        .Cells(row_count, 11) = (close_price - open_price) / open_price
        .Cells(row_count, 11).NumberFormat = "0.00%"
        .Cells(row_count, 12) = total_stock
        'Change cell colors depending on yearly change
        If .Cells(row_count, 10) > 0 Then
            .Cells(row_count, 10).Interior.ColorIndex = 4
        Else
            .Cells(row_count, 10).Interior.ColorIndex = 3
        End If
        
        row_count = row_count + 1
        ticker = .Cells(i, 1)
        open_price = .Cells(i, 3)
        close_price = .Cells(i, 6)
        total_stock = 0
    End If
Next i

'number of rows of output from previous section
endrow2 = .Cells(Rows.count, 11).End(xlUp).Row

'ticker for each value searching for
Dim ticker_inc As String
ticker_inc = .Cells(2, 9)
Dim ticker_dec As String
ticker_dec = .Cells(2, 9)
Dim ticker_vol As String
ticker_vol = .Cells(2, 9)

'variables for each value searching for
Dim value_inc As Double
value_inc = .Cells(2, 11)
Dim value_dec As Double
value_dec = .Cells(2, 11)
Dim value_vol As Double
value_vol = .Cells(2, 12)

'loop through output from previous section
For j = 3 To endrow2
    'find the greatest value for each
    If value_inc < .Cells(j, 11) Then
        value_inc = .Cells(j, 11)
        ticker_inc = .Cells(j, 9)
    ElseIf value_dec > .Cells(j, 11) Then
        value_dec = .Cells(j, 11)
        ticker_dec = .Cells(j, 9)
    End If
    
    If value_vol < .Cells(j, 12) Then
        value_vol = .Cells(j, 12)
        ticker_vol = .Cells(j, 9)
    End If
    
 Next j

'assign greatest values to cells
.Cells(2, 16) = ticker_inc
.Cells(3, 16) = ticker_dec
.Cells(4, 16) = ticker_vol

.Cells(2, 17) = value_inc
.Cells(3, 17) = value_dec
.Cells(4, 17) = value_vol

    
End With
Next ws

End Sub
