Attribute VB_Name = "Module1"

Sub Worksheet_loop():
    'Declare ws variable to loop through all worksheets
    Dim ws As Worksheet
    
    'Turn off screen updating to speed up the macro
    Application.ScreenUpdating = False
    
    'Loop through all worksheets
    For Each ws In Worksheets
        ws.Select
        Call VBAchallenge
    Next ws
    
    'Turn screen updating back on
    Application.ScreenUpdating = True
    
End Sub


Sub VBAchallenge():

'Count the number of total rows
Dim rowcount As Double
rowcount = Cells(Rows.Count, "A").End(xlUp).Row

'Create headers for new columns
[I1] = "Ticker"
[J1] = "Yearly Change"
[K1] = "Percent Change"
[L1] = "Total Stock Volume"
[O1] = "Ticker"
[P1] = "Value"
[N2] = "Greatest % Inc"
[N3] = "Greatest % Dec"
[N4] = "Greatest Total Vol"

'Make new headers bold
Range("I1", Range("I1").End(xlToRight)).Font.Bold = True
Range("N1:P4").Font.Bold = True

'Autofit all columns and center
Columns("A:P").AutoFit

'Copy first ticker to table
Cells(2, 9).Value = Cells(2, 1).Value

'Set row for next ticker to be copied to
tickerrow = 2

'Declare variables to calculate yearly change
Dim openprice, closeprice, yearlychange As Double

'Set initial value for openprice
openprice = [C2]

'Declare variable to calculate percentchange
Dim percentchange As Double

'Declare variable to calculate total volume and set initial value to 0
Dim totalstockvalue As LongLong
totalstockvalue = 0

'Loop through all tickers
For i = 2 To (rowcount)

    'Add row's volume to totalstockvalue
    totalstockvalue = totalstockvalue + Cells(i, 7).Value
    
    'Check for new ticker to copy to table
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
    'Print totalstockvalue to table
    Cells(tickerrow, 12) = totalstockvalue
    
    'Reset totalstockvalue to 0
    totalstockvalue = 0
    
    'Create var and store the new ticker
    Dim ticker As String
    ticker = Cells(i, 1).Value
    
    'Print new ticker to table
    Range("I" & tickerrow).Value = ticker
    
    'Store closing price
    closeprice = Cells(i, 6).Value
    
    'Calculate yearly change and print to table
    yearlychange = closeprice - openprice
    Cells(tickerrow, 10).Value = yearlychange
    
    'Calculate percent change and print to table
    If openprice = 0 Then
    Cells(tickerrow, 11).Value = "NA"
    
    Else
    percentchange = (closeprice - openprice) / openprice
    Cells(tickerrow, 11).Value = percentchange
    
    End If
    
    'Reset open price value for next ticker
    openprice = Cells(i + 1, 3).Value
    
     'Add one to tickerrow
    tickerrow = tickerrow + 1
    
    'If ticker is same as previous row
    Else
    
    End If
    
Next i

'Count number of rows in table
Dim tablerowcount As Integer
tablerowcount = Cells(Rows.Count, "I").End(xlUp).Row

'Format Column K as percent with 2 digits
Range("K2:K" & tablerowcount).NumberFormat = "0.00%"

'Format new columns so text is centered
Range("I1:L" & tablerowcount, "N1:P4").HorizontalAlignment = xlCenter

'Add commas to volume numbers to make more readable
Range("L2:L" & tablerowcount).NumberFormat = "###,###,###,##0"

'Add conditional formatting to Column J, yearly change
For j = 1 To tablerowcount
    If Cells(j, 10).Value < 0 Then
    Cells(j, 10).Interior.ColorIndex = 3
    Else: Cells(j, 10).Interior.ColorIndex = 4
    End If
Next j

'Declare variables and create For loop to determine row with greatest % increase
Dim maxpercent As Double
Dim maxticker As String

maxpercent = 0.001

For k = 2 To tablerowcount
    If (Cells(k, 11).Value <> "NA") Then
        If (Cells(k, 11).Value > maxpercent) Then
            maxpercent = Cells(k, 11).Value
            maxticker = Cells(k, 9).Value
        End If
    ElseIf (Cells(k, 11).Value = "NA") Then
    End If
Next k

'Print values to table for greatest % increase
[O2] = maxticker
[P2] = maxpercent

'Declare variables and create For loop to determine row with greatest % increase
Dim minpercent As Double
Dim minticker As String

minpercent = 0

For m = 2 To tablerowcount
    If (Cells(m, 11).Value < minpercent) Then
        minpercent = Cells(m, 11).Value
        minticker = Cells(m, 9).Value
    End If
Next m

'Print values to table for greatest % increase
[O3] = minticker
[P3] = minpercent

'Declare variables and create For loop to determine stock with greatest total volume
Dim maxvolume As LongLong
Dim maxvolticker As String

maxvolume = 1

For n = 2 To tablerowcount
    If (Cells(n, 12).Value > maxvolume) Then
        maxvolume = Cells(n, 12).Value
        maxvolticker = Cells(n, 9).Value
    End If
Next n

'Print values to table for greatest total volume
[O4] = maxvolticker
[P4] = maxvolume

'Update formatting for 2nd part of table
Range("P2:P3").NumberFormat = "0.00%"
Range("O2:P4").Font.Bold = False
Range("P4").NumberFormat = "###,###,###,#00"

End Sub
