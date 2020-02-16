Attribute VB_Name = "Module1"
Sub Stonks()

'defining all our variables'
For Each ws In Worksheets


Dim Ticker As String

Dim StockTotal As Double

Dim OpenPrice As Double
Dim ClosePrice As Double
Dim PriceDif As Double
Dim PricePrct As Double

Dim SummaryTable As Integer


StockTotal = 0

'starting off summtable at row two since headers will be first row'

SummaryTable = 2

'Headers'

ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Stock Volume Total"
ws.Cells(1, 11).Value = "Price Difference"
ws.Cells(1, 12).Value = "Price Percentage"


For i = 2 To ws.Cells(Rows.Count, "A").End(xlUp).Row


'Stating that if the value above doesn't equal value ; that must be the start of the year'

        If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
            
            OpenPrice = ws.Cells(i, 3)
            
            StockTotal = StockTotal + ws.Cells(i, 7).Value
            
'if the beg year is not applicable, then move on to the next row and add the stock volume up'
            
        ElseIf ws.Cells(i + 1, 1).Value = ws.Cells(i, 1).Value Then
        
            StockTotal = StockTotal + ws.Cells(i, 7).Value
            
'when next row doesn't equal current row start the outputting and adding up process'
            
        ElseIf ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
            Ticker = ws.Cells(i, 1).Value
            
            StockTotal = StockTotal + ws.Cells(i, 7).Value
            
            ClosePrice = ws.Cells(i, 6).Value
            
            PriceDif = (ClosePrice - OpenPrice)
            
'Nested if statment incase Opening Price is 0 -> don't want to get an error; going to default and make price prct 0'
                
                If OpenPrice = 0 Then
                    
                    PricePrct = 0
                
                Else
                    
                    PricePrct = (PriceDif / OpenPrice)
                    
                End If
                
            
'placing all the numbers onto the sheet
        
            ws.Range("I" & SummaryTable).Value = Ticker
            
            ws.Range("J" & SummaryTable).Value = StockTotal
            
            ws.Range("K" & SummaryTable).Value = PriceDif
            
            ws.Range("L" & SummaryTable).Value = PricePrct
            
'Changed the PricePrct to be an actual percentage'
            
            ws.Range("L" & SummaryTable).NumberFormat = "0.00%"
        
'Now we're changing the colors of the Percent'
    
                If PriceDif < 0 Then
                
                    ws.Range("K" & SummaryTable).Interior.ColorIndex = 3
                    
                Else
                    
                    ws.Range("K" & SummaryTable).Interior.ColorIndex = 4
                
                End If
                
'Adding one to the summary table to go to the next row'
            
            SummaryTable = SummaryTable + 1
            
    
'now we're starting over; so stocktotal needs to be back to 0'
            
            StockTotal = 0
            
        End If
    
Next i

'starting our next loop to look for greatest change; start by define all variables'


Dim GreatestVol As Double
Dim GreatPrct As Double
Dim LeastPrct As Double

Dim VolTick As String

GreatestVol = 0
GreatPrct = 0
LeastPrct = 0

'set up titles'

ws.Cells(1, 15).Value = "Ticker"
ws.Cells(1, 16).Value = "Value"

ws.Cells(2, 14).Value = "Greatest % Increase"
ws.Cells(3, 14).Value = "Greatest % Decrease"
ws.Cells(4, 14).Value = "Greatest Total Volume"

'start first for loop looking for volume'

 For i = 2 To ws.Cells(Rows.Count, "I").End(xlUp).Row

    If ws.Cells(i, 10).Value > GreatestVol Then
        GreatestVol = ws.Cells(i, 10).Value
        VolTick = ws.Cells(i, 9).Value
        
        ws.Cells(4, 15).Value = VolTick
        ws.Cells(4, 16).Value = GreatestVol
        
    End If
    
Next i

'start second for loop looking for greatest percent'

For i = 2 To ws.Cells(Rows.Count, "I").End(xlUp).Row

    If ws.Cells(i, 12).Value > GreatPrct Then
        GreatPrct = ws.Cells(i, 12).Value
        VolTick = ws.Cells(i, 9).Value
        
        ws.Cells(2, 15).Value = VolTick
        ws.Cells(2, 16).Value = GreatPrct
        ws.Range("P2").NumberFormat = "0.00%"
        
    End If

Next i

'start third for loop looking for least percent'

For i = 2 To ws.Cells(Rows.Count, "I").End(xlUp).Row

    If ws.Cells(i, 12).Value < LeastPrct Then
        LeastPrct = ws.Cells(i, 12).Value
        VolTick = ws.Cells(i, 9).Value
        
        ws.Cells(3, 15).Value = VolTick
        ws.Cells(3, 16).Value = LeastPrct
        ws.Range("P3").NumberFormat = "0.00%"
        
    End If

Next i
            
    
Next ws

End Sub
