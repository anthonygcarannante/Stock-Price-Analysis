Attribute VB_Name = "Module1"
Sub StockSorting()

Dim Ticker As String
Dim Total_Volume As LongLong
Dim TickerTable_Row As Integer
Dim Top_Row As Long
Dim Ann_Open_Price As Double
Dim Ann_Close_Price As Double
Dim Yearly_Change_OC As Double
Dim Yearly_Change_OC_Pct As Double
Dim Max_Pct As Double
Dim Min_Pct As Double
Dim Max_Vol As Double
Dim Max_Pct_Ticker As String
Dim Min_Pct_Ticker As String
Dim Max_Vol_Ticker As String

For Each ws In Worksheets

    ' Create variable to keep track of running row count in new summary table
    TickerTable_Row = 2
    Total_Volume = 0
    Max_Pct = 0
    Min_Pct = 0
    Max_Vol = 0
    
    ' Set and format column headers for new tables
    ws.Range("J1") = "Ticker"
    ws.Range("K1") = "Annual Volume"
    ws.Range("L1") = "Annual Price Change"
    ws.Range("M1") = "Annual Percent Change"
    ws.Range("O2") = "Greatest % Increase"
    ws.Range("O3") = "Greatest % Decrease"
    ws.Range("O4") = "Greatest Total Volume"
    ws.Range("P1") = "Ticker"
    ws.Range("Q1") = "Value"
    
    ' Find the last row in souce data sheet with data
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    LastColumn = ws.Cells(1, Columns.Count).End(xlToLeft).Column
    ' Initialize the Open Price for first ticker in list
    Ann_Open_Price = ws.Range("C2").Value
    
        For i = 2 To LastRow
        
            ' Compare cell to next cell. If not the same tickers...
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                ' Record Ticker, Print Unique Ticker to Cell J2 to start Summary Table
                Ticker = ws.Cells(i, 1).Value
                ws.Range("J" & TickerTable_Row).Value = Ticker
        
                ' Add last total_volume data point, print to Cell K2
                ' Store the EOY closing price for the current ticker
                Total_Volume = Total_Volume + ws.Cells(i, 7).Value
                ws.Range("K" & TickerTable_Row).Value = Total_Volume
                Ann_Close_Price = ws.Cells(i, 6).Value
    
                ' Yearly Change in Price
                Yearly_Change_OC = (Ann_Close_Price - Ann_Open_Price)
                ws.Range("L" & TickerTable_Row).Value = Yearly_Change_OC
        
                ' Yearly Change in Percentage
                If Ann_Open_Price <> 0 Then
                    Yearly_Change_OC_Pct = Yearly_Change_OC / Ann_Open_Price
                    ws.Range("M" & TickerTable_Row).Value = Yearly_Change_OC_Pct
                Else
                    MsgBox (Ticker & " in Row " & CStr(i) & ": Open Price equals " & Ann_Open_Price)
                End If
                
                ' Formatting
                If Yearly_Change_OC > 0 And Yearly_Change_OC_Pct > 0 Then
                    ws.Cells(TickerTable_Row, 12).Interior.ColorIndex = 4
                    ws.Cells(TickerTable_Row, 13).Interior.ColorIndex = 4
                Else
                    ws.Cells(TickerTable_Row, 12).Interior.ColorIndex = 3
                    ws.Cells(TickerTable_Row, 13).Interior.ColorIndex = 3
                End If
                
                ' Hard part with new Max, Min pct changes and volume
                If Yearly_Change_OC_Pct > Max_Pct Then
                    Max_Pct = Yearly_Change_OC_Pct
                    Max_Ticker = Ticker
                ElseIf Yearly_Change_OC_Pct < Min_Pct Then
                    Min_Pct = Yearly_Change_OC_Pct
                    Min_Ticker = Ticker
                End If
                
                If Total_Volume > Max_Vol Then
                    Max_Vol = Total_Volume
                    Max_Vol_Ticker = Ticker
                End If
                
                ws.Range("P2").Value = Max_Ticker
                ws.Range("Q2").Value = Max_Pct
                ws.Range("P3").Value = Min_Ticker
                ws.Range("Q3").Value = Min_Pct
                ws.Range("P4").Value = Max_Vol_Ticker
                ws.Range("Q4").Value = Max_Vol
                
                ' Start next entry on next row of summary table
                TickerTable_Row = TickerTable_Row + 1
                
                ' Capture Open Price of the next Ticker for calculation in next loop, Reset Total Volume
                Ann_Open_Price = ws.Cells((i + 1), 3).Value
                Total_Volume = 0
                Yearly_Change_OC = 0
                Yearly_Change_OC_Pct = 0
                
            Else
                
                ' Calculate total volume as we are looping through each row
                Total_Volume = Total_Volume + ws.Cells(i, 7).Value
            
            End If
        
        Next i
    
    ' Format ws.Cells
    ws.Range("L2:L" & TickerTable_Row).NumberFormat = "0.00"
    ws.Range("M2:M" & TickerTable_Row).NumberFormat = "0.00%"
    ws.Range("Q2:Q3").NumberFormat = "0.00%"
    ws.Range("J:J, K:K, L:L, M:M, O:O, P:P, Q:Q").EntireColumn.AutoFit

Next ws

End Sub




