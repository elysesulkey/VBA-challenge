Attribute VB_Name = "Module1"
Sub StockMarket()

'Set worksheet
Dim Current As Worksheet

'Loop through worksheets in workbook
For Each Current In Worksheets

'Free up space for faster processing
Application.ScreenUpdating = False
Application.DisplayAlerts = False

'Set variables
Dim Ticker, GreatestIncreaseTicker, GreatestDecreaseTicker, GreatestTotalVolumeTicker As String
Dim OpenPrice, ClosePrice, YearlyChangeValue, PercentChangeValue, GreatestIncreaseValue, GreatestDecreaseValue As Double
Dim StockVolume, TotalStockVolumeValue, GreatestTotalVolumeValue, LastRow3 As LongLong
Dim i, LastRow, LastRow2, SummaryTableRow As Long

Ticker = " "
GreatestIncreaseTicker = " "
GreatestDecreaseTicker = " "
OpenPrice = 0
ClosePrice = 0
YearlyChangeValue = 0
PercentChangeValue = 0
TotalStockVolumeValue = 0
GreatestIncreaseValue = 0
GreatestDecreaseValue = 0
GreatestTotalVolumeValue = 0
SummaryTableRow = 2

'Find last row of data
LastRow = Current.Cells(Rows.Count, 1).End(xlUp).Row

'Create summary table titles
Current.Range("I1:L1") = Array("Ticker", "Yearly Change", "Percent Change", "Total Stock Volume")
Current.Range("N2") = "Greatest % Increase"
Current.Range("N3") = "Greatest % Decrease"
Current.Range("N4") = "Greatest Total Volume"
Current.Range("O1:P1") = Array("Ticker", "Value")

'Set first ticker's open price
OpenPrice = Current.Cells(2, 3).Value
  
  'Loop through data
    For i = 2 To LastRow
    
        'Determine if ticker change
        If Current.Cells(i + 1, 1).Value <> Current.Cells(i, 1).Value Then
        
            'Calculate values for ticker
            Ticker = Current.Cells(i, 1).Value
            ClosePrice = Current.Cells(i, 6).Value
            
            YearlyChangeValue = YearlyChangeValue + (ClosePrice - OpenPrice)
            
                  If OpenPrice <> 0 Then
                            
                            PercentChangeValue = YearlyChangeValue / OpenPrice
                                
                        Else
            
                            PercentChangeValue = 0
                                        
                        End If
                    
            StockVolume = Current.Cells(i, 7).Value
        
            TotalStockVolumeValue = TotalStockVolumeValue + StockVolume
                            
             'Add values to summary table
            Current.Range("I" & SummaryTableRow).Value = Ticker
            Current.Range("J" & SummaryTableRow).Value = YearlyChangeValue
            Current.Range("K" & SummaryTableRow).Value = PercentChangeValue
            Current.Range("L" & SummaryTableRow).Value = TotalStockVolumeValue
                    
            SummaryTableRow = SummaryTableRow + 1
            
            'Reset values
            YearlyChangeValue = 0
            PercentChangeValue = 0
            TotalStockVolumeValue = 0
            OpenPrice = Current.Cells(i + 1, 3).Value
            
        Else
        'Add to current ticker volume
        
            StockVolume = Current.Cells(i, 7).Value

            TotalStockVolumeValue = TotalStockVolumeValue + StockVolume
            
        End If
         
    Next i

    'Format conditionals on summary table
    LastRow2 = Current.Cells(Rows.Count, 10).End(xlUp).Row
    
    For i = 2 To LastRow2
  
        If Current.Cells(i, 10).Value >= 0 Then
                  
            Current.Cells(i, 10).Interior.ColorIndex = 4
                  
        ElseIf Current.Cells(i, 10).Value < 0 Then
                 
            Current.Cells(i, 10).Interior.ColorIndex = 3
                 
        End If
        
        Current.Cells(i, 11).NumberFormat = "0.00%"
        
    Next i

'Find values for bonus summary table
LastRow3 = Current.Cells(Rows.Count, 9).End(xlUp).Row
                     
 For i = 2 To LastRow3
 
    'Compare values
    If Current.Cells(i + 1, 11) > GreatestIncreaseValue Then
    
        GreatestIncreaseTicker = Current.Cells(i + 1, 9).Value
        GreatestIncreaseValue = Current.Cells(i + 1, 11).Value
        
        Current.Range("O2") = GreatestIncreaseTicker
        Current.Range("P2") = GreatestIncreaseValue
    
    ElseIf Current.Cells(i + 1, 11) < GreatestDecreaseValue Then
            
        GreatestDecreaseTicker = Current.Cells(i + 1, 9).Value
        GreatestDecreaseValue = Current.Cells(i + 1, 11).Value
                    
        Current.Range("O3") = GreatestDecreaseTicker
        Current.Range("P3") = GreatestDecreaseValue

     ElseIf Current.Cells(i + 1, 12) > GreatestTotalVolumeValue Then
    
        GreatestTotalVolumeTicker = Current.Cells(i + 1, 9).Value
        GreatestTotalVolumeValue = Current.Cells(i + 1, 12).Value
        
        Current.Range("O4") = GreatestTotalVolumeTicker
        Current.Range("P4") = GreatestTotalVolumeValue
     
     Else
     'Print summary table values
     
        Current.Range("O2") = GreatestIncreaseTicker
        Current.Range("P2") = GreatestIncreaseValue
        Current.Range("O3") = GreatestDecreaseTicker
        Current.Range("P3") = GreatestDecreaseValue
        Current.Range("O4") = GreatestTotalVolumeTicker
        Current.Range("P4") = GreatestTotalVolumeValue

    End If
    
    'Format summary table numbers
    Current.Cells(2, 16).NumberFormat = "0.00%"
    Current.Cells(3, 16).NumberFormat = "0.00%"

 Next i
 
Application.ScreenUpdating = True
Application.DisplayAlerts = True
 
Next Current

End Sub



