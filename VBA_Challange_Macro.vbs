Attribute VB_Name = "Module1"
Sub stocks()

Dim TickerSymbol As String
Dim OpenPrice As Double
Dim ClosePrice As Double
Dim DateOpen As Integer
Dim DateClose As Integer
Dim LastRow As Long
LastRow = Range("A" & Rows.Count).End(xlUp).Row
Dim TotalVolume As Double
Dim Diff As Double
Dim ConditionValue As Double
Dim ws As Worksheet

'loops through all worksheets in the workbook and applies the ws prefix to all cell locations
For Each ws In ActiveWorkbook.Worksheets
ws.Activate

  

'Keep track of the location for each ticker in the summary table
Dim Summary_Table_Row As Integer
Summary_Table_Row = 2
TotalVolume = 0
BeginingTicker = True
  
'Add Additional headers for calculations
Range("I1").Value = ("Ticker")
Range("J1").Value = ("Yearly Change")
Range("K1").Value = ("Percent Change")
Range("L1").Value = ("Total Stock Volume")
Range("O1").Value = ("Ticker")
Range("P1").Value = ("Value")
Range("N2").Value = "Greatest % increase"
Range("N3").Value = "Greatest % decrease"
Range("N4").Value = "Greatest Total Volume"
  
  ' Loop through all tickers
  For i = 2 To LastRow
  
    'Get the stock open price for first counter
        If BeginingTicker = True Then
            OpenPrice = Cells(i, 3).Value
        End If
    
    'Print Percentage Change into Column K
    If OpenPrice And ClosePrice > 0 Then
        Range("K" & Summary_Table_Row).Value = (ClosePrice - OpenPrice) / (OpenPrice)
    End If
        
                
    ' Check if we are still within the same Stock ticker
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
       
         ' Set the Stock name
         TickerSymbol = Cells(i, 1).Value
        
        ' Print the TickerSymbol in the Summary Table
          Range("I" & Summary_Table_Row).Value = TickerSymbol
       
       ' Print the OpenPrice to the Summary Table Checking OpenPrice
         'Range("M" & Summary_Table_Row).Value = OpenPrice
         
        ' Print Difference price into Column L
         Range("J" & Summary_Table_Row).Value = ClosePrice - OpenPrice
         
         ' Color the cells with Green for Positive and Red For Negative
                If Range("J" & Summary_Table_Row).Value > 0 Then
                    Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                Else
                    Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                End If
       
         'Add one to the summary table row
          Summary_Table_Row = Summary_Table_Row + 1
          
          'Reset Volume to 0
           TotalVolume = 0
           
        'Reset biginning stock flag
            BeginingTicker = True
    
    ' If the cell immediately following a row is the same Symbol...
    Else
        
            ClosePrice = Cells(i + 1, 6).Value
                 
           'Add to the ClosePrice Column K checking results
           'Range("N" & Summary_Table_Row).Value = ClosePrice
            
            TotalVolume = TotalVolume + Cells(i, 7).Value
            Range("L" & Summary_Table_Row).Value = TotalVolume
              
                     
            'Reset biginning stock flag
            BeginingTicker = False
      
        End If
   
   
   
  Next i
 
    ' Start Loop For Final Results
            For i = 2 To LastRow
                If Range("K" & i).Value > Range("P2").Value Then
                    Range("P2").Value = Range("K" & i).Value
                    Range("O2").Value = Range("I" & i).Value
                End If

                If Range("K" & i).Value < Range("P3").Value Then
                    Range("P3").Value = Range("K" & i).Value
                    Range("O3").Value = Range("I" & i).Value
                End If

                If Range("L" & i).Value > Range("P4").Value Then
                    Range("P4").Value = Range("L" & i).Value
                    Range("O4").Value = Range("I" & i).Value
                End If

            Next i
    
            'Modify the cell with percentage formating
            Range("P2").NumberFormat = "0.00%"
            Range("P3").NumberFormat = "0.00%"
    
 
Next ws


End Sub
