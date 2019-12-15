Sub StockAnalysis():
'loop through worksheets
 
For Each ws In Worksheets

'Create a worksheet

Dim Worksheet As String

'Grab worksheet Name
Worksheet = ws.Name

MsgBox (Worksheet)

'declare variables

Dim tickername As String
Dim Openstock As Double
Dim Closestock As Double
Dim YearlyChange As Double
Dim PercentChange As Variant
Dim TotalStockVolume As Variant
Dim Rownumber As Integer

Rownumber = 2
TotalStockVolume = 0
Openstock = ws.Cells(2, 3).Value
Closestock = 0
YearlyChange = 0
PercentChange = 0

'Assign summary column names
ws.Range("H1").Value = "Ticker"
ws.Range("I1").Value = "Yearly Change"
ws.Range("J1").Value = "Percent Change"
ws.Range("K1").Value = "Total Stock Volume"

'Assign summary maximum values column names
ws.Range("N1").Value = "Ticker"
ws.Range("O1").Value = "Value"
ws.Range("M2").Value = "Greatest % Increase"
ws.Range("M3").Value = "Greatest % Decrease"
ws.Range("M4").Value = "Greatest Total Volume"

'determine last row in the summary table for our range


'delare value to hold ranges
Dim Percrange As Range
Dim maxtotalrange As Range

'Determine ranges
'Range("C3:H14").NumberFormat = "#'###'##0"
Set Percrange = ws.Range("J:J")
Percrange.NumberFormat = "0"
Set maxtotalrange = ws.Range("K:K")

'declare variables to hold maximum values
Dim MinPercChange As Double
Dim MaxPercChange As Double
Dim maxtotalvolume As Variant

'determine max values
MinPercChange = Application.WorksheetFunction.Min(Percrange)
MaxPercChange = Application.WorksheetFunction.Max(Percrange)
maxtotalvolume = Application.WorksheetFunction.Max(maxtotalrange)

'send maximum values to excel
ws.Range("O2").Value = MaxPercChange
ws.Range("O3").Value = MinPercChange
ws.Range("O4").Value = maxtotalvolume

ws.Range("O2").NumberFormat = "0.00%"
ws.Range("O3").NumberFormat = "0.00%"

'tickername for the greatest % increase, greatest % decrease and greatest total volume
'declare variables to hold gratest change ticker names
Dim MinPercChangeticker As String
Dim MaxPercChangeticker As String
Dim maxtotalvolumeticker As String

'Last row
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

For j = 2 To LastRow
'ws.Cells(j, 10).NumberFormat = "0"
'ws.Cells(2, 15).NumberFormat = "0"

'tickername for the greatest % increase
If ws.Cells(j, 10).Value = ws.Cells(2, 15).Value Then
 
 MaxPercChangeticker = ws.Cells(j, 8).Value
 

ws.Range("N2").Value = MaxPercChangeticker
'tickername for the greatest % decrease

ElseIf ws.Cells(j, 10).Value = ws.Cells(3, 15).Value Then

 MinPercChangeticker = ws.Cells(j, 8).Value
 ws.Range("N3").Value = MinPercChangeticker
 
'tickername for the greatest volume total
ElseIf ws.Cells(j, 11).Value = ws.Cells(4, 15).Value Then

 maxtotalvolumeticker = ws.Cells(j, 8).Value
 ws.Range("N4").Value = maxtotalvolumeticker

End If

Next j


    

 'Loop through each row to determine closing stock and ticker name
 For i = 2 To LastRow
 
  
     
'Find row where ticker name changes

    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
    
        'find ticker name
        tickername = ws.Cells(i, 1).Value
            
        'Calculate total volume of stock
        TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value

        'Closing price at the end of that year
         Closestock = ws.Cells(i, 6).Value
         
        'Change from opening price at the beginning of a given year to the closing price at the end of that year
         YearlyChange = Closestock - Openstock
                
            'Use if to eliminate division by 0 overflow errors
             If YearlyChange <> 0 And Openstock <> 0 Then
                        
             PercentChange = (YearlyChange / Openstock)
             Else
             PercentChange = 0
             End If
        
           
           'send ticker value to spreadsheet
            ws.Range("H" & Rownumber).Value = tickername
            
            'send TotalStockVolume value to spreadsheet
            ws.Range("K" & Rownumber).Value = TotalStockVolume
            
            'send YearlyChange value to spreadsheet
            ws.Range("I" & Rownumber).Value = YearlyChange
            
             'send PercentChange value to spreadsheet
            ws.Range("J" & Rownumber).Value = PercentChange
            
            'Format percentage
            ws.Range("J" & Rownumber).NumberFormat = "0.00%"
            
            'highlight positive change in green and negative change in red
            
            If YearlyChange < 0 Then
                ws.Range("I" & Rownumber).Interior.ColorIndex = 3
            Else
               ws.Range("I" & Rownumber).Interior.ColorIndex = 4
               End If
        
                 
               
            'Reset because we are going to start counting another ticker
            TotalStockVolume = 0
            Closestock = 0
            Openstock = ws.Cells(i + 1, 3).Value
            YearlyChange = 0
            PercentChange = 0
                 
            Rownumber = Rownumber + 1
         
    
       
        Else
                 'Calculate total volume of stock if condition is not met
                  TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value

                   Closestock = 0
                       
                       
                        
       End If
       
       
       Next i
                        
Next ws

End Sub

