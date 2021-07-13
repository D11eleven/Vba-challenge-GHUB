Attribute VB_Name = "Module1"
Sub WallStreetBonus()
Attribute WallStreetBonus.VB_Description = "AlphaStocks Macro - Creates Summary Table"
Attribute WallStreetBonus.VB_ProcData.VB_Invoke_Func = "z\n14"

'Enables code to work through all sheets of workbook

Dim ws As Worksheet
For Each ws In Worksheets

'Declare Variables

Dim column As Integer
column = 1

Dim I As Double
I = 2


Dim Ticker As String
Dim Total_Volume As Double

'Initiate T_V
Total_Volume = 0

Dim Opening_Price As Double
'This assigns value to Opening_Price at beginning of sheet - there is no ticker
'change for first stock at the beginning

Opening_Price = ws.Cells(I, 3).Value


Dim Closing_Price As Double
Dim Price_Change As Double
Dim Price_Change_Percent As Double


'Assigns Headers and Autofits columns


ws.Cells(1, 9).Value = "TICKER"
ws.Cells(1, 10).Value = "YEAR CHANGE"
ws.Cells(1, 11).Value = "YEAR CHANGE (%)"
ws.Cells(1, 12).Value = "TOTAL VOLUME"

ws.Cells(1, 17).Value = "TICKER"
ws.Cells(1, 18).Value = "VALUE"

ws.Cells(2, 16).Value = "GREATEST % INCREASE"
ws.Cells(3, 16).Value = "GREATEST % DECREASE"
ws.Cells(4, 16).Value = "GREATEST TOTAL VOLUME"

Columns(9).AutoFit
Columns(10).AutoFit
Columns(11).AutoFit
Columns(12).AutoFit
Columns(16).AutoFit
Columns(17).AutoFit
Columns(18).AutoFit







'Keep track of location for each stock in the summary
Dim Summary_Table_Row As Double
Summary_Table_Row = 2



Dim LastRow As Double

'counts the number of rows
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row



'will count through every row til end
For I = 2 To LastRow


    'Check to see if row is still within same stock(ticker)
    If ws.Cells(I + 1, column).Value <> ws.Cells(I, column).Value Then
    
  
        'Set ticker symbol
        Ticker = ws.Cells(I, 1).Value
  
        'Add to the Volume Total
        Total_Volume = Total_Volume + ws.Cells(I, 7)
  
        'Print Ticker to Summary Table
        ws.Range("I" & Summary_Table_Row).Value = Ticker
        
        'Print Total Volume
        ws.Range("L" & Summary_Table_Row).Value = Total_Volume
  
       'Set closing_price as last entry of a group
        Closing_Price = ws.Cells(I, 6).Value
  
        
        
        'Calculate Price Change
        Price_Change = (Closing_Price - Opening_Price)
        
        
        
        'Print Price_Change to Summary Table
        ws.Range("J" & Summary_Table_Row).Value = Price_Change
        
        'some stocks had no opening price
        
        If Opening_Price = 0 Then
        Price_Change_Percent = 0
        Else
        
        'Price_Change Percent  formatted cell will * 100
        'Actual formula is (price change / opening price)* 100 for percentage
    
        Price_Change_Percent = (Price_Change / Opening_Price)
        End If
        
        
        'Print Price_Change_Percent to Summary Table
        ws.Range("K" & Summary_Table_Row).Value = Price_Change_Percent
        ws.Range("K" & Summary_Table_Row).NumberFormat = ".00%"
        
        '
        
        'Color code negative Red, zero and positive Green
        
        If Price_Change_Percent < 0 Then
        ws.Range("K" & Summary_Table_Row).Interior.ColorIndex = 3
        
        
        Else
            ws.Range("K" & Summary_Table_Row).Interior.ColorIndex = 4
        
        End If
  
        'Set opening_price as first entry of a group
        'Note: O_P will now use this value as opposed to very first stock which was
        'assigned value before loop
        
        Opening_Price = ws.Cells(I + 1, 3).Value
        
        'Using worksheet function to find MAX & MINs and then move to 2nd(Bonus)table
        
        ws.Range("R2").Value = Application.WorksheetFunction.Max(ws.Columns("K"))
        ws.Range("R2").NumberFormat = ".00%"
        ws.Range("R3").Value = Application.WorksheetFunction.Min(ws.Columns("K"))
        ws.Range("R3").NumberFormat = ".00%"
        ws.Range("R4").Value = Application.WorksheetFunction.Max(ws.Columns("L"))
        
       
       If ws.Range("K" & Summary_Table_Row).Value = ws.Range("R2").Value Then
       'find ticker symbol in
       ws.Range("Q2").Value = ws.Range("I" & Summary_Table_Row).Value
       
       End If
             
       
       If ws.Range("K" & Summary_Table_Row).Value = ws.Range("R3").Value Then
       'find ticker symbol in
       ws.Range("Q3").Value = ws.Range("I" & Summary_Table_Row).Value
       
       End If
       
             
       If ws.Range("L" & Summary_Table_Row).Value = ws.Range("R4").Value Then
       'find ticker symbol in
       ws.Range("Q4").Value = ws.Range("I" & Summary_Table_Row).Value
       
       End If
       
       
        
       
        'Reset Total Volume
         Total_Volume = 0
        
        
        'Add one to summary table
        Summary_Table_Row = Summary_Table_Row + 1
        
  
    'If cell is following row is same stock(ticker)
    Else
  
        'Add to Total_Volume
         Total_Volume = Total_Volume + ws.Cells(I, 7).Value
         
    End If
    
    'next row
     Next I
    
   Next ws

End Sub









'Unused Code Bad ideas DISREGARD past this point


 'BONUS IDEAS  DISREGARD
 
    
    
'Sub MinMaxTotals()

   ' For Each ws In Worksheets

    'Dim curr_row As Integer
    'Dim LastRow1 As Long
    
    'Dim MaxTicker As String
    'Dim MinTicker As String
    'Dim TotalVolTicker As String
    
    
   ' Dim Greatest_Price_Change_Percent_Increase As Double

    'Dim Greatest_Price_Change_Percent_Decrease As Double


    'Dim Greatest_Total_Volume As Double
    
    
    
    
    'Summary Table for Min,Max PercentChange and Highest Volume
'Dim Summary_Table1_Row As Double
'Summary_Table1_Row = 2
    
    'counts number of rows
    
    'LastRow1 = ws.Cells(Rows.Count, 11).End(xlUp).Row
    
    'For curr_row = 2 To LastRow1
   
    
    'If Cells(curr_row, 11).Value >= Cells(curr_row + 1, 11).Value Then
        'Greatest_Price_Change_Percent_Increase = Cells(curr_row, 11).Value
        'MaxTicker = Cells(curr_row, 9).Value
    'Else
       'Cells(curr_row + 1, 11).Value = Greatest_Price_Change_Percent_Increase
        'Cells(curr_row + 1, 9).Value = MaxTicker
    'Else
    ' What if a tie Cells(I
    
        'Range("Q2").Value = MaxTicker
       ' Range("R2").Value = Greatest_Price_Change_Percent_Increase
       'Range("R3").Value = Greatest_Price_Change_Percent_Decrease
        'Cells(2, 17).Value = MaxTicker
        'Cells(2, 18).Value = Greatest_Price_Change_Percent_Increase
    
    'End If
    
    
    
    'If Cells(curr_row, 11).Value <= Cells(i + 1, 11).Value Then
        'Cells(curr_row, 11).Value = Greatest_Price_Change_Percent_Decrease
        'Cells(curr_row, 9).Value = MinTicker
    'Else
       'Cells(curr_row + 1, 11).Value = Greatest_Price_Change_Percent_Decrease
        'Cells(curr_row + 1, 9).Value = MinTicker
        
        'what if a tie
        
        'Range("Q3").Value = MinTicker
        'Range("R3").Value = Greatest_Price_Change_Percent_Decrease
        
        'Cells(3, 17).Value = MinTicker
        'Cells(3, 18).Value = Greatest_Price_Change_Percent_Decrease
        
    'End If
    
   ' If Cells(curr_row, 12).Value >= Cells(curr_row + 1, 12).Value Then
       ' Cells(curr_row, 11).Value = Greatest_Total_Volume
        'Cells(curr_row, 9).Value = TotalVolTicker
   'Else
        'Cells(curr_row + 1, 11).Value = Greatest_Total_Volume
        'Cells(curr_row + 1, 9).Value = TotalVolTicker
        
        'Range("Q4").Value = MaxTicker
        'Range("R4").Value = Greatest_Total_Volume
        'Cells(4, 17).Value = TotalVolTicker
        'Cells(4, 18).Value = Greatest_Total_Volume
        
   ' End If
    
   ' Next curr_row
    
   ' Next ws
    
    
    
    
    
    'Max = Application.WorksheetFunction.Max(Columns("K"))
    'Range("J" & Summary_Table_Row1).Value = Price_Change
   
    'Min = Application.WorksheetFunction.Min(Columns("k"))
    'Range("J" & Summary_Table_Row1).Value = Price_Change


'End Sub

'END BONUS IDEAS

'Bonus

    'Dim MaxTicker As String
    'Dim MinTicker As String
    'Dim TotalVolTicker As String
    
    
   ' Dim Greatest_Price_Change_Percent_Increase As Double

    'Dim Greatest_Price_Change_Percent_Decrease As Double


    'Dim Greatest_Total_Volume As Double
    
'END BONUS STUFF


    


'if Price_Change_Percent > Price_Change_Percent then
'Greatest_Price_Change_Percent_Increase = Price_Change_Percent

  'Print Price_Change Percent  (*100)
        'Range("K" & Summary_Table_Row).Value = Price_Change
        
        ' Range("R2").Value = Greatest_Price_Change_Percent_Increase
       'Range("R3").Value = Greatest_Price_Change_Percent_Decrease
       
       'logic to find ticker
       'run a loop through table and find max, min, vol max previously found from worksheet function
       'above. when value is found, take the value from ticker column and print to table ... good luck!
       
        'Dim B As Integer
        'B = 2
        'Dim LastRow2 As Double
        
        'Dim Summary_Table_Row2 As Double
        'Summary_Table_Row2 = 2

        'counts the number of rows
        'LastRow2 = ws.Cells(Rows.Count, 11).End(xlUp).Row



        'will count through every row til end
        'For B = 2 To LastRow2


