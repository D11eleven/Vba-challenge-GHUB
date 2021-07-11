Attribute VB_Name = "Module1"
Sub AlphaStocks()
Attribute AlphaStocks.VB_Description = "AlphaStocks Macro - Creates Summary Table"
Attribute AlphaStocks.VB_ProcData.VB_Invoke_Func = "z\n14"


For Each ws In Worksheets

Dim column As Integer
column = 1

Dim i As Double
i = 2

Dim Ticker As String
Dim Total_Volume As Double


Total_Volume = 0

Dim Opening_Price As Double
Opening_Price = Cells(i, 3).Value


Dim Closing_Price As Double
Dim Price_Change As Double
Dim Price_Change_Percent As Double


'if Price_Change_Percent > Price_Change_Percent then
'Greatest_Price_Change_Percent_Increase = Price_Change_Percent


Cells(1, 9).Value = "TICKER"
Cells(1, 10).Value = "YEAR CHANGE"
Cells(1, 11).Value = "YEAR CHANGE (%)"
Cells(1, 12).Value = "TOTAL VOLUME"

'Cells(1, 17).Value = "TICKER"
'Cells(1, 18).Value = "VALUE"

'Cells(2, 16).Value = "GREATEST % INCREASE"
'Cells(3, 16).Value = "GREATEST % DECREASE"
'Cells(4, 16).Value = "GREATEST TOTAL VOLUME"

Columns(9).AutoFit
Columns(10).AutoFit
Columns(11).AutoFit
Columns(12).AutoFit
'Columns(16).AutoFit
'Columns(17).AutoFit
'Columns(18).AutoFit







'Keep track of location for each stock in the summary
Dim Summary_Table_Row As Double
Summary_Table_Row = 2



Dim LastRow As Double

'counts the number of rows
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row



'will count through every row til end
For i = 2 To LastRow

   

    

    'Check to see if still within same stock(ticker)
    If Cells(i + 1, column).Value <> Cells(i, column).Value Then
    
    
    
    
  
        'Set ticker symbol
        Ticker = Cells(i, 1).Value
  
        'Add to the Volume Total
        Total_Volume = Total_Volume + Cells(i, 7)
  
        'Print Ticker to Summary Table
        Range("I" & Summary_Table_Row).Value = Ticker
        
        'Print Total Volume
        Range("L" & Summary_Table_Row).Value = Total_Volume
  
       'Set closing_price as last entry of a group
        Closing_Price = Cells(i, 6).Value
  
        
        
        'Calculate Price Change
        Price_Change = (Closing_Price - Opening_Price)
        
        
        
        'Print Price_Change to Summary Table
        Range("J" & Summary_Table_Row).Value = Price_Change
        
        If Opening_Price = 0 Then
        Price_Change_Percent = 0
        Else
        
        'Price_Change Percent  (*100)
        Price_Change_Percent = (Price_Change / Opening_Price) * 100
        End If
        
        
        
        
        
        
        
        
        'Print Price_Change Percent  (*100)
        'Range("K" & Summary_Table_Row).Value = Price_Change
        
        'Print Price_Change_Percent to Summary Table
        Range("K" & Summary_Table_Row).Value = Price_Change_Percent
        Range("K" & Summary_Table_Row).NumberFormat = "0.00"
        
        '
        
        
        
        If Price_Change_Percent < 0 Then
        Range("K" & Summary_Table_Row).Interior.ColorIndex = 3
        
        
        Else
            Range("K" & Summary_Table_Row).Interior.ColorIndex = 4
        
        End If
  
        'Set opening_price as first entry of a group
        Opening_Price = Cells(i + 1, 3).Value
        
        
        
         
          
        
        
       
        'Reset Total Volume
         Total_Volume = 0
        
        
        
        
        
        
        'Add one to summary table
        Summary_Table_Row = Summary_Table_Row + 1

  
        
         
  
    'If cell after following row is same stock(ticker)
    Else
  
        'Add to Total_Volume
         Total_Volume = Total_Volume + Cells(i, 7).Value
         
    
   
    
    End If
    
     Next i
    
   Next ws
   

End Sub
