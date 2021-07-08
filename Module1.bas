Attribute VB_Name = "Module1"
Sub AlphaStocks()
Attribute AlphaStocks.VB_Description = "AlphaStocks Macro - Creates Summary Table"
Attribute AlphaStocks.VB_ProcData.VB_Invoke_Func = "z\n14"


For Each ws In Worksheets

Dim column As Integer
column = 1

Dim Ticker As String
Dim Total_Volume As Double

Total_Volume = 0

Dim Opening_Price As Double
Dim Closing_Price As Double
Dim Price_Change As Double
Dim Price_Change_Percent As Double

Cells(1, 9).Value = "TICKER"
Cells(1, 10).Value = "YEAR CHANGE"
Cells(1, 11).Value = "YEAR CHANGE (%)"
Cells(1, 12).Value = "TOTAL VOLUME"

'Keep track of location for each stock in the summary
Dim Summary_Table_Row As Integer
Summary_Table_Row = 2

Dim LastRow As Long

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
  
        'Print Price_Change to Summary Table
        'Range("J" & Summary_Table_Row).Value = Price_Change
  
        'Print Price_Change_Percent to Summary Table
        'Range("K" & Summary_Table_Row).Value = Price_Change_Percent
  
        'Add one to summary table
        Summary_Table_Row = Summary_Table_Row + 1

  
        'Reset Total Volume
         Total_Volume = 0
         
  
    'If cell after following row is same stock(ticker)
    Else
  
        'Add to Total_Volume
         Total_Volume = Total_Volume + Cells(i, 7).Value
    
    
    End If
    
     Next i
    
   
   Next ws
   

End Sub
