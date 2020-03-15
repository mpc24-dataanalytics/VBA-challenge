VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub HW_VBA_WallStreet_StockM_Analysis()

        
        'Create a scrip that will loop through all stocks for one year
         'Dim ws As Worksheet
    'Set variable to hold total volume of stocks
        Dim total_volume As Double
        Range("I1").Value = "Ticker"
    Range("J1").Value = "Total Stock Volume"
    
          rowcount = Cells(Rows.Count, "A").End(xlUp).row
    For i = 2 To rowcount
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
     total_volume = total_volume + Cells(i, 7)
     
     Range("I" & 2 + j).Value = Cells(i, 1).Value
            ' Print total
            Range("J" & 2 + j).Value = total_volume
            ' Reset Total
            total_volume = 0
            ' Move to next row
            j = j + 1
     Else
     total_volume = total_volume + Cells(i, 7)
     
     End If
     Next i
     End Sub
     
        'Start loop
         'For Each ws In Worksheets
    
        'Create column labels to summarize data
            'ws.Cells(1, 7).Value = "Ticker"
            'ws.Cells(1, 7).Value = "Yearly Change"
            'ws.Cells(1, 7).Value = "Percent Change"
            'ws.Cells(1, 7).Value = "Total Stock Volume"

        'Set variable to hold the ticker symbol
         'Dim ticker_symbol As String

        'Set variable to hold the yearly change and percent change
    


        'Set variable to hold total volume of stocks
       
            
            

