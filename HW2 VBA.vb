'(Module 1) Aggreates ticker and volume
Sub GO()
'declare variables
  Dim Total As Double
  Dim SummaryRow As Integer
  Dim LastRow As Long

    
    
    
'set values
    Cells(1, 9).Value = "Ticker"
    Cells(1, 9).Font.FontStyle = "Bold"
    Cells(1, 10).Value = "Volume"
    Cells(1, 10).Font.FontStyle = "Bold"
    SummaryRow = 2
    Total = 0
    LastRow = ActiveSheet.UsedRange.Rows.Count
    
   
    
'Set loop to check for <> tickers
    For i = 2 To LastRow
        If Cells(i + 1, 1) <> Cells(i + 1 - 1, 1) Then
            'Post ticker
            Cells(SummaryRow, "I") = Cells(i - 1, 1).Value
            'Post volume
            Cells(SummaryRow, "J") = Total
            'Reset counter
            Total = 0
            'Next row
            SummaryRow = SummaryRow + 1
            Total = Total + Cells(i, "G").Value
        End If
            'Add to volume total if = tickers
            Total = Total + Cells(i, "G").Value
    Next i

End Sub

-----------------------------------------------------------------------------------------------------

'(Module 2)Resets worksheet
Sub RESET()
    'Clear Content
    Range("I:I").ClearContents
    Range("J:J").ClearContents

End Sub

------------------------------------------------------------------------------------------------------

'(Moduel 3)Runs module 1 on all worksheets
Sub Run_All()
    Dim ws_num As Integer
    Dim z As Integer
	
	
    'set ws_num
    ws_num = ThisWorkbook.Worksheets.count
		
    
    'loop
    For z = 1 To ws_num
        'activate worksheet
        Sheets(z).Activate
		 
        'declare variables
          Dim Total As Double
          Dim SummaryRow As Integer
          Dim LastRow As Long
              
                 
        'set values
            Cells(1, 9).Value = "Ticker"
            Cells(1, 9).Font.FontStyle = "Bold"
            Cells(1, 10).Value = "Volume"
            Cells(1, 10).Font.FontStyle = "Bold"
            SummaryRow = 2
            Total = 0
            LastRow = ActiveSheet.UsedRange.Rows.Count
            
           
            
        'Set loop to check for <> tickers
            For i = 2 To LastRow
                If Cells(i + 1, 1) <> Cells(i + 1 - 1, 1) Then
                    'Post ticker
                    Cells(SummaryRow, "I") = Cells(i - 1, 1).Value
                    'Post volume
                    Cells(SummaryRow, "J") = Total
                    'Reset counter
                    Total = 0
                    'Next row
                    SummaryRow = SummaryRow + 1
                    Total = Total + Cells(i, "G").Value
                End If
                    'Add to volume total if = tickers
                    Total = Total + Cells(i, "G").Value
            Next i

    Next z
   
   
End Sub

