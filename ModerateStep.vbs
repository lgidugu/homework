Attribute VB_Name = "ModerateStep"
Sub TestData()

    For Each ws In Worksheets
    
        Dim WorksheetName As String
    
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
                    
            WorksheetName = ws.Name
            'Declaring Vatiables to get Yearly Change, Percentage Change
            
            Dim TickerSymbol As String
            Dim TotalVolume As Double
            TotalVolume = 0
            Dim Year As Integer
            
            Dim YearlyChange As Double
            Dim InitOpenPr As Double
            InitOpenPr = Cells(2, 3).Value
            Dim YearEndPr As Double
            Dim PercentageChange As Double
            
            
            Dim Summary_Ticker_Symbol As Integer
            Summary_Ticker_Symbol = 2
                
                For I = 2 To LastRow
                    TotalVolume = TotalVolume + Cells(I, 7).Value
                   
                    
                    If Cells(I + 1, 1) <> Cells(I, 1) Then
                    TickerSymbol = Cells(I, 1).Value
                    YearEndPr = Cells(I, 3).Value
                    YearlyChange = YearEndPr - InitOpenPr
                    
                    
                        If InitOpenPr = 0 Then
                        PercentageChange = 1
                                                
                        Else
                        PercentageChange = (YearEndPr - InitOpenPr) / (InitOpenPr)
                        End If
                    InitOpenPr = Cells(I + 1, 3).Value
                    
                    Range("I" & Summary_Ticker_Symbol).Value = TickerSymbol
                    Range("J" & Summary_Ticker_Symbol).Value = YearlyChange
                    Range("K" & Summary_Ticker_Symbol).Value = PercentageChange
                    Range("L" & Summary_Ticker_Symbol).Value = TotalVolume
                    Summary_Ticker_Symbol = Summary_Ticker_Symbol + 1
                    TotalVolume = 0
                   End If
                                                         
           Next I
               
              'For loop for formatting cells
               For Each Summ_Ticker_Table In Worksheets
               LastRowIJKL = Summ_Ticker_Table.Cells(Rows.Count, 9).End(xlUp).Row
               
               
               For j = 2 To LastRowIJKL
                                          
                 If Cells(j, 10) > 0 Then
                    Cells(j, 10).Interior.ColorIndex = 4
                    Else
                    Cells(j, 10).Interior.ColorIndex = 3
                    End If
                    
                    Cells(j, 11).Style = "Currency"
                    Cells(j, 11).NumberFormat = "##.##%"
                Next j
                Next Summ_Ticker_Table
    Next ws
End Sub

