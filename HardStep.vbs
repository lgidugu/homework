Attribute VB_Name = "HardStep"
Sub TestData()

    For Each ws In Worksheets
     
        Dim WorksheetName As String
    
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
                    
            WorksheetName = ws.Name
            
           'New column labels for all the worksheets
            Range("I1").Value = "<Ticker Symbol>"
            Range("J1").Value = "Yearly Change"
            Range("K1").Value = "Percent Change"
            Range("L1").Value = "<Total Stock Volume>"
            Range("N2").Value = "Greatest % Increase"
            Range("N3").Value = "Greatest % Decrease"
            Range("N4").Value = "Greatest Total Volume"
            Range("O1").Value = "Ticker"
            Range("P1").Value = "Value"
            
            'Declaring variables
            Dim TickerSymbol As String
            Dim TotalVolume As Double
            TotalVolume = 0
            Dim Year As Integer
            
            Dim YearlyChange As Double
            
            Dim InitOpenPr As Double
            InitOpenPr = Cells(2, 3).Value
            
            Dim YearEndPr As Double
            Dim PercentageChange As Double
            
            Dim GrtPerctch As Double
            GrtPerctch = 0
            Dim GrtPerctch_ticker As String
            GrtPerctch_ticker = " "
            
            Dim PerctDecr As Double
            PerctDecr = 0
            Dim PerctDecr_ticker As String
            PerctDecr_ticker = " "
            
            Dim GrtTotalVol As Double
            GrtTotalVolume = 0
            Dim GrtTotalVolume_ticker As String
            GrtTotalVolume_ticker = " "
            
            
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
                        
                        InitOpenPr = ws.Cells(I + 1, 3).Value
                            
                            If PercentageChange > GrtPerctch Then
                                GrtPerctch = PercentageChange
                                GrtPerctch_ticker = Cells(I, 1)
                            End If
                                        
                            If PercentageChange < PerctDecr Then
                                PerctDecr = PercentageChange
                                PerctDecr_ticker = Cells(I, 1)
                            End If
                        Range("I" & Summary_Ticker_Symbol).Value = TickerSymbol
                        Range("J" & Summary_Ticker_Symbol).Value = YearlyChange
                        Range("K" & Summary_Ticker_Symbol).Value = PercentageChange
                        Range("L" & Summary_Ticker_Symbol).Value = TotalVolume
                    
                        If TotalVolume > GrtTotalVolume Then
                            GrtTotalVolume = TotalVolume
                            GrtTotalVolume_ticker = Cells(I, 1)
                        End If
                    
                        Summary_Ticker_Symbol = Summary_Ticker_Symbol + 1
                        TotalVolume = 0
                   End If
                                                         
                Next I
                    Cells(4, 16) = GrtTotalVolume
                    Cells(4, 15) = GrtTotalVolume_ticker
                    Cells(3, 16) = PerctDecr
                    Cells(3, 15) = PerctDecr_ticker
                    Cells(2, 16) = GrtPerctch
                    Cells(2, 15) = GrtPerctch_ticker
                    
                      Cells(3, 16).Style = "Currency"
                      Cells(3, 16).NumberFormat = "##.##%"
                      Cells(2, 16).Style = "Currency"
                      Cells(2, 16).NumberFormat = "##.##%"
                    
                    
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


