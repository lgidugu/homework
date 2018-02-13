Attribute VB_Name = "ChallengeStep"
 Sub WorksheetLoop()

    'Changes made from HardStep Code, so the code can be used for all the worksheets
    For Each ws In Worksheets
    
        Dim WorksheetName As String
    
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
                    
            WorksheetName = ws.Name
            
            'Column Lablels  for all newly added columns
            
            ws.Range("I1").Value = "<Ticker Symbol>"
            ws.Range("J1").Value = "Yearly Change"
            ws.Range("K1").Value = "Percent Change"
            ws.Range("L1").Value = "<Total Stock Volume>"
            ws.Range("N2").Value = "Greatest % Increase"
            ws.Range("N3").Value = "Greatest % Decrease"
            ws.Range("N4").Value = "Greatest Total Volume"
            ws.Range("O1").Value = "Ticker"
            ws.Range("P1").Value = "Value"
            
            'Declaring Variables
            Dim TickerSymbol As String
            Dim TotalVolume As Double
            TotalVolume = 0
            Dim Year As Integer
            
            Dim YearlyChange As Double
            
            Dim InitOpenPr As Double
            InitOpenPr = ws.Cells(2, 3).Value
            
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
                    TotalVolume = TotalVolume + ws.Cells(I, 7).Value
                   
                    
                    If Cells(I + 1, 1) <> Cells(I, 1) Then
                        TickerSymbol = ws.Cells(I, 1).Value
                        YearEndPr = ws.Cells(I, 3).Value
                        YearlyChange = YearEndPr - InitOpenPr
                       
                        If InitOpenPr = 0 Then
                        PercentageChange = 1
                                                
                        Else
                        PercentageChange = (YearEndPr - InitOpenPr) / (InitOpenPr)
                        
                        End If
                        InitOpenPr = ws.Cells(I + 1, 3).Value
                        
                            If PercentageChange > GrtPerctch Then
                                GrtPerctch = PercentageChange
                                GrtPerctch_ticker = ws.Cells(I, 1)
                            End If
                                        
                            If PercentageChange < PerctDecr Then
                                PerctDecr = PercentageChange
                                PerctDecr_ticker = ws.Cells(I, 1)
                            End If
                        ws.Range("I" & Summary_Ticker_Symbol).Value = TickerSymbol
                        ws.Range("J" & Summary_Ticker_Symbol).Value = YearlyChange
                        ws.Range("K" & Summary_Ticker_Symbol).Value = PercentageChange
                        ws.Range("L" & Summary_Ticker_Symbol).Value = TotalVolume
                    
                        If TotalVolume > GrtTotalVolume Then
                            GrtTotalVolume = TotalVolume
                            GrtTotalVolume_ticker = ws.Cells(I, 1)
                        End If
                    
                        Summary_Ticker_Symbol = Summary_Ticker_Symbol + 1
                        TotalVolume = 0
                   End If
                                                         
                Next I
                    ws.Cells(4, 16) = GrtTotalVolume
                    ws.Cells(4, 15) = GrtTotalVolume_ticker
                    ws.Cells(3, 16) = PerctDecr
                    ws.Cells(3, 15) = PerctDecr_ticker
                    ws.Cells(2, 16) = GrtPerctch
                    ws.Cells(2, 15) = GrtPerctch_ticker
                    
                      ws.Cells(3, 16).Style = "Currency"
                      ws.Cells(3, 16).NumberFormat = "##.##%"
                      ws.Cells(2, 16).Style = "Currency"
                      ws.Cells(2, 16).NumberFormat = "##.##%"
                    
                    
               LastRowIJKL = ws.Cells(Rows.Count, 9).End(xlUp).Row
               
               
               For j = 2 To LastRowIJKL
                                          
                 If ws.Cells(j, 10) > 0 Then
                    ws.Cells(j, 10).Interior.ColorIndex = 4
                    Else
                    ws.Cells(j, 10).Interior.ColorIndex = 3
                    End If
                    
                    ws.Cells(j, 11).Style = "Currency"
                    ws.Cells(j, 11).NumberFormat = "##.##%"
                Next j
    Next ws
    
End Sub



