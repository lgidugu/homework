Attribute VB_Name = "EasyStep"
Sub TestData()

    For Each ws In Worksheets
    
        Dim WorksheetName As String
    
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
                    
            WorksheetName = ws.Name
            'Declaring Variables
            Dim TickerSymbol As String
            Dim TotalVolume As Double
            TotalVolume = 0
            Dim Year As Integer
            
            'For loop to identify different tickers and total volume for each ticker symbol
            
            Dim Summary_Ticker_Symbol As Integer
            Summary_Ticker_Symbol = 2
                
                For I = 2 To LastRow
                    TotalVolume = TotalVolume + Cells(I, 7).Value
                    
                    If Cells(I + 1, 1) <> Cells(I, 1) Then
                    TickerSymbol = Cells(I, 1).Value
                                    
                    Range("I" & Summary_Ticker_Symbol).Value = TickerSymbol
                    Range("J" & Summary_Ticker_Symbol).Value = TotalVolume
                    Summary_Ticker_Symbol = Summary_Ticker_Symbol + 1
                    TotalVolume = 0
                   End If
                                                         
           Next I
    Next ws
End Sub
    
    

