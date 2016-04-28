Module ScheduledTransactions
    Public ArbTicker As String = ""
    Public SetUpArb As Boolean = False
    Public ArbUnderway As Boolean = False
    Public ArbStockAsk As Double = 0
    Public ArbOptionAsk As Double = 0
    Public ArbNumberOfStocks As Double = 0
    Public ArbNumberOfOptions As Double = 0

    Public CurrentTotalValue As Double = 0
    Public ExcessValue As Double = 0

    Public Function NumberInIP(underlier)
        If IsInIP(underlier) Then
            For Each myRow As DataRow In myDataSet.Tables("InitialPositionTable").Rows
                If myRow("Symbol").Trim() = underlier Then
                    Return myRow("Units")
                End If
            Next
        Else
            Return Nothing
        End If
        Return Nothing
    End Function

    'Public Function IstheArbLegal(underlier)
    '    If IsInIP(underlier) = True Then
    '        If NumberInIP(underlier) < 0 Then
    '            Return 
    '        End If
    ''    End If


    'End Function

    'ExecuteAlgoTransaction moved to Smart Hedger and modified
    Public Sub DoScheduledTransactions()

        Globals.Dashboard.ExecuteOptionTransactionBtn.Enabled = False
        Globals.Dashboard.ExecuteStockTransactionBtn.Enabled = False


        If NeedMoreCapital = True And MarginTripped = False Then
            'AAPL
            If currentDate.ToShortDateString = "5/11/2016" Or currentDate.ToShortDateString = "8/10/2016" Then
                ArbNumberOfOptions = Math.Min((100000 / GetAsk("AAPL_COCTB", currentDate)), Math.Abs((ExcessMargin - 1000000) / GetAsk("AAPL_COCTB", currentDate)))
                ExecuteAlgoTransaction("SellShort", ArbNumberOfOptions, "AAPL_COCTB")
                ArbNumberOfStocks = ArbNumberOfOptions / (1 / (Math.Abs(CalcDelta("AAPL_COCTB", currentDate))))
                ExecuteAlgoTransaction("Buy", ArbNumberOfStocks, "AAPL")
                ArbUnderway = True
                lastTransactionDate = currentDate
                CAccount = CAccountAT
                margin = marginAT
            End If



            'BLK
            If currentDate.ToShortDateString = "6/21/2016" Or currentDate.ToShortDateString = "9/20/2016" Then
                ArbNumberOfOptions = Math.Min((100000 / GetAsk("BLK_COCTB", currentDate)), Math.Abs((ExcessMargin - 1000000) / GetAsk("BLK_COCTB", currentDate)))
                ExecuteAlgoTransaction("SellShort", ArbNumberOfOptions, "BLK_COCTB")
                ArbNumberOfStocks = ArbNumberOfOptions / (1 / (Math.Abs(CalcDelta("BLK_COCTB", currentDate))))
                ExecuteAlgoTransaction("Buy", ArbNumberOfStocks, "BLK")
                ArbUnderway = True
                lastTransactionDate = currentDate
                CAccount = CAccountAT
                margin = marginAT
            End If



            'HSY
            If currentDate.ToShortDateString = "6/13/2016" Or currentDate.ToShortDateString = "9/12/2016" Then
                ArbNumberOfOptions = Math.Min((100000 / GetAsk("HSY_COCTB", currentDate)), Math.Abs((ExcessMargin - 1000000) / GetAsk("HSY_COCTB", currentDate)))
                ExecuteAlgoTransaction("SellShort", ArbNumberOfOptions, "HSY_COCTB")
                ArbNumberOfStocks = ArbNumberOfOptions / (1 / (Math.Abs(CalcDelta("HSY_COCTB", currentDate))))
                ExecuteAlgoTransaction("Buy", ArbNumberOfStocks, "HSY")
                ArbUnderway = True
                lastTransactionDate = currentDate
                CAccount = CAccountAT
                margin = marginAT
            End If



            'NKE: Risky Weekend Arb
            'If currentDate.ToShortDateString = "7/1/2016" Or currentDate.ToShortDateString = "9/30/2016" Then
            '    ArbNumberOfOptions = Math.Min((100000 / GetAsk("NKE_COCTB", currentDate)), Math.Abs((ExcessMargin - 1000000) / GetAsk("NKE_COCTB", currentDate)))
            '    ExecuteAlgoTransaction("SellShort", ArbNumberOfOptions, "NKE_COCTB")
            '    ArbNumberOfStocks = ArbNumberOfOptions / (1 / (Math.Abs(CalcDelta("NKE_COCTB", currentDate))))
            '    ExecuteAlgoTransaction("Buy", ArbNumberOfStocks, "NKE")
            '    ArbUnderway = True
            '    lastTransactionDate = currentDate
            '    CAccount = CAccountAT
            '    margin = marginAT
            'End If



            ''WMT: Risky Weekend Arb
            'If currentDate.ToShortDateString = "6/3/2016" Or currentDate.ToShortDateString = "9/2/2016" Then
            '    ArbNumberOfOptions = Math.Min((100000 / GetAsk("WMT_COCTB", currentDate)), Math.Abs((ExcessMargin - 1000000) / GetAsk("WMT_COCTB", currentDate)))
            '    ExecuteAlgoTransaction("SellShort", ArbNumberOfOptions, "WMT_COCTB")
            '    ArbNumberOfStocks = ArbNumberOfOptions / (1 / (Math.Abs(CalcDelta("WMT_COCTB", currentDate))))
            '    ExecuteAlgoTransaction("Buy", ArbNumberOfStocks, "WMT")
            '    ArbUnderway = True
            '    lastTransactionDate = currentDate
            '    CAccount = CAccountAT
            '    margin = marginAT
            'End If



            'XOM
            If currentDate.ToShortDateString = "6/8/2016" Or currentDate.ToShortDateString = "9/7/2016" Then
                ArbNumberOfOptions = Math.Min((100000 / GetAsk("XOM_COCTB", currentDate)), Math.Abs((ExcessMargin - 1000000) / GetAsk("XOM_COCTB", currentDate)))
                ExecuteAlgoTransaction("SellShort", ArbNumberOfOptions, "XOM_COCTB")
                ArbNumberOfStocks = ArbNumberOfOptions / (1 / (Math.Abs(CalcDelta("XOM_COCTB", currentDate))))
                ExecuteAlgoTransaction("Buy", ArbNumberOfStocks, "XOM")
                ArbUnderway = True
                lastTransactionDate = currentDate
                CAccount = CAccountAT
                margin = marginAT
            End If



            'For Each myRow As DataRow In myDataSet.Tables("StockMarketOneDayTable").Rows
            '    If myRow("DivDate").ToShortDateString = myRow("Date").AddDays(1).ToShortDateString Then
            '        ArbTicker = myRow("Ticker").Trim()
            '        ArbAsk = myRow("Ask")
            '        ArbNumberOfStocks = 1000 / ArbAsk
            '        CalcDelta(ArbTicker, currentDate)
            '        ArbNumberOfOptions = ArbNumberOfStocks * (1 / (Math.Abs(CalcDelta(ArbTicker, currentDate))))
            '        ' MessageBox.Show("Beep. Boop. Identified " + ArbTicker + " as a DivArb.", "Success?", MessageBoxButtons.OK, MessageBoxIcon.Error)
            '        'ArbUnderway = True
            '    End If
            'Next


            'For Each myRow As DataRow In myDataSet.Tables("StockMarketOneDayTable").Rows
            '    If myRow("DivDate").ToShortDateString = myRow("Date").AddDays(1).ToShortDateString And myRow("Ticker").Trim() = ArbTicker Then
            '        ExecuteAlgoTransaction("SellShort", ArbNumberOfOptions, ArbTicker + "_COCTA")
            '        ExecuteAlgoTransaction("Buy", ArbNumberOfStocks, ArbTicker)

            '        ' MessageBox.Show("Beep. Boop. Bought Shares of " + ArbTicker, "Success?", MessageBoxButtons.OK, MessageBoxIcon.Error)
            '    End If
            'Next


            'For Each myRow As DataRow In myDataSet.Tables("StockMarketOneDayTable").Rows
            '    If myRow("DivDate").ToShortDateString = myRow("Date").ToShortDateString And myRow("Ticker").Trim() = ArbTicker And ArbUnits > 0 Then
            '        ExecuteAlgoTransaction("CashDiv", ArbNumberOfStocks, ArbTicker)
            '        ' MessageBox.Show("Beep. Boop. Cashed Dividend.", "Success?", MessageBoxButtons.OK, MessageBoxIcon.Error)
            '        'ArbYesterday = True
            '    End If
            'Next

            'For Each myRow As DataRow In myDataSet.Tables("StockMarketOneDayTable").Rows
            '    If ArbYesterday = True And myRow("DivDate").ToShortDateString <> myRow("Date").ToShortDateString And myRow("Ticker").Trim() = ArbTicker Then
            '        ExecuteAlgoTransaction("Sell", ArbNumber, ArbTicker)
            '        ExecuteAlgoTransaction("Buy", ArbHedgeVol, ArbTicker + "_COCTE")
            '        ' MessageBox.Show("Beep. Boop. Sold shares.", "Success?", MessageBoxButtons.OK, MessageBoxIcon.Error)
            '        ArbYesterday = False
            '        ArbUnderway = False
            '    End If
            'Next

        End If

        If ArbUnderway = True Then
            'AAPL
            If currentDate.ToShortDateString = "5/12/2016" Or currentDate.ToShortDateString = "8/11/2016" Then
                ExecuteAlgoTransaction("CashDiv", GetCurrPositionInAP("AAPL") + NumberInIP("AAPL"), "AAPL")
            End If
            If currentDate.ToShortDateString = "5/13/2016" Or currentDate.ToShortDateString = "8/12/2016" Then
                ExecuteAlgoTransaction("Buy", ArbNumberOfOptions, "AAPL_COCTB")
                ExecuteAlgoTransaction("Sell", ArbNumberOfStocks, "AAPL")
                ArbUnderway = False
            End If

            'BLK
            If currentDate.ToShortDateString = "6/22/2016" Or currentDate.ToShortDateString = "9/21/2016" Then
                ExecuteAlgoTransaction("CashDiv", GetCurrPositionInAP("BLK") + NumberInIP("BLK"), "BLK")
            End If
            If currentDate.ToShortDateString = "6/23/2016" Or currentDate.ToShortDateString = "9/22/2016" Then
                ExecuteAlgoTransaction("Buy", ArbNumberOfOptions, "BLK_COCTB")
                ExecuteAlgoTransaction("Sell", ArbNumberOfStocks, "BLK")
                ArbUnderway = False
            End If

            'HSY
            If currentDate.ToShortDateString = "6/14/2016" Or currentDate.ToShortDateString = "9/13/2016" Then
                ExecuteAlgoTransaction("CashDiv", GetCurrPositionInAP("HSY") + NumberInIP("HSY"), "HSY")
            End If
            If ArbUnderway = True And currentDate.ToShortDateString = "6/15/2016" Or currentDate.ToShortDateString = "9/14/2016" Then
                ExecuteAlgoTransaction("Buy", ArbNumberOfOptions, "HSY_COCTB")
                ExecuteAlgoTransaction("Sell", ArbNumberOfStocks, "HSY")
                ArbUnderway = False
            End If

            ''NKE: Risky Weekend Arb
            'If ArbUnderway = True And currentDate.ToShortDateString = "7/4/2016" Or currentDate.ToShortDateString = "10/3/2016" Then
            '    ExecuteAlgoTransaction("CashDiv", GetCurrPositionInAP("NKE") + NumberInIP("NKE"), "NKE")
            'End If
            'If ArbUnderway = True And currentDate.ToShortDateString = "7/5/2016" Or currentDate.ToShortDateString = "10/4/2016" Then
            '    ExecuteAlgoTransaction("Buy", ArbNumberOfOptions, "NKE_COCTB")
            '    ExecuteAlgoTransaction("Sell", ArbNumberOfStocks, "NKE")
            '    ArbUnderway = False
            'End If

            ''WMT: Risky Weekend Arb
            'If ArbUnderway = True And currentDate.ToShortDateString = "6/6/2016" Or currentDate.ToShortDateString = "9/5/2016" Then
            '    ExecuteAlgoTransaction("CashDiv", GetCurrPositionInAP("WMT") + NumberInIP("WMT"), "WMT")
            'End If
            'If ArbUnderway = True And currentDate.ToShortDateString = "6/7/2016" Or currentDate.ToShortDateString = "9/6/2016" Then
            '    ExecuteAlgoTransaction("Buy", ArbNumberOfOptions, "WMT_COCTB")
            '    ExecuteAlgoTransaction("Sell", ArbNumberOfStocks, "WMT")
            '    ArbUnderway = False
            'End If

            'XOM
            If currentDate.ToShortDateString = "6/9/2016" Or currentDate.ToShortDateString = "9/8/2016" Then
                ExecuteAlgoTransaction("CashDiv", GetCurrPositionInAP("XOM") + NumberInIP("XOM"), "XOM")
            End If
            If currentDate.ToShortDateString = "6/10/2016" Or currentDate.ToShortDateString = "9/9/2016" Then
                ExecuteAlgoTransaction("Buy", ArbNumberOfOptions, "XOM_COCTB")
                ExecuteAlgoTransaction("Sell", ArbNumberOfStocks, "XOM")
                ArbUnderway = False
            End If

        End If

        Globals.Dashboard.ExecuteOptionTransactionBtn.Enabled = True
        Globals.Dashboard.ExecuteStockTransactionBtn.Enabled = True
    End Sub

    'Public Function FindArbTickers(currentDate)
    '    For Each myRow As DataRow In myDataSet.Tables("StockMarketOneDayTable").Rows
    '        If myRow("DivDate").ToShortDateString = myRow("Date").AddDays(1).ToShortDateString Then
    '            Return myRow("Ticker")
    '        Else
    '            Return Nothing
    '        End If
    '    Next
    '    Return Nothing
    'End Function

    'Public Sub TestArb()
    '    ArbTicker = FindArbTickers(currentDate)
    '    MessageBox.Show("Beep. Boop. Returned" + ArbTicker + ".", "Success?", MessageBoxButtons.OK, MessageBoxIcon.Error)
    'End Sub

    ''Public Sub MoneyMachine(targetDate As Date, currentDate As Date)
    ''    If currentDate = targetDate.AddDays(-1) Then

    ''    End If

    ''End Sub

End Module
