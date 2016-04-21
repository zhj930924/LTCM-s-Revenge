Module ScheduledTransactions
    Public ArbTicker As String = ""
    Public SetUpArb As Boolean = False
    Public ArbUnderway As Boolean = False
    Public ArbAsk As Double = 0
    Public ArbNumber As Double = 0
    Public ArbHedgeVol As Double = 0



    'ExecuteAlgoTransaction moved to Smart Hedger and modified
    Public Sub DoScheduledTransactions()
        Globals.Dashboard.ExecuteOptionTransactionBtn.Enabled = False
        Globals.Dashboard.ExecuteStockTransactionBtn.Enabled = False

        For Each myRow As DataRow In myDataSet.Tables("StockMarketOneDayTable").Rows
            If myRow("DivDate").ToShortDateString = myRow("Date").AddDays(1).ToShortDateString And ArbUnderway = False Then
                ArbTicker = myRow("Ticker").Trim()
                ArbAsk = myRow("Ask")
                ArbNumber = 4000000 / ArbAsk
                CalcDelta(ArbTicker, currentDate)
                ArbHedgeVol = ArbNumber * (1 / (Math.Abs(CalcDelta(ArbTicker, currentDate))))
                ' MessageBox.Show("Beep. Boop. Identified " + ArbTicker + " as a DivArb.", "Success?", MessageBoxButtons.OK, MessageBoxIcon.Error)
                ArbUnderway = True
            End If
        Next

        For Each myRow As DataRow In myDataSet.Tables("StockMarketOneDayTable").Rows
            If myRow("DivDate").ToShortDateString = myRow("Date").AddDays(1).ToShortDateString And myRow("Ticker").Trim() = ArbTicker Then

                ExecuteAlgoTransaction("Buy", ArbNumber, ArbTicker)
                ExecuteAlgoTransaction("SellShort", ArbHedgeVol, ArbTicker + "_COCTE")
                ' MessageBox.Show("Beep. Boop. Bought Shares of " + ArbTicker, "Success?", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        Next

        For Each myRow As DataRow In myDataSet.Tables("StockMarketOneDayTable").Rows
            If myRow("DivDate").ToShortDateString = myRow("Date").ToShortDateString And myRow("Ticker").Trim() = ArbTicker Then
                ExecuteAlgoTransaction("CashDiv", ArbNumber, ArbTicker)
                ' MessageBox.Show("Beep. Boop. Cashed Dividend.", "Success?", MessageBoxButtons.OK, MessageBoxIcon.Error)
                ArbYesterday = True
            End If
        Next

        For Each myRow As DataRow In myDataSet.Tables("StockMarketOneDayTable").Rows
            If ArbYesterday = True And myRow("DivDate").ToShortDateString <> myRow("Date").ToShortDateString And myRow("Ticker").Trim() = ArbTicker Then
                ExecuteAlgoTransaction("Sell", ArbNumber, ArbTicker)
                ExecuteAlgoTransaction("Buy", ArbHedgeVol, ArbTicker + "_COCTE")
                ' MessageBox.Show("Beep. Boop. Sold shares.", "Success?", MessageBoxButtons.OK, MessageBoxIcon.Error)
                ArbYesterday = False
                ArbUnderway = False
            End If
        Next




        Globals.Dashboard.ExecuteOptionTransactionBtn.Enabled = True
        Globals.Dashboard.ExecuteStockTransactionBtn.Enabled = True
    End Sub

    Public Function FindArbTickers(currentDate)
        For Each myRow As DataRow In myDataSet.Tables("StockMarketOneDayTable").Rows
            If myRow("DivDate").ToShortDateString = myRow("Date").AddDays(1).ToShortDateString Then
                Return myRow("Ticker")
            Else
                Return Nothing
            End If
        Next
        Return Nothing
    End Function

    Public Sub TestArb()
        ArbTicker = FindArbTickers(currentDate)
        MessageBox.Show("Beep. Boop. Returned" + ArbTicker + ".", "Success?", MessageBoxButtons.OK, MessageBoxIcon.Error)
    End Sub

    'Public Sub MoneyMachine(targetDate As Date, currentDate As Date)
    '    If currentDate = targetDate.AddDays(-1) Then

    '    End If

    'End Sub

End Module
