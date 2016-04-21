Module ScheduledTransactions

    Public Sub DoScheduledTransactions()
        Globals.Dashboard.ExecuteOptionTransactionButton.Enabled = False
        Globals.Dashboard.ExecuteStockTransactionButton.Enabled = False

        'If currentDate.ToShortDateString = "5/13/2015" Then
        'ExecuteAlgoTransaction("Buy", 100000, "AAPL")
        'End If

        'If currentDate.ToShortDateString = "5/14/2015" Then
        'ExecuteAlgoTransaction("CashDiv", 100000, "AAPL")
        'End If

        'If currentDate.ToShortDateString = "5/15/2015" Then
        'ExecuteAlgoTransaction("Sell", 100000, "AAPL")
        'End If

        'If currentDate.ToShortDateString = "6/25/2015" Then
        'ExecuteAlgoTransaction("Buy", 100000, "BAC")
        'End If

        'If currentDate.ToShortDateString = "6/26/2015" Then
        'ExecuteAlgoTransaction("CashDiv", 100000, "BAC")
        'End If

        'If currentDate.ToShortDateString = "6/29/2015" Then
        '    ExecuteAlgoTransaction("Sell", 100000, "BAC")
        'End If

        If currentDate.ToShortDateString = "6/9/2016" Then
            ExecuteAlgoTransaction("Buy", 5000, "XOM")
        End If

        If currentDate.ToShortDateString = "6/9/2016" Then
            ExecuteAlgoTransaction("SellShort", 10000, "XOM_POCTD")
        End If

        If currentDate.ToShortDateString = "6/10/2016" Then
            ExecuteAlgoTransaction("CashDiv", 5000, "XOM")
        End If

        If currentDate.ToShortDateString = "6/11/2016" Then
            ExecuteAlgoTransaction("Buy", 10000, "XOM_POCTD")
        End If

        If currentDate.ToShortDateString = "6/11/2016" Then
            ExecuteAlgoTransaction("Sell", 5000, "XOM")
        End If

        Globals.Dashboard.ExecuteOptionTransactionButton.Enabled = True
        Globals.Dashboard.ExecuteStockTransactionButton.Enabled = True

    End Sub

End Module
