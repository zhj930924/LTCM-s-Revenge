Module DBUtilities
    Public Sub DownloadConfirmationTickets()
        DownloadTableUsingSQL("Select * from " + ConfirmationTicketTableName + " order by symbol", ConfirmationTicketTableName)
    End Sub

    Public Sub ShowConfirmationTickets()
        Globals.Portfolio.Activate()
        Globals.Portfolio.ConfirmationTicketsLO.AutoSetDataBoundColumnHeaders = True
        Globals.Portfolio.ConfirmationTicketsLO.DataSource = myDataSet.Tables(ConfirmationTicketTableName)
    End Sub

    Public Function DownloadCurrentDate2() As Date
        Dim temp As String = ""
        myCommand.CommandText = "Select Value from EnvironmentVariable where Name = 'CurrentDate'"
        Try
            temp = myCommand.ExecuteScalar()
            Globals.Dashboard.CurrentDateCell.Value = Date.Parse(temp).ToLongDateString()
            Return Date.Parse(temp)
        Catch myException As Exception
            Return currentDate
        End Try
    End Function

    Public Function StockPricesExist() As Boolean
        Dim temp As String = ""
        myCommand.CommandText = "Select top 1 Bid from Stockmarket"
        Try
            temp = myCommand.ExecuteScalar()
            If temp = Nothing Or temp = "" Then
                MessageBox.Show("Beep. Boop. No stock prices yet.", "Empty DB", MessageBoxButtons.OK)
                Return False
            End If
        Catch myException As Exception
            MessageBox.Show("Beep. Boop. No stock prices yet." + "Maybe this will help: " + myException.Message, "Failed to execute a query", MessageBoxButtons.OK)
            Return False
        End Try
        Return True
    End Function

    Public Sub ClearTeamPortfolioOnDb()
        ExecuteNonQuery("Delete from " + portfolioTableName)
    End Sub

    Public Sub UploadPosition(sym As String, newValue As Double)
        Try
            newValue = Math.Round(newValue, 2)
            sym = sym.Trim()
            myCommand.CommandText = "Delete from " + portfolioTableName + " Where Symbol = '" + sym + "';"
            myCommand.ExecuteNonQuery()

            If (newValue <> 0) Or (sym = "CAccount") Then
                myCommand.CommandText = String.Format("Insert into {0} Values ('{1}', '{2}')", portfolioTableName, sym, newValue)
                myCommand.ExecuteNonQuery()
            End If
        Catch myException As Exception
            MessageBox.Show("Beep. Boop. I could not set " + sym + ". Maybe this will help: " + myException.Message)

        End Try
    End Sub

    Public Sub DownloadStockDataToChart(ticker As String)
        Dim SQL As String
        SQL = "Select date, bid, ask from StockMarket where ticker = '" + ticker + "';"
        DownloadTableUsingSQL(SQL, "StockDataToChart")
    End Sub

    Public Sub DownloadOptionDataToChart(symbol As String)
        Dim SQL As String
        SQL = "Select date, bid, ask from OptionMarket where symbol = '" + symbol + "';"
        DownloadTableUsingSQL(SQL, "OptionDataToChart")
    End Sub

    Public Function DownloadDividend(ticker As String, targetDate As Date)
        Dim temp As String = "0"
        Dim mySql As String = ""

        If (targetDate.DayOfWeek = DayOfWeek.Saturday) Then
            targetDate = targetDate.AddDays(-1)
        End If
        If (targetDate.DayOfWeek = DayOfWeek.Sunday) Then
            targetDate = targetDate.AddDays(-2)
        End If

        If IsAStock(ticker) Then
            mySql = "Select dividend from StockMarket where ticker = '" + ticker + "' and date = '" + targetDate.ToShortDateString() + "'"
        End If

        Try
            myCommand.CommandText = mySql
            temp = myCommand.ExecuteScalar()
        Catch ex As Exception
            MessageBox.Show("Beep. Boop. I could not retrieve dividend for " + ticker + ". This is the query you created: " + mySql + " and this is what the DB said" + ex.Message, "Ouch!", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        Return Double.Parse(temp)

    End Function

    Public Sub ExecuteNonQuery(SQLString As String)
        Try
            myCommand.CommandText = SQLString
            myCommand.ExecuteNonQuery()
        Catch myException As Exception
            MessageBox.Show("Beep. Boop. " + myException.Message, "Likely SQL Problem", MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try
    End Sub

    Public Sub DownloadTransactionQueue(team As String)
        DownloadTableUsingSQL("Select * from TransactionQueue where teamID = '" + teamID + "' order by RowID desc", "TransactionQueueTable")
    End Sub

    Public Sub ShowTransactionQueue()
        Globals.TransactionQueue.Activate()
        Globals.TransactionQueue.TransactionQueueLO.AutoSetDataBoundColumnHeaders = True
        Globals.TransactionQueue.TransactionQueueLO.DataSource = myDataSet.Tables("TransactionQueueTable")
    End Sub

    Public Function DownloadLastTransactionDate(targetDate As Date) As Date
        Dim temp As String = ""
        myCommand.CommandText = String.Format("Select max(date) from TransactionQueue where teamid = {0} and date <= '{1}'", teamID, targetDate.ToShortDateString())
        Try
            temp = myCommand.ExecuteScalar()
            Return Date.Parse(temp)
        Catch myException As Exception
            MessageBox.Show("Last transaction not found. Set LastTransactionDate to StartDate ", "Transaction Queue", MessageBoxButtons.OK)
            Return startDate
        End Try
    End Function

    Public Function DownloadCAccount() As Double
        Dim temp As String = ""
        myCommand.CommandText = "Select Units from " + portfolioTableName + " where Symbol = 'CAccount'"
        Try
            temp = myCommand.ExecuteScalar()
            Return Double.Parse(temp)
        Catch myException As Exception
            MessageBox.Show("Beep. Boop. Couldn't retrieve the CAccount. I reported $0. " + "Maybe this will help: " + myException.Message, "Likely SQL problem", MessageBoxButtons.OK)
            DownloadEnvironmentVariable()
            initialCAccount = GetInitialCAccount()
            UploadPosition("CAccount", initialCAccount)
            Return initialCAccount
        End Try
    End Function

    Public Function DownloadCurrentDate() As Date
        Dim temp As String = ""
        myCommand.CommandText = "Select Value from EnvironmentVariable where Name = 'CurrentDate'"
        Try
            temp = myCommand.ExecuteScalar()
            Globals.Dashboard.CurrentDateCell.Value = Date.Parse(temp).ToLongDateString()
            Return Date.Parse(temp)
        Catch myException As Exception
            Return currentDate
        End Try
    End Function

    Public Sub DownloadTickers()
        DownloadTableUsingSQL("Select distinct ticker from StockMarket order by ticker", "TickerTable")
    End Sub

    Public Sub DownloadSymbols()
        DownloadTableUsingSQL("Select distinct symbol from OptionMarket order by symbol", "SymbolTable")
    End Sub

    Public Function DownloadAsk(symbol As String, targetDate As Date)
        Dim temp As String = "0"
        Dim mySql As String = ""

        If (targetDate.DayOfWeek = DayOfWeek.Saturday) Then
            targetDate = targetDate.AddDays(-1)
        End If
        If (targetDate.DayOfWeek = DayOfWeek.Sunday) Then
            targetDate = targetDate.AddDays(-2)
        End If
        If IsAStock(symbol) Then
            mySql = "Select Ask from StockMarket where ticker = '" + symbol + "' and date = '" + targetDate.ToShortDateString() + "'"
        Else
            mySql = "Select Ask from OptionMarket where symbol = '" + symbol + "' and date = '" + targetDate.ToShortDateString() + "'"
        End If

        Try
            myCommand.CommandText = mySql
            temp = myCommand.ExecuteScalar()
        Catch ex As Exception
            MessageBox.Show("Beep. Boop. I could not retrieve the ask for " + symbol + ". This is the query you created " + mySql + " and this is what the DB said " + ex.Message, "Beep!", MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try
        Return Double.Parse(temp)
    End Function

    Public Function DownloadBid(symbol As String, targetDate As Date)
        Dim temp As String = "0"
        Dim mySql As String = ""

        If (targetDate.DayOfWeek = DayOfWeek.Saturday) Then
            targetDate = targetDate.AddDays(-1)
        End If
        If (targetDate.DayOfWeek = DayOfWeek.Sunday) Then
            targetDate = targetDate.AddDays(-2)
        End If
        If IsAStock(symbol) Then
            mySql = "Select Bid from StockMarket where ticker = '" + symbol + "' and date = '" + targetDate.ToShortDateString() + "'"
        Else
            mySql = "Select Bid from OptionMarket where symbol = '" + symbol + "' and date = '" + targetDate.ToShortDateString() + "'"
        End If

        Try
            myCommand.CommandText = mySql
            temp = myCommand.ExecuteScalar()
        Catch ex As Exception
            MessageBox.Show("Beep. Boop. I could not retrieve the bid for " + symbol + ". This is the query you created " + mySql + " and this is what the DB said " + ex.Message, "Beep!", MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try
        Return Double.Parse(temp)
    End Function

    Public Sub DownloadPricesForOneDay(targetDate As Date)
        If (targetDate.DayOfWeek = DayOfWeek.Saturday) Then
            targetDate = targetDate.AddDays(-1)
        End If
        If (targetDate.DayOfWeek = DayOfWeek.Sunday) Then
            targetDate = targetDate.AddDays(-2)
        End If

        If targetDate.Date <> lastPriceDownloadDate.Date Then
            Dim mySQL As String
            mySQL = "Select * from StockMarket where Date = '" + targetDate.ToShortDateString() + "';"
            DownloadTableUsingSQL(mySQL, "StockMarketOneDayTable")

            mySQL = "Select * from OptionMarket where Date = '" + targetDate.ToShortDateString() + "';"
            DownloadTableUsingSQL(mySQL, "OptionMarketOneDayTable")
        End If
    End Sub

    Dim myConnection As SqlClient.SqlConnection
    Dim myCommand As SqlClient.SqlCommand
    Dim myDataAdapter As SqlClient.SqlDataAdapter
    Public myDataSet As DataSet
    Dim myDataTable As DataTable
    Dim mySQLString As String

    Public Sub CreateAndConnectTheADOObjects()
        myConnection = New SqlClient.SqlConnection
        myCommand = New SqlClient.SqlCommand
        myCommand.Connection = myConnection
        myDataAdapter = New SqlClient.SqlDataAdapter
        myDataAdapter.SelectCommand = myCommand
        myDataSet = New DataSet
    End Sub

    Public Function OpenDBConnection() As Boolean
        Select Case activeDB
            Case "Alpha"
                myConnection.ConnectionString = "Data Source=f-sg6m-s4.comm.virginia.edu;" + "Initial Catalog=HedgeTournamentALPHA;Integrated Security=True"
            Case "Beta"
                myConnection.ConnectionString = "Data Source=f-sg6m-s4.comm.virginia.edu;" + "Initial Catalog=HedgeTournamentBETA;Integrated Security=True"
            Case "Gamma"
                myConnection.ConnectionString = "Data Source=f-sg6m-s4.comm.virginia.edu;" + "Initial Catalog=HedgeTournamentGAMMA;Integrated Security=True"
        End Select
        Try
            myConnection.Open()
            Return True
        Catch myException As Exception
            MessageBox.Show("Hello. It's me. The DB is not responding. " + myException.Message, "Connection problem", MessageBoxButtons.OK)
            Return False
        End Try
    End Function

    Public Sub DownloadInitialPositions()
        If Not excludeIP Then
            DownloadTableUsingSQL("Select * from InitialPosition order by symbol", "InitialPositionTable")
        End If
    End Sub

    Public Sub ShowInitialPositions()
        Globals.Portfolio.Activate()
        Globals.Portfolio.InitialPositionsLO.AutoSetDataBoundColumnHeaders = True
        Globals.Portfolio.InitialPositionsLO.DataSource = myDataSet.Tables("InitialPositionTable")
    End Sub

    Public Sub DownloadTableUsingSQL(mySQL As String, NameOfTheResultTable As String)
        ClearDataSetTable(NameOfTheResultTable)
        myCommand.CommandText = mySQL
        Try
            myDataAdapter.Fill(myDataSet, NameOfTheResultTable)
        Catch myException As Exception
            MessageBox.Show("I must have called 1000 times, but I could not download " + NameOfTheResultTable + " using " + mySQL + ". " + "No corrective action was taken. Maybe this will help: " + myException.Message, "Likely SQL problem.", MessageBoxButtons.OK)
        End Try
    End Sub

    Public Sub ClearDataSetTable(TableToClear As String)
        If myDataSet.Tables.Contains(TableToClear) Then
            myDataSet.Tables(TableToClear).Clear()
        End If
    End Sub

    Public Sub DownloadAcquiredPositions()
        DownloadTableUsingSQL("Select * from " + portfolioTableName + " order by symbol", portfolioTableName)
    End Sub

    Public Sub ShowAcquiredPositions()
        Globals.Portfolio.Activate()
        Globals.Portfolio.AcquiredPositionsLO.AutoSetDataBoundColumnHeaders = True
        Globals.Portfolio.AcquiredPositionsLO.DataSource = myDataSet.Tables(portfolioTableName)
    End Sub

    Public Sub DownloadStockMarket()
        DownloadTableUsingSQL("Select * from StockMarket order by date desc", "StockMarketTable")
    End Sub

    Public Sub ShowStockMarket()
        Globals.Markets.Activate()
        Globals.Markets.StockMarketLO.AutoSetDataBoundColumnHeaders = True
        Globals.Markets.StockMarketLO.DataSource = myDataSet.Tables("StockMarketTable")
    End Sub

    Public Sub DownloadOptionMarket()
        DownloadTableUsingSQL("Select * from OptionMarket order by date desc", "OptionMarketTable")
    End Sub

    Public Sub ShowOptionMarket()
        Globals.Markets.Activate()
        Globals.Markets.OptionMarketLO.AutoSetDataBoundColumnHeaders = True
        Globals.Markets.OptionMarketLO.DataSource = myDataSet.Tables("OptionMarketTable")
    End Sub

    Public Sub DownloadStockIndex()
        DownloadTableUsingSQL("Select * from StockIndex order by date desc", "StockIndexTable")
    End Sub

    Public Sub ShowStockIndex()
        Globals.Markets.Activate()
        Globals.Markets.StockIndexLO.AutoSetDataBoundColumnHeaders = True
        Globals.Markets.StockIndexLO.DataSource = myDataSet.Tables("StockIndexTable")
    End Sub

    Public Sub DownloadEnvironmentVariable()
        DownloadTableUsingSQL("Select * from EnvironmentVariable", "EnvironmentVariableTable")
    End Sub
    Public Sub ShowEnvironmentVariable()
        Globals.Environment.Activate()
        Globals.Environment.SettingsLO.AutoSetDataBoundColumnHeaders = True
        Globals.Environment.SettingsLO.DataSource = myDataSet.Tables("EnvironmentVariableTable")
    End Sub

    Public Sub DownloadTransactionCost()
        DownloadTableUsingSQL("Select * from TransactionCost", "TransactionCostTable")
    End Sub

    Public Sub ShowTransactionCost()
        Globals.Environment.Activate()
        Globals.Environment.TransactionCostsLO.AutoSetDataBoundColumnHeaders = True
        Globals.Environment.TransactionCostsLO.DataSource = myDataSet.Tables("TransactionCostTable")
    End Sub

End Module
