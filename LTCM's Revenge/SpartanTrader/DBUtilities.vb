Module DBUtilities

    'This Module contains procedures for managing DB connections and manipulating data
    Public Function StockPricesExist() As Boolean
        Dim temp As String = ""
        myCommand.CommandText = "Select top 1 Bid from Stockmarket"
        Try
            temp = myCommand.ExecuteScalar()
            If temp = Nothing Or temp = "" Then
                MessageBox.Show("No stock prices yet, Dave.", "Empty DB", MessageBoxButtons.OK)
                Return False
            End If
        Catch ex As Exception
            MessageBox.Show("No stock prices yet, Dave." + "Maybe this will help: " + ex.Message, "Failed to execute a query", MessageBoxButtons.OK)
            Return False
        End Try
        Return True
    End Function


    'HM19
    Public Sub ClearTeamPortfolioOnDB()
        ExecuteNonQuery("Delete from " + portfolioTableName)
    End Sub

    Public Sub UploadPosition(sym As String, newValue As Double)
        Try
            newValue = Math.Round(newValue, 2)
            sym = sym.Trim()
            myCommand.CommandText = "Delete from " + portfolioTableName + " where Symbol = '" + sym + "';"
            myCommand.ExecuteNonQuery()

            ' if new position value is 0, we skip it. exception: CAccount
            If (newValue <> 0) Or (sym = "CAccount") Then
                myCommand.CommandText = String.Format("Insert into {0} Values ('{1}','{2}')", portfolioTableName, sym, newValue)
                myCommand.ExecuteNonQuery()
            End If
        Catch ex As Exception
            MessageBox.Show("I could not set " + sym + ", Dave. " + "Maybe this will help: " + ex.Message)
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

    'HM 17
    Public Function DownloadDividend(ticker As String, targetDate As Date)
        Dim temp As String = "0"
        Dim mySQL As String = ""

        'last day in which the markets are open is friday
        If (targetDate.DayOfWeek = DayOfWeek.Saturday) Then
            targetDate = targetDate.AddDays(-1)
        End If
        If (targetDate.DayOfWeek = DayOfWeek.Sunday) Then
            targetDate = targetDate.AddDays(-2)
        End If

        If IsAStock(ticker) Then
            mySQL = "Select dividend from StockMarket where ticker = '" _
                + ticker + "' and date = '" + targetDate.ToShortDateString() + "'"
        End If

        Try
            myCommand.CommandText = mySQL
            temp = myCommand.ExecuteScalar()
        Catch ex As Exception
            MessageBox.Show("Holy batmobile! I could not retrieve the dividend for " + ticker + ". This is the query you created " _
                            + mySQL + " and this is what the DB said " + ex.Message, "Ouch!", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        Return Double.Parse(temp)
    End Function

    Public Sub ExecuteNonQuery(SQLString As String)
        Try
            myCommand.CommandText = SQLString
            myCommand.ExecuteNonQuery()
        Catch ex As Exception
            MessageBox.Show("You query failed, Dave." + "Maybe this will help: " + ex.Message, "Likely SQL problem", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Public Sub DownloadTransactionQueue(team As String)
        DownloadTableUsingSQL("Select * from TransactionQueue where teamID = '" + team + "' order by rowID desc", "TransactionQueueTable")
            End Sub


    Public Sub ShowTransactionQueue()
        Globals.TransactionQueue.Activate()
        Globals.TransactionQueue.TransactionQueueLO.AutoSetDataBoundColumnHeaders = True
        Globals.TransactionQueue.TransactionQueueLO.DataSource = myDataSet.Tables("TransactionQueueTable")
    End Sub

    'HM16
    Public Function DownloadLastTransactionDate(targetDate As Date) As Date
        Dim temp As String = ""
        myCommand.CommandText = String.Format("Select max(date) from TransactionQueue where teamid = {0} and date <= '{1}'", teamID,
                                              targetDate.ToShortDateString())
        Try
            temp = myCommand.ExecuteScalar()
            Return Date.Parse(temp)
        Catch ex As Exception
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
            MessageBox.Show("Holy Batmobile! I could not retrive the CAccount. So I reset it to the initial value",
                            "No CAccount", MessageBoxButtons.OK, MessageBoxIcon.Information)
            DownloadEnvironmentVariable()
            initialCAccount = GetInitialCAccount()
            UploadPosition("CAccount", initialCAccount)
            Return initialCAccount
        End Try
    End Function


    'HM14
    Public Sub DownloadTickers()
        DownloadTableUsingSQL("Select distinct ticker from StockMarket order by ticker", "TickerTable")
    End Sub

    Public Sub DownloadSymbols()
        DownloadTableUsingSQL("Select distinct symbol from OptionMarket order by symbol", "SymbolTable")
    End Sub

    Public Function DownloadAsk(symbol As String, targetDate As Date)
        Dim temp As String = "0"
        Dim mySql As String = ""

        'last day in which the markets are open is friday
        If (targetDate.DayOfWeek = DayOfWeek.Saturday) Then
            targetDate = targetDate.AddDays(-1)
        End If

        If (targetDate.DayOfWeek = DayOfWeek.Sunday) Then
            targetDate = targetDate.AddDays(-2)
        End If

        If isAStock(symbol) Then
            mySql = "Select Ask from StockMarket where ticker = '" + symbol + "' And date = '" + targetDate.ToShortDateString() + "'"
        Else
            mySql = "Select Ask from OptionMarket where symbol = '" + symbol + "' And date = '" + targetDate.ToShortDateString() + "'"
        End If

        Try
            myCommand.CommandText = mySql
            temp = myCommand.ExecuteScalar()
        Catch ex As Exception
            MessageBox.Show("Holy Batmobile! I could not retrive the ask for " + symbol + ". This is the query you created " +
                            mySql + " and this is what the DB said " + ex.Message, "Ouch!", MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try
        Return Double.Parse(temp)
    End Function

    Public Function DownloadBid(symbol As String, targetDate As Date)
        Dim temp As String = "0"
        Dim mySql As String = ""

        'last day in which the markets are open is friday
        If (targetDate.DayOfWeek = DayOfWeek.Saturday) Then
            targetDate = targetDate.AddDays(-1)
        End If

        If (targetDate.DayOfWeek = DayOfWeek.Sunday) Then
            targetDate = targetDate.AddDays(-2)
        End If

        If isAStock(symbol) Then
            mySql = "Select Bid from StockMarket where ticker = '" + symbol + "' And date = '" + targetDate.ToShortDateString() + "'"
        Else
            mySql = "Select Bid from OptionMarket where symbol = '" + symbol + "' And date = '" + targetDate.ToShortDateString() + "'"
        End If

        Try
            myCommand.CommandText = mySql
            temp = myCommand.ExecuteScalar()
        Catch ex As Exception
            MessageBox.Show("Holy Batmobile! I could not retrive the bid for " + symbol + ". This is the query you created " +
                            mySql + " and this is what the DB said " + ex.Message, "Ouch!", MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try
        Return Double.Parse(temp)
    End Function

    Public Function DownloadCurrentDate() As Date
        Dim temp As String = ""
        myCommand.CommandText = "Select Value from EnvironmentVariable where name = 'CurrentDate'"
        Try
            temp = myCommand.ExecuteScalar()
            Globals.Dashboard.CurrentDateCell.Value = Date.Parse(temp).ToLongDateString()
            Return Date.Parse(temp)
        Catch ex As Exception
            Return currentDate
        End Try
    End Function

    Public Function DownloadCurrentDate2() As Date
        Dim temp As String = ""
        myCommand.CommandText = "Select Value from EnvironmentVariable where name = 'CurrentDate'"
        Try
            temp = myCommand.ExecuteScalar()
            Globals.Dashboard.CurrentDateCell.Value = Date.Parse(temp).ToLongDateString()
            Return Date.Parse(temp)
        Catch ex As Exception
            Return currentDate
        End Try
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
            mySQL = "Select * from StockMarket where date = '" + targetDate.ToShortDateString() + "';"
            DownloadTableUsingSQL(mySQL, "StockMarketOneDayTable")

            mySQL = "Select * from OptionMarket where date = '" + targetDate.ToShortDateString() + "';"
            DownloadTableUsingSQL(mySQL, "OptionMarketOneDayTable")
            lastPriceDownloadDate = targetDate
        End If

    End Sub

    'HM13
    'ADO-related objects (also global variables)
    Dim myConnection As SqlClient.SqlConnection
    Dim myCommand As SqlClient.SqlCommand
    Dim myDataAdaptor As SqlClient.SqlDataAdapter
    Public myDataSet As DataSet
    Dim myDataTable As DataTable
    Dim mySQLString As String

    Public Sub CreateAndConnectTheADOObjects()
        'Create the connection and set the connection string
        myConnection = New SqlClient.SqlConnection
        'Create the command and set the connection
        myCommand = New SqlClient.SqlCommand
        myCommand.Connection = myConnection
        'Create the data adaptor and set the selectCommand
        myDataAdaptor = New SqlClient.SqlDataAdapter
        myDataAdaptor.SelectCommand = myCommand
        'Create the dataset
        myDataSet = New DataSet
    End Sub

    Public Function OpenDBConnection() As Boolean
        Select Case activeDB
            Case "Alpha"
                myConnection.ConnectionString = "Data Source=f-sg6m-s4.comm.virginia.edu;" +
                    "Initial Catalog=HedgeTournamentAlpha;Integrated Security=True"
            Case "Beta"
                myConnection.ConnectionString = "Data Source=f-sg6m-s4.comm.virginia.edu;" +
                    "Initial Catalog=HedgeTournamentBeta;Integrated Security=True"
            Case "Gamma"
                myConnection.ConnectionString = "Data Source=f-sg6m-s4.comm.virginia.edu;" +
                    "Initial Catalog=HedgeTournamentGamma;Integrated Security=True"
        End Select
        Try
            myConnection.Open()
            Return True 'true = success
        Catch myException As Exception
            MessageBox.Show("I am calling, but the DB is not responding, Dave. " + myException.Message, "Connection problem", MessageBoxButtons.OK)
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
            myDataAdaptor.Fill(myDataSet, NameOfTheResultTable)
        Catch myException As Exception
            MessageBox.Show("Dave, I could not download " + NameOfTheResultTable + " using " + mySQL + ". " +
                "No corrective action was taken. Maybe this will help:" + myException.Message, "Likely SQL problem.", MessageBoxButtons.OK)
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

    Public Sub DownloadConfirmationTickets()
        DownloadTableUsingSQL("Select * from " + ConfirmationTicketTableName + " order by symbol", ConfirmationTicketTableName)
    End Sub

    Public Sub ShowConfrimationTickets()
        Globals.Portfolio.Activate()
        Globals.Portfolio.ConfirmationTicketsLO.AutoSetDataBoundColumnHeaders = True
        Globals.Portfolio.ConfirmationTicketsLO.DataSource = myDataSet.Tables(ConfirmationTicketTableName)
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
        Globals.Markets.SP500LO.AutoSetDataBoundColumnHeaders = True
        Globals.Markets.SP500LO.DataSource = myDataSet.Tables("StockIndexTable")
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
