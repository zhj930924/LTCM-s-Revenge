Module ProcessAutomation
    Public WithEvents spyTimer As Timer
    Public WithEvents secondsTimer As Timer

    Public Sub WarmStart()
        StopTimers()
        ClearOldData()
        CreateAndConnectTheADOObjects()
        If OpenDBConnection() = False Then
            Exit Sub
        End If
        If StockPricesExist() = False Then
            Exit Sub
        End If
        DownloadStaticData()
        DownloadTeamData()
        Select Case traderMode
            Case "Manual"
                currentDate = DownloadCurrentDate()
                DownloadPricesForOneDay(currentDate)
                SetupCharts()
                DailyRoutine()
            Case "Synch"
                currentDate = DownloadCurrentDate()
                DownloadPricesForOneDay(currentDate)
                SetupCharts()
                StartTimers()
            Case "Simulation"
                Globals.Dashboard.Activate()
                currentDate = startDate
                lastTransactionDate = startDate
                DownloadPricesForOneDay(currentDate)
                SetupCharts()
                While currentDate < DownloadCurrentDate()
                    DailyRoutine()
                    currentDate = currentDate.AddDays(1)
                    For i As Integer = 0 To 2
                        Application.DoEvents()
                    Next
                End While
            Case "Auto"
                currentDate = DownloadCurrentDate()
                DownloadPricesForOneDay(currentDate)
                SetupCharts()
                StartTimers()
        End Select
    End Sub

    Public Sub DailyRoutine()
        myTransaction.Clear()
        Globals.Dashboard.CurrentDateCell.Value = currentDate.ToLongDateString()
        DownloadPricesForOneDay(currentDate)
        CalcFinancialMetrics(currentDate)
        DisplayFinancialMetrics(currentDate)
        ' Globals.Dashboard.DisplayFamilyDeltas()
        Select Case traderMode
            Case "Manual"
                'RecommendHedge()
            Case "Synch"
                DoScheduledTransactions()
                CalcFinancialMetrics(currentDate)
                DisplayFinancialMetrics(currentDate)
                'RecommendHedge()
            Case "Simulation"
                DoScheduledTransactions()
                CalcFinancialMetrics(currentDate)
                DisplayFinancialMetrics(currentDate)
                'AlgoHedgeAll()
            Case "Auto"
                DoScheduledTransactions()
                CalcFinancialMetrics(currentDate)
                DisplayFinancialMetrics(currentDate)
                'AlgoHedgeAll()
        End Select
    End Sub
    Public Sub StopTimers()
        If IsNothing(spyTimer) Then
            'skip
        Else
            spyTimer.Stop()
            secondsTimer.Stop()
            Globals.Dashboard.SetSeconds(0)
        End If

    End Sub

    Public Sub StartTimers()
        spyTimer = New Timer
        spyTimer.Interval = 2000
        secondsTimer = New Timer
        secondsTimer.Interval = 1000
        Globals.Dashboard.SetSeconds(0)
        secondsTimer.Start()
        spyTimer.Start()
    End Sub

    Private Sub secondsTimer_Tick() Handles spyTimer.Tick
        Globals.Dashboard.SetSeconds(secondsLeft - 1)
    End Sub

    Private Sub spyTimer_Tick() Handles spyTimer.Tick
        Dim tempNewDate As Date
        tempNewDate = DownloadCurrentDate2()
        If tempNewDate.Date <> currentDate.Date Then
            currentDate = tempNewDate
            Globals.Dashboard.SetSeconds(60)
            DailyRoutine()
        End If
    End Sub

    Public Sub CalcFinancialMetrics(targetDate As Date)

        margin = CalcMargin(targetDate)
        IP = CalcIPValue(targetDate)
        AP = CalcAPValue(targetDate)
        TPV = IP + AP + CAccount + CalcInterestSLT(targetDate)
        TaTPV = CalcTaTPV(targetDate)
        TE = TPV - TaTPV
        If TE > 0 Then TE = TE / 4
        TEpercent = TE / TaTPV
        TPVNoHedge = CalcTPVNoHedge(targetDate)
        Globals.Dashboard.UpdateTPVTrackingTable(targetDate, TPV, TaTPV, TPVNoHedge)
        ' If traderMode <> "Simulation" Then
        'RecommendHedges()
        ' End If

    End Sub

    Public Sub DisplayFinancialMetrics(targetDate As Date)
        Try
            Globals.Dashboard.CAccountCell.Value = CAccount
            Globals.Dashboard.MarginCell.Value = margin
            Globals.Dashboard.MarginPercCell.Value = margin * 0.3
            Globals.Dashboard.maxMarginCell.Value = maxMargins

            Globals.Dashboard.InterestSLTCell.Value = interestSLT
            Globals.Dashboard.IPCell.Value = IP
            Globals.Dashboard.APCell.Value = AP

            Globals.Dashboard.TPVatStartCell.Value = TPVatStart
            Globals.Dashboard.TPVCell.Value = TPV
            Globals.Dashboard.TaTPVCell.Value = TaTPV
            Globals.Dashboard.TECell.Value = TE
            Globals.Dashboard.TEPercCell.Value = TEpercent

        Catch ex As Exception

        End Try
    End Sub


    Public Sub DownloadTeamData()
        CAccount = DownloadCAccount()
        DownloadAcquiredPositions()
        lastTransactionDate = DownloadLastTransactionDate(endDate)
    End Sub

    Public Sub SetupCharts()
        Globals.Dashboard.FillTPVTrackingTable()
        Globals.Dashboard.SetupTrackingChart()
        Globals.FinCharts.SetUpFinCharts()
    End Sub

    Public Sub ClearOldData()
        Globals.Markets.StockMarketLO.DataSource = Nothing
        Globals.Markets.OptionMarketLO.DataSource = Nothing
        Globals.Markets.StockIndexLO.DataSource = Nothing
        Globals.Portfolio.InitialPositionsLO.DataSource = Nothing
        Globals.Portfolio.AcquiredPositionsLO.DataSource = Nothing
        Globals.Environment.SettingsLO.DataSource = Nothing
        Globals.Environment.TransactionCostsLO.DataSource = Nothing
        Globals.TransactionQueue.TransactionQueueLO.DataSource = Nothing
        Globals.Portfolio.ConfirmationTicketsLO.DataSource = Nothing
        Globals.Dashboard.AlgoTransactionsLogTBox.Text = "Algo Transaction Log - Ready."
        Globals.Dashboard.Activate()
        lastPriceDownloadDate = "1/1/2000"
        'ResetRecommendations()
        'DisplayRecommendations()
    End Sub

    Public Sub DownloadStaticData()
        DownloadInitialPositions()
        DownloadTransactionCost()
        DownloadEnvironmentVariable()
        DownloadTickers()
        DownloadSymbols()

        Globals.Dashboard.LoadCBoxes()

        initialCAccount = GetInitialCAccount()
        iRate = GetIRate()
        startDate = GetStartDate()
        endDate = GetEndDate()
        maxMargins = GetMaxMargins()
        TPVatStart = CalcTPVatStart()
    End Sub

    Public Sub ststart()

        Globals.ThisWorkbook.Application.DisplayFormulaBar = False
        Globals.Dashboard.Activate()
        Globals.Ribbons.SpartanTraderRibbon.BetaButton_Click(Nothing, Nothing)
    End Sub

    Public Sub stQuit()

        Globals.ThisWorkbook.Saved = True
        Globals.ThisWorkbook.Application.DisplayAlerts = False
        Globals.ThisWorkbook.Application.DisplayFormulaBar = True
        Globals.ThisWorkbook.Application.Quit()
        StopTimers()
    End Sub


End Module
