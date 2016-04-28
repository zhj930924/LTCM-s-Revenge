Module ProcessAutomation

    Public WithEvents spyTimer As Timer
    Public WithEvents secondsTimer As Timer

    Public Sub WarmStart()
        StopTimers()
        ClearOldData()
        CreateAndConnectTheADOObjects()
        If OpenDBConnection() = False Then
            Exit Sub ' beacuse it could not connect
        End If
        If StockPricesExist() = False Then
            Exit Sub 'because the DB is empty()
        End If
        DownloadStaticData()
        DownloadteamData()
        SetUpRecommendations()
        Select Case traderMode
            Case "Manual"
                currentDate = DownloadCurrentDate()
                SetupCharts()
                DailyRoutine()
            Case "Synch"
                currentDate = DownloadCurrentDate()
                SetupCharts()
                StartTimers()
            Case "Simulation"
                Globals.Dashboard.Activate()
                currentDate = startDate
                lastTransactionDate = startDate
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
        myTransaction.clear()
        Globals.Dashboard.CurrentDateCell.Value = currentDate.ToLongDateString()
        DownloadPricesForOneDay(currentDate)
        CalcFinancialMetrics(currentDate)
        DisplayFinancialMetrics(currentDate)

        'Globals.Dashboard.DisplayFamilyDeltas()
        Select Case traderMode
            Case "Manual"
                'nothing at this point
            Case "Synch"

                CalcFinancialMetrics(currentDate)
                'DoScheduledTransactions()
                DisplayFinancialMetrics(currentDate)
            Case "Simulation"

                CalcFinancialMetrics(currentDate)
                DoScheduledTransactions()
                DisplayFinancialMetrics(currentDate)
                AlgoHedgeAll()
            Case "Auto"

                CalcFinancialMetrics(currentDate)
                DoScheduledTransactions()
                DisplayFinancialMetrics(currentDate)
                AlgoHedgeAll()
        End Select
    End Sub

    Public Sub StartTimers()
        spyTimer = New Timer
        spyTimer.Interval = 2000  ' do not change these settings
        secondsTimer = New Timer
        secondsTimer.Interval = 1000
        Globals.Dashboard.SetSeconds(0)   ' reset the screen countdown
        secondsTimer.Start()
        spyTimer.Start()
    End Sub
    Public Sub StopTimers()
        If IsNothing(spyTimer) Then
            'skip
        Else
            spyTimer.Stop()
            secondsTimer.Stop()
            Globals.Dashboard.SetSeconds(0)  ' reset the screen timer
        End If
    End Sub

    Private Sub secondsTimer_tick() Handles secondsTimer.Tick
        Globals.Dashboard.SetSeconds(secondsLeft - 1)
    End Sub

    Private Sub spyTimer_Tick() Handles spyTimer.Tick
        Dim tempNewDate As Date
        tempNewDate = DownloadCurrentDate2()
        If tempNewDate.Date <> currentDate.Date Then   ' it is a new day!
            currentDate = tempNewDate
            Globals.Dashboard.SetSeconds(60)
            DailyRoutine()
        End If
    End Sub

    Public Sub CalcFinancialMetrics(targetDate As Date)
        'interestSLT = CalcInterestSLT(targetDate)  <--- deleted
        ' CAccount = CAccount + interestSLT   <--- deleted
        margin = CalcMargin(targetDate)
        IP = CalcIPValue(targetDate)
        AP = CalcAPValue(targetDate)
        TPV = IP + AP + CAccount
        TaTPV = CalcTaTPV(targetDate)
        If TaTPV >= TPV Then
            NeedMoreCapital = True
            ExcessMargin = maxMargins - margin
        ElseIf TaTPV < TPV Then
            NeedMoreCapital = False
        End If

        If margin > maxMargins Or margin * 0.3 >= CAccount Then
            MarginTripped = True
        End If

        TE = TPV - TaTPV
        If TE > 0 Then TE = TE / 4  'If a gain then...
        TEpercent = TE / TaTPV
        TPVNoHedge = CalcTPVNoHedge(targetDate)
        Globals.Dashboard.UpdateTPVTrackingTable(targetDate, TPV, TaTPV, TPVNoHedge)
        RecommendHedges()
    End Sub
    Public Sub DownloadteamData()
        CAccount = DownloadCAccount()
        DownloadAcquiredPositions()
        lastTransactionDate = DownloadLastTransactionDate(endDate)
    End Sub

    Public Sub SetupCharts()
        Globals.Dashboard.FillTPVTrackingTable()
        Globals.Dashboard.SetupTrackingChart()
        Globals.FinCharts.SetUpFinCharts()
    End Sub

    'HM14  21
    Public Sub ClearOldData()
        Globals.Markets.StockMarketLO.DataSource = Nothing
        Globals.Markets.OptionMarketLO.DataSource = Nothing
        Globals.Markets.SP500LO.DataSource = Nothing
        Globals.Portfolio.InitialPositionsLO.DataSource = Nothing
        Globals.Portfolio.AcquiredPositionsLO.DataSource = Nothing
        Globals.Environment.SettingsLO.DataSource = Nothing
        Globals.Environment.TransactionCostsLO.DataSource = Nothing
        Globals.Portfolio.ConfirmationTicketsLO.DataSource = Nothing
        Globals.Dashboard.AlgoTransactionsLogTBox.Text = "Algo Transaction Log - Ready."
        Globals.Dashboard.Activate()
        lastPriceDownloadDate = "1/1/2000"
        ResetRecommendations()
        DisplayRecommendations()
    End Sub

    Public Sub DownloadStaticData()
        DownloadInitialPositions()
        DownloadTransactionCost()
        DownloadEnvironmentVariable()
        DownloadTickers()
        DownloadSymbols()
        Globals.Dashboard.LoadCBoxes()
        ' we use 'get' to indicate that we are extracting data from the dataset, and 'download'
        ' to indicate that we are extracting data from the database
        initialCAccount = GetInitialCAccount()
        iRate = GetIRate()
        startDate = GetStartDate()
        endDate = GetEndDate()
        maxMargins = GetMaxMargins()
        TPVatStart = CalcTPVAtStart()
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
            'do nothing
        End Try

    End Sub

    'HM13
    Public Sub stStart()
        Globals.ThisWorkbook.Application.DisplayFormulaBar = False
        Globals.Dashboard.Activate()
        Globals.Ribbons.SpartanTraderRibbon.BetaBtn_Click(Nothing, Nothing)
    End Sub

    Public Sub stQuit()
        StopTimers()
        Globals.ThisWorkbook.Application.DisplayAlerts = False
        Globals.ThisWorkbook.Application.DisplayFormulaBar = True
        Globals.ThisWorkbook.Application.Quit()
    End Sub



End Module
