Imports Microsoft.Office.Tools.Ribbon

Public Class SpartanTraderRibbon

    Private Sub TransactionQueueBtn_Click(sender As Object, e As RibbonControlEventArgs) Handles TransactionQueueBtn.Click
        DownloadTransactionQueue(teamID)
        ShowTransactionQueue()
    End Sub

    'HM16
    Private Sub DashboardBtn_Click(sender As Object, e As RibbonControlEventArgs) Handles DashboardBtn.Click
        Globals.Dashboard.Activate()
    End Sub

    'HM13
    Private Sub SpartanTraderRibbon_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load
        'This activates the custom ribbon at start
        Globals.Ribbons.SpartanTraderRibbon.RibbonUI.ActivateTabMso("TabAddIns")
    End Sub

    Private Sub AlphaBtn_Click(sender As Object, e As RibbonControlEventArgs) Handles AlphaBtn.Click
        AlphaBtn.Checked = True
        BetaBtn.Checked = False
        GammaBtn.Checked = False
        activeDB = "Alpha"
        ManualBtn_Click(Nothing, Nothing)
        'WarmStart()
    End Sub

    Public Sub BetaBtn_Click(sender As Object, e As RibbonControlEventArgs) Handles BetaBtn.Click
        AlphaBtn.Checked = False
        BetaBtn.Checked = True
        GammaBtn.Checked = False
        activeDB = "Beta"
        WarmStart()
    End Sub

    Private Sub GammaBtn_Click(sender As Object, e As RibbonControlEventArgs) Handles GammaBtn.Click
        AlphaBtn.Checked = False
        BetaBtn.Checked = False
        GammaBtn.Checked = True
        activeDB = "Gamma"
        WarmStart()
    End Sub

    Public Sub QuitBtn_Click(sender As Object, e As RibbonControlEventArgs) Handles QuitBtn.Click
        stQuit()
    End Sub

    Private Sub InitialPositionsBtn_Click(sender As Object, e As RibbonControlEventArgs) Handles InitialPositionsBtn.Click
        DownloadInitialPositions()
        ShowInitialPositions()
    End Sub

    Public Sub AcquiredPositionsBtn_Click(sender As Object, e As RibbonControlEventArgs) Handles AcquiredPositionsBtn.Click
        DownloadAcquiredPositions()
        ShowAcquiredPositions()
    End Sub

    Private Sub StockMktBtn_Click(sender As Object, e As RibbonControlEventArgs) Handles StockMktBtn.Click
        DownloadStockMarket()
        ShowStockMarket()
    End Sub

    Private Sub OptionMktBtn_Click(sender As Object, e As RibbonControlEventArgs) Handles OptionMktBtn.Click
        DownloadOptionMarket()
        ShowOptionMarket()
    End Sub

    Private Sub SP500Btn_Click(sender As Object, e As RibbonControlEventArgs) Handles SP500Btn.Click
        DownloadStockIndex()
        ShowStockIndex()
    End Sub

    Private Sub SettingsBtn_Click(sender As Object, e As RibbonControlEventArgs) Handles SettingsBtn.Click
        DownloadEnvironmentVariable()
        ShowEnvironmentVariable()
    End Sub

    Private Sub TCostsBtn_Click(sender As Object, e As RibbonControlEventArgs) Handles TCostsBtn.Click
        DownloadTransactionCost()
        ShowTransactionCost()
    End Sub

    Private Sub FinChartsBtn_Click(sender As Object, e As RibbonControlEventArgs) Handles FinChartsBtn.Click
        Globals.FinCharts.Activate()
    End Sub

    Private Sub ResetAPBtn_Click(sender As Object, e As RibbonControlEventArgs) Handles ResetAPBtn.Click
        ResetAP()
    End Sub

    Private Sub EditAPBtn_Click(sender As Object, e As RibbonControlEventArgs) Handles EditAPBtn.Click
        AcquiredPositionsBtn_Click(Nothing, Nothing)
        For i = 1 To 20
            myDataSet.Tables(portfolioTableName).Rows.Add()
        Next
        Globals.Portfolio.Activate()
    End Sub

    Private Sub UploadAPBtn_Click(sender As Object, e As RibbonControlEventArgs) Handles UploadAPBtn.Click
        UploadScreenPortfolioToDB()
    End Sub

    Private Sub ExcludeIPBtn_Click(sender As Object, e As RibbonControlEventArgs) Handles ExcludeIPBtn.Click
        excludeIP = ExcludeIPBtn.Checked
        ManualBtn_Click(Nothing, Nothing)
        'WarmStart()
    End Sub

    Private Sub ConfirmationBtn_Click(sender As Object, e As RibbonControlEventArgs) Handles ConfirmationBtn.Click
        DownloadConfirmationTickets()
        ShowConfrimationTickets()
    End Sub

    Private Sub ManualBtn_Click(sender As Object, e As RibbonControlEventArgs) Handles ManualBtn.Click
        Globals.Dashboard.CurrentDateCell.Value = "Spartan Trader is Offline"
        traderMode = "Manual"
        TurnOffControlButtons()
        ManualBtn.Checked = True
        WarmStart()
    End Sub

    Private Sub SynchBtn_Click(sender As Object, e As RibbonControlEventArgs) Handles SynchBtn.Click
        Globals.Dashboard.CurrentDateCell.Value = "Spartan Trader is Offline"
        traderMode = "Synch"
        TurnOffControlButtons()
        SynchBtn.Checked = True
        WarmStart()
    End Sub

    Private Sub SimulationBtn_Click(sender As Object, e As RibbonControlEventArgs) Handles SimulationBtn.Click
        Globals.Dashboard.CurrentDateCell.Value = "Spartan Trader is Offline"
        traderMode = "Simulation"
        TurnOffControlButtons()
        SimulationBtn.Checked = True
        WarmStart()
    End Sub

    Private Sub AutoHedgeBtn_Click(sender As Object, e As RibbonControlEventArgs) Handles AutoHedgeBtn.Click
        Globals.Dashboard.CurrentDateCell.Value = "Spartan Trader is Offline"
        traderMode = "Auto"
        TurnOffControlButtons()
        AutoHedgeBtn.Checked = True
        WarmStart()
    End Sub

    Private Sub TurnOffControlButtons()
        ManualBtn.Checked = False
        SynchBtn.Checked = False
        SimulationBtn.Checked = False
        AutoHedgeBtn.Checked = False
    End Sub
End Class
