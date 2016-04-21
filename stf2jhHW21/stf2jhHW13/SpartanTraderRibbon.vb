Imports Microsoft.Office.Tools.Ribbon

Public Class stRibbon

    Private Sub FinChartsButton_Click(sender As Object, e As RibbonControlEventArgs) Handles FinChartsButton.Click
        Globals.FinCharts.Activate()
    End Sub



    Private Sub SpartanTraderRibbon_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load
        Globals.Ribbons.SpartanTraderRibbon.RibbonUI.ActivateTabMso("TabAddIns")
    End Sub

    Public Sub QuitButton_Click(sender As Object, e As RibbonControlEventArgs) Handles QuitButton.Click
        stQuit()
    End Sub

    Public Sub BetaButton_Click(sender As Object, e As RibbonControlEventArgs) Handles BetaButton.Click
        AlphaButton.Checked = False
        BetaButton.Checked = True
        GammaButton.Checked = False
        activeDB = "Beta"
        ManualButton_Click(Nothing, Nothing)
    End Sub

    Public Sub DashboardButton_Click(sender As Object, e As RibbonControlEventArgs) Handles DashboardButton.Click
        Globals.Dashboard.Activate()
    End Sub

    Public Sub AlphaButton_Click(sender As Object, e As RibbonControlEventArgs) Handles AlphaButton.Click
        AlphaButton.Checked = True
        BetaButton.Checked = False
        GammaButton.Checked = False
        activeDB = "Alpha"
        ManualButton_Click(Nothing, Nothing)
    End Sub

    Private Sub InitialPositionsButton_Click(sender As Object, e As RibbonControlEventArgs) Handles InitialPositionsButton.Click
        DownloadInitialPositions()
        ShowInitialPositions()
    End Sub

    Public Sub AcquiredPositionsButton_Click(sender As Object, e As RibbonControlEventArgs) Handles AcquiredPositionsButton.Click
        DownloadAcquiredPositions()
        ShowAcquiredPositions()
    End Sub

    Private Sub StockMktButton_Click(sender As Object, e As RibbonControlEventArgs) Handles StockMktButton.Click
        DownloadStockMarket()
        ShowStockMarket()
    End Sub

    Private Sub OptionMktButton_Click(sender As Object, e As RibbonControlEventArgs) Handles OptionMktButton.Click
        DownloadOptionMarket()
        ShowOptionMarket()
    End Sub

    Private Sub SP500MktButton_Click(sender As Object, e As RibbonControlEventArgs) Handles SP500Button.Click
        DownloadStockIndex()
        ShowStockIndex()
    End Sub

    Private Sub SettingsButton_Click(sender As Object, e As RibbonControlEventArgs) Handles SettingsButton.Click
        DownloadEnvironmentVariable()
        ShowEnvironmentVariable()
    End Sub

    Private Sub TransactionButton_Click(sender As Object, e As RibbonControlEventArgs) Handles TransactionButton.Click
        DownloadTransactionCost()
        ShowTransactionCost()
    End Sub

    Private Sub TransactionQButton_Click(sender As Object, e As RibbonControlEventArgs) Handles TransactionQButton.Click
        DownloadTransactionQueue(teamID)
        ShowTransactionQueue()
    End Sub

    Private Sub ResetAPButton_Click(sender As Object, e As RibbonControlEventArgs) Handles ResetAPButton.Click
        ResetAP()
    End Sub

    Private Sub EditAPButton_Click(sender As Object, e As RibbonControlEventArgs) Handles EditAPButton.Click
        AcquiredPositionsButton_Click(Nothing, Nothing)
        For i = 1 To 20
            myDataSet.Tables(portfolioTableName).Rows.Add()
        Next
        Globals.Portfolio.Activate()
    End Sub

    Private Sub UploadAPButton_Click(sender As Object, e As RibbonControlEventArgs) Handles UploadAPButton.Click
        UploadScreenPortfolioToDB()
    End Sub

    Private Sub ExcludeIPButton_Click(sender As Object, e As RibbonControlEventArgs) Handles ExcludeIPButton.Click
        excludeIP = ExcludeIPButton.Checked
        ManualButton_Click(Nothing, Nothing)
    End Sub

    Private Sub ConfirmationButton_Click(sender As Object, e As RibbonControlEventArgs) Handles ConfirmationButton.Click
        DownloadConfirmationTickets()
        ShowConfirmationTickets()
    End Sub

    Private Sub ManualButton_Click(sender As Object, e As RibbonControlEventArgs) Handles ManualButton.Click
        Globals.Dashboard.CurrentDateCell.Value = "Spartan Trader is Offline"
        traderMode = "Manual"
        TurnOffControlButtons()
        ManualButton.Checked = True
        WarmStart()
    End Sub

    Private Sub TurnOffControlButtons()
        ManualButton.Checked = False
        SynchButton.Checked = False
        SimulationButton.Checked = False
        AutoHedgeButton.Checked = False
    End Sub

    Private Sub SynchButton_Click(sender As Object, e As RibbonControlEventArgs) Handles SynchButton.Click
        Globals.Dashboard.CurrentDateCell.Value = "Spartan Trader is Offline"
        traderMode = "Synch"
        TurnOffControlButtons()
        SynchButton.Checked = True
        WarmStart()
    End Sub

    Private Sub SimulationButton_Click(sender As Object, e As RibbonControlEventArgs) Handles SimulationButton.Click
        Globals.Dashboard.CurrentDateCell.Value = "Spartan Trader is Offline"
        traderMode = "Simulation"
        TurnOffControlButtons()
        SimulationButton.Checked = True
        WarmStart()
    End Sub

    Private Sub AutoHedgeButton_Click(sender As Object, e As RibbonControlEventArgs) Handles AutoHedgeButton.Click
        Globals.Dashboard.CurrentDateCell.Value = "Spartan Trader is Offline"
        traderMode = "Auto"
        TurnOffControlButtons()
        AutoHedgeButton.Checked = True
        WarmStart()
    End Sub
End Class

