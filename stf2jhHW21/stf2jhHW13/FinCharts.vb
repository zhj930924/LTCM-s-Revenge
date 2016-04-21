
Public Class FinCharts

    Private Sub Sheet6_Startup() Handles Me.Startup

    End Sub

    Private Sub Sheet6_Shutdown() Handles Me.Shutdown

    End Sub

    Public Sub SetupFinCharts()
        TickerLBox.Items.Clear()
        For Each dr As DataRow In myDataSet.Tables("TickerTable").Rows
            TickerLBox.Items.Add(dr("ticker").ToString().Trim())
        Next

        StockDataToChartLO.AutoSetDataBoundColumnHeaders = True

        StockChart.ChartType = Excel.XlChartType.xlLine
        Dim y As Excel.Axis = StockChart.Axes(Excel.XlAxisType.xlValue)
        y.HasTitle = False
        y.HasMinorGridlines = True
        y.MinorTickMark = Excel.XlTickMark.xlTickMarkOutside
        y.TickLabels.NumberFormat = "$###.00"

        Dim x As Excel.Axis = StockChart.Axes(Excel.XlAxisType.xlCategory)
        x.CategoryType = Excel.XlCategoryType.xlTimeScale
        x.MajorTickMark = Excel.XlTickMark.xlTickMarkCross
        x.BaseUnit = Excel.XlTimeUnit.xlDays
        x.TickLabels.NumberFormat = "[$-409]d-mmm;@"

        SymbolLBox.Items.Clear()
        For Each dr As DataRow In myDataSet.Tables("SymbolTable").Rows
            SymbolLBox.Items.Add(dr("symbol").ToString().Trim())
        Next

        OptionDataToChartLO.AutoSetDataBoundColumnHeaders = True

        OptionChart.ChartType = Excel.XlChartType.xlLine
        Dim y2 As Excel.Axis = StockChart.Axes(Excel.XlAxisType.xlValue)
        y2.HasTitle = False
        y2.HasMinorGridlines = True
        y2.MinorTickMark = Excel.XlTickMark.xlTickMarkOutside
        y2.TickLabels.NumberFormat = "$###.00"

        Dim x2 As Excel.Axis = StockChart.Axes(Excel.XlAxisType.xlCategory)
        x2.CategoryType = Excel.XlCategoryType.xlTimeScale
        x2.MajorTickMark = Excel.XlTickMark.xlTickMarkCross
        x2.BaseUnit = Excel.XlTimeUnit.xlDays
        x2.TickLabels.NumberFormat = "[$-409]d-mmm;@"
    End Sub

    Private Sub TickerLBox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles TickerLBox.SelectedIndexChanged
        DownloadStockDataToChart(TickerLBox.SelectedItem.Trim())
        StockDataToChartLO.DataSource = myDataSet.Tables("StockDataToChart")
        StockChart.ChartTitle.Text = "Daily Closings for " + TickerLBox.SelectedItem.ToString()
        Dim y As Excel.Axis = StockChart.Axes(Excel.XlAxisType.xlValue)
        y.MinimumScale = Math.Truncate(FindMinBid("StockDataToChart") / 10) * 10
    End Sub

    Private Sub SymbolLBox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles SymbolLBox.SelectedIndexChanged
        DownloadOptionDataToChart(SymbolLBox.SelectedItem.Trim())
        OptionDataToChartLO.DataSource = myDataSet.Tables("OptionDataToChart")
        OptionChart.ChartTitle.Text = "Daily Closings for " + SymbolLBox.SelectedItem.ToString()
        Dim y As Excel.Axis = OptionChart.Axes(Excel.XlAxisType.xlValue)
        y.MinimumScale = Math.Truncate(FindMinBid("OptionDataToChart") / 10) * 10
    End Sub

    Public Function FindMinBid(tableName As String) As Double
        Dim tempMin As Double = 100000000
        For Each myRow As DataRow In myDataSet.Tables(tableName).Rows
            tempMin = Math.Min(myRow("Bid"), tempMin)
        Next
        Return tempMin
    End Function
End Class
