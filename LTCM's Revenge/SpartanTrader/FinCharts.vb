
Public Class FinCharts

    Private Sub Sheet6_Startup() Handles Me.Startup

    End Sub

    Private Sub Sheet6_Shutdown() Handles Me.Shutdown

    End Sub



    Public Sub SetUpFinCharts()
        'set up the listbox with the tickers
        TickerLBox.Items.Clear()
        For Each dr As DataRow In myDataSet.Tables("TickerTable").Rows
            TickerLBox.Items.Add(dr("Ticker").ToString().Trim())
        Next

        'set up the LO
        StockDataToChartLO.AutoSetDataBoundColumnHeaders = True

        'format the chart
        StockChart.ChartStyle = Excel.XlChartType.xlLine
        StockChart.ApplyLayout(3)
        StockChart.ChartStyle = 6

        'format the y axis as $
        Dim y As Excel.Axis = StockChart.Axes(Excel.XlAxisType.xlValue)
        y.HasTitle = False
        y.HasMinorGridlines = True
        y.MinorTickMark = Excel.XlTickMark.xlTickMarkOutside
        y.TickLabels.NumberFormat = "$###.00"

        'format the x axis as dates
        Dim x As Excel.Axis = StockChart.Axes(Excel.XlAxisType.xlCategory)
        x.CategoryType = Excel.XlCategoryType.xlTimeScale
        x.MajorTickMark = Excel.XlTickMark.xlTickMarkCross
        x.BaseUnit = Excel.XlTimeUnit.xlDays
        x.TickLabels.NumberFormat = "[$-409]d-mmm;@"

        'set up the listbox with the symbols
        SymbolLBox.Items.Clear()
        For Each drr As DataRow In myDataSet.Tables("SymbolTable").Rows
            SymbolLBox.Items.Add(drr("Symbol").ToString().Trim())
        Next

        'set up the LO
        OptionDataToChartLO.AutoSetDataBoundColumnHeaders = True

        'format the chart
        OptionChart.ChartStyle = Excel.XlChartType.xlLine
        OptionChart.ApplyLayout(3)
        OptionChart.ChartStyle = 6

        'format the y axis as $
        Dim a As Excel.Axis = OptionChart.Axes(Excel.XlAxisType.xlValue)
        a.HasTitle = False
        a.HasMinorGridlines = True
        a.MinorTickMark = Excel.XlTickMark.xlTickMarkOutside
        a.TickLabels.NumberFormat = "$###.00"

        'format the x axis as dates
        Dim b As Excel.Axis = OptionChart.Axes(Excel.XlAxisType.xlCategory)
        b.CategoryType = Excel.XlCategoryType.xlTimeScale
        b.MajorTickMark = Excel.XlTickMark.xlTickMarkCross
        b.BaseUnit = Excel.XlTimeUnit.xlDays
        b.TickLabels.NumberFormat = "[$-409]d-mmm;@"


        'TrackingChart.SetSourceData(TrackingChartLO.Range)
        'Dim s As Excel.SeriesCollection = TrackingChart.SeriesCollection
        's(0).format.line.Weight = 2
        's(0).format.line.ForeColor.RGB = System.Drawing.Color.DarkOrange
        's(1).format.line.Weight = 2
        's(1).format.line.ForeColor.RGB = System.Drawing.Color.Gray
        's(2).format.line.Weight = 2
        's(2).format.line.ForeColor.RGB = System.Drawing.Color.DarkBlue


    End Sub

    Private Sub TickerLBox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles TickerLBox.SelectedIndexChanged
        'Application.ScreenUpdating = False
        DownloadStockDataToChart(TickerLBox.SelectedItem.Trim())
        StockDataToChartLO.DataSource = myDataSet.Tables("StockDataToChart")
        StockChart.SetSourceData(StockDataToChartLO.Range)
        'Application.ScreenUpdating = True
        StockChart.ChartTitle.Text = "Daily closing for " + TickerLBox.SelectedItem.ToString()
        Dim y As Excel.Axis = StockChart.Axes(Excel.XlAxisType.xlValue)
        ' this line sets the scale of the chart for better viewing
        y.MinimumScale = Math.Truncate((FindminBid("StockDataToChart") / 10)) * 10
    End Sub

    Private Sub SymbolLBox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles SymbolLBox.SelectedIndexChanged
        'Application.ScreenUpdating = False
        DownloadOptionDataToChart(SymbolLBox.SelectedItem.Trim())
        OptionDataToChartLO.DataSource = myDataSet.Tables("OptionDataToChart")
        OptionChart.SetSourceData(OptionDataToChartLO.Range)
        'Application.ScreenUpdating = True
        OptionChart.ChartTitle.Text = "Daily closing for " + SymbolLBox.SelectedItem.ToString()
        Dim y As Excel.Axis = OptionChart.Axes(Excel.XlAxisType.xlValue)
        ' this line sets the scale of the chart for better viewing
        y.MinimumScale = Math.Truncate((FindminBid("OptionDataToChart") / 10)) * 10
    End Sub

    Public Function FindMinBid(tableName As String) As Double
        Dim tempMin As Double = 100000000
        For Each myRow As DataRow In myDataSet.Tables(tableName).Rows
            tempMin = Math.Min(myRow("Bid"), tempMin)
        Next
        Return tempMin
    End Function
End Class
