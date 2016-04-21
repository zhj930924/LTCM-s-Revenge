
Public Class Dashboard
    Public Sub PrintToAlgoLog(textToShow As String)
        If AlgoTransactionsLogTBox.TextLength > 32000 Then
            AlgoTransactionsLogTBox.Clear()
        End If
        AlgoTransactionsLogTBox.Text = " > " + textToShow + vbNewLine + AlgoTransactionsLogTBox.Text
    End Sub

    Public Sub SetSeconds(secs As Integer)
        Try
            If Math.Abs(secs) > 60 Then
                secs = 0
            End If
            secondsLeft = secs
            If Globals.ThisWorkbook.ActiveSheet.Name = "Dashboard" Then
                SecondsCell.Value = Math.Abs(secs)
                Select Case secs
                    Case Is < 0
                        SecondsCell.Font.Color = System.Drawing.Color.Orange
                    Case Is <= 5
                        SecondsCell.Font.Color = System.Drawing.Color.Red
                    Case Is <= 10
                        SecondsCell.Font.Color = System.Drawing.Color.Yellow
                    Case Else
                        SecondsCell.Font.Color = System.Drawing.Color.LightGreen
                End Select
            End If
        Catch
            ' Nada
        End Try
    End Sub

    Public Sub LoadCBoxes()
        TickerCBox.Items.Clear()
        For Each myRow As DataRow In myDataSet.Tables("TickerTable").Rows
            TickerCBox.Items.Add(myRow("Ticker").ToString().Trim())
        Next
        TickerCBox.Text = "Select Ticker"

        SymbolCBox.Items.Clear()
        For Each myRow As DataRow In myDataSet.Tables("SymbolTable").Rows
            SymbolCBox.Items.Add(myRow("Symbol").ToString().Trim())
        Next
        SymbolCBox.Text = "Select Symbol"
    End Sub

    Private Sub SellShortStockButton_Click(sender As Object, e As EventArgs) Handles SellShortStockButton.Click
        myTransaction.Clear()
        myTransaction.trType = "SellShort"
        myTransaction.typeOfPrice = "Bid"
        myTransaction.typeOfSecurity = "Stock"
        If myTransaction.IsStockInputValid() = True Then
            myTransaction.ComputeTransactionProperties()
            myTransaction.DisplayTransactionData()
        End If

    End Sub

    Private Sub SellStockButton_Click(sender As Object, e As EventArgs) Handles SellStockButton.Click
        myTransaction.Clear()
        myTransaction.trType = "Sell"
        myTransaction.typeOfPrice = "Bid"
        myTransaction.typeOfSecurity = "Stock"
        If myTransaction.IsStockInputValid() = True Then
            myTransaction.ComputeTransactionProperties()
            myTransaction.DisplayTransactionData()
        End If

    End Sub

    Private Sub CashDivButton_Click(sender As Object, e As EventArgs) Handles CashDivButton.Click
        myTransaction.Clear()
        myTransaction.trType = "CashDiv"
        myTransaction.typeOfPrice = "Div"
        myTransaction.typeOfSecurity = "Stock"
        If myTransaction.IsStockInputValid() = True Then
            myTransaction.ComputeTransactionProperties()
            myTransaction.DisplayTransactionData()
        End If

    End Sub

    Private Sub ExecuteStockTransactionButton_Click(sender As Object, e As EventArgs) Handles ExecuteStockTransactionButton.Click
        If myTransaction.IsStockInputValid() = True Then
            myTransaction.ComputeTransactionProperties()
            If IsTransactionValid(myTransaction.trType, myTransaction.symbol, myTransaction.qty) Then
                myTransaction.ExecuteTransaction()
                HighlightTransaction()
                CalcFinancialMetrics(currentDate)
                myTransaction.DisplayTransactionData()
                DisplayFinancialMetrics(currentDate)
            End If
        Else
                MessageBox.Show("Beep. Boop. Stock input not valid.")
        End If
    End Sub

    Public Sub HighlightTransaction()
        Globals.Dashboard.Range("C4:C6").Font.Color = RGB(0, 255, 0)
    End Sub

    Public Sub ClearTransactionHighlight()
        Globals.Dashboard.Range("C4:C6").Font.Color = RGB(255, 255, 255)
    End Sub

    Private Sub BuyStockButton_Click(sender As Object, e As EventArgs) Handles BuyStockButton.Click
        myTransaction.Clear()
        myTransaction.trType = "Buy"
        myTransaction.typeOfPrice = "Ask"
        myTransaction.typeOfSecurity = "Stock"
        If myTransaction.IsStockInputValid() = True Then
            myTransaction.ComputeTransactionProperties()
            myTransaction.DisplayTransactionData()
        End If

    End Sub


    Public Sub SetupTrackingChart()
        TrackingChart.ChartType = Excel.XlChartType.xlLine
        TrackingChart.ChartStyle = 8
        TrackingChart.ApplyLayout(3)
        TrackingChart.HasTitle = False
        TrackingChart.HasLegend = True

        Dim y As Excel.Axis = TrackingChart.Axes(Excel.XlAxisType.xlValue)
        y.HasTitle = False
        y.HasMinorGridlines = True
        y.MinorTickMark = Excel.XlTickMark.xlTickMarkOutside
        y.TickLabels.NumberFormat = "$#,###"
        y.MinimumScaleIsAuto = False
        y.MaximumScaleIsAuto = True

        Dim x As Excel.Axis = TrackingChart.Axes(Excel.XlAxisType.xlCategory)
        x.CategoryType = Excel.XlCategoryType.xlTimeScale
        x.MajorTickMark = Excel.XlTickMark.xlTickMarkCross
        x.BaseUnit = Excel.XlTimeUnit.xlDays
        x.TickLabels.NumberFormat = "[$-409]d-mmm;@"

        TrackingChart.SetSourceData(TrackingChartLO.Range)
        Dim s As Excel.SeriesCollection = TrackingChart.SeriesCollection
        s(0).Format.Line.Weight = 2
        s(0).Format.Line.ForeColor.RGB = System.Drawing.Color.DarkOrange
        s(1).Format.Line.Weight = 2
        s(1).Format.Line.ForeColor.RGB = System.Drawing.Color.Gray
        s(2).Format.Line.Weight = 2
        s(2).Format.Line.ForeColor.RGB = System.Drawing.Color.DarkBlue
    End Sub

    Public Sub FillTPVTrackingTable()
        If myDataSet.Tables.Contains("TPVTrackingTable") Then
            myDataSet.Tables("TPVTrackingTable").Clear()
        Else
            myDataSet.Tables.Add("TPVTrackingTable")
            myDataSet.Tables("TPVTrackingTable").Columns.Add("Date", GetType(Date))
            myDataSet.Tables("TPVTrackingTable").Columns.Add("TaTPV", GetType(Double))
            myDataSet.Tables("TPVTrackingTable").Columns.Add("NoHedge", GetType(Double))
            myDataSet.Tables("TPVTrackingTable").Columns.Add("TPV", GetType(Double))
        End If

        Dim tempTaTPV, tempTPV, tempNoHedge As Double
        Dim targetdate As Date

        For i As Integer = 14 To 0 Step -1
            targetdate = currentDate.AddDays(-i)
            If targetdate >= startDate Then
                DownloadPricesForOneDay(targetdate)
                tempTPV = CalcTPV(targetdate)
                tempTaTPV = CalcTaTPV(targetdate)
                tempNoHedge = CalcTPVNoHedge(targetdate)
                UpdateTPVTrackingTable(targetdate, tempTPV, tempTaTPV, tempNoHedge)
            End If

        Next
        TrackingChartLO.DataSource = myDataSet.Tables("TPVTrackingTable")
    End Sub

    Public Sub UpdateTPVTrackingTable(targetDate As Date, tpvInput As Double, tatpvInput As Double, noHedgeInput As Double)
        For Each myRow As DataRow In myDataSet.Tables("TPVTrackingTable").Rows
            If myRow("Date") = targetDate.ToShortDateString Then
                myRow("TPV") = tpvInput
                myRow("TaTPV") = tatpvInput
                myRow("NoHedge") = noHedgeInput
                Return
            End If
        Next
        myDataSet.Tables("TPVTrackingTable").Rows.Add(targetDate, tatpvInput, noHedgeInput, tpvInput)

        Try
            Dim y As Excel.Axis = TrackingChart.Axes(Excel.XlAxisType.xlValue)
            y.MinimumScale = Math.Truncate((FindMinInTPVTrackingTable() / 10000000)) * 10000000
        Catch

        End Try
    End Sub

    Public Function FindMinInTPVTrackingTable() As Double
        Dim TempMin As Double = 100000000
        For Each myRow As DataRow In myDataSet.Tables("TPVTrackingTable").Rows
            TempMin = Math.Min(myRow("TPV"), TempMin)
            TempMin = Math.Min(myRow("TATPV"), TempMin)
            TempMin = Math.Min(myRow("NoHedge"), TempMin)
        Next
        Return TempMin
    End Function

    Private Sub StockQtyTBox_TextChanged(sender As Object, e As EventArgs) Handles StockQtyTBox.TextChanged

    End Sub

    Private Sub BuyOptionButton_Click(sender As Object, e As EventArgs) Handles BuyOptionButton.Click
        myTransaction.Clear()
        myTransaction.trType = "Buy"
        myTransaction.typeOfPrice = "Ask"
        myTransaction.typeOfSecurity = "Option"
        If myTransaction.IsOptionInputValid() = True Then
            myTransaction.ComputeTransactionProperties()
            myTransaction.DisplayTransactionData()
        End If
    End Sub

    Private Sub SellOptionButton_Click(sender As Object, e As EventArgs) Handles SellOptionButton.Click
        myTransaction.Clear()
        myTransaction.trType = "Sell"
        myTransaction.typeOfPrice = "Bid"
        myTransaction.typeOfSecurity = "Option"
        If myTransaction.IsOptionInputValid() = True Then
            myTransaction.ComputeTransactionProperties()
            myTransaction.DisplayTransactionData()
        End If
    End Sub

    Private Sub SellShortOptionButton_Click(sender As Object, e As EventArgs) Handles SellShortOptionButton.Click
        myTransaction.Clear()
        myTransaction.trType = "SellShort"
        myTransaction.typeOfPrice = "Bid"
        myTransaction.typeOfSecurity = "Option"
        If myTransaction.IsOptionInputValid() = True Then
            myTransaction.ComputeTransactionProperties()
            myTransaction.DisplayTransactionData()
        End If
    End Sub

    Private Sub ExerciseOptionButton_Click(sender As Object, e As EventArgs) Handles ExerciseOptionButton.Click
        myTransaction.Clear()
        myTransaction.trType = ""
        myTransaction.typeOfPrice = "Strike"
        myTransaction.typeOfSecurity = "Option"
        If myTransaction.IsOptionInputValid() = True Then
            myTransaction.ComputeTransactionProperties()
            myTransaction.DisplayTransactionData()
        End If
    End Sub

    Private Sub ExecuteOptionTransactionButton_Click(sender As Object, e As EventArgs) Handles ExecuteOptionTransactionButton.Click
        If myTransaction.IsOptionInputValid() = True Then
            myTransaction.ComputeTransactionProperties()
            If IsTransactionValid(myTransaction.trType, myTransaction.symbol, myTransaction.qty) Then
                myTransaction.ExecuteTransaction()
                HighlightTransaction()
                CalcFinancialMetrics(currentDate)
                myTransaction.DisplayTransactionData()
                DisplayFinancialMetrics(currentDate)
            End If
        Else
                MessageBox.Show("Beep. Boop. Option input not valid.")
        End If
    End Sub

    Private Sub Trade0Button_Click(sender As Object, e As EventArgs) Handles Trade0Button.Click
        AlgoHedge(0)
    End Sub

    Private Sub Trade1Button_Click(sender As Object, e As EventArgs) Handles Trade1Button.Click
        AlgoHedge(1)
    End Sub

    Private Sub Trade2Button_Click(sender As Object, e As EventArgs) Handles Trade2Button.Click
        AlgoHedge(2)
    End Sub

    Private Sub Trade3Button_Click(sender As Object, e As EventArgs) Handles Trade3Button.Click
        AlgoHedge(3)
    End Sub

    Private Sub Trade4Button_Click(sender As Object, e As EventArgs) Handles Trade4Button.Click
        AlgoHedge(4)
    End Sub

    Private Sub Trade5Button_Click(sender As Object, e As EventArgs) Handles Trade5Button.Click
        AlgoHedge(5)
    End Sub

    Private Sub Trade6Button_Click(sender As Object, e As EventArgs) Handles Trade6Button.Click
        AlgoHedge(6)
    End Sub

    Private Sub Trade7Button_Click(sender As Object, e As EventArgs) Handles Trade7Button.Click
        AlgoHedge(7)
    End Sub

    Private Sub Trade8Button_Click(sender As Object, e As EventArgs) Handles Trade8Button.Click
        AlgoHedge(8)
    End Sub

    Private Sub Trade9Button_Click(sender As Object, e As EventArgs) Handles Trade9Button.Click
        AlgoHedge(9)
    End Sub

    Private Sub Trade10Button_Click(sender As Object, e As EventArgs) Handles Trade10Button.Click
        AlgoHedge(10)
    End Sub

    Private Sub Trade11Button_Click(sender As Object, e As EventArgs) Handles Trade11Button.Click
        AlgoHedge(11)
    End Sub
End Class
