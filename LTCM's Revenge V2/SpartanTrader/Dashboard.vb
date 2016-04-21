
Public Class Dashboard

    Private Sub Sheet1_Startup() Handles Me.Startup

    End Sub

    Private Sub Sheet1_Shutdown() Handles Me.Shutdown

    End Sub






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
        Catch ex As Exception
            'skip
        End Try
    End Sub
    'HM19
    Public Sub LoadCBoxes()
        TickerCBox.Items.Clear()
        For Each myrow As DataRow In myDataSet.Tables("TickerTable").Rows
            TickerCBox.Items.Add(myrow("Ticker").ToString().Trim())
        Next
        TickerCBox.Text = "Select Ticker"

        SymbolCbox.Items.Clear()
        For Each myrow As DataRow In myDataSet.Tables("SymbolTable").Rows
            SymbolCbox.Items.Add(myrow("Symbol").ToString().Trim())
        Next
        SymbolCbox.Text = "Select Symbol"
    End Sub

    Private Sub BuyStockBtn_Click(sender As Object, e As EventArgs) Handles BuyStockBtn.Click
        myTransaction.clear()
        myTransaction.trType = "Buy"
        myTransaction.typeOfPrice = "Ask"
        myTransaction.typeOfSecurity = "Stock"
        If myTransaction.IsStockInputValid() = True Then
            myTransaction.ComputeTransactionProperties()
            myTransaction.DisplayTransactionData()
        End If
    End Sub

    Private Sub SellStockBtn_Click(sender As Object, e As EventArgs) Handles SellStockBtn.Click
        myTransaction.clear()
        myTransaction.trType = "Sell"
        myTransaction.typeOfPrice = "Bid"
        myTransaction.typeOfSecurity = "Stock"
        If myTransaction.IsStockInputValid() = True Then
            myTransaction.ComputeTransactionProperties()
            myTransaction.DisplayTransactionData()
        End If
    End Sub

    Private Sub SellShortStockBtn_Click(sender As Object, e As EventArgs) Handles SellShortStockBtn.Click
        myTransaction.clear()
        myTransaction.trType = "SellShort"
        myTransaction.typeOfPrice = "Bid"
        myTransaction.typeOfSecurity = "Stock"
        If myTransaction.IsStockInputValid() = True Then
            myTransaction.ComputeTransactionProperties()
            myTransaction.DisplayTransactionData()
        End If
    End Sub

    Private Sub CashDivBtn_Click(sender As Object, e As EventArgs) Handles CashDivBtn.Click
        myTransaction.clear()
        myTransaction.trType = "CashDiv"
        myTransaction.typeOfPrice = "Div"
        myTransaction.typeOfSecurity = "Stock"
        If myTransaction.IsStockInputValid() = True Then
            myTransaction.ComputeTransactionProperties()
            myTransaction.DisplayTransactionData()
        End If
    End Sub

    Private Sub ExecuteStockTransactionBtn_Click(sender As Object, e As EventArgs) Handles ExecuteStockTransactionBtn.Click
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
            MessageBox.Show("I cannot do this for you, Dave. Stock input not valid.")
        End If
    End Sub

    Public Sub HighlightTransaction()
        Globals.Dashboard.Range("C4:C6").Font.Color = RGB(0, 255, 0)
        ' for color and their codes see http://condor.depaul.edu/sjost/it236/documents/colorNames.htm
    End Sub

    Public Sub ClearTransactionHighlight()
        Globals.Dashboard.Range("C4:C6").Font.Color = RGB(255, 255, 255)
        ' for color and their codes see http://condor.depaul.edu/sjost/it236/documents/colorNames.htm
    End Sub

    'HM15
    Public Sub SetupTrackingChart()
        'format the chart - feel free to change any formatting
        TrackingChart.ChartType = Excel.XlChartType.xlLine
        TrackingChart.ChartStyle = 8
        TrackingChart.ApplyLayout(3)
        TrackingChart.HasTitle = False
        TrackingChart.HasLegend = True

        'format the y axis as $
        Dim y As Excel.Axis = TrackingChart.Axes(Excel.XlAxisType.xlValue)
        y.HasTitle = False
        y.HasMinorGridlines = True
        y.MinorTickMark = Excel.XlTickMark.xlTickMarkOutside
        y.TickLabels.NumberFormat = "$#,###"
        y.MinimumScaleIsAuto = False
        y.MaximumScaleIsAuto = True

        ' format the x axis as dates
        Dim x As Excel.Axis = TrackingChart.Axes(Excel.XlAxisType.xlCategory)
        x.CategoryType = Excel.XlCategoryType.xlTimeScale
        x.MajorTickMark = Excel.XlTickMark.xlTickMarkCross
        x.BaseUnit = Excel.XlTimeUnit.xlDays
        x.TickLabels.NumberFormat = "[$-409]d-mmm;@"

        TrackingChart.SetSourceData(TrackingChartLO.Range)
        Dim s As Excel.SeriesCollection = TrackingChart.SeriesCollection
        s(0).format.line.Weight = 2
        s(0).format.line.ForeColor.RGB = System.Drawing.Color.DarkOrange
        s(1).format.line.Weight = 2
        s(1).format.line.ForeColor.RGB = System.Drawing.Color.Gray
        s(2).format.line.Weight = 2
        s(2).format.line.ForeColor.RGB = System.Drawing.Color.DarkBlue
    End Sub
    Private Sub Chart_1_SelectEvent(ElementID As Integer, Arg1 As Integer, Arg2 As Integer) Handles TrackingChart.SelectEvent

    End Sub

    Public Sub FillTPVTrackingTable()
        'clear the table
        If myDataSet.Tables.Contains("TPVTrackingTable") Then
            myDataSet.Tables("TPVTrackingTable").Clear()
        Else
            'create the table
            myDataSet.Tables.Add("TPVTrackingTable")
            myDataSet.Tables("TPVTrackingTable").Columns.Add("Date", GetType(Date))
            myDataSet.Tables("TPVTrackingTable").Columns.Add("TaTPV", GetType(Double))
            myDataSet.Tables("TPVTrackingTable").Columns.Add("NoHedge", GetType(Double))
            myDataSet.Tables("TPVTrackingTable").Columns.Add("TPV", GetType(Double))
        End If

        'fill it
        Dim tempTaTPV, tempTPV, tempNoHedge As Double
        Dim targetdate As Date

        For i As Integer = 14 To 0 Step -1
            targetdate = currentDate.AddDays(-i)
            If targetdate >= startDate Then
                DownloadPricesForOneDay(targetdate)
                tempTPV = CalCTPV(targetdate)
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
            'this line sets the scale of the chart for better viewing
            Dim y As Excel.Axis = TrackingChart.Axes(Excel.XlAxisType.xlValue)
            y.MinimumScale = Math.Truncate((FindminInTPVTrackingTable() / 10000000)) * 10000000
        Catch ex As Exception
            'skip screen refresh errors
        End Try
    End Sub

    Public Function FindminInTPVTrackingTable() As Double
        Dim tempMin As Double = 100000000
        For Each myRow As DataRow In myDataSet.Tables("TPVTrackingTable").Rows
            tempMin = Math.Min(myRow("TPV"), tempMin)
            tempMin = Math.Min(myRow("TaTPV"), tempMin)
            tempMin = Math.Min(myRow("NoHedge"), tempMin)
        Next
        Return tempMin
    End Function

    Private Sub BuyOptionBtn_Click(sender As Object, e As EventArgs) Handles BuyOptionBtn.Click
        myTransaction.clear()
        myTransaction.trType = "Buy"
        myTransaction.typeOfPrice = "Ask"
        myTransaction.typeOfSecurity = "Option"
        If myTransaction.IsOptionInputValid() = True Then
            myTransaction.ComputeTransactionProperties()
            myTransaction.DisplayTransactionData()
        End If
    End Sub

    Private Sub SellOptionBtn_Click(sender As Object, e As EventArgs) Handles SellOptionBtn.Click
        myTransaction.clear()
        myTransaction.trType = "Sell"
        myTransaction.typeOfPrice = "Bid"
        myTransaction.typeOfSecurity = "Option"
        If myTransaction.IsOptionInputValid() = True Then
            myTransaction.ComputeTransactionProperties()
            myTransaction.DisplayTransactionData()
        End If
    End Sub

    Private Sub SellShortOptionBtn_Click(sender As Object, e As EventArgs) Handles SellShortOptionBtn.Click
        myTransaction.clear()
        myTransaction.trType = "SellShort"
        myTransaction.typeOfPrice = "Bid"
        myTransaction.typeOfSecurity = "Option"
        If myTransaction.IsOptionInputValid() = True Then
            myTransaction.ComputeTransactionProperties()
            myTransaction.DisplayTransactionData()
        End If
    End Sub

    Private Sub ExerciseOptionBtn_Click(sender As Object, e As EventArgs) Handles ExerciseOptionBtn.Click
        myTransaction.clear()
        myTransaction.trType = ""
        myTransaction.typeOfPrice = "Strike"
        myTransaction.typeOfSecurity = "Option"
        If myTransaction.IsOptionInputValid() = True Then
            myTransaction.ComputeTransactionProperties()
            myTransaction.DisplayTransactionData()
        End If
    End Sub


    Private Sub ExecuteOptionTransactionBtn_Click(sender As Object, e As EventArgs) Handles ExecuteOptionTransactionBtn.Click
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
            MessageBox.Show("I cannot do this for you, Dave. Option input not valid.")
        End If
    End Sub

    Private Sub Trade0Btn_Click(sender As Object, e As EventArgs) Handles Trade0Btn.Click
        AlgoHedge(0)
    End Sub

    Private Sub Trade1Btn_Click(sender As Object, e As EventArgs) Handles Trade1Btn.Click
        AlgoHedge(1)
    End Sub

    Private Sub Trade2Btn_Click(sender As Object, e As EventArgs) Handles Trade2Btn.Click
        AlgoHedge(2)
    End Sub

    Private Sub Trade3Btn_Click(sender As Object, e As EventArgs) Handles Trade3Btn.Click
        AlgoHedge(3)
    End Sub

    Private Sub Trade4Btn_Click(sender As Object, e As EventArgs) Handles Trade4Btn.Click
        AlgoHedge(4)
    End Sub

    Private Sub Trade5Btn_Click(sender As Object, e As EventArgs) Handles Trade5Btn.Click
        AlgoHedge(5)
    End Sub

    Private Sub Trade6Btn_Click(sender As Object, e As EventArgs) Handles Trade6Btn.Click
        AlgoHedge(6)
    End Sub

    Private Sub Trade7Btn_Click(sender As Object, e As EventArgs) Handles Trade7Btn.Click
        AlgoHedge(7)
    End Sub

    Private Sub Trade8Btn_Click(sender As Object, e As EventArgs) Handles Trade8Btn.Click
        AlgoHedge(8)
    End Sub

    Private Sub Trade9Btn_Click(sender As Object, e As EventArgs) Handles Trade9Btn.Click
        AlgoHedge(9)
    End Sub

    Private Sub Trade10Btn_Click(sender As Object, e As EventArgs) Handles Trade10Btn.Click
        AlgoHedge(10)
    End Sub

    Private Sub Trade11Btn_Click(sender As Object, e As EventArgs) Handles Trade11Btn.Click
        AlgoHedge(11)
    End Sub


End Class
