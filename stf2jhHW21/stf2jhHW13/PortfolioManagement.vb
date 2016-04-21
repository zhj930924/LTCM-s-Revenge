Module PortfolioManagement

    Public Sub ResetAP()
        If MessageBox.Show("Beep. Boop. Are you sure?", "Reset AP?", MessageBoxButtons.YesNo, MessageBoxIcon.Hand) = DialogResult.Yes Then
            DownloadEnvironmentVariable()
            initialCAccount = GetInitialCAccount()
            ClearTeamPortfolioOnDb()
            UploadPosition("CAccount", initialCAccount)
            WarmStart()
        End If
    End Sub

    Public Sub UploadScreenPortfolioToDB()
        If Globals.ThisWorkbook.ActiveSheet.Name <> "Portfolio" Then
            MessageBox.Show("Beep. Boop. Are you looking at the Portfolio that you want me to upload?", "Portfolio Not Active", MessageBoxButtons.OK, MessageBoxIcon.Hand)
            Return
        End If

        Dim tempSymbol, tempUnits As String
        If Globals.Portfolio.AcquiredPositionsLO.IsSelected Then
            MessageBox.Show("Beep. Boop. Click outside the ListObject to confirm data entry.", "Edit In progress", MessageBoxButtons.OK, MessageBoxIcon.Hand)
            Return
        End If

        ClearTeamPortfolioOnDb()
        For i As Integer = 1 To Globals.Portfolio.AcquiredPositionsLO.DataBodyRange.Rows.Count()
            tempSymbol = Globals.Portfolio.AcquiredPositionsLO.DataBodyRange.Cells(i, 1).Value
            tempUnits = Globals.Portfolio.AcquiredPositionsLO.DataBodyRange.Cells(i, 2).Value
            If IsAPEntryValid(tempSymbol, tempUnits) Then
                UploadPosition(tempSymbol, tempUnits)
            End If
        Next
        Globals.Ribbons.SpartanTraderRibbon.AcquiredPositionsButton_Click(Nothing, Nothing)
        CAccount = GetCurrPositionInAP("CAccount")
        If StockPricesExist() Then
            CalcFinancialMetrics(currentDate)
            DisplayFinancialMetrics(currentDate)
        End If
    End Sub

    Public Function GetCurrPositionInAP(symbol) As Double
        For Each myRow As DataRow In myDataSet.Tables(portfolioTableName).Rows
            If myRow("Symbol").ToString().Trim() = symbol Then
                Return Double.Parse(myRow("Units"))
            End If
        Next
        Return 0
    End Function

    Public Function CalcTPVNoHedge(targetDate As Date) As Double
        Dim ts As TimeSpan = targetDate.Date - startDate.Date
        Dim t As Double = ts.Days / 365.25
        Dim interest As Double = initialCAccount * (Math.Exp(iRate * t) - 1)
        Return (CalcIPValue(targetDate) + initialCAccount + interest)
    End Function

    Public Function CalcTPV(targetdate As Date)
        Return (CalcIPValue(targetdate) + CalcAPValue(targetdate) + CAccount + CalcInterestSLT(targetdate))
    End Function

    Public Function CalcMargin(targetDate As Date) As Double
        Dim tempMargin As Double = 0
        Dim tempSymbol As String
        Dim tempUnits As Double
        If myDataSet.Tables.Contains("InitialPositionTable") Then
            For Each myRow As DataRow In myDataSet.Tables("InitialPositionTable").Rows
                tempSymbol = myRow("Symbol").ToString().Trim()
                tempUnits = myRow("Units")
                If tempUnits < 0 Then
                    tempMargin = tempMargin + (-tempUnits * CalcMTM(tempSymbol, targetDate))
                End If
            Next
        End If

        For Each myRow As DataRow In myDataSet.Tables(portfolioTableName).Rows
            tempSymbol = myRow("Symbol").ToString().Trim()
            tempUnits = myRow("Units")
            If (tempUnits < 0) And (tempSymbol <> "CAccount") Then
                tempMargin = tempMargin + (-tempUnits * CalcMTM(tempSymbol, targetDate))
            End If
        Next
        Return tempMargin

    End Function

    Public Function CalcAPValue(targetDate As Date) As Double
        Dim tempAP As Double = 0
        Dim tempSymbol As String
        Dim tempUnits As Double
        For Each myRow As DataRow In myDataSet.Tables(portfolioTableName).Rows
            tempSymbol = myRow("Symbol").ToString().Trim()
            tempUnits = myRow("Units")
            If tempSymbol <> "CAccount" Then
                tempAP = tempAP + (tempUnits * CalcMTM(tempSymbol, targetDate))
            End If
        Next
        Return tempAP
    End Function

    Public Function CalcInterestSLT(toThisDay As Date) As Double
        Dim interest As Double = 0
        Dim ts As TimeSpan = toThisDay.Date - lastTransactionDate.Date
        Dim t As Double = ts.Days / 365.25
        interest = CAccount * (Math.Exp(iRate * t) - 1)
        Return interest
    End Function

    Public Function CalcTaTPV(targetDate As Date) As Double
        Dim ts As TimeSpan = targetDate.Date - startDate.Date
        Dim t As Double = ts.Days / 365.25
        Return TPVatStart * Math.Exp(iRate * t)
    End Function

    Public Function CalcTPVatStart() As Double
        Return CalcIPValue(startDate) + initialCAccount
    End Function

    Public Function CalcIPValue(targetDate As Date) As Double
        Dim tempCumulativeValue As Double = 0
        Dim tempSymbol As String
        Dim tempUnits As Double

        If myDataSet.Tables.Contains("InitialPositionTable") Then
            For Each myRow As DataRow In myDataSet.Tables("InitialPositionTable").Rows
                tempSymbol = myRow("Symbol").ToString().Trim
                tempUnits = myRow("Units")
                tempCumulativeValue = tempCumulativeValue + (tempUnits * CalcMTM(tempSymbol, targetDate))
            Next
        End If

        Return tempCumulativeValue

    End Function

    Public Function CalcMTM(symbol As String, targetDate As Date) As Double
        Return (GetAsk(symbol, targetDate) + GetBid(symbol, targetDate)) / 2
    End Function

    Public Function IsAStock(Symbol As String) As Boolean
        Symbol = Symbol.Trim()
        For Each myRow As DataRow In myDataSet.Tables("TickerTable").Rows
            If myRow("Ticker").trim() = Symbol Then
                Return True
            End If
        Next
        Return False
    End Function

    Public Function IsAnOption(Symbol As String) As Boolean
        Symbol = Symbol.Trim()
        For Each myRow As DataRow In myDataSet.Tables("SymbolTable").Rows
            If myRow("Symbol").trim() = Symbol Then
                Return True
            End If
        Next
        Return False
    End Function

End Module
