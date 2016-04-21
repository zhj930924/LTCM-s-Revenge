Class Transaction
    Public underlier As String
    Public price As Double = 0
    Public trType As String = ""
    Public symbol As String = ""
    Public typeOfSecurity As String = ""
    Public qty As Double = 0
    Public transCost As Double = 0
    Public totValue As Double = 0
    Public typeOfPrice As String = ""
    Public optionType As String = ""
    Public delta As Double = 0
    Public strike As Double = 0

    Public Function IsOptionInputValid() As Boolean
        If Globals.Dashboard.SymbolCBox.SelectedItem = Nothing Then
            MessageBox.Show("Beep. Boop. No Symbol", "No symbol", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        Else
            symbol = Globals.Dashboard.SymbolCBox.SelectedItem.Trim()
            optionType = GetOptionType(symbol)
        End If

        If trType = "" Then
            If optionType = "Put" Then
                trType = "X-Put"
            Else
                trType = "X-Call"
            End If
        End If

        Try
            qty = Integer.Parse(Globals.Dashboard.OptionQtyTBox.Text)
        Catch
            MessageBox.Show("Beep. Boop. No quantity", "No quantity", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End Try
        Return True
    End Function

    Public Sub UpdateAP()
        Select Case trType
            Case "Buy"
                UploadPosition(symbol, GetCurrPositionInAP(symbol) + qty)
            Case "Sell"
                UploadPosition(symbol, GetCurrPositionInAP(symbol) - qty)
            Case "SellShort"
                UploadPosition(symbol, GetCurrPositionInAP(symbol) - qty)
            Case "CashDiv"
                'Does nada
            Case "X-Put"
                Dim UnderlierPosition As Double = GetCurrPositionInAP(GetUnderlier(symbol))
                If GetCurrPositionInAP(symbol) > 0 Then
                    UploadPosition(symbol, GetCurrPositionInAP(symbol) - qty)
                    UploadPosition(GetUnderlier(symbol), UnderlierPosition - qty)
                Else
                    UploadPosition(symbol, GetCurrPositionInAP(symbol) + qty)
                    UploadPosition(GetUnderlier(symbol), UnderlierPosition + qty)
                End If
            Case "X-Call"
                Dim UnderlierPosition As Double = GetCurrPositionInAP(GetUnderlier(symbol))
                If GetCurrPositionInAP(symbol) > 0 Then
                    UploadPosition(symbol, GetCurrPositionInAP(symbol) - qty)
                    UploadPosition(GetUnderlier(symbol), UnderlierPosition + qty)
                Else
                    UploadPosition(symbol, GetCurrPositionInAP(symbol) + qty)
                    UploadPosition(GetUnderlier(symbol), UnderlierPosition - qty)
                End If
        End Select
        UploadPosition("CAccount", CAccountAT)
        DownloadAcquiredPositions()
    End Sub

    Public Sub Clear()
        price = 0
        trType = ""
        symbol = ""
        qty = 0
        transCost = 0
        totValue = 0
        typeOfPrice = ""
        optionType = ""
        delta = 0
        strike = 0
        Globals.Dashboard.ClearTransactionHighlight()
    End Sub

    Public Sub ExecuteTransaction()
        Dim mySQL As String
        mySQL = String.Format("INSERT INTO TransactionQueue (Date, TeamID, Symbol, Type, Qty, Price, Cost, TotValue, " + "InterestSinceLastTransaction, CashPositionAfterTransaction, TotMargin) VALUES " + "('{0}', '{1}', '{2}', '{3}', '{4}', '{5}', '{6}', '{7}', '{8}', '{9}', '{10}')", currentDate.ToShortDateString, teamID, symbol, trType, qty, price, transCost, totValue, interestSLT, CAccountAt, marginAT)
        ExecuteNonQuery(mySQL)
        lastTransactionDate = currentDate
        CAccount = CAccountAT
        margin = marginAT
        UpdateAP()
    End Sub

    Public Function EffectofTransactiononMargin() As Double
        Dim currPosition As Integer = 0
        Dim underlierPosition As Integer = 0
        Dim effect As Double = 0
        Dim underlier As String '<_ !!!!!!TESTING

        Select Case trType
            Case "Sell"
                Return 0
            Case "Buy"
                currPosition = GetCurrPositionInAP(symbol)
                If currPosition >= 0 Then
                    Return 0
                Else
                    If qty >= Math.Abs(currPosition) Then
                        Return currPosition * CalcMTM(symbol, currentDate)
                    Else
                        Return -(qty * CalcMTM(symbol, currentDate))
                    End If
                End If
            Case "SellShort"
                Return qty * CalcMTM(symbol, currentDate)
            Case "CashDiv"
                Return 0
            Case "X-Call"
                Dim OptionEffect As Double = 0
                currPosition = GetCurrPositionInAP(symbol)
                underlier = GetUnderlier(symbol)
                underlierPosition = GetCurrPositionInAP(underlier)

                If currPosition < 0 Then
                    OptionEffect = -qty * CalcMTM(symbol, currentDate)

                Else
                    OptionEffect = 0
                End If

                If currPosition >= 0 Then
                    If underlierPosition >= 0 Then
                        Return OptionEffect
                    Else
                        If qty >= Math.Abs(underlierPosition) Then
                            Return OptionEffect + (underlierPosition * CalcMTM(underlier, currentDate))
                        Else
                            Return OptionEffect - (qty * CalcMTM(underlier, currentDate))
                        End If
                    End If

                Else
                    If underlierPosition <= 0 Then
                        Return OptionEffect - (qty * CalcMTM(underlier, currentDate))
                    Else
                        If underlierPosition >= qty Then
                            Return OptionEffect
                        Else
                            Return OptionEffect + ((qty - underlierPosition) * CalcMTM(underlier, currentDate))
                        End If
                    End If
                End If

            Case "X-Put"
                Dim OptionEffect As Double = 0
                currPosition = GetCurrPositionInAP(symbol)
                underlier = GetUnderlier(symbol)
                underlierPosition = GetCurrPositionInAP(underlier)

                If currPosition < 0 Then
                    OptionEffect = -qty * CalcMTM(symbol, currentDate)

                Else
                    OptionEffect = 0
                End If

                If currPosition < 0 Then
                    If underlierPosition >= 0 Then
                        Return OptionEffect
                    Else
                        If qty >= Math.Abs(underlierPosition) Then
                            Return OptionEffect + (underlierPosition * CalcMTM(underlier, currentDate))
                        Else
                            Return OptionEffect - (qty * CalcMTM(underlier, currentDate))
                        End If
                    End If

                Else
                    If underlierPosition <= 0 Then
                        Return OptionEffect - (qty * CalcMTM(underlier, currentDate))
                    Else
                        If underlierPosition >= qty Then
                            Return OptionEffect
                        Else
                            Return OptionEffect + ((qty - underlierPosition) * CalcMTM(underlier, currentDate))
                        End If
                    End If
                End If

        End Select
        MessageBox.Show("Beep. Boop. Couldn'r figure out impact of " + symbol + " on margin. Returned $0.")
        Return 0
    End Function

    Public Function IsStockInputValid() As Boolean
        If Globals.Dashboard.TickerCBox.SelectedItem = Nothing Then
            MessageBox.Show("Beep. Boop. No Ticker", "No ticker", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        Else
            symbol = Globals.Dashboard.TickerCBox.SelectedItem
        End If

        If trType = "" Then
            MessageBox.Show("Beep. Boop. No transaction type", "No transaction type", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End If

        Try
            qty = Integer.Parse(Globals.Dashboard.StockQtyTBox.Text)
        Catch
            MessageBox.Show("Beep. Boop. No quantity", "No quantity", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End Try
        Return True
    End Function

    Public Sub ComputeTransactionProperties()
        If IsAStock(symbol) Then
            typeOfSecurity = "Stock"
        Else
            typeOfSecurity = "Option"
            strike = GetStrike(symbol)
        End If
        Select Case trType
            Case "Buy"
                typeOfPrice = "Ask"
            Case "Sell"
                typeOfPrice = "Bid"
            Case "SellShort"
                typeOfPrice = "Bid"
            Case "CashDiv"
                typeOfPrice = "Div"
            Case "X-Call"
                typeOfPrice = "Strike"
            Case "X-Put"
                typeOfPrice = "Strike"
        End Select
        Select Case typeOfPrice
            Case "Bid"
                price = GetBid(symbol, currentDate)
            Case "Ask"
                price = GetAsk(symbol, currentDate)
            Case "Div"
                price = GetDividend(symbol, currentDate)
            Case "Strike"
                price = strike
            Case Else
                price = 0
        End Select

        delta = CalcDelta(symbol, currentDate)
        transCost = CalcTransCost()
        totValue = CalcTotValue()
        interestSLT = CalcInterestSLT(currentDate)
        CAccountAT = CAccount + totValue + interestSLT
        marginAT = margin + EffectofTransactiononMargin()
    End Sub

    Public Sub DisplayTransactionData()
        Try
            Globals.Dashboard.PriceCell.Value = price
            Globals.Dashboard.TypeCell.Value = trType
            Globals.Dashboard.SymbolCell.Value = symbol
            Globals.Dashboard.QtyCell.Value = qty
            Globals.Dashboard.TransCostCell.Value = transCost
            Globals.Dashboard.TotValueCell.Value = totValue
            Globals.Dashboard.InterestSLTCell.Value = interestSLT
            Globals.Dashboard.DeltaCell.Value = delta
            Globals.Dashboard.StrikeCell.Value = strike
            Globals.Dashboard.CAccountATCell.Value = CAccountAT
            Globals.Dashboard.MarginATCell.Value = marginAT
        Catch

        End Try
    End Sub

    Public Function CalcTransCost() As Double
        Return GetTCostCoefficient(symbol, trType) * Math.Abs(qty) * price
    End Function

    Public Function CalcTotValue() As Double
        Select Case trType
            Case "Buy"
                Return -(price * qty) - transCost
            Case "Sell"
                Return (price * qty) - transCost
            Case "SellShort"
                Return (price * qty) - transCost
            Case "CashDiv"
                Return (price * qty) - transCost
            Case "X-Put"
                Return (price * qty) - transCost
            Case "X-Call"
                Return -(price * qty) - transCost
            Case Else
                Return 0
        End Select
    End Function

End Class

