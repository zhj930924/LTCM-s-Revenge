Public Class Recommendation

    Public posInArray As Integer = 0
    Public underlier As String = ""
    Public familyDelta As Double = 0
    Public familyGamma As Double = 0
    Public recSymbol As String = ""
    Public hedgeQty As Double = 0
    Public recTrType As String = ""
    Public recScore As Double = 0
    Public vol As Double = 0
    Public delta As Double = 0
    Public gamma As Double = 0
    Public maxShort As Double = 0
    Public maxBuy As Double = 0
    Public recCurrPos As Double = 0
    Public bestTrType As String = ""
    Public bestScore As Double = 0
    Public bestQty As Double = 0
    Public bestSymbol As String = ""
    Public DeltaGammaRatio As String 

    Public Sub FindBestHedge()
        If familyDelta > 0 Then
            'these base score weight are arbitrary. Improve them!
            CalcScoreForCashingDividend(0)
            CalcScoreForSellingStock(800)
            CalcScoreForSellingCall(700)
            CalcScoreForSellingShortCall(600)
            CalcScoreForBuyingBackPut(500)
            CalcScoreForBuyingPut(400)
            CalcScoreForSellingShortStock(200)
        Else  ' famdelta < 0
            CalcScoreForCashingDividend(0)
            CalcScoreForSellingPut(800)
            CalcScoreForBuyingBackCall(600)
            CalcScoreForBuyingBackStock(500)
            CalcScoreForBuyingStock(700)
            CalcScoreForBuyingCall(700)
            CalcScoreForSellingShortPut(300)
        End If
    End Sub

    Public Function NeedToHedge() As Boolean
        ' here you decide whether you need to hedge this family
        ' this is just an example with an arbitrary threshold  $5000 IS VERY TIGHT
        If Math.Abs(familyDelta) < 5000 Then
            If Math.Abs(familyGamma) < 5000 Then
                Return False
            Else
                Return True
            End If
            Return True
        End If
        Return True
    End Function

    Public Sub ResetRecommendation()
        bestTrType = "Hold"
        bestSymbol = "--"
        bestQty = 0
        bestScore = 0
        familyDelta = 0
        FamilyGamma = 0
    End Sub

    Public Sub DisplayRecommendation()
        Globals.Dashboard.FamilyDeltaRange.Cells(posInArray + 1, 1).Value = familyDelta
        Globals.Dashboard.RecommendationRange.Cells(posInArray + 1, 1).Value = bestTrType
        Globals.Dashboard.SymbolRange.Cells(posInArray + 1, 1).Value = bestSymbol
        Globals.Dashboard.QtyRange.Cells(posInArray + 1, 1).Value = bestQty
        Globals.Dashboard.FamilyGammaRange.Cells(posInArray + 1, 1).Value = familyGamma
        Globals.Dashboard.DeltaGammaRatioRange.Cells(posInArray + 1, 1).Value = DeltaGammaRatio
    End Sub

    'Public Function CalcQtyNeededToHedge(sym As String) As Integer
    '    Dim FamilyDeltaTarget As Double = 0
    '    Dim q As Double
    '    delta = CalcDelta(sym, currentDate)

    '    If Math.Abs(delta) < 0.05 Then
    '        '  arbitrary threshold!
    '        Return 0
    '    End If
    '    ' can change familydeltaTarget if you want to hedge to non-zero deltas
    '    q = (FamilyDeltaTarget - familyDelta) / delta
    '    Return Math.Abs(Math.Round(q))
    'End Function

    Public Function CalcQtyNeededToHedge(sym As String) As Integer
        Dim FamilyDeltaTarget As Double = 0
        Dim FamilyGammaTarget As Double = 0
        Dim q As Double

        delta = CalcDelta(sym, currentDate)
        gamma = CalcGamma(sym, currentDate)

        If Math.Abs(delta) < 0.05 Then
            If Math.Abs(gamma) < 0.05 Then
                '  arbitrary threshold!
                Return 0
            ElseIf Math.Abs(gamma) >= 0.05 Then
                q = (FamilyGammaTarget - familyGamma) / gamma
                Return Math.Abs(Math.Round(q))
            End If
        End If
        ' can change familydeltaTarget if you want to hedge to non-zero deltas
        q = (FamilyDeltaTarget - familyDelta) / delta
        Return Math.Abs(Math.Round(q))
    End Function

    Public Function TooCloseToMaxMargins() As Boolean
        If ((maxMargins - margin) < 500000) Then   ' Arbitrary threshold
            Return True
        Else
            Return False
        End If
    End Function

    Public Function MaxShortWithinConstraints(sym As String) As Double
        Dim q As Double = 0
        Dim maxAllowableIncreaseInMargins As Double = 0
        If TooCloseToMaxMargins() = True Then
            Return 0
        Else
            maxAllowableIncreaseInMargins = (maxMargins - margin) - 2000000 ' $2mil is an arbitrary cushion
            If maxAllowableIncreaseInMargins < 0 Then
                Return 0
            Else
                q = maxAllowableIncreaseInMargins / GetBid(sym, currentDate)
                Return Math.Truncate(q)
            End If
        End If
    End Function

    Public Function AvailableCashIsLow() As Boolean
        Dim availableCash As Double = CAccount - margin * 0.3
        If availableCash < 1000000 Then   ' Arbitrary threshold: $1mil
            Return True
        Else
            Return False
        End If
    End Function

    Public Function MaxPurchasePossible(sym As String) As Double
        Dim ask As Double = 0
        Dim q As Double = 0
        Dim availableCash As Double = (CAccount - margin * 0.3)
        availableCash = availableCash * 0.95 - 1000000  ' 5% is a cushion to pay for t-costs, $1mil is a buffer
        ask = GetAsk(sym, currentDate)
        If availableCash > 0 And ask > 0 Then
            q = availableCash / ask
            Return Math.Truncate(q)
        Else
            Return 0
        End If
    End Function

    ' ------- TACTIC SCORING --------------------------------------------------------------------------------------------------
    Public Function NumberInIP(underlier)
        If IsInIP(underlier) Then
            For Each myRow As DataRow In myDataSet.Tables("InitialPositionTable").Rows
                If myRow("Symbol").Trim() = underlier Then
                    Return myRow("Units")
                End If
            Next
        Else
            Return Nothing
        End If
        Return Nothing
    End Function

    Private Sub CalcScoreForCashingDividend(baseScore As Integer)
        Dim adjust As Integer = 0
        recCurrPos = GetCurrPositionInAP(underlier)

        For Each myRow As DataRow In myDataSet.Tables("StockMarketOneDayTable").Rows
            If myRow("DivDate").ToShortDateString = myRow("Date").ToShortDateString And myRow("Ticker").Trim() = underlier And myRow("Dividend") > 0 Then
                adjust = 100000
                bestTrType = "CashDiv"
                bestSymbol = underlier
                bestQty = NumberInIP(underlier)
                bestScore = (baseScore + adjust)
            End If
        Next

    End Sub


    Private Sub CalcScoreForSellingStock(baseScore As Integer)
        Dim adjust As Integer = 0
        If IsInIP(underlier) Then
            Exit Sub   ' cannot sell if in IP - no changes to best hedge
        End If

        recCurrPos = GetCurrPositionInAP(underlier)
        If recCurrPos <= 0 Then ' we cannot sell since we are not long
            Exit Sub
        End If

        hedgeQty = CalcQtyNeededToHedge(underlier)
        If hedgeQty = 0 Then
            Exit Sub  ' nothing to do
        End If

        If recCurrPos < hedgeQty Then ' you have fewer than needed
            hedgeQty = recCurrPos ' sell all you have
            adjust = -50
        End If

        If (baseScore + adjust) > bestScore Then
            bestTrType = "Sell"
            bestSymbol = underlier
            bestQty = hedgeQty
            bestScore = (baseScore + adjust)
        End If
    End Sub

    Private Sub CalcScoreForSellingCall(baseScore As Integer)
        Dim adjust As Integer = 0
        For Each dr As DataRow In myDataSet.Tables(portfolioTableName).Rows
            adjust = 0
            recSymbol = dr("Symbol").ToString().Trim()
            If IsAStock(recSymbol) Or recSymbol = "CAccount" Then
                'skip
            Else
                If (GetOptionType(recSymbol) = "Call") And (GetUnderlier(recSymbol) = underlier) Then
                    recCurrPos = dr("Units")
                    If recCurrPos > 0 Then
                        hedgeQty = CalcQtyNeededToHedge(recSymbol)
                        If hedgeQty > 0 Then
                            If recCurrPos < hedgeQty Then
                                hedgeQty = recCurrPos ' sell all you have
                                adjust = -50  'because incomplete hedge
                            End If
                            'If CAccount > (margin * 0.3) + 1000000 And TPV > TaTPV + 1000000 Then
                            '    adjust = -100
                            'End If '<- V5Change
                            If (baseScore + adjust) > bestScore Then
                                bestTrType = "Sell"
                                bestSymbol = recSymbol
                                bestQty = hedgeQty
                                bestScore = (baseScore + adjust)
                            End If
                        End If
                    End If
                End If
            End If
        Next
    End Sub

    Private Sub CalcScoreForSellingShortCall(baseScore As Integer)
        Dim adjust As Integer = 0
        If TooCloseToMaxMargins() Then
            Exit Sub ' no more credit
        End If
        ' only these options will be considered, in this order - you might add/subtract to/from the list
        For Each partialSymbol As String In {"_COCTA", "_COCTB", "_COCTC", "_COCTD", "_COCTE"}
            adjust = 0
            recSymbol = underlier + partialSymbol
            If Not IsInIP(recSymbol) Then
                recCurrPos = GetCurrPositionInAP(recSymbol)
                If recCurrPos <= 0 Then   ' if long cannot sell short
                    hedgeQty = CalcQtyNeededToHedge(recSymbol)
                    maxShort = MaxShortWithinConstraints(recSymbol)
                    If hedgeQty > maxShort Then
                        hedgeQty = maxShort
                        adjust = -50
                    End If
                    If CAccount > (margin * 0.3) + 1000000 Then
                        hedgeQty = maxBuy
                        adjust = -300
                    End If '<-V5Change
                    If hedgeQty > 0 And (baseScore + adjust) > bestScore Then
                        bestTrType = "SellShort"
                        bestSymbol = recSymbol
                        bestQty = hedgeQty
                        bestScore = (baseScore + adjust)
                    End If
                End If
            End If
        Next
    End Sub

    Public Sub CalcScoreForBuyingBackPut(baseScore As Integer)
        Dim adjust As Integer = 0
        If AvailableCashIsLow() Then
            Exit Sub
        End If
        For Each dr As DataRow In myDataSet.Tables(portfolioTableName).Rows
            adjust = 0
            recSymbol = dr("Symbol").ToString().Trim()
            If IsAStock(recSymbol) Or recSymbol = "CAccount" Then
                ' skip
            Else
                If GetOptionType(recSymbol) = "Put" And GetUnderlier(recSymbol) = underlier Then
                    recCurrPos = dr("Units")
                    If recCurrPos < 0 Then
                        hedgeQty = CalcQtyNeededToHedge(recSymbol)
                        If hedgeQty > 0 Then
                            If hedgeQty > Math.Abs(recCurrPos) Then
                                hedgeQty = Math.Abs(recCurrPos) ' buy back all that you have
                            End If
                            ' how much can you afford?
                            maxBuy = MaxPurchasePossible(recSymbol)
                            If maxBuy < hedgeQty Then
                                hedgeQty = maxBuy
                                adjust = -50
                            End If
                            If hedgeQty > 0 And (baseScore + adjust) > bestScore Then
                                bestTrType = "Buy"
                                bestSymbol = recSymbol
                                bestQty = hedgeQty
                                bestScore = (baseScore + adjust)
                            End If
                        End If
                    End If
                End If
            End If
        Next
    End Sub

    Public Sub CalcScoreForBuyingPut(baseScore As Integer)
        Dim adjust As Integer = 0
        If AvailableCashIsLow() Then
            Exit Sub
        End If
        ' arbitrarily only considers OCT options in this order - you can change that
        For Each partialSymbol As String In {"_POCTE", "_POCTD", "_POCTC", "_POCTB", "_POCTA"}
            adjust = 0
            recSymbol = underlier + partialSymbol
            If Not IsInIP(recSymbol) Then
                recCurrPos = GetCurrPositionInAP(recSymbol)
                If recCurrPos >= 0 Then ' if short it is a buyback
                    hedgeQty = CalcQtyNeededToHedge(recSymbol)
                    maxBuy = MaxPurchasePossible(recSymbol)  ' how much can you afford?
                    If maxBuy < hedgeQty Then
                        hedgeQty = maxBuy
                        adjust = -50

                    End If
                    If hedgeQty > 0 And (baseScore + adjust) > bestScore Then
                        bestTrType = "Buy"
                        bestSymbol = recSymbol
                        bestQty = hedgeQty
                        bestScore = (baseScore + adjust)
                    End If
                End If
            End If
        Next
    End Sub

    Public Sub CalcScoreForSellingShortStock(baseScore As Integer)
        Dim adjust As Integer = 0
        If TooCloseToMaxMargins() Then
            Exit Sub
        End If
        If Not IsInIP(underlier) Then
            recCurrPos = GetCurrPositionInAP(underlier)
            If recCurrPos <= 0 Then ' if long we cannot sell short
                hedgeQty = CalcQtyNeededToHedge(underlier)
                maxShort = MaxShortWithinConstraints(underlier)
                If hedgeQty > maxShort Then
                    hedgeQty = maxShort
                    adjust = -50
                End If
                If CAccount > (margin * 0.3) + 1000000 Then
                    hedgeQty = maxBuy
                    adjust = -300
                End If '<-V5Change
                If hedgeQty > 0 And (baseScore + adjust) > bestScore Then
                    bestTrType = "SellShort"
                    bestSymbol = underlier
                    bestQty = hedgeQty
                    bestScore = (baseScore + adjust)
                End If
            End If
        End If
    End Sub

    Public Sub CalcScoreForSellingPut(baseScore As Integer)
        Dim adjust As Integer = 0
        For Each dr As DataRow In myDataSet.Tables(portfolioTableName).Rows
            adjust = 0
            recSymbol = dr("Symbol").ToString().Trim()
            If IsAStock(recSymbol) Or recSymbol = "CAccount" Then
                ' skip
            Else
                If (GetOptionType(recSymbol) = "Put") And (GetUnderlier(recSymbol) = underlier) Then
                    recCurrPos = GetCurrPositionInAP(recSymbol)
                    If recCurrPos > 0 Then
                        hedgeQty = CalcQtyNeededToHedge(recSymbol)
                        'If CAccount > (margin * 0.3) + 1000000 And TPV > TaTPV + 1000000 Then
                        '    adjust = -100
                        'End If '<- V5Change
                        If recCurrPos < hedgeQty Then
                            hedgeQty = recCurrPos
                            adjust = -50
                            If hedgeQty > 0 And (baseScore + adjust) > bestScore Then
                                bestTrType = "Sell"
                                bestSymbol = recSymbol
                                bestQty = hedgeQty
                                bestScore = (baseScore + adjust)
                            End If
                        End If
                    End If
                End If
            End If
        Next
    End Sub

    Public Sub CalcScoreForSellingShortPut(baseScore As Integer)
        Dim adjust As Integer = 0
        If TooCloseToMaxMargins() Then
            Exit Sub
        End If
        ' arbitrary order, arbitrary exclusion of jul options
        For Each partialSymbol As String In {"_POCTE", "_POCTD", "_POCTC", "_POCTB", "_POCTA"}
            adjust = 0
            recSymbol = underlier + partialSymbol
            If Not IsInIP(recSymbol) Then
                recCurrPos = GetCurrPositionInAP(recSymbol)
                If recCurrPos <= 0 Then
                    hedgeQty = CalcQtyNeededToHedge(recSymbol)
                    maxShort = MaxShortWithinConstraints(recSymbol)
                    If maxShort < hedgeQty Then
                        hedgeQty = maxShort
                        adjust = -50
                    End If
                    If CAccount > (margin * 0.3) + 1000000 Then
                        hedgeQty = maxBuy
                        adjust = -100
                    End If '<-V5Change
                    If hedgeQty > 0 And (baseScore + adjust) > bestScore Then
                        bestTrType = "SellShort"
                        bestSymbol = recSymbol
                        bestQty = hedgeQty
                        bestScore = (baseScore + adjust)
                    End If
                End If
            End If
        Next
    End Sub

    Public Sub CalcScoreForBuyingBackCall(baseScore As Integer)
        Dim adjust As Integer = 0
        If AvailableCashIsLow() Then
            Exit Sub
        End If

        For Each dr As DataRow In myDataSet.Tables(portfolioTableName).Rows
            recSymbol = dr("Symbol").ToString().Trim()
            If IsAStock(recSymbol) Or recSymbol = "CAccount" Then
                ' skip
            Else
                If GetOptionType(recSymbol) = "Call" And GetUnderlier(recSymbol) = underlier Then
                    recCurrPos = dr("Units")
                    If recCurrPos < 0 Then
                        hedgeQty = CalcQtyNeededToHedge(recSymbol)
                        If Math.Abs(recCurrPos) < hedgeQty Then
                            hedgeQty = Math.Abs(recCurrPos) ' buy back all that you have
                            adjust = -50
                        End If
                        maxBuy = MaxPurchasePossible(recSymbol)
                        If maxBuy < hedgeQty Then
                            hedgeQty = maxBuy
                            adjust = -50
                        End If
                        If hedgeQty > 0 And (baseScore + adjust) > bestScore Then
                            bestTrType = "Buy"
                            bestSymbol = recSymbol
                            bestQty = hedgeQty
                            bestScore = (baseScore + adjust)
                        End If
                    End If
                End If
            End If
        Next
    End Sub

    Public Sub CalcScoreForBuyingBackStock(baseScore As Integer)
        Dim adjust As Integer = 0
        If AvailableCashIsLow() Then
            Exit Sub
        End If
        If Not IsInIP(underlier) Then
            recCurrPos = GetCurrPositionInAP(underlier)
            If recCurrPos < 0 Then
                hedgeQty = CalcQtyNeededToHedge(underlier)
                If Math.Abs(recCurrPos) < hedgeQty Then
                    hedgeQty = Math.Abs(recCurrPos) ' buy back all that you have
                    adjust = -50
                End If
                maxBuy = MaxPurchasePossible(underlier) ' how much can you afford?
                If maxBuy < hedgeQty Then
                    hedgeQty = maxBuy
                    adjust = -50
                End If
                If hedgeQty > 0 And (baseScore + adjust) > bestScore Then
                    bestTrType = "Buy"
                    bestSymbol = underlier
                    bestQty = hedgeQty
                    bestScore = (baseScore + adjust)
                End If
            End If
        End If
    End Sub

    Public Sub CalcScoreForBuyingCall(baseScore As Integer)
        Dim adjust As Integer = 0
        If AvailableCashIsLow() Then
            Exit Sub
        End If
        ' only considers OCT options you can change this
        For Each partialSymbol As String In {"_COCTA", "_COCTB", "_COCTC", "_COCTD", "_COCTE"}
            adjust = 0
            recSymbol = underlier + partialSymbol
            If Not IsInIP(recSymbol) Then
                recCurrPos = GetCurrPositionInAP(recSymbol)
                If recCurrPos >= 0 Then ' if short is a buyback
                    hedgeQty = CalcQtyNeededToHedge(recSymbol)
                    maxBuy = MaxPurchasePossible(recSymbol) ' how much can you afford?
                    If maxBuy < hedgeQty Then
                        hedgeQty = maxBuy
                        adjust = -50
                    End If
                    If hedgeQty > 0 And (baseScore + adjust) > bestScore Then
                        bestTrType = "Buy"
                        bestSymbol = recSymbol
                        bestQty = hedgeQty
                        bestScore = (baseScore + adjust)
                    End If
                End If
            End If
        Next
    End Sub

    Public Sub CalcScoreForBuyingStock(baseScore As Integer)
        Dim adjust As Integer = 0
        If AvailableCashIsLow() Then
            Exit Sub
        End If
        If IsInIP(underlier) Then
            Exit Sub
        End If
        recCurrPos = GetCurrPositionInAP(underlier)
        If recCurrPos < 0 Then     ' if short the we need a buyback
            Exit Sub
        End If
        hedgeQty = CalcQtyNeededToHedge(underlier)
        ' how much can you afford?
        maxBuy = MaxPurchasePossible(underlier)
        If maxBuy < hedgeQty Then
            hedgeQty = maxBuy
            adjust = -50
        End If
        If hedgeQty > 0 And (baseScore + adjust) > bestScore Then
            bestTrType = "Buy"
            bestSymbol = underlier
            bestQty = hedgeQty
            bestScore = baseScore + adjust
        End If
    End Sub



End Class
