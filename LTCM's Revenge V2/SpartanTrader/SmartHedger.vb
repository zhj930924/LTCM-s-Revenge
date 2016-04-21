Module SmartHedger

    Public Sub RecommendHedges()
        ResetRecommendations()
        CalcFamilyDeltas()
        DisplayRecommendations()
        If HedgingToday() = False Then  ' first: are hedging today?
            Exit Sub
        End If
        For i As Integer = 0 To 11
            If RecArray(i).NeedToHedge() = True Then  ' second, do we want to hedge this ticker?
                RecArray(i).FindBestHedge()
                RecArray(i).DisplayRecommendation()
            End If
        Next
    End Sub

    Public Sub AlgoHedgeAll()
        For i As Integer = 0 To 11
            AlgoHedge(i)
        Next
    End Sub

    Public Sub AlgoHedge(i As Integer)
        Dim tType As String = ""
        Dim tQty As Double = 0
        Dim tSymbol As String = ""
        tType = RecArray(i).bestTrType
        tQty = RecArray(i).bestQty
        tSymbol = RecArray(i).bestSymbol
        If tType <> "Hold" And tQty <> 0 Then
            ExecuteAlgoTransaction(tType, tQty, tSymbol)
            CalcFinancialMetrics(currentDate)
            DisplayFinancialMetrics(currentDate)
        End If
    End Sub

    Public Sub ExecuteAlgoTransaction(tType As String, tQty As Integer, tSymbol As String)
        Dim AlgoTr As Transaction = New Transaction()
        AlgoTr.trType = tType
        AlgoTr.qty = tQty
        AlgoTr.symbol = tSymbol
        AlgoTr.ComputeTransactionProperties()
        If IsTransactionValid(AlgoTr.trType, AlgoTr.symbol, AlgoTr.qty) Then
            AlgoTr.ExecuteTransaction()
            Globals.Dashboard.PrintToAlgoLog(String.Format("{0}: {1} {2} {3} {4:C0} Delta: {5:N3}",
                currentDate.ToShortDateString(),
                AlgoTr.trType, AlgoTr.qty, AlgoTr.symbol, AlgoTr.totValue, AlgoTr.delta))
        End If
    End Sub

    Public Function HedgingToday() As Boolean
        'here you decide whether you want to hedge today - can add more conditions!
        If currentDate.DayOfWeek = DayOfWeek.Saturday Or
             currentDate.DayOfWeek = DayOfWeek.Sunday Then
            Return False
        End If
        Return True
    End Function

    Public Sub ResetRecommendations()
        If Not IsNothing(RecArray(0)) Then
            For i As Integer = 0 To 11
                RecArray(i).ResetRecommendation()
            Next
        End If
    End Sub

    Public Sub DisplayRecommendations()
        If Not IsNothing(RecArray(0)) Then
            For i As Integer = 0 To 11
                RecArray(i).DisplayRecommendation()
            Next
        End If
    End Sub

    Public Sub CalcFamilyDeltas()
        For i As Integer = 0 To 11
            RecArray(i).familyDelta = CalcFamilyDelta(RecArray(i).underlier)
        Next
    End Sub

    Public Sub SetUpRecommendations()
        Dim tempSym As String = ""
        Dim tempVol As Double = 0

        For i As Integer = 0 To 11
            RecArray(i) = New Recommendation()

            RecArray(i).posInArray = i
            tempSym = myDataSet.Tables("TickerTable").Rows(i)("Ticker")
            tempSym = tempSym.Trim()
            RecArray(i).underlier = tempSym
            Globals.Dashboard.UnderlierRange.Cells(i + 1, 1).Value = "[" + i.ToString() + "]   " + tempSym

            RecArray(i).vol = GetVol(tempSym)  ' you need to set your own vols
            Globals.Dashboard.VolatilityRange.Cells(i + 1, 1).Value = RecArray(i).vol
        Next
    End Sub


End Module