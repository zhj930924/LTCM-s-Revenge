Module BlackScholes

    Public Function CalcFamilyDelta(tkr As String) As Double
        Dim tempFamDelta As Double = 0
        Dim delta As Double = 0
        Dim sym As String
        tkr = tkr.Trim()
        'AP
        For Each dr As DataRow In myDataSet.Tables(portfolioTableName).Rows
            sym = dr("Symbol").ToString().Trim()
            If IsInTheFamily(sym, tkr) Then
                delta = CalcDelta(sym, currentDate)
                tempFamDelta = tempFamDelta + delta * dr("Units")
            End If
        Next
        'IP
        If myDataSet.Tables.Contains("InitialPositionTable") Then ' else skip
            For Each dr As DataRow In myDataSet.Tables("InitialPositionTable").Rows
                sym = dr("Symbol").ToString().Trim()
                If IsInTheFamily(sym, tkr) Then
                    delta = CalcDelta(sym, currentDate)
                    tempFamDelta = tempFamDelta + delta * dr("Units")
                End If
            Next
        End If
        Return tempFamDelta
    End Function

    Public Function IsInTheFamily(sym As String, familyTicker As String) As Boolean
        If sym = "CAccount" Then
            Return False
        End If
        If IsAStock(sym) Then
            If sym = familyTicker Then
                Return True
            Else
                Return False
            End If
        Else
            If GetUnderlier(sym) = familyTicker Then
                Return True
            Else
                Return False
            End If
        End If
    End Function

    Public Function CalcDelta(symbol As String, targetDate As Date) As Double
        Dim sigma As Double
        Dim K As Double    'strike
        Dim S As Double    'spot price to the stock
        Dim r As Double = iRate
        Dim t As Double    'time to expiration
        Dim ts As TimeSpan
        Dim underlier As String
        Dim d1 As Double
        'checks
        If symbol = "CAccount" Then
            Return 0
        End If
        If IsAStock(symbol) Then
            Return 1
        End If   'from now on only options
        If targetDate.Date >= GetExpiration(symbol).Date Then
            Return 0
        End If
        If GetAsk(symbol, targetDate) = 0 Then
            Return 0
        End If
        'data collection
        underlier = GetUnderlier(symbol)
        sigma = GetVol(underlier)
        K = GetStrike(symbol)
        S = CalcMTM(underlier, targetDate)
        ts = GetExpiration(symbol).Date - targetDate.Date
        t = ts.Days / 365.25
        'BS formula
        d1 = (Math.Log(S / K) + (r + sigma * sigma / 2) * t) / (sigma * Math.Sqrt(t))
        If GetOptionType(symbol).Trim() = "Call" Then
            Return Globals.ThisWorkbook.Application.WorksheetFunction.Norm_S_Dist(d1, True)
        End If
        If GetOptionType(symbol).Trim() = "Put" Then
            Return (Globals.ThisWorkbook.Application.WorksheetFunction.Norm_S_Dist(d1, True) - 1)
        End If
        Return 0
    End Function


    ' THESE vols are not accurate

    Public Function GetVol(symbol As String) As Double
        Select Case symbol
            Case "AAPL"
                Return 0.2
            Case "AMZN"
                Return 0.25
            Case "BABA"
                Return 0.35
            Case "BAC"
                Return 0.15
            Case "CMG"
                Return 0.2
            Case "DIS"
                Return 0.1
            Case "GOOG"
                Return 0.2
            Case "KMB"
                Return 0.25
            Case "SHLD"
                Return 0.3
            Case "SNE"
                Return 0.2
            Case "UA"
                Return 0.15
            Case "VFC"
                Return 0.25
            Case Else
                Return 0.25
        End Select

    End Function


End Module
