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

    Public Function CalcGamma(symbol As String, targetDate As Date) As Double
        Dim sigma As Double
        Dim K As Double    'strike
        Dim S As Double    'spot price to the stock
        Dim r As Double = iRate
        Dim t As Double    'time to expiration
        Dim ts As TimeSpan
        Dim underlier As String
        Dim d1 As Double
        Dim Gamma As Double
        'checks
        If symbol = "CAccount" Then
            Return 0
        End If
        If IsAStock(symbol) Then
            Return 0
        End If   'from now on only options
        If targetDate.Date >= GetExpiration(symbol).Date Then
            Return 0
        End If
        'If GetAsk(symbol, targetDate) = 0 Then
        '    Return 0
        'End If
        'data collection
        underlier = GetUnderlier(symbol)
        sigma = GetVol(underlier)
        K = GetStrike(symbol)
        S = CalcMTM(underlier, targetDate)
        ts = GetExpiration(symbol).Date - targetDate.Date
        t = ts.Days / 365.25
        'BS formula
        d1 = (Math.Log(S / K) + (r + sigma * sigma / 2) * t) / (sigma * Math.Sqrt(t))

        Gamma = (1 / Math.Sqrt(2 * Math.PI)) * Math.Exp(-0.5 * d1 * d1) * (1 / (S * sigma * Math.Sqrt(t)))
        'MessageBox.Show("Beep. Boop. Calculated " + Gamma + " as the Gamma for " + underlier + ".", "Success?", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Return Gamma

    End Function

    Public Function CalcFamilyGamma(tkr As String) As Double
        Dim tempFamGamma As Double = 0
        Dim Gamma As Double = 0
        Dim sym As String
        tkr = tkr.Trim()
        'AP
        For Each dr As DataRow In myDataSet.Tables(portfolioTableName).Rows
            sym = dr("Symbol").ToString().Trim()
            If IsInTheFamily(sym, tkr) Then
                Gamma = CalcGamma(sym, currentDate)
                tempFamGamma = tempFamGamma + Gamma * dr("Units")
            End If
        Next
        'IP
        If myDataSet.Tables.Contains("InitialPositionTable") Then ' else skip
            For Each dr As DataRow In myDataSet.Tables("InitialPositionTable").Rows
                sym = dr("Symbol").ToString().Trim()
                If IsInTheFamily(sym, tkr) Then
                    Gamma = CalcGamma(sym, currentDate)
                    tempFamGamma = tempFamGamma + Gamma * dr("Units")
                End If
            Next
        End If
        Return tempFamGamma
    End Function

    ' THESE vols are not accurate

    Public Function GetVol(symbol As String) As Double
        Select Case symbol
            Case "AAPL"
                Return 0.205683
            Case "BABA"
                Return 0.345512
            Case "BLK"
                Return 0.301195
            Case "COP"
                Return 0.5281
            Case "COST"
                Return 0.186249
            Case "DB"
                Return 0.547739
            Case "FIT"
                Return 0.883353
            Case "HSY"
                Return 0.152654
            Case "LNKD"
                Return 1.142
            Case "NKE"
                Return 0.308082
            Case "WMT"
                Return 0.15922
            Case "XOM"
                Return 0.204368
            Case Else
                Return 0.25
        End Select

    End Function


End Module
