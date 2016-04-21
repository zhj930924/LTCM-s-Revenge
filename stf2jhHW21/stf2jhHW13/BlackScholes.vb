Module BlackScholes
    Public Function CalcFamilyDelta(tkr As String) As Double
        Dim tempFamDelta As Double = 0
        Dim delta As Double = 0
        Dim sym As String
        tkr = tkr.Trim()

        For Each dr As DataRow In myDataSet.Tables(portfolioTableName).Rows
            sym = dr("Symbol").ToString().Trim()
            If IsIntheFamily(sym, tkr) Then
                delta = CalcDelta(sym, currentDate)
                tempFamDelta = tempFamDelta + delta * dr("Units")
            End If
        Next

        If myDataSet.Tables.Contains("InitialPositionTable") Then
            For Each dr As DataRow In myDataSet.Tables("InitialPositionTable").Rows
                sym = dr("Symbol").ToString().Trim()
                If IsIntheFamily(sym, tkr) Then
                    delta = CalcDelta(sym, currentDate)
                    tempFamDelta = tempFamDelta + delta * dr("Units")
                End If
            Next
        End If
        Return tempFamDelta

    End Function

    Public Function IsIntheFamily(sym As String, familyTicker As String) As Boolean
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

    Public Function CalcDelta(Symbol As String, targetDate As Date) As Double
        Dim sigma As Double
        Dim K As Double
        Dim S As Double
        Dim r As Double = iRate
        Dim t As Double
        Dim ts As TimeSpan
        Dim underlier As String
        Dim d1 As Double

        'Checks
        If Symbol = "CAccount" Then
            Return 0
        End If
        If IsAStock(Symbol) Then
            Return 1
        End If
        If targetDate.Date >= GetExpiration(Symbol).Date Then
            Return 0
        End If
        If GetAsk(Symbol, targetDate) = 0 Then
            Return 0
        End If

        'Data collection
        underlier = GetUnderlier(Symbol)
        sigma = GetVol(underlier)
        K = GetStrike(Symbol)
        S = CalcMTM(underlier, targetDate)
        ts = GetExpiration(Symbol).Date - targetDate.Date
        t = ts.Days / 365.25
        'BS Formula
        d1 = (Math.Log(S / K) + (r + sigma * sigma / 2) * t) / (sigma * Math.Sqrt(t))
        If GetOptionType(Symbol).Trim() = "Call" Then
            Return Globals.ThisWorkbook.Application.WorksheetFunction.Norm_S_Dist(d1, True)
        End If
        If GetOptionType(Symbol).Trim() = "Put" Then
            Return (Globals.ThisWorkbook.Application.WorksheetFunction.Norm_S_Dist(d1, True) - 1)
        End If
        Return 0

    End Function

    Public Function GetVol(symbol As String) As Double
        Select Case symbol
            Case "AAPL" '<- TEST NUMBERS
                Return 0.281
            Case "COP"
                Return 0.25
            Case "BABA"
                Return 0.25
            Case "COST"
                Return 0.25
            Case "BLK"
                Return 0.25
            Case "DB"
                Return 0.25
            Case "WMT"
                Return 0.25
            Case "LNKD"
                Return 0.25
            Case "XOM"
                Return 0.25
            Case "HSY"
                Return 0.25
            Case "FIT"
                Return 0.25
            Case "NKE"
                Return 0.25
            Case Else
                Return 0.25
        End Select
    End Function
End Module
