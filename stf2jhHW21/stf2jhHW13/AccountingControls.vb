Module AccountingControls
    Public Function IsTransactionValid(ttype As String, sym As String, qty As Double) As Boolean
        If IsInIP(ttype) And (ttype <> "CashDiv" Or ttype <> "X-Put" Or ttype <> "X-Call") Then
            MessageBox.Show("Beep. Boop. Security in IP. Not sent.", "Accounting Controls", MessageBoxButtons.OK, MessageBoxIcon.Hand)
            Return False
        End If

        If ttype = "Sell" And (qty > GetCurrPositionInAP(sym)) Then
            MessageBox.Show("Beep. Boop. Selling more than you have. Not sent.", "Accounting Controls", MessageBoxButtons.OK, MessageBoxIcon.Hand)
            Return False
        End If

        If CAccountAT < 0 Then
            MessageBox.Show("Beep. Boop. Not money to do that. Not sent.", "Accounting Controls", MessageBoxButtons.OK, MessageBoxIcon.Hand)
            Return False
        End If

        If marginAT > maxMargins And (marginAT > margin) Then
            MessageBox.Show("Beep. Boop. No margin. Not Sent.", "Accounting Controls", MessageBoxButtons.OK, MessageBoxIcon.Hand)
            Return False
        End If

        If CAccountAT < (0.3 * margin) And (marginAT > margin) Then
            MessageBox.Show("Beep. Boop. Margin violation. Not Sent.", "Accounting Controls", MessageBoxButtons.OK, MessageBoxIcon.Hand)
            Return False
        End If

        Return True
    End Function

    Public Function IsAPEntryValid(sym As String, unitAsText As String) As Boolean
        If sym = "" Or unitAsText = "" Then
            Return False
        End If

        sym = sym.Trim()
        If Not IsNumeric(unitAsText) Then
            Return False
        End If

        If sym = "CAccount" Then
            Return True
        End If

        If Double.Parse(unitAsText) = 0 Then
            Return False
        End If
        If Not (IsAStock(sym) Or IsAnOption(sym)) Then
            MessageBox.Show("Beep.Boop.Unknown security (" + sym + ")", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End If
        Return True
    End Function
End Module
