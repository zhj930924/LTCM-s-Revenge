Module AccountingControls

    Public Function IsTransactionValid(ttype As String, sym As String, qty As Double) As Boolean
        If IsInIP(ttype) And (ttype <> "CashDiv" Or ttype <> "X-Put" Or ttype <> "X-Call") Then
            MessageBox.Show("Holy BatSmoke! Security in IP. Not sent.", "Accounting Controls", MessageBoxButtons.OK, MessageBoxIcon.Hand)
            Return False
        End If

        If ttype = "Sell" And (qty > GetCurrPositionInAP(sym)) Then
            MessageBox.Show("Holy BatSmoke! Selling more than you have. Not Sent.", "Accounting Controls", MessageBoxButtons.OK, MessageBoxIcon.Hand)
            Return False
        End If

        If CAccountAT < 0 Then
            MessageBox.Show("Holy BatSmoke! No money to do that. Not sent.", "Accounting Controls", MessageBoxButtons.OK, MessageBoxIcon.Hand)
            Return False
        End If

        If marginAT > maxMargins And (marginAT > margin) Then
            MessageBox.Show("Holy BatSmoke! No margin. Not sent.", "Accounting Controls", MessageBoxButtons.OK, MessageBoxIcon.Hand)
            Return False
        End If

        If CAccount < (0.3 * margin) And (marginAT > margin) Then
            MessageBox.Show("Holy BatSmoke! Margin violation. Not sent.", "Accounting Controls", MessageBoxButtons.OK, MessageBoxIcon.Hand)
            Return False
        End If

        Return True ' if all controls are passed

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
            MessageBox.Show("Holy batpencil! I am afraid I cannot process this, Dave. Unknown security (" + sym + ")", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End If
        Return True
    End Function
End Module
