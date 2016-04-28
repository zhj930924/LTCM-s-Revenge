
Public Class Volatility

    Private Sub Sheet1_Startup() Handles Me.Startup
        
    End Sub

    Private Sub Sheet1_Shutdown() Handles Me.Shutdown

    End Sub

    Private Sub GoalSeekBtn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GoalSeekBtn.Click
        ' init volatility to an arbitrary value
        Range("G28").Value = 0.25
        ' set all cells so that they copy that initial value
        Range("K3:K22").Formula = "=$G$28"
        ' Goalseek tries to bring the target to 0.05 (choosing = 0 risks making it fail to
        ' converge - you can choose other thresholds)
        Range("R26").GoalSeek(0.05, Range("G28"))
    End Sub

    Private Sub SolverBtn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SolverBtn.Click
        ' if VSTO can't find Solver.xlam, just copy and paste Solver.xlam and Solver.dll
        ' "from C:\Program Files\Microsoft Office\Office14\Library\SOLVER" or (wherever it is) into "My Documents".
        Application.Run("Solver.xlam!SolverReset")
        Application.Run("Solver.xlam!SolverOk", "$R$26", 2, "", "K3:K22")
        ' the third parameter means "find the Min", 1 is "find the Max", 3 means "make it equal to" the fouth parameter.
        Application.Run("Solver.xlam!SolverSolve", "True")
    End Sub

    Public Sub LoopingThroughTable()
        ' 22 To 10002
        ' Alpha Options: T{0}:AA{1}, Stock: AC{0}:AH{1}
        ' Beta Options: AJ{0}:AQ{1}, Stock: AS{0}:AX{1}
        ' Gamma Options: AZ{0}:BG{1}, Stock: BI{0}:BN{1}
        For x = 1 To 41
            ' Stocks
            Table3.DataBodyRange.Value = Range(String.Format("BI{0}:BN{1}", x + 2, x + 2)).Value
            ' Options
            Table2.DataBodyRange.Value = Range(String.Format("AZ{0}:BG{1}", 20 * x - 17, 20 * x + 2)).Value

            ' Time
            Range(String.Format("B{0}", x + 32)).Value = Range("B3").Value

            ' Ticker
            Range(String.Format("C{0}", x + 32)).Value = Range("I3").Value

            ' GoalSeek Average
            ' GoalSeekBtn_Click(Nothing, Nothing)
            ' Range(String.Format("D{0}", x + 32)).Value = Range("K26").Value

            ' Solver Average
            SolverBtn_Click(Nothing, Nothing)
            Range(String.Format("E{0}", x + 32)).Value = Range("K26").Value
        Next
    End Sub

    Public Sub StartBtn_Click(sender As Object, e As EventArgs) Handles StartBtn.Click

        LoopingThroughTable()

    End Sub

    Public Sub MyDelayMacro()

        For iCount = 1 To 100000
        Next iCount

    End Sub

End Class
