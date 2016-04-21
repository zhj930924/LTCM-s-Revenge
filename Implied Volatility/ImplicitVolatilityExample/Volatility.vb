
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

End Class
