Partial Class stRibbon
    Inherits Microsoft.Office.Tools.Ribbon.RibbonBase

    <System.Diagnostics.DebuggerNonUserCode()> _
    Public Sub New(ByVal container As System.ComponentModel.IContainer)
        MyClass.New()

        'Required for Windows.Forms Class Composition Designer support
        If (container IsNot Nothing) Then
            container.Add(Me)
        End If

    End Sub

    <System.Diagnostics.DebuggerNonUserCode()> _
    Public Sub New()
        MyBase.New(Globals.Factory.GetRibbonFactory())

        'This call is required by the Component Designer.
        InitializeComponent()

    End Sub

    'Component overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Component Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Component Designer
    'It can be modified using the Component Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(stRibbon))
        Me.Tab1 = Me.Factory.CreateRibbonTab
        Me.Group1 = Me.Factory.CreateRibbonGroup
        Me.AlphaButton = Me.Factory.CreateRibbonToggleButton
        Me.BetaButton = Me.Factory.CreateRibbonToggleButton
        Me.GammaButton = Me.Factory.CreateRibbonToggleButton
        Me.Control = Me.Factory.CreateRibbonGroup
        Me.ManualButton = Me.Factory.CreateRibbonToggleButton
        Me.SynchButton = Me.Factory.CreateRibbonToggleButton
        Me.SimulationButton = Me.Factory.CreateRibbonToggleButton
        Me.AutoHedgeButton = Me.Factory.CreateRibbonToggleButton
        Me.Group2 = Me.Factory.CreateRibbonGroup
        Me.DashboardButton = Me.Factory.CreateRibbonButton
        Me.FinChartsButton = Me.Factory.CreateRibbonButton
        Me.Group3 = Me.Factory.CreateRibbonGroup
        Me.InitialPositionsButton = Me.Factory.CreateRibbonButton
        Me.AcquiredPositionsButton = Me.Factory.CreateRibbonButton
        Me.TransactionQButton = Me.Factory.CreateRibbonButton
        Me.ResetAPButton = Me.Factory.CreateRibbonButton
        Me.EditAPButton = Me.Factory.CreateRibbonButton
        Me.UploadAPButton = Me.Factory.CreateRibbonButton
        Me.ConfirmationButton = Me.Factory.CreateRibbonButton
        Me.Group4 = Me.Factory.CreateRibbonGroup
        Me.StockMktButton = Me.Factory.CreateRibbonButton
        Me.OptionMktButton = Me.Factory.CreateRibbonButton
        Me.SP500Button = Me.Factory.CreateRibbonButton
        Me.SettingsButton = Me.Factory.CreateRibbonButton
        Me.TransactionButton = Me.Factory.CreateRibbonButton
        Me.Group5 = Me.Factory.CreateRibbonGroup
        Me.QuitButton = Me.Factory.CreateRibbonButton
        Me.ExcludeIPButton = Me.Factory.CreateRibbonToggleButton
        Me.Tab1.SuspendLayout()
        Me.Group1.SuspendLayout()
        Me.Control.SuspendLayout()
        Me.Group2.SuspendLayout()
        Me.Group3.SuspendLayout()
        Me.Group4.SuspendLayout()
        Me.Group5.SuspendLayout()
        Me.SuspendLayout()
        '
        'Tab1
        '
        Me.Tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office
        Me.Tab1.Groups.Add(Me.Group1)
        Me.Tab1.Groups.Add(Me.Control)
        Me.Tab1.Groups.Add(Me.Group2)
        Me.Tab1.Groups.Add(Me.Group3)
        Me.Tab1.Groups.Add(Me.Group4)
        Me.Tab1.Groups.Add(Me.Group5)
        resources.ApplyResources(Me.Tab1, "Tab1")
        Me.Tab1.Name = "Tab1"
        '
        'Group1
        '
        Me.Group1.Items.Add(Me.AlphaButton)
        Me.Group1.Items.Add(Me.BetaButton)
        Me.Group1.Items.Add(Me.GammaButton)
        resources.ApplyResources(Me.Group1, "Group1")
        Me.Group1.Name = "Group1"
        '
        'AlphaButton
        '
        resources.ApplyResources(Me.AlphaButton, "AlphaButton")
        Me.AlphaButton.Name = "AlphaButton"
        '
        'BetaButton
        '
        resources.ApplyResources(Me.BetaButton, "BetaButton")
        Me.BetaButton.Name = "BetaButton"
        '
        'GammaButton
        '
        resources.ApplyResources(Me.GammaButton, "GammaButton")
        Me.GammaButton.Name = "GammaButton"
        '
        'Control
        '
        Me.Control.Items.Add(Me.ManualButton)
        Me.Control.Items.Add(Me.SynchButton)
        Me.Control.Items.Add(Me.SimulationButton)
        Me.Control.Items.Add(Me.AutoHedgeButton)
        resources.ApplyResources(Me.Control, "Control")
        Me.Control.Name = "Control"
        '
        'ManualButton
        '
        resources.ApplyResources(Me.ManualButton, "ManualButton")
        Me.ManualButton.Name = "ManualButton"
        '
        'SynchButton
        '
        resources.ApplyResources(Me.SynchButton, "SynchButton")
        Me.SynchButton.Name = "SynchButton"
        '
        'SimulationButton
        '
        resources.ApplyResources(Me.SimulationButton, "SimulationButton")
        Me.SimulationButton.Name = "SimulationButton"
        '
        'AutoHedgeButton
        '
        resources.ApplyResources(Me.AutoHedgeButton, "AutoHedgeButton")
        Me.AutoHedgeButton.Name = "AutoHedgeButton"
        '
        'Group2
        '
        Me.Group2.Items.Add(Me.DashboardButton)
        Me.Group2.Items.Add(Me.FinChartsButton)
        resources.ApplyResources(Me.Group2, "Group2")
        Me.Group2.Name = "Group2"
        '
        'DashboardButton
        '
        resources.ApplyResources(Me.DashboardButton, "DashboardButton")
        Me.DashboardButton.Name = "DashboardButton"
        '
        'FinChartsButton
        '
        resources.ApplyResources(Me.FinChartsButton, "FinChartsButton")
        Me.FinChartsButton.Name = "FinChartsButton"
        '
        'Group3
        '
        Me.Group3.Items.Add(Me.InitialPositionsButton)
        Me.Group3.Items.Add(Me.AcquiredPositionsButton)
        Me.Group3.Items.Add(Me.TransactionQButton)
        Me.Group3.Items.Add(Me.ResetAPButton)
        Me.Group3.Items.Add(Me.EditAPButton)
        Me.Group3.Items.Add(Me.UploadAPButton)
        Me.Group3.Items.Add(Me.ConfirmationButton)
        resources.ApplyResources(Me.Group3, "Group3")
        Me.Group3.Name = "Group3"
        '
        'InitialPositionsButton
        '
        resources.ApplyResources(Me.InitialPositionsButton, "InitialPositionsButton")
        Me.InitialPositionsButton.Name = "InitialPositionsButton"
        '
        'AcquiredPositionsButton
        '
        resources.ApplyResources(Me.AcquiredPositionsButton, "AcquiredPositionsButton")
        Me.AcquiredPositionsButton.Name = "AcquiredPositionsButton"
        '
        'TransactionQButton
        '
        resources.ApplyResources(Me.TransactionQButton, "TransactionQButton")
        Me.TransactionQButton.Name = "TransactionQButton"
        '
        'ResetAPButton
        '
        resources.ApplyResources(Me.ResetAPButton, "ResetAPButton")
        Me.ResetAPButton.Name = "ResetAPButton"
        '
        'EditAPButton
        '
        resources.ApplyResources(Me.EditAPButton, "EditAPButton")
        Me.EditAPButton.Name = "EditAPButton"
        '
        'UploadAPButton
        '
        resources.ApplyResources(Me.UploadAPButton, "UploadAPButton")
        Me.UploadAPButton.Name = "UploadAPButton"
        '
        'ConfirmationButton
        '
        resources.ApplyResources(Me.ConfirmationButton, "ConfirmationButton")
        Me.ConfirmationButton.Name = "ConfirmationButton"
        '
        'Group4
        '
        Me.Group4.Items.Add(Me.StockMktButton)
        Me.Group4.Items.Add(Me.OptionMktButton)
        Me.Group4.Items.Add(Me.SP500Button)
        Me.Group4.Items.Add(Me.SettingsButton)
        Me.Group4.Items.Add(Me.TransactionButton)
        resources.ApplyResources(Me.Group4, "Group4")
        Me.Group4.Name = "Group4"
        '
        'StockMktButton
        '
        resources.ApplyResources(Me.StockMktButton, "StockMktButton")
        Me.StockMktButton.Name = "StockMktButton"
        '
        'OptionMktButton
        '
        resources.ApplyResources(Me.OptionMktButton, "OptionMktButton")
        Me.OptionMktButton.Name = "OptionMktButton"
        '
        'SP500Button
        '
        resources.ApplyResources(Me.SP500Button, "SP500Button")
        Me.SP500Button.Name = "SP500Button"
        '
        'SettingsButton
        '
        resources.ApplyResources(Me.SettingsButton, "SettingsButton")
        Me.SettingsButton.Name = "SettingsButton"
        '
        'TransactionButton
        '
        resources.ApplyResources(Me.TransactionButton, "TransactionButton")
        Me.TransactionButton.Name = "TransactionButton"
        '
        'Group5
        '
        Me.Group5.Items.Add(Me.QuitButton)
        Me.Group5.Items.Add(Me.ExcludeIPButton)
        resources.ApplyResources(Me.Group5, "Group5")
        Me.Group5.Name = "Group5"
        '
        'QuitButton
        '
        resources.ApplyResources(Me.QuitButton, "QuitButton")
        Me.QuitButton.Name = "QuitButton"
        '
        'ExcludeIPButton
        '
        resources.ApplyResources(Me.ExcludeIPButton, "ExcludeIPButton")
        Me.ExcludeIPButton.Name = "ExcludeIPButton"
        '
        'stRibbon
        '
        Me.Name = "stRibbon"
        Me.RibbonType = "Microsoft.Excel.Workbook"
        Me.Tabs.Add(Me.Tab1)
        Me.Tab1.ResumeLayout(False)
        Me.Tab1.PerformLayout()
        Me.Group1.ResumeLayout(False)
        Me.Group1.PerformLayout()
        Me.Control.ResumeLayout(False)
        Me.Control.PerformLayout()
        Me.Group2.ResumeLayout(False)
        Me.Group2.PerformLayout()
        Me.Group3.ResumeLayout(False)
        Me.Group3.PerformLayout()
        Me.Group4.ResumeLayout(False)
        Me.Group4.PerformLayout()
        Me.Group5.ResumeLayout(False)
        Me.Group5.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents Tab1 As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents Group1 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents AlphaButton As Microsoft.Office.Tools.Ribbon.RibbonToggleButton
    Friend WithEvents BetaButton As Microsoft.Office.Tools.Ribbon.RibbonToggleButton
    Friend WithEvents GammaButton As Microsoft.Office.Tools.Ribbon.RibbonToggleButton
    Friend WithEvents Group2 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents DashboardButton As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Group3 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents InitialPositionsButton As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents AcquiredPositionsButton As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Group4 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents StockMktButton As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents OptionMktButton As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents SP500Button As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents SettingsButton As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents TransactionButton As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Group5 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents QuitButton As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents TransactionQButton As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents FinChartsButton As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ResetAPButton As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents EditAPButton As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents UploadAPButton As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ExcludeIPButton As Microsoft.Office.Tools.Ribbon.RibbonToggleButton
    Friend WithEvents Control As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents ManualButton As Microsoft.Office.Tools.Ribbon.RibbonToggleButton
    Friend WithEvents SynchButton As Microsoft.Office.Tools.Ribbon.RibbonToggleButton
    Friend WithEvents SimulationButton As Microsoft.Office.Tools.Ribbon.RibbonToggleButton
    Friend WithEvents AutoHedgeButton As Microsoft.Office.Tools.Ribbon.RibbonToggleButton
    Friend WithEvents ConfirmationButton As Microsoft.Office.Tools.Ribbon.RibbonButton
End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()> _
    Friend ReadOnly Property SpartanTraderRibbon() As stRibbon
        Get
            Return Me.GetRibbon(Of stRibbon)()
        End Get
    End Property
End Class
