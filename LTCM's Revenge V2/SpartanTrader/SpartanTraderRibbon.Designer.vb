Partial Class SpartanTraderRibbon
    Inherits Microsoft.Office.Tools.Ribbon.RibbonBase

    <System.Diagnostics.DebuggerNonUserCode()>
    Public Sub New(ByVal container As System.ComponentModel.IContainer)
        MyClass.New()

        'Required for Windows.Forms Class Composition Designer support
        If (container IsNot Nothing) Then
            container.Add(Me)
        End If

    End Sub

    <System.Diagnostics.DebuggerNonUserCode()>
    Public Sub New()
        MyBase.New(Globals.Factory.GetRibbonFactory())

        'This call is required by the Component Designer.
        InitializeComponent()

    End Sub

    'Component overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
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
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.Tab1 = Me.Factory.CreateRibbonTab
        Me.Group1 = Me.Factory.CreateRibbonGroup
        Me.AlphaBtn = Me.Factory.CreateRibbonToggleButton
        Me.BetaBtn = Me.Factory.CreateRibbonToggleButton
        Me.GammaBtn = Me.Factory.CreateRibbonToggleButton
        Me.Group2 = Me.Factory.CreateRibbonGroup
        Me.DashboardBtn = Me.Factory.CreateRibbonButton
        Me.FinChartsBtn = Me.Factory.CreateRibbonButton
        Me.Group3 = Me.Factory.CreateRibbonGroup
        Me.InitialPositionsBtn = Me.Factory.CreateRibbonButton
        Me.AcquiredPositionsBtn = Me.Factory.CreateRibbonButton
        Me.TransactionQueueBtn = Me.Factory.CreateRibbonButton
        Me.ResetAPBtn = Me.Factory.CreateRibbonButton
        Me.EditAPBtn = Me.Factory.CreateRibbonButton
        Me.UploadAPBtn = Me.Factory.CreateRibbonButton
        Me.Group4 = Me.Factory.CreateRibbonGroup
        Me.StockMktBtn = Me.Factory.CreateRibbonButton
        Me.OptionMktBtn = Me.Factory.CreateRibbonButton
        Me.SP500Btn = Me.Factory.CreateRibbonButton
        Me.SettingsBtn = Me.Factory.CreateRibbonButton
        Me.TCostsBtn = Me.Factory.CreateRibbonButton
        Me.Group5 = Me.Factory.CreateRibbonGroup
        Me.QuitBtn = Me.Factory.CreateRibbonButton
        Me.ExcludeIPBtn = Me.Factory.CreateRibbonToggleButton
        Me.Control = Me.Factory.CreateRibbonGroup
        Me.ManualBtn = Me.Factory.CreateRibbonToggleButton
        Me.SynchBtn = Me.Factory.CreateRibbonToggleButton
        Me.SimulationBtn = Me.Factory.CreateRibbonToggleButton
        Me.AutoHedgeBtn = Me.Factory.CreateRibbonToggleButton
        Me.ConfirmationBtn = Me.Factory.CreateRibbonButton
        Me.Tab1.SuspendLayout()
        Me.Group1.SuspendLayout()
        Me.Group2.SuspendLayout()
        Me.Group3.SuspendLayout()
        Me.Group4.SuspendLayout()
        Me.Group5.SuspendLayout()
        Me.Control.SuspendLayout()
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
        Me.Tab1.Label = "Spartan Trader"
        Me.Tab1.Name = "Tab1"
        '
        'Group1
        '
        Me.Group1.Items.Add(Me.AlphaBtn)
        Me.Group1.Items.Add(Me.BetaBtn)
        Me.Group1.Items.Add(Me.GammaBtn)
        Me.Group1.Label = "Database"
        Me.Group1.Name = "Group1"
        '
        'AlphaBtn
        '
        Me.AlphaBtn.Label = "Alpha"
        Me.AlphaBtn.Name = "AlphaBtn"
        '
        'BetaBtn
        '
        Me.BetaBtn.Label = "Beta"
        Me.BetaBtn.Name = "BetaBtn"
        '
        'GammaBtn
        '
        Me.GammaBtn.Label = "Gamma"
        Me.GammaBtn.Name = "GammaBtn"
        '
        'Group2
        '
        Me.Group2.Items.Add(Me.DashboardBtn)
        Me.Group2.Items.Add(Me.FinChartsBtn)
        Me.Group2.Label = "Dashboard"
        Me.Group2.Name = "Group2"
        '
        'DashboardBtn
        '
        Me.DashboardBtn.Label = "Dashboard"
        Me.DashboardBtn.Name = "DashboardBtn"
        '
        'FinChartsBtn
        '
        Me.FinChartsBtn.Label = "FinCharts"
        Me.FinChartsBtn.Name = "FinChartsBtn"
        '
        'Group3
        '
        Me.Group3.Items.Add(Me.InitialPositionsBtn)
        Me.Group3.Items.Add(Me.AcquiredPositionsBtn)
        Me.Group3.Items.Add(Me.TransactionQueueBtn)
        Me.Group3.Items.Add(Me.ResetAPBtn)
        Me.Group3.Items.Add(Me.EditAPBtn)
        Me.Group3.Items.Add(Me.UploadAPBtn)
        Me.Group3.Items.Add(Me.ConfirmationBtn)
        Me.Group3.Label = "Protfolio Management"
        Me.Group3.Name = "Group3"
        '
        'InitialPositionsBtn
        '
        Me.InitialPositionsBtn.Label = "Initial Positions"
        Me.InitialPositionsBtn.Name = "InitialPositionsBtn"
        '
        'AcquiredPositionsBtn
        '
        Me.AcquiredPositionsBtn.Label = "Acquired Positions"
        Me.AcquiredPositionsBtn.Name = "AcquiredPositionsBtn"
        '
        'TransactionQueueBtn
        '
        Me.TransactionQueueBtn.Label = "Transaction Q"
        Me.TransactionQueueBtn.Name = "TransactionQueueBtn"
        '
        'ResetAPBtn
        '
        Me.ResetAPBtn.Label = "Reset AP"
        Me.ResetAPBtn.Name = "ResetAPBtn"
        '
        'EditAPBtn
        '
        Me.EditAPBtn.Label = "Edit AP"
        Me.EditAPBtn.Name = "EditAPBtn"
        '
        'UploadAPBtn
        '
        Me.UploadAPBtn.Label = "Upload AP"
        Me.UploadAPBtn.Name = "UploadAPBtn"
        '
        'Group4
        '
        Me.Group4.Items.Add(Me.StockMktBtn)
        Me.Group4.Items.Add(Me.OptionMktBtn)
        Me.Group4.Items.Add(Me.SP500Btn)
        Me.Group4.Items.Add(Me.SettingsBtn)
        Me.Group4.Items.Add(Me.TCostsBtn)
        Me.Group4.Label = "Business Intelligence"
        Me.Group4.Name = "Group4"
        '
        'StockMktBtn
        '
        Me.StockMktBtn.Label = "Stock Mkt"
        Me.StockMktBtn.Name = "StockMktBtn"
        '
        'OptionMktBtn
        '
        Me.OptionMktBtn.Label = "Option Mkt"
        Me.OptionMktBtn.Name = "OptionMktBtn"
        '
        'SP500Btn
        '
        Me.SP500Btn.Label = "SP 500"
        Me.SP500Btn.Name = "SP500Btn"
        '
        'SettingsBtn
        '
        Me.SettingsBtn.Label = "Settings"
        Me.SettingsBtn.Name = "SettingsBtn"
        '
        'TCostsBtn
        '
        Me.TCostsBtn.Label = "T Costs"
        Me.TCostsBtn.Name = "TCostsBtn"
        '
        'Group5
        '
        Me.Group5.Items.Add(Me.QuitBtn)
        Me.Group5.Items.Add(Me.ExcludeIPBtn)
        Me.Group5.Label = "Control"
        Me.Group5.Name = "Group5"
        '
        'QuitBtn
        '
        Me.QuitBtn.Label = "Quit"
        Me.QuitBtn.Name = "QuitBtn"
        '
        'ExcludeIPBtn
        '
        Me.ExcludeIPBtn.Label = "Exclude IP"
        Me.ExcludeIPBtn.Name = "ExcludeIPBtn"
        '
        'Control
        '
        Me.Control.Items.Add(Me.ManualBtn)
        Me.Control.Items.Add(Me.SynchBtn)
        Me.Control.Items.Add(Me.SimulationBtn)
        Me.Control.Items.Add(Me.AutoHedgeBtn)
        Me.Control.Label = "Control"
        Me.Control.Name = "Control"
        '
        'ManualBtn
        '
        Me.ManualBtn.Label = "Manual"
        Me.ManualBtn.Name = "ManualBtn"
        '
        'SynchBtn
        '
        Me.SynchBtn.Label = "Synch"
        Me.SynchBtn.Name = "SynchBtn"
        '
        'SimulationBtn
        '
        Me.SimulationBtn.Label = "Simulation"
        Me.SimulationBtn.Name = "SimulationBtn"
        '
        'AutoHedgeBtn
        '
        Me.AutoHedgeBtn.Label = "Auto Hedge"
        Me.AutoHedgeBtn.Name = "AutoHedgeBtn"
        '
        'ConfirmationBtn
        '
        Me.ConfirmationBtn.Label = "Confirmation"
        Me.ConfirmationBtn.Name = "ConfirmationBtn"
        '
        'SpartanTraderRibbon
        '
        Me.Name = "SpartanTraderRibbon"
        Me.RibbonType = "Microsoft.Excel.Workbook"
        Me.Tabs.Add(Me.Tab1)
        Me.Tab1.ResumeLayout(False)
        Me.Tab1.PerformLayout()
        Me.Group1.ResumeLayout(False)
        Me.Group1.PerformLayout()
        Me.Group2.ResumeLayout(False)
        Me.Group2.PerformLayout()
        Me.Group3.ResumeLayout(False)
        Me.Group3.PerformLayout()
        Me.Group4.ResumeLayout(False)
        Me.Group4.PerformLayout()
        Me.Group5.ResumeLayout(False)
        Me.Group5.PerformLayout()
        Me.Control.ResumeLayout(False)
        Me.Control.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents Tab1 As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents Group1 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents AlphaBtn As Microsoft.Office.Tools.Ribbon.RibbonToggleButton
    Friend WithEvents BetaBtn As Microsoft.Office.Tools.Ribbon.RibbonToggleButton
    Friend WithEvents GammaBtn As Microsoft.Office.Tools.Ribbon.RibbonToggleButton
    Friend WithEvents Group2 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents DashboardBtn As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Group3 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents InitialPositionsBtn As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents AcquiredPositionsBtn As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Group4 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents StockMktBtn As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents OptionMktBtn As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents SP500Btn As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents SettingsBtn As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents TCostsBtn As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Group5 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents QuitBtn As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents TransactionQueueBtn As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ResetAPBtn As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents EditAPBtn As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents UploadAPBtn As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents FinChartsBtn As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ExcludeIPBtn As Microsoft.Office.Tools.Ribbon.RibbonToggleButton
    Friend WithEvents Control As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents ManualBtn As Microsoft.Office.Tools.Ribbon.RibbonToggleButton
    Friend WithEvents SynchBtn As Microsoft.Office.Tools.Ribbon.RibbonToggleButton
    Friend WithEvents SimulationBtn As Microsoft.Office.Tools.Ribbon.RibbonToggleButton
    Friend WithEvents AutoHedgeBtn As Microsoft.Office.Tools.Ribbon.RibbonToggleButton
    Friend WithEvents ConfirmationBtn As Microsoft.Office.Tools.Ribbon.RibbonButton
End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()> _
    Friend ReadOnly Property SpartanTraderRibbon() As SpartanTraderRibbon
        Get
            Return Me.GetRibbon(Of SpartanTraderRibbon)()
        End Get
    End Property
End Class
