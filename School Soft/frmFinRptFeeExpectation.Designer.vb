<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmFinRptFeeExpectation
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
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

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.pnlFeeExpectation = New System.Windows.Forms.Panel()
        Me.pnlVoteSummary = New System.Windows.Forms.Panel()
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.cboTerm = New System.Windows.Forms.ComboBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.cboYear = New System.Windows.Forms.ComboBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.btnClose = New System.Windows.Forms.Button()
        Me.btnLoad = New System.Windows.Forms.Button()
        Me.StatusStrip12 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip13 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip14 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip5 = New System.Windows.Forms.StatusStrip()
        Me.crtVwFeeExpectation = New CrystalDecisions.Windows.Forms.CrystalReportViewer()
        Me.StatusStrip2 = New System.Windows.Forms.StatusStrip()
        Me.ToolStripStatusLabel1 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.pnlFeeExpectation.SuspendLayout()
        Me.pnlVoteSummary.SuspendLayout()
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.StatusStrip2.SuspendLayout()
        Me.SuspendLayout()
        '
        'pnlFeeExpectation
        '
        Me.pnlFeeExpectation.BackColor = System.Drawing.SystemColors.Control
        Me.pnlFeeExpectation.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pnlFeeExpectation.Controls.Add(Me.pnlVoteSummary)
        Me.pnlFeeExpectation.Location = New System.Drawing.Point(12, 12)
        Me.pnlFeeExpectation.Name = "pnlFeeExpectation"
        Me.pnlFeeExpectation.Size = New System.Drawing.Size(999, 549)
        Me.pnlFeeExpectation.TabIndex = 0
        '
        'pnlVoteSummary
        '
        Me.pnlVoteSummary.BackColor = System.Drawing.SystemColors.Control
        Me.pnlVoteSummary.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pnlVoteSummary.Controls.Add(Me.SplitContainer1)
        Me.pnlVoteSummary.Controls.Add(Me.StatusStrip2)
        Me.pnlVoteSummary.Location = New System.Drawing.Point(-1, -1)
        Me.pnlVoteSummary.Name = "pnlVoteSummary"
        Me.pnlVoteSummary.Size = New System.Drawing.Size(999, 564)
        Me.pnlVoteSummary.TabIndex = 1
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer1.Location = New System.Drawing.Point(0, 25)
        Me.SplitContainer1.Name = "SplitContainer1"
        Me.SplitContainer1.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.Panel1)
        Me.SplitContainer1.Panel1.Controls.Add(Me.StatusStrip12)
        Me.SplitContainer1.Panel1.Controls.Add(Me.StatusStrip13)
        Me.SplitContainer1.Panel1.Controls.Add(Me.StatusStrip14)
        Me.SplitContainer1.Panel1.Controls.Add(Me.StatusStrip5)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.crtVwFeeExpectation)
        Me.SplitContainer1.Size = New System.Drawing.Size(997, 537)
        Me.SplitContainer1.SplitterDistance = 51
        Me.SplitContainer1.SplitterWidth = 1
        Me.SplitContainer1.TabIndex = 25
        '
        'Panel1
        '
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel1.Controls.Add(Me.cboTerm)
        Me.Panel1.Controls.Add(Me.Label2)
        Me.Panel1.Controls.Add(Me.cboYear)
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Controls.Add(Me.btnClose)
        Me.Panel1.Controls.Add(Me.btnLoad)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel1.Location = New System.Drawing.Point(5, 5)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(987, 41)
        Me.Panel1.TabIndex = 65
        '
        'cboTerm
        '
        Me.cboTerm.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboTerm.FormattingEnabled = True
        Me.cboTerm.Location = New System.Drawing.Point(92, 9)
        Me.cboTerm.Name = "cboTerm"
        Me.cboTerm.Size = New System.Drawing.Size(138, 21)
        Me.cboTerm.TabIndex = 51
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(18, 12)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(68, 13)
        Me.Label2.TabIndex = 50
        Me.Label2.Text = "Term Name :"
        '
        'cboYear
        '
        Me.cboYear.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboYear.FormattingEnabled = True
        Me.cboYear.Location = New System.Drawing.Point(369, 9)
        Me.cboYear.Name = "cboYear"
        Me.cboYear.Size = New System.Drawing.Size(123, 21)
        Me.cboYear.TabIndex = 49
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(328, 12)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(35, 13)
        Me.Label1.TabIndex = 48
        Me.Label1.Text = "Year :"
        '
        'btnClose
        '
        Me.btnClose.Location = New System.Drawing.Point(891, 7)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(75, 23)
        Me.btnClose.TabIndex = 46
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'btnLoad
        '
        Me.btnLoad.Location = New System.Drawing.Point(674, 7)
        Me.btnLoad.Name = "btnLoad"
        Me.btnLoad.Size = New System.Drawing.Size(75, 23)
        Me.btnLoad.TabIndex = 45
        Me.btnLoad.Text = "Load"
        Me.btnLoad.UseVisualStyleBackColor = True
        '
        'StatusStrip12
        '
        Me.StatusStrip12.AutoSize = False
        Me.StatusStrip12.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip12.Dock = System.Windows.Forms.DockStyle.Left
        Me.StatusStrip12.Location = New System.Drawing.Point(0, 5)
        Me.StatusStrip12.Name = "StatusStrip12"
        Me.StatusStrip12.Size = New System.Drawing.Size(5, 41)
        Me.StatusStrip12.SizingGrip = False
        Me.StatusStrip12.TabIndex = 64
        Me.StatusStrip12.Text = "StatusStrip12"
        '
        'StatusStrip13
        '
        Me.StatusStrip13.AutoSize = False
        Me.StatusStrip13.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip13.Location = New System.Drawing.Point(0, 46)
        Me.StatusStrip13.Name = "StatusStrip13"
        Me.StatusStrip13.Size = New System.Drawing.Size(992, 5)
        Me.StatusStrip13.SizingGrip = False
        Me.StatusStrip13.TabIndex = 63
        Me.StatusStrip13.Text = "StatusStrip13"
        '
        'StatusStrip14
        '
        Me.StatusStrip14.AutoSize = False
        Me.StatusStrip14.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip14.Dock = System.Windows.Forms.DockStyle.Top
        Me.StatusStrip14.Location = New System.Drawing.Point(0, 0)
        Me.StatusStrip14.Name = "StatusStrip14"
        Me.StatusStrip14.Size = New System.Drawing.Size(992, 5)
        Me.StatusStrip14.SizingGrip = False
        Me.StatusStrip14.TabIndex = 62
        Me.StatusStrip14.Text = "StatusStrip14"
        '
        'StatusStrip5
        '
        Me.StatusStrip5.AutoSize = False
        Me.StatusStrip5.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip5.Dock = System.Windows.Forms.DockStyle.Right
        Me.StatusStrip5.Location = New System.Drawing.Point(992, 0)
        Me.StatusStrip5.Name = "StatusStrip5"
        Me.StatusStrip5.Size = New System.Drawing.Size(5, 51)
        Me.StatusStrip5.SizingGrip = False
        Me.StatusStrip5.TabIndex = 61
        Me.StatusStrip5.Text = "StatusStrip5"
        '
        'crtVwFeeExpectation
        '
        Me.crtVwFeeExpectation.ActiveViewIndex = -1
        Me.crtVwFeeExpectation.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.crtVwFeeExpectation.Cursor = System.Windows.Forms.Cursors.Default
        Me.crtVwFeeExpectation.DisplayStatusBar = False
        Me.crtVwFeeExpectation.Dock = System.Windows.Forms.DockStyle.Fill
        Me.crtVwFeeExpectation.Location = New System.Drawing.Point(0, 0)
        Me.crtVwFeeExpectation.Name = "crtVwFeeExpectation"
        Me.crtVwFeeExpectation.ShowGroupTreeButton = False
        Me.crtVwFeeExpectation.ShowLogo = False
        Me.crtVwFeeExpectation.ShowParameterPanelButton = False
        Me.crtVwFeeExpectation.ShowTextSearchButton = False
        Me.crtVwFeeExpectation.Size = New System.Drawing.Size(997, 485)
        Me.crtVwFeeExpectation.TabIndex = 37
        Me.crtVwFeeExpectation.ToolPanelView = CrystalDecisions.Windows.Forms.ToolPanelViewType.None
        '
        'StatusStrip2
        '
        Me.StatusStrip2.AutoSize = False
        Me.StatusStrip2.BackColor = System.Drawing.Color.LightSeaGreen
        Me.StatusStrip2.Dock = System.Windows.Forms.DockStyle.Top
        Me.StatusStrip2.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripStatusLabel1})
        Me.StatusStrip2.Location = New System.Drawing.Point(0, 0)
        Me.StatusStrip2.Name = "StatusStrip2"
        Me.StatusStrip2.Size = New System.Drawing.Size(997, 25)
        Me.StatusStrip2.SizingGrip = False
        Me.StatusStrip2.TabIndex = 24
        Me.StatusStrip2.Text = "StatusStrip2"
        '
        'ToolStripStatusLabel1
        '
        Me.ToolStripStatusLabel1.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ToolStripStatusLabel1.Name = "ToolStripStatusLabel1"
        Me.ToolStripStatusLabel1.Size = New System.Drawing.Size(185, 20)
        Me.ToolStripStatusLabel1.Text = "FINANCE FEE EXPECTATIONS"
        '
        'frmFinRptFeeExpectation
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.LightSkyBlue
        Me.ClientSize = New System.Drawing.Size(1023, 574)
        Me.Controls.Add(Me.pnlFeeExpectation)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Name = "frmFinRptFeeExpectation"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Fee Expectation"
        Me.pnlFeeExpectation.ResumeLayout(False)
        Me.pnlVoteSummary.ResumeLayout(False)
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.ResumeLayout(False)
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.StatusStrip2.ResumeLayout(False)
        Me.StatusStrip2.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents pnlFeeExpectation As System.Windows.Forms.Panel
    Friend WithEvents pnlVoteSummary As System.Windows.Forms.Panel
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents cboTerm As System.Windows.Forms.ComboBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents cboYear As System.Windows.Forms.ComboBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents btnLoad As System.Windows.Forms.Button
    Friend WithEvents StatusStrip12 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip13 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip14 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip5 As System.Windows.Forms.StatusStrip
    Friend WithEvents crtVwFeeExpectation As CrystalDecisions.Windows.Forms.CrystalReportViewer
    Friend WithEvents StatusStrip2 As System.Windows.Forms.StatusStrip
    Friend WithEvents ToolStripStatusLabel1 As System.Windows.Forms.ToolStripStatusLabel
End Class
