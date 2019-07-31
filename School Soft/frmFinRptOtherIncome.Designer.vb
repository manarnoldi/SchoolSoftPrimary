<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmFinRptOtherIncome
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
        Me.pnlOtherIncome = New System.Windows.Forms.Panel()
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.cboTerm = New System.Windows.Forms.ComboBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.rbTermly = New System.Windows.Forms.RadioButton()
        Me.btnClose = New System.Windows.Forms.Button()
        Me.btnLoad = New System.Windows.Forms.Button()
        Me.dtpDate = New System.Windows.Forms.DateTimePicker()
        Me.lblDateLabel = New System.Windows.Forms.Label()
        Me.rbMonthly = New System.Windows.Forms.RadioButton()
        Me.rbYearly = New System.Windows.Forms.RadioButton()
        Me.StatusStrip12 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip13 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip14 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip5 = New System.Windows.Forms.StatusStrip()
        Me.crtVwOtherInc = New CrystalDecisions.Windows.Forms.CrystalReportViewer()
        Me.StatusStrip2 = New System.Windows.Forms.StatusStrip()
        Me.ToolStripStatusLabel1 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.pnlOtherIncome.SuspendLayout()
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.StatusStrip2.SuspendLayout()
        Me.SuspendLayout()
        '
        'pnlOtherIncome
        '
        Me.pnlOtherIncome.BackColor = System.Drawing.SystemColors.Control
        Me.pnlOtherIncome.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pnlOtherIncome.Controls.Add(Me.SplitContainer1)
        Me.pnlOtherIncome.Controls.Add(Me.StatusStrip2)
        Me.pnlOtherIncome.Location = New System.Drawing.Point(12, 12)
        Me.pnlOtherIncome.Name = "pnlOtherIncome"
        Me.pnlOtherIncome.Size = New System.Drawing.Size(999, 564)
        Me.pnlOtherIncome.TabIndex = 1
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
        Me.SplitContainer1.Panel2.Controls.Add(Me.crtVwOtherInc)
        Me.SplitContainer1.Size = New System.Drawing.Size(997, 537)
        Me.SplitContainer1.SplitterDistance = 51
        Me.SplitContainer1.SplitterWidth = 1
        Me.SplitContainer1.TabIndex = 25
        '
        'Panel1
        '
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel1.Controls.Add(Me.cboTerm)
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Controls.Add(Me.rbTermly)
        Me.Panel1.Controls.Add(Me.btnClose)
        Me.Panel1.Controls.Add(Me.btnLoad)
        Me.Panel1.Controls.Add(Me.dtpDate)
        Me.Panel1.Controls.Add(Me.lblDateLabel)
        Me.Panel1.Controls.Add(Me.rbMonthly)
        Me.Panel1.Controls.Add(Me.rbYearly)
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
        Me.cboTerm.Location = New System.Drawing.Point(395, 11)
        Me.cboTerm.Name = "cboTerm"
        Me.cboTerm.Size = New System.Drawing.Size(104, 21)
        Me.cboTerm.TabIndex = 49
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(355, 14)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(34, 13)
        Me.Label1.TabIndex = 48
        Me.Label1.Text = "Term:"
        '
        'rbTermly
        '
        Me.rbTermly.AutoSize = True
        Me.rbTermly.Location = New System.Drawing.Point(127, 12)
        Me.rbTermly.Name = "rbTermly"
        Me.rbTermly.Size = New System.Drawing.Size(56, 17)
        Me.rbTermly.TabIndex = 47
        Me.rbTermly.Text = "Termly"
        Me.rbTermly.UseVisualStyleBackColor = True
        '
        'btnClose
        '
        Me.btnClose.Location = New System.Drawing.Point(904, 9)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(75, 23)
        Me.btnClose.TabIndex = 46
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'btnLoad
        '
        Me.btnLoad.Location = New System.Drawing.Point(784, 9)
        Me.btnLoad.Name = "btnLoad"
        Me.btnLoad.Size = New System.Drawing.Size(75, 23)
        Me.btnLoad.TabIndex = 45
        Me.btnLoad.Text = "Load"
        Me.btnLoad.UseVisualStyleBackColor = True
        '
        'dtpDate
        '
        Me.dtpDate.CustomFormat = " MMMM, yyyy"
        Me.dtpDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpDate.Location = New System.Drawing.Point(619, 11)
        Me.dtpDate.Name = "dtpDate"
        Me.dtpDate.Size = New System.Drawing.Size(129, 20)
        Me.dtpDate.TabIndex = 5
        Me.dtpDate.Value = New Date(2014, 12, 9, 0, 0, 0, 0)
        '
        'lblDateLabel
        '
        Me.lblDateLabel.AutoSize = True
        Me.lblDateLabel.Location = New System.Drawing.Point(539, 14)
        Me.lblDateLabel.Name = "lblDateLabel"
        Me.lblDateLabel.Size = New System.Drawing.Size(76, 13)
        Me.lblDateLabel.TabIndex = 4
        Me.lblDateLabel.Text = "Select Month :"
        '
        'rbMonthly
        '
        Me.rbMonthly.AutoSize = True
        Me.rbMonthly.Location = New System.Drawing.Point(15, 12)
        Me.rbMonthly.Name = "rbMonthly"
        Me.rbMonthly.Size = New System.Drawing.Size(62, 17)
        Me.rbMonthly.TabIndex = 3
        Me.rbMonthly.Text = "Monthly"
        Me.rbMonthly.UseVisualStyleBackColor = True
        '
        'rbYearly
        '
        Me.rbYearly.AutoSize = True
        Me.rbYearly.Location = New System.Drawing.Point(244, 12)
        Me.rbYearly.Name = "rbYearly"
        Me.rbYearly.Size = New System.Drawing.Size(54, 17)
        Me.rbYearly.TabIndex = 2
        Me.rbYearly.Text = "Yearly"
        Me.rbYearly.UseVisualStyleBackColor = True
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
        'crtVwOtherInc
        '
        Me.crtVwOtherInc.ActiveViewIndex = -1
        Me.crtVwOtherInc.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.crtVwOtherInc.Cursor = System.Windows.Forms.Cursors.Default
        Me.crtVwOtherInc.DisplayStatusBar = False
        Me.crtVwOtherInc.Dock = System.Windows.Forms.DockStyle.Fill
        Me.crtVwOtherInc.Location = New System.Drawing.Point(0, 0)
        Me.crtVwOtherInc.Name = "crtVwOtherInc"
        Me.crtVwOtherInc.ShowGroupTreeButton = False
        Me.crtVwOtherInc.ShowLogo = False
        Me.crtVwOtherInc.ShowParameterPanelButton = False
        Me.crtVwOtherInc.ShowTextSearchButton = False
        Me.crtVwOtherInc.Size = New System.Drawing.Size(997, 485)
        Me.crtVwOtherInc.TabIndex = 37
        Me.crtVwOtherInc.ToolPanelView = CrystalDecisions.Windows.Forms.ToolPanelViewType.None
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
        Me.ToolStripStatusLabel1.Size = New System.Drawing.Size(217, 20)
        Me.ToolStripStatusLabel1.Text = "FINANCE OTHER INCOME REPORT"
        '
        'frmFinRptOtherIncome
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.LightSkyBlue
        Me.ClientSize = New System.Drawing.Size(1023, 588)
        Me.Controls.Add(Me.pnlOtherIncome)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Name = "frmFinRptOtherIncome"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Other Income"
        Me.pnlOtherIncome.ResumeLayout(False)
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
    Friend WithEvents pnlOtherIncome As System.Windows.Forms.Panel
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents cboTerm As System.Windows.Forms.ComboBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents rbTermly As System.Windows.Forms.RadioButton
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents btnLoad As System.Windows.Forms.Button
    Friend WithEvents dtpDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents lblDateLabel As System.Windows.Forms.Label
    Friend WithEvents rbMonthly As System.Windows.Forms.RadioButton
    Friend WithEvents rbYearly As System.Windows.Forms.RadioButton
    Friend WithEvents StatusStrip12 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip13 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip14 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip5 As System.Windows.Forms.StatusStrip
    Friend WithEvents crtVwOtherInc As CrystalDecisions.Windows.Forms.CrystalReportViewer
    Friend WithEvents StatusStrip2 As System.Windows.Forms.StatusStrip
    Friend WithEvents ToolStripStatusLabel1 As System.Windows.Forms.ToolStripStatusLabel
End Class
