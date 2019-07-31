<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmFinRptAccountsAdj
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
        Me.pnlAccountAdj = New System.Windows.Forms.Panel()
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.btnLoad = New System.Windows.Forms.Button()
        Me.dtpDate = New System.Windows.Forms.DateTimePicker()
        Me.lblDateLabel = New System.Windows.Forms.Label()
        Me.rbYearly = New System.Windows.Forms.RadioButton()
        Me.StatusStrip12 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip13 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip14 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip5 = New System.Windows.Forms.StatusStrip()
        Me.crtVwFinAccAdjRpt = New CrystalDecisions.Windows.Forms.CrystalReportViewer()
        Me.StatusStrip2 = New System.Windows.Forms.StatusStrip()
        Me.ToolStripStatusLabel1 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.btnClose = New System.Windows.Forms.Button()
        Me.pnlAccountAdj.SuspendLayout()
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.StatusStrip2.SuspendLayout()
        Me.SuspendLayout()
        '
        'pnlAccountAdj
        '
        Me.pnlAccountAdj.BackColor = System.Drawing.SystemColors.Control
        Me.pnlAccountAdj.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pnlAccountAdj.Controls.Add(Me.SplitContainer1)
        Me.pnlAccountAdj.Controls.Add(Me.StatusStrip2)
        Me.pnlAccountAdj.Location = New System.Drawing.Point(11, 13)
        Me.pnlAccountAdj.Name = "pnlAccountAdj"
        Me.pnlAccountAdj.Size = New System.Drawing.Size(966, 584)
        Me.pnlAccountAdj.TabIndex = 1
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
        Me.SplitContainer1.Panel2.Controls.Add(Me.crtVwFinAccAdjRpt)
        Me.SplitContainer1.Size = New System.Drawing.Size(964, 557)
        Me.SplitContainer1.SplitterDistance = 62
        Me.SplitContainer1.SplitterWidth = 1
        Me.SplitContainer1.TabIndex = 24
        '
        'Panel1
        '
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel1.Controls.Add(Me.btnClose)
        Me.Panel1.Controls.Add(Me.btnLoad)
        Me.Panel1.Controls.Add(Me.dtpDate)
        Me.Panel1.Controls.Add(Me.lblDateLabel)
        Me.Panel1.Controls.Add(Me.rbYearly)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel1.Location = New System.Drawing.Point(5, 5)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(954, 52)
        Me.Panel1.TabIndex = 65
        '
        'btnLoad
        '
        Me.btnLoad.Location = New System.Drawing.Point(473, 13)
        Me.btnLoad.Name = "btnLoad"
        Me.btnLoad.Size = New System.Drawing.Size(75, 23)
        Me.btnLoad.TabIndex = 45
        Me.btnLoad.Text = "Load"
        Me.btnLoad.UseVisualStyleBackColor = True
        '
        'dtpDate
        '
        Me.dtpDate.CustomFormat = "            yyyy"
        Me.dtpDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpDate.Location = New System.Drawing.Point(284, 15)
        Me.dtpDate.Name = "dtpDate"
        Me.dtpDate.ShowUpDown = True
        Me.dtpDate.Size = New System.Drawing.Size(121, 20)
        Me.dtpDate.TabIndex = 5
        Me.dtpDate.Value = New Date(2014, 12, 5, 0, 0, 0, 0)
        '
        'lblDateLabel
        '
        Me.lblDateLabel.AutoSize = True
        Me.lblDateLabel.Location = New System.Drawing.Point(195, 18)
        Me.lblDateLabel.Name = "lblDateLabel"
        Me.lblDateLabel.Size = New System.Drawing.Size(68, 13)
        Me.lblDateLabel.TabIndex = 4
        Me.lblDateLabel.Text = "Select Year :"
        '
        'rbYearly
        '
        Me.rbYearly.AutoSize = True
        Me.rbYearly.Location = New System.Drawing.Point(22, 16)
        Me.rbYearly.Name = "rbYearly"
        Me.rbYearly.Size = New System.Drawing.Size(89, 17)
        Me.rbYearly.TabIndex = 2
        Me.rbYearly.Text = "Yearly Report"
        Me.rbYearly.UseVisualStyleBackColor = True
        '
        'StatusStrip12
        '
        Me.StatusStrip12.AutoSize = False
        Me.StatusStrip12.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip12.Dock = System.Windows.Forms.DockStyle.Left
        Me.StatusStrip12.Location = New System.Drawing.Point(0, 5)
        Me.StatusStrip12.Name = "StatusStrip12"
        Me.StatusStrip12.Size = New System.Drawing.Size(5, 52)
        Me.StatusStrip12.SizingGrip = False
        Me.StatusStrip12.TabIndex = 64
        Me.StatusStrip12.Text = "StatusStrip12"
        '
        'StatusStrip13
        '
        Me.StatusStrip13.AutoSize = False
        Me.StatusStrip13.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip13.Location = New System.Drawing.Point(0, 57)
        Me.StatusStrip13.Name = "StatusStrip13"
        Me.StatusStrip13.Size = New System.Drawing.Size(959, 5)
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
        Me.StatusStrip14.Size = New System.Drawing.Size(959, 5)
        Me.StatusStrip14.SizingGrip = False
        Me.StatusStrip14.TabIndex = 62
        Me.StatusStrip14.Text = "StatusStrip14"
        '
        'StatusStrip5
        '
        Me.StatusStrip5.AutoSize = False
        Me.StatusStrip5.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip5.Dock = System.Windows.Forms.DockStyle.Right
        Me.StatusStrip5.Location = New System.Drawing.Point(959, 0)
        Me.StatusStrip5.Name = "StatusStrip5"
        Me.StatusStrip5.Size = New System.Drawing.Size(5, 62)
        Me.StatusStrip5.SizingGrip = False
        Me.StatusStrip5.TabIndex = 61
        Me.StatusStrip5.Text = "StatusStrip5"
        '
        'crtVwFinAccAdjRpt
        '
        Me.crtVwFinAccAdjRpt.ActiveViewIndex = -1
        Me.crtVwFinAccAdjRpt.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.crtVwFinAccAdjRpt.Cursor = System.Windows.Forms.Cursors.Default
        Me.crtVwFinAccAdjRpt.DisplayStatusBar = False
        Me.crtVwFinAccAdjRpt.Dock = System.Windows.Forms.DockStyle.Fill
        Me.crtVwFinAccAdjRpt.Location = New System.Drawing.Point(0, 0)
        Me.crtVwFinAccAdjRpt.Name = "crtVwFinAccAdjRpt"
        Me.crtVwFinAccAdjRpt.ShowGroupTreeButton = False
        Me.crtVwFinAccAdjRpt.ShowLogo = False
        Me.crtVwFinAccAdjRpt.ShowParameterPanelButton = False
        Me.crtVwFinAccAdjRpt.ShowTextSearchButton = False
        Me.crtVwFinAccAdjRpt.Size = New System.Drawing.Size(964, 494)
        Me.crtVwFinAccAdjRpt.TabIndex = 37
        Me.crtVwFinAccAdjRpt.ToolPanelView = CrystalDecisions.Windows.Forms.ToolPanelViewType.None
        '
        'StatusStrip2
        '
        Me.StatusStrip2.AutoSize = False
        Me.StatusStrip2.BackColor = System.Drawing.Color.LightSeaGreen
        Me.StatusStrip2.Dock = System.Windows.Forms.DockStyle.Top
        Me.StatusStrip2.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripStatusLabel1})
        Me.StatusStrip2.Location = New System.Drawing.Point(0, 0)
        Me.StatusStrip2.Name = "StatusStrip2"
        Me.StatusStrip2.Size = New System.Drawing.Size(964, 25)
        Me.StatusStrip2.SizingGrip = False
        Me.StatusStrip2.TabIndex = 23
        Me.StatusStrip2.Text = "StatusStrip2"
        '
        'ToolStripStatusLabel1
        '
        Me.ToolStripStatusLabel1.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ToolStripStatusLabel1.Name = "ToolStripStatusLabel1"
        Me.ToolStripStatusLabel1.Size = New System.Drawing.Size(286, 20)
        Me.ToolStripStatusLabel1.Text = "FINANCE ACCOUNTS ADJUSTMENTS REPORT"
        '
        'btnClose
        '
        Me.btnClose.Location = New System.Drawing.Point(660, 13)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(75, 23)
        Me.btnClose.TabIndex = 46
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'frmFinRptAccountsAdj
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.LightSkyBlue
        Me.ClientSize = New System.Drawing.Size(989, 610)
        Me.Controls.Add(Me.pnlAccountAdj)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Name = "frmFinRptAccountsAdj"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Accounts Adjustment"
        Me.pnlAccountAdj.ResumeLayout(False)
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
    Friend WithEvents pnlAccountAdj As System.Windows.Forms.Panel
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents btnLoad As System.Windows.Forms.Button
    Friend WithEvents dtpDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents lblDateLabel As System.Windows.Forms.Label
    Friend WithEvents rbYearly As System.Windows.Forms.RadioButton
    Friend WithEvents StatusStrip12 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip13 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip14 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip5 As System.Windows.Forms.StatusStrip
    Friend WithEvents crtVwFinAccAdjRpt As CrystalDecisions.Windows.Forms.CrystalReportViewer
    Friend WithEvents StatusStrip2 As System.Windows.Forms.StatusStrip
    Friend WithEvents ToolStripStatusLabel1 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents btnClose As System.Windows.Forms.Button
End Class
