<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmInvRptStockAtHand
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
        Me.pnlInvStockAtHandRpt = New System.Windows.Forms.Panel()
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.btnClose = New System.Windows.Forms.Button()
        Me.StatusStrip12 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip13 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip14 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip5 = New System.Windows.Forms.StatusStrip()
        Me.crtVwRptInvStockAtHand = New CrystalDecisions.Windows.Forms.CrystalReportViewer()
        Me.StatusStrip2 = New System.Windows.Forms.StatusStrip()
        Me.ToolStripStatusLabel1 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.btnLoad = New System.Windows.Forms.Button()
        Me.pnlInvStockAtHandRpt.SuspendLayout()
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.StatusStrip2.SuspendLayout()
        Me.SuspendLayout()
        '
        'pnlInvStockAtHandRpt
        '
        Me.pnlInvStockAtHandRpt.BackColor = System.Drawing.SystemColors.Control
        Me.pnlInvStockAtHandRpt.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pnlInvStockAtHandRpt.Controls.Add(Me.SplitContainer1)
        Me.pnlInvStockAtHandRpt.Controls.Add(Me.StatusStrip2)
        Me.pnlInvStockAtHandRpt.Location = New System.Drawing.Point(12, 12)
        Me.pnlInvStockAtHandRpt.Name = "pnlInvStockAtHandRpt"
        Me.pnlInvStockAtHandRpt.Size = New System.Drawing.Size(999, 564)
        Me.pnlInvStockAtHandRpt.TabIndex = 4
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
        Me.SplitContainer1.Panel2.Controls.Add(Me.crtVwRptInvStockAtHand)
        Me.SplitContainer1.Size = New System.Drawing.Size(997, 537)
        Me.SplitContainer1.SplitterDistance = 51
        Me.SplitContainer1.SplitterWidth = 1
        Me.SplitContainer1.TabIndex = 25
        '
        'Panel1
        '
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel1.Controls.Add(Me.btnLoad)
        Me.Panel1.Controls.Add(Me.btnClose)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel1.Location = New System.Drawing.Point(5, 5)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(987, 41)
        Me.Panel1.TabIndex = 65
        '
        'btnClose
        '
        Me.btnClose.Location = New System.Drawing.Point(887, 8)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(75, 23)
        Me.btnClose.TabIndex = 46
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = True
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
        'crtVwRptInvStockAtHand
        '
        Me.crtVwRptInvStockAtHand.ActiveViewIndex = -1
        Me.crtVwRptInvStockAtHand.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.crtVwRptInvStockAtHand.Cursor = System.Windows.Forms.Cursors.Default
        Me.crtVwRptInvStockAtHand.DisplayStatusBar = False
        Me.crtVwRptInvStockAtHand.Dock = System.Windows.Forms.DockStyle.Fill
        Me.crtVwRptInvStockAtHand.Location = New System.Drawing.Point(0, 0)
        Me.crtVwRptInvStockAtHand.Name = "crtVwRptInvStockAtHand"
        Me.crtVwRptInvStockAtHand.ShowGroupTreeButton = False
        Me.crtVwRptInvStockAtHand.ShowLogo = False
        Me.crtVwRptInvStockAtHand.ShowParameterPanelButton = False
        Me.crtVwRptInvStockAtHand.ShowTextSearchButton = False
        Me.crtVwRptInvStockAtHand.Size = New System.Drawing.Size(997, 485)
        Me.crtVwRptInvStockAtHand.TabIndex = 37
        Me.crtVwRptInvStockAtHand.ToolPanelView = CrystalDecisions.Windows.Forms.ToolPanelViewType.None
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
        Me.ToolStripStatusLabel1.Size = New System.Drawing.Size(165, 20)
        Me.ToolStripStatusLabel1.Text = "STOCK AT HAND REPORT"
        '
        'btnLoad
        '
        Me.btnLoad.Location = New System.Drawing.Point(618, 8)
        Me.btnLoad.Name = "btnLoad"
        Me.btnLoad.Size = New System.Drawing.Size(75, 23)
        Me.btnLoad.TabIndex = 47
        Me.btnLoad.Text = "Load"
        Me.btnLoad.UseVisualStyleBackColor = True
        '
        'frmInvRptStockAtHand
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.LightSkyBlue
        Me.ClientSize = New System.Drawing.Size(1023, 588)
        Me.Controls.Add(Me.pnlInvStockAtHandRpt)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Name = "frmInvRptStockAtHand"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Stock At Hand"
        Me.pnlInvStockAtHandRpt.ResumeLayout(False)
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.ResumeLayout(False)
        Me.Panel1.ResumeLayout(False)
        Me.StatusStrip2.ResumeLayout(False)
        Me.StatusStrip2.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents pnlInvStockAtHandRpt As System.Windows.Forms.Panel
    Friend WithEvents StatusStrip2 As System.Windows.Forms.StatusStrip
    Friend WithEvents ToolStripStatusLabel1 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents StatusStrip12 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip13 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip14 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip5 As System.Windows.Forms.StatusStrip
    Friend WithEvents crtVwRptInvStockAtHand As CrystalDecisions.Windows.Forms.CrystalReportViewer
    Friend WithEvents btnLoad As System.Windows.Forms.Button
End Class
