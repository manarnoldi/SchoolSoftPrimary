<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmResultsViewing
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
        Me.pnlResultsViewing = New System.Windows.Forms.Panel()
        Me.crtViewResultsSummary = New CrystalDecisions.Windows.Forms.CrystalReportViewer()
        Me.StatusStrip2 = New System.Windows.Forms.StatusStrip()
        Me.ToolStripStatusLabel1 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip()
        Me.pnlResultsViewing.SuspendLayout()
        Me.StatusStrip2.SuspendLayout()
        Me.SuspendLayout()
        '
        'pnlResultsViewing
        '
        Me.pnlResultsViewing.BackColor = System.Drawing.SystemColors.Control
        Me.pnlResultsViewing.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pnlResultsViewing.Controls.Add(Me.crtViewResultsSummary)
        Me.pnlResultsViewing.Controls.Add(Me.StatusStrip2)
        Me.pnlResultsViewing.Controls.Add(Me.StatusStrip1)
        Me.pnlResultsViewing.Location = New System.Drawing.Point(12, 12)
        Me.pnlResultsViewing.Name = "pnlResultsViewing"
        Me.pnlResultsViewing.Size = New System.Drawing.Size(1013, 584)
        Me.pnlResultsViewing.TabIndex = 0
        '
        'crtViewResultsSummary
        '
        Me.crtViewResultsSummary.ActiveViewIndex = -1
        Me.crtViewResultsSummary.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.crtViewResultsSummary.Cursor = System.Windows.Forms.Cursors.Default
        Me.crtViewResultsSummary.DisplayStatusBar = False
        Me.crtViewResultsSummary.Dock = System.Windows.Forms.DockStyle.Fill
        Me.crtViewResultsSummary.Location = New System.Drawing.Point(0, 25)
        Me.crtViewResultsSummary.Name = "crtViewResultsSummary"
        Me.crtViewResultsSummary.ShowGroupTreeButton = False
        Me.crtViewResultsSummary.ShowLogo = False
        Me.crtViewResultsSummary.ShowParameterPanelButton = False
        Me.crtViewResultsSummary.ShowTextSearchButton = False
        Me.crtViewResultsSummary.Size = New System.Drawing.Size(1011, 532)
        Me.crtViewResultsSummary.TabIndex = 35
        Me.crtViewResultsSummary.ToolPanelView = CrystalDecisions.Windows.Forms.ToolPanelViewType.None
        '
        'StatusStrip2
        '
        Me.StatusStrip2.AutoSize = False
        Me.StatusStrip2.BackColor = System.Drawing.Color.LightSeaGreen
        Me.StatusStrip2.Dock = System.Windows.Forms.DockStyle.Top
        Me.StatusStrip2.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripStatusLabel1})
        Me.StatusStrip2.Location = New System.Drawing.Point(0, 0)
        Me.StatusStrip2.Name = "StatusStrip2"
        Me.StatusStrip2.Size = New System.Drawing.Size(1011, 25)
        Me.StatusStrip2.SizingGrip = False
        Me.StatusStrip2.TabIndex = 34
        Me.StatusStrip2.Text = "StatusStrip2"
        '
        'ToolStripStatusLabel1
        '
        Me.ToolStripStatusLabel1.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ToolStripStatusLabel1.Name = "ToolStripStatusLabel1"
        Me.ToolStripStatusLabel1.Size = New System.Drawing.Size(170, 20)
        Me.ToolStripStatusLabel1.Text = "EXAM RESULTS SUMMARY"
        '
        'StatusStrip1
        '
        Me.StatusStrip1.AutoSize = False
        Me.StatusStrip1.BackColor = System.Drawing.Color.LightSeaGreen
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 557)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(1011, 25)
        Me.StatusStrip1.SizingGrip = False
        Me.StatusStrip1.TabIndex = 33
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'frmResultsViewing
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.LightSkyBlue
        Me.ClientSize = New System.Drawing.Size(1037, 608)
        Me.Controls.Add(Me.pnlResultsViewing)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Name = "frmResultsViewing"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Results Viewing"
        Me.pnlResultsViewing.ResumeLayout(False)
        Me.StatusStrip2.ResumeLayout(False)
        Me.StatusStrip2.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents pnlResultsViewing As System.Windows.Forms.Panel
    Friend WithEvents crtViewResultsSummary As CrystalDecisions.Windows.Forms.CrystalReportViewer
    Friend WithEvents StatusStrip2 As System.Windows.Forms.StatusStrip
    Friend WithEvents ToolStripStatusLabel1 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
End Class
