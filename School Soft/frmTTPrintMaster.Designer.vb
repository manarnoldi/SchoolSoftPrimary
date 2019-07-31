<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmTTPrintMaster
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
        Me.pnlPrintMaster = New System.Windows.Forms.Panel()
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.btnPrint = New System.Windows.Forms.Button()
        Me.cboYear = New System.Windows.Forms.ComboBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.btnClose = New System.Windows.Forms.Button()
        Me.btnPreview = New System.Windows.Forms.Button()
        Me.StatusStrip6 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip7 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip8 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip9 = New System.Windows.Forms.StatusStrip()
        Me.crtVwPrintMasterCopy = New CrystalDecisions.Windows.Forms.CrystalReportViewer()
        Me.StatusStrip2 = New System.Windows.Forms.StatusStrip()
        Me.ToolStripStatusLabel1 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip()
        Me.pnlPrintMaster.SuspendLayout()
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.StatusStrip2.SuspendLayout()
        Me.SuspendLayout()
        '
        'pnlPrintMaster
        '
        Me.pnlPrintMaster.BackColor = System.Drawing.SystemColors.Control
        Me.pnlPrintMaster.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pnlPrintMaster.Controls.Add(Me.SplitContainer1)
        Me.pnlPrintMaster.Controls.Add(Me.StatusStrip2)
        Me.pnlPrintMaster.Controls.Add(Me.StatusStrip1)
        Me.pnlPrintMaster.Location = New System.Drawing.Point(12, 12)
        Me.pnlPrintMaster.Name = "pnlPrintMaster"
        Me.pnlPrintMaster.Size = New System.Drawing.Size(1019, 609)
        Me.pnlPrintMaster.TabIndex = 1
        '
        'SplitContainer1
        '
        Me.SplitContainer1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.SplitContainer1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer1.Location = New System.Drawing.Point(0, 25)
        Me.SplitContainer1.Name = "SplitContainer1"
        Me.SplitContainer1.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.Panel2)
        Me.SplitContainer1.Panel1.Controls.Add(Me.StatusStrip6)
        Me.SplitContainer1.Panel1.Controls.Add(Me.StatusStrip7)
        Me.SplitContainer1.Panel1.Controls.Add(Me.StatusStrip8)
        Me.SplitContainer1.Panel1.Controls.Add(Me.StatusStrip9)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.crtVwPrintMasterCopy)
        Me.SplitContainer1.Size = New System.Drawing.Size(1017, 557)
        Me.SplitContainer1.SplitterDistance = 67
        Me.SplitContainer1.SplitterWidth = 1
        Me.SplitContainer1.TabIndex = 21
        '
        'Panel2
        '
        Me.Panel2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel2.Controls.Add(Me.btnPrint)
        Me.Panel2.Controls.Add(Me.cboYear)
        Me.Panel2.Controls.Add(Me.Label3)
        Me.Panel2.Controls.Add(Me.btnClose)
        Me.Panel2.Controls.Add(Me.btnPreview)
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel2.Location = New System.Drawing.Point(5, 5)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(1005, 55)
        Me.Panel2.TabIndex = 59
        '
        'btnPrint
        '
        Me.btnPrint.Location = New System.Drawing.Point(688, 8)
        Me.btnPrint.Name = "btnPrint"
        Me.btnPrint.Size = New System.Drawing.Size(75, 23)
        Me.btnPrint.TabIndex = 10
        Me.btnPrint.Text = "Print"
        Me.btnPrint.UseVisualStyleBackColor = True
        '
        'cboYear
        '
        Me.cboYear.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboYear.FormattingEnabled = True
        Me.cboYear.Location = New System.Drawing.Point(52, 10)
        Me.cboYear.Name = "cboYear"
        Me.cboYear.Size = New System.Drawing.Size(101, 21)
        Me.cboYear.TabIndex = 9
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(14, 13)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(32, 13)
        Me.Label3.TabIndex = 8
        Me.Label3.Text = "Year:"
        '
        'btnClose
        '
        Me.btnClose.Location = New System.Drawing.Point(856, 8)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(75, 23)
        Me.btnClose.TabIndex = 3
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'btnPreview
        '
        Me.btnPreview.Location = New System.Drawing.Point(524, 8)
        Me.btnPreview.Name = "btnPreview"
        Me.btnPreview.Size = New System.Drawing.Size(75, 23)
        Me.btnPreview.TabIndex = 2
        Me.btnPreview.Text = "Preview"
        Me.btnPreview.UseVisualStyleBackColor = True
        '
        'StatusStrip6
        '
        Me.StatusStrip6.AutoSize = False
        Me.StatusStrip6.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip6.Dock = System.Windows.Forms.DockStyle.Right
        Me.StatusStrip6.Location = New System.Drawing.Point(1010, 5)
        Me.StatusStrip6.Name = "StatusStrip6"
        Me.StatusStrip6.Size = New System.Drawing.Size(5, 55)
        Me.StatusStrip6.SizingGrip = False
        Me.StatusStrip6.TabIndex = 58
        Me.StatusStrip6.Text = "StatusStrip6"
        '
        'StatusStrip7
        '
        Me.StatusStrip7.AutoSize = False
        Me.StatusStrip7.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip7.Dock = System.Windows.Forms.DockStyle.Left
        Me.StatusStrip7.Location = New System.Drawing.Point(0, 5)
        Me.StatusStrip7.Name = "StatusStrip7"
        Me.StatusStrip7.Size = New System.Drawing.Size(5, 55)
        Me.StatusStrip7.SizingGrip = False
        Me.StatusStrip7.TabIndex = 57
        Me.StatusStrip7.Text = "StatusStrip7"
        '
        'StatusStrip8
        '
        Me.StatusStrip8.AutoSize = False
        Me.StatusStrip8.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip8.Location = New System.Drawing.Point(0, 60)
        Me.StatusStrip8.Name = "StatusStrip8"
        Me.StatusStrip8.Size = New System.Drawing.Size(1015, 5)
        Me.StatusStrip8.SizingGrip = False
        Me.StatusStrip8.TabIndex = 56
        Me.StatusStrip8.Text = "StatusStrip8"
        '
        'StatusStrip9
        '
        Me.StatusStrip9.AutoSize = False
        Me.StatusStrip9.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip9.Dock = System.Windows.Forms.DockStyle.Top
        Me.StatusStrip9.Location = New System.Drawing.Point(0, 0)
        Me.StatusStrip9.Name = "StatusStrip9"
        Me.StatusStrip9.Size = New System.Drawing.Size(1015, 5)
        Me.StatusStrip9.SizingGrip = False
        Me.StatusStrip9.TabIndex = 55
        Me.StatusStrip9.Text = "StatusStrip9"
        '
        'crtVwPrintMasterCopy
        '
        Me.crtVwPrintMasterCopy.ActiveViewIndex = -1
        Me.crtVwPrintMasterCopy.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.crtVwPrintMasterCopy.Cursor = System.Windows.Forms.Cursors.Default
        Me.crtVwPrintMasterCopy.DisplayStatusBar = False
        Me.crtVwPrintMasterCopy.Dock = System.Windows.Forms.DockStyle.Fill
        Me.crtVwPrintMasterCopy.Location = New System.Drawing.Point(0, 0)
        Me.crtVwPrintMasterCopy.Name = "crtVwPrintMasterCopy"
        Me.crtVwPrintMasterCopy.ShowGroupTreeButton = False
        Me.crtVwPrintMasterCopy.ShowLogo = False
        Me.crtVwPrintMasterCopy.ShowParameterPanelButton = False
        Me.crtVwPrintMasterCopy.ShowTextSearchButton = False
        Me.crtVwPrintMasterCopy.Size = New System.Drawing.Size(1015, 487)
        Me.crtVwPrintMasterCopy.TabIndex = 36
        Me.crtVwPrintMasterCopy.ToolPanelView = CrystalDecisions.Windows.Forms.ToolPanelViewType.None
        '
        'StatusStrip2
        '
        Me.StatusStrip2.AutoSize = False
        Me.StatusStrip2.BackColor = System.Drawing.Color.LightSeaGreen
        Me.StatusStrip2.Dock = System.Windows.Forms.DockStyle.Top
        Me.StatusStrip2.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripStatusLabel1})
        Me.StatusStrip2.Location = New System.Drawing.Point(0, 0)
        Me.StatusStrip2.Name = "StatusStrip2"
        Me.StatusStrip2.Size = New System.Drawing.Size(1017, 25)
        Me.StatusStrip2.SizingGrip = False
        Me.StatusStrip2.TabIndex = 20
        Me.StatusStrip2.Text = "StatusStrip2"
        '
        'ToolStripStatusLabel1
        '
        Me.ToolStripStatusLabel1.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ToolStripStatusLabel1.Name = "ToolStripStatusLabel1"
        Me.ToolStripStatusLabel1.Size = New System.Drawing.Size(212, 20)
        Me.ToolStripStatusLabel1.Text = "TIMETABLE PRINT MASTER COPY"
        '
        'StatusStrip1
        '
        Me.StatusStrip1.AutoSize = False
        Me.StatusStrip1.BackColor = System.Drawing.Color.LightSeaGreen
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 582)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(1017, 25)
        Me.StatusStrip1.SizingGrip = False
        Me.StatusStrip1.TabIndex = 19
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'frmTTPrintMaster
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.LightSkyBlue
        Me.ClientSize = New System.Drawing.Size(1043, 633)
        Me.Controls.Add(Me.pnlPrintMaster)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Name = "frmTTPrintMaster"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "frmTTPrintMaster"
        Me.pnlPrintMaster.ResumeLayout(False)
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.ResumeLayout(False)
        Me.Panel2.ResumeLayout(False)
        Me.Panel2.PerformLayout()
        Me.StatusStrip2.ResumeLayout(False)
        Me.StatusStrip2.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents pnlPrintMaster As System.Windows.Forms.Panel
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents btnPrint As System.Windows.Forms.Button
    Friend WithEvents cboYear As System.Windows.Forms.ComboBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents btnPreview As System.Windows.Forms.Button
    Friend WithEvents StatusStrip6 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip7 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip8 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip9 As System.Windows.Forms.StatusStrip
    Friend WithEvents crtVwPrintMasterCopy As CrystalDecisions.Windows.Forms.CrystalReportViewer
    Friend WithEvents StatusStrip2 As System.Windows.Forms.StatusStrip
    Friend WithEvents ToolStripStatusLabel1 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
End Class
