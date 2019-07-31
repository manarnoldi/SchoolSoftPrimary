<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmRptstaffsubj
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
        Me.pnlClassLists = New System.Windows.Forms.Panel()
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.btnClose = New System.Windows.Forms.Button()
        Me.btnPrint = New System.Windows.Forms.Button()
        Me.btnLoad = New System.Windows.Forms.Button()
        Me.cboYear = New System.Windows.Forms.ComboBox()
        Me.cboClassName = New System.Windows.Forms.ComboBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.StatusStrip5 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip4 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip3 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip10 = New System.Windows.Forms.StatusStrip()
        Me.crtVwClassLists = New CrystalDecisions.Windows.Forms.CrystalReportViewer()
        Me.StatusStrip2 = New System.Windows.Forms.StatusStrip()
        Me.ToolStripStatusLabel1 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.btnExport = New System.Windows.Forms.Button()
        Me.pnlClassLists.SuspendLayout()
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.StatusStrip2.SuspendLayout()
        Me.SuspendLayout()
        '
        'pnlClassLists
        '
        Me.pnlClassLists.BackColor = System.Drawing.SystemColors.Control
        Me.pnlClassLists.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pnlClassLists.Controls.Add(Me.SplitContainer1)
        Me.pnlClassLists.Controls.Add(Me.StatusStrip2)
        Me.pnlClassLists.Location = New System.Drawing.Point(10, 11)
        Me.pnlClassLists.Name = "pnlClassLists"
        Me.pnlClassLists.Size = New System.Drawing.Size(1004, 587)
        Me.pnlClassLists.TabIndex = 2
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
        Me.SplitContainer1.Panel1.Controls.Add(Me.Panel2)
        Me.SplitContainer1.Panel1.Controls.Add(Me.StatusStrip5)
        Me.SplitContainer1.Panel1.Controls.Add(Me.StatusStrip4)
        Me.SplitContainer1.Panel1.Controls.Add(Me.StatusStrip3)
        Me.SplitContainer1.Panel1.Controls.Add(Me.StatusStrip10)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.crtVwClassLists)
        Me.SplitContainer1.Size = New System.Drawing.Size(1002, 560)
        Me.SplitContainer1.SplitterDistance = 47
        Me.SplitContainer1.SplitterWidth = 1
        Me.SplitContainer1.TabIndex = 32
        '
        'Panel2
        '
        Me.Panel2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel2.Controls.Add(Me.btnExport)
        Me.Panel2.Controls.Add(Me.btnClose)
        Me.Panel2.Controls.Add(Me.btnPrint)
        Me.Panel2.Controls.Add(Me.btnLoad)
        Me.Panel2.Controls.Add(Me.cboYear)
        Me.Panel2.Controls.Add(Me.cboClassName)
        Me.Panel2.Controls.Add(Me.Label3)
        Me.Panel2.Controls.Add(Me.Label1)
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel2.Location = New System.Drawing.Point(5, 5)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(992, 37)
        Me.Panel2.TabIndex = 42
        '
        'btnClose
        '
        Me.btnClose.Location = New System.Drawing.Point(908, 5)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(75, 23)
        Me.btnClose.TabIndex = 8
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'btnPrint
        '
        Me.btnPrint.Location = New System.Drawing.Point(633, 5)
        Me.btnPrint.Name = "btnPrint"
        Me.btnPrint.Size = New System.Drawing.Size(75, 23)
        Me.btnPrint.TabIndex = 7
        Me.btnPrint.Text = "Print"
        Me.btnPrint.UseVisualStyleBackColor = True
        '
        'btnLoad
        '
        Me.btnLoad.Location = New System.Drawing.Point(500, 5)
        Me.btnLoad.Name = "btnLoad"
        Me.btnLoad.Size = New System.Drawing.Size(75, 23)
        Me.btnLoad.TabIndex = 6
        Me.btnLoad.Text = "Preview"
        Me.btnLoad.UseVisualStyleBackColor = True
        '
        'cboYear
        '
        Me.cboYear.DropDownHeight = 70
        Me.cboYear.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboYear.FormattingEnabled = True
        Me.cboYear.IntegralHeight = False
        Me.cboYear.Location = New System.Drawing.Point(316, 7)
        Me.cboYear.Name = "cboYear"
        Me.cboYear.Size = New System.Drawing.Size(99, 21)
        Me.cboYear.TabIndex = 5
        '
        'cboClassName
        '
        Me.cboClassName.DropDownHeight = 70
        Me.cboClassName.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboClassName.FormattingEnabled = True
        Me.cboClassName.IntegralHeight = False
        Me.cboClassName.Location = New System.Drawing.Point(85, 7)
        Me.cboClassName.Name = "cboClassName"
        Me.cboClassName.Size = New System.Drawing.Size(121, 21)
        Me.cboClassName.TabIndex = 3
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(266, 10)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(35, 13)
        Me.Label3.TabIndex = 2
        Me.Label3.Text = "Year :"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(10, 10)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(69, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Class Name :"
        '
        'StatusStrip5
        '
        Me.StatusStrip5.AutoSize = False
        Me.StatusStrip5.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip5.Dock = System.Windows.Forms.DockStyle.Right
        Me.StatusStrip5.Location = New System.Drawing.Point(997, 5)
        Me.StatusStrip5.Name = "StatusStrip5"
        Me.StatusStrip5.Size = New System.Drawing.Size(5, 37)
        Me.StatusStrip5.SizingGrip = False
        Me.StatusStrip5.TabIndex = 41
        Me.StatusStrip5.Text = "StatusStrip5"
        '
        'StatusStrip4
        '
        Me.StatusStrip4.AutoSize = False
        Me.StatusStrip4.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip4.Dock = System.Windows.Forms.DockStyle.Left
        Me.StatusStrip4.Location = New System.Drawing.Point(0, 5)
        Me.StatusStrip4.Name = "StatusStrip4"
        Me.StatusStrip4.Size = New System.Drawing.Size(5, 37)
        Me.StatusStrip4.SizingGrip = False
        Me.StatusStrip4.TabIndex = 40
        Me.StatusStrip4.Text = "StatusStrip4"
        '
        'StatusStrip3
        '
        Me.StatusStrip3.AutoSize = False
        Me.StatusStrip3.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip3.Location = New System.Drawing.Point(0, 42)
        Me.StatusStrip3.Name = "StatusStrip3"
        Me.StatusStrip3.Size = New System.Drawing.Size(1002, 5)
        Me.StatusStrip3.SizingGrip = False
        Me.StatusStrip3.TabIndex = 39
        Me.StatusStrip3.Text = "StatusStrip3"
        '
        'StatusStrip10
        '
        Me.StatusStrip10.AutoSize = False
        Me.StatusStrip10.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip10.Dock = System.Windows.Forms.DockStyle.Top
        Me.StatusStrip10.Location = New System.Drawing.Point(0, 0)
        Me.StatusStrip10.Name = "StatusStrip10"
        Me.StatusStrip10.Size = New System.Drawing.Size(1002, 5)
        Me.StatusStrip10.SizingGrip = False
        Me.StatusStrip10.TabIndex = 38
        Me.StatusStrip10.Text = "StatusStrip10"
        '
        'crtVwClassLists
        '
        Me.crtVwClassLists.ActiveViewIndex = -1
        Me.crtVwClassLists.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.crtVwClassLists.Cursor = System.Windows.Forms.Cursors.Default
        Me.crtVwClassLists.DisplayStatusBar = False
        Me.crtVwClassLists.Dock = System.Windows.Forms.DockStyle.Fill
        Me.crtVwClassLists.Location = New System.Drawing.Point(0, 0)
        Me.crtVwClassLists.Name = "crtVwClassLists"
        Me.crtVwClassLists.ShowGroupTreeButton = False
        Me.crtVwClassLists.ShowLogo = False
        Me.crtVwClassLists.ShowParameterPanelButton = False
        Me.crtVwClassLists.ShowTextSearchButton = False
        Me.crtVwClassLists.Size = New System.Drawing.Size(1002, 512)
        Me.crtVwClassLists.TabIndex = 36
        Me.crtVwClassLists.ToolPanelView = CrystalDecisions.Windows.Forms.ToolPanelViewType.None
        '
        'StatusStrip2
        '
        Me.StatusStrip2.AutoSize = False
        Me.StatusStrip2.BackColor = System.Drawing.Color.LightSeaGreen
        Me.StatusStrip2.Dock = System.Windows.Forms.DockStyle.Top
        Me.StatusStrip2.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripStatusLabel1})
        Me.StatusStrip2.Location = New System.Drawing.Point(0, 0)
        Me.StatusStrip2.Name = "StatusStrip2"
        Me.StatusStrip2.Size = New System.Drawing.Size(1002, 25)
        Me.StatusStrip2.SizingGrip = False
        Me.StatusStrip2.TabIndex = 30
        Me.StatusStrip2.Text = "StatusStrip2"
        '
        'ToolStripStatusLabel1
        '
        Me.ToolStripStatusLabel1.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ToolStripStatusLabel1.Name = "ToolStripStatusLabel1"
        Me.ToolStripStatusLabel1.Size = New System.Drawing.Size(194, 20)
        Me.ToolStripStatusLabel1.Text = "STAFF SUBJECT ALLOCATIONS"
        '
        'btnExport
        '
        Me.btnExport.Location = New System.Drawing.Point(769, 5)
        Me.btnExport.Name = "btnExport"
        Me.btnExport.Size = New System.Drawing.Size(75, 23)
        Me.btnExport.TabIndex = 9
        Me.btnExport.Text = "Export"
        Me.btnExport.UseVisualStyleBackColor = True
        '
        'frmRptstaffsubj
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.LightSkyBlue
        Me.ClientSize = New System.Drawing.Size(1026, 610)
        Me.Controls.Add(Me.pnlClassLists)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Name = "frmRptstaffsubj"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Staff Subjects"
        Me.pnlClassLists.ResumeLayout(False)
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
    Friend WithEvents pnlClassLists As System.Windows.Forms.Panel
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents btnPrint As System.Windows.Forms.Button
    Friend WithEvents btnLoad As System.Windows.Forms.Button
    Friend WithEvents cboYear As System.Windows.Forms.ComboBox
    Friend WithEvents cboClassName As System.Windows.Forms.ComboBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents StatusStrip5 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip4 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip3 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip10 As System.Windows.Forms.StatusStrip
    Friend WithEvents crtVwClassLists As CrystalDecisions.Windows.Forms.CrystalReportViewer
    Friend WithEvents StatusStrip2 As System.Windows.Forms.StatusStrip
    Friend WithEvents ToolStripStatusLabel1 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents btnExport As System.Windows.Forms.Button
End Class
