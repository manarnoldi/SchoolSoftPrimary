<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmInvIssueRequestNote
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
        Me.pnlIssueReqNote = New System.Windows.Forms.Panel()
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.btnClose = New System.Windows.Forms.Button()
        Me.btnLoad = New System.Windows.Forms.Button()
        Me.cboIssueNoteNo = New System.Windows.Forms.ComboBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.cboYear = New System.Windows.Forms.ComboBox()
        Me.cboMonth = New System.Windows.Forms.ComboBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.StatusStrip3 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip4 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip5 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip6 = New System.Windows.Forms.StatusStrip()
        Me.crtVwIssueReqNote = New CrystalDecisions.Windows.Forms.CrystalReportViewer()
        Me.StatusStrip2 = New System.Windows.Forms.StatusStrip()
        Me.ToolStripStatusLabel1 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.pnlIssueReqNote.SuspendLayout()
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.StatusStrip2.SuspendLayout()
        Me.SuspendLayout()
        '
        'pnlIssueReqNote
        '
        Me.pnlIssueReqNote.BackColor = System.Drawing.SystemColors.Control
        Me.pnlIssueReqNote.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pnlIssueReqNote.Controls.Add(Me.SplitContainer1)
        Me.pnlIssueReqNote.Controls.Add(Me.StatusStrip2)
        Me.pnlIssueReqNote.Location = New System.Drawing.Point(12, 12)
        Me.pnlIssueReqNote.Name = "pnlIssueReqNote"
        Me.pnlIssueReqNote.Size = New System.Drawing.Size(1004, 567)
        Me.pnlIssueReqNote.TabIndex = 0
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
        Me.SplitContainer1.Panel1.Controls.Add(Me.Panel1)
        Me.SplitContainer1.Panel1.Controls.Add(Me.StatusStrip3)
        Me.SplitContainer1.Panel1.Controls.Add(Me.StatusStrip4)
        Me.SplitContainer1.Panel1.Controls.Add(Me.StatusStrip5)
        Me.SplitContainer1.Panel1.Controls.Add(Me.StatusStrip6)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.crtVwIssueReqNote)
        Me.SplitContainer1.Size = New System.Drawing.Size(1002, 540)
        Me.SplitContainer1.SplitterDistance = 62
        Me.SplitContainer1.SplitterWidth = 1
        Me.SplitContainer1.TabIndex = 39
        '
        'Panel1
        '
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel1.Controls.Add(Me.btnClose)
        Me.Panel1.Controls.Add(Me.btnLoad)
        Me.Panel1.Controls.Add(Me.cboIssueNoteNo)
        Me.Panel1.Controls.Add(Me.Label3)
        Me.Panel1.Controls.Add(Me.cboYear)
        Me.Panel1.Controls.Add(Me.cboMonth)
        Me.Panel1.Controls.Add(Me.Label2)
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel1.Location = New System.Drawing.Point(5, 5)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(990, 50)
        Me.Panel1.TabIndex = 77
        '
        'btnClose
        '
        Me.btnClose.Location = New System.Drawing.Point(897, 13)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(75, 23)
        Me.btnClose.TabIndex = 52
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'btnLoad
        '
        Me.btnLoad.Location = New System.Drawing.Point(713, 13)
        Me.btnLoad.Name = "btnLoad"
        Me.btnLoad.Size = New System.Drawing.Size(75, 23)
        Me.btnLoad.TabIndex = 51
        Me.btnLoad.Text = "Load"
        Me.btnLoad.UseVisualStyleBackColor = True
        '
        'cboIssueNoteNo
        '
        Me.cboIssueNoteNo.DropDownHeight = 80
        Me.cboIssueNoteNo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboIssueNoteNo.FormattingEnabled = True
        Me.cboIssueNoteNo.IntegralHeight = False
        Me.cboIssueNoteNo.Location = New System.Drawing.Point(477, 15)
        Me.cboIssueNoteNo.Name = "cboIssueNoteNo"
        Me.cboIssueNoteNo.Size = New System.Drawing.Size(155, 21)
        Me.cboIssueNoteNo.TabIndex = 50
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(396, 18)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(75, 13)
        Me.Label3.TabIndex = 49
        Me.Label3.Text = "IssueNote No:"
        '
        'cboYear
        '
        Me.cboYear.DropDownHeight = 80
        Me.cboYear.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboYear.FormattingEnabled = True
        Me.cboYear.IntegralHeight = False
        Me.cboYear.Location = New System.Drawing.Point(280, 15)
        Me.cboYear.Name = "cboYear"
        Me.cboYear.Size = New System.Drawing.Size(96, 21)
        Me.cboYear.TabIndex = 48
        '
        'cboMonth
        '
        Me.cboMonth.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboMonth.FormattingEnabled = True
        Me.cboMonth.IntegralHeight = False
        Me.cboMonth.Items.AddRange(New Object() {"January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"})
        Me.cboMonth.Location = New System.Drawing.Point(73, 15)
        Me.cboMonth.Name = "cboMonth"
        Me.cboMonth.Size = New System.Drawing.Size(153, 21)
        Me.cboMonth.TabIndex = 47
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(242, 18)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(32, 13)
        Me.Label2.TabIndex = 46
        Me.Label2.Text = "Year:"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(21, 18)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(43, 13)
        Me.Label1.TabIndex = 45
        Me.Label1.Text = "Month :"
        '
        'StatusStrip3
        '
        Me.StatusStrip3.AutoSize = False
        Me.StatusStrip3.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip3.Dock = System.Windows.Forms.DockStyle.Right
        Me.StatusStrip3.Location = New System.Drawing.Point(995, 5)
        Me.StatusStrip3.Name = "StatusStrip3"
        Me.StatusStrip3.Size = New System.Drawing.Size(5, 50)
        Me.StatusStrip3.SizingGrip = False
        Me.StatusStrip3.TabIndex = 76
        Me.StatusStrip3.Text = "StatusStrip3"
        '
        'StatusStrip4
        '
        Me.StatusStrip4.AutoSize = False
        Me.StatusStrip4.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip4.Dock = System.Windows.Forms.DockStyle.Left
        Me.StatusStrip4.Location = New System.Drawing.Point(0, 5)
        Me.StatusStrip4.Name = "StatusStrip4"
        Me.StatusStrip4.Size = New System.Drawing.Size(5, 50)
        Me.StatusStrip4.SizingGrip = False
        Me.StatusStrip4.TabIndex = 75
        Me.StatusStrip4.Text = "StatusStrip4"
        '
        'StatusStrip5
        '
        Me.StatusStrip5.AutoSize = False
        Me.StatusStrip5.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip5.Location = New System.Drawing.Point(0, 55)
        Me.StatusStrip5.Name = "StatusStrip5"
        Me.StatusStrip5.Size = New System.Drawing.Size(1000, 5)
        Me.StatusStrip5.SizingGrip = False
        Me.StatusStrip5.TabIndex = 74
        Me.StatusStrip5.Text = "StatusStrip5"
        '
        'StatusStrip6
        '
        Me.StatusStrip6.AutoSize = False
        Me.StatusStrip6.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip6.Dock = System.Windows.Forms.DockStyle.Top
        Me.StatusStrip6.Location = New System.Drawing.Point(0, 0)
        Me.StatusStrip6.Name = "StatusStrip6"
        Me.StatusStrip6.Size = New System.Drawing.Size(1000, 5)
        Me.StatusStrip6.SizingGrip = False
        Me.StatusStrip6.TabIndex = 73
        Me.StatusStrip6.Text = "StatusStrip6"
        '
        'crtVwIssueReqNote
        '
        Me.crtVwIssueReqNote.ActiveViewIndex = -1
        Me.crtVwIssueReqNote.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.crtVwIssueReqNote.Cursor = System.Windows.Forms.Cursors.Default
        Me.crtVwIssueReqNote.DisplayStatusBar = False
        Me.crtVwIssueReqNote.Dock = System.Windows.Forms.DockStyle.Fill
        Me.crtVwIssueReqNote.Location = New System.Drawing.Point(0, 0)
        Me.crtVwIssueReqNote.Name = "crtVwIssueReqNote"
        Me.crtVwIssueReqNote.ShowGroupTreeButton = False
        Me.crtVwIssueReqNote.ShowLogo = False
        Me.crtVwIssueReqNote.ShowParameterPanelButton = False
        Me.crtVwIssueReqNote.ShowTextSearchButton = False
        Me.crtVwIssueReqNote.Size = New System.Drawing.Size(1000, 475)
        Me.crtVwIssueReqNote.TabIndex = 37
        Me.crtVwIssueReqNote.ToolPanelView = CrystalDecisions.Windows.Forms.ToolPanelViewType.None
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
        Me.StatusStrip2.TabIndex = 38
        Me.StatusStrip2.Text = "StatusStrip2"
        '
        'ToolStripStatusLabel1
        '
        Me.ToolStripStatusLabel1.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ToolStripStatusLabel1.Name = "ToolStripStatusLabel1"
        Me.ToolStripStatusLabel1.Size = New System.Drawing.Size(218, 20)
        Me.ToolStripStatusLabel1.Text = "INVENTORY ISSUE REQUEST NOTE"
        '
        'frmInvIssueRequestNote
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.LightSkyBlue
        Me.ClientSize = New System.Drawing.Size(1028, 591)
        Me.Controls.Add(Me.pnlIssueReqNote)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Name = "frmInvIssueRequestNote"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Issue Note"
        Me.pnlIssueReqNote.ResumeLayout(False)
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
    Friend WithEvents pnlIssueReqNote As System.Windows.Forms.Panel
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents StatusStrip2 As System.Windows.Forms.StatusStrip
    Friend WithEvents ToolStripStatusLabel1 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents StatusStrip3 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip4 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip5 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip6 As System.Windows.Forms.StatusStrip
    Friend WithEvents crtVwIssueReqNote As CrystalDecisions.Windows.Forms.CrystalReportViewer
    Friend WithEvents btnLoad As System.Windows.Forms.Button
    Friend WithEvents cboIssueNoteNo As System.Windows.Forms.ComboBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents cboYear As System.Windows.Forms.ComboBox
    Friend WithEvents cboMonth As System.Windows.Forms.ComboBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents btnClose As System.Windows.Forms.Button
End Class
