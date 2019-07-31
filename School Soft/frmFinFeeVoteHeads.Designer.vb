<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmFinFeeVoteHeads
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
        Me.components = New System.ComponentModel.Container()
        Me.pnlFeeVoteHeads = New System.Windows.Forms.Panel()
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.numVotePriority = New System.Windows.Forms.NumericUpDown()
        Me.btnUpdate = New System.Windows.Forms.Button()
        Me.btnSave = New System.Windows.Forms.Button()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.cboVoteHeadName = New System.Windows.Forms.ComboBox()
        Me.StatusStrip12 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip13 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip14 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip5 = New System.Windows.Forms.StatusStrip()
        Me.SplitContainer2 = New System.Windows.Forms.SplitContainer()
        Me.lstVoteHeads = New System.Windows.Forms.ListView()
        Me.ColumnHeader2 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader3 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ContextMenuStrip1 = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.UPDATEToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.DELETEToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.CLOSEToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.btnView = New System.Windows.Forms.Button()
        Me.btnClose = New System.Windows.Forms.Button()
        Me.StatusStrip3 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip4 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip6 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip7 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip2 = New System.Windows.Forms.StatusStrip()
        Me.ToolStripStatusLabel1 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip()
        Me.pnlFeeVoteHeads.SuspendLayout()
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.Panel2.SuspendLayout()
        CType(Me.numVotePriority, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.SplitContainer2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer2.Panel1.SuspendLayout()
        Me.SplitContainer2.Panel2.SuspendLayout()
        Me.SplitContainer2.SuspendLayout()
        Me.ContextMenuStrip1.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.StatusStrip2.SuspendLayout()
        Me.SuspendLayout()
        '
        'pnlFeeVoteHeads
        '
        Me.pnlFeeVoteHeads.BackColor = System.Drawing.SystemColors.Control
        Me.pnlFeeVoteHeads.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pnlFeeVoteHeads.Controls.Add(Me.SplitContainer1)
        Me.pnlFeeVoteHeads.Controls.Add(Me.StatusStrip2)
        Me.pnlFeeVoteHeads.Controls.Add(Me.StatusStrip1)
        Me.pnlFeeVoteHeads.Location = New System.Drawing.Point(11, 12)
        Me.pnlFeeVoteHeads.Name = "pnlFeeVoteHeads"
        Me.pnlFeeVoteHeads.Size = New System.Drawing.Size(417, 485)
        Me.pnlFeeVoteHeads.TabIndex = 5
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
        Me.SplitContainer1.Panel1.Controls.Add(Me.StatusStrip12)
        Me.SplitContainer1.Panel1.Controls.Add(Me.StatusStrip13)
        Me.SplitContainer1.Panel1.Controls.Add(Me.StatusStrip14)
        Me.SplitContainer1.Panel1.Controls.Add(Me.StatusStrip5)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.SplitContainer2)
        Me.SplitContainer1.Size = New System.Drawing.Size(415, 433)
        Me.SplitContainer1.SplitterDistance = 125
        Me.SplitContainer1.SplitterWidth = 1
        Me.SplitContainer1.TabIndex = 23
        '
        'Panel2
        '
        Me.Panel2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel2.Controls.Add(Me.numVotePriority)
        Me.Panel2.Controls.Add(Me.btnUpdate)
        Me.Panel2.Controls.Add(Me.btnSave)
        Me.Panel2.Controls.Add(Me.Label4)
        Me.Panel2.Controls.Add(Me.Label1)
        Me.Panel2.Controls.Add(Me.cboVoteHeadName)
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel2.Location = New System.Drawing.Point(5, 5)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(405, 115)
        Me.Panel2.TabIndex = 61
        '
        'numVotePriority
        '
        Me.numVotePriority.Location = New System.Drawing.Point(119, 48)
        Me.numVotePriority.Maximum = New Decimal(New Integer() {50, 0, 0, 0})
        Me.numVotePriority.Minimum = New Decimal(New Integer() {1, 0, 0, 0})
        Me.numVotePriority.Name = "numVotePriority"
        Me.numVotePriority.ReadOnly = True
        Me.numVotePriority.Size = New System.Drawing.Size(269, 20)
        Me.numVotePriority.TabIndex = 11
        Me.numVotePriority.Value = New Decimal(New Integer() {1, 0, 0, 0})
        '
        'btnUpdate
        '
        Me.btnUpdate.Enabled = False
        Me.btnUpdate.Location = New System.Drawing.Point(313, 79)
        Me.btnUpdate.Name = "btnUpdate"
        Me.btnUpdate.Size = New System.Drawing.Size(75, 23)
        Me.btnUpdate.TabIndex = 10
        Me.btnUpdate.Text = "Update"
        Me.btnUpdate.UseVisualStyleBackColor = True
        '
        'btnSave
        '
        Me.btnSave.Location = New System.Drawing.Point(119, 79)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(75, 23)
        Me.btnSave.TabIndex = 4
        Me.btnSave.Text = "Save"
        Me.btnSave.UseVisualStyleBackColor = True
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(13, 50)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(92, 13)
        Me.Label4.TabIndex = 6
        Me.Label4.Text = "VoteHead Priority:"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(13, 17)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(89, 13)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "VoteHead Name:"
        '
        'cboVoteHeadName
        '
        Me.cboVoteHeadName.FormattingEnabled = True
        Me.cboVoteHeadName.Location = New System.Drawing.Point(119, 14)
        Me.cboVoteHeadName.Name = "cboVoteHeadName"
        Me.cboVoteHeadName.Size = New System.Drawing.Size(269, 21)
        Me.cboVoteHeadName.TabIndex = 0
        '
        'StatusStrip12
        '
        Me.StatusStrip12.AutoSize = False
        Me.StatusStrip12.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip12.Dock = System.Windows.Forms.DockStyle.Left
        Me.StatusStrip12.Location = New System.Drawing.Point(0, 5)
        Me.StatusStrip12.Name = "StatusStrip12"
        Me.StatusStrip12.Size = New System.Drawing.Size(5, 115)
        Me.StatusStrip12.SizingGrip = False
        Me.StatusStrip12.TabIndex = 60
        Me.StatusStrip12.Text = "StatusStrip12"
        '
        'StatusStrip13
        '
        Me.StatusStrip13.AutoSize = False
        Me.StatusStrip13.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip13.Location = New System.Drawing.Point(0, 120)
        Me.StatusStrip13.Name = "StatusStrip13"
        Me.StatusStrip13.Size = New System.Drawing.Size(410, 5)
        Me.StatusStrip13.SizingGrip = False
        Me.StatusStrip13.TabIndex = 59
        Me.StatusStrip13.Text = "StatusStrip13"
        '
        'StatusStrip14
        '
        Me.StatusStrip14.AutoSize = False
        Me.StatusStrip14.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip14.Dock = System.Windows.Forms.DockStyle.Top
        Me.StatusStrip14.Location = New System.Drawing.Point(0, 0)
        Me.StatusStrip14.Name = "StatusStrip14"
        Me.StatusStrip14.Size = New System.Drawing.Size(410, 5)
        Me.StatusStrip14.SizingGrip = False
        Me.StatusStrip14.TabIndex = 58
        Me.StatusStrip14.Text = "StatusStrip14"
        '
        'StatusStrip5
        '
        Me.StatusStrip5.AutoSize = False
        Me.StatusStrip5.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip5.Dock = System.Windows.Forms.DockStyle.Right
        Me.StatusStrip5.Location = New System.Drawing.Point(410, 0)
        Me.StatusStrip5.Name = "StatusStrip5"
        Me.StatusStrip5.Size = New System.Drawing.Size(5, 125)
        Me.StatusStrip5.SizingGrip = False
        Me.StatusStrip5.TabIndex = 57
        Me.StatusStrip5.Text = "StatusStrip5"
        '
        'SplitContainer2
        '
        Me.SplitContainer2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.SplitContainer2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer2.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer2.Name = "SplitContainer2"
        Me.SplitContainer2.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'SplitContainer2.Panel1
        '
        Me.SplitContainer2.Panel1.Controls.Add(Me.lstVoteHeads)
        '
        'SplitContainer2.Panel2
        '
        Me.SplitContainer2.Panel2.Controls.Add(Me.Panel1)
        Me.SplitContainer2.Panel2.Controls.Add(Me.StatusStrip3)
        Me.SplitContainer2.Panel2.Controls.Add(Me.StatusStrip4)
        Me.SplitContainer2.Panel2.Controls.Add(Me.StatusStrip6)
        Me.SplitContainer2.Panel2.Controls.Add(Me.StatusStrip7)
        Me.SplitContainer2.Size = New System.Drawing.Size(415, 307)
        Me.SplitContainer2.SplitterDistance = 244
        Me.SplitContainer2.SplitterWidth = 1
        Me.SplitContainer2.TabIndex = 0
        '
        'lstVoteHeads
        '
        Me.lstVoteHeads.CheckBoxes = True
        Me.lstVoteHeads.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColumnHeader2, Me.ColumnHeader3})
        Me.lstVoteHeads.ContextMenuStrip = Me.ContextMenuStrip1
        Me.lstVoteHeads.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lstVoteHeads.FullRowSelect = True
        Me.lstVoteHeads.GridLines = True
        Me.lstVoteHeads.Location = New System.Drawing.Point(0, 0)
        Me.lstVoteHeads.Name = "lstVoteHeads"
        Me.lstVoteHeads.Size = New System.Drawing.Size(413, 242)
        Me.lstVoteHeads.TabIndex = 28
        Me.lstVoteHeads.UseCompatibleStateImageBehavior = False
        Me.lstVoteHeads.View = System.Windows.Forms.View.Details
        '
        'ColumnHeader2
        '
        Me.ColumnHeader2.Text = "VoteHead Name"
        Me.ColumnHeader2.Width = 300
        '
        'ColumnHeader3
        '
        Me.ColumnHeader3.Text = "Priority"
        Me.ColumnHeader3.Width = 100
        '
        'ContextMenuStrip1
        '
        Me.ContextMenuStrip1.BackColor = System.Drawing.Color.LightSeaGreen
        Me.ContextMenuStrip1.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ContextMenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.UPDATEToolStripMenuItem, Me.DELETEToolStripMenuItem, Me.CLOSEToolStripMenuItem})
        Me.ContextMenuStrip1.Name = "ContextMenuStrip1"
        Me.ContextMenuStrip1.Size = New System.Drawing.Size(123, 70)
        '
        'UPDATEToolStripMenuItem
        '
        Me.UPDATEToolStripMenuItem.Name = "UPDATEToolStripMenuItem"
        Me.UPDATEToolStripMenuItem.Size = New System.Drawing.Size(122, 22)
        Me.UPDATEToolStripMenuItem.Text = "EDIT"
        '
        'DELETEToolStripMenuItem
        '
        Me.DELETEToolStripMenuItem.Name = "DELETEToolStripMenuItem"
        Me.DELETEToolStripMenuItem.Size = New System.Drawing.Size(122, 22)
        Me.DELETEToolStripMenuItem.Text = "DELETE"
        '
        'CLOSEToolStripMenuItem
        '
        Me.CLOSEToolStripMenuItem.Name = "CLOSEToolStripMenuItem"
        Me.CLOSEToolStripMenuItem.Size = New System.Drawing.Size(122, 22)
        Me.CLOSEToolStripMenuItem.Text = "CLOSE"
        '
        'Panel1
        '
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel1.Controls.Add(Me.btnView)
        Me.Panel1.Controls.Add(Me.btnClose)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel1.Location = New System.Drawing.Point(5, 5)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(403, 50)
        Me.Panel1.TabIndex = 65
        '
        'btnView
        '
        Me.btnView.Location = New System.Drawing.Point(118, 13)
        Me.btnView.Name = "btnView"
        Me.btnView.Size = New System.Drawing.Size(75, 23)
        Me.btnView.TabIndex = 8
        Me.btnView.Text = "View"
        Me.btnView.UseVisualStyleBackColor = True
        '
        'btnClose
        '
        Me.btnClose.Location = New System.Drawing.Point(312, 13)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(75, 23)
        Me.btnClose.TabIndex = 7
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'StatusStrip3
        '
        Me.StatusStrip3.AutoSize = False
        Me.StatusStrip3.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip3.Dock = System.Windows.Forms.DockStyle.Left
        Me.StatusStrip3.Location = New System.Drawing.Point(0, 5)
        Me.StatusStrip3.Name = "StatusStrip3"
        Me.StatusStrip3.Size = New System.Drawing.Size(5, 50)
        Me.StatusStrip3.SizingGrip = False
        Me.StatusStrip3.TabIndex = 64
        Me.StatusStrip3.Text = "StatusStrip3"
        '
        'StatusStrip4
        '
        Me.StatusStrip4.AutoSize = False
        Me.StatusStrip4.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip4.Location = New System.Drawing.Point(0, 55)
        Me.StatusStrip4.Name = "StatusStrip4"
        Me.StatusStrip4.Size = New System.Drawing.Size(408, 5)
        Me.StatusStrip4.SizingGrip = False
        Me.StatusStrip4.TabIndex = 63
        Me.StatusStrip4.Text = "StatusStrip4"
        '
        'StatusStrip6
        '
        Me.StatusStrip6.AutoSize = False
        Me.StatusStrip6.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip6.Dock = System.Windows.Forms.DockStyle.Top
        Me.StatusStrip6.Location = New System.Drawing.Point(0, 0)
        Me.StatusStrip6.Name = "StatusStrip6"
        Me.StatusStrip6.Size = New System.Drawing.Size(408, 5)
        Me.StatusStrip6.SizingGrip = False
        Me.StatusStrip6.TabIndex = 62
        Me.StatusStrip6.Text = "StatusStrip6"
        '
        'StatusStrip7
        '
        Me.StatusStrip7.AutoSize = False
        Me.StatusStrip7.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip7.Dock = System.Windows.Forms.DockStyle.Right
        Me.StatusStrip7.Location = New System.Drawing.Point(408, 0)
        Me.StatusStrip7.Name = "StatusStrip7"
        Me.StatusStrip7.Size = New System.Drawing.Size(5, 60)
        Me.StatusStrip7.SizingGrip = False
        Me.StatusStrip7.TabIndex = 61
        Me.StatusStrip7.Text = "StatusStrip7"
        '
        'StatusStrip2
        '
        Me.StatusStrip2.AutoSize = False
        Me.StatusStrip2.BackColor = System.Drawing.Color.LightSeaGreen
        Me.StatusStrip2.Dock = System.Windows.Forms.DockStyle.Top
        Me.StatusStrip2.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripStatusLabel1})
        Me.StatusStrip2.Location = New System.Drawing.Point(0, 0)
        Me.StatusStrip2.Name = "StatusStrip2"
        Me.StatusStrip2.Size = New System.Drawing.Size(415, 25)
        Me.StatusStrip2.SizingGrip = False
        Me.StatusStrip2.TabIndex = 22
        Me.StatusStrip2.Text = "StatusStrip2"
        '
        'ToolStripStatusLabel1
        '
        Me.ToolStripStatusLabel1.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ToolStripStatusLabel1.Name = "ToolStripStatusLabel1"
        Me.ToolStripStatusLabel1.Size = New System.Drawing.Size(88, 20)
        Me.ToolStripStatusLabel1.Text = "VOTE HEADS"
        '
        'StatusStrip1
        '
        Me.StatusStrip1.AutoSize = False
        Me.StatusStrip1.BackColor = System.Drawing.Color.LightSeaGreen
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 458)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(415, 25)
        Me.StatusStrip1.SizingGrip = False
        Me.StatusStrip1.TabIndex = 21
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'frmFinFeeVoteHeads
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.LightSkyBlue
        Me.ClientSize = New System.Drawing.Size(438, 508)
        Me.Controls.Add(Me.pnlFeeVoteHeads)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Name = "frmFinFeeVoteHeads"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "VoteHeads"
        Me.pnlFeeVoteHeads.ResumeLayout(False)
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.ResumeLayout(False)
        Me.Panel2.ResumeLayout(False)
        Me.Panel2.PerformLayout()
        CType(Me.numVotePriority, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer2.Panel1.ResumeLayout(False)
        Me.SplitContainer2.Panel2.ResumeLayout(False)
        CType(Me.SplitContainer2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer2.ResumeLayout(False)
        Me.ContextMenuStrip1.ResumeLayout(False)
        Me.Panel1.ResumeLayout(False)
        Me.StatusStrip2.ResumeLayout(False)
        Me.StatusStrip2.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents pnlFeeVoteHeads As System.Windows.Forms.Panel
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents btnUpdate As System.Windows.Forms.Button
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents cboVoteHeadName As System.Windows.Forms.ComboBox
    Friend WithEvents StatusStrip12 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip13 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip14 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip5 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip2 As System.Windows.Forms.StatusStrip
    Friend WithEvents ToolStripStatusLabel1 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents numVotePriority As System.Windows.Forms.NumericUpDown
    Friend WithEvents SplitContainer2 As System.Windows.Forms.SplitContainer
    Friend WithEvents lstVoteHeads As System.Windows.Forms.ListView
    Friend WithEvents ColumnHeader2 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader3 As System.Windows.Forms.ColumnHeader
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents StatusStrip3 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip4 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip6 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip7 As System.Windows.Forms.StatusStrip
    Friend WithEvents btnView As System.Windows.Forms.Button
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents ContextMenuStrip1 As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents UPDATEToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents DELETEToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents CLOSEToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
End Class
