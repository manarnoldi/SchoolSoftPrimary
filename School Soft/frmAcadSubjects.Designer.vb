<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmAcadSubjects
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
        Me.pnlAcadSubj = New System.Windows.Forms.Panel()
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
        Me.cboSubjGroup = New System.Windows.Forms.ComboBox()
        Me.txtSubjName = New System.Windows.Forms.TextBox()
        Me.txtSubjCode = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.StatusStrip5 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip4 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip3 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip10 = New System.Windows.Forms.StatusStrip()
        Me.SplitContainer2 = New System.Windows.Forms.SplitContainer()
        Me.lstSubjects = New System.Windows.Forms.ListView()
        Me.ColumnHeader1 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader2 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader3 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ContextMenuStrip1 = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.UPDATEToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.DELETEToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.CLOSEToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem()
        Me.btnView = New System.Windows.Forms.Button()
        Me.btnUpdate = New System.Windows.Forms.Button()
        Me.btnSave = New System.Windows.Forms.Button()
        Me.btnClose = New System.Windows.Forms.Button()
        Me.StatusStrip6 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip7 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip8 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip9 = New System.Windows.Forms.StatusStrip()
        Me.ToolStripStatusLabel2 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.StatusStrip2 = New System.Windows.Forms.StatusStrip()
        Me.ToolStripStatusLabel1 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.pnlAcadSubj.SuspendLayout()
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        CType(Me.SplitContainer2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer2.Panel1.SuspendLayout()
        Me.SplitContainer2.Panel2.SuspendLayout()
        Me.SplitContainer2.SuspendLayout()
        Me.ContextMenuStrip1.SuspendLayout()
        Me.StatusStrip9.SuspendLayout()
        Me.StatusStrip2.SuspendLayout()
        Me.SuspendLayout()
        '
        'pnlAcadSubj
        '
        Me.pnlAcadSubj.BackColor = System.Drawing.SystemColors.Control
        Me.pnlAcadSubj.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pnlAcadSubj.Controls.Add(Me.SplitContainer1)
        Me.pnlAcadSubj.Controls.Add(Me.StatusStrip2)
        Me.pnlAcadSubj.Controls.Add(Me.StatusStrip1)
        Me.pnlAcadSubj.Location = New System.Drawing.Point(12, 12)
        Me.pnlAcadSubj.Name = "pnlAcadSubj"
        Me.pnlAcadSubj.Size = New System.Drawing.Size(489, 530)
        Me.pnlAcadSubj.TabIndex = 0
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
        Me.SplitContainer1.Panel1.Controls.Add(Me.cboSubjGroup)
        Me.SplitContainer1.Panel1.Controls.Add(Me.txtSubjName)
        Me.SplitContainer1.Panel1.Controls.Add(Me.txtSubjCode)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label3)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label2)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label1)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Panel1)
        Me.SplitContainer1.Panel1.Controls.Add(Me.StatusStrip5)
        Me.SplitContainer1.Panel1.Controls.Add(Me.StatusStrip4)
        Me.SplitContainer1.Panel1.Controls.Add(Me.StatusStrip3)
        Me.SplitContainer1.Panel1.Controls.Add(Me.StatusStrip10)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.SplitContainer2)
        Me.SplitContainer1.Size = New System.Drawing.Size(487, 478)
        Me.SplitContainer1.SplitterDistance = 152
        Me.SplitContainer1.SplitterWidth = 1
        Me.SplitContainer1.TabIndex = 29
        '
        'cboSubjGroup
        '
        Me.cboSubjGroup.FormattingEnabled = True
        Me.cboSubjGroup.Location = New System.Drawing.Point(160, 102)
        Me.cboSubjGroup.Name = "cboSubjGroup"
        Me.cboSubjGroup.Size = New System.Drawing.Size(269, 21)
        Me.cboSubjGroup.TabIndex = 43
        '
        'txtSubjName
        '
        Me.txtSubjName.Location = New System.Drawing.Point(160, 64)
        Me.txtSubjName.Name = "txtSubjName"
        Me.txtSubjName.Size = New System.Drawing.Size(269, 20)
        Me.txtSubjName.TabIndex = 42
        '
        'txtSubjCode
        '
        Me.txtSubjCode.Location = New System.Drawing.Point(160, 26)
        Me.txtSubjCode.Name = "txtSubjCode"
        Me.txtSubjCode.Size = New System.Drawing.Size(269, 20)
        Me.txtSubjCode.TabIndex = 41
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(51, 105)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(100, 13)
        Me.Label3.TabIndex = 40
        Me.Label3.Text = "SUBJECT GROUP:"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(51, 67)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(92, 13)
        Me.Label2.TabIndex = 39
        Me.Label2.Text = "SUBJECT NAME:"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(51, 29)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(91, 13)
        Me.Label1.TabIndex = 38
        Me.Label1.Text = "SUBJECT CODE:"
        '
        'StatusStrip5
        '
        Me.StatusStrip5.AutoSize = False
        Me.StatusStrip5.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip5.Dock = System.Windows.Forms.DockStyle.Right
        Me.StatusStrip5.Location = New System.Drawing.Point(480, 5)
        Me.StatusStrip5.Name = "StatusStrip5"
        Me.StatusStrip5.Size = New System.Drawing.Size(5, 140)
        Me.StatusStrip5.SizingGrip = False
        Me.StatusStrip5.TabIndex = 37
        Me.StatusStrip5.Text = "StatusStrip5"
        '
        'StatusStrip4
        '
        Me.StatusStrip4.AutoSize = False
        Me.StatusStrip4.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip4.Dock = System.Windows.Forms.DockStyle.Left
        Me.StatusStrip4.Location = New System.Drawing.Point(0, 5)
        Me.StatusStrip4.Name = "StatusStrip4"
        Me.StatusStrip4.Size = New System.Drawing.Size(5, 140)
        Me.StatusStrip4.SizingGrip = False
        Me.StatusStrip4.TabIndex = 36
        Me.StatusStrip4.Text = "StatusStrip4"
        '
        'StatusStrip3
        '
        Me.StatusStrip3.AutoSize = False
        Me.StatusStrip3.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip3.Location = New System.Drawing.Point(0, 145)
        Me.StatusStrip3.Name = "StatusStrip3"
        Me.StatusStrip3.Size = New System.Drawing.Size(485, 5)
        Me.StatusStrip3.SizingGrip = False
        Me.StatusStrip3.TabIndex = 35
        Me.StatusStrip3.Text = "StatusStrip3"
        '
        'StatusStrip10
        '
        Me.StatusStrip10.AutoSize = False
        Me.StatusStrip10.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip10.Dock = System.Windows.Forms.DockStyle.Top
        Me.StatusStrip10.Location = New System.Drawing.Point(0, 0)
        Me.StatusStrip10.Name = "StatusStrip10"
        Me.StatusStrip10.Size = New System.Drawing.Size(485, 5)
        Me.StatusStrip10.SizingGrip = False
        Me.StatusStrip10.TabIndex = 34
        Me.StatusStrip10.Text = "StatusStrip10"
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
        Me.SplitContainer2.Panel1.Controls.Add(Me.lstSubjects)
        '
        'SplitContainer2.Panel2
        '
        Me.SplitContainer2.Panel2.Controls.Add(Me.btnView)
        Me.SplitContainer2.Panel2.Controls.Add(Me.btnUpdate)
        Me.SplitContainer2.Panel2.Controls.Add(Me.btnSave)
        Me.SplitContainer2.Panel2.Controls.Add(Me.btnClose)
        Me.SplitContainer2.Panel2.Controls.Add(Me.Panel2)
        Me.SplitContainer2.Panel2.Controls.Add(Me.StatusStrip6)
        Me.SplitContainer2.Panel2.Controls.Add(Me.StatusStrip7)
        Me.SplitContainer2.Panel2.Controls.Add(Me.StatusStrip8)
        Me.SplitContainer2.Panel2.Controls.Add(Me.StatusStrip9)
        Me.SplitContainer2.Size = New System.Drawing.Size(487, 325)
        Me.SplitContainer2.SplitterDistance = 258
        Me.SplitContainer2.SplitterWidth = 1
        Me.SplitContainer2.TabIndex = 30
        '
        'lstSubjects
        '
        Me.lstSubjects.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColumnHeader1, Me.ColumnHeader2, Me.ColumnHeader3})
        Me.lstSubjects.ContextMenuStrip = Me.ContextMenuStrip1
        Me.lstSubjects.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lstSubjects.FullRowSelect = True
        Me.lstSubjects.GridLines = True
        Me.lstSubjects.Location = New System.Drawing.Point(0, 0)
        Me.lstSubjects.Name = "lstSubjects"
        Me.lstSubjects.Size = New System.Drawing.Size(485, 256)
        Me.lstSubjects.TabIndex = 29
        Me.lstSubjects.UseCompatibleStateImageBehavior = False
        Me.lstSubjects.View = System.Windows.Forms.View.Details
        '
        'ColumnHeader1
        '
        Me.ColumnHeader1.Text = "SUB CODE"
        Me.ColumnHeader1.Width = 100
        '
        'ColumnHeader2
        '
        Me.ColumnHeader2.Text = "SUBJECT NAME"
        Me.ColumnHeader2.Width = 229
        '
        'ColumnHeader3
        '
        Me.ColumnHeader3.Text = "SUB GROUP"
        Me.ColumnHeader3.Width = 150
        '
        'ContextMenuStrip1
        '
        Me.ContextMenuStrip1.BackColor = System.Drawing.Color.LightSeaGreen
        Me.ContextMenuStrip1.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ContextMenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.UPDATEToolStripMenuItem, Me.DELETEToolStripMenuItem, Me.CLOSEToolStripMenuItem1})
        Me.ContextMenuStrip1.Name = "ContextMenuStrip1"
        Me.ContextMenuStrip1.Size = New System.Drawing.Size(127, 70)
        '
        'UPDATEToolStripMenuItem
        '
        Me.UPDATEToolStripMenuItem.Name = "UPDATEToolStripMenuItem"
        Me.UPDATEToolStripMenuItem.Size = New System.Drawing.Size(126, 22)
        Me.UPDATEToolStripMenuItem.Text = "UPDATE"
        '
        'DELETEToolStripMenuItem
        '
        Me.DELETEToolStripMenuItem.Name = "DELETEToolStripMenuItem"
        Me.DELETEToolStripMenuItem.Size = New System.Drawing.Size(126, 22)
        Me.DELETEToolStripMenuItem.Text = "DELETE"
        '
        'CLOSEToolStripMenuItem1
        '
        Me.CLOSEToolStripMenuItem1.Name = "CLOSEToolStripMenuItem1"
        Me.CLOSEToolStripMenuItem1.Size = New System.Drawing.Size(126, 22)
        Me.CLOSEToolStripMenuItem1.Text = "CLOSE"
        '
        'btnView
        '
        Me.btnView.Location = New System.Drawing.Point(11, 28)
        Me.btnView.Name = "btnView"
        Me.btnView.Size = New System.Drawing.Size(75, 23)
        Me.btnView.TabIndex = 57
        Me.btnView.Text = "View"
        Me.btnView.UseVisualStyleBackColor = True
        '
        'btnUpdate
        '
        Me.btnUpdate.Enabled = False
        Me.btnUpdate.Location = New System.Drawing.Point(257, 28)
        Me.btnUpdate.Name = "btnUpdate"
        Me.btnUpdate.Size = New System.Drawing.Size(97, 23)
        Me.btnUpdate.TabIndex = 55
        Me.btnUpdate.Text = "Save Changes"
        Me.btnUpdate.UseVisualStyleBackColor = True
        '
        'btnSave
        '
        Me.btnSave.Location = New System.Drawing.Point(133, 28)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(75, 23)
        Me.btnSave.TabIndex = 54
        Me.btnSave.Text = "Save"
        Me.btnSave.UseVisualStyleBackColor = True
        '
        'btnClose
        '
        Me.btnClose.Location = New System.Drawing.Point(398, 28)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(75, 23)
        Me.btnClose.TabIndex = 53
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'StatusStrip6
        '
        Me.StatusStrip6.AutoSize = False
        Me.StatusStrip6.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip6.Dock = System.Windows.Forms.DockStyle.Right
        Me.StatusStrip6.Location = New System.Drawing.Point(480, 22)
        Me.StatusStrip6.Name = "StatusStrip6"
        Me.StatusStrip6.Size = New System.Drawing.Size(5, 37)
        Me.StatusStrip6.SizingGrip = False
        Me.StatusStrip6.TabIndex = 37
        Me.StatusStrip6.Text = "StatusStrip6"
        '
        'StatusStrip7
        '
        Me.StatusStrip7.AutoSize = False
        Me.StatusStrip7.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip7.Dock = System.Windows.Forms.DockStyle.Left
        Me.StatusStrip7.Location = New System.Drawing.Point(0, 22)
        Me.StatusStrip7.Name = "StatusStrip7"
        Me.StatusStrip7.Size = New System.Drawing.Size(5, 37)
        Me.StatusStrip7.SizingGrip = False
        Me.StatusStrip7.TabIndex = 36
        Me.StatusStrip7.Text = "StatusStrip7"
        '
        'StatusStrip8
        '
        Me.StatusStrip8.AutoSize = False
        Me.StatusStrip8.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip8.Location = New System.Drawing.Point(0, 59)
        Me.StatusStrip8.Name = "StatusStrip8"
        Me.StatusStrip8.Size = New System.Drawing.Size(485, 5)
        Me.StatusStrip8.SizingGrip = False
        Me.StatusStrip8.TabIndex = 35
        Me.StatusStrip8.Text = "StatusStrip8"
        '
        'StatusStrip9
        '
        Me.StatusStrip9.AutoSize = False
        Me.StatusStrip9.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip9.Dock = System.Windows.Forms.DockStyle.Top
        Me.StatusStrip9.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripStatusLabel2})
        Me.StatusStrip9.Location = New System.Drawing.Point(0, 0)
        Me.StatusStrip9.Name = "StatusStrip9"
        Me.StatusStrip9.Size = New System.Drawing.Size(485, 22)
        Me.StatusStrip9.SizingGrip = False
        Me.StatusStrip9.TabIndex = 34
        '
        'ToolStripStatusLabel2
        '
        Me.ToolStripStatusLabel2.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ToolStripStatusLabel2.Name = "ToolStripStatusLabel2"
        Me.ToolStripStatusLabel2.Size = New System.Drawing.Size(85, 17)
        Me.ToolStripStatusLabel2.Text = "User Actions"
        '
        'StatusStrip2
        '
        Me.StatusStrip2.AutoSize = False
        Me.StatusStrip2.BackColor = System.Drawing.Color.LightSeaGreen
        Me.StatusStrip2.Dock = System.Windows.Forms.DockStyle.Top
        Me.StatusStrip2.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripStatusLabel1})
        Me.StatusStrip2.Location = New System.Drawing.Point(0, 0)
        Me.StatusStrip2.Name = "StatusStrip2"
        Me.StatusStrip2.Size = New System.Drawing.Size(487, 25)
        Me.StatusStrip2.SizingGrip = False
        Me.StatusStrip2.TabIndex = 28
        Me.StatusStrip2.Text = "StatusStrip2"
        '
        'ToolStripStatusLabel1
        '
        Me.ToolStripStatusLabel1.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ToolStripStatusLabel1.Name = "ToolStripStatusLabel1"
        Me.ToolStripStatusLabel1.Size = New System.Drawing.Size(130, 20)
        Me.ToolStripStatusLabel1.Text = "SUBJECTS REGISTER"
        '
        'StatusStrip1
        '
        Me.StatusStrip1.AutoSize = False
        Me.StatusStrip1.BackColor = System.Drawing.Color.LightSeaGreen
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 503)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(487, 25)
        Me.StatusStrip1.SizingGrip = False
        Me.StatusStrip1.TabIndex = 27
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'Panel1
        '
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel1.Location = New System.Drawing.Point(5, 5)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(475, 140)
        Me.Panel1.TabIndex = 44
        '
        'Panel2
        '
        Me.Panel2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel2.Location = New System.Drawing.Point(5, 22)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(475, 37)
        Me.Panel2.TabIndex = 58
        '
        'frmAcadSubjects
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.LightSkyBlue
        Me.ClientSize = New System.Drawing.Size(512, 554)
        Me.Controls.Add(Me.pnlAcadSubj)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Name = "frmAcadSubjects"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Subjects"
        Me.pnlAcadSubj.ResumeLayout(False)
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel1.PerformLayout()
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.ResumeLayout(False)
        Me.SplitContainer2.Panel1.ResumeLayout(False)
        Me.SplitContainer2.Panel2.ResumeLayout(False)
        CType(Me.SplitContainer2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer2.ResumeLayout(False)
        Me.ContextMenuStrip1.ResumeLayout(False)
        Me.StatusStrip9.ResumeLayout(False)
        Me.StatusStrip9.PerformLayout()
        Me.StatusStrip2.ResumeLayout(False)
        Me.StatusStrip2.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents pnlAcadSubj As System.Windows.Forms.Panel
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents SplitContainer2 As System.Windows.Forms.SplitContainer
    Friend WithEvents StatusStrip2 As System.Windows.Forms.StatusStrip
    Friend WithEvents ToolStripStatusLabel1 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents lstSubjects As System.Windows.Forms.ListView
    Friend WithEvents cboSubjGroup As System.Windows.Forms.ComboBox
    Friend WithEvents txtSubjName As System.Windows.Forms.TextBox
    Friend WithEvents txtSubjCode As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents StatusStrip5 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip4 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip3 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip10 As System.Windows.Forms.StatusStrip
    Friend WithEvents ColumnHeader1 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader2 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader3 As System.Windows.Forms.ColumnHeader
    Friend WithEvents StatusStrip6 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip7 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip8 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip9 As System.Windows.Forms.StatusStrip
    Friend WithEvents ToolStripStatusLabel2 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents btnView As System.Windows.Forms.Button
    Friend WithEvents btnUpdate As System.Windows.Forms.Button
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents ContextMenuStrip1 As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents UPDATEToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents DELETEToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents CLOSEToolStripMenuItem1 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
End Class
