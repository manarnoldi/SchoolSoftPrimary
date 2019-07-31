<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmAcadExaminationNames
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
        Me.pnlExamNames = New System.Windows.Forms.Panel()
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.cboExamName = New System.Windows.Forms.ComboBox()
        Me.cboExamType = New System.Windows.Forms.ComboBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.btnSave = New System.Windows.Forms.Button()
        Me.StatusStrip12 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip13 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip14 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip5 = New System.Windows.Forms.StatusStrip()
        Me.SplitContainer2 = New System.Windows.Forms.SplitContainer()
        Me.lstViewExamNames = New System.Windows.Forms.ListView()
        Me.ColumnHeader1 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader2 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ContextMenuStrip1 = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.UPDATEToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.DELETEToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.CLOSEToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.btnClear = New System.Windows.Forms.Button()
        Me.btnUpdate = New System.Windows.Forms.Button()
        Me.btnClose = New System.Windows.Forms.Button()
        Me.btnView = New System.Windows.Forms.Button()
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip3 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip4 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip6 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip2 = New System.Windows.Forms.StatusStrip()
        Me.ToolStripStatusLabel1 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.pnlExamNames.SuspendLayout()
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.Panel1.SuspendLayout()
        CType(Me.SplitContainer2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer2.Panel1.SuspendLayout()
        Me.SplitContainer2.Panel2.SuspendLayout()
        Me.SplitContainer2.SuspendLayout()
        Me.ContextMenuStrip1.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.StatusStrip2.SuspendLayout()
        Me.SuspendLayout()
        '
        'pnlExamNames
        '
        Me.pnlExamNames.BackColor = System.Drawing.SystemColors.Control
        Me.pnlExamNames.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pnlExamNames.Controls.Add(Me.SplitContainer1)
        Me.pnlExamNames.Controls.Add(Me.StatusStrip2)
        Me.pnlExamNames.Location = New System.Drawing.Point(12, 12)
        Me.pnlExamNames.Name = "pnlExamNames"
        Me.pnlExamNames.Size = New System.Drawing.Size(410, 461)
        Me.pnlExamNames.TabIndex = 0
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
        Me.SplitContainer1.Panel1.Controls.Add(Me.StatusStrip12)
        Me.SplitContainer1.Panel1.Controls.Add(Me.StatusStrip13)
        Me.SplitContainer1.Panel1.Controls.Add(Me.StatusStrip14)
        Me.SplitContainer1.Panel1.Controls.Add(Me.StatusStrip5)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.SplitContainer2)
        Me.SplitContainer1.Size = New System.Drawing.Size(408, 434)
        Me.SplitContainer1.SplitterDistance = 92
        Me.SplitContainer1.SplitterWidth = 1
        Me.SplitContainer1.TabIndex = 26
        '
        'Panel1
        '
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel1.Controls.Add(Me.cboExamName)
        Me.Panel1.Controls.Add(Me.cboExamType)
        Me.Panel1.Controls.Add(Me.Label2)
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Controls.Add(Me.btnSave)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel1.Location = New System.Drawing.Point(5, 5)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(396, 80)
        Me.Panel1.TabIndex = 53
        '
        'cboExamName
        '
        Me.cboExamName.FormattingEnabled = True
        Me.cboExamName.Location = New System.Drawing.Point(91, 44)
        Me.cboExamName.Name = "cboExamName"
        Me.cboExamName.Size = New System.Drawing.Size(281, 21)
        Me.cboExamName.TabIndex = 3
        '
        'cboExamType
        '
        Me.cboExamType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboExamType.FormattingEnabled = True
        Me.cboExamType.Location = New System.Drawing.Point(90, 11)
        Me.cboExamType.Name = "cboExamType"
        Me.cboExamType.Size = New System.Drawing.Size(192, 21)
        Me.cboExamType.TabIndex = 2
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(21, 47)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(67, 13)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "Exam Name:"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(21, 14)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(63, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Exam Type:"
        '
        'btnSave
        '
        Me.btnSave.Location = New System.Drawing.Point(297, 9)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(75, 23)
        Me.btnSave.TabIndex = 2
        Me.btnSave.Text = "Save"
        Me.btnSave.UseVisualStyleBackColor = True
        '
        'StatusStrip12
        '
        Me.StatusStrip12.AutoSize = False
        Me.StatusStrip12.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip12.Dock = System.Windows.Forms.DockStyle.Left
        Me.StatusStrip12.Location = New System.Drawing.Point(0, 5)
        Me.StatusStrip12.Name = "StatusStrip12"
        Me.StatusStrip12.Size = New System.Drawing.Size(5, 80)
        Me.StatusStrip12.SizingGrip = False
        Me.StatusStrip12.TabIndex = 52
        Me.StatusStrip12.Text = "StatusStrip12"
        '
        'StatusStrip13
        '
        Me.StatusStrip13.AutoSize = False
        Me.StatusStrip13.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip13.Location = New System.Drawing.Point(0, 85)
        Me.StatusStrip13.Name = "StatusStrip13"
        Me.StatusStrip13.Size = New System.Drawing.Size(401, 5)
        Me.StatusStrip13.SizingGrip = False
        Me.StatusStrip13.TabIndex = 51
        Me.StatusStrip13.Text = "StatusStrip13"
        '
        'StatusStrip14
        '
        Me.StatusStrip14.AutoSize = False
        Me.StatusStrip14.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip14.Dock = System.Windows.Forms.DockStyle.Top
        Me.StatusStrip14.Location = New System.Drawing.Point(0, 0)
        Me.StatusStrip14.Name = "StatusStrip14"
        Me.StatusStrip14.Size = New System.Drawing.Size(401, 5)
        Me.StatusStrip14.SizingGrip = False
        Me.StatusStrip14.TabIndex = 50
        Me.StatusStrip14.Text = "StatusStrip14"
        '
        'StatusStrip5
        '
        Me.StatusStrip5.AutoSize = False
        Me.StatusStrip5.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip5.Dock = System.Windows.Forms.DockStyle.Right
        Me.StatusStrip5.Location = New System.Drawing.Point(401, 0)
        Me.StatusStrip5.Name = "StatusStrip5"
        Me.StatusStrip5.Size = New System.Drawing.Size(5, 90)
        Me.StatusStrip5.SizingGrip = False
        Me.StatusStrip5.TabIndex = 49
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
        Me.SplitContainer2.Panel1.Controls.Add(Me.lstViewExamNames)
        '
        'SplitContainer2.Panel2
        '
        Me.SplitContainer2.Panel2.Controls.Add(Me.Panel2)
        Me.SplitContainer2.Panel2.Controls.Add(Me.StatusStrip1)
        Me.SplitContainer2.Panel2.Controls.Add(Me.StatusStrip3)
        Me.SplitContainer2.Panel2.Controls.Add(Me.StatusStrip4)
        Me.SplitContainer2.Panel2.Controls.Add(Me.StatusStrip6)
        Me.SplitContainer2.Size = New System.Drawing.Size(408, 341)
        Me.SplitContainer2.SplitterDistance = 286
        Me.SplitContainer2.SplitterWidth = 1
        Me.SplitContainer2.TabIndex = 27
        '
        'lstViewExamNames
        '
        Me.lstViewExamNames.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lstViewExamNames.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColumnHeader1, Me.ColumnHeader2})
        Me.lstViewExamNames.ContextMenuStrip = Me.ContextMenuStrip1
        Me.lstViewExamNames.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lstViewExamNames.FullRowSelect = True
        Me.lstViewExamNames.GridLines = True
        Me.lstViewExamNames.Location = New System.Drawing.Point(0, 0)
        Me.lstViewExamNames.Name = "lstViewExamNames"
        Me.lstViewExamNames.Size = New System.Drawing.Size(406, 284)
        Me.lstViewExamNames.TabIndex = 31
        Me.lstViewExamNames.UseCompatibleStateImageBehavior = False
        Me.lstViewExamNames.View = System.Windows.Forms.View.Details
        '
        'ColumnHeader1
        '
        Me.ColumnHeader1.Text = "Exam Type"
        Me.ColumnHeader1.Width = 150
        '
        'ColumnHeader2
        '
        Me.ColumnHeader2.Text = "Exam Name"
        Me.ColumnHeader2.Width = 250
        '
        'ContextMenuStrip1
        '
        Me.ContextMenuStrip1.BackColor = System.Drawing.Color.LightSeaGreen
        Me.ContextMenuStrip1.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ContextMenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.UPDATEToolStripMenuItem, Me.DELETEToolStripMenuItem, Me.CLOSEToolStripMenuItem})
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
        'CLOSEToolStripMenuItem
        '
        Me.CLOSEToolStripMenuItem.Name = "CLOSEToolStripMenuItem"
        Me.CLOSEToolStripMenuItem.Size = New System.Drawing.Size(126, 22)
        Me.CLOSEToolStripMenuItem.Text = "CLOSE"
        '
        'Panel2
        '
        Me.Panel2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel2.Controls.Add(Me.btnClear)
        Me.Panel2.Controls.Add(Me.btnUpdate)
        Me.Panel2.Controls.Add(Me.btnClose)
        Me.Panel2.Controls.Add(Me.btnView)
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel2.Location = New System.Drawing.Point(5, 5)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(396, 42)
        Me.Panel2.TabIndex = 54
        '
        'btnClear
        '
        Me.btnClear.Location = New System.Drawing.Point(110, 8)
        Me.btnClear.Name = "btnClear"
        Me.btnClear.Size = New System.Drawing.Size(75, 23)
        Me.btnClear.TabIndex = 5
        Me.btnClear.Text = "Clear"
        Me.btnClear.UseVisualStyleBackColor = True
        '
        'btnUpdate
        '
        Me.btnUpdate.Location = New System.Drawing.Point(210, 8)
        Me.btnUpdate.Name = "btnUpdate"
        Me.btnUpdate.Size = New System.Drawing.Size(75, 23)
        Me.btnUpdate.TabIndex = 4
        Me.btnUpdate.Text = "Update"
        Me.btnUpdate.UseVisualStyleBackColor = True
        '
        'btnClose
        '
        Me.btnClose.Location = New System.Drawing.Point(312, 8)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(75, 23)
        Me.btnClose.TabIndex = 4
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'btnView
        '
        Me.btnView.Location = New System.Drawing.Point(8, 8)
        Me.btnView.Name = "btnView"
        Me.btnView.Size = New System.Drawing.Size(75, 23)
        Me.btnView.TabIndex = 0
        Me.btnView.Text = "View"
        Me.btnView.UseVisualStyleBackColor = True
        '
        'StatusStrip1
        '
        Me.StatusStrip1.AutoSize = False
        Me.StatusStrip1.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip1.Dock = System.Windows.Forms.DockStyle.Left
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 5)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(5, 42)
        Me.StatusStrip1.SizingGrip = False
        Me.StatusStrip1.TabIndex = 52
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'StatusStrip3
        '
        Me.StatusStrip3.AutoSize = False
        Me.StatusStrip3.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip3.Location = New System.Drawing.Point(0, 47)
        Me.StatusStrip3.Name = "StatusStrip3"
        Me.StatusStrip3.Size = New System.Drawing.Size(401, 5)
        Me.StatusStrip3.SizingGrip = False
        Me.StatusStrip3.TabIndex = 51
        Me.StatusStrip3.Text = "StatusStrip3"
        '
        'StatusStrip4
        '
        Me.StatusStrip4.AutoSize = False
        Me.StatusStrip4.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip4.Dock = System.Windows.Forms.DockStyle.Top
        Me.StatusStrip4.Location = New System.Drawing.Point(0, 0)
        Me.StatusStrip4.Name = "StatusStrip4"
        Me.StatusStrip4.Size = New System.Drawing.Size(401, 5)
        Me.StatusStrip4.SizingGrip = False
        Me.StatusStrip4.TabIndex = 50
        Me.StatusStrip4.Text = "StatusStrip4"
        '
        'StatusStrip6
        '
        Me.StatusStrip6.AutoSize = False
        Me.StatusStrip6.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip6.Dock = System.Windows.Forms.DockStyle.Right
        Me.StatusStrip6.Location = New System.Drawing.Point(401, 0)
        Me.StatusStrip6.Name = "StatusStrip6"
        Me.StatusStrip6.Size = New System.Drawing.Size(5, 52)
        Me.StatusStrip6.SizingGrip = False
        Me.StatusStrip6.TabIndex = 49
        Me.StatusStrip6.Text = "StatusStrip6"
        '
        'StatusStrip2
        '
        Me.StatusStrip2.AutoSize = False
        Me.StatusStrip2.BackColor = System.Drawing.Color.LightSeaGreen
        Me.StatusStrip2.Dock = System.Windows.Forms.DockStyle.Top
        Me.StatusStrip2.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripStatusLabel1})
        Me.StatusStrip2.Location = New System.Drawing.Point(0, 0)
        Me.StatusStrip2.Name = "StatusStrip2"
        Me.StatusStrip2.Size = New System.Drawing.Size(408, 25)
        Me.StatusStrip2.SizingGrip = False
        Me.StatusStrip2.TabIndex = 25
        Me.StatusStrip2.Text = "StatusStrip2"
        '
        'ToolStripStatusLabel1
        '
        Me.ToolStripStatusLabel1.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ToolStripStatusLabel1.Name = "ToolStripStatusLabel1"
        Me.ToolStripStatusLabel1.Size = New System.Drawing.Size(94, 20)
        Me.ToolStripStatusLabel1.Text = "EXAM NAMES"
        '
        'frmAcadExaminationNames
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.LightSkyBlue
        Me.ClientSize = New System.Drawing.Size(434, 485)
        Me.Controls.Add(Me.pnlExamNames)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Name = "frmAcadExaminationNames"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Exam Names"
        Me.pnlExamNames.ResumeLayout(False)
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.ResumeLayout(False)
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.SplitContainer2.Panel1.ResumeLayout(False)
        Me.SplitContainer2.Panel2.ResumeLayout(False)
        CType(Me.SplitContainer2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer2.ResumeLayout(False)
        Me.ContextMenuStrip1.ResumeLayout(False)
        Me.Panel2.ResumeLayout(False)
        Me.StatusStrip2.ResumeLayout(False)
        Me.StatusStrip2.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents pnlExamNames As System.Windows.Forms.Panel
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents SplitContainer2 As System.Windows.Forms.SplitContainer
    Friend WithEvents StatusStrip2 As System.Windows.Forms.StatusStrip
    Friend WithEvents ToolStripStatusLabel1 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents lstViewExamNames As System.Windows.Forms.ListView
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents cboExamName As System.Windows.Forms.ComboBox
    Friend WithEvents cboExamType As System.Windows.Forms.ComboBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents StatusStrip12 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip13 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip14 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip5 As System.Windows.Forms.StatusStrip
    Friend WithEvents ColumnHeader1 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader2 As System.Windows.Forms.ColumnHeader
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents btnView As System.Windows.Forms.Button
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip3 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip4 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip6 As System.Windows.Forms.StatusStrip
    Friend WithEvents btnUpdate As System.Windows.Forms.Button
    Friend WithEvents ContextMenuStrip1 As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents UPDATEToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents DELETEToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents CLOSEToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents btnClear As System.Windows.Forms.Button
End Class
