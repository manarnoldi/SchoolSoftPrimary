<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmClassHeads
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
        Me.pnlClassHeads = New System.Windows.Forms.Panel()
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
        Me.btnView = New System.Windows.Forms.Button()
        Me.btnDelete = New System.Windows.Forms.Button()
        Me.btnUpdate = New System.Windows.Forms.Button()
        Me.btnSave = New System.Windows.Forms.Button()
        Me.btnClose = New System.Windows.Forms.Button()
        Me.txtStudName = New System.Windows.Forms.TextBox()
        Me.cboStudNo = New System.Windows.Forms.ComboBox()
        Me.txtTeacherName = New System.Windows.Forms.TextBox()
        Me.cboTeacherNo = New System.Windows.Forms.ComboBox()
        Me.cboClassYear = New System.Windows.Forms.ComboBox()
        Me.cboClassStream = New System.Windows.Forms.ComboBox()
        Me.cboClassName = New System.Windows.Forms.ComboBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.StatusStrip5 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip4 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip3 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip10 = New System.Windows.Forms.StatusStrip()
        Me.lstClassHeads = New System.Windows.Forms.ListView()
        Me.ColumnHeader1 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader2 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader3 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader4 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader5 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader6 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader7 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.StatusStrip2 = New System.Windows.Forms.StatusStrip()
        Me.ToolStripStatusLabel1 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip()
        Me.pnlClassHeads.SuspendLayout()
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.StatusStrip2.SuspendLayout()
        Me.SuspendLayout()
        '
        'pnlClassHeads
        '
        Me.pnlClassHeads.BackColor = System.Drawing.SystemColors.Control
        Me.pnlClassHeads.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pnlClassHeads.Controls.Add(Me.SplitContainer1)
        Me.pnlClassHeads.Controls.Add(Me.StatusStrip2)
        Me.pnlClassHeads.Controls.Add(Me.StatusStrip1)
        Me.pnlClassHeads.Location = New System.Drawing.Point(12, 12)
        Me.pnlClassHeads.Name = "pnlClassHeads"
        Me.pnlClassHeads.Size = New System.Drawing.Size(892, 495)
        Me.pnlClassHeads.TabIndex = 0
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
        Me.SplitContainer1.Panel1.Controls.Add(Me.btnView)
        Me.SplitContainer1.Panel1.Controls.Add(Me.btnDelete)
        Me.SplitContainer1.Panel1.Controls.Add(Me.btnUpdate)
        Me.SplitContainer1.Panel1.Controls.Add(Me.btnSave)
        Me.SplitContainer1.Panel1.Controls.Add(Me.btnClose)
        Me.SplitContainer1.Panel1.Controls.Add(Me.txtStudName)
        Me.SplitContainer1.Panel1.Controls.Add(Me.cboStudNo)
        Me.SplitContainer1.Panel1.Controls.Add(Me.txtTeacherName)
        Me.SplitContainer1.Panel1.Controls.Add(Me.cboTeacherNo)
        Me.SplitContainer1.Panel1.Controls.Add(Me.cboClassYear)
        Me.SplitContainer1.Panel1.Controls.Add(Me.cboClassStream)
        Me.SplitContainer1.Panel1.Controls.Add(Me.cboClassName)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label7)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label6)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label5)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label4)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label3)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label2)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label1)
        Me.SplitContainer1.Panel1.Controls.Add(Me.StatusStrip5)
        Me.SplitContainer1.Panel1.Controls.Add(Me.StatusStrip4)
        Me.SplitContainer1.Panel1.Controls.Add(Me.StatusStrip3)
        Me.SplitContainer1.Panel1.Controls.Add(Me.StatusStrip10)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.lstClassHeads)
        Me.SplitContainer1.Size = New System.Drawing.Size(890, 443)
        Me.SplitContainer1.SplitterDistance = 114
        Me.SplitContainer1.SplitterWidth = 1
        Me.SplitContainer1.TabIndex = 23
        '
        'btnView
        '
        Me.btnView.Location = New System.Drawing.Point(324, 79)
        Me.btnView.Name = "btnView"
        Me.btnView.Size = New System.Drawing.Size(87, 23)
        Me.btnView.TabIndex = 57
        Me.btnView.Text = "View"
        Me.btnView.UseVisualStyleBackColor = True
        '
        'btnDelete
        '
        Me.btnDelete.Enabled = False
        Me.btnDelete.Location = New System.Drawing.Point(686, 79)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(75, 23)
        Me.btnDelete.TabIndex = 56
        Me.btnDelete.Text = "Delete"
        Me.btnDelete.UseVisualStyleBackColor = True
        '
        'btnUpdate
        '
        Me.btnUpdate.Enabled = False
        Me.btnUpdate.Location = New System.Drawing.Point(568, 79)
        Me.btnUpdate.Name = "btnUpdate"
        Me.btnUpdate.Size = New System.Drawing.Size(75, 23)
        Me.btnUpdate.TabIndex = 55
        Me.btnUpdate.Text = "Update"
        Me.btnUpdate.UseVisualStyleBackColor = True
        '
        'btnSave
        '
        Me.btnSave.Location = New System.Drawing.Point(451, 79)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(75, 23)
        Me.btnSave.TabIndex = 54
        Me.btnSave.Text = "Save"
        Me.btnSave.UseVisualStyleBackColor = True
        '
        'btnClose
        '
        Me.btnClose.Location = New System.Drawing.Point(799, 79)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(75, 23)
        Me.btnClose.TabIndex = 53
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'txtStudName
        '
        Me.txtStudName.Location = New System.Drawing.Point(520, 48)
        Me.txtStudName.Name = "txtStudName"
        Me.txtStudName.ReadOnly = True
        Me.txtStudName.Size = New System.Drawing.Size(354, 20)
        Me.txtStudName.TabIndex = 51
        '
        'cboStudNo
        '
        Me.cboStudNo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboStudNo.FormattingEnabled = True
        Me.cboStudNo.Location = New System.Drawing.Point(324, 48)
        Me.cboStudNo.Name = "cboStudNo"
        Me.cboStudNo.Size = New System.Drawing.Size(87, 21)
        Me.cboStudNo.TabIndex = 50
        '
        'txtTeacherName
        '
        Me.txtTeacherName.Location = New System.Drawing.Point(520, 16)
        Me.txtTeacherName.Name = "txtTeacherName"
        Me.txtTeacherName.ReadOnly = True
        Me.txtTeacherName.Size = New System.Drawing.Size(354, 20)
        Me.txtTeacherName.TabIndex = 49
        '
        'cboTeacherNo
        '
        Me.cboTeacherNo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboTeacherNo.FormattingEnabled = True
        Me.cboTeacherNo.Location = New System.Drawing.Point(324, 16)
        Me.cboTeacherNo.Name = "cboTeacherNo"
        Me.cboTeacherNo.Size = New System.Drawing.Size(87, 21)
        Me.cboTeacherNo.TabIndex = 48
        '
        'cboClassYear
        '
        Me.cboClassYear.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboClassYear.FormattingEnabled = True
        Me.cboClassYear.Location = New System.Drawing.Point(100, 81)
        Me.cboClassYear.Name = "cboClassYear"
        Me.cboClassYear.Size = New System.Drawing.Size(121, 21)
        Me.cboClassYear.TabIndex = 47
        '
        'cboClassStream
        '
        Me.cboClassStream.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboClassStream.FormattingEnabled = True
        Me.cboClassStream.Location = New System.Drawing.Point(100, 48)
        Me.cboClassStream.Name = "cboClassStream"
        Me.cboClassStream.Size = New System.Drawing.Size(121, 21)
        Me.cboClassStream.TabIndex = 46
        '
        'cboClassName
        '
        Me.cboClassName.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboClassName.FormattingEnabled = True
        Me.cboClassName.Location = New System.Drawing.Point(100, 16)
        Me.cboClassName.Name = "cboClassName"
        Me.cboClassName.Size = New System.Drawing.Size(121, 21)
        Me.cboClassName.TabIndex = 45
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(433, 51)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(78, 13)
        Me.Label7.TabIndex = 44
        Me.Label7.Text = "Student Name:"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(247, 51)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(64, 13)
        Me.Label6.TabIndex = 43
        Me.Label6.Text = "Student No:"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(433, 19)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(81, 13)
        Me.Label5.TabIndex = 42
        Me.Label5.Text = "Teacher Name:"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(247, 19)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(67, 13)
        Me.Label4.TabIndex = 41
        Me.Label4.Text = "Teacher No:"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(19, 84)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(60, 13)
        Me.Label3.TabIndex = 40
        Me.Label3.Text = "Class Year:"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(19, 51)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(71, 13)
        Me.Label2.TabIndex = 39
        Me.Label2.Text = "Class Stream:"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(19, 19)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(66, 13)
        Me.Label1.TabIndex = 38
        Me.Label1.Text = "Class Name:"
        '
        'StatusStrip5
        '
        Me.StatusStrip5.AutoSize = False
        Me.StatusStrip5.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip5.Dock = System.Windows.Forms.DockStyle.Right
        Me.StatusStrip5.Location = New System.Drawing.Point(883, 5)
        Me.StatusStrip5.Name = "StatusStrip5"
        Me.StatusStrip5.Size = New System.Drawing.Size(5, 102)
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
        Me.StatusStrip4.Size = New System.Drawing.Size(5, 102)
        Me.StatusStrip4.SizingGrip = False
        Me.StatusStrip4.TabIndex = 36
        Me.StatusStrip4.Text = "StatusStrip4"
        '
        'StatusStrip3
        '
        Me.StatusStrip3.AutoSize = False
        Me.StatusStrip3.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip3.Location = New System.Drawing.Point(0, 107)
        Me.StatusStrip3.Name = "StatusStrip3"
        Me.StatusStrip3.Size = New System.Drawing.Size(888, 5)
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
        Me.StatusStrip10.Size = New System.Drawing.Size(888, 5)
        Me.StatusStrip10.SizingGrip = False
        Me.StatusStrip10.TabIndex = 34
        Me.StatusStrip10.Text = "StatusStrip10"
        '
        'lstClassHeads
        '
        Me.lstClassHeads.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColumnHeader1, Me.ColumnHeader2, Me.ColumnHeader3, Me.ColumnHeader4, Me.ColumnHeader5, Me.ColumnHeader6, Me.ColumnHeader7})
        Me.lstClassHeads.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lstClassHeads.FullRowSelect = True
        Me.lstClassHeads.GridLines = True
        Me.lstClassHeads.Location = New System.Drawing.Point(0, 0)
        Me.lstClassHeads.Name = "lstClassHeads"
        Me.lstClassHeads.Size = New System.Drawing.Size(888, 326)
        Me.lstClassHeads.TabIndex = 26
        Me.lstClassHeads.UseCompatibleStateImageBehavior = False
        Me.lstClassHeads.View = System.Windows.Forms.View.Details
        '
        'ColumnHeader1
        '
        Me.ColumnHeader1.Text = "Class Name"
        Me.ColumnHeader1.Width = 80
        '
        'ColumnHeader2
        '
        Me.ColumnHeader2.Text = "Stream"
        Me.ColumnHeader2.Width = 50
        '
        'ColumnHeader3
        '
        Me.ColumnHeader3.Text = "Class Year"
        Me.ColumnHeader3.Width = 70
        '
        'ColumnHeader4
        '
        Me.ColumnHeader4.Text = "Stud No"
        Me.ColumnHeader4.Width = 70
        '
        'ColumnHeader5
        '
        Me.ColumnHeader5.Text = "Student Name"
        Me.ColumnHeader5.Width = 270
        '
        'ColumnHeader6
        '
        Me.ColumnHeader6.Text = "Teacher No"
        Me.ColumnHeader6.Width = 70
        '
        'ColumnHeader7
        '
        Me.ColumnHeader7.Text = "Teacher Name"
        Me.ColumnHeader7.Width = 270
        '
        'StatusStrip2
        '
        Me.StatusStrip2.AutoSize = False
        Me.StatusStrip2.BackColor = System.Drawing.Color.LightSeaGreen
        Me.StatusStrip2.Dock = System.Windows.Forms.DockStyle.Top
        Me.StatusStrip2.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripStatusLabel1})
        Me.StatusStrip2.Location = New System.Drawing.Point(0, 0)
        Me.StatusStrip2.Name = "StatusStrip2"
        Me.StatusStrip2.Size = New System.Drawing.Size(890, 25)
        Me.StatusStrip2.SizingGrip = False
        Me.StatusStrip2.TabIndex = 22
        Me.StatusStrip2.Text = "StatusStrip2"
        '
        'ToolStripStatusLabel1
        '
        Me.ToolStripStatusLabel1.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ToolStripStatusLabel1.Name = "ToolStripStatusLabel1"
        Me.ToolStripStatusLabel1.Size = New System.Drawing.Size(143, 20)
        Me.ToolStripStatusLabel1.Text = "ASSIGN CLASS HEADS"
        '
        'StatusStrip1
        '
        Me.StatusStrip1.AutoSize = False
        Me.StatusStrip1.BackColor = System.Drawing.Color.LightSeaGreen
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 468)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(890, 25)
        Me.StatusStrip1.SizingGrip = False
        Me.StatusStrip1.TabIndex = 21
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'frmClassHeads
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.LightSkyBlue
        Me.ClientSize = New System.Drawing.Size(916, 519)
        Me.Controls.Add(Me.pnlClassHeads)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Name = "frmClassHeads"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Class Heads"
        Me.pnlClassHeads.ResumeLayout(False)
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel1.PerformLayout()
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.ResumeLayout(False)
        Me.StatusStrip2.ResumeLayout(False)
        Me.StatusStrip2.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents pnlClassHeads As System.Windows.Forms.Panel
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents StatusStrip2 As System.Windows.Forms.StatusStrip
    Friend WithEvents ToolStripStatusLabel1 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents lstClassHeads As System.Windows.Forms.ListView
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents StatusStrip5 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip4 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip3 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip10 As System.Windows.Forms.StatusStrip
    Friend WithEvents txtStudName As System.Windows.Forms.TextBox
    Friend WithEvents txtTeacherName As System.Windows.Forms.TextBox
    Friend WithEvents cboClassYear As System.Windows.Forms.ComboBox
    Friend WithEvents cboClassStream As System.Windows.Forms.ComboBox
    Friend WithEvents cboClassName As System.Windows.Forms.ComboBox
    Friend WithEvents btnView As System.Windows.Forms.Button
    Friend WithEvents btnDelete As System.Windows.Forms.Button
    Friend WithEvents btnUpdate As System.Windows.Forms.Button
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents ColumnHeader1 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader2 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader3 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader4 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader5 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader6 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader7 As System.Windows.Forms.ColumnHeader
    Friend WithEvents cboStudNo As System.Windows.Forms.ComboBox
    Friend WithEvents cboTeacherNo As System.Windows.Forms.ComboBox
End Class
