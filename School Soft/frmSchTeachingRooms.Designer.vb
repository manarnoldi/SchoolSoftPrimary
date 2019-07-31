<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmSchTeachingRooms
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
        Me.pnlTeachingRooms = New System.Windows.Forms.Panel()
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
        Me.txtRoomName = New System.Windows.Forms.TextBox()
        Me.cboRoomType = New System.Windows.Forms.ComboBox()
        Me.rbTrue = New System.Windows.Forms.RadioButton()
        Me.rbFalse = New System.Windows.Forms.RadioButton()
        Me.txtRoomCapacity = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.StatusStrip5 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip4 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip3 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip10 = New System.Windows.Forms.StatusStrip()
        Me.SplitContainer2 = New System.Windows.Forms.SplitContainer()
        Me.lstTeachingRooms = New System.Windows.Forms.ListView()
        Me.ColumnHeader1 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader2 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader3 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader4 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.btnView = New System.Windows.Forms.Button()
        Me.btnDelete = New System.Windows.Forms.Button()
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
        Me.pnlTeachingRooms.SuspendLayout()
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        CType(Me.SplitContainer2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer2.Panel1.SuspendLayout()
        Me.SplitContainer2.Panel2.SuspendLayout()
        Me.SplitContainer2.SuspendLayout()
        Me.StatusStrip9.SuspendLayout()
        Me.StatusStrip2.SuspendLayout()
        Me.SuspendLayout()
        '
        'pnlTeachingRooms
        '
        Me.pnlTeachingRooms.BackColor = System.Drawing.SystemColors.Control
        Me.pnlTeachingRooms.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pnlTeachingRooms.Controls.Add(Me.SplitContainer1)
        Me.pnlTeachingRooms.Controls.Add(Me.StatusStrip2)
        Me.pnlTeachingRooms.Controls.Add(Me.StatusStrip1)
        Me.pnlTeachingRooms.Location = New System.Drawing.Point(12, 12)
        Me.pnlTeachingRooms.Name = "pnlTeachingRooms"
        Me.pnlTeachingRooms.Size = New System.Drawing.Size(554, 478)
        Me.pnlTeachingRooms.TabIndex = 0
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
        Me.SplitContainer1.Panel1.Controls.Add(Me.txtRoomName)
        Me.SplitContainer1.Panel1.Controls.Add(Me.cboRoomType)
        Me.SplitContainer1.Panel1.Controls.Add(Me.rbTrue)
        Me.SplitContainer1.Panel1.Controls.Add(Me.rbFalse)
        Me.SplitContainer1.Panel1.Controls.Add(Me.txtRoomCapacity)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label5)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label4)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label3)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label1)
        Me.SplitContainer1.Panel1.Controls.Add(Me.StatusStrip5)
        Me.SplitContainer1.Panel1.Controls.Add(Me.StatusStrip4)
        Me.SplitContainer1.Panel1.Controls.Add(Me.StatusStrip3)
        Me.SplitContainer1.Panel1.Controls.Add(Me.StatusStrip10)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.SplitContainer2)
        Me.SplitContainer1.Size = New System.Drawing.Size(552, 426)
        Me.SplitContainer1.SplitterDistance = 103
        Me.SplitContainer1.SplitterWidth = 1
        Me.SplitContainer1.TabIndex = 19
        '
        'txtRoomName
        '
        Me.txtRoomName.Location = New System.Drawing.Point(90, 59)
        Me.txtRoomName.Name = "txtRoomName"
        Me.txtRoomName.Size = New System.Drawing.Size(221, 20)
        Me.txtRoomName.TabIndex = 35
        '
        'cboRoomType
        '
        Me.cboRoomType.FormattingEnabled = True
        Me.cboRoomType.Location = New System.Drawing.Point(90, 23)
        Me.cboRoomType.Name = "cboRoomType"
        Me.cboRoomType.Size = New System.Drawing.Size(221, 21)
        Me.cboRoomType.TabIndex = 34
        '
        'rbTrue
        '
        Me.rbTrue.AutoSize = True
        Me.rbTrue.Location = New System.Drawing.Point(410, 24)
        Me.rbTrue.Name = "rbTrue"
        Me.rbTrue.Size = New System.Drawing.Size(47, 17)
        Me.rbTrue.TabIndex = 33
        Me.rbTrue.TabStop = True
        Me.rbTrue.Text = "True"
        Me.rbTrue.UseVisualStyleBackColor = True
        '
        'rbFalse
        '
        Me.rbFalse.AutoSize = True
        Me.rbFalse.Location = New System.Drawing.Point(480, 24)
        Me.rbFalse.Name = "rbFalse"
        Me.rbFalse.Size = New System.Drawing.Size(50, 17)
        Me.rbFalse.TabIndex = 32
        Me.rbFalse.TabStop = True
        Me.rbFalse.Text = "False"
        Me.rbFalse.UseVisualStyleBackColor = True
        '
        'txtRoomCapacity
        '
        Me.txtRoomCapacity.Location = New System.Drawing.Point(405, 59)
        Me.txtRoomCapacity.Name = "txtRoomCapacity"
        Me.txtRoomCapacity.Size = New System.Drawing.Size(129, 20)
        Me.txtRoomCapacity.TabIndex = 31
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(317, 62)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(82, 13)
        Me.Label5.TabIndex = 30
        Me.Label5.Text = "Room Capacity:"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(19, 62)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(69, 13)
        Me.Label4.TabIndex = 29
        Me.Label4.Text = "Room Name:"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(317, 26)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(83, 13)
        Me.Label3.TabIndex = 28
        Me.Label3.Text = "Room Sharable:"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(19, 26)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(65, 13)
        Me.Label1.TabIndex = 26
        Me.Label1.Text = "Room Type:"
        '
        'StatusStrip5
        '
        Me.StatusStrip5.AutoSize = False
        Me.StatusStrip5.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip5.Dock = System.Windows.Forms.DockStyle.Right
        Me.StatusStrip5.Location = New System.Drawing.Point(542, 10)
        Me.StatusStrip5.Name = "StatusStrip5"
        Me.StatusStrip5.Size = New System.Drawing.Size(10, 83)
        Me.StatusStrip5.SizingGrip = False
        Me.StatusStrip5.TabIndex = 25
        Me.StatusStrip5.Text = "StatusStrip5"
        '
        'StatusStrip4
        '
        Me.StatusStrip4.AutoSize = False
        Me.StatusStrip4.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip4.Dock = System.Windows.Forms.DockStyle.Left
        Me.StatusStrip4.Location = New System.Drawing.Point(0, 10)
        Me.StatusStrip4.Name = "StatusStrip4"
        Me.StatusStrip4.Size = New System.Drawing.Size(10, 83)
        Me.StatusStrip4.SizingGrip = False
        Me.StatusStrip4.TabIndex = 24
        Me.StatusStrip4.Text = "StatusStrip4"
        '
        'StatusStrip3
        '
        Me.StatusStrip3.AutoSize = False
        Me.StatusStrip3.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip3.Location = New System.Drawing.Point(0, 93)
        Me.StatusStrip3.Name = "StatusStrip3"
        Me.StatusStrip3.Size = New System.Drawing.Size(552, 10)
        Me.StatusStrip3.SizingGrip = False
        Me.StatusStrip3.TabIndex = 23
        Me.StatusStrip3.Text = "StatusStrip3"
        '
        'StatusStrip10
        '
        Me.StatusStrip10.AutoSize = False
        Me.StatusStrip10.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip10.Dock = System.Windows.Forms.DockStyle.Top
        Me.StatusStrip10.Location = New System.Drawing.Point(0, 0)
        Me.StatusStrip10.Name = "StatusStrip10"
        Me.StatusStrip10.Size = New System.Drawing.Size(552, 10)
        Me.StatusStrip10.SizingGrip = False
        Me.StatusStrip10.TabIndex = 22
        Me.StatusStrip10.Text = "StatusStrip10"
        '
        'SplitContainer2
        '
        Me.SplitContainer2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer2.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer2.Name = "SplitContainer2"
        Me.SplitContainer2.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'SplitContainer2.Panel1
        '
        Me.SplitContainer2.Panel1.Controls.Add(Me.lstTeachingRooms)
        '
        'SplitContainer2.Panel2
        '
        Me.SplitContainer2.Panel2.Controls.Add(Me.btnView)
        Me.SplitContainer2.Panel2.Controls.Add(Me.btnDelete)
        Me.SplitContainer2.Panel2.Controls.Add(Me.btnUpdate)
        Me.SplitContainer2.Panel2.Controls.Add(Me.btnSave)
        Me.SplitContainer2.Panel2.Controls.Add(Me.btnClose)
        Me.SplitContainer2.Panel2.Controls.Add(Me.StatusStrip6)
        Me.SplitContainer2.Panel2.Controls.Add(Me.StatusStrip7)
        Me.SplitContainer2.Panel2.Controls.Add(Me.StatusStrip8)
        Me.SplitContainer2.Panel2.Controls.Add(Me.StatusStrip9)
        Me.SplitContainer2.Size = New System.Drawing.Size(552, 322)
        Me.SplitContainer2.SplitterDistance = 259
        Me.SplitContainer2.SplitterWidth = 1
        Me.SplitContainer2.TabIndex = 20
        '
        'lstTeachingRooms
        '
        Me.lstTeachingRooms.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColumnHeader1, Me.ColumnHeader2, Me.ColumnHeader3, Me.ColumnHeader4})
        Me.lstTeachingRooms.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lstTeachingRooms.FullRowSelect = True
        Me.lstTeachingRooms.GridLines = True
        Me.lstTeachingRooms.Location = New System.Drawing.Point(0, 0)
        Me.lstTeachingRooms.Name = "lstTeachingRooms"
        Me.lstTeachingRooms.Size = New System.Drawing.Size(552, 259)
        Me.lstTeachingRooms.TabIndex = 23
        Me.lstTeachingRooms.UseCompatibleStateImageBehavior = False
        Me.lstTeachingRooms.View = System.Windows.Forms.View.Details
        '
        'ColumnHeader1
        '
        Me.ColumnHeader1.Text = "ROOM TYPE"
        Me.ColumnHeader1.Width = 145
        '
        'ColumnHeader2
        '
        Me.ColumnHeader2.Text = "ROOM NAME"
        Me.ColumnHeader2.Width = 235
        '
        'ColumnHeader3
        '
        Me.ColumnHeader3.Text = "CAPACITY"
        Me.ColumnHeader3.Width = 70
        '
        'ColumnHeader4
        '
        Me.ColumnHeader4.Text = "SHARABLE"
        Me.ColumnHeader4.Width = 80
        '
        'btnView
        '
        Me.btnView.Location = New System.Drawing.Point(15, 24)
        Me.btnView.Name = "btnView"
        Me.btnView.Size = New System.Drawing.Size(75, 23)
        Me.btnView.TabIndex = 37
        Me.btnView.Text = "View"
        Me.btnView.UseVisualStyleBackColor = True
        '
        'btnDelete
        '
        Me.btnDelete.Enabled = False
        Me.btnDelete.Location = New System.Drawing.Point(348, 24)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(75, 23)
        Me.btnDelete.TabIndex = 36
        Me.btnDelete.Text = "Delete"
        Me.btnDelete.UseVisualStyleBackColor = True
        '
        'btnUpdate
        '
        Me.btnUpdate.Enabled = False
        Me.btnUpdate.Location = New System.Drawing.Point(235, 24)
        Me.btnUpdate.Name = "btnUpdate"
        Me.btnUpdate.Size = New System.Drawing.Size(75, 23)
        Me.btnUpdate.TabIndex = 35
        Me.btnUpdate.Text = "Update"
        Me.btnUpdate.UseVisualStyleBackColor = True
        '
        'btnSave
        '
        Me.btnSave.Location = New System.Drawing.Point(123, 24)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(75, 23)
        Me.btnSave.TabIndex = 34
        Me.btnSave.Text = "Save"
        Me.btnSave.UseVisualStyleBackColor = True
        '
        'btnClose
        '
        Me.btnClose.Location = New System.Drawing.Point(462, 24)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(75, 23)
        Me.btnClose.TabIndex = 33
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'StatusStrip6
        '
        Me.StatusStrip6.AutoSize = False
        Me.StatusStrip6.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip6.Dock = System.Windows.Forms.DockStyle.Right
        Me.StatusStrip6.Location = New System.Drawing.Point(542, 20)
        Me.StatusStrip6.Name = "StatusStrip6"
        Me.StatusStrip6.Size = New System.Drawing.Size(10, 32)
        Me.StatusStrip6.SizingGrip = False
        Me.StatusStrip6.TabIndex = 31
        Me.StatusStrip6.Text = "StatusStrip6"
        '
        'StatusStrip7
        '
        Me.StatusStrip7.AutoSize = False
        Me.StatusStrip7.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip7.Dock = System.Windows.Forms.DockStyle.Left
        Me.StatusStrip7.Location = New System.Drawing.Point(0, 20)
        Me.StatusStrip7.Name = "StatusStrip7"
        Me.StatusStrip7.Size = New System.Drawing.Size(10, 32)
        Me.StatusStrip7.SizingGrip = False
        Me.StatusStrip7.TabIndex = 30
        Me.StatusStrip7.Text = "StatusStrip7"
        '
        'StatusStrip8
        '
        Me.StatusStrip8.AutoSize = False
        Me.StatusStrip8.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip8.Location = New System.Drawing.Point(0, 52)
        Me.StatusStrip8.Name = "StatusStrip8"
        Me.StatusStrip8.Size = New System.Drawing.Size(552, 10)
        Me.StatusStrip8.SizingGrip = False
        Me.StatusStrip8.TabIndex = 29
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
        Me.StatusStrip9.Size = New System.Drawing.Size(552, 20)
        Me.StatusStrip9.SizingGrip = False
        Me.StatusStrip9.TabIndex = 28
        Me.StatusStrip9.Text = "StatusStrip9"
        '
        'ToolStripStatusLabel2
        '
        Me.ToolStripStatusLabel2.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ToolStripStatusLabel2.Name = "ToolStripStatusLabel2"
        Me.ToolStripStatusLabel2.Size = New System.Drawing.Size(77, 15)
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
        Me.StatusStrip2.Size = New System.Drawing.Size(552, 25)
        Me.StatusStrip2.SizingGrip = False
        Me.StatusStrip2.TabIndex = 18
        Me.StatusStrip2.Text = "StatusStrip2"
        '
        'ToolStripStatusLabel1
        '
        Me.ToolStripStatusLabel1.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ToolStripStatusLabel1.Name = "ToolStripStatusLabel1"
        Me.ToolStripStatusLabel1.Size = New System.Drawing.Size(180, 20)
        Me.ToolStripStatusLabel1.Text = "SCHOOL TEACHING ROOMS"
        '
        'StatusStrip1
        '
        Me.StatusStrip1.AutoSize = False
        Me.StatusStrip1.BackColor = System.Drawing.Color.LightSeaGreen
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 451)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(552, 25)
        Me.StatusStrip1.SizingGrip = False
        Me.StatusStrip1.TabIndex = 17
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'frmSchTeachingRooms
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.LightBlue
        Me.ClientSize = New System.Drawing.Size(576, 502)
        Me.Controls.Add(Me.pnlTeachingRooms)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Name = "frmSchTeachingRooms"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Teaching Rooms"
        Me.pnlTeachingRooms.ResumeLayout(False)
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel1.PerformLayout()
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.ResumeLayout(False)
        Me.SplitContainer2.Panel1.ResumeLayout(False)
        Me.SplitContainer2.Panel2.ResumeLayout(False)
        CType(Me.SplitContainer2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer2.ResumeLayout(False)
        Me.StatusStrip9.ResumeLayout(False)
        Me.StatusStrip9.PerformLayout()
        Me.StatusStrip2.ResumeLayout(False)
        Me.StatusStrip2.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents pnlTeachingRooms As System.Windows.Forms.Panel
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents SplitContainer2 As System.Windows.Forms.SplitContainer
    Friend WithEvents StatusStrip2 As System.Windows.Forms.StatusStrip
    Friend WithEvents ToolStripStatusLabel1 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents lstTeachingRooms As System.Windows.Forms.ListView
    Friend WithEvents ColumnHeader1 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader2 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader3 As System.Windows.Forms.ColumnHeader
    Friend WithEvents StatusStrip5 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip4 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip3 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip10 As System.Windows.Forms.StatusStrip
    Friend WithEvents txtRoomName As System.Windows.Forms.TextBox
    Friend WithEvents cboRoomType As System.Windows.Forms.ComboBox
    Friend WithEvents rbTrue As System.Windows.Forms.RadioButton
    Friend WithEvents rbFalse As System.Windows.Forms.RadioButton
    Friend WithEvents txtRoomCapacity As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents StatusStrip6 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip7 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip8 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip9 As System.Windows.Forms.StatusStrip
    Friend WithEvents ToolStripStatusLabel2 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents btnView As System.Windows.Forms.Button
    Friend WithEvents btnDelete As System.Windows.Forms.Button
    Friend WithEvents btnUpdate As System.Windows.Forms.Button
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents ColumnHeader4 As System.Windows.Forms.ColumnHeader
End Class
