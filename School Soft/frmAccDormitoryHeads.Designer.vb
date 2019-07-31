<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmAccDormitoryHeads
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
        Me.pnlDormHeads = New System.Windows.Forms.Panel()
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.txtBorder = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.txtStudName = New System.Windows.Forms.TextBox()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.cboStudNo = New System.Windows.Forms.ComboBox()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.txtEmpName = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.cboEmpNo = New System.Windows.Forms.ComboBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.cboEmpType = New System.Windows.Forms.ComboBox()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.cboYear = New System.Windows.Forms.ComboBox()
        Me.btnClose = New System.Windows.Forms.Button()
        Me.btnUpdate = New System.Windows.Forms.Button()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.btnSave = New System.Windows.Forms.Button()
        Me.btnView = New System.Windows.Forms.Button()
        Me.cboDormName = New System.Windows.Forms.ComboBox()
        Me.StatusStrip12 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip13 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip14 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip5 = New System.Windows.Forms.StatusStrip()
        Me.lstDormDetails = New System.Windows.Forms.ListView()
        Me.ColumnHeader1 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader2 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader3 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader4 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader5 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.StatusStrip2 = New System.Windows.Forms.StatusStrip()
        Me.ToolStripStatusLabel1 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.ContextMenuStrip1 = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.UPDATEToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.DELETEToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.CLOSEToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.pnlDormHeads.SuspendLayout()
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.StatusStrip2.SuspendLayout()
        Me.ContextMenuStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'pnlDormHeads
        '
        Me.pnlDormHeads.BackColor = System.Drawing.SystemColors.Control
        Me.pnlDormHeads.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pnlDormHeads.Controls.Add(Me.SplitContainer1)
        Me.pnlDormHeads.Controls.Add(Me.StatusStrip2)
        Me.pnlDormHeads.Location = New System.Drawing.Point(12, 12)
        Me.pnlDormHeads.Name = "pnlDormHeads"
        Me.pnlDormHeads.Size = New System.Drawing.Size(884, 515)
        Me.pnlDormHeads.TabIndex = 5
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
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.lstDormDetails)
        Me.SplitContainer1.Size = New System.Drawing.Size(882, 488)
        Me.SplitContainer1.SplitterDistance = 140
        Me.SplitContainer1.SplitterWidth = 1
        Me.SplitContainer1.TabIndex = 23
        '
        'Panel1
        '
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel1.Controls.Add(Me.Label9)
        Me.Panel1.Controls.Add(Me.GroupBox2)
        Me.Panel1.Controls.Add(Me.GroupBox1)
        Me.Panel1.Controls.Add(Me.Panel2)
        Me.Panel1.Controls.Add(Me.StatusStrip12)
        Me.Panel1.Controls.Add(Me.StatusStrip13)
        Me.Panel1.Controls.Add(Me.StatusStrip14)
        Me.Panel1.Controls.Add(Me.StatusStrip5)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel1.Location = New System.Drawing.Point(0, 0)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(882, 140)
        Me.Panel1.TabIndex = 74
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(17, 19)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(72, 13)
        Me.Label9.TabIndex = 3
        Me.Label9.Text = "Dorm Name : "
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.txtBorder)
        Me.GroupBox2.Controls.Add(Me.Label4)
        Me.GroupBox2.Controls.Add(Me.txtStudName)
        Me.GroupBox2.Controls.Add(Me.Label8)
        Me.GroupBox2.Controls.Add(Me.Label7)
        Me.GroupBox2.Controls.Add(Me.cboStudNo)
        Me.GroupBox2.Location = New System.Drawing.Point(424, 11)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(441, 79)
        Me.GroupBox2.TabIndex = 1
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Student Head Details :"
        '
        'txtBorder
        '
        Me.txtBorder.Location = New System.Drawing.Point(313, 18)
        Me.txtBorder.Name = "txtBorder"
        Me.txtBorder.ReadOnly = True
        Me.txtBorder.Size = New System.Drawing.Size(117, 20)
        Me.txtBorder.TabIndex = 10
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(254, 21)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(53, 13)
        Me.Label4.TabIndex = 11
        Me.Label4.Text = "Boarder : "
        '
        'txtStudName
        '
        Me.txtStudName.Location = New System.Drawing.Point(96, 50)
        Me.txtStudName.Name = "txtStudName"
        Me.txtStudName.ReadOnly = True
        Me.txtStudName.Size = New System.Drawing.Size(334, 20)
        Me.txtStudName.TabIndex = 5
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(6, 53)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(84, 13)
        Me.Label8.TabIndex = 9
        Me.Label8.Text = "Student Name : "
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(6, 21)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(78, 13)
        Me.Label7.TabIndex = 8
        Me.Label7.Text = "Stud Number : "
        '
        'cboStudNo
        '
        Me.cboStudNo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboStudNo.FormattingEnabled = True
        Me.cboStudNo.Location = New System.Drawing.Point(96, 18)
        Me.cboStudNo.Name = "cboStudNo"
        Me.cboStudNo.Size = New System.Drawing.Size(139, 21)
        Me.cboStudNo.TabIndex = 7
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.txtEmpName)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.cboEmpNo)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Controls.Add(Me.cboEmpType)
        Me.GroupBox1.Location = New System.Drawing.Point(14, 42)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(395, 81)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Staff Head Details :"
        '
        'txtEmpName
        '
        Me.txtEmpName.Location = New System.Drawing.Point(80, 51)
        Me.txtEmpName.Name = "txtEmpName"
        Me.txtEmpName.ReadOnly = True
        Me.txtEmpName.Size = New System.Drawing.Size(308, 20)
        Me.txtEmpName.TabIndex = 4
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(10, 54)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(68, 13)
        Me.Label3.TabIndex = 3
        Me.Label3.Text = "Emp Name : "
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(206, 22)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(54, 13)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "Emp No : "
        '
        'cboEmpNo
        '
        Me.cboEmpNo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboEmpNo.FormattingEnabled = True
        Me.cboEmpNo.Location = New System.Drawing.Point(267, 19)
        Me.cboEmpNo.Name = "cboEmpNo"
        Me.cboEmpNo.Size = New System.Drawing.Size(121, 21)
        Me.cboEmpNo.TabIndex = 1
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(10, 22)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(64, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Emp Type : "
        '
        'cboEmpType
        '
        Me.cboEmpType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboEmpType.FormattingEnabled = True
        Me.cboEmpType.Location = New System.Drawing.Point(80, 19)
        Me.cboEmpType.Name = "cboEmpType"
        Me.cboEmpType.Size = New System.Drawing.Size(121, 21)
        Me.cboEmpType.TabIndex = 0
        '
        'Panel2
        '
        Me.Panel2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel2.Controls.Add(Me.cboYear)
        Me.Panel2.Controls.Add(Me.btnClose)
        Me.Panel2.Controls.Add(Me.btnUpdate)
        Me.Panel2.Controls.Add(Me.Label10)
        Me.Panel2.Controls.Add(Me.btnSave)
        Me.Panel2.Controls.Add(Me.btnView)
        Me.Panel2.Controls.Add(Me.cboDormName)
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel2.Location = New System.Drawing.Point(5, 5)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(870, 128)
        Me.Panel2.TabIndex = 65
        '
        'cboYear
        '
        Me.cboYear.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboYear.FormattingEnabled = True
        Me.cboYear.Location = New System.Drawing.Point(316, 10)
        Me.cboYear.Name = "cboYear"
        Me.cboYear.Size = New System.Drawing.Size(87, 21)
        Me.cboYear.TabIndex = 5
        '
        'btnClose
        '
        Me.btnClose.Location = New System.Drawing.Point(784, 94)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(75, 23)
        Me.btnClose.TabIndex = 0
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'btnUpdate
        '
        Me.btnUpdate.Enabled = False
        Me.btnUpdate.Location = New System.Drawing.Point(662, 94)
        Me.btnUpdate.Name = "btnUpdate"
        Me.btnUpdate.Size = New System.Drawing.Size(75, 23)
        Me.btnUpdate.TabIndex = 2
        Me.btnUpdate.Text = "Update"
        Me.btnUpdate.UseVisualStyleBackColor = True
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(272, 13)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(38, 13)
        Me.Label10.TabIndex = 6
        Me.Label10.Text = "Year : "
        '
        'btnSave
        '
        Me.btnSave.Location = New System.Drawing.Point(538, 94)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(75, 23)
        Me.btnSave.TabIndex = 1
        Me.btnSave.Text = "Save"
        Me.btnSave.UseVisualStyleBackColor = True
        '
        'btnView
        '
        Me.btnView.Location = New System.Drawing.Point(418, 94)
        Me.btnView.Name = "btnView"
        Me.btnView.Size = New System.Drawing.Size(75, 23)
        Me.btnView.TabIndex = 3
        Me.btnView.Text = "View"
        Me.btnView.UseVisualStyleBackColor = True
        '
        'cboDormName
        '
        Me.cboDormName.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboDormName.FormattingEnabled = True
        Me.cboDormName.Location = New System.Drawing.Point(88, 10)
        Me.cboDormName.Name = "cboDormName"
        Me.cboDormName.Size = New System.Drawing.Size(175, 21)
        Me.cboDormName.TabIndex = 5
        '
        'StatusStrip12
        '
        Me.StatusStrip12.AutoSize = False
        Me.StatusStrip12.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip12.Dock = System.Windows.Forms.DockStyle.Left
        Me.StatusStrip12.Location = New System.Drawing.Point(0, 5)
        Me.StatusStrip12.Name = "StatusStrip12"
        Me.StatusStrip12.Size = New System.Drawing.Size(5, 128)
        Me.StatusStrip12.SizingGrip = False
        Me.StatusStrip12.TabIndex = 64
        Me.StatusStrip12.Text = "StatusStrip12"
        '
        'StatusStrip13
        '
        Me.StatusStrip13.AutoSize = False
        Me.StatusStrip13.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip13.Location = New System.Drawing.Point(0, 133)
        Me.StatusStrip13.Name = "StatusStrip13"
        Me.StatusStrip13.Size = New System.Drawing.Size(875, 5)
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
        Me.StatusStrip14.Size = New System.Drawing.Size(875, 5)
        Me.StatusStrip14.SizingGrip = False
        Me.StatusStrip14.TabIndex = 62
        Me.StatusStrip14.Text = "StatusStrip14"
        '
        'StatusStrip5
        '
        Me.StatusStrip5.AutoSize = False
        Me.StatusStrip5.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip5.Dock = System.Windows.Forms.DockStyle.Right
        Me.StatusStrip5.Location = New System.Drawing.Point(875, 0)
        Me.StatusStrip5.Name = "StatusStrip5"
        Me.StatusStrip5.Size = New System.Drawing.Size(5, 138)
        Me.StatusStrip5.SizingGrip = False
        Me.StatusStrip5.TabIndex = 61
        Me.StatusStrip5.Text = "StatusStrip5"
        '
        'lstDormDetails
        '
        Me.lstDormDetails.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lstDormDetails.CheckBoxes = True
        Me.lstDormDetails.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColumnHeader1, Me.ColumnHeader2, Me.ColumnHeader3, Me.ColumnHeader4, Me.ColumnHeader5})
        Me.lstDormDetails.ContextMenuStrip = Me.ContextMenuStrip1
        Me.lstDormDetails.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lstDormDetails.FullRowSelect = True
        Me.lstDormDetails.GridLines = True
        Me.lstDormDetails.Location = New System.Drawing.Point(0, 0)
        Me.lstDormDetails.Name = "lstDormDetails"
        Me.lstDormDetails.Size = New System.Drawing.Size(882, 347)
        Me.lstDormDetails.TabIndex = 31
        Me.lstDormDetails.UseCompatibleStateImageBehavior = False
        Me.lstDormDetails.View = System.Windows.Forms.View.Details
        '
        'ColumnHeader1
        '
        Me.ColumnHeader1.Text = "Dorm Name"
        Me.ColumnHeader1.Width = 190
        '
        'ColumnHeader2
        '
        Me.ColumnHeader2.Text = "Staff No"
        Me.ColumnHeader2.Width = 80
        '
        'ColumnHeader3
        '
        Me.ColumnHeader3.Text = "Staff Name"
        Me.ColumnHeader3.Width = 250
        '
        'ColumnHeader4
        '
        Me.ColumnHeader4.Text = "Adm No"
        Me.ColumnHeader4.Width = 80
        '
        'ColumnHeader5
        '
        Me.ColumnHeader5.Text = "Student Name"
        Me.ColumnHeader5.Width = 250
        '
        'StatusStrip2
        '
        Me.StatusStrip2.AutoSize = False
        Me.StatusStrip2.BackColor = System.Drawing.Color.LightSeaGreen
        Me.StatusStrip2.Dock = System.Windows.Forms.DockStyle.Top
        Me.StatusStrip2.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripStatusLabel1})
        Me.StatusStrip2.Location = New System.Drawing.Point(0, 0)
        Me.StatusStrip2.Name = "StatusStrip2"
        Me.StatusStrip2.Size = New System.Drawing.Size(882, 25)
        Me.StatusStrip2.SizingGrip = False
        Me.StatusStrip2.TabIndex = 22
        Me.StatusStrip2.Text = "StatusStrip2"
        '
        'ToolStripStatusLabel1
        '
        Me.ToolStripStatusLabel1.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ToolStripStatusLabel1.Name = "ToolStripStatusLabel1"
        Me.ToolStripStatusLabel1.Size = New System.Drawing.Size(244, 20)
        Me.ToolStripStatusLabel1.Text = "ACCOMODATION DORMITORY HEADS"
        '
        'ContextMenuStrip1
        '
        Me.ContextMenuStrip1.BackColor = System.Drawing.Color.LightSeaGreen
        Me.ContextMenuStrip1.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ContextMenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.UPDATEToolStripMenuItem, Me.DELETEToolStripMenuItem, Me.CLOSEToolStripMenuItem})
        Me.ContextMenuStrip1.Name = "ContextMenuStrip1"
        Me.ContextMenuStrip1.Size = New System.Drawing.Size(153, 92)
        '
        'UPDATEToolStripMenuItem
        '
        Me.UPDATEToolStripMenuItem.Name = "UPDATEToolStripMenuItem"
        Me.UPDATEToolStripMenuItem.Size = New System.Drawing.Size(152, 22)
        Me.UPDATEToolStripMenuItem.Text = "EDIT"
        '
        'DELETEToolStripMenuItem
        '
        Me.DELETEToolStripMenuItem.Name = "DELETEToolStripMenuItem"
        Me.DELETEToolStripMenuItem.Size = New System.Drawing.Size(152, 22)
        Me.DELETEToolStripMenuItem.Text = "DELETE"
        '
        'CLOSEToolStripMenuItem
        '
        Me.CLOSEToolStripMenuItem.Name = "CLOSEToolStripMenuItem"
        Me.CLOSEToolStripMenuItem.Size = New System.Drawing.Size(152, 22)
        Me.CLOSEToolStripMenuItem.Text = "CLOSE"
        '
        'frmAccDormitoryHeads
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.LightSkyBlue
        Me.ClientSize = New System.Drawing.Size(906, 539)
        Me.Controls.Add(Me.pnlDormHeads)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.MaximizeBox = False
        Me.Name = "frmAccDormitoryHeads"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.Text = "Dormitory Heads"
        Me.pnlDormHeads.ResumeLayout(False)
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.ResumeLayout(False)
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.Panel2.ResumeLayout(False)
        Me.Panel2.PerformLayout()
        Me.StatusStrip2.ResumeLayout(False)
        Me.StatusStrip2.PerformLayout()
        Me.ContextMenuStrip1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents pnlDormHeads As System.Windows.Forms.Panel
    Friend WithEvents StatusStrip2 As System.Windows.Forms.StatusStrip
    Friend WithEvents ToolStripStatusLabel1 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents btnUpdate As System.Windows.Forms.Button
    Friend WithEvents btnView As System.Windows.Forms.Button
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents cboDormName As System.Windows.Forms.ComboBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents txtBorder As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents txtStudName As System.Windows.Forms.TextBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents cboStudNo As System.Windows.Forms.ComboBox
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents txtEmpName As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents cboEmpNo As System.Windows.Forms.ComboBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents cboEmpType As System.Windows.Forms.ComboBox
    Friend WithEvents lstDormDetails As System.Windows.Forms.ListView
    Friend WithEvents ColumnHeader1 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader2 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader3 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader4 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader5 As System.Windows.Forms.ColumnHeader
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents StatusStrip12 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip13 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip14 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip5 As System.Windows.Forms.StatusStrip
    Friend WithEvents cboYear As System.Windows.Forms.ComboBox
    Friend WithEvents ContextMenuStrip1 As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents UPDATEToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents DELETEToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents CLOSEToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
End Class
