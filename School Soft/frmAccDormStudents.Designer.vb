<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmAccDormStudents
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
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Panel5 = New System.Windows.Forms.Panel()
        Me.cboHouse = New System.Windows.Forms.ComboBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.txtGender = New System.Windows.Forms.TextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.txtStudentName = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.btnSave = New System.Windows.Forms.Button()
        Me.btnLoad = New System.Windows.Forms.Button()
        Me.chckBoxAutoAssign = New System.Windows.Forms.CheckBox()
        Me.txtAdmNo = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.StatusStrip4 = New System.Windows.Forms.StatusStrip()
        Me.ToolStripStatusLabel1 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.StatusStrip3 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip2 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip5 = New System.Windows.Forms.StatusStrip()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.SplitContainer2 = New System.Windows.Forms.SplitContainer()
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.Panel6 = New System.Windows.Forms.Panel()
        Me.lstDormDetails = New System.Windows.Forms.ListView()
        Me.ColumnHeader1 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader5 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader2 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader3 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader4 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.StatusStrip6 = New System.Windows.Forms.StatusStrip()
        Me.ToolStripStatusLabel2 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.StatusStrip7 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip8 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip9 = New System.Windows.Forms.StatusStrip()
        Me.Panel4 = New System.Windows.Forms.Panel()
        Me.Panel7 = New System.Windows.Forms.Panel()
        Me.btnDelete = New System.Windows.Forms.Button()
        Me.btnUpdate = New System.Windows.Forms.Button()
        Me.btnView = New System.Windows.Forms.Button()
        Me.DTPDateOfAllocation = New System.Windows.Forms.DateTimePicker()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.btnClose = New System.Windows.Forms.Button()
        Me.StatusStrip10 = New System.Windows.Forms.StatusStrip()
        Me.ToolStripStatusLabel3 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.StatusStrip11 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip12 = New System.Windows.Forms.StatusStrip()
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
        Me.pnlRegisterDorms = New System.Windows.Forms.Panel()
        Me.StatusStrip14 = New System.Windows.Forms.StatusStrip()
        Me.ToolStripStatusLabel4 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.Panel1.SuspendLayout()
        Me.Panel5.SuspendLayout()
        Me.StatusStrip4.SuspendLayout()
        Me.Panel2.SuspendLayout()
        CType(Me.SplitContainer2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer2.Panel1.SuspendLayout()
        Me.SplitContainer2.Panel2.SuspendLayout()
        Me.SplitContainer2.SuspendLayout()
        Me.Panel3.SuspendLayout()
        Me.Panel6.SuspendLayout()
        Me.StatusStrip6.SuspendLayout()
        Me.Panel4.SuspendLayout()
        Me.Panel7.SuspendLayout()
        Me.StatusStrip10.SuspendLayout()
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.pnlRegisterDorms.SuspendLayout()
        Me.StatusStrip14.SuspendLayout()
        Me.SuspendLayout()
        '
        'StatusStrip1
        '
        Me.StatusStrip1.BackColor = System.Drawing.Color.LightSeaGreen
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 452)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(743, 22)
        Me.StatusStrip1.SizingGrip = False
        Me.StatusStrip1.TabIndex = 75
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.Panel5)
        Me.Panel1.Controls.Add(Me.StatusStrip4)
        Me.Panel1.Controls.Add(Me.StatusStrip3)
        Me.Panel1.Controls.Add(Me.StatusStrip2)
        Me.Panel1.Controls.Add(Me.StatusStrip5)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel1.Location = New System.Drawing.Point(0, 0)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(743, 111)
        Me.Panel1.TabIndex = 0
        '
        'Panel5
        '
        Me.Panel5.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel5.Controls.Add(Me.cboHouse)
        Me.Panel5.Controls.Add(Me.Label7)
        Me.Panel5.Controls.Add(Me.txtGender)
        Me.Panel5.Controls.Add(Me.Label6)
        Me.Panel5.Controls.Add(Me.txtStudentName)
        Me.Panel5.Controls.Add(Me.Label2)
        Me.Panel5.Controls.Add(Me.btnSave)
        Me.Panel5.Controls.Add(Me.btnLoad)
        Me.Panel5.Controls.Add(Me.chckBoxAutoAssign)
        Me.Panel5.Controls.Add(Me.txtAdmNo)
        Me.Panel5.Controls.Add(Me.Label1)
        Me.Panel5.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel5.Location = New System.Drawing.Point(5, 20)
        Me.Panel5.Name = "Panel5"
        Me.Panel5.Size = New System.Drawing.Size(733, 86)
        Me.Panel5.TabIndex = 8
        '
        'cboHouse
        '
        Me.cboHouse.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboHouse.FormattingEnabled = True
        Me.cboHouse.Location = New System.Drawing.Point(478, 16)
        Me.cboHouse.Name = "cboHouse"
        Me.cboHouse.Size = New System.Drawing.Size(238, 21)
        Me.cboHouse.TabIndex = 24
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(398, 19)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(74, 13)
        Me.Label7.TabIndex = 23
        Me.Label7.Text = "Select House:"
        '
        'txtGender
        '
        Me.txtGender.BackColor = System.Drawing.SystemColors.Control
        Me.txtGender.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtGender.Location = New System.Drawing.Point(569, 51)
        Me.txtGender.Name = "txtGender"
        Me.txtGender.Size = New System.Drawing.Size(147, 20)
        Me.txtGender.TabIndex = 22
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(515, 53)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(48, 13)
        Me.Label6.TabIndex = 18
        Me.Label6.Text = "Gender :"
        '
        'txtStudentName
        '
        Me.txtStudentName.BackColor = System.Drawing.SystemColors.Control
        Me.txtStudentName.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtStudentName.Location = New System.Drawing.Point(99, 49)
        Me.txtStudentName.Name = "txtStudentName"
        Me.txtStudentName.Size = New System.Drawing.Size(282, 20)
        Me.txtStudentName.TabIndex = 17
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(15, 53)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(78, 13)
        Me.Label2.TabIndex = 13
        Me.Label2.Text = "Student Name:"
        '
        'btnSave
        '
        Me.btnSave.Location = New System.Drawing.Point(401, 48)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(70, 23)
        Me.btnSave.TabIndex = 13
        Me.btnSave.Text = "Save"
        Me.btnSave.UseVisualStyleBackColor = True
        '
        'btnLoad
        '
        Me.btnLoad.Location = New System.Drawing.Point(223, 14)
        Me.btnLoad.Name = "btnLoad"
        Me.btnLoad.Size = New System.Drawing.Size(61, 23)
        Me.btnLoad.TabIndex = 12
        Me.btnLoad.Text = "Load"
        Me.btnLoad.UseVisualStyleBackColor = True
        '
        'chckBoxAutoAssign
        '
        Me.chckBoxAutoAssign.AutoSize = True
        Me.chckBoxAutoAssign.Location = New System.Drawing.Point(299, 18)
        Me.chckBoxAutoAssign.Name = "chckBoxAutoAssign"
        Me.chckBoxAutoAssign.Size = New System.Drawing.Size(82, 17)
        Me.chckBoxAutoAssign.TabIndex = 0
        Me.chckBoxAutoAssign.Text = "Auto Assign"
        Me.chckBoxAutoAssign.UseVisualStyleBackColor = True
        '
        'txtAdmNo
        '
        Me.txtAdmNo.Location = New System.Drawing.Point(99, 16)
        Me.txtAdmNo.Name = "txtAdmNo"
        Me.txtAdmNo.Size = New System.Drawing.Size(112, 20)
        Me.txtAdmNo.TabIndex = 11
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(15, 19)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(64, 13)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Student No:"
        '
        'StatusStrip4
        '
        Me.StatusStrip4.AutoSize = False
        Me.StatusStrip4.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip4.Dock = System.Windows.Forms.DockStyle.Top
        Me.StatusStrip4.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripStatusLabel1})
        Me.StatusStrip4.Location = New System.Drawing.Point(5, 0)
        Me.StatusStrip4.Name = "StatusStrip4"
        Me.StatusStrip4.Size = New System.Drawing.Size(733, 20)
        Me.StatusStrip4.SizingGrip = False
        Me.StatusStrip4.TabIndex = 7
        Me.StatusStrip4.Text = "StatusStrip4"
        '
        'ToolStripStatusLabel1
        '
        Me.ToolStripStatusLabel1.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ToolStripStatusLabel1.Name = "ToolStripStatusLabel1"
        Me.ToolStripStatusLabel1.Size = New System.Drawing.Size(93, 15)
        Me.ToolStripStatusLabel1.Text = "Student Details"
        '
        'StatusStrip3
        '
        Me.StatusStrip3.AutoSize = False
        Me.StatusStrip3.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip3.Dock = System.Windows.Forms.DockStyle.Right
        Me.StatusStrip3.Location = New System.Drawing.Point(738, 0)
        Me.StatusStrip3.Name = "StatusStrip3"
        Me.StatusStrip3.Size = New System.Drawing.Size(5, 106)
        Me.StatusStrip3.SizingGrip = False
        Me.StatusStrip3.TabIndex = 6
        Me.StatusStrip3.Text = "StatusStrip3"
        '
        'StatusStrip2
        '
        Me.StatusStrip2.AutoSize = False
        Me.StatusStrip2.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip2.Dock = System.Windows.Forms.DockStyle.Left
        Me.StatusStrip2.Location = New System.Drawing.Point(0, 0)
        Me.StatusStrip2.Name = "StatusStrip2"
        Me.StatusStrip2.Size = New System.Drawing.Size(5, 106)
        Me.StatusStrip2.SizingGrip = False
        Me.StatusStrip2.TabIndex = 5
        Me.StatusStrip2.Text = "StatusStrip2"
        '
        'StatusStrip5
        '
        Me.StatusStrip5.AutoSize = False
        Me.StatusStrip5.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip5.Location = New System.Drawing.Point(0, 106)
        Me.StatusStrip5.Name = "StatusStrip5"
        Me.StatusStrip5.Size = New System.Drawing.Size(743, 5)
        Me.StatusStrip5.SizingGrip = False
        Me.StatusStrip5.TabIndex = 4
        Me.StatusStrip5.Text = "StatusStrip5"
        '
        'Panel2
        '
        Me.Panel2.Controls.Add(Me.SplitContainer2)
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel2.Location = New System.Drawing.Point(0, 0)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(743, 315)
        Me.Panel2.TabIndex = 1
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
        Me.SplitContainer2.Panel1.Controls.Add(Me.Panel3)
        '
        'SplitContainer2.Panel2
        '
        Me.SplitContainer2.Panel2.Controls.Add(Me.Panel4)
        Me.SplitContainer2.Size = New System.Drawing.Size(743, 315)
        Me.SplitContainer2.SplitterDistance = 244
        Me.SplitContainer2.SplitterWidth = 1
        Me.SplitContainer2.TabIndex = 77
        '
        'Panel3
        '
        Me.Panel3.Controls.Add(Me.Panel6)
        Me.Panel3.Controls.Add(Me.StatusStrip6)
        Me.Panel3.Controls.Add(Me.StatusStrip7)
        Me.Panel3.Controls.Add(Me.StatusStrip8)
        Me.Panel3.Controls.Add(Me.StatusStrip9)
        Me.Panel3.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel3.Location = New System.Drawing.Point(0, 0)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(743, 244)
        Me.Panel3.TabIndex = 0
        '
        'Panel6
        '
        Me.Panel6.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel6.Controls.Add(Me.lstDormDetails)
        Me.Panel6.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel6.Location = New System.Drawing.Point(5, 20)
        Me.Panel6.Name = "Panel6"
        Me.Panel6.Size = New System.Drawing.Size(733, 219)
        Me.Panel6.TabIndex = 9
        '
        'lstDormDetails
        '
        Me.lstDormDetails.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.lstDormDetails.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColumnHeader1, Me.ColumnHeader5, Me.ColumnHeader2, Me.ColumnHeader3, Me.ColumnHeader4})
        Me.lstDormDetails.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lstDormDetails.FullRowSelect = True
        Me.lstDormDetails.GridLines = True
        Me.lstDormDetails.Location = New System.Drawing.Point(0, 0)
        Me.lstDormDetails.Name = "lstDormDetails"
        Me.lstDormDetails.Size = New System.Drawing.Size(731, 217)
        Me.lstDormDetails.TabIndex = 0
        Me.lstDormDetails.UseCompatibleStateImageBehavior = False
        Me.lstDormDetails.View = System.Windows.Forms.View.Details
        '
        'ColumnHeader1
        '
        Me.ColumnHeader1.Text = "Name"
        Me.ColumnHeader1.Width = 300
        '
        'ColumnHeader5
        '
        Me.ColumnHeader5.Text = "Gender Type"
        Me.ColumnHeader5.Width = 100
        '
        'ColumnHeader2
        '
        Me.ColumnHeader2.Text = "Capacity"
        Me.ColumnHeader2.Width = 100
        '
        'ColumnHeader3
        '
        Me.ColumnHeader3.Text = "Enrolment"
        Me.ColumnHeader3.Width = 100
        '
        'ColumnHeader4
        '
        Me.ColumnHeader4.Text = "Available"
        Me.ColumnHeader4.Width = 100
        '
        'StatusStrip6
        '
        Me.StatusStrip6.AutoSize = False
        Me.StatusStrip6.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip6.Dock = System.Windows.Forms.DockStyle.Top
        Me.StatusStrip6.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripStatusLabel2})
        Me.StatusStrip6.Location = New System.Drawing.Point(5, 0)
        Me.StatusStrip6.Name = "StatusStrip6"
        Me.StatusStrip6.Size = New System.Drawing.Size(733, 20)
        Me.StatusStrip6.SizingGrip = False
        Me.StatusStrip6.TabIndex = 7
        Me.StatusStrip6.Text = "StatusStrip6"
        '
        'ToolStripStatusLabel2
        '
        Me.ToolStripStatusLabel2.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ToolStripStatusLabel2.Name = "ToolStripStatusLabel2"
        Me.ToolStripStatusLabel2.Size = New System.Drawing.Size(106, 15)
        Me.ToolStripStatusLabel2.Text = "Dormitory Details"
        '
        'StatusStrip7
        '
        Me.StatusStrip7.AutoSize = False
        Me.StatusStrip7.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip7.Dock = System.Windows.Forms.DockStyle.Right
        Me.StatusStrip7.Location = New System.Drawing.Point(738, 0)
        Me.StatusStrip7.Name = "StatusStrip7"
        Me.StatusStrip7.Size = New System.Drawing.Size(5, 239)
        Me.StatusStrip7.SizingGrip = False
        Me.StatusStrip7.TabIndex = 6
        Me.StatusStrip7.Text = "StatusStrip7"
        '
        'StatusStrip8
        '
        Me.StatusStrip8.AutoSize = False
        Me.StatusStrip8.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip8.Dock = System.Windows.Forms.DockStyle.Left
        Me.StatusStrip8.Location = New System.Drawing.Point(0, 0)
        Me.StatusStrip8.Name = "StatusStrip8"
        Me.StatusStrip8.Size = New System.Drawing.Size(5, 239)
        Me.StatusStrip8.SizingGrip = False
        Me.StatusStrip8.TabIndex = 5
        Me.StatusStrip8.Text = "StatusStrip8"
        '
        'StatusStrip9
        '
        Me.StatusStrip9.AutoSize = False
        Me.StatusStrip9.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip9.Location = New System.Drawing.Point(0, 239)
        Me.StatusStrip9.Name = "StatusStrip9"
        Me.StatusStrip9.Size = New System.Drawing.Size(743, 5)
        Me.StatusStrip9.SizingGrip = False
        Me.StatusStrip9.TabIndex = 4
        Me.StatusStrip9.Text = "StatusStrip9"
        '
        'Panel4
        '
        Me.Panel4.Controls.Add(Me.Panel7)
        Me.Panel4.Controls.Add(Me.StatusStrip10)
        Me.Panel4.Controls.Add(Me.StatusStrip11)
        Me.Panel4.Controls.Add(Me.StatusStrip12)
        Me.Panel4.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel4.Location = New System.Drawing.Point(0, 0)
        Me.Panel4.Name = "Panel4"
        Me.Panel4.Size = New System.Drawing.Size(743, 70)
        Me.Panel4.TabIndex = 1
        '
        'Panel7
        '
        Me.Panel7.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel7.Controls.Add(Me.btnDelete)
        Me.Panel7.Controls.Add(Me.btnUpdate)
        Me.Panel7.Controls.Add(Me.btnView)
        Me.Panel7.Controls.Add(Me.DTPDateOfAllocation)
        Me.Panel7.Controls.Add(Me.Label4)
        Me.Panel7.Controls.Add(Me.btnClose)
        Me.Panel7.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel7.Location = New System.Drawing.Point(5, 20)
        Me.Panel7.Name = "Panel7"
        Me.Panel7.Size = New System.Drawing.Size(733, 50)
        Me.Panel7.TabIndex = 12
        '
        'btnDelete
        '
        Me.btnDelete.Enabled = False
        Me.btnDelete.Location = New System.Drawing.Point(521, 12)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(70, 23)
        Me.btnDelete.TabIndex = 28
        Me.btnDelete.Text = "Delete"
        Me.btnDelete.UseVisualStyleBackColor = True
        '
        'btnUpdate
        '
        Me.btnUpdate.Enabled = False
        Me.btnUpdate.Location = New System.Drawing.Point(391, 12)
        Me.btnUpdate.Name = "btnUpdate"
        Me.btnUpdate.Size = New System.Drawing.Size(70, 23)
        Me.btnUpdate.TabIndex = 27
        Me.btnUpdate.Text = "Update"
        Me.btnUpdate.UseVisualStyleBackColor = True
        '
        'btnView
        '
        Me.btnView.Location = New System.Drawing.Point(261, 12)
        Me.btnView.Name = "btnView"
        Me.btnView.Size = New System.Drawing.Size(70, 23)
        Me.btnView.TabIndex = 26
        Me.btnView.Text = "View"
        Me.btnView.UseVisualStyleBackColor = True
        '
        'DTPDateOfAllocation
        '
        Me.DTPDateOfAllocation.CustomFormat = "   dd-MM-yyyy"
        Me.DTPDateOfAllocation.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DTPDateOfAllocation.Location = New System.Drawing.Point(116, 14)
        Me.DTPDateOfAllocation.Name = "DTPDateOfAllocation"
        Me.DTPDateOfAllocation.Size = New System.Drawing.Size(107, 20)
        Me.DTPDateOfAllocation.TabIndex = 25
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(14, 17)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(96, 13)
        Me.Label4.TabIndex = 24
        Me.Label4.Text = "Date Of Allocation:"
        '
        'btnClose
        '
        Me.btnClose.Location = New System.Drawing.Point(646, 12)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(70, 23)
        Me.btnClose.TabIndex = 14
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'StatusStrip10
        '
        Me.StatusStrip10.AutoSize = False
        Me.StatusStrip10.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip10.Dock = System.Windows.Forms.DockStyle.Top
        Me.StatusStrip10.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripStatusLabel3})
        Me.StatusStrip10.Location = New System.Drawing.Point(5, 0)
        Me.StatusStrip10.Name = "StatusStrip10"
        Me.StatusStrip10.Size = New System.Drawing.Size(733, 20)
        Me.StatusStrip10.SizingGrip = False
        Me.StatusStrip10.TabIndex = 11
        Me.StatusStrip10.Text = "StatusStrip10"
        '
        'ToolStripStatusLabel3
        '
        Me.ToolStripStatusLabel3.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ToolStripStatusLabel3.Name = "ToolStripStatusLabel3"
        Me.ToolStripStatusLabel3.Size = New System.Drawing.Size(77, 15)
        Me.ToolStripStatusLabel3.Text = "User Actions"
        '
        'StatusStrip11
        '
        Me.StatusStrip11.AutoSize = False
        Me.StatusStrip11.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip11.Dock = System.Windows.Forms.DockStyle.Right
        Me.StatusStrip11.Location = New System.Drawing.Point(738, 0)
        Me.StatusStrip11.Name = "StatusStrip11"
        Me.StatusStrip11.Size = New System.Drawing.Size(5, 70)
        Me.StatusStrip11.SizingGrip = False
        Me.StatusStrip11.TabIndex = 10
        Me.StatusStrip11.Text = "StatusStrip11"
        '
        'StatusStrip12
        '
        Me.StatusStrip12.AutoSize = False
        Me.StatusStrip12.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip12.Dock = System.Windows.Forms.DockStyle.Left
        Me.StatusStrip12.Location = New System.Drawing.Point(0, 0)
        Me.StatusStrip12.Name = "StatusStrip12"
        Me.StatusStrip12.Size = New System.Drawing.Size(5, 70)
        Me.StatusStrip12.SizingGrip = False
        Me.StatusStrip12.TabIndex = 9
        Me.StatusStrip12.Text = "StatusStrip12"
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
        Me.SplitContainer1.Panel2.Controls.Add(Me.Panel2)
        Me.SplitContainer1.Size = New System.Drawing.Size(743, 427)
        Me.SplitContainer1.SplitterDistance = 111
        Me.SplitContainer1.SplitterWidth = 1
        Me.SplitContainer1.TabIndex = 76
        '
        'pnlRegisterDorms
        '
        Me.pnlRegisterDorms.BackColor = System.Drawing.SystemColors.Control
        Me.pnlRegisterDorms.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pnlRegisterDorms.Controls.Add(Me.SplitContainer1)
        Me.pnlRegisterDorms.Controls.Add(Me.StatusStrip14)
        Me.pnlRegisterDorms.Controls.Add(Me.StatusStrip1)
        Me.pnlRegisterDorms.Location = New System.Drawing.Point(12, 12)
        Me.pnlRegisterDorms.Name = "pnlRegisterDorms"
        Me.pnlRegisterDorms.Size = New System.Drawing.Size(745, 476)
        Me.pnlRegisterDorms.TabIndex = 1
        '
        'StatusStrip14
        '
        Me.StatusStrip14.AutoSize = False
        Me.StatusStrip14.BackColor = System.Drawing.Color.LightSeaGreen
        Me.StatusStrip14.Dock = System.Windows.Forms.DockStyle.Top
        Me.StatusStrip14.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripStatusLabel4})
        Me.StatusStrip14.Location = New System.Drawing.Point(0, 0)
        Me.StatusStrip14.Name = "StatusStrip14"
        Me.StatusStrip14.Size = New System.Drawing.Size(743, 25)
        Me.StatusStrip14.SizingGrip = False
        Me.StatusStrip14.TabIndex = 77
        Me.StatusStrip14.Text = "StatusStrip14"
        '
        'ToolStripStatusLabel4
        '
        Me.ToolStripStatusLabel4.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ToolStripStatusLabel4.Name = "ToolStripStatusLabel4"
        Me.ToolStripStatusLabel4.Size = New System.Drawing.Size(332, 20)
        Me.ToolStripStatusLabel4.Text = "ACCOMODATION ASSIGN STUDENT TO DORMITORY"
        '
        'frmAccDormStudents
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.LightSkyBlue
        Me.ClientSize = New System.Drawing.Size(770, 500)
        Me.Controls.Add(Me.pnlRegisterDorms)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.MaximizeBox = False
        Me.Name = "frmAccDormStudents"
        Me.ShowIcon = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Assign Students"
        Me.Panel1.ResumeLayout(False)
        Me.Panel5.ResumeLayout(False)
        Me.Panel5.PerformLayout()
        Me.StatusStrip4.ResumeLayout(False)
        Me.StatusStrip4.PerformLayout()
        Me.Panel2.ResumeLayout(False)
        Me.SplitContainer2.Panel1.ResumeLayout(False)
        Me.SplitContainer2.Panel2.ResumeLayout(False)
        CType(Me.SplitContainer2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer2.ResumeLayout(False)
        Me.Panel3.ResumeLayout(False)
        Me.Panel6.ResumeLayout(False)
        Me.StatusStrip6.ResumeLayout(False)
        Me.StatusStrip6.PerformLayout()
        Me.Panel4.ResumeLayout(False)
        Me.Panel7.ResumeLayout(False)
        Me.Panel7.PerformLayout()
        Me.StatusStrip10.ResumeLayout(False)
        Me.StatusStrip10.PerformLayout()
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.ResumeLayout(False)
        Me.pnlRegisterDorms.ResumeLayout(False)
        Me.pnlRegisterDorms.PerformLayout()
        Me.StatusStrip14.ResumeLayout(False)
        Me.StatusStrip14.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents StatusStrip4 As System.Windows.Forms.StatusStrip
    Friend WithEvents ToolStripStatusLabel1 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents StatusStrip3 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip2 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip5 As System.Windows.Forms.StatusStrip
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents pnlRegisterDorms As System.Windows.Forms.Panel
    Friend WithEvents SplitContainer2 As System.Windows.Forms.SplitContainer
    Friend WithEvents Panel3 As System.Windows.Forms.Panel
    Friend WithEvents StatusStrip6 As System.Windows.Forms.StatusStrip
    Friend WithEvents ToolStripStatusLabel2 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents StatusStrip7 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip8 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip9 As System.Windows.Forms.StatusStrip
    Friend WithEvents Panel4 As System.Windows.Forms.Panel
    Friend WithEvents StatusStrip10 As System.Windows.Forms.StatusStrip
    Friend WithEvents ToolStripStatusLabel3 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents StatusStrip11 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip12 As System.Windows.Forms.StatusStrip
    Friend WithEvents Panel5 As System.Windows.Forms.Panel
    Friend WithEvents Panel6 As System.Windows.Forms.Panel
    Friend WithEvents Panel7 As System.Windows.Forms.Panel
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents chckBoxAutoAssign As System.Windows.Forms.CheckBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents btnLoad As System.Windows.Forms.Button
    Friend WithEvents txtAdmNo As System.Windows.Forms.TextBox
    Friend WithEvents txtGender As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents txtStudentName As System.Windows.Forms.TextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents DTPDateOfAllocation As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents lstDormDetails As System.Windows.Forms.ListView
    Friend WithEvents ColumnHeader1 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader2 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader3 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader4 As System.Windows.Forms.ColumnHeader
    Friend WithEvents cboHouse As System.Windows.Forms.ComboBox
    Friend WithEvents btnView As System.Windows.Forms.Button
    Friend WithEvents StatusStrip14 As System.Windows.Forms.StatusStrip
    Friend WithEvents ToolStripStatusLabel4 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents ColumnHeader5 As System.Windows.Forms.ColumnHeader
    Friend WithEvents btnUpdate As System.Windows.Forms.Button
    Friend WithEvents btnDelete As System.Windows.Forms.Button
End Class
