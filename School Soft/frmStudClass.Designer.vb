<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmStudClass
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
        Me.pnlStudClass = New System.Windows.Forms.Panel()
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
        Me.SplitContainer2 = New System.Windows.Forms.SplitContainer()
        Me.txtAdmNoVw = New System.Windows.Forms.TextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.txtFullName = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.btnLoad = New System.Windows.Forms.Button()
        Me.txtAdmNo = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.StatusStrip11 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip12 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip13 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip14 = New System.Windows.Forms.StatusStrip()
        Me.ToolStripStatusLabel2 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.cboClassYear = New System.Windows.Forms.ComboBox()
        Me.cboClassStream = New System.Windows.Forms.ComboBox()
        Me.cboClassName = New System.Windows.Forms.ComboBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.StatusStrip6 = New System.Windows.Forms.StatusStrip()
        Me.ToolStripStatusLabel3 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.StatusStrip3 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip4 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip5 = New System.Windows.Forms.StatusStrip()
        Me.SplitContainer3 = New System.Windows.Forms.SplitContainer()
        Me.lstStudentClass = New System.Windows.Forms.ListView()
        Me.ColumnHeader1 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader2 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader3 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader4 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader5 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.btnView = New System.Windows.Forms.Button()
        Me.btnDelete = New System.Windows.Forms.Button()
        Me.btnUpdate = New System.Windows.Forms.Button()
        Me.btnSave = New System.Windows.Forms.Button()
        Me.btnClose = New System.Windows.Forms.Button()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.StatusStrip7 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip8 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip9 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip10 = New System.Windows.Forms.StatusStrip()
        Me.ToolStripStatusLabel4 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.StatusStrip2 = New System.Windows.Forms.StatusStrip()
        Me.ToolStripStatusLabel1 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip()
        Me.pnlStudClass.SuspendLayout()
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        CType(Me.SplitContainer2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer2.Panel1.SuspendLayout()
        Me.SplitContainer2.Panel2.SuspendLayout()
        Me.SplitContainer2.SuspendLayout()
        Me.StatusStrip14.SuspendLayout()
        Me.StatusStrip6.SuspendLayout()
        CType(Me.SplitContainer3, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer3.Panel1.SuspendLayout()
        Me.SplitContainer3.Panel2.SuspendLayout()
        Me.SplitContainer3.SuspendLayout()
        Me.StatusStrip10.SuspendLayout()
        Me.StatusStrip2.SuspendLayout()
        Me.SuspendLayout()
        '
        'pnlStudClass
        '
        Me.pnlStudClass.BackColor = System.Drawing.SystemColors.Control
        Me.pnlStudClass.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pnlStudClass.Controls.Add(Me.SplitContainer1)
        Me.pnlStudClass.Controls.Add(Me.StatusStrip2)
        Me.pnlStudClass.Controls.Add(Me.StatusStrip1)
        Me.pnlStudClass.Location = New System.Drawing.Point(12, 12)
        Me.pnlStudClass.Name = "pnlStudClass"
        Me.pnlStudClass.Size = New System.Drawing.Size(775, 548)
        Me.pnlStudClass.TabIndex = 0
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
        Me.SplitContainer1.Panel1.Controls.Add(Me.SplitContainer2)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.SplitContainer3)
        Me.SplitContainer1.Size = New System.Drawing.Size(773, 496)
        Me.SplitContainer1.SplitterDistance = 137
        Me.SplitContainer1.SplitterWidth = 1
        Me.SplitContainer1.TabIndex = 25
        '
        'SplitContainer2
        '
        Me.SplitContainer2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.SplitContainer2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer2.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer2.Name = "SplitContainer2"
        '
        'SplitContainer2.Panel1
        '
        Me.SplitContainer2.Panel1.Controls.Add(Me.txtAdmNoVw)
        Me.SplitContainer2.Panel1.Controls.Add(Me.Label6)
        Me.SplitContainer2.Panel1.Controls.Add(Me.txtFullName)
        Me.SplitContainer2.Panel1.Controls.Add(Me.Label5)
        Me.SplitContainer2.Panel1.Controls.Add(Me.btnLoad)
        Me.SplitContainer2.Panel1.Controls.Add(Me.txtAdmNo)
        Me.SplitContainer2.Panel1.Controls.Add(Me.Label4)
        Me.SplitContainer2.Panel1.Controls.Add(Me.Panel2)
        Me.SplitContainer2.Panel1.Controls.Add(Me.StatusStrip11)
        Me.SplitContainer2.Panel1.Controls.Add(Me.StatusStrip12)
        Me.SplitContainer2.Panel1.Controls.Add(Me.StatusStrip13)
        Me.SplitContainer2.Panel1.Controls.Add(Me.StatusStrip14)
        '
        'SplitContainer2.Panel2
        '
        Me.SplitContainer2.Panel2.Controls.Add(Me.cboClassYear)
        Me.SplitContainer2.Panel2.Controls.Add(Me.cboClassStream)
        Me.SplitContainer2.Panel2.Controls.Add(Me.cboClassName)
        Me.SplitContainer2.Panel2.Controls.Add(Me.Label3)
        Me.SplitContainer2.Panel2.Controls.Add(Me.Label2)
        Me.SplitContainer2.Panel2.Controls.Add(Me.Label1)
        Me.SplitContainer2.Panel2.Controls.Add(Me.Panel3)
        Me.SplitContainer2.Panel2.Controls.Add(Me.StatusStrip6)
        Me.SplitContainer2.Panel2.Controls.Add(Me.StatusStrip3)
        Me.SplitContainer2.Panel2.Controls.Add(Me.StatusStrip4)
        Me.SplitContainer2.Panel2.Controls.Add(Me.StatusStrip5)
        Me.SplitContainer2.Size = New System.Drawing.Size(773, 137)
        Me.SplitContainer2.SplitterDistance = 452
        Me.SplitContainer2.SplitterWidth = 1
        Me.SplitContainer2.TabIndex = 26
        '
        'txtAdmNoVw
        '
        Me.txtAdmNoVw.Location = New System.Drawing.Point(81, 99)
        Me.txtAdmNoVw.Name = "txtAdmNoVw"
        Me.txtAdmNoVw.ReadOnly = True
        Me.txtAdmNoVw.Size = New System.Drawing.Size(153, 20)
        Me.txtAdmNoVw.TabIndex = 49
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(15, 102)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(48, 13)
        Me.Label6.TabIndex = 48
        Me.Label6.Text = "Adm No:"
        '
        'txtFullName
        '
        Me.txtFullName.Location = New System.Drawing.Point(81, 66)
        Me.txtFullName.Name = "txtFullName"
        Me.txtFullName.ReadOnly = True
        Me.txtFullName.Size = New System.Drawing.Size(357, 20)
        Me.txtFullName.TabIndex = 47
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(15, 69)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(54, 13)
        Me.Label5.TabIndex = 46
        Me.Label5.Text = "FullName:"
        '
        'btnLoad
        '
        Me.btnLoad.Location = New System.Drawing.Point(201, 30)
        Me.btnLoad.Name = "btnLoad"
        Me.btnLoad.Size = New System.Drawing.Size(75, 23)
        Me.btnLoad.TabIndex = 45
        Me.btnLoad.Text = "Load"
        Me.btnLoad.UseVisualStyleBackColor = True
        '
        'txtAdmNo
        '
        Me.txtAdmNo.Location = New System.Drawing.Point(81, 32)
        Me.txtAdmNo.Name = "txtAdmNo"
        Me.txtAdmNo.Size = New System.Drawing.Size(112, 20)
        Me.txtAdmNo.TabIndex = 44
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(15, 35)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(48, 13)
        Me.Label4.TabIndex = 43
        Me.Label4.Text = "Adm No:"
        '
        'Panel2
        '
        Me.Panel2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel2.Location = New System.Drawing.Point(5, 22)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(440, 108)
        Me.Panel2.TabIndex = 64
        '
        'StatusStrip11
        '
        Me.StatusStrip11.AutoSize = False
        Me.StatusStrip11.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip11.Dock = System.Windows.Forms.DockStyle.Right
        Me.StatusStrip11.Location = New System.Drawing.Point(445, 22)
        Me.StatusStrip11.Name = "StatusStrip11"
        Me.StatusStrip11.Size = New System.Drawing.Size(5, 108)
        Me.StatusStrip11.SizingGrip = False
        Me.StatusStrip11.TabIndex = 41
        Me.StatusStrip11.Text = "StatusStrip11"
        '
        'StatusStrip12
        '
        Me.StatusStrip12.AutoSize = False
        Me.StatusStrip12.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip12.Dock = System.Windows.Forms.DockStyle.Left
        Me.StatusStrip12.Location = New System.Drawing.Point(0, 22)
        Me.StatusStrip12.Name = "StatusStrip12"
        Me.StatusStrip12.Size = New System.Drawing.Size(5, 108)
        Me.StatusStrip12.SizingGrip = False
        Me.StatusStrip12.TabIndex = 40
        Me.StatusStrip12.Text = "StatusStrip12"
        '
        'StatusStrip13
        '
        Me.StatusStrip13.AutoSize = False
        Me.StatusStrip13.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip13.Location = New System.Drawing.Point(0, 130)
        Me.StatusStrip13.Name = "StatusStrip13"
        Me.StatusStrip13.Size = New System.Drawing.Size(450, 5)
        Me.StatusStrip13.SizingGrip = False
        Me.StatusStrip13.TabIndex = 39
        Me.StatusStrip13.Text = "StatusStrip13"
        '
        'StatusStrip14
        '
        Me.StatusStrip14.AutoSize = False
        Me.StatusStrip14.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip14.Dock = System.Windows.Forms.DockStyle.Top
        Me.StatusStrip14.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripStatusLabel2})
        Me.StatusStrip14.Location = New System.Drawing.Point(0, 0)
        Me.StatusStrip14.Name = "StatusStrip14"
        Me.StatusStrip14.Size = New System.Drawing.Size(450, 22)
        Me.StatusStrip14.SizingGrip = False
        Me.StatusStrip14.TabIndex = 38
        Me.StatusStrip14.Text = "StatusStrip14"
        '
        'ToolStripStatusLabel2
        '
        Me.ToolStripStatusLabel2.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
        Me.ToolStripStatusLabel2.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ToolStripStatusLabel2.Name = "ToolStripStatusLabel2"
        Me.ToolStripStatusLabel2.Size = New System.Drawing.Size(103, 17)
        Me.ToolStripStatusLabel2.Text = "Student Details"
        '
        'cboClassYear
        '
        Me.cboClassYear.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboClassYear.FormattingEnabled = True
        Me.cboClassYear.Location = New System.Drawing.Point(107, 99)
        Me.cboClassYear.Name = "cboClassYear"
        Me.cboClassYear.Size = New System.Drawing.Size(183, 21)
        Me.cboClassYear.TabIndex = 47
        '
        'cboClassStream
        '
        Me.cboClassStream.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboClassStream.FormattingEnabled = True
        Me.cboClassStream.Location = New System.Drawing.Point(107, 66)
        Me.cboClassStream.Name = "cboClassStream"
        Me.cboClassStream.Size = New System.Drawing.Size(183, 21)
        Me.cboClassStream.TabIndex = 46
        '
        'cboClassName
        '
        Me.cboClassName.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboClassName.FormattingEnabled = True
        Me.cboClassName.Location = New System.Drawing.Point(107, 32)
        Me.cboClassName.Name = "cboClassName"
        Me.cboClassName.Size = New System.Drawing.Size(183, 21)
        Me.cboClassName.TabIndex = 45
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(24, 102)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(60, 13)
        Me.Label3.TabIndex = 44
        Me.Label3.Text = "Class Year:"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(24, 69)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(71, 13)
        Me.Label2.TabIndex = 43
        Me.Label2.Text = "Class Stream:"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(24, 35)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(66, 13)
        Me.Label1.TabIndex = 42
        Me.Label1.Text = "Class Name:"
        '
        'Panel3
        '
        Me.Panel3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel3.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel3.Location = New System.Drawing.Point(5, 22)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(308, 108)
        Me.Panel3.TabIndex = 64
        '
        'StatusStrip6
        '
        Me.StatusStrip6.AutoSize = False
        Me.StatusStrip6.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip6.Dock = System.Windows.Forms.DockStyle.Top
        Me.StatusStrip6.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripStatusLabel3})
        Me.StatusStrip6.Location = New System.Drawing.Point(5, 0)
        Me.StatusStrip6.Name = "StatusStrip6"
        Me.StatusStrip6.Size = New System.Drawing.Size(308, 22)
        Me.StatusStrip6.SizingGrip = False
        Me.StatusStrip6.TabIndex = 42
        Me.StatusStrip6.Text = "StatusStrip6"
        '
        'ToolStripStatusLabel3
        '
        Me.ToolStripStatusLabel3.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ToolStripStatusLabel3.Name = "ToolStripStatusLabel3"
        Me.ToolStripStatusLabel3.Size = New System.Drawing.Size(86, 17)
        Me.ToolStripStatusLabel3.Text = "Class Details"
        '
        'StatusStrip3
        '
        Me.StatusStrip3.AutoSize = False
        Me.StatusStrip3.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip3.Dock = System.Windows.Forms.DockStyle.Right
        Me.StatusStrip3.Location = New System.Drawing.Point(313, 0)
        Me.StatusStrip3.Name = "StatusStrip3"
        Me.StatusStrip3.Size = New System.Drawing.Size(5, 130)
        Me.StatusStrip3.SizingGrip = False
        Me.StatusStrip3.TabIndex = 41
        Me.StatusStrip3.Text = "StatusStrip3"
        '
        'StatusStrip4
        '
        Me.StatusStrip4.AutoSize = False
        Me.StatusStrip4.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip4.Dock = System.Windows.Forms.DockStyle.Left
        Me.StatusStrip4.Location = New System.Drawing.Point(0, 0)
        Me.StatusStrip4.Name = "StatusStrip4"
        Me.StatusStrip4.Size = New System.Drawing.Size(5, 130)
        Me.StatusStrip4.SizingGrip = False
        Me.StatusStrip4.TabIndex = 40
        Me.StatusStrip4.Text = "StatusStrip4"
        '
        'StatusStrip5
        '
        Me.StatusStrip5.AutoSize = False
        Me.StatusStrip5.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip5.Location = New System.Drawing.Point(0, 130)
        Me.StatusStrip5.Name = "StatusStrip5"
        Me.StatusStrip5.Size = New System.Drawing.Size(318, 5)
        Me.StatusStrip5.SizingGrip = False
        Me.StatusStrip5.TabIndex = 39
        Me.StatusStrip5.Text = "StatusStrip5"
        '
        'SplitContainer3
        '
        Me.SplitContainer3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.SplitContainer3.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer3.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer3.Name = "SplitContainer3"
        Me.SplitContainer3.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'SplitContainer3.Panel1
        '
        Me.SplitContainer3.Panel1.Controls.Add(Me.lstStudentClass)
        '
        'SplitContainer3.Panel2
        '
        Me.SplitContainer3.Panel2.Controls.Add(Me.btnView)
        Me.SplitContainer3.Panel2.Controls.Add(Me.btnDelete)
        Me.SplitContainer3.Panel2.Controls.Add(Me.btnUpdate)
        Me.SplitContainer3.Panel2.Controls.Add(Me.btnSave)
        Me.SplitContainer3.Panel2.Controls.Add(Me.btnClose)
        Me.SplitContainer3.Panel2.Controls.Add(Me.Panel1)
        Me.SplitContainer3.Panel2.Controls.Add(Me.StatusStrip7)
        Me.SplitContainer3.Panel2.Controls.Add(Me.StatusStrip8)
        Me.SplitContainer3.Panel2.Controls.Add(Me.StatusStrip9)
        Me.SplitContainer3.Panel2.Controls.Add(Me.StatusStrip10)
        Me.SplitContainer3.Size = New System.Drawing.Size(773, 358)
        Me.SplitContainer3.SplitterDistance = 287
        Me.SplitContainer3.SplitterWidth = 1
        Me.SplitContainer3.TabIndex = 26
        '
        'lstStudentClass
        '
        Me.lstStudentClass.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColumnHeader1, Me.ColumnHeader2, Me.ColumnHeader3, Me.ColumnHeader4, Me.ColumnHeader5})
        Me.lstStudentClass.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lstStudentClass.FullRowSelect = True
        Me.lstStudentClass.GridLines = True
        Me.lstStudentClass.Location = New System.Drawing.Point(0, 0)
        Me.lstStudentClass.Name = "lstStudentClass"
        Me.lstStudentClass.Size = New System.Drawing.Size(771, 285)
        Me.lstStudentClass.TabIndex = 27
        Me.lstStudentClass.UseCompatibleStateImageBehavior = False
        Me.lstStudentClass.View = System.Windows.Forms.View.Details
        '
        'ColumnHeader1
        '
        Me.ColumnHeader1.Text = "Stud Number"
        Me.ColumnHeader1.Width = 100
        '
        'ColumnHeader2
        '
        Me.ColumnHeader2.Text = "Full Name"
        Me.ColumnHeader2.Width = 350
        '
        'ColumnHeader3
        '
        Me.ColumnHeader3.Text = "Class Name"
        Me.ColumnHeader3.Width = 100
        '
        'ColumnHeader4
        '
        Me.ColumnHeader4.Text = "Class Stream"
        Me.ColumnHeader4.Width = 100
        '
        'ColumnHeader5
        '
        Me.ColumnHeader5.Text = "Class Year"
        Me.ColumnHeader5.Width = 100
        '
        'btnView
        '
        Me.btnView.Location = New System.Drawing.Point(190, 31)
        Me.btnView.Name = "btnView"
        Me.btnView.Size = New System.Drawing.Size(75, 23)
        Me.btnView.TabIndex = 62
        Me.btnView.Text = "View"
        Me.btnView.UseVisualStyleBackColor = True
        '
        'btnDelete
        '
        Me.btnDelete.Enabled = False
        Me.btnDelete.Location = New System.Drawing.Point(542, 31)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(75, 23)
        Me.btnDelete.TabIndex = 61
        Me.btnDelete.Text = "Delete"
        Me.btnDelete.UseVisualStyleBackColor = True
        '
        'btnUpdate
        '
        Me.btnUpdate.Enabled = False
        Me.btnUpdate.Location = New System.Drawing.Point(423, 31)
        Me.btnUpdate.Name = "btnUpdate"
        Me.btnUpdate.Size = New System.Drawing.Size(75, 23)
        Me.btnUpdate.TabIndex = 60
        Me.btnUpdate.Text = "Update"
        Me.btnUpdate.UseVisualStyleBackColor = True
        '
        'btnSave
        '
        Me.btnSave.Location = New System.Drawing.Point(306, 31)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(75, 23)
        Me.btnSave.TabIndex = 59
        Me.btnSave.Text = "Save"
        Me.btnSave.UseVisualStyleBackColor = True
        '
        'btnClose
        '
        Me.btnClose.Location = New System.Drawing.Point(658, 31)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(75, 23)
        Me.btnClose.TabIndex = 58
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'Panel1
        '
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel1.Location = New System.Drawing.Point(5, 22)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(761, 41)
        Me.Panel1.TabIndex = 63
        '
        'StatusStrip7
        '
        Me.StatusStrip7.AutoSize = False
        Me.StatusStrip7.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip7.Dock = System.Windows.Forms.DockStyle.Right
        Me.StatusStrip7.Location = New System.Drawing.Point(766, 22)
        Me.StatusStrip7.Name = "StatusStrip7"
        Me.StatusStrip7.Size = New System.Drawing.Size(5, 41)
        Me.StatusStrip7.SizingGrip = False
        Me.StatusStrip7.TabIndex = 41
        Me.StatusStrip7.Text = "StatusStrip7"
        '
        'StatusStrip8
        '
        Me.StatusStrip8.AutoSize = False
        Me.StatusStrip8.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip8.Dock = System.Windows.Forms.DockStyle.Left
        Me.StatusStrip8.Location = New System.Drawing.Point(0, 22)
        Me.StatusStrip8.Name = "StatusStrip8"
        Me.StatusStrip8.Size = New System.Drawing.Size(5, 41)
        Me.StatusStrip8.SizingGrip = False
        Me.StatusStrip8.TabIndex = 40
        Me.StatusStrip8.Text = "StatusStrip8"
        '
        'StatusStrip9
        '
        Me.StatusStrip9.AutoSize = False
        Me.StatusStrip9.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip9.Location = New System.Drawing.Point(0, 63)
        Me.StatusStrip9.Name = "StatusStrip9"
        Me.StatusStrip9.Size = New System.Drawing.Size(771, 5)
        Me.StatusStrip9.SizingGrip = False
        Me.StatusStrip9.TabIndex = 39
        Me.StatusStrip9.Text = "StatusStrip9"
        '
        'StatusStrip10
        '
        Me.StatusStrip10.AutoSize = False
        Me.StatusStrip10.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip10.Dock = System.Windows.Forms.DockStyle.Top
        Me.StatusStrip10.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripStatusLabel4})
        Me.StatusStrip10.Location = New System.Drawing.Point(0, 0)
        Me.StatusStrip10.Name = "StatusStrip10"
        Me.StatusStrip10.Size = New System.Drawing.Size(771, 22)
        Me.StatusStrip10.SizingGrip = False
        Me.StatusStrip10.TabIndex = 38
        Me.StatusStrip10.Text = "StatusStrip10"
        '
        'ToolStripStatusLabel4
        '
        Me.ToolStripStatusLabel4.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ToolStripStatusLabel4.Name = "ToolStripStatusLabel4"
        Me.ToolStripStatusLabel4.Size = New System.Drawing.Size(85, 17)
        Me.ToolStripStatusLabel4.Text = "User Actions"
        '
        'StatusStrip2
        '
        Me.StatusStrip2.AutoSize = False
        Me.StatusStrip2.BackColor = System.Drawing.Color.LightSeaGreen
        Me.StatusStrip2.Dock = System.Windows.Forms.DockStyle.Top
        Me.StatusStrip2.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripStatusLabel1})
        Me.StatusStrip2.Location = New System.Drawing.Point(0, 0)
        Me.StatusStrip2.Name = "StatusStrip2"
        Me.StatusStrip2.Size = New System.Drawing.Size(773, 25)
        Me.StatusStrip2.SizingGrip = False
        Me.StatusStrip2.TabIndex = 24
        Me.StatusStrip2.Text = "StatusStrip2"
        '
        'ToolStripStatusLabel1
        '
        Me.ToolStripStatusLabel1.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ToolStripStatusLabel1.Name = "ToolStripStatusLabel1"
        Me.ToolStripStatusLabel1.Size = New System.Drawing.Size(159, 20)
        Me.ToolStripStatusLabel1.Text = "ASSIGN STUDENT CLASS"
        '
        'StatusStrip1
        '
        Me.StatusStrip1.AutoSize = False
        Me.StatusStrip1.BackColor = System.Drawing.Color.LightSeaGreen
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 521)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(773, 25)
        Me.StatusStrip1.SizingGrip = False
        Me.StatusStrip1.TabIndex = 23
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'frmStudClass
        '
        Me.AcceptButton = Me.btnSave
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.LightSkyBlue
        Me.ClientSize = New System.Drawing.Size(799, 572)
        Me.Controls.Add(Me.pnlStudClass)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Name = "frmStudClass"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Student-Class"
        Me.pnlStudClass.ResumeLayout(False)
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.ResumeLayout(False)
        Me.SplitContainer2.Panel1.ResumeLayout(False)
        Me.SplitContainer2.Panel1.PerformLayout()
        Me.SplitContainer2.Panel2.ResumeLayout(False)
        Me.SplitContainer2.Panel2.PerformLayout()
        CType(Me.SplitContainer2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer2.ResumeLayout(False)
        Me.StatusStrip14.ResumeLayout(False)
        Me.StatusStrip14.PerformLayout()
        Me.StatusStrip6.ResumeLayout(False)
        Me.StatusStrip6.PerformLayout()
        Me.SplitContainer3.Panel1.ResumeLayout(False)
        Me.SplitContainer3.Panel2.ResumeLayout(False)
        CType(Me.SplitContainer3, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer3.ResumeLayout(False)
        Me.StatusStrip10.ResumeLayout(False)
        Me.StatusStrip10.PerformLayout()
        Me.StatusStrip2.ResumeLayout(False)
        Me.StatusStrip2.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents pnlStudClass As System.Windows.Forms.Panel
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents SplitContainer2 As System.Windows.Forms.SplitContainer
    Friend WithEvents SplitContainer3 As System.Windows.Forms.SplitContainer
    Friend WithEvents StatusStrip2 As System.Windows.Forms.StatusStrip
    Friend WithEvents ToolStripStatusLabel1 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip11 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip12 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip13 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip14 As System.Windows.Forms.StatusStrip
    Friend WithEvents ToolStripStatusLabel2 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents StatusStrip6 As System.Windows.Forms.StatusStrip
    Friend WithEvents ToolStripStatusLabel3 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents StatusStrip3 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip4 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip5 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip7 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip8 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip9 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip10 As System.Windows.Forms.StatusStrip
    Friend WithEvents ToolStripStatusLabel4 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents lstStudentClass As System.Windows.Forms.ListView
    Friend WithEvents txtAdmNoVw As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents txtFullName As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents btnLoad As System.Windows.Forms.Button
    Friend WithEvents txtAdmNo As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
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
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents Panel3 As System.Windows.Forms.Panel
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
End Class
