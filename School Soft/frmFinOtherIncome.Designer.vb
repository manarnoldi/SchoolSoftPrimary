<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmFinOtherIncome
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
        Me.pnlOtherInc = New System.Windows.Forms.Panel()
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.cboAccountNo = New System.Windows.Forms.ComboBox()
        Me.btnRemove = New System.Windows.Forms.Button()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.cboYear = New System.Windows.Forms.ComboBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.cbomodeName = New System.Windows.Forms.ComboBox()
        Me.btnAdd = New System.Windows.Forms.Button()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.cboTermName = New System.Windows.Forms.ComboBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.txtAmount = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.cboRecFrom = New System.Windows.Forms.ComboBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.cboSourceName = New System.Windows.Forms.ComboBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.dtpIncDate = New System.Windows.Forms.DateTimePicker()
        Me.StatusStrip3 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip4 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip6 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip7 = New System.Windows.Forms.StatusStrip()
        Me.SplitContainer2 = New System.Windows.Forms.SplitContainer()
        Me.lstOtherInc = New System.Windows.Forms.ListView()
        Me.ColumnHeader6 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader7 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader8 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader9 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader10 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader11 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.btnClear = New System.Windows.Forms.Button()
        Me.btnClose = New System.Windows.Forms.Button()
        Me.btnUpdate = New System.Windows.Forms.Button()
        Me.StatusStrip5 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip8 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip9 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip10 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip2 = New System.Windows.Forms.StatusStrip()
        Me.ToolStripStatusLabel1 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip()
        Me.pnlOtherInc.SuspendLayout()
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.Panel2.SuspendLayout()
        CType(Me.SplitContainer2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer2.Panel1.SuspendLayout()
        Me.SplitContainer2.Panel2.SuspendLayout()
        Me.SplitContainer2.SuspendLayout()
        Me.Panel3.SuspendLayout()
        Me.StatusStrip2.SuspendLayout()
        Me.SuspendLayout()
        '
        'pnlOtherInc
        '
        Me.pnlOtherInc.BackColor = System.Drawing.SystemColors.Control
        Me.pnlOtherInc.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pnlOtherInc.Controls.Add(Me.SplitContainer1)
        Me.pnlOtherInc.Controls.Add(Me.StatusStrip2)
        Me.pnlOtherInc.Controls.Add(Me.StatusStrip1)
        Me.pnlOtherInc.Location = New System.Drawing.Point(12, 12)
        Me.pnlOtherInc.Name = "pnlOtherInc"
        Me.pnlOtherInc.Size = New System.Drawing.Size(818, 497)
        Me.pnlOtherInc.TabIndex = 0
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
        Me.SplitContainer1.Panel1.Controls.Add(Me.Panel2)
        Me.SplitContainer1.Panel1.Controls.Add(Me.StatusStrip3)
        Me.SplitContainer1.Panel1.Controls.Add(Me.StatusStrip4)
        Me.SplitContainer1.Panel1.Controls.Add(Me.StatusStrip6)
        Me.SplitContainer1.Panel1.Controls.Add(Me.StatusStrip7)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.SplitContainer2)
        Me.SplitContainer1.Size = New System.Drawing.Size(816, 445)
        Me.SplitContainer1.SplitterDistance = 118
        Me.SplitContainer1.SplitterWidth = 1
        Me.SplitContainer1.TabIndex = 29
        '
        'Panel2
        '
        Me.Panel2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel2.Controls.Add(Me.cboAccountNo)
        Me.Panel2.Controls.Add(Me.btnRemove)
        Me.Panel2.Controls.Add(Me.Label8)
        Me.Panel2.Controls.Add(Me.cboYear)
        Me.Panel2.Controls.Add(Me.Label7)
        Me.Panel2.Controls.Add(Me.cbomodeName)
        Me.Panel2.Controls.Add(Me.btnAdd)
        Me.Panel2.Controls.Add(Me.Label6)
        Me.Panel2.Controls.Add(Me.cboTermName)
        Me.Panel2.Controls.Add(Me.Label5)
        Me.Panel2.Controls.Add(Me.txtAmount)
        Me.Panel2.Controls.Add(Me.Label4)
        Me.Panel2.Controls.Add(Me.cboRecFrom)
        Me.Panel2.Controls.Add(Me.Label3)
        Me.Panel2.Controls.Add(Me.cboSourceName)
        Me.Panel2.Controls.Add(Me.Label2)
        Me.Panel2.Controls.Add(Me.Label14)
        Me.Panel2.Controls.Add(Me.dtpIncDate)
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel2.Location = New System.Drawing.Point(5, 5)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(804, 106)
        Me.Panel2.TabIndex = 74
        '
        'cboAccountNo
        '
        Me.cboAccountNo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboAccountNo.FormattingEnabled = True
        Me.cboAccountNo.Location = New System.Drawing.Point(500, 39)
        Me.cboAccountNo.Name = "cboAccountNo"
        Me.cboAccountNo.Size = New System.Drawing.Size(284, 21)
        Me.cboAccountNo.TabIndex = 32
        '
        'btnRemove
        '
        Me.btnRemove.Location = New System.Drawing.Point(709, 71)
        Me.btnRemove.Name = "btnRemove"
        Me.btnRemove.Size = New System.Drawing.Size(75, 23)
        Me.btnRemove.TabIndex = 36
        Me.btnRemove.Text = "Remove"
        Me.btnRemove.UseVisualStyleBackColor = True
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label8.Location = New System.Drawing.Point(404, 43)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(90, 13)
        Me.Label8.TabIndex = 31
        Me.Label8.Text = "Account Number:"
        '
        'cboYear
        '
        Me.cboYear.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboYear.FormattingEnabled = True
        Me.cboYear.Location = New System.Drawing.Point(275, 73)
        Me.cboYear.Name = "cboYear"
        Me.cboYear.Size = New System.Drawing.Size(123, 21)
        Me.cboYear.TabIndex = 38
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.Location = New System.Drawing.Point(404, 11)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(74, 13)
        Me.Label7.TabIndex = 29
        Me.Label7.Text = "Deposited To:"
        '
        'cbomodeName
        '
        Me.cbomodeName.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbomodeName.FormattingEnabled = True
        Me.cbomodeName.Location = New System.Drawing.Point(484, 8)
        Me.cbomodeName.Name = "cbomodeName"
        Me.cbomodeName.Size = New System.Drawing.Size(106, 21)
        Me.cbomodeName.TabIndex = 30
        '
        'btnAdd
        '
        Me.btnAdd.Location = New System.Drawing.Point(599, 71)
        Me.btnAdd.Name = "btnAdd"
        Me.btnAdd.Size = New System.Drawing.Size(56, 23)
        Me.btnAdd.TabIndex = 35
        Me.btnAdd.Text = "Add"
        Me.btnAdd.UseVisualStyleBackColor = True
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.Location = New System.Drawing.Point(237, 76)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(32, 13)
        Me.Label6.TabIndex = 37
        Me.Label6.Text = "Year:"
        '
        'cboTermName
        '
        Me.cboTermName.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboTermName.FormattingEnabled = True
        Me.cboTermName.Location = New System.Drawing.Point(105, 73)
        Me.cboTermName.Name = "cboTermName"
        Me.cboTermName.Size = New System.Drawing.Size(126, 21)
        Me.cboTermName.TabIndex = 36
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.Location = New System.Drawing.Point(13, 76)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(65, 13)
        Me.Label5.TabIndex = 35
        Me.Label5.Text = "Term Name:"
        '
        'txtAmount
        '
        Me.txtAmount.Location = New System.Drawing.Point(456, 73)
        Me.txtAmount.Name = "txtAmount"
        Me.txtAmount.Size = New System.Drawing.Size(104, 20)
        Me.txtAmount.TabIndex = 34
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(404, 76)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(46, 13)
        Me.Label4.TabIndex = 33
        Me.Label4.Text = "Amount:"
        '
        'cboRecFrom
        '
        Me.cboRecFrom.FormattingEnabled = True
        Me.cboRecFrom.Location = New System.Drawing.Point(105, 39)
        Me.cboRecFrom.Name = "cboRecFrom"
        Me.cboRecFrom.Size = New System.Drawing.Size(293, 21)
        Me.cboRecFrom.TabIndex = 32
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(13, 43)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(73, 13)
        Me.Label3.TabIndex = 31
        Me.Label3.Text = "Receipt From:"
        '
        'cboSourceName
        '
        Me.cboSourceName.FormattingEnabled = True
        Me.cboSourceName.Location = New System.Drawing.Point(105, 8)
        Me.cboSourceName.Name = "cboSourceName"
        Me.cboSourceName.Size = New System.Drawing.Size(293, 21)
        Me.cboSourceName.TabIndex = 30
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(13, 11)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(75, 13)
        Me.Label2.TabIndex = 29
        Me.Label2.Text = "Source Name:"
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label14.Location = New System.Drawing.Point(596, 11)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(82, 13)
        Me.Label14.TabIndex = 25
        Me.Label14.Text = "Date Received:"
        '
        'dtpIncDate
        '
        Me.dtpIncDate.CustomFormat = "  dd-MM-yyyy"
        Me.dtpIncDate.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtpIncDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpIncDate.Location = New System.Drawing.Point(684, 8)
        Me.dtpIncDate.Name = "dtpIncDate"
        Me.dtpIncDate.Size = New System.Drawing.Size(100, 20)
        Me.dtpIncDate.TabIndex = 26
        '
        'StatusStrip3
        '
        Me.StatusStrip3.AutoSize = False
        Me.StatusStrip3.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip3.Dock = System.Windows.Forms.DockStyle.Left
        Me.StatusStrip3.Location = New System.Drawing.Point(0, 5)
        Me.StatusStrip3.Name = "StatusStrip3"
        Me.StatusStrip3.Size = New System.Drawing.Size(5, 106)
        Me.StatusStrip3.SizingGrip = False
        Me.StatusStrip3.TabIndex = 73
        Me.StatusStrip3.Text = "StatusStrip3"
        '
        'StatusStrip4
        '
        Me.StatusStrip4.AutoSize = False
        Me.StatusStrip4.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip4.Location = New System.Drawing.Point(0, 111)
        Me.StatusStrip4.Name = "StatusStrip4"
        Me.StatusStrip4.Size = New System.Drawing.Size(809, 5)
        Me.StatusStrip4.SizingGrip = False
        Me.StatusStrip4.TabIndex = 72
        Me.StatusStrip4.Text = "StatusStrip4"
        '
        'StatusStrip6
        '
        Me.StatusStrip6.AutoSize = False
        Me.StatusStrip6.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip6.Dock = System.Windows.Forms.DockStyle.Top
        Me.StatusStrip6.Location = New System.Drawing.Point(0, 0)
        Me.StatusStrip6.Name = "StatusStrip6"
        Me.StatusStrip6.Size = New System.Drawing.Size(809, 5)
        Me.StatusStrip6.SizingGrip = False
        Me.StatusStrip6.TabIndex = 71
        Me.StatusStrip6.Text = "StatusStrip6"
        '
        'StatusStrip7
        '
        Me.StatusStrip7.AutoSize = False
        Me.StatusStrip7.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip7.Dock = System.Windows.Forms.DockStyle.Right
        Me.StatusStrip7.Location = New System.Drawing.Point(809, 0)
        Me.StatusStrip7.Name = "StatusStrip7"
        Me.StatusStrip7.Size = New System.Drawing.Size(5, 116)
        Me.StatusStrip7.SizingGrip = False
        Me.StatusStrip7.TabIndex = 70
        Me.StatusStrip7.Text = "StatusStrip7"
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
        Me.SplitContainer2.Panel1.Controls.Add(Me.lstOtherInc)
        '
        'SplitContainer2.Panel2
        '
        Me.SplitContainer2.Panel2.Controls.Add(Me.Panel3)
        Me.SplitContainer2.Panel2.Controls.Add(Me.StatusStrip5)
        Me.SplitContainer2.Panel2.Controls.Add(Me.StatusStrip8)
        Me.SplitContainer2.Panel2.Controls.Add(Me.StatusStrip9)
        Me.SplitContainer2.Panel2.Controls.Add(Me.StatusStrip10)
        Me.SplitContainer2.Size = New System.Drawing.Size(816, 326)
        Me.SplitContainer2.SplitterDistance = 275
        Me.SplitContainer2.SplitterWidth = 1
        Me.SplitContainer2.TabIndex = 30
        '
        'lstOtherInc
        '
        Me.lstOtherInc.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColumnHeader6, Me.ColumnHeader7, Me.ColumnHeader8, Me.ColumnHeader9, Me.ColumnHeader10, Me.ColumnHeader11})
        Me.lstOtherInc.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lstOtherInc.FullRowSelect = True
        Me.lstOtherInc.GridLines = True
        Me.lstOtherInc.Location = New System.Drawing.Point(0, 0)
        Me.lstOtherInc.Name = "lstOtherInc"
        Me.lstOtherInc.Size = New System.Drawing.Size(814, 273)
        Me.lstOtherInc.TabIndex = 90
        Me.lstOtherInc.UseCompatibleStateImageBehavior = False
        Me.lstOtherInc.View = System.Windows.Forms.View.Details
        '
        'ColumnHeader6
        '
        Me.ColumnHeader6.Text = "Source Name"
        Me.ColumnHeader6.Width = 165
        '
        'ColumnHeader7
        '
        Me.ColumnHeader7.Text = "Receipt From"
        Me.ColumnHeader7.Width = 200
        '
        'ColumnHeader8
        '
        Me.ColumnHeader8.Text = "Term Name"
        Me.ColumnHeader8.Width = 80
        '
        'ColumnHeader9
        '
        Me.ColumnHeader9.Text = "Year"
        Me.ColumnHeader9.Width = 80
        '
        'ColumnHeader10
        '
        Me.ColumnHeader10.Text = "Account Number"
        Me.ColumnHeader10.Width = 200
        '
        'ColumnHeader11
        '
        Me.ColumnHeader11.Text = "Amount"
        Me.ColumnHeader11.Width = 80
        '
        'Panel3
        '
        Me.Panel3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel3.Controls.Add(Me.btnClear)
        Me.Panel3.Controls.Add(Me.btnClose)
        Me.Panel3.Controls.Add(Me.btnUpdate)
        Me.Panel3.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel3.Location = New System.Drawing.Point(5, 5)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(804, 38)
        Me.Panel3.TabIndex = 74
        '
        'btnClear
        '
        Me.btnClear.Location = New System.Drawing.Point(16, 6)
        Me.btnClear.Name = "btnClear"
        Me.btnClear.Size = New System.Drawing.Size(75, 23)
        Me.btnClear.TabIndex = 37
        Me.btnClear.Text = "Clear"
        Me.btnClear.UseVisualStyleBackColor = True
        '
        'btnClose
        '
        Me.btnClose.Location = New System.Drawing.Point(709, 6)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(75, 23)
        Me.btnClose.TabIndex = 24
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'btnUpdate
        '
        Me.btnUpdate.Location = New System.Drawing.Point(485, 6)
        Me.btnUpdate.Name = "btnUpdate"
        Me.btnUpdate.Size = New System.Drawing.Size(75, 23)
        Me.btnUpdate.TabIndex = 23
        Me.btnUpdate.Text = "Update"
        Me.btnUpdate.UseVisualStyleBackColor = True
        '
        'StatusStrip5
        '
        Me.StatusStrip5.AutoSize = False
        Me.StatusStrip5.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip5.Dock = System.Windows.Forms.DockStyle.Left
        Me.StatusStrip5.Location = New System.Drawing.Point(0, 5)
        Me.StatusStrip5.Name = "StatusStrip5"
        Me.StatusStrip5.Size = New System.Drawing.Size(5, 38)
        Me.StatusStrip5.SizingGrip = False
        Me.StatusStrip5.TabIndex = 73
        Me.StatusStrip5.Text = "StatusStrip5"
        '
        'StatusStrip8
        '
        Me.StatusStrip8.AutoSize = False
        Me.StatusStrip8.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip8.Location = New System.Drawing.Point(0, 43)
        Me.StatusStrip8.Name = "StatusStrip8"
        Me.StatusStrip8.Size = New System.Drawing.Size(809, 5)
        Me.StatusStrip8.SizingGrip = False
        Me.StatusStrip8.TabIndex = 72
        Me.StatusStrip8.Text = "StatusStrip8"
        '
        'StatusStrip9
        '
        Me.StatusStrip9.AutoSize = False
        Me.StatusStrip9.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip9.Dock = System.Windows.Forms.DockStyle.Top
        Me.StatusStrip9.Location = New System.Drawing.Point(0, 0)
        Me.StatusStrip9.Name = "StatusStrip9"
        Me.StatusStrip9.Size = New System.Drawing.Size(809, 5)
        Me.StatusStrip9.SizingGrip = False
        Me.StatusStrip9.TabIndex = 71
        Me.StatusStrip9.Text = "StatusStrip9"
        '
        'StatusStrip10
        '
        Me.StatusStrip10.AutoSize = False
        Me.StatusStrip10.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip10.Dock = System.Windows.Forms.DockStyle.Right
        Me.StatusStrip10.Location = New System.Drawing.Point(809, 0)
        Me.StatusStrip10.Name = "StatusStrip10"
        Me.StatusStrip10.Size = New System.Drawing.Size(5, 48)
        Me.StatusStrip10.SizingGrip = False
        Me.StatusStrip10.TabIndex = 70
        Me.StatusStrip10.Text = "StatusStrip10"
        '
        'StatusStrip2
        '
        Me.StatusStrip2.AutoSize = False
        Me.StatusStrip2.BackColor = System.Drawing.Color.LightSeaGreen
        Me.StatusStrip2.Dock = System.Windows.Forms.DockStyle.Top
        Me.StatusStrip2.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripStatusLabel1})
        Me.StatusStrip2.Location = New System.Drawing.Point(0, 0)
        Me.StatusStrip2.Name = "StatusStrip2"
        Me.StatusStrip2.Size = New System.Drawing.Size(816, 25)
        Me.StatusStrip2.SizingGrip = False
        Me.StatusStrip2.TabIndex = 28
        Me.StatusStrip2.Text = "StatusStrip2"
        '
        'ToolStripStatusLabel1
        '
        Me.ToolStripStatusLabel1.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ToolStripStatusLabel1.Name = "ToolStripStatusLabel1"
        Me.ToolStripStatusLabel1.Size = New System.Drawing.Size(105, 20)
        Me.ToolStripStatusLabel1.Text = "OTHER INCOME"
        '
        'StatusStrip1
        '
        Me.StatusStrip1.AutoSize = False
        Me.StatusStrip1.BackColor = System.Drawing.Color.LightSeaGreen
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 470)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(816, 25)
        Me.StatusStrip1.SizingGrip = False
        Me.StatusStrip1.TabIndex = 27
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'frmFinOtherIncome
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.LightSkyBlue
        Me.ClientSize = New System.Drawing.Size(843, 521)
        Me.Controls.Add(Me.pnlOtherInc)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Name = "frmFinOtherIncome"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Other Income"
        Me.pnlOtherInc.ResumeLayout(False)
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.ResumeLayout(False)
        Me.Panel2.ResumeLayout(False)
        Me.Panel2.PerformLayout()
        Me.SplitContainer2.Panel1.ResumeLayout(False)
        Me.SplitContainer2.Panel2.ResumeLayout(False)
        CType(Me.SplitContainer2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer2.ResumeLayout(False)
        Me.Panel3.ResumeLayout(False)
        Me.StatusStrip2.ResumeLayout(False)
        Me.StatusStrip2.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents pnlOtherInc As System.Windows.Forms.Panel
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents SplitContainer2 As System.Windows.Forms.SplitContainer
    Friend WithEvents StatusStrip2 As System.Windows.Forms.StatusStrip
    Friend WithEvents ToolStripStatusLabel1 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents lstOtherInc As System.Windows.Forms.ListView
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents StatusStrip3 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip4 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip6 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip7 As System.Windows.Forms.StatusStrip
    Friend WithEvents Panel3 As System.Windows.Forms.Panel
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents btnUpdate As System.Windows.Forms.Button
    Friend WithEvents StatusStrip5 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip8 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip9 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip10 As System.Windows.Forms.StatusStrip
    Friend WithEvents dtpIncDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents btnRemove As System.Windows.Forms.Button
    Friend WithEvents cboYear As System.Windows.Forms.ComboBox
    Friend WithEvents btnAdd As System.Windows.Forms.Button
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents cboTermName As System.Windows.Forms.ComboBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents txtAmount As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents cboRecFrom As System.Windows.Forms.ComboBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents cboSourceName As System.Windows.Forms.ComboBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents cboAccountNo As System.Windows.Forms.ComboBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents cbomodeName As System.Windows.Forms.ComboBox
    Friend WithEvents btnClear As System.Windows.Forms.Button
    Friend WithEvents ColumnHeader6 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader7 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader8 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader9 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader10 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader11 As System.Windows.Forms.ColumnHeader
End Class
