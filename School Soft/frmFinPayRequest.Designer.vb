<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmFinPayRequest
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
        Me.pnlPayReq = New System.Windows.Forms.Panel()
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.cboPayYear = New System.Windows.Forms.ComboBox()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.cboPayTerm = New System.Windows.Forms.ComboBox()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.btnAdd = New System.Windows.Forms.Button()
        Me.btnRemove = New System.Windows.Forms.Button()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.txtVoucherNo = New System.Windows.Forms.TextBox()
        Me.dtpPayDate = New System.Windows.Forms.DateTimePicker()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.cboMoneyPaidTo = New System.Windows.Forms.ComboBox()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.txtAmount = New System.Windows.Forms.TextBox()
        Me.cboExpName = New System.Windows.Forms.ComboBox()
        Me.cboExpCat = New System.Windows.Forms.ComboBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.txtBalance = New System.Windows.Forms.TextBox()
        Me.cboAccountNo = New System.Windows.Forms.ComboBox()
        Me.cboPayMode = New System.Windows.Forms.ComboBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.StatusStrip3 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip4 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip6 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip7 = New System.Windows.Forms.StatusStrip()
        Me.SplitContainer2 = New System.Windows.Forms.SplitContainer()
        Me.lstPayRequest = New System.Windows.Forms.ListView()
        Me.ColumnHeader1 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader2 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader3 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader4 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader5 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader6 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader7 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.txtVendorName = New System.Windows.Forms.TextBox()
        Me.btnUpdate = New System.Windows.Forms.Button()
        Me.btnClose = New System.Windows.Forms.Button()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.StatusStrip5 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip8 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip9 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip10 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip2 = New System.Windows.Forms.StatusStrip()
        Me.ToolStripStatusLabel1 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip()
        Me.pnlPayReq.SuspendLayout()
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.Panel1.SuspendLayout()
        CType(Me.SplitContainer2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer2.Panel1.SuspendLayout()
        Me.SplitContainer2.Panel2.SuspendLayout()
        Me.SplitContainer2.SuspendLayout()
        Me.Panel3.SuspendLayout()
        Me.StatusStrip2.SuspendLayout()
        Me.SuspendLayout()
        '
        'pnlPayReq
        '
        Me.pnlPayReq.BackColor = System.Drawing.SystemColors.Control
        Me.pnlPayReq.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pnlPayReq.Controls.Add(Me.SplitContainer1)
        Me.pnlPayReq.Controls.Add(Me.StatusStrip2)
        Me.pnlPayReq.Controls.Add(Me.StatusStrip1)
        Me.pnlPayReq.Location = New System.Drawing.Point(12, 12)
        Me.pnlPayReq.Name = "pnlPayReq"
        Me.pnlPayReq.Size = New System.Drawing.Size(963, 524)
        Me.pnlPayReq.TabIndex = 0
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
        Me.SplitContainer1.Panel1.Controls.Add(Me.StatusStrip6)
        Me.SplitContainer1.Panel1.Controls.Add(Me.StatusStrip7)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.SplitContainer2)
        Me.SplitContainer1.Size = New System.Drawing.Size(961, 472)
        Me.SplitContainer1.SplitterDistance = 131
        Me.SplitContainer1.SplitterWidth = 1
        Me.SplitContainer1.TabIndex = 33
        '
        'Panel1
        '
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel1.Controls.Add(Me.cboPayYear)
        Me.Panel1.Controls.Add(Me.Label12)
        Me.Panel1.Controls.Add(Me.cboPayTerm)
        Me.Panel1.Controls.Add(Me.Label11)
        Me.Panel1.Controls.Add(Me.btnAdd)
        Me.Panel1.Controls.Add(Me.btnRemove)
        Me.Panel1.Controls.Add(Me.Label10)
        Me.Panel1.Controls.Add(Me.txtVoucherNo)
        Me.Panel1.Controls.Add(Me.dtpPayDate)
        Me.Panel1.Controls.Add(Me.Label9)
        Me.Panel1.Controls.Add(Me.cboMoneyPaidTo)
        Me.Panel1.Controls.Add(Me.Label8)
        Me.Panel1.Controls.Add(Me.Label7)
        Me.Panel1.Controls.Add(Me.txtAmount)
        Me.Panel1.Controls.Add(Me.cboExpName)
        Me.Panel1.Controls.Add(Me.cboExpCat)
        Me.Panel1.Controls.Add(Me.Label5)
        Me.Panel1.Controls.Add(Me.Label6)
        Me.Panel1.Controls.Add(Me.txtBalance)
        Me.Panel1.Controls.Add(Me.cboAccountNo)
        Me.Panel1.Controls.Add(Me.cboPayMode)
        Me.Panel1.Controls.Add(Me.Label4)
        Me.Panel1.Controls.Add(Me.Label3)
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel1.Location = New System.Drawing.Point(5, 5)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(949, 119)
        Me.Panel1.TabIndex = 88
        '
        'cboPayYear
        '
        Me.cboPayYear.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboPayYear.FormattingEnabled = True
        Me.cboPayYear.Location = New System.Drawing.Point(329, 6)
        Me.cboPayYear.Name = "cboPayYear"
        Me.cboPayYear.Size = New System.Drawing.Size(136, 21)
        Me.cboPayYear.TabIndex = 49
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Location = New System.Drawing.Point(244, 9)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(79, 13)
        Me.Label12.TabIndex = 48
        Me.Label12.Text = "Payment Year :"
        '
        'cboPayTerm
        '
        Me.cboPayTerm.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboPayTerm.FormattingEnabled = True
        Me.cboPayTerm.Location = New System.Drawing.Point(109, 6)
        Me.cboPayTerm.Name = "cboPayTerm"
        Me.cboPayTerm.Size = New System.Drawing.Size(129, 21)
        Me.cboPayTerm.TabIndex = 47
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(13, 9)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(78, 13)
        Me.Label11.TabIndex = 46
        Me.Label11.Text = "Payment Term:"
        '
        'btnAdd
        '
        Me.btnAdd.Location = New System.Drawing.Point(744, 89)
        Me.btnAdd.Name = "btnAdd"
        Me.btnAdd.Size = New System.Drawing.Size(75, 23)
        Me.btnAdd.TabIndex = 45
        Me.btnAdd.Text = "Add"
        Me.btnAdd.UseVisualStyleBackColor = True
        '
        'btnRemove
        '
        Me.btnRemove.Location = New System.Drawing.Point(859, 89)
        Me.btnRemove.Name = "btnRemove"
        Me.btnRemove.Size = New System.Drawing.Size(75, 23)
        Me.btnRemove.TabIndex = 44
        Me.btnRemove.Text = "Remove"
        Me.btnRemove.UseVisualStyleBackColor = True
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(244, 38)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(67, 13)
        Me.Label10.TabIndex = 43
        Me.Label10.Text = "Voucher No:"
        '
        'txtVoucherNo
        '
        Me.txtVoucherNo.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtVoucherNo.Location = New System.Drawing.Point(317, 35)
        Me.txtVoucherNo.Name = "txtVoucherNo"
        Me.txtVoucherNo.ReadOnly = True
        Me.txtVoucherNo.Size = New System.Drawing.Size(148, 20)
        Me.txtVoucherNo.TabIndex = 42
        '
        'dtpPayDate
        '
        Me.dtpPayDate.CustomFormat = "  dd - MM - yyyy"
        Me.dtpPayDate.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtpPayDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpPayDate.Location = New System.Drawing.Point(355, 91)
        Me.dtpPayDate.Name = "dtpPayDate"
        Me.dtpPayDate.Size = New System.Drawing.Size(110, 20)
        Me.dtpPayDate.TabIndex = 41
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(272, 94)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(77, 13)
        Me.Label9.TabIndex = 40
        Me.Label9.Text = "Payment Date:"
        '
        'cboMoneyPaidTo
        '
        Me.cboMoneyPaidTo.FormattingEnabled = True
        Me.cboMoneyPaidTo.Location = New System.Drawing.Point(584, 65)
        Me.cboMoneyPaidTo.Name = "cboMoneyPaidTo"
        Me.cboMoneyPaidTo.Size = New System.Drawing.Size(350, 21)
        Me.cboMoneyPaidTo.TabIndex = 39
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(482, 68)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(82, 13)
        Me.Label8.TabIndex = 38
        Me.Label8.Text = "Money Paid To:"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(482, 94)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(90, 13)
        Me.Label7.TabIndex = 37
        Me.Label7.Text = "Expense Amount:"
        '
        'txtAmount
        '
        Me.txtAmount.Location = New System.Drawing.Point(584, 91)
        Me.txtAmount.Name = "txtAmount"
        Me.txtAmount.Size = New System.Drawing.Size(140, 20)
        Me.txtAmount.TabIndex = 36
        '
        'cboExpName
        '
        Me.cboExpName.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboExpName.FormattingEnabled = True
        Me.cboExpName.Location = New System.Drawing.Point(584, 37)
        Me.cboExpName.Name = "cboExpName"
        Me.cboExpName.Size = New System.Drawing.Size(350, 21)
        Me.cboExpName.TabIndex = 35
        '
        'cboExpCat
        '
        Me.cboExpCat.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboExpCat.FormattingEnabled = True
        Me.cboExpCat.Location = New System.Drawing.Point(584, 8)
        Me.cboExpCat.Name = "cboExpCat"
        Me.cboExpCat.Size = New System.Drawing.Size(350, 21)
        Me.cboExpCat.TabIndex = 34
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(482, 40)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(82, 13)
        Me.Label5.TabIndex = 33
        Me.Label5.Text = "Expense Name:"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(482, 11)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(96, 13)
        Me.Label6.TabIndex = 32
        Me.Label6.Text = "Expense Category:"
        '
        'txtBalance
        '
        Me.txtBalance.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtBalance.Location = New System.Drawing.Point(109, 91)
        Me.txtBalance.Name = "txtBalance"
        Me.txtBalance.ReadOnly = True
        Me.txtBalance.Size = New System.Drawing.Size(157, 20)
        Me.txtBalance.TabIndex = 31
        '
        'cboAccountNo
        '
        Me.cboAccountNo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboAccountNo.FormattingEnabled = True
        Me.cboAccountNo.Location = New System.Drawing.Point(109, 64)
        Me.cboAccountNo.Name = "cboAccountNo"
        Me.cboAccountNo.Size = New System.Drawing.Size(356, 21)
        Me.cboAccountNo.TabIndex = 30
        '
        'cboPayMode
        '
        Me.cboPayMode.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboPayMode.FormattingEnabled = True
        Me.cboPayMode.Location = New System.Drawing.Point(109, 35)
        Me.cboPayMode.Name = "cboPayMode"
        Me.cboPayMode.Size = New System.Drawing.Size(129, 21)
        Me.cboPayMode.TabIndex = 28
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(13, 94)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(89, 13)
        Me.Label4.TabIndex = 3
        Me.Label4.Text = "Balance ( KSH ) :"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(13, 67)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(90, 13)
        Me.Label3.TabIndex = 2
        Me.Label3.Text = "Account Number:"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(13, 38)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(81, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Payment Mode:"
        '
        'StatusStrip3
        '
        Me.StatusStrip3.AutoSize = False
        Me.StatusStrip3.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip3.Dock = System.Windows.Forms.DockStyle.Left
        Me.StatusStrip3.Location = New System.Drawing.Point(0, 5)
        Me.StatusStrip3.Name = "StatusStrip3"
        Me.StatusStrip3.Size = New System.Drawing.Size(5, 119)
        Me.StatusStrip3.SizingGrip = False
        Me.StatusStrip3.TabIndex = 87
        Me.StatusStrip3.Text = "StatusStrip3"
        '
        'StatusStrip4
        '
        Me.StatusStrip4.AutoSize = False
        Me.StatusStrip4.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip4.Location = New System.Drawing.Point(0, 124)
        Me.StatusStrip4.Name = "StatusStrip4"
        Me.StatusStrip4.Size = New System.Drawing.Size(954, 5)
        Me.StatusStrip4.SizingGrip = False
        Me.StatusStrip4.TabIndex = 86
        Me.StatusStrip4.Text = "StatusStrip4"
        '
        'StatusStrip6
        '
        Me.StatusStrip6.AutoSize = False
        Me.StatusStrip6.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip6.Dock = System.Windows.Forms.DockStyle.Top
        Me.StatusStrip6.Location = New System.Drawing.Point(0, 0)
        Me.StatusStrip6.Name = "StatusStrip6"
        Me.StatusStrip6.Size = New System.Drawing.Size(954, 5)
        Me.StatusStrip6.SizingGrip = False
        Me.StatusStrip6.TabIndex = 85
        Me.StatusStrip6.Text = "StatusStrip6"
        '
        'StatusStrip7
        '
        Me.StatusStrip7.AutoSize = False
        Me.StatusStrip7.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip7.Dock = System.Windows.Forms.DockStyle.Right
        Me.StatusStrip7.Location = New System.Drawing.Point(954, 0)
        Me.StatusStrip7.Name = "StatusStrip7"
        Me.StatusStrip7.Size = New System.Drawing.Size(5, 129)
        Me.StatusStrip7.SizingGrip = False
        Me.StatusStrip7.TabIndex = 84
        Me.StatusStrip7.Text = "StatusStrip7"
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
        Me.SplitContainer2.Panel1.Controls.Add(Me.lstPayRequest)
        '
        'SplitContainer2.Panel2
        '
        Me.SplitContainer2.Panel2.Controls.Add(Me.Panel3)
        Me.SplitContainer2.Panel2.Controls.Add(Me.StatusStrip5)
        Me.SplitContainer2.Panel2.Controls.Add(Me.StatusStrip8)
        Me.SplitContainer2.Panel2.Controls.Add(Me.StatusStrip9)
        Me.SplitContainer2.Panel2.Controls.Add(Me.StatusStrip10)
        Me.SplitContainer2.Size = New System.Drawing.Size(959, 338)
        Me.SplitContainer2.SplitterDistance = 287
        Me.SplitContainer2.SplitterWidth = 1
        Me.SplitContainer2.TabIndex = 34
        '
        'lstPayRequest
        '
        Me.lstPayRequest.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColumnHeader1, Me.ColumnHeader2, Me.ColumnHeader3, Me.ColumnHeader4, Me.ColumnHeader5, Me.ColumnHeader6, Me.ColumnHeader7})
        Me.lstPayRequest.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lstPayRequest.FullRowSelect = True
        Me.lstPayRequest.GridLines = True
        Me.lstPayRequest.Location = New System.Drawing.Point(0, 0)
        Me.lstPayRequest.Name = "lstPayRequest"
        Me.lstPayRequest.Size = New System.Drawing.Size(959, 287)
        Me.lstPayRequest.TabIndex = 93
        Me.lstPayRequest.UseCompatibleStateImageBehavior = False
        Me.lstPayRequest.View = System.Windows.Forms.View.Details
        '
        'ColumnHeader1
        '
        Me.ColumnHeader1.Text = "Voucher No"
        Me.ColumnHeader1.Width = 130
        '
        'ColumnHeader2
        '
        Me.ColumnHeader2.Text = "Payment Mode"
        Me.ColumnHeader2.Width = 100
        '
        'ColumnHeader3
        '
        Me.ColumnHeader3.Text = "Account Number"
        Me.ColumnHeader3.Width = 170
        '
        'ColumnHeader4
        '
        Me.ColumnHeader4.Text = "Request Date"
        Me.ColumnHeader4.Width = 100
        '
        'ColumnHeader5
        '
        Me.ColumnHeader5.Text = "Expense Name"
        Me.ColumnHeader5.Width = 150
        '
        'ColumnHeader6
        '
        Me.ColumnHeader6.Text = "Paid To"
        Me.ColumnHeader6.Width = 200
        '
        'ColumnHeader7
        '
        Me.ColumnHeader7.Text = "Amount"
        Me.ColumnHeader7.Width = 100
        '
        'Panel3
        '
        Me.Panel3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel3.Controls.Add(Me.txtVendorName)
        Me.Panel3.Controls.Add(Me.btnUpdate)
        Me.Panel3.Controls.Add(Me.btnClose)
        Me.Panel3.Controls.Add(Me.Label2)
        Me.Panel3.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel3.Location = New System.Drawing.Point(5, 5)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(949, 40)
        Me.Panel3.TabIndex = 84
        '
        'txtVendorName
        '
        Me.txtVendorName.Location = New System.Drawing.Point(124, 9)
        Me.txtVendorName.Name = "txtVendorName"
        Me.txtVendorName.ReadOnly = True
        Me.txtVendorName.Size = New System.Drawing.Size(356, 20)
        Me.txtVendorName.TabIndex = 46
        '
        'btnUpdate
        '
        Me.btnUpdate.Location = New System.Drawing.Point(649, 7)
        Me.btnUpdate.Name = "btnUpdate"
        Me.btnUpdate.Size = New System.Drawing.Size(75, 23)
        Me.btnUpdate.TabIndex = 27
        Me.btnUpdate.Text = "Save"
        Me.btnUpdate.UseVisualStyleBackColor = True
        '
        'btnClose
        '
        Me.btnClose.Location = New System.Drawing.Point(859, 7)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(75, 23)
        Me.btnClose.TabIndex = 23
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(28, 12)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(75, 13)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "Vendor Name:"
        '
        'StatusStrip5
        '
        Me.StatusStrip5.AutoSize = False
        Me.StatusStrip5.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip5.Dock = System.Windows.Forms.DockStyle.Left
        Me.StatusStrip5.Location = New System.Drawing.Point(0, 5)
        Me.StatusStrip5.Name = "StatusStrip5"
        Me.StatusStrip5.Size = New System.Drawing.Size(5, 40)
        Me.StatusStrip5.SizingGrip = False
        Me.StatusStrip5.TabIndex = 83
        Me.StatusStrip5.Text = "StatusStrip5"
        '
        'StatusStrip8
        '
        Me.StatusStrip8.AutoSize = False
        Me.StatusStrip8.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip8.Location = New System.Drawing.Point(0, 45)
        Me.StatusStrip8.Name = "StatusStrip8"
        Me.StatusStrip8.Size = New System.Drawing.Size(954, 5)
        Me.StatusStrip8.SizingGrip = False
        Me.StatusStrip8.TabIndex = 82
        Me.StatusStrip8.Text = "StatusStrip8"
        '
        'StatusStrip9
        '
        Me.StatusStrip9.AutoSize = False
        Me.StatusStrip9.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip9.Dock = System.Windows.Forms.DockStyle.Top
        Me.StatusStrip9.Location = New System.Drawing.Point(0, 0)
        Me.StatusStrip9.Name = "StatusStrip9"
        Me.StatusStrip9.Size = New System.Drawing.Size(954, 5)
        Me.StatusStrip9.SizingGrip = False
        Me.StatusStrip9.TabIndex = 81
        Me.StatusStrip9.Text = "StatusStrip9"
        '
        'StatusStrip10
        '
        Me.StatusStrip10.AutoSize = False
        Me.StatusStrip10.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip10.Dock = System.Windows.Forms.DockStyle.Right
        Me.StatusStrip10.Location = New System.Drawing.Point(954, 0)
        Me.StatusStrip10.Name = "StatusStrip10"
        Me.StatusStrip10.Size = New System.Drawing.Size(5, 50)
        Me.StatusStrip10.SizingGrip = False
        Me.StatusStrip10.TabIndex = 80
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
        Me.StatusStrip2.Size = New System.Drawing.Size(961, 25)
        Me.StatusStrip2.SizingGrip = False
        Me.StatusStrip2.TabIndex = 32
        Me.StatusStrip2.Text = "StatusStrip2"
        '
        'ToolStripStatusLabel1
        '
        Me.ToolStripStatusLabel1.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ToolStripStatusLabel1.Name = "ToolStripStatusLabel1"
        Me.ToolStripStatusLabel1.Size = New System.Drawing.Size(128, 20)
        Me.ToolStripStatusLabel1.Text = "PAYMENT REQUEST"
        '
        'StatusStrip1
        '
        Me.StatusStrip1.AutoSize = False
        Me.StatusStrip1.BackColor = System.Drawing.Color.LightSeaGreen
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 497)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(961, 25)
        Me.StatusStrip1.SizingGrip = False
        Me.StatusStrip1.TabIndex = 31
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'frmFinPayRequest
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.LightSkyBlue
        Me.ClientSize = New System.Drawing.Size(985, 546)
        Me.Controls.Add(Me.pnlPayReq)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Name = "frmFinPayRequest"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Payment Request"
        Me.pnlPayReq.ResumeLayout(False)
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
        Me.Panel3.ResumeLayout(False)
        Me.Panel3.PerformLayout()
        Me.StatusStrip2.ResumeLayout(False)
        Me.StatusStrip2.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents pnlPayReq As System.Windows.Forms.Panel
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents SplitContainer2 As System.Windows.Forms.SplitContainer
    Friend WithEvents StatusStrip2 As System.Windows.Forms.StatusStrip
    Friend WithEvents ToolStripStatusLabel1 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents lstPayRequest As System.Windows.Forms.ListView
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents StatusStrip3 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip4 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip6 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip7 As System.Windows.Forms.StatusStrip
    Friend WithEvents Panel3 As System.Windows.Forms.Panel
    Friend WithEvents btnUpdate As System.Windows.Forms.Button
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents StatusStrip5 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip8 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip9 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip10 As System.Windows.Forms.StatusStrip
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents cboAccountNo As System.Windows.Forms.ComboBox
    Friend WithEvents cboPayMode As System.Windows.Forms.ComboBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents cboMoneyPaidTo As System.Windows.Forms.ComboBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents txtAmount As System.Windows.Forms.TextBox
    Friend WithEvents cboExpName As System.Windows.Forms.ComboBox
    Friend WithEvents cboExpCat As System.Windows.Forms.ComboBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents btnAdd As System.Windows.Forms.Button
    Friend WithEvents btnRemove As System.Windows.Forms.Button
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents txtVoucherNo As System.Windows.Forms.TextBox
    Friend WithEvents dtpPayDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents ColumnHeader1 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader2 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader3 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader4 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader5 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader6 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader7 As System.Windows.Forms.ColumnHeader
    Friend WithEvents cboPayYear As System.Windows.Forms.ComboBox
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents cboPayTerm As System.Windows.Forms.ComboBox
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents txtBalance As System.Windows.Forms.TextBox
    Friend WithEvents txtVendorName As System.Windows.Forms.TextBox
End Class
