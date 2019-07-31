<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmInvReceipts
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
        Me.pnlInvReceipts = New System.Windows.Forms.Panel()
        Me.SplitContainer2 = New System.Windows.Forms.SplitContainer()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.GroupBox4 = New System.Windows.Forms.GroupBox()
        Me.btnAdd = New System.Windows.Forms.Button()
        Me.btnClose = New System.Windows.Forms.Button()
        Me.btnRemove = New System.Windows.Forms.Button()
        Me.btnSave = New System.Windows.Forms.Button()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.txtSuppName = New System.Windows.Forms.TextBox()
        Me.cboSuppCode = New System.Windows.Forms.ComboBox()
        Me.cboSuppCat = New System.Windows.Forms.ComboBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.txtQtyRec = New System.Windows.Forms.TextBox()
        Me.txtUOM = New System.Windows.Forms.TextBox()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.txtCostPerUnit = New System.Windows.Forms.TextBox()
        Me.txtStockQty = New System.Windows.Forms.TextBox()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.txtTotalCost = New System.Windows.Forms.TextBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.txtItemName = New System.Windows.Forms.TextBox()
        Me.cboItemCode = New System.Windows.Forms.ComboBox()
        Me.cboItemCategory = New System.Windows.Forms.ComboBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.txtRecNo = New System.Windows.Forms.TextBox()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.dtpRecDate = New System.Windows.Forms.DateTimePicker()
        Me.StatusStrip11 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip12 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip13 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip14 = New System.Windows.Forms.StatusStrip()
        Me.lstInvReceipts = New System.Windows.Forms.ListView()
        Me.ColumnHeader1 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader2 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader3 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader4 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader6 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader7 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader5 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader8 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader9 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.StatusStrip2 = New System.Windows.Forms.StatusStrip()
        Me.ToolStripStatusLabel1 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip()
        Me.pnlInvReceipts.SuspendLayout()
        CType(Me.SplitContainer2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer2.Panel1.SuspendLayout()
        Me.SplitContainer2.Panel2.SuspendLayout()
        Me.SplitContainer2.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.GroupBox4.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.StatusStrip2.SuspendLayout()
        Me.SuspendLayout()
        '
        'pnlInvReceipts
        '
        Me.pnlInvReceipts.BackColor = System.Drawing.SystemColors.Control
        Me.pnlInvReceipts.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pnlInvReceipts.Controls.Add(Me.SplitContainer2)
        Me.pnlInvReceipts.Controls.Add(Me.StatusStrip2)
        Me.pnlInvReceipts.Controls.Add(Me.StatusStrip1)
        Me.pnlInvReceipts.Location = New System.Drawing.Point(12, 12)
        Me.pnlInvReceipts.Name = "pnlInvReceipts"
        Me.pnlInvReceipts.Size = New System.Drawing.Size(951, 541)
        Me.pnlInvReceipts.TabIndex = 0
        '
        'SplitContainer2
        '
        Me.SplitContainer2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.SplitContainer2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer2.Location = New System.Drawing.Point(0, 25)
        Me.SplitContainer2.Name = "SplitContainer2"
        Me.SplitContainer2.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'SplitContainer2.Panel1
        '
        Me.SplitContainer2.Panel1.Controls.Add(Me.Panel1)
        Me.SplitContainer2.Panel1.Controls.Add(Me.StatusStrip11)
        Me.SplitContainer2.Panel1.Controls.Add(Me.StatusStrip12)
        Me.SplitContainer2.Panel1.Controls.Add(Me.StatusStrip13)
        Me.SplitContainer2.Panel1.Controls.Add(Me.StatusStrip14)
        '
        'SplitContainer2.Panel2
        '
        Me.SplitContainer2.Panel2.Controls.Add(Me.lstInvReceipts)
        Me.SplitContainer2.Size = New System.Drawing.Size(949, 489)
        Me.SplitContainer2.SplitterDistance = 211
        Me.SplitContainer2.SplitterWidth = 1
        Me.SplitContainer2.TabIndex = 35
        '
        'Panel1
        '
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel1.Controls.Add(Me.GroupBox4)
        Me.Panel1.Controls.Add(Me.Label13)
        Me.Panel1.Controls.Add(Me.GroupBox2)
        Me.Panel1.Controls.Add(Me.GroupBox3)
        Me.Panel1.Controls.Add(Me.GroupBox1)
        Me.Panel1.Controls.Add(Me.txtRecNo)
        Me.Panel1.Controls.Add(Me.Label10)
        Me.Panel1.Controls.Add(Me.dtpRecDate)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel1.Location = New System.Drawing.Point(5, 5)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(937, 199)
        Me.Panel1.TabIndex = 63
        '
        'GroupBox4
        '
        Me.GroupBox4.Controls.Add(Me.btnAdd)
        Me.GroupBox4.Controls.Add(Me.btnClose)
        Me.GroupBox4.Controls.Add(Me.btnRemove)
        Me.GroupBox4.Controls.Add(Me.btnSave)
        Me.GroupBox4.Location = New System.Drawing.Point(474, 141)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Size = New System.Drawing.Size(450, 51)
        Me.GroupBox4.TabIndex = 86
        Me.GroupBox4.TabStop = False
        '
        'btnAdd
        '
        Me.btnAdd.Location = New System.Drawing.Point(12, 17)
        Me.btnAdd.Name = "btnAdd"
        Me.btnAdd.Size = New System.Drawing.Size(75, 23)
        Me.btnAdd.TabIndex = 82
        Me.btnAdd.Text = "Add"
        Me.btnAdd.UseVisualStyleBackColor = True
        '
        'btnClose
        '
        Me.btnClose.Location = New System.Drawing.Point(366, 17)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(75, 23)
        Me.btnClose.TabIndex = 83
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'btnRemove
        '
        Me.btnRemove.Location = New System.Drawing.Point(129, 17)
        Me.btnRemove.Name = "btnRemove"
        Me.btnRemove.Size = New System.Drawing.Size(75, 23)
        Me.btnRemove.TabIndex = 83
        Me.btnRemove.Text = "Remove"
        Me.btnRemove.UseVisualStyleBackColor = True
        '
        'btnSave
        '
        Me.btnSave.Location = New System.Drawing.Point(249, 17)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(75, 23)
        Me.btnSave.TabIndex = 84
        Me.btnSave.Text = "Save"
        Me.btnSave.UseVisualStyleBackColor = True
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Location = New System.Drawing.Point(693, 12)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(67, 13)
        Me.Label13.TabIndex = 57
        Me.Label13.Text = "Receipt No :"
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.txtSuppName)
        Me.GroupBox2.Controls.Add(Me.cboSuppCode)
        Me.GroupBox2.Controls.Add(Me.cboSuppCat)
        Me.GroupBox2.Controls.Add(Me.Label3)
        Me.GroupBox2.Controls.Add(Me.Label2)
        Me.GroupBox2.Controls.Add(Me.Label1)
        Me.GroupBox2.Location = New System.Drawing.Point(3, 3)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(462, 92)
        Me.GroupBox2.TabIndex = 1
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Supplier Details"
        '
        'txtSuppName
        '
        Me.txtSuppName.Location = New System.Drawing.Point(69, 58)
        Me.txtSuppName.Name = "txtSuppName"
        Me.txtSuppName.ReadOnly = True
        Me.txtSuppName.Size = New System.Drawing.Size(385, 20)
        Me.txtSuppName.TabIndex = 24
        '
        'cboSuppCode
        '
        Me.cboSuppCode.DropDownHeight = 100
        Me.cboSuppCode.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboSuppCode.FormattingEnabled = True
        Me.cboSuppCode.IntegralHeight = False
        Me.cboSuppCode.Location = New System.Drawing.Point(291, 25)
        Me.cboSuppCode.Name = "cboSuppCode"
        Me.cboSuppCode.Size = New System.Drawing.Size(163, 21)
        Me.cboSuppCode.TabIndex = 12
        '
        'cboSuppCat
        '
        Me.cboSuppCat.DropDownHeight = 100
        Me.cboSuppCat.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboSuppCat.FormattingEnabled = True
        Me.cboSuppCat.IntegralHeight = False
        Me.cboSuppCat.Location = New System.Drawing.Point(69, 25)
        Me.cboSuppCat.Name = "cboSuppCat"
        Me.cboSuppCat.Size = New System.Drawing.Size(173, 21)
        Me.cboSuppCat.TabIndex = 11
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(9, 28)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(55, 13)
        Me.Label3.TabIndex = 3
        Me.Label3.Text = "Category :"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(247, 28)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(38, 13)
        Me.Label2.TabIndex = 5
        Me.Label2.Text = "Code :"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(9, 61)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(41, 13)
        Me.Label1.TabIndex = 4
        Me.Label1.Text = "Name :"
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.Label12)
        Me.GroupBox3.Controls.Add(Me.txtQtyRec)
        Me.GroupBox3.Controls.Add(Me.txtUOM)
        Me.GroupBox3.Controls.Add(Me.Label9)
        Me.GroupBox3.Controls.Add(Me.Label11)
        Me.GroupBox3.Controls.Add(Me.txtCostPerUnit)
        Me.GroupBox3.Controls.Add(Me.txtStockQty)
        Me.GroupBox3.Controls.Add(Me.Label8)
        Me.GroupBox3.Controls.Add(Me.txtTotalCost)
        Me.GroupBox3.Controls.Add(Me.Label7)
        Me.GroupBox3.Location = New System.Drawing.Point(474, 41)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(450, 91)
        Me.GroupBox3.TabIndex = 59
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Receipt Details"
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Location = New System.Drawing.Point(310, 24)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(38, 13)
        Me.Label12.TabIndex = 42
        Me.Label12.Text = "UOM :"
        '
        'txtQtyRec
        '
        Me.txtQtyRec.Location = New System.Drawing.Point(79, 21)
        Me.txtQtyRec.Name = "txtQtyRec"
        Me.txtQtyRec.Size = New System.Drawing.Size(76, 20)
        Me.txtQtyRec.TabIndex = 38
        '
        'txtUOM
        '
        Me.txtUOM.Location = New System.Drawing.Point(354, 21)
        Me.txtUOM.Name = "txtUOM"
        Me.txtUOM.ReadOnly = True
        Me.txtUOM.Size = New System.Drawing.Size(87, 20)
        Me.txtUOM.TabIndex = 41
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(219, 58)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(88, 13)
        Me.Label9.TabIndex = 33
        Me.Label9.Text = "Total Cost (Ksh) :"
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(6, 58)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(60, 13)
        Me.Label11.TabIndex = 39
        Me.Label11.Text = "Stock Qty :"
        '
        'txtCostPerUnit
        '
        Me.txtCostPerUnit.Location = New System.Drawing.Point(222, 21)
        Me.txtCostPerUnit.Name = "txtCostPerUnit"
        Me.txtCostPerUnit.ReadOnly = True
        Me.txtCostPerUnit.Size = New System.Drawing.Size(82, 20)
        Me.txtCostPerUnit.TabIndex = 36
        '
        'txtStockQty
        '
        Me.txtStockQty.Location = New System.Drawing.Point(79, 54)
        Me.txtStockQty.Name = "txtStockQty"
        Me.txtStockQty.ReadOnly = True
        Me.txtStockQty.Size = New System.Drawing.Size(125, 20)
        Me.txtStockQty.TabIndex = 40
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(161, 24)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(56, 13)
        Me.Label8.TabIndex = 32
        Me.Label8.Text = "Unit Cost :"
        '
        'txtTotalCost
        '
        Me.txtTotalCost.Location = New System.Drawing.Point(313, 54)
        Me.txtTotalCost.Name = "txtTotalCost"
        Me.txtTotalCost.ReadOnly = True
        Me.txtTotalCost.Size = New System.Drawing.Size(128, 20)
        Me.txtTotalCost.TabIndex = 37
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(8, 24)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(67, 13)
        Me.Label7.TabIndex = 31
        Me.Label7.Text = "Recept Qty :"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.txtItemName)
        Me.GroupBox1.Controls.Add(Me.cboItemCode)
        Me.GroupBox1.Controls.Add(Me.cboItemCategory)
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Controls.Add(Me.Label5)
        Me.GroupBox1.Controls.Add(Me.Label6)
        Me.GroupBox1.Location = New System.Drawing.Point(3, 101)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(462, 91)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Item Details"
        '
        'txtItemName
        '
        Me.txtItemName.Location = New System.Drawing.Point(67, 57)
        Me.txtItemName.Name = "txtItemName"
        Me.txtItemName.ReadOnly = True
        Me.txtItemName.Size = New System.Drawing.Size(387, 20)
        Me.txtItemName.TabIndex = 30
        '
        'cboItemCode
        '
        Me.cboItemCode.DropDownHeight = 100
        Me.cboItemCode.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboItemCode.FormattingEnabled = True
        Me.cboItemCode.IntegralHeight = False
        Me.cboItemCode.Location = New System.Drawing.Point(289, 24)
        Me.cboItemCode.Name = "cboItemCode"
        Me.cboItemCode.Size = New System.Drawing.Size(165, 21)
        Me.cboItemCode.TabIndex = 29
        '
        'cboItemCategory
        '
        Me.cboItemCategory.DropDownHeight = 100
        Me.cboItemCategory.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboItemCategory.FormattingEnabled = True
        Me.cboItemCategory.IntegralHeight = False
        Me.cboItemCategory.Location = New System.Drawing.Point(67, 24)
        Me.cboItemCategory.Name = "cboItemCategory"
        Me.cboItemCategory.Size = New System.Drawing.Size(173, 21)
        Me.cboItemCategory.TabIndex = 28
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(7, 27)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(55, 13)
        Me.Label4.TabIndex = 25
        Me.Label4.Text = "Category :"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(245, 27)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(38, 13)
        Me.Label5.TabIndex = 27
        Me.Label5.Text = "Code :"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(7, 60)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(41, 13)
        Me.Label6.TabIndex = 26
        Me.Label6.Text = "Name :"
        '
        'txtRecNo
        '
        Me.txtRecNo.Location = New System.Drawing.Point(772, 9)
        Me.txtRecNo.Name = "txtRecNo"
        Me.txtRecNo.ReadOnly = True
        Me.txtRecNo.Size = New System.Drawing.Size(152, 20)
        Me.txtRecNo.TabIndex = 58
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(471, 12)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(76, 13)
        Me.Label10.TabIndex = 34
        Me.Label10.Text = "Receipt Date :"
        '
        'dtpRecDate
        '
        Me.dtpRecDate.CustomFormat = " dd - MM - yyyy"
        Me.dtpRecDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpRecDate.Location = New System.Drawing.Point(553, 9)
        Me.dtpRecDate.Name = "dtpRecDate"
        Me.dtpRecDate.Size = New System.Drawing.Size(125, 20)
        Me.dtpRecDate.TabIndex = 35
        '
        'StatusStrip11
        '
        Me.StatusStrip11.AutoSize = False
        Me.StatusStrip11.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip11.Dock = System.Windows.Forms.DockStyle.Right
        Me.StatusStrip11.Location = New System.Drawing.Point(942, 5)
        Me.StatusStrip11.Name = "StatusStrip11"
        Me.StatusStrip11.Size = New System.Drawing.Size(5, 199)
        Me.StatusStrip11.SizingGrip = False
        Me.StatusStrip11.TabIndex = 62
        Me.StatusStrip11.Text = "StatusStrip11"
        '
        'StatusStrip12
        '
        Me.StatusStrip12.AutoSize = False
        Me.StatusStrip12.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip12.Dock = System.Windows.Forms.DockStyle.Left
        Me.StatusStrip12.Location = New System.Drawing.Point(0, 5)
        Me.StatusStrip12.Name = "StatusStrip12"
        Me.StatusStrip12.Size = New System.Drawing.Size(5, 199)
        Me.StatusStrip12.SizingGrip = False
        Me.StatusStrip12.TabIndex = 61
        Me.StatusStrip12.Text = "StatusStrip12"
        '
        'StatusStrip13
        '
        Me.StatusStrip13.AutoSize = False
        Me.StatusStrip13.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip13.Location = New System.Drawing.Point(0, 204)
        Me.StatusStrip13.Name = "StatusStrip13"
        Me.StatusStrip13.Size = New System.Drawing.Size(947, 5)
        Me.StatusStrip13.SizingGrip = False
        Me.StatusStrip13.TabIndex = 60
        Me.StatusStrip13.Text = "StatusStrip13"
        '
        'StatusStrip14
        '
        Me.StatusStrip14.AutoSize = False
        Me.StatusStrip14.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip14.Dock = System.Windows.Forms.DockStyle.Top
        Me.StatusStrip14.Location = New System.Drawing.Point(0, 0)
        Me.StatusStrip14.Name = "StatusStrip14"
        Me.StatusStrip14.Size = New System.Drawing.Size(947, 5)
        Me.StatusStrip14.SizingGrip = False
        Me.StatusStrip14.TabIndex = 59
        Me.StatusStrip14.Text = "StatusStrip14"
        '
        'lstInvReceipts
        '
        Me.lstInvReceipts.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.lstInvReceipts.CheckBoxes = True
        Me.lstInvReceipts.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColumnHeader1, Me.ColumnHeader2, Me.ColumnHeader3, Me.ColumnHeader4, Me.ColumnHeader6, Me.ColumnHeader7, Me.ColumnHeader5, Me.ColumnHeader8, Me.ColumnHeader9})
        Me.lstInvReceipts.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lstInvReceipts.FullRowSelect = True
        Me.lstInvReceipts.GridLines = True
        Me.lstInvReceipts.Location = New System.Drawing.Point(0, 0)
        Me.lstInvReceipts.Name = "lstInvReceipts"
        Me.lstInvReceipts.Size = New System.Drawing.Size(947, 275)
        Me.lstInvReceipts.TabIndex = 6
        Me.lstInvReceipts.UseCompatibleStateImageBehavior = False
        Me.lstInvReceipts.View = System.Windows.Forms.View.Details
        '
        'ColumnHeader1
        '
        Me.ColumnHeader1.Text = "ReceiptNo"
        Me.ColumnHeader1.Width = 140
        '
        'ColumnHeader2
        '
        Me.ColumnHeader2.Text = "ItemCode"
        Me.ColumnHeader2.Width = 90
        '
        'ColumnHeader3
        '
        Me.ColumnHeader3.Text = "Item Name"
        Me.ColumnHeader3.Width = 200
        '
        'ColumnHeader4
        '
        Me.ColumnHeader4.Text = "VendorCode"
        Me.ColumnHeader4.Width = 110
        '
        'ColumnHeader6
        '
        Me.ColumnHeader6.Text = "UOM"
        Me.ColumnHeader6.Width = 70
        '
        'ColumnHeader7
        '
        Me.ColumnHeader7.Text = "Date Rec"
        Me.ColumnHeader7.Width = 80
        '
        'ColumnHeader5
        '
        Me.ColumnHeader5.Text = "Qty Rec"
        Me.ColumnHeader5.Width = 70
        '
        'ColumnHeader8
        '
        Me.ColumnHeader8.Text = "Unit Cost"
        Me.ColumnHeader8.Width = 80
        '
        'ColumnHeader9
        '
        Me.ColumnHeader9.Text = "Total Cost"
        Me.ColumnHeader9.Width = 80
        '
        'StatusStrip2
        '
        Me.StatusStrip2.AutoSize = False
        Me.StatusStrip2.BackColor = System.Drawing.Color.LightSeaGreen
        Me.StatusStrip2.Dock = System.Windows.Forms.DockStyle.Top
        Me.StatusStrip2.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripStatusLabel1})
        Me.StatusStrip2.Location = New System.Drawing.Point(0, 0)
        Me.StatusStrip2.Name = "StatusStrip2"
        Me.StatusStrip2.Size = New System.Drawing.Size(949, 25)
        Me.StatusStrip2.SizingGrip = False
        Me.StatusStrip2.TabIndex = 32
        Me.StatusStrip2.Text = "StatusStrip2"
        '
        'ToolStripStatusLabel1
        '
        Me.ToolStripStatusLabel1.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ToolStripStatusLabel1.Name = "ToolStripStatusLabel1"
        Me.ToolStripStatusLabel1.Size = New System.Drawing.Size(143, 20)
        Me.ToolStripStatusLabel1.Text = "INVENTORY RECEIPTS"
        '
        'StatusStrip1
        '
        Me.StatusStrip1.AutoSize = False
        Me.StatusStrip1.BackColor = System.Drawing.Color.LightSeaGreen
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 514)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(949, 25)
        Me.StatusStrip1.SizingGrip = False
        Me.StatusStrip1.TabIndex = 31
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'frmInvReceipts
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.LightSkyBlue
        Me.ClientSize = New System.Drawing.Size(974, 563)
        Me.Controls.Add(Me.pnlInvReceipts)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Name = "frmInvReceipts"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Inventory Receipts"
        Me.pnlInvReceipts.ResumeLayout(False)
        Me.SplitContainer2.Panel1.ResumeLayout(False)
        Me.SplitContainer2.Panel2.ResumeLayout(False)
        CType(Me.SplitContainer2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer2.ResumeLayout(False)
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.GroupBox4.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.StatusStrip2.ResumeLayout(False)
        Me.StatusStrip2.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents pnlInvReceipts As System.Windows.Forms.Panel
    Friend WithEvents StatusStrip2 As System.Windows.Forms.StatusStrip
    Friend WithEvents ToolStripStatusLabel1 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents SplitContainer2 As System.Windows.Forms.SplitContainer
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents txtSuppName As System.Windows.Forms.TextBox
    Friend WithEvents cboSuppCode As System.Windows.Forms.ComboBox
    Friend WithEvents cboSuppCat As System.Windows.Forms.ComboBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents txtItemName As System.Windows.Forms.TextBox
    Friend WithEvents cboItemCode As System.Windows.Forms.ComboBox
    Friend WithEvents cboItemCategory As System.Windows.Forms.ComboBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents StatusStrip11 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip12 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip13 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip14 As System.Windows.Forms.StatusStrip
    Friend WithEvents txtQtyRec As System.Windows.Forms.TextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents txtTotalCost As System.Windows.Forms.TextBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents txtCostPerUnit As System.Windows.Forms.TextBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents dtpRecDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents txtStockQty As System.Windows.Forms.TextBox
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents txtUOM As System.Windows.Forms.TextBox
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents txtRecNo As System.Windows.Forms.TextBox
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents GroupBox4 As System.Windows.Forms.GroupBox
    Friend WithEvents btnAdd As System.Windows.Forms.Button
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents btnRemove As System.Windows.Forms.Button
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents lstInvReceipts As System.Windows.Forms.ListView
    Friend WithEvents ColumnHeader1 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader2 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader3 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader4 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader5 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader6 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader7 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader8 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader9 As System.Windows.Forms.ColumnHeader
End Class
