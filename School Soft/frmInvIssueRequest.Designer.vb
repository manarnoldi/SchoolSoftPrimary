<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmInvIssueRequest
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
        Me.pnlIssueReq = New System.Windows.Forms.Panel()
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.GroupBox4 = New System.Windows.Forms.GroupBox()
        Me.btnAdd = New System.Windows.Forms.Button()
        Me.btnClose = New System.Windows.Forms.Button()
        Me.btnRemove = New System.Windows.Forms.Button()
        Me.btnSave = New System.Windows.Forms.Button()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.txtUOM = New System.Windows.Forms.TextBox()
        Me.Label15 = New System.Windows.Forms.Label()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.txtQtyInStock = New System.Windows.Forms.TextBox()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.txtIssueQty = New System.Windows.Forms.TextBox()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.txtTotalItemCost = New System.Windows.Forms.TextBox()
        Me.txtCostPerUnit = New System.Windows.Forms.TextBox()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.cboItemCategory = New System.Windows.Forms.ComboBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.cboItemCode = New System.Windows.Forms.ComboBox()
        Me.txtItemName = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.dtpReqDate = New System.Windows.Forms.DateTimePicker()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.cboIssueToType = New System.Windows.Forms.ComboBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.cboIssueTypeNo = New System.Windows.Forms.ComboBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.txtEmployeeName = New System.Windows.Forms.TextBox()
        Me.txtIssueNo = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.StatusStrip11 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip12 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip13 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip14 = New System.Windows.Forms.StatusStrip()
        Me.lstIssueRequest = New System.Windows.Forms.ListView()
        Me.ColumnHeader1 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader2 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader8 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader3 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader4 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader5 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader7 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader6 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader9 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader10 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.StatusStrip2 = New System.Windows.Forms.StatusStrip()
        Me.ToolStripStatusLabel1 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip()
        Me.miniToolStrip = New System.Windows.Forms.StatusStrip()
        Me.pnlIssueReq.SuspendLayout()
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.GroupBox4.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.StatusStrip2.SuspendLayout()
        Me.SuspendLayout()
        '
        'pnlIssueReq
        '
        Me.pnlIssueReq.BackColor = System.Drawing.SystemColors.Control
        Me.pnlIssueReq.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pnlIssueReq.Controls.Add(Me.SplitContainer1)
        Me.pnlIssueReq.Controls.Add(Me.StatusStrip2)
        Me.pnlIssueReq.Controls.Add(Me.StatusStrip1)
        Me.pnlIssueReq.Location = New System.Drawing.Point(12, 12)
        Me.pnlIssueReq.Name = "pnlIssueReq"
        Me.pnlIssueReq.Size = New System.Drawing.Size(951, 541)
        Me.pnlIssueReq.TabIndex = 0
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
        Me.SplitContainer1.Panel1.Controls.Add(Me.StatusStrip11)
        Me.SplitContainer1.Panel1.Controls.Add(Me.StatusStrip12)
        Me.SplitContainer1.Panel1.Controls.Add(Me.StatusStrip13)
        Me.SplitContainer1.Panel1.Controls.Add(Me.StatusStrip14)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.lstIssueRequest)
        Me.SplitContainer1.Size = New System.Drawing.Size(949, 489)
        Me.SplitContainer1.SplitterDistance = 204
        Me.SplitContainer1.SplitterWidth = 1
        Me.SplitContainer1.TabIndex = 35
        '
        'Panel1
        '
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel1.Controls.Add(Me.GroupBox4)
        Me.Panel1.Controls.Add(Me.GroupBox3)
        Me.Panel1.Controls.Add(Me.GroupBox1)
        Me.Panel1.Controls.Add(Me.dtpReqDate)
        Me.Panel1.Controls.Add(Me.Label11)
        Me.Panel1.Controls.Add(Me.GroupBox2)
        Me.Panel1.Controls.Add(Me.txtIssueNo)
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel1.Location = New System.Drawing.Point(5, 5)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(937, 192)
        Me.Panel1.TabIndex = 77
        '
        'GroupBox4
        '
        Me.GroupBox4.Controls.Add(Me.btnAdd)
        Me.GroupBox4.Controls.Add(Me.btnClose)
        Me.GroupBox4.Controls.Add(Me.btnRemove)
        Me.GroupBox4.Controls.Add(Me.btnSave)
        Me.GroupBox4.Location = New System.Drawing.Point(474, 131)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Size = New System.Drawing.Size(450, 51)
        Me.GroupBox4.TabIndex = 85
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
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.txtUOM)
        Me.GroupBox3.Controls.Add(Me.Label15)
        Me.GroupBox3.Controls.Add(Me.Label14)
        Me.GroupBox3.Controls.Add(Me.txtQtyInStock)
        Me.GroupBox3.Controls.Add(Me.Label8)
        Me.GroupBox3.Controls.Add(Me.txtIssueQty)
        Me.GroupBox3.Controls.Add(Me.Label9)
        Me.GroupBox3.Controls.Add(Me.txtTotalItemCost)
        Me.GroupBox3.Controls.Add(Me.txtCostPerUnit)
        Me.GroupBox3.Controls.Add(Me.Label10)
        Me.GroupBox3.Location = New System.Drawing.Point(474, 37)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(450, 88)
        Me.GroupBox3.TabIndex = 0
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Issue Details"
        '
        'txtUOM
        '
        Me.txtUOM.Location = New System.Drawing.Point(355, 22)
        Me.txtUOM.Name = "txtUOM"
        Me.txtUOM.ReadOnly = True
        Me.txtUOM.Size = New System.Drawing.Size(86, 20)
        Me.txtUOM.TabIndex = 78
        '
        'Label15
        '
        Me.Label15.AutoSize = True
        Me.Label15.Location = New System.Drawing.Point(311, 25)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(38, 13)
        Me.Label15.TabIndex = 77
        Me.Label15.Text = "UOM :"
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.Location = New System.Drawing.Point(9, 59)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(72, 13)
        Me.Label14.TabIndex = 75
        Me.Label14.Text = "Qty In Stock :"
        '
        'txtQtyInStock
        '
        Me.txtQtyInStock.Location = New System.Drawing.Point(92, 56)
        Me.txtQtyInStock.Name = "txtQtyInStock"
        Me.txtQtyInStock.ReadOnly = True
        Me.txtQtyInStock.Size = New System.Drawing.Size(150, 20)
        Me.txtQtyInStock.TabIndex = 76
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(9, 25)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(57, 13)
        Me.Label8.TabIndex = 69
        Me.Label8.Text = "Issue Qty :"
        '
        'txtIssueQty
        '
        Me.txtIssueQty.Location = New System.Drawing.Point(72, 22)
        Me.txtIssueQty.Name = "txtIssueQty"
        Me.txtIssueQty.Size = New System.Drawing.Size(86, 20)
        Me.txtIssueQty.TabIndex = 70
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(247, 59)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(61, 13)
        Me.Label9.TabIndex = 71
        Me.Label9.Text = "Total Cost :"
        '
        'txtTotalItemCost
        '
        Me.txtTotalItemCost.Location = New System.Drawing.Point(314, 56)
        Me.txtTotalItemCost.Name = "txtTotalItemCost"
        Me.txtTotalItemCost.ReadOnly = True
        Me.txtTotalItemCost.Size = New System.Drawing.Size(127, 20)
        Me.txtTotalItemCost.TabIndex = 72
        '
        'txtCostPerUnit
        '
        Me.txtCostPerUnit.Location = New System.Drawing.Point(229, 22)
        Me.txtCostPerUnit.Name = "txtCostPerUnit"
        Me.txtCostPerUnit.ReadOnly = True
        Me.txtCostPerUnit.Size = New System.Drawing.Size(73, 20)
        Me.txtCostPerUnit.TabIndex = 74
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(164, 25)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(56, 13)
        Me.Label10.TabIndex = 73
        Me.Label10.Text = "Unit Cost :"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.cboItemCategory)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.cboItemCode)
        Me.GroupBox1.Controls.Add(Me.txtItemName)
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Location = New System.Drawing.Point(9, 3)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(450, 85)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Item Details"
        '
        'cboItemCategory
        '
        Me.cboItemCategory.DropDownHeight = 100
        Me.cboItemCategory.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboItemCategory.FormattingEnabled = True
        Me.cboItemCategory.IntegralHeight = False
        Me.cboItemCategory.Location = New System.Drawing.Point(70, 19)
        Me.cboItemCategory.Name = "cboItemCategory"
        Me.cboItemCategory.Size = New System.Drawing.Size(172, 21)
        Me.cboItemCategory.TabIndex = 58
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(9, 22)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(55, 13)
        Me.Label2.TabIndex = 57
        Me.Label2.Text = "Category :"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(248, 22)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(38, 13)
        Me.Label3.TabIndex = 59
        Me.Label3.Text = "Code :"
        '
        'cboItemCode
        '
        Me.cboItemCode.DropDownHeight = 100
        Me.cboItemCode.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboItemCode.FormattingEnabled = True
        Me.cboItemCode.IntegralHeight = False
        Me.cboItemCode.Location = New System.Drawing.Point(305, 19)
        Me.cboItemCode.Name = "cboItemCode"
        Me.cboItemCode.Size = New System.Drawing.Size(136, 21)
        Me.cboItemCode.TabIndex = 60
        '
        'txtItemName
        '
        Me.txtItemName.Location = New System.Drawing.Point(70, 51)
        Me.txtItemName.Name = "txtItemName"
        Me.txtItemName.ReadOnly = True
        Me.txtItemName.Size = New System.Drawing.Size(371, 20)
        Me.txtItemName.TabIndex = 62
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(9, 54)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(41, 13)
        Me.Label4.TabIndex = 61
        Me.Label4.Text = "Name :"
        '
        'dtpReqDate
        '
        Me.dtpReqDate.CustomFormat = " dd - MM - yyyy"
        Me.dtpReqDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpReqDate.Location = New System.Drawing.Point(536, 10)
        Me.dtpReqDate.Name = "dtpReqDate"
        Me.dtpReqDate.Size = New System.Drawing.Size(123, 20)
        Me.dtpReqDate.TabIndex = 76
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(471, 13)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(59, 13)
        Me.Label11.TabIndex = 75
        Me.Label11.Text = "Req Date :"
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.Label5)
        Me.GroupBox2.Controls.Add(Me.cboIssueToType)
        Me.GroupBox2.Controls.Add(Me.Label6)
        Me.GroupBox2.Controls.Add(Me.cboIssueTypeNo)
        Me.GroupBox2.Controls.Add(Me.Label7)
        Me.GroupBox2.Controls.Add(Me.txtEmployeeName)
        Me.GroupBox2.Location = New System.Drawing.Point(9, 97)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(450, 85)
        Me.GroupBox2.TabIndex = 0
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Issue To Details"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(9, 22)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(53, 13)
        Me.Label5.TabIndex = 63
        Me.Label5.Text = "To Type :"
        '
        'cboIssueToType
        '
        Me.cboIssueToType.DropDownHeight = 100
        Me.cboIssueToType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboIssueToType.FormattingEnabled = True
        Me.cboIssueToType.IntegralHeight = False
        Me.cboIssueToType.Location = New System.Drawing.Point(70, 19)
        Me.cboIssueToType.Name = "cboIssueToType"
        Me.cboIssueToType.Size = New System.Drawing.Size(172, 21)
        Me.cboIssueToType.TabIndex = 64
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(248, 22)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(54, 13)
        Me.Label6.TabIndex = 65
        Me.Label6.Text = "Type No :"
        '
        'cboIssueTypeNo
        '
        Me.cboIssueTypeNo.DropDownHeight = 100
        Me.cboIssueTypeNo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboIssueTypeNo.FormattingEnabled = True
        Me.cboIssueTypeNo.IntegralHeight = False
        Me.cboIssueTypeNo.Location = New System.Drawing.Point(305, 19)
        Me.cboIssueTypeNo.Name = "cboIssueTypeNo"
        Me.cboIssueTypeNo.Size = New System.Drawing.Size(136, 21)
        Me.cboIssueTypeNo.TabIndex = 66
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(9, 54)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(75, 13)
        Me.Label7.TabIndex = 67
        Me.Label7.Text = "Issued Name :"
        '
        'txtEmployeeName
        '
        Me.txtEmployeeName.Location = New System.Drawing.Point(92, 51)
        Me.txtEmployeeName.Name = "txtEmployeeName"
        Me.txtEmployeeName.ReadOnly = True
        Me.txtEmployeeName.Size = New System.Drawing.Size(349, 20)
        Me.txtEmployeeName.TabIndex = 68
        '
        'txtIssueNo
        '
        Me.txtIssueNo.Location = New System.Drawing.Point(763, 10)
        Me.txtIssueNo.Name = "txtIssueNo"
        Me.txtIssueNo.ReadOnly = True
        Me.txtIssueNo.Size = New System.Drawing.Size(161, 20)
        Me.txtIssueNo.TabIndex = 56
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(678, 13)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(78, 13)
        Me.Label1.TabIndex = 55
        Me.Label1.Text = "IssueNote No :"
        '
        'StatusStrip11
        '
        Me.StatusStrip11.AutoSize = False
        Me.StatusStrip11.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip11.Dock = System.Windows.Forms.DockStyle.Right
        Me.StatusStrip11.Location = New System.Drawing.Point(942, 5)
        Me.StatusStrip11.Name = "StatusStrip11"
        Me.StatusStrip11.Size = New System.Drawing.Size(5, 192)
        Me.StatusStrip11.SizingGrip = False
        Me.StatusStrip11.TabIndex = 76
        Me.StatusStrip11.Text = "StatusStrip11"
        '
        'StatusStrip12
        '
        Me.StatusStrip12.AutoSize = False
        Me.StatusStrip12.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip12.Dock = System.Windows.Forms.DockStyle.Left
        Me.StatusStrip12.Location = New System.Drawing.Point(0, 5)
        Me.StatusStrip12.Name = "StatusStrip12"
        Me.StatusStrip12.Size = New System.Drawing.Size(5, 192)
        Me.StatusStrip12.SizingGrip = False
        Me.StatusStrip12.TabIndex = 75
        Me.StatusStrip12.Text = "StatusStrip12"
        '
        'StatusStrip13
        '
        Me.StatusStrip13.AutoSize = False
        Me.StatusStrip13.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip13.Location = New System.Drawing.Point(0, 197)
        Me.StatusStrip13.Name = "StatusStrip13"
        Me.StatusStrip13.Size = New System.Drawing.Size(947, 5)
        Me.StatusStrip13.SizingGrip = False
        Me.StatusStrip13.TabIndex = 74
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
        Me.StatusStrip14.TabIndex = 73
        Me.StatusStrip14.Text = "StatusStrip14"
        '
        'lstIssueRequest
        '
        Me.lstIssueRequest.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.lstIssueRequest.CheckBoxes = True
        Me.lstIssueRequest.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColumnHeader1, Me.ColumnHeader2, Me.ColumnHeader8, Me.ColumnHeader3, Me.ColumnHeader4, Me.ColumnHeader5, Me.ColumnHeader7, Me.ColumnHeader6, Me.ColumnHeader9, Me.ColumnHeader10})
        Me.lstIssueRequest.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lstIssueRequest.FullRowSelect = True
        Me.lstIssueRequest.GridLines = True
        Me.lstIssueRequest.Location = New System.Drawing.Point(0, 0)
        Me.lstIssueRequest.Name = "lstIssueRequest"
        Me.lstIssueRequest.Size = New System.Drawing.Size(947, 282)
        Me.lstIssueRequest.TabIndex = 5
        Me.lstIssueRequest.UseCompatibleStateImageBehavior = False
        Me.lstIssueRequest.View = System.Windows.Forms.View.Details
        '
        'ColumnHeader1
        '
        Me.ColumnHeader1.Text = "Issue No"
        Me.ColumnHeader1.Width = 130
        '
        'ColumnHeader2
        '
        Me.ColumnHeader2.Text = "Item Name"
        Me.ColumnHeader2.Width = 130
        '
        'ColumnHeader8
        '
        Me.ColumnHeader8.Text = "UOM"
        Me.ColumnHeader8.Width = 50
        '
        'ColumnHeader3
        '
        Me.ColumnHeader3.Text = "IssueToTypeName"
        Me.ColumnHeader3.Width = 110
        '
        'ColumnHeader4
        '
        Me.ColumnHeader4.Text = "Type No"
        '
        'ColumnHeader5
        '
        Me.ColumnHeader5.Text = "Issued To Name"
        Me.ColumnHeader5.Width = 170
        '
        'ColumnHeader7
        '
        Me.ColumnHeader7.Text = "Issue Date"
        Me.ColumnHeader7.Width = 80
        '
        'ColumnHeader6
        '
        Me.ColumnHeader6.Text = "Issue Qty"
        '
        'ColumnHeader9
        '
        Me.ColumnHeader9.Text = "Unit Cost"
        '
        'ColumnHeader10
        '
        Me.ColumnHeader10.Text = "Total Cost"
        Me.ColumnHeader10.Width = 70
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
        Me.StatusStrip2.TabIndex = 34
        Me.StatusStrip2.Text = "StatusStrip2"
        '
        'ToolStripStatusLabel1
        '
        Me.ToolStripStatusLabel1.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ToolStripStatusLabel1.Name = "ToolStripStatusLabel1"
        Me.ToolStripStatusLabel1.Size = New System.Drawing.Size(180, 20)
        Me.ToolStripStatusLabel1.Text = "INVENTORY ISSUE REQUEST"
        '
        'StatusStrip1
        '
        Me.StatusStrip1.AutoSize = False
        Me.StatusStrip1.BackColor = System.Drawing.Color.LightSeaGreen
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 514)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(949, 25)
        Me.StatusStrip1.SizingGrip = False
        Me.StatusStrip1.TabIndex = 33
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'miniToolStrip
        '
        Me.miniToolStrip.AutoSize = False
        Me.miniToolStrip.BackColor = System.Drawing.Color.MediumAquamarine
        Me.miniToolStrip.Dock = System.Windows.Forms.DockStyle.None
        Me.miniToolStrip.Location = New System.Drawing.Point(1, -7)
        Me.miniToolStrip.Name = "miniToolStrip"
        Me.miniToolStrip.Size = New System.Drawing.Size(985, 5)
        Me.miniToolStrip.SizingGrip = False
        Me.miniToolStrip.TabIndex = 63
        '
        'frmInvIssueRequest
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.LightSkyBlue
        Me.ClientSize = New System.Drawing.Size(974, 563)
        Me.Controls.Add(Me.pnlIssueReq)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Name = "frmInvIssueRequest"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Issue Request"
        Me.pnlIssueReq.ResumeLayout(False)
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.ResumeLayout(False)
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.GroupBox4.ResumeLayout(False)
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.StatusStrip2.ResumeLayout(False)
        Me.StatusStrip2.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents pnlIssueReq As System.Windows.Forms.Panel
    Friend WithEvents StatusStrip2 As System.Windows.Forms.StatusStrip
    Friend WithEvents ToolStripStatusLabel1 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents btnRemove As System.Windows.Forms.Button
    Friend WithEvents btnAdd As System.Windows.Forms.Button
    Friend WithEvents dtpReqDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents txtUOM As System.Windows.Forms.TextBox
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents txtQtyInStock As System.Windows.Forms.TextBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents txtIssueQty As System.Windows.Forms.TextBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents txtTotalItemCost As System.Windows.Forms.TextBox
    Friend WithEvents txtCostPerUnit As System.Windows.Forms.TextBox
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents cboIssueToType As System.Windows.Forms.ComboBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents cboIssueTypeNo As System.Windows.Forms.ComboBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents txtEmployeeName As System.Windows.Forms.TextBox
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents cboItemCategory As System.Windows.Forms.ComboBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents cboItemCode As System.Windows.Forms.ComboBox
    Friend WithEvents txtItemName As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtIssueNo As System.Windows.Forms.TextBox
    Friend WithEvents StatusStrip11 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip12 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip13 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip14 As System.Windows.Forms.StatusStrip
    Friend WithEvents lstIssueRequest As System.Windows.Forms.ListView
    Friend WithEvents miniToolStrip As System.Windows.Forms.StatusStrip
    Friend WithEvents GroupBox4 As System.Windows.Forms.GroupBox
    Friend WithEvents ColumnHeader1 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader2 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader3 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader4 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader5 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader6 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader7 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader8 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader9 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader10 As System.Windows.Forms.ColumnHeader
End Class
