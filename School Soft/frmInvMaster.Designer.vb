<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmInvMaster
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
        Me.pnlInvMaster = New System.Windows.Forms.Panel()
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.txtItemName = New System.Windows.Forms.TextBox()
        Me.txtItemCode = New System.Windows.Forms.TextBox()
        Me.txtSearchItemName = New System.Windows.Forms.TextBox()
        Me.txtReorderLevel = New System.Windows.Forms.TextBox()
        Me.txtCostPerUnit = New System.Windows.Forms.TextBox()
        Me.txtQtyInStock = New System.Windows.Forms.TextBox()
        Me.cbReturnFalse = New System.Windows.Forms.CheckBox()
        Me.cbReturnTrue = New System.Windows.Forms.CheckBox()
        Me.cboUOM = New System.Windows.Forms.ComboBox()
        Me.cboItemCategory = New System.Windows.Forms.ComboBox()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.StatusStrip11 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip12 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip13 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip14 = New System.Windows.Forms.StatusStrip()
        Me.SplitContainer2 = New System.Windows.Forms.SplitContainer()
        Me.lstInvMaster = New System.Windows.Forms.ListView()
        Me.ColumnHeader1 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader2 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader3 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader4 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader5 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader6 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader7 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader8 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader9 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ContextMenuStrip1 = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.UPDATEToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.DELETEToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.CLOSEToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.btnClear = New System.Windows.Forms.Button()
        Me.btnView = New System.Windows.Forms.Button()
        Me.btnSave = New System.Windows.Forms.Button()
        Me.btnUpdate = New System.Windows.Forms.Button()
        Me.btnClose = New System.Windows.Forms.Button()
        Me.StatusStrip3 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip4 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip5 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip6 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip2 = New System.Windows.Forms.StatusStrip()
        Me.ToolStripStatusLabel1 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip()
        Me.pnlInvMaster.SuspendLayout()
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.Panel2.SuspendLayout()
        CType(Me.SplitContainer2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer2.Panel1.SuspendLayout()
        Me.SplitContainer2.Panel2.SuspendLayout()
        Me.SplitContainer2.SuspendLayout()
        Me.ContextMenuStrip1.SuspendLayout()
        Me.Panel3.SuspendLayout()
        Me.StatusStrip2.SuspendLayout()
        Me.SuspendLayout()
        '
        'pnlInvMaster
        '
        Me.pnlInvMaster.BackColor = System.Drawing.SystemColors.Control
        Me.pnlInvMaster.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pnlInvMaster.Controls.Add(Me.SplitContainer1)
        Me.pnlInvMaster.Controls.Add(Me.StatusStrip2)
        Me.pnlInvMaster.Controls.Add(Me.StatusStrip1)
        Me.pnlInvMaster.Location = New System.Drawing.Point(12, 12)
        Me.pnlInvMaster.Name = "pnlInvMaster"
        Me.pnlInvMaster.Size = New System.Drawing.Size(975, 541)
        Me.pnlInvMaster.TabIndex = 0
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
        Me.SplitContainer1.Panel1.Controls.Add(Me.StatusStrip11)
        Me.SplitContainer1.Panel1.Controls.Add(Me.StatusStrip12)
        Me.SplitContainer1.Panel1.Controls.Add(Me.StatusStrip13)
        Me.SplitContainer1.Panel1.Controls.Add(Me.StatusStrip14)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.SplitContainer2)
        Me.SplitContainer1.Size = New System.Drawing.Size(973, 489)
        Me.SplitContainer1.SplitterDistance = 116
        Me.SplitContainer1.SplitterWidth = 1
        Me.SplitContainer1.TabIndex = 29
        '
        'Panel2
        '
        Me.Panel2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel2.Controls.Add(Me.txtItemName)
        Me.Panel2.Controls.Add(Me.txtItemCode)
        Me.Panel2.Controls.Add(Me.txtSearchItemName)
        Me.Panel2.Controls.Add(Me.txtReorderLevel)
        Me.Panel2.Controls.Add(Me.txtCostPerUnit)
        Me.Panel2.Controls.Add(Me.txtQtyInStock)
        Me.Panel2.Controls.Add(Me.cbReturnFalse)
        Me.Panel2.Controls.Add(Me.cbReturnTrue)
        Me.Panel2.Controls.Add(Me.cboUOM)
        Me.Panel2.Controls.Add(Me.cboItemCategory)
        Me.Panel2.Controls.Add(Me.Label9)
        Me.Panel2.Controls.Add(Me.Label8)
        Me.Panel2.Controls.Add(Me.Label7)
        Me.Panel2.Controls.Add(Me.Label6)
        Me.Panel2.Controls.Add(Me.Label5)
        Me.Panel2.Controls.Add(Me.Label4)
        Me.Panel2.Controls.Add(Me.Label3)
        Me.Panel2.Controls.Add(Me.Label2)
        Me.Panel2.Controls.Add(Me.Label1)
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel2.Location = New System.Drawing.Point(5, 5)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(961, 104)
        Me.Panel2.TabIndex = 54
        '
        'txtItemName
        '
        Me.txtItemName.Location = New System.Drawing.Point(99, 40)
        Me.txtItemName.Name = "txtItemName"
        Me.txtItemName.Size = New System.Drawing.Size(366, 20)
        Me.txtItemName.TabIndex = 18
        '
        'txtItemCode
        '
        Me.txtItemCode.Location = New System.Drawing.Point(478, 11)
        Me.txtItemCode.Name = "txtItemCode"
        Me.txtItemCode.ReadOnly = True
        Me.txtItemCode.Size = New System.Drawing.Size(232, 20)
        Me.txtItemCode.TabIndex = 17
        '
        'txtSearchItemName
        '
        Me.txtSearchItemName.Location = New System.Drawing.Point(158, 70)
        Me.txtSearchItemName.Name = "txtSearchItemName"
        Me.txtSearchItemName.Size = New System.Drawing.Size(307, 20)
        Me.txtSearchItemName.TabIndex = 16
        '
        'txtReorderLevel
        '
        Me.txtReorderLevel.Location = New System.Drawing.Point(815, 11)
        Me.txtReorderLevel.Name = "txtReorderLevel"
        Me.txtReorderLevel.Size = New System.Drawing.Size(128, 20)
        Me.txtReorderLevel.TabIndex = 15
        '
        'txtCostPerUnit
        '
        Me.txtCostPerUnit.Location = New System.Drawing.Point(815, 40)
        Me.txtCostPerUnit.Name = "txtCostPerUnit"
        Me.txtCostPerUnit.Size = New System.Drawing.Size(128, 20)
        Me.txtCostPerUnit.TabIndex = 14
        '
        'txtQtyInStock
        '
        Me.txtQtyInStock.Location = New System.Drawing.Point(815, 70)
        Me.txtQtyInStock.Name = "txtQtyInStock"
        Me.txtQtyInStock.ReadOnly = True
        Me.txtQtyInStock.Size = New System.Drawing.Size(128, 20)
        Me.txtQtyInStock.TabIndex = 13
        '
        'cbReturnFalse
        '
        Me.cbReturnFalse.AutoSize = True
        Me.cbReturnFalse.Location = New System.Drawing.Point(659, 72)
        Me.cbReturnFalse.Name = "cbReturnFalse"
        Me.cbReturnFalse.Size = New System.Drawing.Size(51, 17)
        Me.cbReturnFalse.TabIndex = 12
        Me.cbReturnFalse.Text = "False"
        Me.cbReturnFalse.UseVisualStyleBackColor = True
        '
        'cbReturnTrue
        '
        Me.cbReturnTrue.AutoSize = True
        Me.cbReturnTrue.Location = New System.Drawing.Point(571, 72)
        Me.cbReturnTrue.Name = "cbReturnTrue"
        Me.cbReturnTrue.Size = New System.Drawing.Size(48, 17)
        Me.cbReturnTrue.TabIndex = 11
        Me.cbReturnTrue.Text = "True"
        Me.cbReturnTrue.UseVisualStyleBackColor = True
        '
        'cboUOM
        '
        Me.cboUOM.FormattingEnabled = True
        Me.cboUOM.Location = New System.Drawing.Point(571, 40)
        Me.cboUOM.Name = "cboUOM"
        Me.cboUOM.Size = New System.Drawing.Size(139, 21)
        Me.cboUOM.TabIndex = 10
        '
        'cboItemCategory
        '
        Me.cboItemCategory.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboItemCategory.FormattingEnabled = True
        Me.cboItemCategory.Location = New System.Drawing.Point(99, 11)
        Me.cboItemCategory.Name = "cboItemCategory"
        Me.cboItemCategory.Size = New System.Drawing.Size(282, 21)
        Me.cboItemCategory.TabIndex = 9
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(15, 73)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(132, 13)
        Me.Label9.TabIndex = 8
        Me.Label9.Text = "Quick Search Item Name :"
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(475, 73)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(65, 13)
        Me.Label8.TabIndex = 7
        Me.Label8.Text = "Returnable :"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(728, 73)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(72, 13)
        Me.Label7.TabIndex = 6
        Me.Label7.Text = "Qty In Stock :"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(728, 43)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(75, 13)
        Me.Label6.TabIndex = 5
        Me.Label6.Text = "Cost Per Unit :"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(728, 14)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(80, 13)
        Me.Label5.TabIndex = 4
        Me.Label5.Text = "Reorder Level :"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(475, 43)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(90, 13)
        Me.Label4.TabIndex = 3
        Me.Label4.Text = "Unit Of Measure :"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(15, 14)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(78, 13)
        Me.Label3.TabIndex = 2
        Me.Label3.Text = "Item Category :"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(15, 43)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(64, 13)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "Item Name :"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(404, 14)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(61, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Item Code :"
        '
        'StatusStrip11
        '
        Me.StatusStrip11.AutoSize = False
        Me.StatusStrip11.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip11.Dock = System.Windows.Forms.DockStyle.Right
        Me.StatusStrip11.Location = New System.Drawing.Point(966, 5)
        Me.StatusStrip11.Name = "StatusStrip11"
        Me.StatusStrip11.Size = New System.Drawing.Size(5, 104)
        Me.StatusStrip11.SizingGrip = False
        Me.StatusStrip11.TabIndex = 53
        Me.StatusStrip11.Text = "StatusStrip11"
        '
        'StatusStrip12
        '
        Me.StatusStrip12.AutoSize = False
        Me.StatusStrip12.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip12.Dock = System.Windows.Forms.DockStyle.Left
        Me.StatusStrip12.Location = New System.Drawing.Point(0, 5)
        Me.StatusStrip12.Name = "StatusStrip12"
        Me.StatusStrip12.Size = New System.Drawing.Size(5, 104)
        Me.StatusStrip12.SizingGrip = False
        Me.StatusStrip12.TabIndex = 52
        Me.StatusStrip12.Text = "StatusStrip12"
        '
        'StatusStrip13
        '
        Me.StatusStrip13.AutoSize = False
        Me.StatusStrip13.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip13.Location = New System.Drawing.Point(0, 109)
        Me.StatusStrip13.Name = "StatusStrip13"
        Me.StatusStrip13.Size = New System.Drawing.Size(971, 5)
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
        Me.StatusStrip14.Size = New System.Drawing.Size(971, 5)
        Me.StatusStrip14.SizingGrip = False
        Me.StatusStrip14.TabIndex = 50
        Me.StatusStrip14.Text = "StatusStrip14"
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
        Me.SplitContainer2.Panel1.Controls.Add(Me.lstInvMaster)
        '
        'SplitContainer2.Panel2
        '
        Me.SplitContainer2.Panel2.Controls.Add(Me.Panel3)
        Me.SplitContainer2.Panel2.Controls.Add(Me.StatusStrip3)
        Me.SplitContainer2.Panel2.Controls.Add(Me.StatusStrip4)
        Me.SplitContainer2.Panel2.Controls.Add(Me.StatusStrip5)
        Me.SplitContainer2.Panel2.Controls.Add(Me.StatusStrip6)
        Me.SplitContainer2.Size = New System.Drawing.Size(973, 372)
        Me.SplitContainer2.SplitterDistance = 319
        Me.SplitContainer2.SplitterWidth = 1
        Me.SplitContainer2.TabIndex = 30
        '
        'lstInvMaster
        '
        Me.lstInvMaster.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.lstInvMaster.CheckBoxes = True
        Me.lstInvMaster.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColumnHeader1, Me.ColumnHeader2, Me.ColumnHeader3, Me.ColumnHeader4, Me.ColumnHeader5, Me.ColumnHeader6, Me.ColumnHeader7, Me.ColumnHeader8, Me.ColumnHeader9})
        Me.lstInvMaster.ContextMenuStrip = Me.ContextMenuStrip1
        Me.lstInvMaster.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lstInvMaster.FullRowSelect = True
        Me.lstInvMaster.GridLines = True
        Me.lstInvMaster.Location = New System.Drawing.Point(0, 0)
        Me.lstInvMaster.Name = "lstInvMaster"
        Me.lstInvMaster.Size = New System.Drawing.Size(971, 317)
        Me.lstInvMaster.TabIndex = 0
        Me.lstInvMaster.UseCompatibleStateImageBehavior = False
        Me.lstInvMaster.View = System.Windows.Forms.View.Details
        '
        'ColumnHeader1
        '
        Me.ColumnHeader1.Text = "Item Code"
        Me.ColumnHeader1.Width = 110
        '
        'ColumnHeader2
        '
        Me.ColumnHeader2.Text = "Item Name"
        Me.ColumnHeader2.Width = 250
        '
        'ColumnHeader3
        '
        Me.ColumnHeader3.Text = "Item Category"
        Me.ColumnHeader3.Width = 120
        '
        'ColumnHeader4
        '
        Me.ColumnHeader4.Text = "Unit Of Measure"
        Me.ColumnHeader4.Width = 90
        '
        'ColumnHeader5
        '
        Me.ColumnHeader5.Text = "Returnable"
        Me.ColumnHeader5.Width = 80
        '
        'ColumnHeader6
        '
        Me.ColumnHeader6.Text = "Reorder Level"
        Me.ColumnHeader6.Width = 80
        '
        'ColumnHeader7
        '
        Me.ColumnHeader7.Text = "Cost Per Unit"
        Me.ColumnHeader7.Width = 75
        '
        'ColumnHeader8
        '
        Me.ColumnHeader8.Text = "Qty InStock"
        Me.ColumnHeader8.Width = 70
        '
        'ColumnHeader9
        '
        Me.ColumnHeader9.Text = "Item Value"
        Me.ColumnHeader9.Width = 70
        '
        'ContextMenuStrip1
        '
        Me.ContextMenuStrip1.BackColor = System.Drawing.Color.LightSeaGreen
        Me.ContextMenuStrip1.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ContextMenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.UPDATEToolStripMenuItem, Me.DELETEToolStripMenuItem, Me.CLOSEToolStripMenuItem})
        Me.ContextMenuStrip1.Name = "ContextMenuStrip1"
        Me.ContextMenuStrip1.Size = New System.Drawing.Size(123, 70)
        '
        'UPDATEToolStripMenuItem
        '
        Me.UPDATEToolStripMenuItem.Name = "UPDATEToolStripMenuItem"
        Me.UPDATEToolStripMenuItem.Size = New System.Drawing.Size(122, 22)
        Me.UPDATEToolStripMenuItem.Text = "EDIT"
        '
        'DELETEToolStripMenuItem
        '
        Me.DELETEToolStripMenuItem.Name = "DELETEToolStripMenuItem"
        Me.DELETEToolStripMenuItem.Size = New System.Drawing.Size(122, 22)
        Me.DELETEToolStripMenuItem.Text = "DELETE"
        '
        'CLOSEToolStripMenuItem
        '
        Me.CLOSEToolStripMenuItem.Name = "CLOSEToolStripMenuItem"
        Me.CLOSEToolStripMenuItem.Size = New System.Drawing.Size(122, 22)
        Me.CLOSEToolStripMenuItem.Text = "CLOSE"
        '
        'Panel3
        '
        Me.Panel3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel3.Controls.Add(Me.btnClear)
        Me.Panel3.Controls.Add(Me.btnView)
        Me.Panel3.Controls.Add(Me.btnSave)
        Me.Panel3.Controls.Add(Me.btnUpdate)
        Me.Panel3.Controls.Add(Me.btnClose)
        Me.Panel3.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel3.Location = New System.Drawing.Point(5, 5)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(961, 40)
        Me.Panel3.TabIndex = 58
        '
        'btnClear
        '
        Me.btnClear.Location = New System.Drawing.Point(129, 8)
        Me.btnClear.Name = "btnClear"
        Me.btnClear.Size = New System.Drawing.Size(75, 23)
        Me.btnClear.TabIndex = 5
        Me.btnClear.Text = "Clear"
        Me.btnClear.UseVisualStyleBackColor = True
        '
        'btnView
        '
        Me.btnView.Location = New System.Drawing.Point(316, 8)
        Me.btnView.Name = "btnView"
        Me.btnView.Size = New System.Drawing.Size(75, 23)
        Me.btnView.TabIndex = 4
        Me.btnView.Text = "View"
        Me.btnView.UseVisualStyleBackColor = True
        '
        'btnSave
        '
        Me.btnSave.Location = New System.Drawing.Point(501, 8)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(75, 23)
        Me.btnSave.TabIndex = 3
        Me.btnSave.Text = "Save"
        Me.btnSave.UseVisualStyleBackColor = True
        '
        'btnUpdate
        '
        Me.btnUpdate.Enabled = False
        Me.btnUpdate.Location = New System.Drawing.Point(682, 8)
        Me.btnUpdate.Name = "btnUpdate"
        Me.btnUpdate.Size = New System.Drawing.Size(75, 23)
        Me.btnUpdate.TabIndex = 2
        Me.btnUpdate.Text = "Update"
        Me.btnUpdate.UseVisualStyleBackColor = True
        '
        'btnClose
        '
        Me.btnClose.Location = New System.Drawing.Point(868, 8)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(75, 23)
        Me.btnClose.TabIndex = 0
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'StatusStrip3
        '
        Me.StatusStrip3.AutoSize = False
        Me.StatusStrip3.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip3.Dock = System.Windows.Forms.DockStyle.Right
        Me.StatusStrip3.Location = New System.Drawing.Point(966, 5)
        Me.StatusStrip3.Name = "StatusStrip3"
        Me.StatusStrip3.Size = New System.Drawing.Size(5, 40)
        Me.StatusStrip3.SizingGrip = False
        Me.StatusStrip3.TabIndex = 57
        Me.StatusStrip3.Text = "StatusStrip3"
        '
        'StatusStrip4
        '
        Me.StatusStrip4.AutoSize = False
        Me.StatusStrip4.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip4.Dock = System.Windows.Forms.DockStyle.Left
        Me.StatusStrip4.Location = New System.Drawing.Point(0, 5)
        Me.StatusStrip4.Name = "StatusStrip4"
        Me.StatusStrip4.Size = New System.Drawing.Size(5, 40)
        Me.StatusStrip4.SizingGrip = False
        Me.StatusStrip4.TabIndex = 56
        Me.StatusStrip4.Text = "StatusStrip4"
        '
        'StatusStrip5
        '
        Me.StatusStrip5.AutoSize = False
        Me.StatusStrip5.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip5.Location = New System.Drawing.Point(0, 45)
        Me.StatusStrip5.Name = "StatusStrip5"
        Me.StatusStrip5.Size = New System.Drawing.Size(971, 5)
        Me.StatusStrip5.SizingGrip = False
        Me.StatusStrip5.TabIndex = 55
        Me.StatusStrip5.Text = "StatusStrip5"
        '
        'StatusStrip6
        '
        Me.StatusStrip6.AutoSize = False
        Me.StatusStrip6.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip6.Dock = System.Windows.Forms.DockStyle.Top
        Me.StatusStrip6.Location = New System.Drawing.Point(0, 0)
        Me.StatusStrip6.Name = "StatusStrip6"
        Me.StatusStrip6.Size = New System.Drawing.Size(971, 5)
        Me.StatusStrip6.SizingGrip = False
        Me.StatusStrip6.TabIndex = 54
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
        Me.StatusStrip2.Size = New System.Drawing.Size(973, 25)
        Me.StatusStrip2.SizingGrip = False
        Me.StatusStrip2.TabIndex = 28
        Me.StatusStrip2.Text = "StatusStrip2"
        '
        'ToolStripStatusLabel1
        '
        Me.ToolStripStatusLabel1.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ToolStripStatusLabel1.Name = "ToolStripStatusLabel1"
        Me.ToolStripStatusLabel1.Size = New System.Drawing.Size(172, 20)
        Me.ToolStripStatusLabel1.Text = "INVENTORY ITEM MASTER"
        '
        'StatusStrip1
        '
        Me.StatusStrip1.AutoSize = False
        Me.StatusStrip1.BackColor = System.Drawing.Color.LightSeaGreen
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 514)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(973, 25)
        Me.StatusStrip1.SizingGrip = False
        Me.StatusStrip1.TabIndex = 27
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'frmInvMaster
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.LightSkyBlue
        Me.ClientSize = New System.Drawing.Size(999, 565)
        Me.Controls.Add(Me.pnlInvMaster)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Name = "frmInvMaster"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Inventory Master"
        Me.pnlInvMaster.ResumeLayout(False)
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
        Me.ContextMenuStrip1.ResumeLayout(False)
        Me.Panel3.ResumeLayout(False)
        Me.StatusStrip2.ResumeLayout(False)
        Me.StatusStrip2.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents pnlInvMaster As System.Windows.Forms.Panel
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents SplitContainer2 As System.Windows.Forms.SplitContainer
    Friend WithEvents StatusStrip2 As System.Windows.Forms.StatusStrip
    Friend WithEvents ToolStripStatusLabel1 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip11 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip12 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip13 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip14 As System.Windows.Forms.StatusStrip
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents txtItemName As System.Windows.Forms.TextBox
    Friend WithEvents txtItemCode As System.Windows.Forms.TextBox
    Friend WithEvents txtSearchItemName As System.Windows.Forms.TextBox
    Friend WithEvents txtReorderLevel As System.Windows.Forms.TextBox
    Friend WithEvents txtCostPerUnit As System.Windows.Forms.TextBox
    Friend WithEvents txtQtyInStock As System.Windows.Forms.TextBox
    Friend WithEvents cbReturnFalse As System.Windows.Forms.CheckBox
    Friend WithEvents cbReturnTrue As System.Windows.Forms.CheckBox
    Friend WithEvents cboUOM As System.Windows.Forms.ComboBox
    Friend WithEvents cboItemCategory As System.Windows.Forms.ComboBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents lstInvMaster As System.Windows.Forms.ListView
    Friend WithEvents ColumnHeader1 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader2 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader3 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader4 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader5 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader6 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader7 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader8 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader9 As System.Windows.Forms.ColumnHeader
    Friend WithEvents Panel3 As System.Windows.Forms.Panel
    Friend WithEvents btnView As System.Windows.Forms.Button
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents btnUpdate As System.Windows.Forms.Button
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents StatusStrip3 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip4 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip5 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip6 As System.Windows.Forms.StatusStrip
    Friend WithEvents ContextMenuStrip1 As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents UPDATEToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents DELETEToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents CLOSEToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents btnClear As System.Windows.Forms.Button
End Class
