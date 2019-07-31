<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmInvVendorMaster
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
        Me.pnlVendorMaster = New System.Windows.Forms.Panel()
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.txtVendorAddress = New System.Windows.Forms.TextBox()
        Me.txtVendorEmail = New System.Windows.Forms.TextBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.txtVendorContacts = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.txtVendorName = New System.Windows.Forms.TextBox()
        Me.txtVendorCode = New System.Windows.Forms.TextBox()
        Me.txtSearchVendorName = New System.Windows.Forms.TextBox()
        Me.cboVendorLocation = New System.Windows.Forms.ComboBox()
        Me.cboVendorCategory = New System.Windows.Forms.ComboBox()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.StatusStrip11 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip12 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip13 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip14 = New System.Windows.Forms.StatusStrip()
        Me.SplitContainer2 = New System.Windows.Forms.SplitContainer()
        Me.lstVendorMaster = New System.Windows.Forms.ListView()
        Me.ColumnHeader1 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader2 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader3 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader4 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader5 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader6 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader7 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader8 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ContextMenuStrip1 = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.UPDATEToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.DELETEToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.CLOSEToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.dtpRegDate = New System.Windows.Forms.DateTimePicker()
        Me.btnView = New System.Windows.Forms.Button()
        Me.btnSave = New System.Windows.Forms.Button()
        Me.btnUpdate = New System.Windows.Forms.Button()
        Me.btnClose = New System.Windows.Forms.Button()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.StatusStrip3 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip4 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip5 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip6 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip2 = New System.Windows.Forms.StatusStrip()
        Me.ToolStripStatusLabel1 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip()
        Me.pnlVendorMaster.SuspendLayout()
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
        'pnlVendorMaster
        '
        Me.pnlVendorMaster.BackColor = System.Drawing.SystemColors.Control
        Me.pnlVendorMaster.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pnlVendorMaster.Controls.Add(Me.SplitContainer1)
        Me.pnlVendorMaster.Controls.Add(Me.StatusStrip2)
        Me.pnlVendorMaster.Controls.Add(Me.StatusStrip1)
        Me.pnlVendorMaster.Location = New System.Drawing.Point(12, 12)
        Me.pnlVendorMaster.Name = "pnlVendorMaster"
        Me.pnlVendorMaster.Size = New System.Drawing.Size(957, 535)
        Me.pnlVendorMaster.TabIndex = 0
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
        Me.SplitContainer1.Size = New System.Drawing.Size(955, 483)
        Me.SplitContainer1.SplitterDistance = 113
        Me.SplitContainer1.SplitterWidth = 1
        Me.SplitContainer1.TabIndex = 31
        '
        'Panel2
        '
        Me.Panel2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel2.Controls.Add(Me.txtVendorAddress)
        Me.Panel2.Controls.Add(Me.txtVendorEmail)
        Me.Panel2.Controls.Add(Me.Label7)
        Me.Panel2.Controls.Add(Me.txtVendorContacts)
        Me.Panel2.Controls.Add(Me.Label2)
        Me.Panel2.Controls.Add(Me.Label5)
        Me.Panel2.Controls.Add(Me.Label6)
        Me.Panel2.Controls.Add(Me.txtVendorName)
        Me.Panel2.Controls.Add(Me.txtVendorCode)
        Me.Panel2.Controls.Add(Me.txtSearchVendorName)
        Me.Panel2.Controls.Add(Me.cboVendorLocation)
        Me.Panel2.Controls.Add(Me.cboVendorCategory)
        Me.Panel2.Controls.Add(Me.Label9)
        Me.Panel2.Controls.Add(Me.Label4)
        Me.Panel2.Controls.Add(Me.Label3)
        Me.Panel2.Controls.Add(Me.Label1)
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel2.Location = New System.Drawing.Point(5, 5)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(943, 101)
        Me.Panel2.TabIndex = 59
        '
        'txtVendorAddress
        '
        Me.txtVendorAddress.Location = New System.Drawing.Point(738, 10)
        Me.txtVendorAddress.Name = "txtVendorAddress"
        Me.txtVendorAddress.Size = New System.Drawing.Size(192, 20)
        Me.txtVendorAddress.TabIndex = 26
        '
        'txtVendorEmail
        '
        Me.txtVendorEmail.Location = New System.Drawing.Point(391, 70)
        Me.txtVendorEmail.Name = "txtVendorEmail"
        Me.txtVendorEmail.Size = New System.Drawing.Size(232, 20)
        Me.txtVendorEmail.TabIndex = 25
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(293, 73)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(75, 13)
        Me.Label7.TabIndex = 24
        Me.Label7.Text = "Vendor Email :"
        '
        'txtVendorContacts
        '
        Me.txtVendorContacts.Location = New System.Drawing.Point(739, 70)
        Me.txtVendorContacts.Name = "txtVendorContacts"
        Me.txtVendorContacts.Size = New System.Drawing.Size(192, 20)
        Me.txtVendorContacts.TabIndex = 23
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(641, 13)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(88, 13)
        Me.Label2.TabIndex = 20
        Me.Label2.Text = "Vendor Address :"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(15, 43)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(78, 13)
        Me.Label5.TabIndex = 19
        Me.Label5.Text = "Vendor Name :"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(641, 73)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(92, 13)
        Me.Label6.TabIndex = 22
        Me.Label6.Text = "Vendor Contacts :"
        '
        'txtVendorName
        '
        Me.txtVendorName.Location = New System.Drawing.Point(109, 40)
        Me.txtVendorName.Name = "txtVendorName"
        Me.txtVendorName.Size = New System.Drawing.Size(514, 20)
        Me.txtVendorName.TabIndex = 18
        '
        'txtVendorCode
        '
        Me.txtVendorCode.Location = New System.Drawing.Point(109, 10)
        Me.txtVendorCode.Name = "txtVendorCode"
        Me.txtVendorCode.ReadOnly = True
        Me.txtVendorCode.Size = New System.Drawing.Size(169, 20)
        Me.txtVendorCode.TabIndex = 17
        '
        'txtSearchVendorName
        '
        Me.txtSearchVendorName.Location = New System.Drawing.Point(109, 70)
        Me.txtSearchVendorName.Name = "txtSearchVendorName"
        Me.txtSearchVendorName.Size = New System.Drawing.Size(169, 20)
        Me.txtSearchVendorName.TabIndex = 16
        '
        'cboVendorLocation
        '
        Me.cboVendorLocation.FormattingEnabled = True
        Me.cboVendorLocation.Location = New System.Drawing.Point(738, 40)
        Me.cboVendorLocation.Name = "cboVendorLocation"
        Me.cboVendorLocation.Size = New System.Drawing.Size(193, 21)
        Me.cboVendorLocation.TabIndex = 10
        '
        'cboVendorCategory
        '
        Me.cboVendorCategory.FormattingEnabled = True
        Me.cboVendorCategory.Location = New System.Drawing.Point(391, 10)
        Me.cboVendorCategory.Name = "cboVendorCategory"
        Me.cboVendorCategory.Size = New System.Drawing.Size(232, 21)
        Me.cboVendorCategory.TabIndex = 9
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(15, 73)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(84, 13)
        Me.Label9.TabIndex = 8
        Me.Label9.Text = "Search Vendor :"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(641, 43)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(91, 13)
        Me.Label4.TabIndex = 3
        Me.Label4.Text = "Vendor Location :"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(293, 13)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(92, 13)
        Me.Label3.TabIndex = 2
        Me.Label3.Text = "Vendor Category :"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(15, 13)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(75, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Vendor Code :"
        '
        'StatusStrip11
        '
        Me.StatusStrip11.AutoSize = False
        Me.StatusStrip11.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip11.Dock = System.Windows.Forms.DockStyle.Right
        Me.StatusStrip11.Location = New System.Drawing.Point(948, 5)
        Me.StatusStrip11.Name = "StatusStrip11"
        Me.StatusStrip11.Size = New System.Drawing.Size(5, 101)
        Me.StatusStrip11.SizingGrip = False
        Me.StatusStrip11.TabIndex = 58
        Me.StatusStrip11.Text = "StatusStrip11"
        '
        'StatusStrip12
        '
        Me.StatusStrip12.AutoSize = False
        Me.StatusStrip12.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip12.Dock = System.Windows.Forms.DockStyle.Left
        Me.StatusStrip12.Location = New System.Drawing.Point(0, 5)
        Me.StatusStrip12.Name = "StatusStrip12"
        Me.StatusStrip12.Size = New System.Drawing.Size(5, 101)
        Me.StatusStrip12.SizingGrip = False
        Me.StatusStrip12.TabIndex = 57
        Me.StatusStrip12.Text = "StatusStrip12"
        '
        'StatusStrip13
        '
        Me.StatusStrip13.AutoSize = False
        Me.StatusStrip13.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip13.Location = New System.Drawing.Point(0, 106)
        Me.StatusStrip13.Name = "StatusStrip13"
        Me.StatusStrip13.Size = New System.Drawing.Size(953, 5)
        Me.StatusStrip13.SizingGrip = False
        Me.StatusStrip13.TabIndex = 56
        Me.StatusStrip13.Text = "StatusStrip13"
        '
        'StatusStrip14
        '
        Me.StatusStrip14.AutoSize = False
        Me.StatusStrip14.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip14.Dock = System.Windows.Forms.DockStyle.Top
        Me.StatusStrip14.Location = New System.Drawing.Point(0, 0)
        Me.StatusStrip14.Name = "StatusStrip14"
        Me.StatusStrip14.Size = New System.Drawing.Size(953, 5)
        Me.StatusStrip14.SizingGrip = False
        Me.StatusStrip14.TabIndex = 55
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
        Me.SplitContainer2.Panel1.Controls.Add(Me.lstVendorMaster)
        '
        'SplitContainer2.Panel2
        '
        Me.SplitContainer2.Panel2.Controls.Add(Me.Panel3)
        Me.SplitContainer2.Panel2.Controls.Add(Me.StatusStrip3)
        Me.SplitContainer2.Panel2.Controls.Add(Me.StatusStrip4)
        Me.SplitContainer2.Panel2.Controls.Add(Me.StatusStrip5)
        Me.SplitContainer2.Panel2.Controls.Add(Me.StatusStrip6)
        Me.SplitContainer2.Size = New System.Drawing.Size(955, 369)
        Me.SplitContainer2.SplitterDistance = 312
        Me.SplitContainer2.SplitterWidth = 1
        Me.SplitContainer2.TabIndex = 32
        '
        'lstVendorMaster
        '
        Me.lstVendorMaster.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.lstVendorMaster.CheckBoxes = True
        Me.lstVendorMaster.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColumnHeader1, Me.ColumnHeader2, Me.ColumnHeader3, Me.ColumnHeader4, Me.ColumnHeader5, Me.ColumnHeader6, Me.ColumnHeader7, Me.ColumnHeader8})
        Me.lstVendorMaster.ContextMenuStrip = Me.ContextMenuStrip1
        Me.lstVendorMaster.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lstVendorMaster.FullRowSelect = True
        Me.lstVendorMaster.GridLines = True
        Me.lstVendorMaster.Location = New System.Drawing.Point(0, 0)
        Me.lstVendorMaster.Name = "lstVendorMaster"
        Me.lstVendorMaster.Size = New System.Drawing.Size(953, 310)
        Me.lstVendorMaster.TabIndex = 1
        Me.lstVendorMaster.UseCompatibleStateImageBehavior = False
        Me.lstVendorMaster.View = System.Windows.Forms.View.Details
        '
        'ColumnHeader1
        '
        Me.ColumnHeader1.Text = "Vendor Code"
        Me.ColumnHeader1.Width = 100
        '
        'ColumnHeader2
        '
        Me.ColumnHeader2.Text = "Vendor Name"
        Me.ColumnHeader2.Width = 200
        '
        'ColumnHeader3
        '
        Me.ColumnHeader3.Text = "Vendor Category"
        Me.ColumnHeader3.Width = 100
        '
        'ColumnHeader4
        '
        Me.ColumnHeader4.Text = "Vendor Address"
        Me.ColumnHeader4.Width = 90
        '
        'ColumnHeader5
        '
        Me.ColumnHeader5.Text = "Vendor Location"
        Me.ColumnHeader5.Width = 90
        '
        'ColumnHeader6
        '
        Me.ColumnHeader6.Text = "Vendor Contacts"
        Me.ColumnHeader6.Width = 95
        '
        'ColumnHeader7
        '
        Me.ColumnHeader7.Text = "Vendor Email Address"
        Me.ColumnHeader7.Width = 180
        '
        'ColumnHeader8
        '
        Me.ColumnHeader8.Text = "Date Registered"
        Me.ColumnHeader8.Width = 90
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
        Me.Panel3.Controls.Add(Me.dtpRegDate)
        Me.Panel3.Controls.Add(Me.btnView)
        Me.Panel3.Controls.Add(Me.btnSave)
        Me.Panel3.Controls.Add(Me.btnUpdate)
        Me.Panel3.Controls.Add(Me.btnClose)
        Me.Panel3.Controls.Add(Me.Label8)
        Me.Panel3.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel3.Location = New System.Drawing.Point(5, 5)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(943, 44)
        Me.Panel3.TabIndex = 63
        '
        'dtpRegDate
        '
        Me.dtpRegDate.CustomFormat = " dd - MM - yyyy"
        Me.dtpRegDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpRegDate.Location = New System.Drawing.Point(111, 8)
        Me.dtpRegDate.Name = "dtpRegDate"
        Me.dtpRegDate.Size = New System.Drawing.Size(129, 20)
        Me.dtpRegDate.TabIndex = 27
        '
        'btnView
        '
        Me.btnView.Location = New System.Drawing.Point(391, 8)
        Me.btnView.Name = "btnView"
        Me.btnView.Size = New System.Drawing.Size(75, 23)
        Me.btnView.TabIndex = 4
        Me.btnView.Text = "View"
        Me.btnView.UseVisualStyleBackColor = True
        '
        'btnSave
        '
        Me.btnSave.Location = New System.Drawing.Point(548, 8)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(75, 23)
        Me.btnSave.TabIndex = 3
        Me.btnSave.Text = "Save"
        Me.btnSave.UseVisualStyleBackColor = True
        '
        'btnUpdate
        '
        Me.btnUpdate.Enabled = False
        Me.btnUpdate.Location = New System.Drawing.Point(697, 8)
        Me.btnUpdate.Name = "btnUpdate"
        Me.btnUpdate.Size = New System.Drawing.Size(75, 23)
        Me.btnUpdate.TabIndex = 2
        Me.btnUpdate.Text = "Update"
        Me.btnUpdate.UseVisualStyleBackColor = True
        '
        'btnClose
        '
        Me.btnClose.Location = New System.Drawing.Point(856, 8)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(75, 23)
        Me.btnClose.TabIndex = 0
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(15, 11)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(90, 13)
        Me.Label8.TabIndex = 7
        Me.Label8.Text = "Date Registered :"
        '
        'StatusStrip3
        '
        Me.StatusStrip3.AutoSize = False
        Me.StatusStrip3.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip3.Dock = System.Windows.Forms.DockStyle.Right
        Me.StatusStrip3.Location = New System.Drawing.Point(948, 5)
        Me.StatusStrip3.Name = "StatusStrip3"
        Me.StatusStrip3.Size = New System.Drawing.Size(5, 44)
        Me.StatusStrip3.SizingGrip = False
        Me.StatusStrip3.TabIndex = 62
        Me.StatusStrip3.Text = "StatusStrip3"
        '
        'StatusStrip4
        '
        Me.StatusStrip4.AutoSize = False
        Me.StatusStrip4.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip4.Dock = System.Windows.Forms.DockStyle.Left
        Me.StatusStrip4.Location = New System.Drawing.Point(0, 5)
        Me.StatusStrip4.Name = "StatusStrip4"
        Me.StatusStrip4.Size = New System.Drawing.Size(5, 44)
        Me.StatusStrip4.SizingGrip = False
        Me.StatusStrip4.TabIndex = 61
        Me.StatusStrip4.Text = "StatusStrip4"
        '
        'StatusStrip5
        '
        Me.StatusStrip5.AutoSize = False
        Me.StatusStrip5.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip5.Location = New System.Drawing.Point(0, 49)
        Me.StatusStrip5.Name = "StatusStrip5"
        Me.StatusStrip5.Size = New System.Drawing.Size(953, 5)
        Me.StatusStrip5.SizingGrip = False
        Me.StatusStrip5.TabIndex = 60
        Me.StatusStrip5.Text = "StatusStrip5"
        '
        'StatusStrip6
        '
        Me.StatusStrip6.AutoSize = False
        Me.StatusStrip6.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip6.Dock = System.Windows.Forms.DockStyle.Top
        Me.StatusStrip6.Location = New System.Drawing.Point(0, 0)
        Me.StatusStrip6.Name = "StatusStrip6"
        Me.StatusStrip6.Size = New System.Drawing.Size(953, 5)
        Me.StatusStrip6.SizingGrip = False
        Me.StatusStrip6.TabIndex = 59
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
        Me.StatusStrip2.Size = New System.Drawing.Size(955, 25)
        Me.StatusStrip2.SizingGrip = False
        Me.StatusStrip2.TabIndex = 30
        Me.StatusStrip2.Text = "StatusStrip2"
        '
        'ToolStripStatusLabel1
        '
        Me.ToolStripStatusLabel1.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ToolStripStatusLabel1.Name = "ToolStripStatusLabel1"
        Me.ToolStripStatusLabel1.Size = New System.Drawing.Size(195, 20)
        Me.ToolStripStatusLabel1.Text = "INVENTORY VENDOR MASTER"
        '
        'StatusStrip1
        '
        Me.StatusStrip1.AutoSize = False
        Me.StatusStrip1.BackColor = System.Drawing.Color.LightSeaGreen
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 508)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(955, 25)
        Me.StatusStrip1.SizingGrip = False
        Me.StatusStrip1.TabIndex = 29
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'frmInvVendorMaster
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.LightSkyBlue
        Me.ClientSize = New System.Drawing.Size(981, 559)
        Me.Controls.Add(Me.pnlVendorMaster)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Name = "frmInvVendorMaster"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Vendor Master"
        Me.pnlVendorMaster.ResumeLayout(False)
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
        Me.Panel3.PerformLayout()
        Me.StatusStrip2.ResumeLayout(False)
        Me.StatusStrip2.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents pnlVendorMaster As System.Windows.Forms.Panel
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents SplitContainer2 As System.Windows.Forms.SplitContainer
    Friend WithEvents StatusStrip2 As System.Windows.Forms.StatusStrip
    Friend WithEvents ToolStripStatusLabel1 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents lstVendorMaster As System.Windows.Forms.ListView
    Friend WithEvents ColumnHeader1 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader2 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader3 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader4 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader5 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader6 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader7 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader8 As System.Windows.Forms.ColumnHeader
    Friend WithEvents Panel3 As System.Windows.Forms.Panel
    Friend WithEvents btnView As System.Windows.Forms.Button
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents btnUpdate As System.Windows.Forms.Button
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents StatusStrip3 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip4 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip5 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip6 As System.Windows.Forms.StatusStrip
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents txtVendorName As System.Windows.Forms.TextBox
    Friend WithEvents txtVendorCode As System.Windows.Forms.TextBox
    Friend WithEvents txtSearchVendorName As System.Windows.Forms.TextBox
    Friend WithEvents cboVendorLocation As System.Windows.Forms.ComboBox
    Friend WithEvents cboVendorCategory As System.Windows.Forms.ComboBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents StatusStrip11 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip12 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip13 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip14 As System.Windows.Forms.StatusStrip
    Friend WithEvents txtVendorEmail As System.Windows.Forms.TextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents txtVendorContacts As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents txtVendorAddress As System.Windows.Forms.TextBox
    Friend WithEvents dtpRegDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents ContextMenuStrip1 As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents UPDATEToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents DELETEToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents CLOSEToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
End Class
