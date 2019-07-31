<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmStudParents
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
        Me.pnlStdParents = New System.Windows.Forms.Panel()
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
        Me.btnRemove = New System.Windows.Forms.Button()
        Me.cboPRelation = New System.Windows.Forms.ComboBox()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.btnAdd = New System.Windows.Forms.Button()
        Me.cboPOccupation = New System.Windows.Forms.ComboBox()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.cboPResidence = New System.Windows.Forms.ComboBox()
        Me.txtPNationalId = New System.Windows.Forms.TextBox()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.txtPEmail = New System.Windows.Forms.TextBox()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.txtPPhoneNo = New System.Windows.Forms.TextBox()
        Me.txtPAddress = New System.Windows.Forms.TextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.txtPLName = New System.Windows.Forms.TextBox()
        Me.txtPMName = New System.Windows.Forms.TextBox()
        Me.txtPFName = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.txtStudName = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.btnLoad = New System.Windows.Forms.Button()
        Me.txtStudNo = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.StatusStrip5 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip4 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip3 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip10 = New System.Windows.Forms.StatusStrip()
        Me.SplitContainer2 = New System.Windows.Forms.SplitContainer()
        Me.lstParentDetails = New System.Windows.Forms.ListView()
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
        Me.EditToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.DeleteToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ExitToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.btnView = New System.Windows.Forms.Button()
        Me.btnUpdate = New System.Windows.Forms.Button()
        Me.btnSave = New System.Windows.Forms.Button()
        Me.btnClose = New System.Windows.Forms.Button()
        Me.StatusStrip6 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip7 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip8 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip9 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip2 = New System.Windows.Forms.StatusStrip()
        Me.ToolStripStatusLabel1 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip()
        Me.pnlStdParents.SuspendLayout()
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        CType(Me.SplitContainer2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer2.Panel1.SuspendLayout()
        Me.SplitContainer2.Panel2.SuspendLayout()
        Me.SplitContainer2.SuspendLayout()
        Me.ContextMenuStrip1.SuspendLayout()
        Me.StatusStrip2.SuspendLayout()
        Me.SuspendLayout()
        '
        'pnlStdParents
        '
        Me.pnlStdParents.BackColor = System.Drawing.SystemColors.Control
        Me.pnlStdParents.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pnlStdParents.Controls.Add(Me.SplitContainer1)
        Me.pnlStdParents.Controls.Add(Me.StatusStrip2)
        Me.pnlStdParents.Controls.Add(Me.StatusStrip1)
        Me.pnlStdParents.Location = New System.Drawing.Point(12, 12)
        Me.pnlStdParents.Name = "pnlStdParents"
        Me.pnlStdParents.Size = New System.Drawing.Size(938, 498)
        Me.pnlStdParents.TabIndex = 0
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
        Me.SplitContainer1.Panel1.Controls.Add(Me.btnRemove)
        Me.SplitContainer1.Panel1.Controls.Add(Me.cboPRelation)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label12)
        Me.SplitContainer1.Panel1.Controls.Add(Me.btnAdd)
        Me.SplitContainer1.Panel1.Controls.Add(Me.cboPOccupation)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label11)
        Me.SplitContainer1.Panel1.Controls.Add(Me.cboPResidence)
        Me.SplitContainer1.Panel1.Controls.Add(Me.txtPNationalId)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label10)
        Me.SplitContainer1.Panel1.Controls.Add(Me.txtPEmail)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label9)
        Me.SplitContainer1.Panel1.Controls.Add(Me.txtPPhoneNo)
        Me.SplitContainer1.Panel1.Controls.Add(Me.txtPAddress)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label6)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label7)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label8)
        Me.SplitContainer1.Panel1.Controls.Add(Me.txtPLName)
        Me.SplitContainer1.Panel1.Controls.Add(Me.txtPMName)
        Me.SplitContainer1.Panel1.Controls.Add(Me.txtPFName)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label5)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label3)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label2)
        Me.SplitContainer1.Panel1.Controls.Add(Me.txtStudName)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label1)
        Me.SplitContainer1.Panel1.Controls.Add(Me.btnLoad)
        Me.SplitContainer1.Panel1.Controls.Add(Me.txtStudNo)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label4)
        Me.SplitContainer1.Panel1.Controls.Add(Me.StatusStrip5)
        Me.SplitContainer1.Panel1.Controls.Add(Me.StatusStrip4)
        Me.SplitContainer1.Panel1.Controls.Add(Me.StatusStrip3)
        Me.SplitContainer1.Panel1.Controls.Add(Me.StatusStrip10)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.SplitContainer2)
        Me.SplitContainer1.Size = New System.Drawing.Size(936, 446)
        Me.SplitContainer1.SplitterDistance = 145
        Me.SplitContainer1.SplitterWidth = 1
        Me.SplitContainer1.TabIndex = 27
        '
        'btnRemove
        '
        Me.btnRemove.Location = New System.Drawing.Point(851, 105)
        Me.btnRemove.Name = "btnRemove"
        Me.btnRemove.Size = New System.Drawing.Size(75, 23)
        Me.btnRemove.TabIndex = 73
        Me.btnRemove.Text = "Remove"
        Me.btnRemove.UseVisualStyleBackColor = True
        '
        'cboPRelation
        '
        Me.cboPRelation.FormattingEnabled = True
        Me.cboPRelation.Location = New System.Drawing.Point(569, 107)
        Me.cboPRelation.Name = "cboPRelation"
        Me.cboPRelation.Size = New System.Drawing.Size(141, 21)
        Me.cboPRelation.TabIndex = 72
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Location = New System.Drawing.Point(517, 110)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(49, 13)
        Me.Label12.TabIndex = 71
        Me.Label12.Text = "Relation:"
        '
        'btnAdd
        '
        Me.btnAdd.Location = New System.Drawing.Point(735, 105)
        Me.btnAdd.Name = "btnAdd"
        Me.btnAdd.Size = New System.Drawing.Size(75, 23)
        Me.btnAdd.TabIndex = 70
        Me.btnAdd.Text = "Add"
        Me.btnAdd.UseVisualStyleBackColor = True
        '
        'cboPOccupation
        '
        Me.cboPOccupation.FormattingEnabled = True
        Me.cboPOccupation.Location = New System.Drawing.Point(735, 75)
        Me.cboPOccupation.Name = "cboPOccupation"
        Me.cboPOccupation.Size = New System.Drawing.Size(186, 21)
        Me.cboPOccupation.TabIndex = 69
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(630, 78)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(99, 13)
        Me.Label11.TabIndex = 68
        Me.Label11.Text = "Parent Occupation:"
        '
        'cboPResidence
        '
        Me.cboPResidence.FormattingEnabled = True
        Me.cboPResidence.Location = New System.Drawing.Point(79, 107)
        Me.cboPResidence.Name = "cboPResidence"
        Me.cboPResidence.Size = New System.Drawing.Size(139, 21)
        Me.cboPResidence.TabIndex = 67
        '
        'txtPNationalId
        '
        Me.txtPNationalId.Location = New System.Drawing.Point(424, 75)
        Me.txtPNationalId.Name = "txtPNationalId"
        Me.txtPNationalId.Size = New System.Drawing.Size(198, 20)
        Me.txtPNationalId.TabIndex = 66
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(313, 78)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(97, 13)
        Me.Label10.TabIndex = 65
        Me.Label10.Text = "Parent National ID:"
        '
        'txtPEmail
        '
        Me.txtPEmail.Location = New System.Drawing.Point(293, 107)
        Me.txtPEmail.Name = "txtPEmail"
        Me.txtPEmail.Size = New System.Drawing.Size(219, 20)
        Me.txtPEmail.TabIndex = 64
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(221, 110)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(69, 13)
        Me.Label9.TabIndex = 63
        Me.Label9.Text = "Parent Email:"
        '
        'txtPPhoneNo
        '
        Me.txtPPhoneNo.Location = New System.Drawing.Point(770, 44)
        Me.txtPPhoneNo.Name = "txtPPhoneNo"
        Me.txtPPhoneNo.Size = New System.Drawing.Size(151, 20)
        Me.txtPPhoneNo.TabIndex = 62
        '
        'txtPAddress
        '
        Me.txtPAddress.Location = New System.Drawing.Point(106, 75)
        Me.txtPAddress.Name = "txtPAddress"
        Me.txtPAddress.Size = New System.Drawing.Size(195, 20)
        Me.txtPAddress.TabIndex = 60
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(675, 47)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(89, 13)
        Me.Label6.TabIndex = 59
        Me.Label6.Text = "Parent PhoneNo:"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(12, 110)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(61, 13)
        Me.Label7.TabIndex = 58
        Me.Label7.Text = "Residence:"
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(12, 78)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(82, 13)
        Me.Label8.TabIndex = 57
        Me.Label8.Text = "Parent Address:"
        '
        'txtPLName
        '
        Me.txtPLName.Location = New System.Drawing.Point(506, 44)
        Me.txtPLName.Name = "txtPLName"
        Me.txtPLName.Size = New System.Drawing.Size(160, 20)
        Me.txtPLName.TabIndex = 56
        '
        'txtPMName
        '
        Me.txtPMName.Location = New System.Drawing.Point(762, 14)
        Me.txtPMName.Name = "txtPMName"
        Me.txtPMName.Size = New System.Drawing.Size(160, 20)
        Me.txtPMName.TabIndex = 55
        '
        'txtPFName
        '
        Me.txtPFName.Location = New System.Drawing.Point(506, 14)
        Me.txtPFName.Name = "txtPFName"
        Me.txtPFName.Size = New System.Drawing.Size(160, 20)
        Me.txtPFName.TabIndex = 54
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(421, 47)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(78, 13)
        Me.Label5.TabIndex = 53
        Me.Label5.Text = "Parent LName:"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(675, 17)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(81, 13)
        Me.Label3.TabIndex = 52
        Me.Label3.Text = "Parent MName:"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(422, 17)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(78, 13)
        Me.Label2.TabIndex = 51
        Me.Label2.Text = "Parent FName:"
        '
        'txtStudName
        '
        Me.txtStudName.Location = New System.Drawing.Point(106, 44)
        Me.txtStudName.Name = "txtStudName"
        Me.txtStudName.ReadOnly = True
        Me.txtStudName.Size = New System.Drawing.Size(307, 20)
        Me.txtStudName.TabIndex = 50
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 47)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(78, 13)
        Me.Label1.TabIndex = 49
        Me.Label1.Text = "Student Name:"
        '
        'btnLoad
        '
        Me.btnLoad.Location = New System.Drawing.Point(338, 12)
        Me.btnLoad.Name = "btnLoad"
        Me.btnLoad.Size = New System.Drawing.Size(75, 23)
        Me.btnLoad.TabIndex = 48
        Me.btnLoad.Text = "Load"
        Me.btnLoad.UseVisualStyleBackColor = True
        '
        'txtStudNo
        '
        Me.txtStudNo.Location = New System.Drawing.Point(106, 14)
        Me.txtStudNo.Name = "txtStudNo"
        Me.txtStudNo.Size = New System.Drawing.Size(195, 20)
        Me.txtStudNo.TabIndex = 47
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(12, 17)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(73, 13)
        Me.Label4.TabIndex = 46
        Me.Label4.Text = "Stud Adm No:"
        '
        'StatusStrip5
        '
        Me.StatusStrip5.AutoSize = False
        Me.StatusStrip5.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip5.Dock = System.Windows.Forms.DockStyle.Right
        Me.StatusStrip5.Location = New System.Drawing.Point(929, 5)
        Me.StatusStrip5.Name = "StatusStrip5"
        Me.StatusStrip5.Size = New System.Drawing.Size(5, 133)
        Me.StatusStrip5.SizingGrip = False
        Me.StatusStrip5.TabIndex = 41
        Me.StatusStrip5.Text = "StatusStrip5"
        '
        'StatusStrip4
        '
        Me.StatusStrip4.AutoSize = False
        Me.StatusStrip4.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip4.Dock = System.Windows.Forms.DockStyle.Left
        Me.StatusStrip4.Location = New System.Drawing.Point(0, 5)
        Me.StatusStrip4.Name = "StatusStrip4"
        Me.StatusStrip4.Size = New System.Drawing.Size(5, 133)
        Me.StatusStrip4.SizingGrip = False
        Me.StatusStrip4.TabIndex = 40
        Me.StatusStrip4.Text = "StatusStrip4"
        '
        'StatusStrip3
        '
        Me.StatusStrip3.AutoSize = False
        Me.StatusStrip3.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip3.Location = New System.Drawing.Point(0, 138)
        Me.StatusStrip3.Name = "StatusStrip3"
        Me.StatusStrip3.Size = New System.Drawing.Size(934, 5)
        Me.StatusStrip3.SizingGrip = False
        Me.StatusStrip3.TabIndex = 39
        Me.StatusStrip3.Text = "StatusStrip3"
        '
        'StatusStrip10
        '
        Me.StatusStrip10.AutoSize = False
        Me.StatusStrip10.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip10.Dock = System.Windows.Forms.DockStyle.Top
        Me.StatusStrip10.Location = New System.Drawing.Point(0, 0)
        Me.StatusStrip10.Name = "StatusStrip10"
        Me.StatusStrip10.Size = New System.Drawing.Size(934, 5)
        Me.StatusStrip10.SizingGrip = False
        Me.StatusStrip10.TabIndex = 38
        Me.StatusStrip10.Text = "StatusStrip10"
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
        Me.SplitContainer2.Panel1.Controls.Add(Me.lstParentDetails)
        '
        'SplitContainer2.Panel2
        '
        Me.SplitContainer2.Panel2.Controls.Add(Me.btnView)
        Me.SplitContainer2.Panel2.Controls.Add(Me.btnUpdate)
        Me.SplitContainer2.Panel2.Controls.Add(Me.btnSave)
        Me.SplitContainer2.Panel2.Controls.Add(Me.btnClose)
        Me.SplitContainer2.Panel2.Controls.Add(Me.StatusStrip6)
        Me.SplitContainer2.Panel2.Controls.Add(Me.StatusStrip7)
        Me.SplitContainer2.Panel2.Controls.Add(Me.StatusStrip8)
        Me.SplitContainer2.Panel2.Controls.Add(Me.StatusStrip9)
        Me.SplitContainer2.Size = New System.Drawing.Size(936, 300)
        Me.SplitContainer2.SplitterDistance = 252
        Me.SplitContainer2.SplitterWidth = 1
        Me.SplitContainer2.TabIndex = 28
        '
        'lstParentDetails
        '
        Me.lstParentDetails.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColumnHeader1, Me.ColumnHeader2, Me.ColumnHeader3, Me.ColumnHeader4, Me.ColumnHeader5, Me.ColumnHeader6, Me.ColumnHeader7, Me.ColumnHeader8, Me.ColumnHeader9})
        Me.lstParentDetails.ContextMenuStrip = Me.ContextMenuStrip1
        Me.lstParentDetails.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lstParentDetails.FullRowSelect = True
        Me.lstParentDetails.GridLines = True
        Me.lstParentDetails.Location = New System.Drawing.Point(0, 0)
        Me.lstParentDetails.Name = "lstParentDetails"
        Me.lstParentDetails.Size = New System.Drawing.Size(934, 250)
        Me.lstParentDetails.TabIndex = 27
        Me.lstParentDetails.UseCompatibleStateImageBehavior = False
        Me.lstParentDetails.View = System.Windows.Forms.View.Details
        '
        'ColumnHeader1
        '
        Me.ColumnHeader1.Text = "Student Adm No"
        Me.ColumnHeader1.Width = 90
        '
        'ColumnHeader2
        '
        Me.ColumnHeader2.Text = "First Name"
        Me.ColumnHeader2.Width = 110
        '
        'ColumnHeader3
        '
        Me.ColumnHeader3.Text = "Middle Name"
        Me.ColumnHeader3.Width = 110
        '
        'ColumnHeader4
        '
        Me.ColumnHeader4.Text = "Last Name"
        Me.ColumnHeader4.Width = 110
        '
        'ColumnHeader5
        '
        Me.ColumnHeader5.Text = "RelationShip"
        Me.ColumnHeader5.Width = 100
        '
        'ColumnHeader6
        '
        Me.ColumnHeader6.Text = "Residence"
        Me.ColumnHeader6.Width = 100
        '
        'ColumnHeader7
        '
        Me.ColumnHeader7.Text = "National ID"
        Me.ColumnHeader7.Width = 100
        '
        'ColumnHeader8
        '
        Me.ColumnHeader8.Text = "Occupation"
        Me.ColumnHeader8.Width = 100
        '
        'ColumnHeader9
        '
        Me.ColumnHeader9.Text = "Phone Number"
        Me.ColumnHeader9.Width = 100
        '
        'ContextMenuStrip1
        '
        Me.ContextMenuStrip1.BackColor = System.Drawing.Color.LightSeaGreen
        Me.ContextMenuStrip1.Font = New System.Drawing.Font("Century Gothic", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ContextMenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.EditToolStripMenuItem, Me.DeleteToolStripMenuItem, Me.ExitToolStripMenuItem})
        Me.ContextMenuStrip1.Name = "ContextMenuStrip1"
        Me.ContextMenuStrip1.Size = New System.Drawing.Size(153, 92)
        '
        'EditToolStripMenuItem
        '
        Me.EditToolStripMenuItem.Name = "EditToolStripMenuItem"
        Me.EditToolStripMenuItem.Size = New System.Drawing.Size(152, 22)
        Me.EditToolStripMenuItem.Text = "Edit"
        '
        'DeleteToolStripMenuItem
        '
        Me.DeleteToolStripMenuItem.Name = "DeleteToolStripMenuItem"
        Me.DeleteToolStripMenuItem.Size = New System.Drawing.Size(152, 22)
        Me.DeleteToolStripMenuItem.Text = "Delete"
        '
        'ExitToolStripMenuItem
        '
        Me.ExitToolStripMenuItem.Name = "ExitToolStripMenuItem"
        Me.ExitToolStripMenuItem.Size = New System.Drawing.Size(152, 22)
        Me.ExitToolStripMenuItem.Text = "Close"
        '
        'btnView
        '
        Me.btnView.Location = New System.Drawing.Point(500, 11)
        Me.btnView.Name = "btnView"
        Me.btnView.Size = New System.Drawing.Size(75, 23)
        Me.btnView.TabIndex = 67
        Me.btnView.Text = "View"
        Me.btnView.UseVisualStyleBackColor = True
        '
        'btnUpdate
        '
        Me.btnUpdate.Enabled = False
        Me.btnUpdate.Location = New System.Drawing.Point(733, 11)
        Me.btnUpdate.Name = "btnUpdate"
        Me.btnUpdate.Size = New System.Drawing.Size(75, 23)
        Me.btnUpdate.TabIndex = 65
        Me.btnUpdate.Text = "Update"
        Me.btnUpdate.UseVisualStyleBackColor = True
        '
        'btnSave
        '
        Me.btnSave.Location = New System.Drawing.Point(616, 11)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(75, 23)
        Me.btnSave.TabIndex = 64
        Me.btnSave.Text = "Save"
        Me.btnSave.UseVisualStyleBackColor = True
        '
        'btnClose
        '
        Me.btnClose.Location = New System.Drawing.Point(846, 11)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(75, 23)
        Me.btnClose.TabIndex = 63
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'StatusStrip6
        '
        Me.StatusStrip6.AutoSize = False
        Me.StatusStrip6.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip6.Dock = System.Windows.Forms.DockStyle.Right
        Me.StatusStrip6.Location = New System.Drawing.Point(929, 5)
        Me.StatusStrip6.Name = "StatusStrip6"
        Me.StatusStrip6.Size = New System.Drawing.Size(5, 35)
        Me.StatusStrip6.SizingGrip = False
        Me.StatusStrip6.TabIndex = 41
        Me.StatusStrip6.Text = "StatusStrip6"
        '
        'StatusStrip7
        '
        Me.StatusStrip7.AutoSize = False
        Me.StatusStrip7.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip7.Dock = System.Windows.Forms.DockStyle.Left
        Me.StatusStrip7.Location = New System.Drawing.Point(0, 5)
        Me.StatusStrip7.Name = "StatusStrip7"
        Me.StatusStrip7.Size = New System.Drawing.Size(5, 35)
        Me.StatusStrip7.SizingGrip = False
        Me.StatusStrip7.TabIndex = 40
        Me.StatusStrip7.Text = "StatusStrip7"
        '
        'StatusStrip8
        '
        Me.StatusStrip8.AutoSize = False
        Me.StatusStrip8.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip8.Location = New System.Drawing.Point(0, 40)
        Me.StatusStrip8.Name = "StatusStrip8"
        Me.StatusStrip8.Size = New System.Drawing.Size(934, 5)
        Me.StatusStrip8.SizingGrip = False
        Me.StatusStrip8.TabIndex = 39
        Me.StatusStrip8.Text = "StatusStrip8"
        '
        'StatusStrip9
        '
        Me.StatusStrip9.AutoSize = False
        Me.StatusStrip9.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip9.Dock = System.Windows.Forms.DockStyle.Top
        Me.StatusStrip9.Location = New System.Drawing.Point(0, 0)
        Me.StatusStrip9.Name = "StatusStrip9"
        Me.StatusStrip9.Size = New System.Drawing.Size(934, 5)
        Me.StatusStrip9.SizingGrip = False
        Me.StatusStrip9.TabIndex = 38
        Me.StatusStrip9.Text = "StatusStrip9"
        '
        'StatusStrip2
        '
        Me.StatusStrip2.AutoSize = False
        Me.StatusStrip2.BackColor = System.Drawing.Color.LightSeaGreen
        Me.StatusStrip2.Dock = System.Windows.Forms.DockStyle.Top
        Me.StatusStrip2.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripStatusLabel1})
        Me.StatusStrip2.Location = New System.Drawing.Point(0, 0)
        Me.StatusStrip2.Name = "StatusStrip2"
        Me.StatusStrip2.Size = New System.Drawing.Size(936, 25)
        Me.StatusStrip2.SizingGrip = False
        Me.StatusStrip2.TabIndex = 26
        Me.StatusStrip2.Text = "StatusStrip2"
        '
        'ToolStripStatusLabel1
        '
        Me.ToolStripStatusLabel1.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ToolStripStatusLabel1.Name = "ToolStripStatusLabel1"
        Me.ToolStripStatusLabel1.Size = New System.Drawing.Size(183, 20)
        Me.ToolStripStatusLabel1.Text = "STUDENT-PARENTS DETAILS"
        '
        'StatusStrip1
        '
        Me.StatusStrip1.AutoSize = False
        Me.StatusStrip1.BackColor = System.Drawing.Color.LightSeaGreen
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 471)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(936, 25)
        Me.StatusStrip1.SizingGrip = False
        Me.StatusStrip1.TabIndex = 25
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'frmStudParents
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.LightSkyBlue
        Me.ClientSize = New System.Drawing.Size(962, 522)
        Me.Controls.Add(Me.pnlStdParents)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Name = "frmStudParents"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Student Parents"
        Me.pnlStdParents.ResumeLayout(False)
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel1.PerformLayout()
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.ResumeLayout(False)
        Me.SplitContainer2.Panel1.ResumeLayout(False)
        Me.SplitContainer2.Panel2.ResumeLayout(False)
        CType(Me.SplitContainer2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer2.ResumeLayout(False)
        Me.ContextMenuStrip1.ResumeLayout(False)
        Me.StatusStrip2.ResumeLayout(False)
        Me.StatusStrip2.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents pnlStdParents As System.Windows.Forms.Panel
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents StatusStrip2 As System.Windows.Forms.StatusStrip
    Friend WithEvents ToolStripStatusLabel1 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents SplitContainer2 As System.Windows.Forms.SplitContainer
    Friend WithEvents StatusStrip5 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip4 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip3 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip10 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip6 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip7 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip8 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip9 As System.Windows.Forms.StatusStrip
    Friend WithEvents lstParentDetails As System.Windows.Forms.ListView
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents btnLoad As System.Windows.Forms.Button
    Friend WithEvents txtStudNo As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents cboPOccupation As System.Windows.Forms.ComboBox
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents cboPResidence As System.Windows.Forms.ComboBox
    Friend WithEvents txtPNationalId As System.Windows.Forms.TextBox
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents txtPEmail As System.Windows.Forms.TextBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents txtPPhoneNo As System.Windows.Forms.TextBox
    Friend WithEvents txtPAddress As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents txtPLName As System.Windows.Forms.TextBox
    Friend WithEvents txtPMName As System.Windows.Forms.TextBox
    Friend WithEvents txtPFName As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtStudName As System.Windows.Forms.TextBox
    Friend WithEvents cboPRelation As System.Windows.Forms.ComboBox
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents btnAdd As System.Windows.Forms.Button
    Friend WithEvents ColumnHeader1 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader2 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader3 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader4 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader5 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader6 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader7 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader8 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader9 As System.Windows.Forms.ColumnHeader
    Friend WithEvents btnView As System.Windows.Forms.Button
    Friend WithEvents btnUpdate As System.Windows.Forms.Button
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents btnRemove As System.Windows.Forms.Button
    Friend WithEvents ContextMenuStrip1 As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents EditToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents DeleteToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ExitToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
End Class
