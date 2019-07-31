<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmSchDepartments
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
        Me.pnlSchDepartment = New System.Windows.Forms.Panel()
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
        Me.cboStaffType = New System.Windows.Forms.ComboBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.txtStaffName = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.cboStaffNo = New System.Windows.Forms.ComboBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.TextBox3 = New System.Windows.Forms.TextBox()
        Me.txtDepCode = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.txtDepName = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.lstDepartMents = New System.Windows.Forms.ListView()
        Me.ColumnHeader1 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader2 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader3 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.StatusStrip5 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip4 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip3 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip10 = New System.Windows.Forms.StatusStrip()
        Me.btnView = New System.Windows.Forms.Button()
        Me.btnDelete = New System.Windows.Forms.Button()
        Me.btnUpdate = New System.Windows.Forms.Button()
        Me.btnSave = New System.Windows.Forms.Button()
        Me.btnClose = New System.Windows.Forms.Button()
        Me.StatusStrip6 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip7 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip8 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip9 = New System.Windows.Forms.StatusStrip()
        Me.ToolStripStatusLabel2 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.StatusStrip2 = New System.Windows.Forms.StatusStrip()
        Me.ToolStripStatusLabel1 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip()
        Me.pnlSchDepartment.SuspendLayout()
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.StatusStrip9.SuspendLayout()
        Me.StatusStrip2.SuspendLayout()
        Me.SuspendLayout()
        '
        'pnlSchDepartment
        '
        Me.pnlSchDepartment.BackColor = System.Drawing.SystemColors.Control
        Me.pnlSchDepartment.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pnlSchDepartment.Controls.Add(Me.SplitContainer1)
        Me.pnlSchDepartment.Controls.Add(Me.StatusStrip2)
        Me.pnlSchDepartment.Controls.Add(Me.StatusStrip1)
        Me.pnlSchDepartment.Location = New System.Drawing.Point(12, 12)
        Me.pnlSchDepartment.Name = "pnlSchDepartment"
        Me.pnlSchDepartment.Size = New System.Drawing.Size(847, 409)
        Me.pnlSchDepartment.TabIndex = 0
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
        Me.SplitContainer1.Panel1.Controls.Add(Me.cboStaffType)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label5)
        Me.SplitContainer1.Panel1.Controls.Add(Me.txtStaffName)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label4)
        Me.SplitContainer1.Panel1.Controls.Add(Me.cboStaffNo)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label3)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TextBox3)
        Me.SplitContainer1.Panel1.Controls.Add(Me.txtDepCode)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label2)
        Me.SplitContainer1.Panel1.Controls.Add(Me.txtDepName)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label1)
        Me.SplitContainer1.Panel1.Controls.Add(Me.lstDepartMents)
        Me.SplitContainer1.Panel1.Controls.Add(Me.StatusStrip5)
        Me.SplitContainer1.Panel1.Controls.Add(Me.StatusStrip4)
        Me.SplitContainer1.Panel1.Controls.Add(Me.StatusStrip3)
        Me.SplitContainer1.Panel1.Controls.Add(Me.StatusStrip10)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.btnView)
        Me.SplitContainer1.Panel2.Controls.Add(Me.btnDelete)
        Me.SplitContainer1.Panel2.Controls.Add(Me.btnUpdate)
        Me.SplitContainer1.Panel2.Controls.Add(Me.btnSave)
        Me.SplitContainer1.Panel2.Controls.Add(Me.btnClose)
        Me.SplitContainer1.Panel2.Controls.Add(Me.StatusStrip6)
        Me.SplitContainer1.Panel2.Controls.Add(Me.StatusStrip7)
        Me.SplitContainer1.Panel2.Controls.Add(Me.StatusStrip8)
        Me.SplitContainer1.Panel2.Controls.Add(Me.StatusStrip9)
        Me.SplitContainer1.Size = New System.Drawing.Size(845, 357)
        Me.SplitContainer1.SplitterDistance = 287
        Me.SplitContainer1.TabIndex = 18
        '
        'cboStaffType
        '
        Me.cboStaffType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboStaffType.FormattingEnabled = True
        Me.cboStaffType.Location = New System.Drawing.Point(23, 196)
        Me.cboStaffType.Name = "cboStaffType"
        Me.cboStaffType.Size = New System.Drawing.Size(179, 21)
        Me.cboStaffType.TabIndex = 33
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(20, 180)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(56, 13)
        Me.Label5.TabIndex = 32
        Me.Label5.Text = "Staff Type"
        '
        'txtStaffName
        '
        Me.txtStaffName.Anchor = CType((System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtStaffName.BackColor = System.Drawing.SystemColors.Control
        Me.txtStaffName.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtStaffName.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtStaffName.Location = New System.Drawing.Point(23, 240)
        Me.txtStaffName.Name = "txtStaffName"
        Me.txtStaffName.ReadOnly = True
        Me.txtStaffName.Size = New System.Drawing.Size(291, 22)
        Me.txtStaffName.TabIndex = 31
        Me.txtStaffName.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(20, 224)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(60, 13)
        Me.Label4.TabIndex = 30
        Me.Label4.Text = "Staff Name"
        '
        'cboStaffNo
        '
        Me.cboStaffNo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboStaffNo.FormattingEnabled = True
        Me.cboStaffNo.Location = New System.Drawing.Point(223, 196)
        Me.cboStaffNo.Name = "cboStaffNo"
        Me.cboStaffNo.Size = New System.Drawing.Size(91, 21)
        Me.cboStaffNo.TabIndex = 29
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(220, 180)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(69, 13)
        Me.Label3.TabIndex = 28
        Me.Label3.Text = "Staff Number"
        '
        'TextBox3
        '
        Me.TextBox3.Anchor = CType((System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TextBox3.BackColor = System.Drawing.Color.MediumAquamarine
        Me.TextBox3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.TextBox3.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox3.Location = New System.Drawing.Point(9, 144)
        Me.TextBox3.Name = "TextBox3"
        Me.TextBox3.ReadOnly = True
        Me.TextBox3.Size = New System.Drawing.Size(311, 22)
        Me.TextBox3.TabIndex = 27
        Me.TextBox3.Text = "Department Head"
        Me.TextBox3.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'txtDepCode
        '
        Me.txtDepCode.BackColor = System.Drawing.SystemColors.Window
        Me.txtDepCode.Location = New System.Drawing.Point(23, 103)
        Me.txtDepCode.Name = "txtDepCode"
        Me.txtDepCode.Size = New System.Drawing.Size(291, 20)
        Me.txtDepCode.TabIndex = 26
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(20, 80)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(90, 13)
        Me.Label2.TabIndex = 25
        Me.Label2.Text = "Department Code"
        '
        'txtDepName
        '
        Me.txtDepName.Location = New System.Drawing.Point(23, 48)
        Me.txtDepName.Name = "txtDepName"
        Me.txtDepName.Size = New System.Drawing.Size(291, 20)
        Me.txtDepName.TabIndex = 24
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(20, 25)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(93, 13)
        Me.Label1.TabIndex = 23
        Me.Label1.Text = "Department Name"
        '
        'lstDepartMents
        '
        Me.lstDepartMents.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColumnHeader1, Me.ColumnHeader2, Me.ColumnHeader3})
        Me.lstDepartMents.Dock = System.Windows.Forms.DockStyle.Right
        Me.lstDepartMents.FullRowSelect = True
        Me.lstDepartMents.GridLines = True
        Me.lstDepartMents.Location = New System.Drawing.Point(320, 10)
        Me.lstDepartMents.Name = "lstDepartMents"
        Me.lstDepartMents.Size = New System.Drawing.Size(513, 265)
        Me.lstDepartMents.TabIndex = 22
        Me.lstDepartMents.UseCompatibleStateImageBehavior = False
        Me.lstDepartMents.View = System.Windows.Forms.View.Details
        '
        'ColumnHeader1
        '
        Me.ColumnHeader1.Text = "CODE"
        Me.ColumnHeader1.Width = 80
        '
        'ColumnHeader2
        '
        Me.ColumnHeader2.Text = "NAME"
        Me.ColumnHeader2.Width = 175
        '
        'ColumnHeader3
        '
        Me.ColumnHeader3.Text = "HEAD NAME"
        Me.ColumnHeader3.Width = 250
        '
        'StatusStrip5
        '
        Me.StatusStrip5.AutoSize = False
        Me.StatusStrip5.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip5.Dock = System.Windows.Forms.DockStyle.Right
        Me.StatusStrip5.Location = New System.Drawing.Point(833, 10)
        Me.StatusStrip5.Name = "StatusStrip5"
        Me.StatusStrip5.Size = New System.Drawing.Size(10, 265)
        Me.StatusStrip5.SizingGrip = False
        Me.StatusStrip5.TabIndex = 21
        Me.StatusStrip5.Text = "StatusStrip5"
        '
        'StatusStrip4
        '
        Me.StatusStrip4.AutoSize = False
        Me.StatusStrip4.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip4.Dock = System.Windows.Forms.DockStyle.Left
        Me.StatusStrip4.Location = New System.Drawing.Point(0, 10)
        Me.StatusStrip4.Name = "StatusStrip4"
        Me.StatusStrip4.Size = New System.Drawing.Size(10, 265)
        Me.StatusStrip4.SizingGrip = False
        Me.StatusStrip4.TabIndex = 20
        Me.StatusStrip4.Text = "StatusStrip4"
        '
        'StatusStrip3
        '
        Me.StatusStrip3.AutoSize = False
        Me.StatusStrip3.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip3.Location = New System.Drawing.Point(0, 275)
        Me.StatusStrip3.Name = "StatusStrip3"
        Me.StatusStrip3.Size = New System.Drawing.Size(843, 10)
        Me.StatusStrip3.SizingGrip = False
        Me.StatusStrip3.TabIndex = 19
        Me.StatusStrip3.Text = "StatusStrip3"
        '
        'StatusStrip10
        '
        Me.StatusStrip10.AutoSize = False
        Me.StatusStrip10.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip10.Dock = System.Windows.Forms.DockStyle.Top
        Me.StatusStrip10.Location = New System.Drawing.Point(0, 0)
        Me.StatusStrip10.Name = "StatusStrip10"
        Me.StatusStrip10.Size = New System.Drawing.Size(843, 10)
        Me.StatusStrip10.SizingGrip = False
        Me.StatusStrip10.TabIndex = 18
        Me.StatusStrip10.Text = "StatusStrip10"
        '
        'btnView
        '
        Me.btnView.Location = New System.Drawing.Point(358, 26)
        Me.btnView.Name = "btnView"
        Me.btnView.Size = New System.Drawing.Size(75, 23)
        Me.btnView.TabIndex = 32
        Me.btnView.Text = "View"
        Me.btnView.UseVisualStyleBackColor = True
        '
        'btnDelete
        '
        Me.btnDelete.Enabled = False
        Me.btnDelete.Location = New System.Drawing.Point(658, 26)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(75, 23)
        Me.btnDelete.TabIndex = 31
        Me.btnDelete.Text = "Delete"
        Me.btnDelete.UseVisualStyleBackColor = True
        '
        'btnUpdate
        '
        Me.btnUpdate.Enabled = False
        Me.btnUpdate.Location = New System.Drawing.Point(559, 26)
        Me.btnUpdate.Name = "btnUpdate"
        Me.btnUpdate.Size = New System.Drawing.Size(75, 23)
        Me.btnUpdate.TabIndex = 30
        Me.btnUpdate.Text = "Update"
        Me.btnUpdate.UseVisualStyleBackColor = True
        '
        'btnSave
        '
        Me.btnSave.Location = New System.Drawing.Point(459, 26)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(75, 23)
        Me.btnSave.TabIndex = 29
        Me.btnSave.Text = "Save"
        Me.btnSave.UseVisualStyleBackColor = True
        '
        'btnClose
        '
        Me.btnClose.Location = New System.Drawing.Point(755, 26)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(75, 23)
        Me.btnClose.TabIndex = 28
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'StatusStrip6
        '
        Me.StatusStrip6.AutoSize = False
        Me.StatusStrip6.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip6.Dock = System.Windows.Forms.DockStyle.Right
        Me.StatusStrip6.Location = New System.Drawing.Point(833, 20)
        Me.StatusStrip6.Name = "StatusStrip6"
        Me.StatusStrip6.Size = New System.Drawing.Size(10, 34)
        Me.StatusStrip6.SizingGrip = False
        Me.StatusStrip6.TabIndex = 27
        Me.StatusStrip6.Text = "StatusStrip6"
        '
        'StatusStrip7
        '
        Me.StatusStrip7.AutoSize = False
        Me.StatusStrip7.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip7.Dock = System.Windows.Forms.DockStyle.Left
        Me.StatusStrip7.Location = New System.Drawing.Point(0, 20)
        Me.StatusStrip7.Name = "StatusStrip7"
        Me.StatusStrip7.Size = New System.Drawing.Size(10, 34)
        Me.StatusStrip7.SizingGrip = False
        Me.StatusStrip7.TabIndex = 26
        Me.StatusStrip7.Text = "StatusStrip7"
        '
        'StatusStrip8
        '
        Me.StatusStrip8.AutoSize = False
        Me.StatusStrip8.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip8.Location = New System.Drawing.Point(0, 54)
        Me.StatusStrip8.Name = "StatusStrip8"
        Me.StatusStrip8.Size = New System.Drawing.Size(843, 10)
        Me.StatusStrip8.SizingGrip = False
        Me.StatusStrip8.TabIndex = 25
        Me.StatusStrip8.Text = "StatusStrip8"
        '
        'StatusStrip9
        '
        Me.StatusStrip9.AutoSize = False
        Me.StatusStrip9.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip9.Dock = System.Windows.Forms.DockStyle.Top
        Me.StatusStrip9.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripStatusLabel2})
        Me.StatusStrip9.Location = New System.Drawing.Point(0, 0)
        Me.StatusStrip9.Name = "StatusStrip9"
        Me.StatusStrip9.Size = New System.Drawing.Size(843, 20)
        Me.StatusStrip9.SizingGrip = False
        Me.StatusStrip9.TabIndex = 24
        Me.StatusStrip9.Text = "StatusStrip9"
        '
        'ToolStripStatusLabel2
        '
        Me.ToolStripStatusLabel2.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ToolStripStatusLabel2.Name = "ToolStripStatusLabel2"
        Me.ToolStripStatusLabel2.Size = New System.Drawing.Size(77, 15)
        Me.ToolStripStatusLabel2.Text = "User Actions"
        '
        'StatusStrip2
        '
        Me.StatusStrip2.AutoSize = False
        Me.StatusStrip2.BackColor = System.Drawing.Color.LightSeaGreen
        Me.StatusStrip2.Dock = System.Windows.Forms.DockStyle.Top
        Me.StatusStrip2.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripStatusLabel1})
        Me.StatusStrip2.Location = New System.Drawing.Point(0, 0)
        Me.StatusStrip2.Name = "StatusStrip2"
        Me.StatusStrip2.Size = New System.Drawing.Size(845, 25)
        Me.StatusStrip2.SizingGrip = False
        Me.StatusStrip2.TabIndex = 16
        Me.StatusStrip2.Text = "StatusStrip2"
        '
        'ToolStripStatusLabel1
        '
        Me.ToolStripStatusLabel1.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ToolStripStatusLabel1.Name = "ToolStripStatusLabel1"
        Me.ToolStripStatusLabel1.Size = New System.Drawing.Size(157, 20)
        Me.ToolStripStatusLabel1.Text = "SCHOOL DEPARTMENTS"
        '
        'StatusStrip1
        '
        Me.StatusStrip1.AutoSize = False
        Me.StatusStrip1.BackColor = System.Drawing.Color.LightSeaGreen
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 382)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(845, 25)
        Me.StatusStrip1.SizingGrip = False
        Me.StatusStrip1.TabIndex = 15
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'frmSchDepartments
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.LightBlue
        Me.ClientSize = New System.Drawing.Size(871, 433)
        Me.Controls.Add(Me.pnlSchDepartment)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Name = "frmSchDepartments"
        Me.ShowIcon = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Departments"
        Me.pnlSchDepartment.ResumeLayout(False)
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel1.PerformLayout()
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.ResumeLayout(False)
        Me.StatusStrip9.ResumeLayout(False)
        Me.StatusStrip9.PerformLayout()
        Me.StatusStrip2.ResumeLayout(False)
        Me.StatusStrip2.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents pnlSchDepartment As System.Windows.Forms.Panel
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents StatusStrip2 As System.Windows.Forms.StatusStrip
    Friend WithEvents ToolStripStatusLabel1 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip5 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip4 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip3 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip10 As System.Windows.Forms.StatusStrip
    Friend WithEvents txtStaffName As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents cboStaffNo As System.Windows.Forms.ComboBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents TextBox3 As System.Windows.Forms.TextBox
    Friend WithEvents txtDepCode As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtDepName As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents lstDepartMents As System.Windows.Forms.ListView
    Friend WithEvents ColumnHeader1 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader2 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader3 As System.Windows.Forms.ColumnHeader
    Friend WithEvents btnView As System.Windows.Forms.Button
    Friend WithEvents btnDelete As System.Windows.Forms.Button
    Friend WithEvents btnUpdate As System.Windows.Forms.Button
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents StatusStrip6 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip7 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip8 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip9 As System.Windows.Forms.StatusStrip
    Friend WithEvents ToolStripStatusLabel2 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents cboStaffType As System.Windows.Forms.ComboBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
End Class
