<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmUsers
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
        Me.pnlUsers = New System.Windows.Forms.Panel()
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
        Me.cboDomain = New System.Windows.Forms.ComboBox()
        Me.txtPassword = New System.Windows.Forms.TextBox()
        Me.txtUserName = New System.Windows.Forms.TextBox()
        Me.cboEmpNo = New System.Windows.Forms.ComboBox()
        Me.cboContactType = New System.Windows.Forms.ComboBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.cboStaffType = New System.Windows.Forms.ComboBox()
        Me.txtFullName = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.StatusStrip5 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip4 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip3 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip6 = New System.Windows.Forms.StatusStrip()
        Me.ToolStripStatusLabel2 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.btnDelete = New System.Windows.Forms.Button()
        Me.btnSave = New System.Windows.Forms.Button()
        Me.btnUpdate = New System.Windows.Forms.Button()
        Me.btnClose = New System.Windows.Forms.Button()
        Me.StatusStrip7 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip8 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip9 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip10 = New System.Windows.Forms.StatusStrip()
        Me.ToolStripStatusLabel3 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.StatusStrip2 = New System.Windows.Forms.StatusStrip()
        Me.ToolStripStatusLabel1 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip()
        Me.pnlUsers.SuspendLayout()
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.StatusStrip6.SuspendLayout()
        Me.StatusStrip10.SuspendLayout()
        Me.StatusStrip2.SuspendLayout()
        Me.SuspendLayout()
        '
        'pnlUsers
        '
        Me.pnlUsers.BackColor = System.Drawing.SystemColors.Control
        Me.pnlUsers.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pnlUsers.Controls.Add(Me.SplitContainer1)
        Me.pnlUsers.Controls.Add(Me.StatusStrip2)
        Me.pnlUsers.Controls.Add(Me.StatusStrip1)
        Me.pnlUsers.Location = New System.Drawing.Point(12, 12)
        Me.pnlUsers.Name = "pnlUsers"
        Me.pnlUsers.Size = New System.Drawing.Size(715, 320)
        Me.pnlUsers.TabIndex = 0
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
        Me.SplitContainer1.Panel1.Controls.Add(Me.cboDomain)
        Me.SplitContainer1.Panel1.Controls.Add(Me.txtPassword)
        Me.SplitContainer1.Panel1.Controls.Add(Me.txtUserName)
        Me.SplitContainer1.Panel1.Controls.Add(Me.cboEmpNo)
        Me.SplitContainer1.Panel1.Controls.Add(Me.cboContactType)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label7)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label6)
        Me.SplitContainer1.Panel1.Controls.Add(Me.cboStaffType)
        Me.SplitContainer1.Panel1.Controls.Add(Me.txtFullName)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label5)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label4)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label3)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label2)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label1)
        Me.SplitContainer1.Panel1.Controls.Add(Me.StatusStrip5)
        Me.SplitContainer1.Panel1.Controls.Add(Me.StatusStrip4)
        Me.SplitContainer1.Panel1.Controls.Add(Me.StatusStrip3)
        Me.SplitContainer1.Panel1.Controls.Add(Me.StatusStrip6)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.btnDelete)
        Me.SplitContainer1.Panel2.Controls.Add(Me.btnSave)
        Me.SplitContainer1.Panel2.Controls.Add(Me.btnUpdate)
        Me.SplitContainer1.Panel2.Controls.Add(Me.btnClose)
        Me.SplitContainer1.Panel2.Controls.Add(Me.StatusStrip7)
        Me.SplitContainer1.Panel2.Controls.Add(Me.StatusStrip8)
        Me.SplitContainer1.Panel2.Controls.Add(Me.StatusStrip9)
        Me.SplitContainer1.Panel2.Controls.Add(Me.StatusStrip10)
        Me.SplitContainer1.Size = New System.Drawing.Size(713, 268)
        Me.SplitContainer1.SplitterDistance = 188
        Me.SplitContainer1.SplitterWidth = 1
        Me.SplitContainer1.TabIndex = 7
        '
        'cboDomain
        '
        Me.cboDomain.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboDomain.FormattingEnabled = True
        Me.cboDomain.Location = New System.Drawing.Point(461, 104)
        Me.cboDomain.Name = "cboDomain"
        Me.cboDomain.Size = New System.Drawing.Size(220, 21)
        Me.cboDomain.TabIndex = 27
        '
        'txtPassword
        '
        Me.txtPassword.Location = New System.Drawing.Point(461, 71)
        Me.txtPassword.Name = "txtPassword"
        Me.txtPassword.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.txtPassword.Size = New System.Drawing.Size(220, 20)
        Me.txtPassword.TabIndex = 26
        '
        'txtUserName
        '
        Me.txtUserName.Location = New System.Drawing.Point(461, 39)
        Me.txtUserName.Name = "txtUserName"
        Me.txtUserName.Size = New System.Drawing.Size(220, 20)
        Me.txtUserName.TabIndex = 25
        '
        'cboEmpNo
        '
        Me.cboEmpNo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboEmpNo.FormattingEnabled = True
        Me.cboEmpNo.Location = New System.Drawing.Point(133, 104)
        Me.cboEmpNo.Name = "cboEmpNo"
        Me.cboEmpNo.Size = New System.Drawing.Size(211, 21)
        Me.cboEmpNo.TabIndex = 24
        '
        'cboContactType
        '
        Me.cboContactType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboContactType.FormattingEnabled = True
        Me.cboContactType.Location = New System.Drawing.Point(133, 71)
        Me.cboContactType.Name = "cboContactType"
        Me.cboContactType.Size = New System.Drawing.Size(211, 21)
        Me.cboContactType.TabIndex = 23
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(34, 74)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(77, 13)
        Me.Label7.TabIndex = 22
        Me.Label7.Text = "Contract Type:"
        Me.Label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(34, 107)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(73, 13)
        Me.Label6.TabIndex = 21
        Me.Label6.Text = "Employee No:"
        Me.Label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'cboStaffType
        '
        Me.cboStaffType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboStaffType.FormattingEnabled = True
        Me.cboStaffType.Location = New System.Drawing.Point(133, 39)
        Me.cboStaffType.Name = "cboStaffType"
        Me.cboStaffType.Size = New System.Drawing.Size(211, 21)
        Me.cboStaffType.TabIndex = 20
        '
        'txtFullName
        '
        Me.txtFullName.Location = New System.Drawing.Point(133, 137)
        Me.txtFullName.Name = "txtFullName"
        Me.txtFullName.ReadOnly = True
        Me.txtFullName.Size = New System.Drawing.Size(548, 20)
        Me.txtFullName.TabIndex = 19
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(378, 107)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(77, 13)
        Me.Label5.TabIndex = 18
        Me.Label5.Text = "Domain Name:"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(378, 74)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(59, 13)
        Me.Label4.TabIndex = 17
        Me.Label4.Text = "PassWord:"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(378, 42)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(63, 13)
        Me.Label3.TabIndex = 16
        Me.Label3.Text = "User Name:"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(34, 140)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(57, 13)
        Me.Label2.TabIndex = 15
        Me.Label2.Text = "Full Name:"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(34, 42)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(59, 13)
        Me.Label1.TabIndex = 14
        Me.Label1.Text = "Staff Type:"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'StatusStrip5
        '
        Me.StatusStrip5.AutoSize = False
        Me.StatusStrip5.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip5.Dock = System.Windows.Forms.DockStyle.Right
        Me.StatusStrip5.Location = New System.Drawing.Point(701, 20)
        Me.StatusStrip5.Name = "StatusStrip5"
        Me.StatusStrip5.Size = New System.Drawing.Size(10, 156)
        Me.StatusStrip5.SizingGrip = False
        Me.StatusStrip5.TabIndex = 13
        Me.StatusStrip5.Text = "StatusStrip5"
        '
        'StatusStrip4
        '
        Me.StatusStrip4.AutoSize = False
        Me.StatusStrip4.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip4.Dock = System.Windows.Forms.DockStyle.Left
        Me.StatusStrip4.Location = New System.Drawing.Point(0, 20)
        Me.StatusStrip4.Name = "StatusStrip4"
        Me.StatusStrip4.Size = New System.Drawing.Size(10, 156)
        Me.StatusStrip4.SizingGrip = False
        Me.StatusStrip4.TabIndex = 12
        Me.StatusStrip4.Text = "StatusStrip4"
        '
        'StatusStrip3
        '
        Me.StatusStrip3.AutoSize = False
        Me.StatusStrip3.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip3.Location = New System.Drawing.Point(0, 176)
        Me.StatusStrip3.Name = "StatusStrip3"
        Me.StatusStrip3.Size = New System.Drawing.Size(711, 10)
        Me.StatusStrip3.SizingGrip = False
        Me.StatusStrip3.TabIndex = 11
        Me.StatusStrip3.Text = "StatusStrip3"
        '
        'StatusStrip6
        '
        Me.StatusStrip6.AutoSize = False
        Me.StatusStrip6.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip6.Dock = System.Windows.Forms.DockStyle.Top
        Me.StatusStrip6.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripStatusLabel2})
        Me.StatusStrip6.Location = New System.Drawing.Point(0, 0)
        Me.StatusStrip6.Name = "StatusStrip6"
        Me.StatusStrip6.Size = New System.Drawing.Size(711, 20)
        Me.StatusStrip6.SizingGrip = False
        Me.StatusStrip6.TabIndex = 10
        Me.StatusStrip6.Text = "StatusStrip6"
        '
        'ToolStripStatusLabel2
        '
        Me.ToolStripStatusLabel2.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ToolStripStatusLabel2.Name = "ToolStripStatusLabel2"
        Me.ToolStripStatusLabel2.Size = New System.Drawing.Size(78, 15)
        Me.ToolStripStatusLabel2.Text = "Enter Details"
        '
        'btnDelete
        '
        Me.btnDelete.Location = New System.Drawing.Point(494, 32)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(75, 23)
        Me.btnDelete.TabIndex = 16
        Me.btnDelete.Text = "Delete"
        Me.btnDelete.UseVisualStyleBackColor = True
        '
        'btnSave
        '
        Me.btnSave.Location = New System.Drawing.Point(269, 32)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(75, 23)
        Me.btnSave.TabIndex = 15
        Me.btnSave.Text = "Save"
        Me.btnSave.UseVisualStyleBackColor = True
        '
        'btnUpdate
        '
        Me.btnUpdate.Location = New System.Drawing.Point(381, 32)
        Me.btnUpdate.Name = "btnUpdate"
        Me.btnUpdate.Size = New System.Drawing.Size(75, 23)
        Me.btnUpdate.TabIndex = 14
        Me.btnUpdate.Text = "Update"
        Me.btnUpdate.UseVisualStyleBackColor = True
        '
        'btnClose
        '
        Me.btnClose.Location = New System.Drawing.Point(606, 32)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(75, 23)
        Me.btnClose.TabIndex = 0
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'StatusStrip7
        '
        Me.StatusStrip7.AutoSize = False
        Me.StatusStrip7.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip7.Dock = System.Windows.Forms.DockStyle.Right
        Me.StatusStrip7.Location = New System.Drawing.Point(701, 20)
        Me.StatusStrip7.Name = "StatusStrip7"
        Me.StatusStrip7.Size = New System.Drawing.Size(10, 47)
        Me.StatusStrip7.SizingGrip = False
        Me.StatusStrip7.TabIndex = 13
        Me.StatusStrip7.Text = "StatusStrip7"
        '
        'StatusStrip8
        '
        Me.StatusStrip8.AutoSize = False
        Me.StatusStrip8.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip8.Dock = System.Windows.Forms.DockStyle.Left
        Me.StatusStrip8.Location = New System.Drawing.Point(0, 20)
        Me.StatusStrip8.Name = "StatusStrip8"
        Me.StatusStrip8.Size = New System.Drawing.Size(10, 47)
        Me.StatusStrip8.SizingGrip = False
        Me.StatusStrip8.TabIndex = 12
        Me.StatusStrip8.Text = "StatusStrip8"
        '
        'StatusStrip9
        '
        Me.StatusStrip9.AutoSize = False
        Me.StatusStrip9.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip9.Location = New System.Drawing.Point(0, 67)
        Me.StatusStrip9.Name = "StatusStrip9"
        Me.StatusStrip9.Size = New System.Drawing.Size(711, 10)
        Me.StatusStrip9.SizingGrip = False
        Me.StatusStrip9.TabIndex = 11
        Me.StatusStrip9.Text = "StatusStrip9"
        '
        'StatusStrip10
        '
        Me.StatusStrip10.AutoSize = False
        Me.StatusStrip10.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip10.Dock = System.Windows.Forms.DockStyle.Top
        Me.StatusStrip10.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripStatusLabel3})
        Me.StatusStrip10.Location = New System.Drawing.Point(0, 0)
        Me.StatusStrip10.Name = "StatusStrip10"
        Me.StatusStrip10.Size = New System.Drawing.Size(711, 20)
        Me.StatusStrip10.SizingGrip = False
        Me.StatusStrip10.TabIndex = 10
        Me.StatusStrip10.Text = "StatusStrip10"
        '
        'ToolStripStatusLabel3
        '
        Me.ToolStripStatusLabel3.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ToolStripStatusLabel3.Name = "ToolStripStatusLabel3"
        Me.ToolStripStatusLabel3.Size = New System.Drawing.Size(80, 15)
        Me.ToolStripStatusLabel3.Text = "Carry Actions"
        '
        'StatusStrip2
        '
        Me.StatusStrip2.AutoSize = False
        Me.StatusStrip2.BackColor = System.Drawing.Color.LightSeaGreen
        Me.StatusStrip2.Dock = System.Windows.Forms.DockStyle.Top
        Me.StatusStrip2.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripStatusLabel1})
        Me.StatusStrip2.Location = New System.Drawing.Point(0, 0)
        Me.StatusStrip2.Name = "StatusStrip2"
        Me.StatusStrip2.Size = New System.Drawing.Size(713, 25)
        Me.StatusStrip2.SizingGrip = False
        Me.StatusStrip2.TabIndex = 6
        Me.StatusStrip2.Text = "StatusStrip2"
        '
        'ToolStripStatusLabel1
        '
        Me.ToolStripStatusLabel1.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ToolStripStatusLabel1.Name = "ToolStripStatusLabel1"
        Me.ToolStripStatusLabel1.Size = New System.Drawing.Size(99, 20)
        Me.ToolStripStatusLabel1.Text = "SYSTEM USERS"
        '
        'StatusStrip1
        '
        Me.StatusStrip1.AutoSize = False
        Me.StatusStrip1.BackColor = System.Drawing.Color.LightSeaGreen
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 293)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(713, 25)
        Me.StatusStrip1.SizingGrip = False
        Me.StatusStrip1.TabIndex = 5
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'frmUsers
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.LightSkyBlue
        Me.ClientSize = New System.Drawing.Size(739, 342)
        Me.Controls.Add(Me.pnlUsers)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.MaximizeBox = False
        Me.Name = "frmUsers"
        Me.ShowIcon = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Users"
        Me.pnlUsers.ResumeLayout(False)
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel1.PerformLayout()
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.ResumeLayout(False)
        Me.StatusStrip6.ResumeLayout(False)
        Me.StatusStrip6.PerformLayout()
        Me.StatusStrip10.ResumeLayout(False)
        Me.StatusStrip10.PerformLayout()
        Me.StatusStrip2.ResumeLayout(False)
        Me.StatusStrip2.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents pnlUsers As System.Windows.Forms.Panel
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip2 As System.Windows.Forms.StatusStrip
    Friend WithEvents ToolStripStatusLabel1 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents StatusStrip5 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip4 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip3 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip6 As System.Windows.Forms.StatusStrip
    Friend WithEvents btnDelete As System.Windows.Forms.Button
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents btnUpdate As System.Windows.Forms.Button
    Friend WithEvents StatusStrip7 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip8 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip9 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip10 As System.Windows.Forms.StatusStrip
    Friend WithEvents cboDomain As System.Windows.Forms.ComboBox
    Friend WithEvents txtPassword As System.Windows.Forms.TextBox
    Friend WithEvents txtUserName As System.Windows.Forms.TextBox
    Friend WithEvents cboEmpNo As System.Windows.Forms.ComboBox
    Friend WithEvents cboContactType As System.Windows.Forms.ComboBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents cboStaffType As System.Windows.Forms.ComboBox
    Friend WithEvents txtFullName As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents ToolStripStatusLabel2 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents ToolStripStatusLabel3 As System.Windows.Forms.ToolStripStatusLabel
End Class
