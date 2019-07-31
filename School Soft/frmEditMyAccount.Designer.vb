<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmEditMyAccount
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
        Me.pnlEditMyAccount = New System.Windows.Forms.Panel()
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
        Me.txtConfirmPassword = New System.Windows.Forms.TextBox()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.StatusStrip7 = New System.Windows.Forms.StatusStrip()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.StatusStrip4 = New System.Windows.Forms.StatusStrip()
        Me.txtNewPassword = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.txtCurrentPassword = New System.Windows.Forms.TextBox()
        Me.StatusStrip9 = New System.Windows.Forms.StatusStrip()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.StatusStrip6 = New System.Windows.Forms.StatusStrip()
        Me.ToolStripStatusLabel2 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.txtEmployeeName = New System.Windows.Forms.TextBox()
        Me.txtUserName = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.btnSave = New System.Windows.Forms.Button()
        Me.btnClose = New System.Windows.Forms.Button()
        Me.StatusStrip5 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip8 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip10 = New System.Windows.Forms.StatusStrip()
        Me.ToolStripStatusLabel3 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.StatusStrip3 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip2 = New System.Windows.Forms.StatusStrip()
        Me.ToolStripStatusLabel1 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip()
        Me.pnlEditMyAccount.SuspendLayout()
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.StatusStrip6.SuspendLayout()
        Me.StatusStrip10.SuspendLayout()
        Me.StatusStrip2.SuspendLayout()
        Me.SuspendLayout()
        '
        'pnlEditMyAccount
        '
        Me.pnlEditMyAccount.BackColor = System.Drawing.SystemColors.Control
        Me.pnlEditMyAccount.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pnlEditMyAccount.Controls.Add(Me.SplitContainer1)
        Me.pnlEditMyAccount.Controls.Add(Me.StatusStrip2)
        Me.pnlEditMyAccount.Controls.Add(Me.StatusStrip1)
        Me.pnlEditMyAccount.Location = New System.Drawing.Point(12, 12)
        Me.pnlEditMyAccount.Name = "pnlEditMyAccount"
        Me.pnlEditMyAccount.Size = New System.Drawing.Size(575, 308)
        Me.pnlEditMyAccount.TabIndex = 0
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
        Me.SplitContainer1.Panel1.Controls.Add(Me.txtConfirmPassword)
        Me.SplitContainer1.Panel1.Controls.Add(Me.PictureBox1)
        Me.SplitContainer1.Panel1.Controls.Add(Me.StatusStrip7)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label6)
        Me.SplitContainer1.Panel1.Controls.Add(Me.StatusStrip4)
        Me.SplitContainer1.Panel1.Controls.Add(Me.txtNewPassword)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label1)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label4)
        Me.SplitContainer1.Panel1.Controls.Add(Me.txtCurrentPassword)
        Me.SplitContainer1.Panel1.Controls.Add(Me.StatusStrip9)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label5)
        Me.SplitContainer1.Panel1.Controls.Add(Me.StatusStrip6)
        Me.SplitContainer1.Panel1.Controls.Add(Me.txtEmployeeName)
        Me.SplitContainer1.Panel1.Controls.Add(Me.txtUserName)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label2)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.btnSave)
        Me.SplitContainer1.Panel2.Controls.Add(Me.btnClose)
        Me.SplitContainer1.Panel2.Controls.Add(Me.StatusStrip5)
        Me.SplitContainer1.Panel2.Controls.Add(Me.StatusStrip8)
        Me.SplitContainer1.Panel2.Controls.Add(Me.StatusStrip10)
        Me.SplitContainer1.Panel2.Controls.Add(Me.StatusStrip3)
        Me.SplitContainer1.Size = New System.Drawing.Size(573, 256)
        Me.SplitContainer1.SplitterDistance = 181
        Me.SplitContainer1.SplitterWidth = 1
        Me.SplitContainer1.TabIndex = 9
        '
        'txtConfirmPassword
        '
        Me.txtConfirmPassword.Location = New System.Drawing.Point(380, 139)
        Me.txtConfirmPassword.Name = "txtConfirmPassword"
        Me.txtConfirmPassword.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.txtConfirmPassword.Size = New System.Drawing.Size(167, 20)
        Me.txtConfirmPassword.TabIndex = 30
        '
        'PictureBox1
        '
        Me.PictureBox1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.PictureBox1.Image = Global.School_Soft.My.Resources.Resources.graduate
        Me.PictureBox1.Location = New System.Drawing.Point(13, 23)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(161, 143)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox1.TabIndex = 9
        Me.PictureBox1.TabStop = False
        '
        'StatusStrip7
        '
        Me.StatusStrip7.AutoSize = False
        Me.StatusStrip7.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip7.Dock = System.Windows.Forms.DockStyle.Right
        Me.StatusStrip7.Location = New System.Drawing.Point(561, 20)
        Me.StatusStrip7.Name = "StatusStrip7"
        Me.StatusStrip7.Size = New System.Drawing.Size(10, 149)
        Me.StatusStrip7.SizingGrip = False
        Me.StatusStrip7.TabIndex = 21
        Me.StatusStrip7.Text = "StatusStrip7"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(377, 120)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(91, 13)
        Me.Label6.TabIndex = 29
        Me.Label6.Text = "Confirm Password"
        '
        'StatusStrip4
        '
        Me.StatusStrip4.AutoSize = False
        Me.StatusStrip4.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip4.Dock = System.Windows.Forms.DockStyle.Left
        Me.StatusStrip4.Location = New System.Drawing.Point(0, 20)
        Me.StatusStrip4.Name = "StatusStrip4"
        Me.StatusStrip4.Size = New System.Drawing.Size(10, 149)
        Me.StatusStrip4.SizingGrip = False
        Me.StatusStrip4.TabIndex = 19
        Me.StatusStrip4.Text = "StatusStrip4"
        '
        'txtNewPassword
        '
        Me.txtNewPassword.Location = New System.Drawing.Point(380, 91)
        Me.txtNewPassword.Name = "txtNewPassword"
        Me.txtNewPassword.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.txtNewPassword.Size = New System.Drawing.Size(167, 20)
        Me.txtNewPassword.TabIndex = 28
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(177, 29)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(84, 13)
        Me.Label1.TabIndex = 10
        Me.Label1.Text = "Employee Name"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(377, 71)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(78, 13)
        Me.Label4.TabIndex = 27
        Me.Label4.Text = "New Password"
        '
        'txtCurrentPassword
        '
        Me.txtCurrentPassword.Location = New System.Drawing.Point(184, 139)
        Me.txtCurrentPassword.Name = "txtCurrentPassword"
        Me.txtCurrentPassword.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.txtCurrentPassword.Size = New System.Drawing.Size(175, 20)
        Me.txtCurrentPassword.TabIndex = 26
        '
        'StatusStrip9
        '
        Me.StatusStrip9.AutoSize = False
        Me.StatusStrip9.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip9.Location = New System.Drawing.Point(0, 169)
        Me.StatusStrip9.Name = "StatusStrip9"
        Me.StatusStrip9.Size = New System.Drawing.Size(571, 10)
        Me.StatusStrip9.SizingGrip = False
        Me.StatusStrip9.TabIndex = 16
        Me.StatusStrip9.Text = "StatusStrip9"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(180, 120)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(90, 13)
        Me.Label5.TabIndex = 25
        Me.Label5.Text = "Current Password"
        '
        'StatusStrip6
        '
        Me.StatusStrip6.AutoSize = False
        Me.StatusStrip6.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip6.Dock = System.Windows.Forms.DockStyle.Top
        Me.StatusStrip6.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripStatusLabel2})
        Me.StatusStrip6.Location = New System.Drawing.Point(0, 0)
        Me.StatusStrip6.Name = "StatusStrip6"
        Me.StatusStrip6.Size = New System.Drawing.Size(571, 20)
        Me.StatusStrip6.SizingGrip = False
        Me.StatusStrip6.TabIndex = 15
        Me.StatusStrip6.Text = "StatusStrip6"
        '
        'ToolStripStatusLabel2
        '
        Me.ToolStripStatusLabel2.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ToolStripStatusLabel2.Name = "ToolStripStatusLabel2"
        Me.ToolStripStatusLabel2.Size = New System.Drawing.Size(78, 15)
        Me.ToolStripStatusLabel2.Text = "Enter Details"
        '
        'txtEmployeeName
        '
        Me.txtEmployeeName.Location = New System.Drawing.Point(184, 48)
        Me.txtEmployeeName.Name = "txtEmployeeName"
        Me.txtEmployeeName.ReadOnly = True
        Me.txtEmployeeName.Size = New System.Drawing.Size(363, 20)
        Me.txtEmployeeName.TabIndex = 20
        '
        'txtUserName
        '
        Me.txtUserName.Location = New System.Drawing.Point(184, 91)
        Me.txtUserName.Name = "txtUserName"
        Me.txtUserName.ReadOnly = True
        Me.txtUserName.Size = New System.Drawing.Size(172, 20)
        Me.txtUserName.TabIndex = 22
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(177, 73)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(60, 13)
        Me.Label2.TabIndex = 21
        Me.Label2.Text = "User Name"
        '
        'btnSave
        '
        Me.btnSave.Location = New System.Drawing.Point(326, 29)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(75, 23)
        Me.btnSave.TabIndex = 26
        Me.btnSave.Text = "Save"
        Me.btnSave.UseVisualStyleBackColor = True
        '
        'btnClose
        '
        Me.btnClose.Location = New System.Drawing.Point(471, 29)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(75, 23)
        Me.btnClose.TabIndex = 25
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'StatusStrip5
        '
        Me.StatusStrip5.AutoSize = False
        Me.StatusStrip5.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip5.Dock = System.Windows.Forms.DockStyle.Right
        Me.StatusStrip5.Location = New System.Drawing.Point(561, 20)
        Me.StatusStrip5.Name = "StatusStrip5"
        Me.StatusStrip5.Size = New System.Drawing.Size(10, 42)
        Me.StatusStrip5.SizingGrip = False
        Me.StatusStrip5.TabIndex = 24
        Me.StatusStrip5.Text = "StatusStrip5"
        '
        'StatusStrip8
        '
        Me.StatusStrip8.AutoSize = False
        Me.StatusStrip8.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip8.Dock = System.Windows.Forms.DockStyle.Left
        Me.StatusStrip8.Location = New System.Drawing.Point(0, 20)
        Me.StatusStrip8.Name = "StatusStrip8"
        Me.StatusStrip8.Size = New System.Drawing.Size(10, 42)
        Me.StatusStrip8.SizingGrip = False
        Me.StatusStrip8.TabIndex = 23
        Me.StatusStrip8.Text = "StatusStrip8"
        '
        'StatusStrip10
        '
        Me.StatusStrip10.AutoSize = False
        Me.StatusStrip10.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip10.Dock = System.Windows.Forms.DockStyle.Top
        Me.StatusStrip10.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripStatusLabel3})
        Me.StatusStrip10.Location = New System.Drawing.Point(0, 0)
        Me.StatusStrip10.Name = "StatusStrip10"
        Me.StatusStrip10.Size = New System.Drawing.Size(571, 20)
        Me.StatusStrip10.SizingGrip = False
        Me.StatusStrip10.TabIndex = 21
        Me.StatusStrip10.Text = "StatusStrip10"
        '
        'ToolStripStatusLabel3
        '
        Me.ToolStripStatusLabel3.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ToolStripStatusLabel3.Name = "ToolStripStatusLabel3"
        Me.ToolStripStatusLabel3.Size = New System.Drawing.Size(80, 15)
        Me.ToolStripStatusLabel3.Text = "Carry Actions"
        '
        'StatusStrip3
        '
        Me.StatusStrip3.AutoSize = False
        Me.StatusStrip3.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip3.Location = New System.Drawing.Point(0, 62)
        Me.StatusStrip3.Name = "StatusStrip3"
        Me.StatusStrip3.Size = New System.Drawing.Size(571, 10)
        Me.StatusStrip3.SizingGrip = False
        Me.StatusStrip3.TabIndex = 22
        Me.StatusStrip3.Text = "StatusStrip3"
        '
        'StatusStrip2
        '
        Me.StatusStrip2.AutoSize = False
        Me.StatusStrip2.BackColor = System.Drawing.Color.LightSeaGreen
        Me.StatusStrip2.Dock = System.Windows.Forms.DockStyle.Top
        Me.StatusStrip2.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripStatusLabel1})
        Me.StatusStrip2.Location = New System.Drawing.Point(0, 0)
        Me.StatusStrip2.Name = "StatusStrip2"
        Me.StatusStrip2.Size = New System.Drawing.Size(573, 25)
        Me.StatusStrip2.SizingGrip = False
        Me.StatusStrip2.TabIndex = 8
        Me.StatusStrip2.Text = "StatusStrip2"
        '
        'ToolStripStatusLabel1
        '
        Me.ToolStripStatusLabel1.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ToolStripStatusLabel1.Name = "ToolStripStatusLabel1"
        Me.ToolStripStatusLabel1.Size = New System.Drawing.Size(127, 20)
        Me.ToolStripStatusLabel1.Text = "EDIT MY ACCOUNT"
        '
        'StatusStrip1
        '
        Me.StatusStrip1.AutoSize = False
        Me.StatusStrip1.BackColor = System.Drawing.Color.LightSeaGreen
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 281)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(573, 25)
        Me.StatusStrip1.SizingGrip = False
        Me.StatusStrip1.TabIndex = 7
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'frmEditMyAccount
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.LightSkyBlue
        Me.ClientSize = New System.Drawing.Size(599, 329)
        Me.Controls.Add(Me.pnlEditMyAccount)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Name = "frmEditMyAccount"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Edit Account"
        Me.pnlEditMyAccount.ResumeLayout(False)
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel1.PerformLayout()
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.ResumeLayout(False)
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.StatusStrip6.ResumeLayout(False)
        Me.StatusStrip6.PerformLayout()
        Me.StatusStrip10.ResumeLayout(False)
        Me.StatusStrip10.PerformLayout()
        Me.StatusStrip2.ResumeLayout(False)
        Me.StatusStrip2.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents pnlEditMyAccount As System.Windows.Forms.Panel
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents StatusStrip2 As System.Windows.Forms.StatusStrip
    Friend WithEvents ToolStripStatusLabel1 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtConfirmPassword As System.Windows.Forms.TextBox
    Friend WithEvents txtEmployeeName As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents txtNewPassword As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtCurrentPassword As System.Windows.Forms.TextBox
    Friend WithEvents txtUserName As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents StatusStrip7 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip4 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip9 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip6 As System.Windows.Forms.StatusStrip
    Friend WithEvents ToolStripStatusLabel2 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents StatusStrip5 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip8 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip10 As System.Windows.Forms.StatusStrip
    Friend WithEvents ToolStripStatusLabel3 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents StatusStrip3 As System.Windows.Forms.StatusStrip
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents btnClose As System.Windows.Forms.Button
End Class
