<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmLogin
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
        Me.pnlLogin = New System.Windows.Forms.Panel()
        Me.cboDomain = New System.Windows.Forms.ComboBox()
        Me.txtPassword = New System.Windows.Forms.TextBox()
        Me.txtUserName = New System.Windows.Forms.TextBox()
        Me.btnClose = New System.Windows.Forms.Button()
        Me.btnLogin = New System.Windows.Forms.Button()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip()
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
        Me.APPLICATIONSETTINGSToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.pnlLogin.SuspendLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.MenuStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'pnlLogin
        '
        Me.pnlLogin.BackColor = System.Drawing.SystemColors.Control
        Me.pnlLogin.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pnlLogin.Controls.Add(Me.cboDomain)
        Me.pnlLogin.Controls.Add(Me.txtPassword)
        Me.pnlLogin.Controls.Add(Me.txtUserName)
        Me.pnlLogin.Controls.Add(Me.btnClose)
        Me.pnlLogin.Controls.Add(Me.btnLogin)
        Me.pnlLogin.Controls.Add(Me.Label3)
        Me.pnlLogin.Controls.Add(Me.Label2)
        Me.pnlLogin.Controls.Add(Me.Label1)
        Me.pnlLogin.Controls.Add(Me.PictureBox1)
        Me.pnlLogin.Controls.Add(Me.StatusStrip1)
        Me.pnlLogin.Controls.Add(Me.MenuStrip1)
        Me.pnlLogin.Location = New System.Drawing.Point(16, 15)
        Me.pnlLogin.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.pnlLogin.Name = "pnlLogin"
        Me.pnlLogin.Size = New System.Drawing.Size(507, 257)
        Me.pnlLogin.TabIndex = 0
        '
        'cboDomain
        '
        Me.cboDomain.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboDomain.FormattingEnabled = True
        Me.cboDomain.Location = New System.Drawing.Point(232, 150)
        Me.cboDomain.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.cboDomain.Name = "cboDomain"
        Me.cboDomain.Size = New System.Drawing.Size(256, 24)
        Me.cboDomain.TabIndex = 12
        '
        'txtPassword
        '
        Me.txtPassword.Location = New System.Drawing.Point(232, 102)
        Me.txtPassword.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.txtPassword.Name = "txtPassword"
        Me.txtPassword.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.txtPassword.Size = New System.Drawing.Size(256, 22)
        Me.txtPassword.TabIndex = 11
        '
        'txtUserName
        '
        Me.txtUserName.Location = New System.Drawing.Point(232, 54)
        Me.txtUserName.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.txtUserName.Name = "txtUserName"
        Me.txtUserName.Size = New System.Drawing.Size(256, 22)
        Me.txtUserName.TabIndex = 10
        '
        'btnClose
        '
        Me.btnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnClose.Location = New System.Drawing.Point(389, 188)
        Me.btnClose.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(100, 28)
        Me.btnClose.TabIndex = 9
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'btnLogin
        '
        Me.btnLogin.Location = New System.Drawing.Point(232, 188)
        Me.btnLogin.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.btnLogin.Name = "btnLogin"
        Me.btnLogin.Size = New System.Drawing.Size(100, 28)
        Me.btnLogin.TabIndex = 8
        Me.btnLogin.Text = "Login"
        Me.btnLogin.UseVisualStyleBackColor = True
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(228, 130)
        Me.Label3.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(56, 17)
        Me.Label3.TabIndex = 7
        Me.Label3.Text = "Domain"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(228, 82)
        Me.Label2.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(73, 17)
        Me.Label2.TabIndex = 6
        Me.Label2.Text = "PassWord"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(228, 34)
        Me.Label1.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(75, 17)
        Me.Label1.TabIndex = 5
        Me.Label1.Text = "UserName"
        '
        'PictureBox1
        '
        Me.PictureBox1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.PictureBox1.Image = Global.School_Soft.My.Resources.Resources.user_login_icon
        Me.PictureBox1.Location = New System.Drawing.Point(4, 34)
        Me.PictureBox1.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(215, 185)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox1.TabIndex = 4
        Me.PictureBox1.TabStop = False
        '
        'StatusStrip1
        '
        Me.StatusStrip1.AutoSize = False
        Me.StatusStrip1.BackColor = System.Drawing.Color.LightSeaGreen
        Me.StatusStrip1.ImageScalingSize = New System.Drawing.Size(20, 20)
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 224)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Padding = New System.Windows.Forms.Padding(1, 0, 19, 0)
        Me.StatusStrip1.Size = New System.Drawing.Size(505, 31)
        Me.StatusStrip1.SizingGrip = False
        Me.StatusStrip1.TabIndex = 3
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'MenuStrip1
        '
        Me.MenuStrip1.AutoSize = False
        Me.MenuStrip1.BackColor = System.Drawing.Color.LightSeaGreen
        Me.MenuStrip1.ImageScalingSize = New System.Drawing.Size(20, 20)
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.APPLICATIONSETTINGSToolStripMenuItem})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Padding = New System.Windows.Forms.Padding(8, 2, 0, 2)
        Me.MenuStrip1.Size = New System.Drawing.Size(505, 31)
        Me.MenuStrip1.TabIndex = 2
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        'APPLICATIONSETTINGSToolStripMenuItem
        '
        Me.APPLICATIONSETTINGSToolStripMenuItem.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.APPLICATIONSETTINGSToolStripMenuItem.Name = "APPLICATIONSETTINGSToolStripMenuItem"
        Me.APPLICATIONSETTINGSToolStripMenuItem.Size = New System.Drawing.Size(208, 27)
        Me.APPLICATIONSETTINGSToolStripMenuItem.Text = "LOGIN TO THE SYSTEM"
        '
        'frmLogin
        '
        Me.AcceptButton = Me.btnLogin
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.LightSkyBlue
        Me.CancelButton = Me.btnClose
        Me.ClientSize = New System.Drawing.Size(539, 286)
        Me.Controls.Add(Me.pnlLogin)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.Name = "frmLogin"
        Me.ShowIcon = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Login"
        Me.pnlLogin.ResumeLayout(False)
        Me.pnlLogin.PerformLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents pnlLogin As System.Windows.Forms.Panel
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents MenuStrip1 As System.Windows.Forms.MenuStrip
    Friend WithEvents APPLICATIONSETTINGSToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents cboDomain As System.Windows.Forms.ComboBox
    Friend WithEvents txtPassword As System.Windows.Forms.TextBox
    Friend WithEvents txtUserName As System.Windows.Forms.TextBox
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents btnLogin As System.Windows.Forms.Button
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
End Class
