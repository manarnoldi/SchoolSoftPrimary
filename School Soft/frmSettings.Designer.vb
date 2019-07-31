<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmSettings
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
        Me.pnlSettings = New System.Windows.Forms.Panel()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.txtUserName = New System.Windows.Forms.TextBox()
        Me.txtDatabaseName = New System.Windows.Forms.TextBox()
        Me.txtServerName = New System.Windows.Forms.TextBox()
        Me.txtPassword = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.StatusStrip5 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip4 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip3 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip2 = New System.Windows.Forms.StatusStrip()
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.btnSave = New System.Windows.Forms.Button()
        Me.btnExit = New System.Windows.Forms.Button()
        Me.StatusStrip6 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip7 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip8 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip9 = New System.Windows.Forms.StatusStrip()
        Me.ToolStripStatusLabel1 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip()
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
        Me.APPLICATIONSETTINGSToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.pnlSettings.SuspendLayout()
        Me.Panel1.SuspendLayout()
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.Panel2.SuspendLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel3.SuspendLayout()
        Me.StatusStrip9.SuspendLayout()
        Me.MenuStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'pnlSettings
        '
        Me.pnlSettings.BackColor = System.Drawing.SystemColors.Control
        Me.pnlSettings.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pnlSettings.Controls.Add(Me.Panel1)
        Me.pnlSettings.Controls.Add(Me.StatusStrip1)
        Me.pnlSettings.Controls.Add(Me.MenuStrip1)
        Me.pnlSettings.Location = New System.Drawing.Point(12, 12)
        Me.pnlSettings.Name = "pnlSettings"
        Me.pnlSettings.Size = New System.Drawing.Size(468, 248)
        Me.pnlSettings.TabIndex = 0
        '
        'Panel1
        '
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel1.Controls.Add(Me.SplitContainer1)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel1.Location = New System.Drawing.Point(0, 25)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(466, 196)
        Me.Panel1.TabIndex = 6
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer1.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer1.Name = "SplitContainer1"
        Me.SplitContainer1.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.Panel2)
        Me.SplitContainer1.Panel1.Controls.Add(Me.StatusStrip5)
        Me.SplitContainer1.Panel1.Controls.Add(Me.StatusStrip4)
        Me.SplitContainer1.Panel1.Controls.Add(Me.StatusStrip3)
        Me.SplitContainer1.Panel1.Controls.Add(Me.StatusStrip2)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.Panel3)
        Me.SplitContainer1.Panel2.Controls.Add(Me.StatusStrip6)
        Me.SplitContainer1.Panel2.Controls.Add(Me.StatusStrip7)
        Me.SplitContainer1.Panel2.Controls.Add(Me.StatusStrip8)
        Me.SplitContainer1.Panel2.Controls.Add(Me.StatusStrip9)
        Me.SplitContainer1.Size = New System.Drawing.Size(464, 194)
        Me.SplitContainer1.SplitterDistance = 129
        Me.SplitContainer1.SplitterWidth = 1
        Me.SplitContainer1.TabIndex = 0
        '
        'Panel2
        '
        Me.Panel2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel2.Controls.Add(Me.txtUserName)
        Me.Panel2.Controls.Add(Me.txtDatabaseName)
        Me.Panel2.Controls.Add(Me.txtServerName)
        Me.Panel2.Controls.Add(Me.txtPassword)
        Me.Panel2.Controls.Add(Me.Label4)
        Me.Panel2.Controls.Add(Me.Label3)
        Me.Panel2.Controls.Add(Me.Label2)
        Me.Panel2.Controls.Add(Me.Label1)
        Me.Panel2.Controls.Add(Me.PictureBox1)
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel2.Location = New System.Drawing.Point(10, 10)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(444, 109)
        Me.Panel2.TabIndex = 10
        '
        'txtUserName
        '
        Me.txtUserName.Location = New System.Drawing.Point(289, 29)
        Me.txtUserName.Name = "txtUserName"
        Me.txtUserName.Size = New System.Drawing.Size(145, 20)
        Me.txtUserName.TabIndex = 8
        '
        'txtDatabaseName
        '
        Me.txtDatabaseName.Location = New System.Drawing.Point(108, 71)
        Me.txtDatabaseName.Name = "txtDatabaseName"
        Me.txtDatabaseName.Size = New System.Drawing.Size(161, 20)
        Me.txtDatabaseName.TabIndex = 7
        '
        'txtServerName
        '
        Me.txtServerName.Location = New System.Drawing.Point(108, 29)
        Me.txtServerName.Name = "txtServerName"
        Me.txtServerName.Size = New System.Drawing.Size(161, 20)
        Me.txtServerName.TabIndex = 6
        '
        'txtPassword
        '
        Me.txtPassword.Location = New System.Drawing.Point(289, 71)
        Me.txtPassword.Name = "txtPassword"
        Me.txtPassword.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.txtPassword.Size = New System.Drawing.Size(145, 20)
        Me.txtPassword.TabIndex = 5
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(286, 55)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(53, 13)
        Me.Label4.TabIndex = 4
        Me.Label4.Text = "Password"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(286, 13)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(57, 13)
        Me.Label3.TabIndex = 3
        Me.Label3.Text = "UserName"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(105, 55)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(84, 13)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "Database Name"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(105, 13)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(125, 13)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Server Name/IP Address"
        '
        'PictureBox1
        '
        Me.PictureBox1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.PictureBox1.Image = Global.School_Soft.My.Resources.Resources.Service_Manager
        Me.PictureBox1.Location = New System.Drawing.Point(2, 3)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(97, 101)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox1.TabIndex = 0
        Me.PictureBox1.TabStop = False
        '
        'StatusStrip5
        '
        Me.StatusStrip5.AutoSize = False
        Me.StatusStrip5.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip5.Dock = System.Windows.Forms.DockStyle.Right
        Me.StatusStrip5.Location = New System.Drawing.Point(454, 10)
        Me.StatusStrip5.Name = "StatusStrip5"
        Me.StatusStrip5.Size = New System.Drawing.Size(10, 109)
        Me.StatusStrip5.SizingGrip = False
        Me.StatusStrip5.TabIndex = 9
        Me.StatusStrip5.Text = "StatusStrip5"
        '
        'StatusStrip4
        '
        Me.StatusStrip4.AutoSize = False
        Me.StatusStrip4.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip4.Dock = System.Windows.Forms.DockStyle.Left
        Me.StatusStrip4.Location = New System.Drawing.Point(0, 10)
        Me.StatusStrip4.Name = "StatusStrip4"
        Me.StatusStrip4.Size = New System.Drawing.Size(10, 109)
        Me.StatusStrip4.SizingGrip = False
        Me.StatusStrip4.TabIndex = 8
        Me.StatusStrip4.Text = "StatusStrip4"
        '
        'StatusStrip3
        '
        Me.StatusStrip3.AutoSize = False
        Me.StatusStrip3.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip3.Location = New System.Drawing.Point(0, 119)
        Me.StatusStrip3.Name = "StatusStrip3"
        Me.StatusStrip3.Size = New System.Drawing.Size(464, 10)
        Me.StatusStrip3.SizingGrip = False
        Me.StatusStrip3.TabIndex = 7
        Me.StatusStrip3.Text = "StatusStrip3"
        '
        'StatusStrip2
        '
        Me.StatusStrip2.AutoSize = False
        Me.StatusStrip2.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip2.Dock = System.Windows.Forms.DockStyle.Top
        Me.StatusStrip2.Location = New System.Drawing.Point(0, 0)
        Me.StatusStrip2.Name = "StatusStrip2"
        Me.StatusStrip2.Size = New System.Drawing.Size(464, 10)
        Me.StatusStrip2.SizingGrip = False
        Me.StatusStrip2.TabIndex = 6
        Me.StatusStrip2.Text = "StatusStrip2"
        '
        'Panel3
        '
        Me.Panel3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel3.Controls.Add(Me.btnSave)
        Me.Panel3.Controls.Add(Me.btnExit)
        Me.Panel3.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel3.Location = New System.Drawing.Point(10, 20)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(444, 34)
        Me.Panel3.TabIndex = 11
        '
        'btnSave
        '
        Me.btnSave.Location = New System.Drawing.Point(239, 4)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(75, 23)
        Me.btnSave.TabIndex = 1
        Me.btnSave.Text = "Save"
        Me.btnSave.UseVisualStyleBackColor = True
        '
        'btnExit
        '
        Me.btnExit.Location = New System.Drawing.Point(359, 4)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(75, 23)
        Me.btnExit.TabIndex = 0
        Me.btnExit.Text = "Exit"
        Me.btnExit.UseVisualStyleBackColor = True
        '
        'StatusStrip6
        '
        Me.StatusStrip6.AutoSize = False
        Me.StatusStrip6.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip6.Dock = System.Windows.Forms.DockStyle.Right
        Me.StatusStrip6.Location = New System.Drawing.Point(454, 20)
        Me.StatusStrip6.Name = "StatusStrip6"
        Me.StatusStrip6.Size = New System.Drawing.Size(10, 34)
        Me.StatusStrip6.SizingGrip = False
        Me.StatusStrip6.TabIndex = 9
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
        Me.StatusStrip7.TabIndex = 8
        Me.StatusStrip7.Text = "StatusStrip7"
        '
        'StatusStrip8
        '
        Me.StatusStrip8.AutoSize = False
        Me.StatusStrip8.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip8.Location = New System.Drawing.Point(0, 54)
        Me.StatusStrip8.Name = "StatusStrip8"
        Me.StatusStrip8.Size = New System.Drawing.Size(464, 10)
        Me.StatusStrip8.SizingGrip = False
        Me.StatusStrip8.TabIndex = 7
        Me.StatusStrip8.Text = "StatusStrip8"
        '
        'StatusStrip9
        '
        Me.StatusStrip9.AutoSize = False
        Me.StatusStrip9.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip9.Dock = System.Windows.Forms.DockStyle.Top
        Me.StatusStrip9.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripStatusLabel1})
        Me.StatusStrip9.Location = New System.Drawing.Point(0, 0)
        Me.StatusStrip9.Name = "StatusStrip9"
        Me.StatusStrip9.Size = New System.Drawing.Size(464, 20)
        Me.StatusStrip9.SizingGrip = False
        Me.StatusStrip9.TabIndex = 6
        Me.StatusStrip9.Text = "StatusStrip9"
        '
        'ToolStripStatusLabel1
        '
        Me.ToolStripStatusLabel1.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ToolStripStatusLabel1.Name = "ToolStripStatusLabel1"
        Me.ToolStripStatusLabel1.Size = New System.Drawing.Size(77, 15)
        Me.ToolStripStatusLabel1.Text = "User Actions"
        '
        'StatusStrip1
        '
        Me.StatusStrip1.AutoSize = False
        Me.StatusStrip1.BackColor = System.Drawing.Color.LightSeaGreen
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 221)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(466, 25)
        Me.StatusStrip1.SizingGrip = False
        Me.StatusStrip1.TabIndex = 1
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'MenuStrip1
        '
        Me.MenuStrip1.AutoSize = False
        Me.MenuStrip1.BackColor = System.Drawing.Color.LightSeaGreen
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.APPLICATIONSETTINGSToolStripMenuItem})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Size = New System.Drawing.Size(466, 25)
        Me.MenuStrip1.TabIndex = 0
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        'APPLICATIONSETTINGSToolStripMenuItem
        '
        Me.APPLICATIONSETTINGSToolStripMenuItem.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.APPLICATIONSETTINGSToolStripMenuItem.Name = "APPLICATIONSETTINGSToolStripMenuItem"
        Me.APPLICATIONSETTINGSToolStripMenuItem.Size = New System.Drawing.Size(168, 21)
        Me.APPLICATIONSETTINGSToolStripMenuItem.Text = "APPLICATION SETTINGS"
        '
        'frmSettings
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.LightSkyBlue
        Me.ClientSize = New System.Drawing.Size(491, 271)
        Me.Controls.Add(Me.pnlSettings)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.MainMenuStrip = Me.MenuStrip1
        Me.Name = "frmSettings"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "frmSettings"
        Me.pnlSettings.ResumeLayout(False)
        Me.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.ResumeLayout(False)
        Me.Panel2.ResumeLayout(False)
        Me.Panel2.PerformLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel3.ResumeLayout(False)
        Me.StatusStrip9.ResumeLayout(False)
        Me.StatusStrip9.PerformLayout()
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents pnlSettings As System.Windows.Forms.Panel
    Friend WithEvents MenuStrip1 As System.Windows.Forms.MenuStrip
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents APPLICATIONSETTINGSToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents StatusStrip5 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip4 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip3 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip2 As System.Windows.Forms.StatusStrip
    Friend WithEvents Panel3 As System.Windows.Forms.Panel
    Friend WithEvents StatusStrip6 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip7 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip8 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip9 As System.Windows.Forms.StatusStrip
    Friend WithEvents ToolStripStatusLabel1 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents txtUserName As System.Windows.Forms.TextBox
    Friend WithEvents txtDatabaseName As System.Windows.Forms.TextBox
    Friend WithEvents txtServerName As System.Windows.Forms.TextBox
    Friend WithEvents txtPassword As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents btnExit As System.Windows.Forms.Button
End Class
