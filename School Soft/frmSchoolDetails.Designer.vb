<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class frmSchoolDetails
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
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
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.pnlSchDetails = New System.Windows.Forms.Panel()
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.txtInitials = New System.Windows.Forms.TextBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.btnBrowse = New System.Windows.Forms.Button()
        Me.txtSchAddress = New System.Windows.Forms.TextBox()
        Me.txtEmailAddress = New System.Windows.Forms.TextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.txtSchName = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.pbImage = New System.Windows.Forms.PictureBox()
        Me.txtTownName = New System.Windows.Forms.TextBox()
        Me.txtTelNumber = New System.Windows.Forms.TextBox()
        Me.StatusStrip5 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip4 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip3 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip10 = New System.Windows.Forms.StatusStrip()
        Me.Panel1 = New System.Windows.Forms.Panel()
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
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
        Me.txtSchoolMotto = New System.Windows.Forms.TextBox()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.pnlSchDetails.SuspendLayout()
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.Panel2.SuspendLayout()
        CType(Me.pbImage, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel1.SuspendLayout()
        Me.StatusStrip9.SuspendLayout()
        Me.StatusStrip2.SuspendLayout()
        Me.SuspendLayout()
        '
        'pnlSchDetails
        '
        Me.pnlSchDetails.BackColor = System.Drawing.SystemColors.Control
        Me.pnlSchDetails.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pnlSchDetails.Controls.Add(Me.SplitContainer1)
        Me.pnlSchDetails.Controls.Add(Me.StatusStrip2)
        Me.pnlSchDetails.Controls.Add(Me.StatusStrip1)
        Me.pnlSchDetails.Location = New System.Drawing.Point(12, 12)
        Me.pnlSchDetails.Name = "pnlSchDetails"
        Me.pnlSchDetails.Size = New System.Drawing.Size(728, 519)
        Me.pnlSchDetails.TabIndex = 0
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
        Me.SplitContainer1.Panel1.Controls.Add(Me.StatusStrip5)
        Me.SplitContainer1.Panel1.Controls.Add(Me.StatusStrip4)
        Me.SplitContainer1.Panel1.Controls.Add(Me.StatusStrip3)
        Me.SplitContainer1.Panel1.Controls.Add(Me.StatusStrip10)
        Me.SplitContainer1.Panel1MinSize = 100
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.Panel1)
        Me.SplitContainer1.Panel2.Controls.Add(Me.StatusStrip6)
        Me.SplitContainer1.Panel2.Controls.Add(Me.StatusStrip7)
        Me.SplitContainer1.Panel2.Controls.Add(Me.StatusStrip8)
        Me.SplitContainer1.Panel2.Controls.Add(Me.StatusStrip9)
        Me.SplitContainer1.Size = New System.Drawing.Size(726, 467)
        Me.SplitContainer1.SplitterDistance = 375
        Me.SplitContainer1.SplitterWidth = 1
        Me.SplitContainer1.TabIndex = 26
        '
        'Panel2
        '
        Me.Panel2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel2.Controls.Add(Me.txtSchoolMotto)
        Me.Panel2.Controls.Add(Me.Label8)
        Me.Panel2.Controls.Add(Me.txtInitials)
        Me.Panel2.Controls.Add(Me.Label7)
        Me.Panel2.Controls.Add(Me.Label1)
        Me.Panel2.Controls.Add(Me.Label3)
        Me.Panel2.Controls.Add(Me.btnBrowse)
        Me.Panel2.Controls.Add(Me.txtSchAddress)
        Me.Panel2.Controls.Add(Me.txtEmailAddress)
        Me.Panel2.Controls.Add(Me.Label6)
        Me.Panel2.Controls.Add(Me.Label2)
        Me.Panel2.Controls.Add(Me.Label4)
        Me.Panel2.Controls.Add(Me.txtSchName)
        Me.Panel2.Controls.Add(Me.Label5)
        Me.Panel2.Controls.Add(Me.pbImage)
        Me.Panel2.Controls.Add(Me.txtTownName)
        Me.Panel2.Controls.Add(Me.txtTelNumber)
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel2.Location = New System.Drawing.Point(10, 10)
        Me.Panel2.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(704, 353)
        Me.Panel2.TabIndex = 27
        '
        'txtInitials
        '
        Me.txtInitials.Location = New System.Drawing.Point(360, 245)
        Me.txtInitials.Name = "txtInitials"
        Me.txtInitials.Size = New System.Drawing.Size(331, 20)
        Me.txtInitials.TabIndex = 27
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Font = New System.Drawing.Font("Century Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.Location = New System.Drawing.Point(246, 245)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(51, 17)
        Me.Label7.TabIndex = 26
        Me.Label7.Text = "Initials:"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Century Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(3, 13)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(134, 17)
        Me.Label1.TabIndex = 15
        Me.Label1.Text = "Upload School Log:"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Century Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(246, 68)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(108, 17)
        Me.Label3.TabIndex = 18
        Me.Label3.Text = "School Address:"
        '
        'btnBrowse
        '
        Me.btnBrowse.Location = New System.Drawing.Point(143, 9)
        Me.btnBrowse.Name = "btnBrowse"
        Me.btnBrowse.Size = New System.Drawing.Size(77, 23)
        Me.btnBrowse.TabIndex = 16
        Me.btnBrowse.Text = "Select"
        Me.btnBrowse.UseVisualStyleBackColor = True
        '
        'txtSchAddress
        '
        Me.txtSchAddress.Location = New System.Drawing.Point(360, 67)
        Me.txtSchAddress.Name = "txtSchAddress"
        Me.txtSchAddress.Size = New System.Drawing.Size(331, 20)
        Me.txtSchAddress.TabIndex = 19
        '
        'txtEmailAddress
        '
        Me.txtEmailAddress.Location = New System.Drawing.Point(360, 203)
        Me.txtEmailAddress.Name = "txtEmailAddress"
        Me.txtEmailAddress.Size = New System.Drawing.Size(331, 20)
        Me.txtEmailAddress.TabIndex = 25
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Font = New System.Drawing.Font("Century Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.Location = New System.Drawing.Point(246, 158)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(101, 17)
        Me.Label6.TabIndex = 22
        Me.Label6.Text = "Telephone No:"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Century Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(246, 20)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(99, 17)
        Me.Label2.TabIndex = 17
        Me.Label2.Text = "School Name:"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Century Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(246, 112)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(90, 17)
        Me.Label4.TabIndex = 20
        Me.Label4.Text = "Town Name:"
        '
        'txtSchName
        '
        Me.txtSchName.Location = New System.Drawing.Point(360, 20)
        Me.txtSchName.Name = "txtSchName"
        Me.txtSchName.Size = New System.Drawing.Size(331, 20)
        Me.txtSchName.TabIndex = 14
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Century Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.Location = New System.Drawing.Point(246, 204)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(100, 17)
        Me.Label5.TabIndex = 21
        Me.Label5.Text = "Email Address:"
        '
        'pbImage
        '
        Me.pbImage.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pbImage.Location = New System.Drawing.Point(6, 41)
        Me.pbImage.Name = "pbImage"
        Me.pbImage.Size = New System.Drawing.Size(214, 203)
        Me.pbImage.TabIndex = 13
        Me.pbImage.TabStop = False
        '
        'txtTownName
        '
        Me.txtTownName.Location = New System.Drawing.Point(360, 111)
        Me.txtTownName.Name = "txtTownName"
        Me.txtTownName.Size = New System.Drawing.Size(331, 20)
        Me.txtTownName.TabIndex = 23
        '
        'txtTelNumber
        '
        Me.txtTelNumber.Location = New System.Drawing.Point(360, 157)
        Me.txtTelNumber.Name = "txtTelNumber"
        Me.txtTelNumber.Size = New System.Drawing.Size(331, 20)
        Me.txtTelNumber.TabIndex = 24
        '
        'StatusStrip5
        '
        Me.StatusStrip5.AutoSize = False
        Me.StatusStrip5.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip5.Dock = System.Windows.Forms.DockStyle.Right
        Me.StatusStrip5.ImageScalingSize = New System.Drawing.Size(20, 20)
        Me.StatusStrip5.Location = New System.Drawing.Point(714, 10)
        Me.StatusStrip5.Name = "StatusStrip5"
        Me.StatusStrip5.Size = New System.Drawing.Size(10, 353)
        Me.StatusStrip5.SizingGrip = False
        Me.StatusStrip5.TabIndex = 17
        Me.StatusStrip5.Text = "StatusStrip5"
        '
        'StatusStrip4
        '
        Me.StatusStrip4.AutoSize = False
        Me.StatusStrip4.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip4.Dock = System.Windows.Forms.DockStyle.Left
        Me.StatusStrip4.ImageScalingSize = New System.Drawing.Size(20, 20)
        Me.StatusStrip4.Location = New System.Drawing.Point(0, 10)
        Me.StatusStrip4.Name = "StatusStrip4"
        Me.StatusStrip4.Size = New System.Drawing.Size(10, 353)
        Me.StatusStrip4.SizingGrip = False
        Me.StatusStrip4.TabIndex = 16
        Me.StatusStrip4.Text = "StatusStrip4"
        '
        'StatusStrip3
        '
        Me.StatusStrip3.AutoSize = False
        Me.StatusStrip3.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip3.ImageScalingSize = New System.Drawing.Size(20, 20)
        Me.StatusStrip3.Location = New System.Drawing.Point(0, 363)
        Me.StatusStrip3.Name = "StatusStrip3"
        Me.StatusStrip3.Size = New System.Drawing.Size(724, 10)
        Me.StatusStrip3.SizingGrip = False
        Me.StatusStrip3.TabIndex = 15
        Me.StatusStrip3.Text = "StatusStrip3"
        '
        'StatusStrip10
        '
        Me.StatusStrip10.AutoSize = False
        Me.StatusStrip10.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip10.Dock = System.Windows.Forms.DockStyle.Top
        Me.StatusStrip10.ImageScalingSize = New System.Drawing.Size(20, 20)
        Me.StatusStrip10.Location = New System.Drawing.Point(0, 0)
        Me.StatusStrip10.Name = "StatusStrip10"
        Me.StatusStrip10.Size = New System.Drawing.Size(724, 10)
        Me.StatusStrip10.SizingGrip = False
        Me.StatusStrip10.TabIndex = 14
        Me.StatusStrip10.Text = "StatusStrip10"
        '
        'Panel1
        '
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel1.Controls.Add(Me.btnSave)
        Me.Panel1.Controls.Add(Me.btnClose)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Left
        Me.Panel1.Location = New System.Drawing.Point(10, 20)
        Me.Panel1.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(706, 59)
        Me.Panel1.TabIndex = 26
        '
        'btnSave
        '
        Me.btnSave.Location = New System.Drawing.Point(360, 11)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(75, 23)
        Me.btnSave.TabIndex = 26
        Me.btnSave.Text = "Update"
        Me.btnSave.UseVisualStyleBackColor = True
        '
        'btnClose
        '
        Me.btnClose.Location = New System.Drawing.Point(615, 11)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(75, 23)
        Me.btnClose.TabIndex = 27
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'StatusStrip6
        '
        Me.StatusStrip6.AutoSize = False
        Me.StatusStrip6.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip6.Dock = System.Windows.Forms.DockStyle.Right
        Me.StatusStrip6.ImageScalingSize = New System.Drawing.Size(20, 20)
        Me.StatusStrip6.Location = New System.Drawing.Point(714, 20)
        Me.StatusStrip6.Name = "StatusStrip6"
        Me.StatusStrip6.Size = New System.Drawing.Size(10, 59)
        Me.StatusStrip6.SizingGrip = False
        Me.StatusStrip6.TabIndex = 19
        Me.StatusStrip6.Text = "StatusStrip6"
        '
        'StatusStrip7
        '
        Me.StatusStrip7.AutoSize = False
        Me.StatusStrip7.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip7.Dock = System.Windows.Forms.DockStyle.Left
        Me.StatusStrip7.ImageScalingSize = New System.Drawing.Size(20, 20)
        Me.StatusStrip7.Location = New System.Drawing.Point(0, 20)
        Me.StatusStrip7.Name = "StatusStrip7"
        Me.StatusStrip7.Size = New System.Drawing.Size(10, 59)
        Me.StatusStrip7.SizingGrip = False
        Me.StatusStrip7.TabIndex = 18
        Me.StatusStrip7.Text = "StatusStrip7"
        '
        'StatusStrip8
        '
        Me.StatusStrip8.AutoSize = False
        Me.StatusStrip8.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip8.ImageScalingSize = New System.Drawing.Size(20, 20)
        Me.StatusStrip8.Location = New System.Drawing.Point(0, 79)
        Me.StatusStrip8.Name = "StatusStrip8"
        Me.StatusStrip8.Size = New System.Drawing.Size(724, 10)
        Me.StatusStrip8.SizingGrip = False
        Me.StatusStrip8.TabIndex = 17
        Me.StatusStrip8.Text = "StatusStrip8"
        '
        'StatusStrip9
        '
        Me.StatusStrip9.AutoSize = False
        Me.StatusStrip9.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip9.Dock = System.Windows.Forms.DockStyle.Top
        Me.StatusStrip9.ImageScalingSize = New System.Drawing.Size(20, 20)
        Me.StatusStrip9.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripStatusLabel2})
        Me.StatusStrip9.Location = New System.Drawing.Point(0, 0)
        Me.StatusStrip9.Name = "StatusStrip9"
        Me.StatusStrip9.Size = New System.Drawing.Size(724, 20)
        Me.StatusStrip9.SizingGrip = False
        Me.StatusStrip9.TabIndex = 16
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
        Me.StatusStrip2.ImageScalingSize = New System.Drawing.Size(20, 20)
        Me.StatusStrip2.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripStatusLabel1})
        Me.StatusStrip2.Location = New System.Drawing.Point(0, 0)
        Me.StatusStrip2.Name = "StatusStrip2"
        Me.StatusStrip2.Size = New System.Drawing.Size(726, 25)
        Me.StatusStrip2.SizingGrip = False
        Me.StatusStrip2.TabIndex = 12
        Me.StatusStrip2.Text = "StatusStrip2"
        '
        'ToolStripStatusLabel1
        '
        Me.ToolStripStatusLabel1.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ToolStripStatusLabel1.Name = "ToolStripStatusLabel1"
        Me.ToolStripStatusLabel1.Size = New System.Drawing.Size(115, 20)
        Me.ToolStripStatusLabel1.Text = "SCHOOL DETAILS"
        '
        'StatusStrip1
        '
        Me.StatusStrip1.AutoSize = False
        Me.StatusStrip1.BackColor = System.Drawing.Color.LightSeaGreen
        Me.StatusStrip1.ImageScalingSize = New System.Drawing.Size(20, 20)
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 492)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(726, 25)
        Me.StatusStrip1.SizingGrip = False
        Me.StatusStrip1.TabIndex = 11
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.FileName = "OpenFileDialog1"
        '
        'txtSchoolMotto
        '
        Me.txtSchoolMotto.Location = New System.Drawing.Point(360, 284)
        Me.txtSchoolMotto.Multiline = True
        Me.txtSchoolMotto.Name = "txtSchoolMotto"
        Me.txtSchoolMotto.Size = New System.Drawing.Size(331, 49)
        Me.txtSchoolMotto.TabIndex = 29
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Font = New System.Drawing.Font("Century Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label8.Location = New System.Drawing.Point(246, 284)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(102, 17)
        Me.Label8.TabIndex = 28
        Me.Label8.Text = "School Motto: "
        '
        'frmSchoolDetails
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.LightBlue
        Me.ClientSize = New System.Drawing.Size(750, 541)
        Me.Controls.Add(Me.pnlSchDetails)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Name = "frmSchoolDetails"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "School Details"
        Me.pnlSchDetails.ResumeLayout(False)
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.ResumeLayout(False)
        Me.Panel2.ResumeLayout(False)
        Me.Panel2.PerformLayout()
        CType(Me.pbImage, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel1.ResumeLayout(False)
        Me.StatusStrip9.ResumeLayout(False)
        Me.StatusStrip9.PerformLayout()
        Me.StatusStrip2.ResumeLayout(False)
        Me.StatusStrip2.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents pnlSchDetails As System.Windows.Forms.Panel
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents txtEmailAddress As System.Windows.Forms.TextBox
    Friend WithEvents txtTelNumber As System.Windows.Forms.TextBox
    Friend WithEvents txtTownName As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents txtSchAddress As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents btnBrowse As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtSchName As System.Windows.Forms.TextBox
    Friend WithEvents pbImage As System.Windows.Forms.PictureBox
    Friend WithEvents StatusStrip2 As System.Windows.Forms.StatusStrip
    Friend WithEvents ToolStripStatusLabel1 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip5 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip4 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip3 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip10 As System.Windows.Forms.StatusStrip
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents StatusStrip6 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip7 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip8 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip9 As System.Windows.Forms.StatusStrip
    Friend WithEvents ToolStripStatusLabel2 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents Panel2 As Panel
    Friend WithEvents Panel1 As Panel
    Friend WithEvents txtInitials As TextBox
    Friend WithEvents Label7 As Label
    Friend WithEvents txtSchoolMotto As TextBox
    Friend WithEvents Label8 As Label
End Class
