<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmTTStaffSetUp
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
        Me.pnlStaffSetup = New System.Windows.Forms.Panel()
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.cboAcadYear = New System.Windows.Forms.ComboBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.txtStaffName = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.cboTeacherNo = New System.Windows.Forms.ComboBox()
        Me.dgStaffSetUp = New System.Windows.Forms.DataGridView()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.StatusStrip5 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip8 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip9 = New System.Windows.Forms.StatusStrip()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.btnClose = New System.Windows.Forms.Button()
        Me.btnUpdate = New System.Windows.Forms.Button()
        Me.StatusStrip3 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip4 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip6 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip7 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip2 = New System.Windows.Forms.StatusStrip()
        Me.ToolStripStatusLabel1 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip()
        Me.pnlStaffSetup.SuspendLayout()
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.Panel2.SuspendLayout()
        CType(Me.dgStaffSetUp, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel1.SuspendLayout()
        Me.StatusStrip2.SuspendLayout()
        Me.SuspendLayout()
        '
        'pnlStaffSetup
        '
        Me.pnlStaffSetup.BackColor = System.Drawing.SystemColors.Control
        Me.pnlStaffSetup.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pnlStaffSetup.Controls.Add(Me.SplitContainer1)
        Me.pnlStaffSetup.Controls.Add(Me.StatusStrip2)
        Me.pnlStaffSetup.Controls.Add(Me.StatusStrip1)
        Me.pnlStaffSetup.Location = New System.Drawing.Point(10, 9)
        Me.pnlStaffSetup.Name = "pnlStaffSetup"
        Me.pnlStaffSetup.Size = New System.Drawing.Size(430, 448)
        Me.pnlStaffSetup.TabIndex = 2
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
        Me.SplitContainer1.Panel1.Controls.Add(Me.StatusStrip8)
        Me.SplitContainer1.Panel1.Controls.Add(Me.StatusStrip9)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.Panel1)
        Me.SplitContainer1.Panel2.Controls.Add(Me.StatusStrip3)
        Me.SplitContainer1.Panel2.Controls.Add(Me.StatusStrip4)
        Me.SplitContainer1.Panel2.Controls.Add(Me.StatusStrip6)
        Me.SplitContainer1.Panel2.Controls.Add(Me.StatusStrip7)
        Me.SplitContainer1.Size = New System.Drawing.Size(428, 396)
        Me.SplitContainer1.SplitterDistance = 344
        Me.SplitContainer1.SplitterWidth = 1
        Me.SplitContainer1.TabIndex = 17
        '
        'Panel2
        '
        Me.Panel2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel2.Controls.Add(Me.cboAcadYear)
        Me.Panel2.Controls.Add(Me.Label4)
        Me.Panel2.Controls.Add(Me.txtStaffName)
        Me.Panel2.Controls.Add(Me.Label3)
        Me.Panel2.Controls.Add(Me.cboTeacherNo)
        Me.Panel2.Controls.Add(Me.dgStaffSetUp)
        Me.Panel2.Controls.Add(Me.Label1)
        Me.Panel2.Controls.Add(Me.Label2)
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel2.Location = New System.Drawing.Point(0, 5)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(421, 332)
        Me.Panel2.TabIndex = 81
        '
        'cboAcadYear
        '
        Me.cboAcadYear.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboAcadYear.FormattingEnabled = True
        Me.cboAcadYear.Location = New System.Drawing.Point(292, 11)
        Me.cboAcadYear.Name = "cboAcadYear"
        Me.cboAcadYear.Size = New System.Drawing.Size(115, 21)
        Me.cboAcadYear.TabIndex = 20
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(226, 14)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(60, 13)
        Me.Label4.TabIndex = 19
        Me.Label4.Text = "Acad Year:"
        '
        'txtStaffName
        '
        Me.txtStaffName.Location = New System.Drawing.Point(93, 44)
        Me.txtStaffName.Name = "txtStaffName"
        Me.txtStaffName.ReadOnly = True
        Me.txtStaffName.Size = New System.Drawing.Size(314, 20)
        Me.txtStaffName.TabIndex = 18
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(1, 75)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(87, 13)
        Me.Label3.TabIndex = 17
        Me.Label3.Text = "Staff Constraints:"
        '
        'cboTeacherNo
        '
        Me.cboTeacherNo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboTeacherNo.FormattingEnabled = True
        Me.cboTeacherNo.Location = New System.Drawing.Point(93, 11)
        Me.cboTeacherNo.Name = "cboTeacherNo"
        Me.cboTeacherNo.Size = New System.Drawing.Size(115, 21)
        Me.cboTeacherNo.TabIndex = 11
        '
        'dgStaffSetUp
        '
        Me.dgStaffSetUp.AllowUserToAddRows = False
        Me.dgStaffSetUp.AllowUserToDeleteRows = False
        Me.dgStaffSetUp.AllowUserToResizeColumns = False
        Me.dgStaffSetUp.AllowUserToResizeRows = False
        Me.dgStaffSetUp.BackgroundColor = System.Drawing.SystemColors.ControlLightLight
        Me.dgStaffSetUp.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgStaffSetUp.Location = New System.Drawing.Point(4, 102)
        Me.dgStaffSetUp.Name = "dgStaffSetUp"
        Me.dgStaffSetUp.RowHeadersWidth = 4
        Me.dgStaffSetUp.ShowCellErrors = False
        Me.dgStaffSetUp.ShowCellToolTips = False
        Me.dgStaffSetUp.ShowEditingIcon = False
        Me.dgStaffSetUp.ShowRowErrors = False
        Me.dgStaffSetUp.Size = New System.Drawing.Size(412, 225)
        Me.dgStaffSetUp.TabIndex = 8
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(15, 47)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(63, 13)
        Me.Label1.TabIndex = 5
        Me.Label1.Text = "Staff Name:"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(15, 14)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(72, 13)
        Me.Label2.TabIndex = 6
        Me.Label2.Text = "Staff Number:"
        '
        'StatusStrip5
        '
        Me.StatusStrip5.AutoSize = False
        Me.StatusStrip5.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip5.Location = New System.Drawing.Point(0, 337)
        Me.StatusStrip5.Name = "StatusStrip5"
        Me.StatusStrip5.Size = New System.Drawing.Size(421, 5)
        Me.StatusStrip5.SizingGrip = False
        Me.StatusStrip5.TabIndex = 80
        Me.StatusStrip5.Text = "StatusStrip5"
        '
        'StatusStrip8
        '
        Me.StatusStrip8.AutoSize = False
        Me.StatusStrip8.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip8.Dock = System.Windows.Forms.DockStyle.Top
        Me.StatusStrip8.Location = New System.Drawing.Point(0, 0)
        Me.StatusStrip8.Name = "StatusStrip8"
        Me.StatusStrip8.Size = New System.Drawing.Size(421, 5)
        Me.StatusStrip8.SizingGrip = False
        Me.StatusStrip8.TabIndex = 79
        Me.StatusStrip8.Text = "StatusStrip8"
        '
        'StatusStrip9
        '
        Me.StatusStrip9.AutoSize = False
        Me.StatusStrip9.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip9.Dock = System.Windows.Forms.DockStyle.Right
        Me.StatusStrip9.Location = New System.Drawing.Point(421, 0)
        Me.StatusStrip9.Name = "StatusStrip9"
        Me.StatusStrip9.Size = New System.Drawing.Size(5, 342)
        Me.StatusStrip9.SizingGrip = False
        Me.StatusStrip9.TabIndex = 78
        Me.StatusStrip9.Text = "StatusStrip9"
        '
        'Panel1
        '
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel1.Controls.Add(Me.btnClose)
        Me.Panel1.Controls.Add(Me.btnUpdate)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel1.Location = New System.Drawing.Point(5, 5)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(416, 39)
        Me.Panel1.TabIndex = 77
        '
        'btnClose
        '
        Me.btnClose.Location = New System.Drawing.Point(338, 7)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(75, 23)
        Me.btnClose.TabIndex = 58
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'btnUpdate
        '
        Me.btnUpdate.Location = New System.Drawing.Point(128, 7)
        Me.btnUpdate.Name = "btnUpdate"
        Me.btnUpdate.Size = New System.Drawing.Size(75, 23)
        Me.btnUpdate.TabIndex = 57
        Me.btnUpdate.Text = "Update"
        Me.btnUpdate.UseVisualStyleBackColor = True
        '
        'StatusStrip3
        '
        Me.StatusStrip3.AutoSize = False
        Me.StatusStrip3.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip3.Dock = System.Windows.Forms.DockStyle.Left
        Me.StatusStrip3.Location = New System.Drawing.Point(0, 5)
        Me.StatusStrip3.Name = "StatusStrip3"
        Me.StatusStrip3.Size = New System.Drawing.Size(5, 39)
        Me.StatusStrip3.SizingGrip = False
        Me.StatusStrip3.TabIndex = 76
        Me.StatusStrip3.Text = "StatusStrip3"
        '
        'StatusStrip4
        '
        Me.StatusStrip4.AutoSize = False
        Me.StatusStrip4.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip4.Location = New System.Drawing.Point(0, 44)
        Me.StatusStrip4.Name = "StatusStrip4"
        Me.StatusStrip4.Size = New System.Drawing.Size(421, 5)
        Me.StatusStrip4.SizingGrip = False
        Me.StatusStrip4.TabIndex = 75
        Me.StatusStrip4.Text = "StatusStrip4"
        '
        'StatusStrip6
        '
        Me.StatusStrip6.AutoSize = False
        Me.StatusStrip6.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip6.Dock = System.Windows.Forms.DockStyle.Top
        Me.StatusStrip6.Location = New System.Drawing.Point(0, 0)
        Me.StatusStrip6.Name = "StatusStrip6"
        Me.StatusStrip6.Size = New System.Drawing.Size(421, 5)
        Me.StatusStrip6.SizingGrip = False
        Me.StatusStrip6.TabIndex = 74
        Me.StatusStrip6.Text = "StatusStrip6"
        '
        'StatusStrip7
        '
        Me.StatusStrip7.AutoSize = False
        Me.StatusStrip7.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip7.Dock = System.Windows.Forms.DockStyle.Right
        Me.StatusStrip7.Location = New System.Drawing.Point(421, 0)
        Me.StatusStrip7.Name = "StatusStrip7"
        Me.StatusStrip7.Size = New System.Drawing.Size(5, 49)
        Me.StatusStrip7.SizingGrip = False
        Me.StatusStrip7.TabIndex = 73
        Me.StatusStrip7.Text = "StatusStrip7"
        '
        'StatusStrip2
        '
        Me.StatusStrip2.AutoSize = False
        Me.StatusStrip2.BackColor = System.Drawing.Color.LightSeaGreen
        Me.StatusStrip2.Dock = System.Windows.Forms.DockStyle.Top
        Me.StatusStrip2.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripStatusLabel1})
        Me.StatusStrip2.Location = New System.Drawing.Point(0, 0)
        Me.StatusStrip2.Name = "StatusStrip2"
        Me.StatusStrip2.Size = New System.Drawing.Size(428, 25)
        Me.StatusStrip2.SizingGrip = False
        Me.StatusStrip2.TabIndex = 16
        Me.StatusStrip2.Text = "StatusStrip2"
        '
        'ToolStripStatusLabel1
        '
        Me.ToolStripStatusLabel1.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ToolStripStatusLabel1.Name = "ToolStripStatusLabel1"
        Me.ToolStripStatusLabel1.Size = New System.Drawing.Size(161, 20)
        Me.ToolStripStatusLabel1.Text = "TIMETABLE STAFF SETUP"
        '
        'StatusStrip1
        '
        Me.StatusStrip1.AutoSize = False
        Me.StatusStrip1.BackColor = System.Drawing.Color.LightSeaGreen
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 421)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(428, 25)
        Me.StatusStrip1.SizingGrip = False
        Me.StatusStrip1.TabIndex = 15
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'frmTTStaffSetUp
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.LightSkyBlue
        Me.ClientSize = New System.Drawing.Size(451, 469)
        Me.Controls.Add(Me.pnlStaffSetup)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Name = "frmTTStaffSetUp"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Teacher SetUp"
        Me.pnlStaffSetup.ResumeLayout(False)
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.ResumeLayout(False)
        Me.Panel2.ResumeLayout(False)
        Me.Panel2.PerformLayout()
        CType(Me.dgStaffSetUp, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel1.ResumeLayout(False)
        Me.StatusStrip2.ResumeLayout(False)
        Me.StatusStrip2.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents pnlStaffSetup As System.Windows.Forms.Panel
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents cboTeacherNo As System.Windows.Forms.ComboBox
    Friend WithEvents dgStaffSetUp As System.Windows.Forms.DataGridView
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents StatusStrip5 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip8 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip9 As System.Windows.Forms.StatusStrip
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents btnUpdate As System.Windows.Forms.Button
    Friend WithEvents StatusStrip3 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip4 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip6 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip7 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip2 As System.Windows.Forms.StatusStrip
    Friend WithEvents ToolStripStatusLabel1 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents txtStaffName As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents cboAcadYear As System.Windows.Forms.ComboBox
End Class
