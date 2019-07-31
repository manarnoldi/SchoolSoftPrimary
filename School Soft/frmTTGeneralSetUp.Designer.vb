<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmTTGeneralSetUp
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
        Me.pnlTTGeneralSetUp = New System.Windows.Forms.Panel()
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.txtDaysRange = New System.Windows.Forms.TextBox()
        Me.cboAcadYear = New System.Windows.Forms.ComboBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.cboLessonDay = New System.Windows.Forms.ComboBox()
        Me.cboDayWeek = New System.Windows.Forms.ComboBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.StatusStrip12 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip13 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip14 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip5 = New System.Windows.Forms.StatusStrip()
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.btnUpdate = New System.Windows.Forms.Button()
        Me.btnClose = New System.Windows.Forms.Button()
        Me.btnSave = New System.Windows.Forms.Button()
        Me.StatusStrip3 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip4 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip6 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip7 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip2 = New System.Windows.Forms.StatusStrip()
        Me.ToolStripStatusLabel1 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.cboBreaksPerDay = New System.Windows.Forms.ComboBox()
        Me.pnlTTGeneralSetUp.SuspendLayout()
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.Panel3.SuspendLayout()
        Me.StatusStrip2.SuspendLayout()
        Me.SuspendLayout()
        '
        'pnlTTGeneralSetUp
        '
        Me.pnlTTGeneralSetUp.BackColor = System.Drawing.SystemColors.Control
        Me.pnlTTGeneralSetUp.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pnlTTGeneralSetUp.Controls.Add(Me.SplitContainer1)
        Me.pnlTTGeneralSetUp.Controls.Add(Me.StatusStrip2)
        Me.pnlTTGeneralSetUp.Controls.Add(Me.StatusStrip1)
        Me.pnlTTGeneralSetUp.Location = New System.Drawing.Point(12, 12)
        Me.pnlTTGeneralSetUp.Name = "pnlTTGeneralSetUp"
        Me.pnlTTGeneralSetUp.Size = New System.Drawing.Size(364, 348)
        Me.pnlTTGeneralSetUp.TabIndex = 0
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
        Me.SplitContainer1.Panel1.Controls.Add(Me.Panel1)
        Me.SplitContainer1.Panel1.Controls.Add(Me.StatusStrip12)
        Me.SplitContainer1.Panel1.Controls.Add(Me.StatusStrip13)
        Me.SplitContainer1.Panel1.Controls.Add(Me.StatusStrip14)
        Me.SplitContainer1.Panel1.Controls.Add(Me.StatusStrip5)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.Panel3)
        Me.SplitContainer1.Panel2.Controls.Add(Me.StatusStrip3)
        Me.SplitContainer1.Panel2.Controls.Add(Me.StatusStrip4)
        Me.SplitContainer1.Panel2.Controls.Add(Me.StatusStrip6)
        Me.SplitContainer1.Panel2.Controls.Add(Me.StatusStrip7)
        Me.SplitContainer1.Size = New System.Drawing.Size(362, 296)
        Me.SplitContainer1.SplitterDistance = 218
        Me.SplitContainer1.SplitterWidth = 1
        Me.SplitContainer1.TabIndex = 15
        '
        'Panel1
        '
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel1.Controls.Add(Me.Label4)
        Me.Panel1.Controls.Add(Me.cboBreaksPerDay)
        Me.Panel1.Controls.Add(Me.Label3)
        Me.Panel1.Controls.Add(Me.txtDaysRange)
        Me.Panel1.Controls.Add(Me.cboAcadYear)
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Controls.Add(Me.Label6)
        Me.Panel1.Controls.Add(Me.cboLessonDay)
        Me.Panel1.Controls.Add(Me.cboDayWeek)
        Me.Panel1.Controls.Add(Me.Label2)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel1.Location = New System.Drawing.Point(5, 5)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(350, 206)
        Me.Panel1.TabIndex = 73
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(16, 163)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(69, 13)
        Me.Label3.TabIndex = 34
        Me.Label3.Text = "Days Range:"
        '
        'txtDaysRange
        '
        Me.txtDaysRange.Location = New System.Drawing.Point(102, 160)
        Me.txtDaysRange.Name = "txtDaysRange"
        Me.txtDaysRange.ReadOnly = True
        Me.txtDaysRange.Size = New System.Drawing.Size(224, 20)
        Me.txtDaysRange.TabIndex = 33
        '
        'cboAcadYear
        '
        Me.cboAcadYear.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboAcadYear.FormattingEnabled = True
        Me.cboAcadYear.Location = New System.Drawing.Point(102, 18)
        Me.cboAcadYear.Name = "cboAcadYear"
        Me.cboAcadYear.Size = New System.Drawing.Size(224, 21)
        Me.cboAcadYear.TabIndex = 32
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(16, 21)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(82, 13)
        Me.Label1.TabIndex = 15
        Me.Label1.Text = "Academic Year:"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(16, 95)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(77, 13)
        Me.Label6.TabIndex = 28
        Me.Label6.Text = "Days / Week :"
        '
        'cboLessonDay
        '
        Me.cboLessonDay.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboLessonDay.FormattingEnabled = True
        Me.cboLessonDay.Items.AddRange(New Object() {"7", "8", "9", "10", "11", "12", "13", "14"})
        Me.cboLessonDay.Location = New System.Drawing.Point(102, 56)
        Me.cboLessonDay.Name = "cboLessonDay"
        Me.cboLessonDay.Size = New System.Drawing.Size(224, 21)
        Me.cboLessonDay.TabIndex = 24
        '
        'cboDayWeek
        '
        Me.cboDayWeek.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboDayWeek.FormattingEnabled = True
        Me.cboDayWeek.Items.AddRange(New Object() {"5", "6", "7"})
        Me.cboDayWeek.Location = New System.Drawing.Point(102, 92)
        Me.cboDayWeek.Name = "cboDayWeek"
        Me.cboDayWeek.Size = New System.Drawing.Size(224, 21)
        Me.cboDayWeek.TabIndex = 29
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(16, 59)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(82, 13)
        Me.Label2.TabIndex = 23
        Me.Label2.Text = "Lessons / Day :"
        '
        'StatusStrip12
        '
        Me.StatusStrip12.AutoSize = False
        Me.StatusStrip12.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip12.Dock = System.Windows.Forms.DockStyle.Left
        Me.StatusStrip12.Location = New System.Drawing.Point(0, 5)
        Me.StatusStrip12.Name = "StatusStrip12"
        Me.StatusStrip12.Size = New System.Drawing.Size(5, 206)
        Me.StatusStrip12.SizingGrip = False
        Me.StatusStrip12.TabIndex = 72
        Me.StatusStrip12.Text = "StatusStrip12"
        '
        'StatusStrip13
        '
        Me.StatusStrip13.AutoSize = False
        Me.StatusStrip13.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip13.Location = New System.Drawing.Point(0, 211)
        Me.StatusStrip13.Name = "StatusStrip13"
        Me.StatusStrip13.Size = New System.Drawing.Size(355, 5)
        Me.StatusStrip13.SizingGrip = False
        Me.StatusStrip13.TabIndex = 71
        Me.StatusStrip13.Text = "StatusStrip13"
        '
        'StatusStrip14
        '
        Me.StatusStrip14.AutoSize = False
        Me.StatusStrip14.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip14.Dock = System.Windows.Forms.DockStyle.Top
        Me.StatusStrip14.Location = New System.Drawing.Point(0, 0)
        Me.StatusStrip14.Name = "StatusStrip14"
        Me.StatusStrip14.Size = New System.Drawing.Size(355, 5)
        Me.StatusStrip14.SizingGrip = False
        Me.StatusStrip14.TabIndex = 70
        Me.StatusStrip14.Text = "StatusStrip14"
        '
        'StatusStrip5
        '
        Me.StatusStrip5.AutoSize = False
        Me.StatusStrip5.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip5.Dock = System.Windows.Forms.DockStyle.Right
        Me.StatusStrip5.Location = New System.Drawing.Point(355, 0)
        Me.StatusStrip5.Name = "StatusStrip5"
        Me.StatusStrip5.Size = New System.Drawing.Size(5, 216)
        Me.StatusStrip5.SizingGrip = False
        Me.StatusStrip5.TabIndex = 69
        Me.StatusStrip5.Text = "StatusStrip5"
        '
        'Panel3
        '
        Me.Panel3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel3.Controls.Add(Me.btnUpdate)
        Me.Panel3.Controls.Add(Me.btnClose)
        Me.Panel3.Controls.Add(Me.btnSave)
        Me.Panel3.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel3.Location = New System.Drawing.Point(5, 5)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(350, 65)
        Me.Panel3.TabIndex = 74
        '
        'btnUpdate
        '
        Me.btnUpdate.Location = New System.Drawing.Point(134, 12)
        Me.btnUpdate.Name = "btnUpdate"
        Me.btnUpdate.Size = New System.Drawing.Size(75, 23)
        Me.btnUpdate.TabIndex = 2
        Me.btnUpdate.Text = "Update"
        Me.btnUpdate.UseVisualStyleBackColor = True
        '
        'btnClose
        '
        Me.btnClose.Location = New System.Drawing.Point(251, 12)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(75, 23)
        Me.btnClose.TabIndex = 1
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'btnSave
        '
        Me.btnSave.Location = New System.Drawing.Point(19, 12)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(75, 23)
        Me.btnSave.TabIndex = 0
        Me.btnSave.Text = "Save"
        Me.btnSave.UseVisualStyleBackColor = True
        '
        'StatusStrip3
        '
        Me.StatusStrip3.AutoSize = False
        Me.StatusStrip3.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip3.Dock = System.Windows.Forms.DockStyle.Left
        Me.StatusStrip3.Location = New System.Drawing.Point(0, 5)
        Me.StatusStrip3.Name = "StatusStrip3"
        Me.StatusStrip3.Size = New System.Drawing.Size(5, 65)
        Me.StatusStrip3.SizingGrip = False
        Me.StatusStrip3.TabIndex = 72
        Me.StatusStrip3.Text = "StatusStrip3"
        '
        'StatusStrip4
        '
        Me.StatusStrip4.AutoSize = False
        Me.StatusStrip4.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip4.Location = New System.Drawing.Point(0, 70)
        Me.StatusStrip4.Name = "StatusStrip4"
        Me.StatusStrip4.Size = New System.Drawing.Size(355, 5)
        Me.StatusStrip4.SizingGrip = False
        Me.StatusStrip4.TabIndex = 71
        Me.StatusStrip4.Text = "StatusStrip4"
        '
        'StatusStrip6
        '
        Me.StatusStrip6.AutoSize = False
        Me.StatusStrip6.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip6.Dock = System.Windows.Forms.DockStyle.Top
        Me.StatusStrip6.Location = New System.Drawing.Point(0, 0)
        Me.StatusStrip6.Name = "StatusStrip6"
        Me.StatusStrip6.Size = New System.Drawing.Size(355, 5)
        Me.StatusStrip6.SizingGrip = False
        Me.StatusStrip6.TabIndex = 70
        Me.StatusStrip6.Text = "StatusStrip6"
        '
        'StatusStrip7
        '
        Me.StatusStrip7.AutoSize = False
        Me.StatusStrip7.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip7.Dock = System.Windows.Forms.DockStyle.Right
        Me.StatusStrip7.Location = New System.Drawing.Point(355, 0)
        Me.StatusStrip7.Name = "StatusStrip7"
        Me.StatusStrip7.Size = New System.Drawing.Size(5, 75)
        Me.StatusStrip7.SizingGrip = False
        Me.StatusStrip7.TabIndex = 69
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
        Me.StatusStrip2.Size = New System.Drawing.Size(362, 25)
        Me.StatusStrip2.SizingGrip = False
        Me.StatusStrip2.TabIndex = 14
        Me.StatusStrip2.Text = "StatusStrip2"
        '
        'ToolStripStatusLabel1
        '
        Me.ToolStripStatusLabel1.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ToolStripStatusLabel1.Name = "ToolStripStatusLabel1"
        Me.ToolStripStatusLabel1.Size = New System.Drawing.Size(181, 20)
        Me.ToolStripStatusLabel1.Text = "TIMETABLE GENERAL SETUP"
        '
        'StatusStrip1
        '
        Me.StatusStrip1.AutoSize = False
        Me.StatusStrip1.BackColor = System.Drawing.Color.LightSeaGreen
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 321)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(362, 25)
        Me.StatusStrip1.SizingGrip = False
        Me.StatusStrip1.TabIndex = 13
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(16, 130)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(87, 13)
        Me.Label4.TabIndex = 35
        Me.Label4.Text = "Breaks Per Day :"
        '
        'cboBreaksPerDay
        '
        Me.cboBreaksPerDay.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboBreaksPerDay.FormattingEnabled = True
        Me.cboBreaksPerDay.Items.AddRange(New Object() {"1", "2", "3", "4", "5"})
        Me.cboBreaksPerDay.Location = New System.Drawing.Point(109, 127)
        Me.cboBreaksPerDay.Name = "cboBreaksPerDay"
        Me.cboBreaksPerDay.Size = New System.Drawing.Size(217, 21)
        Me.cboBreaksPerDay.TabIndex = 36
        '
        'frmTTGeneralSetUp
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.LightSkyBlue
        Me.ClientSize = New System.Drawing.Size(388, 390)
        Me.Controls.Add(Me.pnlTTGeneralSetUp)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Name = "frmTTGeneralSetUp"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = " General SetUp"
        Me.pnlTTGeneralSetUp.ResumeLayout(False)
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.ResumeLayout(False)
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.Panel3.ResumeLayout(False)
        Me.StatusStrip2.ResumeLayout(False)
        Me.StatusStrip2.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents pnlTTGeneralSetUp As System.Windows.Forms.Panel
    Friend WithEvents StatusStrip2 As System.Windows.Forms.StatusStrip
    Friend WithEvents ToolStripStatusLabel1 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents cboAcadYear As System.Windows.Forms.ComboBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents cboDayWeek As System.Windows.Forms.ComboBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents cboLessonDay As System.Windows.Forms.ComboBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents StatusStrip12 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip13 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip14 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip5 As System.Windows.Forms.StatusStrip
    Friend WithEvents Panel3 As System.Windows.Forms.Panel
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents StatusStrip3 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip4 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip6 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip7 As System.Windows.Forms.StatusStrip
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents txtDaysRange As System.Windows.Forms.TextBox
    Friend WithEvents btnUpdate As System.Windows.Forms.Button
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents cboBreaksPerDay As System.Windows.Forms.ComboBox
End Class
