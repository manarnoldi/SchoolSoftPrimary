<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmTTPeriodSetUp
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
        Me.pnlPeriodSetup = New System.Windows.Forms.Panel()
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.cboPeriodNo = New System.Windows.Forms.ComboBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.cboPeriodType = New System.Windows.Forms.ComboBox()
        Me.btnAdd = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.cboYear = New System.Windows.Forms.ComboBox()
        Me.dtpEnd = New System.Windows.Forms.DateTimePicker()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.dtpStart = New System.Windows.Forms.DateTimePicker()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.cboPeriodName = New System.Windows.Forms.ComboBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.lstPeriodSetUp = New System.Windows.Forms.ListView()
        Me.ColumnHeader5 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader1 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader2 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader3 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader4 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.StatusStrip5 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip8 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip9 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip10 = New System.Windows.Forms.StatusStrip()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.btnRemove = New System.Windows.Forms.Button()
        Me.btnView = New System.Windows.Forms.Button()
        Me.btnClose = New System.Windows.Forms.Button()
        Me.btnUpdate = New System.Windows.Forms.Button()
        Me.StatusStrip3 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip4 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip6 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip7 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip2 = New System.Windows.Forms.StatusStrip()
        Me.ToolStripStatusLabel1 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip()
        Me.pnlPeriodSetup.SuspendLayout()
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.StatusStrip2.SuspendLayout()
        Me.SuspendLayout()
        '
        'pnlPeriodSetup
        '
        Me.pnlPeriodSetup.BackColor = System.Drawing.SystemColors.Control
        Me.pnlPeriodSetup.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pnlPeriodSetup.Controls.Add(Me.SplitContainer1)
        Me.pnlPeriodSetup.Controls.Add(Me.StatusStrip2)
        Me.pnlPeriodSetup.Controls.Add(Me.StatusStrip1)
        Me.pnlPeriodSetup.Location = New System.Drawing.Point(12, 12)
        Me.pnlPeriodSetup.Name = "pnlPeriodSetup"
        Me.pnlPeriodSetup.Size = New System.Drawing.Size(492, 502)
        Me.pnlPeriodSetup.TabIndex = 0
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
        Me.SplitContainer1.Panel1.Controls.Add(Me.StatusStrip10)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.Panel1)
        Me.SplitContainer1.Panel2.Controls.Add(Me.StatusStrip3)
        Me.SplitContainer1.Panel2.Controls.Add(Me.StatusStrip4)
        Me.SplitContainer1.Panel2.Controls.Add(Me.StatusStrip6)
        Me.SplitContainer1.Panel2.Controls.Add(Me.StatusStrip7)
        Me.SplitContainer1.Size = New System.Drawing.Size(490, 450)
        Me.SplitContainer1.SplitterDistance = 390
        Me.SplitContainer1.SplitterWidth = 1
        Me.SplitContainer1.TabIndex = 17
        '
        'Panel2
        '
        Me.Panel2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel2.Controls.Add(Me.cboPeriodNo)
        Me.Panel2.Controls.Add(Me.Label2)
        Me.Panel2.Controls.Add(Me.cboPeriodType)
        Me.Panel2.Controls.Add(Me.btnAdd)
        Me.Panel2.Controls.Add(Me.Label1)
        Me.Panel2.Controls.Add(Me.cboYear)
        Me.Panel2.Controls.Add(Me.dtpEnd)
        Me.Panel2.Controls.Add(Me.Label5)
        Me.Panel2.Controls.Add(Me.Label3)
        Me.Panel2.Controls.Add(Me.dtpStart)
        Me.Panel2.Controls.Add(Me.Label4)
        Me.Panel2.Controls.Add(Me.cboPeriodName)
        Me.Panel2.Controls.Add(Me.Label7)
        Me.Panel2.Controls.Add(Me.lstPeriodSetUp)
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel2.Location = New System.Drawing.Point(5, 5)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(478, 378)
        Me.Panel2.TabIndex = 77
        '
        'cboPeriodNo
        '
        Me.cboPeriodNo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboPeriodNo.FormattingEnabled = True
        Me.cboPeriodNo.Location = New System.Drawing.Point(312, 69)
        Me.cboPeriodNo.Name = "cboPeriodNo"
        Me.cboPeriodNo.Size = New System.Drawing.Size(59, 21)
        Me.cboPeriodNo.TabIndex = 67
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(235, 75)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(57, 13)
        Me.Label2.TabIndex = 66
        Me.Label2.Text = "Period No:"
        '
        'cboPeriodType
        '
        Me.cboPeriodType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboPeriodType.FormattingEnabled = True
        Me.cboPeriodType.Items.AddRange(New Object() {"Lesson", "Break", "Lunch"})
        Me.cboPeriodType.Location = New System.Drawing.Point(312, 9)
        Me.cboPeriodType.Name = "cboPeriodType"
        Me.cboPeriodType.Size = New System.Drawing.Size(141, 21)
        Me.cboPeriodType.TabIndex = 63
        '
        'btnAdd
        '
        Me.btnAdd.Location = New System.Drawing.Point(386, 68)
        Me.btnAdd.Name = "btnAdd"
        Me.btnAdd.Size = New System.Drawing.Size(67, 23)
        Me.btnAdd.TabIndex = 61
        Me.btnAdd.Text = "Add"
        Me.btnAdd.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(36, 12)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(71, 13)
        Me.Label1.TabIndex = 64
        Me.Label1.Text = "School Year :"
        '
        'cboYear
        '
        Me.cboYear.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboYear.FormattingEnabled = True
        Me.cboYear.Location = New System.Drawing.Point(117, 9)
        Me.cboYear.Name = "cboYear"
        Me.cboYear.Size = New System.Drawing.Size(104, 21)
        Me.cboYear.TabIndex = 65
        '
        'dtpEnd
        '
        Me.dtpEnd.CustomFormat = "hh:mmtt"
        Me.dtpEnd.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpEnd.Location = New System.Drawing.Point(117, 70)
        Me.dtpEnd.Name = "dtpEnd"
        Me.dtpEnd.ShowUpDown = True
        Me.dtpEnd.Size = New System.Drawing.Size(104, 20)
        Me.dtpEnd.TabIndex = 60
        Me.dtpEnd.Value = New Date(2014, 9, 26, 10, 21, 11, 0)
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(235, 12)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(70, 13)
        Me.Label5.TabIndex = 62
        Me.Label5.Text = "Period Type :"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(36, 73)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(72, 13)
        Me.Label3.TabIndex = 56
        Me.Label3.Text = "Ending Time :"
        '
        'dtpStart
        '
        Me.dtpStart.CustomFormat = "hh:mmtt"
        Me.dtpStart.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpStart.Location = New System.Drawing.Point(117, 39)
        Me.dtpStart.Name = "dtpStart"
        Me.dtpStart.ShowUpDown = True
        Me.dtpStart.Size = New System.Drawing.Size(104, 20)
        Me.dtpStart.TabIndex = 59
        Me.dtpStart.Value = New Date(2014, 9, 26, 11, 41, 4, 0)
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(36, 41)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(75, 13)
        Me.Label4.TabIndex = 55
        Me.Label4.Text = "Starting Time :"
        '
        'cboPeriodName
        '
        Me.cboPeriodName.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboPeriodName.FormattingEnabled = True
        Me.cboPeriodName.Location = New System.Drawing.Point(312, 38)
        Me.cboPeriodName.Name = "cboPeriodName"
        Me.cboPeriodName.Size = New System.Drawing.Size(141, 21)
        Me.cboPeriodName.TabIndex = 58
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(235, 44)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(71, 13)
        Me.Label7.TabIndex = 57
        Me.Label7.Text = "Period Name:"
        '
        'lstPeriodSetUp
        '
        Me.lstPeriodSetUp.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lstPeriodSetUp.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColumnHeader5, Me.ColumnHeader1, Me.ColumnHeader2, Me.ColumnHeader3, Me.ColumnHeader4})
        Me.lstPeriodSetUp.FullRowSelect = True
        Me.lstPeriodSetUp.GridLines = True
        Me.lstPeriodSetUp.Location = New System.Drawing.Point(2, 103)
        Me.lstPeriodSetUp.Name = "lstPeriodSetUp"
        Me.lstPeriodSetUp.Size = New System.Drawing.Size(472, 271)
        Me.lstPeriodSetUp.TabIndex = 54
        Me.lstPeriodSetUp.UseCompatibleStateImageBehavior = False
        Me.lstPeriodSetUp.View = System.Windows.Forms.View.Details
        '
        'ColumnHeader5
        '
        Me.ColumnHeader5.Text = "Period No"
        '
        'ColumnHeader1
        '
        Me.ColumnHeader1.Text = "Period Type"
        Me.ColumnHeader1.Width = 95
        '
        'ColumnHeader2
        '
        Me.ColumnHeader2.Text = "Period Name"
        Me.ColumnHeader2.Width = 130
        '
        'ColumnHeader3
        '
        Me.ColumnHeader3.Text = "Start Time"
        Me.ColumnHeader3.Width = 90
        '
        'ColumnHeader4
        '
        Me.ColumnHeader4.Text = "End Time"
        Me.ColumnHeader4.Width = 90
        '
        'StatusStrip5
        '
        Me.StatusStrip5.AutoSize = False
        Me.StatusStrip5.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip5.Dock = System.Windows.Forms.DockStyle.Left
        Me.StatusStrip5.Location = New System.Drawing.Point(0, 5)
        Me.StatusStrip5.Name = "StatusStrip5"
        Me.StatusStrip5.Size = New System.Drawing.Size(5, 378)
        Me.StatusStrip5.SizingGrip = False
        Me.StatusStrip5.TabIndex = 76
        Me.StatusStrip5.Text = "StatusStrip5"
        '
        'StatusStrip8
        '
        Me.StatusStrip8.AutoSize = False
        Me.StatusStrip8.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip8.Location = New System.Drawing.Point(0, 383)
        Me.StatusStrip8.Name = "StatusStrip8"
        Me.StatusStrip8.Size = New System.Drawing.Size(483, 5)
        Me.StatusStrip8.SizingGrip = False
        Me.StatusStrip8.TabIndex = 75
        Me.StatusStrip8.Text = "StatusStrip8"
        '
        'StatusStrip9
        '
        Me.StatusStrip9.AutoSize = False
        Me.StatusStrip9.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip9.Dock = System.Windows.Forms.DockStyle.Top
        Me.StatusStrip9.Location = New System.Drawing.Point(0, 0)
        Me.StatusStrip9.Name = "StatusStrip9"
        Me.StatusStrip9.Size = New System.Drawing.Size(483, 5)
        Me.StatusStrip9.SizingGrip = False
        Me.StatusStrip9.TabIndex = 74
        Me.StatusStrip9.Text = "StatusStrip9"
        '
        'StatusStrip10
        '
        Me.StatusStrip10.AutoSize = False
        Me.StatusStrip10.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip10.Dock = System.Windows.Forms.DockStyle.Right
        Me.StatusStrip10.Location = New System.Drawing.Point(483, 0)
        Me.StatusStrip10.Name = "StatusStrip10"
        Me.StatusStrip10.Size = New System.Drawing.Size(5, 388)
        Me.StatusStrip10.SizingGrip = False
        Me.StatusStrip10.TabIndex = 73
        Me.StatusStrip10.Text = "StatusStrip10"
        '
        'Panel1
        '
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel1.Controls.Add(Me.btnRemove)
        Me.Panel1.Controls.Add(Me.btnView)
        Me.Panel1.Controls.Add(Me.btnClose)
        Me.Panel1.Controls.Add(Me.btnUpdate)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel1.Location = New System.Drawing.Point(5, 5)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(478, 47)
        Me.Panel1.TabIndex = 77
        '
        'btnRemove
        '
        Me.btnRemove.Location = New System.Drawing.Point(146, 11)
        Me.btnRemove.Name = "btnRemove"
        Me.btnRemove.Size = New System.Drawing.Size(75, 23)
        Me.btnRemove.TabIndex = 58
        Me.btnRemove.Text = "Remove"
        Me.btnRemove.UseVisualStyleBackColor = True
        '
        'btnView
        '
        Me.btnView.Location = New System.Drawing.Point(13, 11)
        Me.btnView.Name = "btnView"
        Me.btnView.Size = New System.Drawing.Size(75, 23)
        Me.btnView.TabIndex = 57
        Me.btnView.Text = "View"
        Me.btnView.UseVisualStyleBackColor = True
        '
        'btnClose
        '
        Me.btnClose.Location = New System.Drawing.Point(398, 11)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(75, 23)
        Me.btnClose.TabIndex = 56
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'btnUpdate
        '
        Me.btnUpdate.Location = New System.Drawing.Point(270, 11)
        Me.btnUpdate.Name = "btnUpdate"
        Me.btnUpdate.Size = New System.Drawing.Size(75, 23)
        Me.btnUpdate.TabIndex = 54
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
        Me.StatusStrip3.Size = New System.Drawing.Size(5, 47)
        Me.StatusStrip3.SizingGrip = False
        Me.StatusStrip3.TabIndex = 76
        Me.StatusStrip3.Text = "StatusStrip3"
        '
        'StatusStrip4
        '
        Me.StatusStrip4.AutoSize = False
        Me.StatusStrip4.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip4.Location = New System.Drawing.Point(0, 52)
        Me.StatusStrip4.Name = "StatusStrip4"
        Me.StatusStrip4.Size = New System.Drawing.Size(483, 5)
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
        Me.StatusStrip6.Size = New System.Drawing.Size(483, 5)
        Me.StatusStrip6.SizingGrip = False
        Me.StatusStrip6.TabIndex = 74
        Me.StatusStrip6.Text = "StatusStrip6"
        '
        'StatusStrip7
        '
        Me.StatusStrip7.AutoSize = False
        Me.StatusStrip7.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip7.Dock = System.Windows.Forms.DockStyle.Right
        Me.StatusStrip7.Location = New System.Drawing.Point(483, 0)
        Me.StatusStrip7.Name = "StatusStrip7"
        Me.StatusStrip7.Size = New System.Drawing.Size(5, 57)
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
        Me.StatusStrip2.Size = New System.Drawing.Size(490, 25)
        Me.StatusStrip2.SizingGrip = False
        Me.StatusStrip2.TabIndex = 16
        Me.StatusStrip2.Text = "StatusStrip2"
        '
        'ToolStripStatusLabel1
        '
        Me.ToolStripStatusLabel1.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ToolStripStatusLabel1.Name = "ToolStripStatusLabel1"
        Me.ToolStripStatusLabel1.Size = New System.Drawing.Size(171, 20)
        Me.ToolStripStatusLabel1.Text = "TIMETABLE PERIOD SETUP"
        '
        'StatusStrip1
        '
        Me.StatusStrip1.AutoSize = False
        Me.StatusStrip1.BackColor = System.Drawing.Color.LightSeaGreen
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 475)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(490, 25)
        Me.StatusStrip1.SizingGrip = False
        Me.StatusStrip1.TabIndex = 15
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'frmTTPeriodSetUp
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.LightSkyBlue
        Me.ClientSize = New System.Drawing.Size(516, 526)
        Me.Controls.Add(Me.pnlPeriodSetup)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Name = "frmTTPeriodSetUp"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Period SetUp"
        Me.pnlPeriodSetup.ResumeLayout(False)
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.ResumeLayout(False)
        Me.Panel2.ResumeLayout(False)
        Me.Panel2.PerformLayout()
        Me.Panel1.ResumeLayout(False)
        Me.StatusStrip2.ResumeLayout(False)
        Me.StatusStrip2.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents pnlPeriodSetup As System.Windows.Forms.Panel
    Friend WithEvents StatusStrip2 As System.Windows.Forms.StatusStrip
    Friend WithEvents ToolStripStatusLabel1 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents btnUpdate As System.Windows.Forms.Button
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents btnAdd As System.Windows.Forms.Button
    Friend WithEvents dtpEnd As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents dtpStart As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents cboPeriodName As System.Windows.Forms.ComboBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents lstPeriodSetUp As System.Windows.Forms.ListView
    Friend WithEvents ColumnHeader1 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader2 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader3 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader4 As System.Windows.Forms.ColumnHeader
    Friend WithEvents StatusStrip5 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip8 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip9 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip10 As System.Windows.Forms.StatusStrip
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents StatusStrip3 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip4 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip6 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip7 As System.Windows.Forms.StatusStrip
    Friend WithEvents cboYear As System.Windows.Forms.ComboBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents cboPeriodType As System.Windows.Forms.ComboBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents btnView As System.Windows.Forms.Button
    Friend WithEvents cboPeriodNo As System.Windows.Forms.ComboBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents ColumnHeader5 As System.Windows.Forms.ColumnHeader
    Friend WithEvents btnRemove As System.Windows.Forms.Button
End Class
