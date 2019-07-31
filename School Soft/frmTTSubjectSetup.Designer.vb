<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmTTSubjectSetup
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
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.pnlSubjectSetup = New System.Windows.Forms.Panel()
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.txtAbbreviation = New System.Windows.Forms.TextBox()
        Me.txtSubjectCode = New System.Windows.Forms.TextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.txtSubColor = New System.Windows.Forms.TextBox()
        Me.cboSubjectColor = New System.Windows.Forms.ComboBox()
        Me.cboSubjectName = New System.Windows.Forms.ComboBox()
        Me.cboYear = New System.Windows.Forms.ComboBox()
        Me.cboSubjectGroup = New System.Windows.Forms.ComboBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.dgSubjectSetUp = New System.Windows.Forms.DataGridView()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.StatusStrip5 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip8 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip9 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip10 = New System.Windows.Forms.StatusStrip()
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
        Me.pnlSubjectSetup.SuspendLayout()
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.Panel2.SuspendLayout()
        CType(Me.dgSubjectSetUp, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel1.SuspendLayout()
        Me.StatusStrip2.SuspendLayout()
        Me.SuspendLayout()
        '
        'pnlSubjectSetup
        '
        Me.pnlSubjectSetup.BackColor = System.Drawing.SystemColors.Control
        Me.pnlSubjectSetup.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pnlSubjectSetup.Controls.Add(Me.SplitContainer1)
        Me.pnlSubjectSetup.Controls.Add(Me.StatusStrip2)
        Me.pnlSubjectSetup.Controls.Add(Me.StatusStrip1)
        Me.pnlSubjectSetup.Location = New System.Drawing.Point(12, 12)
        Me.pnlSubjectSetup.Name = "pnlSubjectSetup"
        Me.pnlSubjectSetup.Size = New System.Drawing.Size(438, 502)
        Me.pnlSubjectSetup.TabIndex = 1
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
        Me.SplitContainer1.Size = New System.Drawing.Size(436, 450)
        Me.SplitContainer1.SplitterDistance = 390
        Me.SplitContainer1.SplitterWidth = 1
        Me.SplitContainer1.TabIndex = 17
        '
        'Panel2
        '
        Me.Panel2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel2.Controls.Add(Me.txtAbbreviation)
        Me.Panel2.Controls.Add(Me.txtSubjectCode)
        Me.Panel2.Controls.Add(Me.Label6)
        Me.Panel2.Controls.Add(Me.txtSubColor)
        Me.Panel2.Controls.Add(Me.cboSubjectColor)
        Me.Panel2.Controls.Add(Me.cboSubjectName)
        Me.Panel2.Controls.Add(Me.cboYear)
        Me.Panel2.Controls.Add(Me.cboSubjectGroup)
        Me.Panel2.Controls.Add(Me.Label5)
        Me.Panel2.Controls.Add(Me.dgSubjectSetUp)
        Me.Panel2.Controls.Add(Me.Label7)
        Me.Panel2.Controls.Add(Me.Label1)
        Me.Panel2.Controls.Add(Me.Label2)
        Me.Panel2.Controls.Add(Me.Label3)
        Me.Panel2.Controls.Add(Me.Label4)
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel2.Location = New System.Drawing.Point(5, 5)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(424, 378)
        Me.Panel2.TabIndex = 81
        '
        'txtAbbreviation
        '
        Me.txtAbbreviation.Location = New System.Drawing.Point(286, 67)
        Me.txtAbbreviation.Name = "txtAbbreviation"
        Me.txtAbbreviation.ReadOnly = True
        Me.txtAbbreviation.Size = New System.Drawing.Size(101, 20)
        Me.txtAbbreviation.TabIndex = 29
        '
        'txtSubjectCode
        '
        Me.txtSubjectCode.Location = New System.Drawing.Point(99, 67)
        Me.txtSubjectCode.Name = "txtSubjectCode"
        Me.txtSubjectCode.ReadOnly = True
        Me.txtSubjectCode.Size = New System.Drawing.Size(101, 20)
        Me.txtSubjectCode.TabIndex = 28
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(3, 128)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(94, 13)
        Me.Label6.TabIndex = 18
        Me.Label6.Text = "Subject Time Offs:"
        '
        'txtSubColor
        '
        Me.txtSubColor.Location = New System.Drawing.Point(331, 95)
        Me.txtSubColor.Multiline = True
        Me.txtSubColor.Name = "txtSubColor"
        Me.txtSubColor.ReadOnly = True
        Me.txtSubColor.Size = New System.Drawing.Size(86, 46)
        Me.txtSubColor.TabIndex = 16
        '
        'cboSubjectColor
        '
        Me.cboSubjectColor.FormattingEnabled = True
        Me.cboSubjectColor.Location = New System.Drawing.Point(99, 95)
        Me.cboSubjectColor.Name = "cboSubjectColor"
        Me.cboSubjectColor.Size = New System.Drawing.Size(225, 21)
        Me.cboSubjectColor.TabIndex = 15
        '
        'cboSubjectName
        '
        Me.cboSubjectName.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboSubjectName.FormattingEnabled = True
        Me.cboSubjectName.Location = New System.Drawing.Point(99, 39)
        Me.cboSubjectName.Name = "cboSubjectName"
        Me.cboSubjectName.Size = New System.Drawing.Size(288, 21)
        Me.cboSubjectName.TabIndex = 12
        '
        'cboYear
        '
        Me.cboYear.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboYear.FormattingEnabled = True
        Me.cboYear.Location = New System.Drawing.Point(330, 11)
        Me.cboYear.Name = "cboYear"
        Me.cboYear.Size = New System.Drawing.Size(86, 21)
        Me.cboYear.TabIndex = 11
        '
        'cboSubjectGroup
        '
        Me.cboSubjectGroup.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboSubjectGroup.FormattingEnabled = True
        Me.cboSubjectGroup.Location = New System.Drawing.Point(99, 11)
        Me.cboSubjectGroup.Name = "cboSubjectGroup"
        Me.cboSubjectGroup.Size = New System.Drawing.Size(187, 21)
        Me.cboSubjectGroup.TabIndex = 11
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(15, 98)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(73, 13)
        Me.Label5.TabIndex = 10
        Me.Label5.Text = "Subject Color:"
        '
        'dgSubjectSetUp
        '
        Me.dgSubjectSetUp.AllowUserToAddRows = False
        Me.dgSubjectSetUp.AllowUserToDeleteRows = False
        Me.dgSubjectSetUp.AllowUserToResizeColumns = False
        Me.dgSubjectSetUp.AllowUserToResizeRows = False
        Me.dgSubjectSetUp.BackgroundColor = System.Drawing.SystemColors.ControlLightLight
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgSubjectSetUp.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle1
        Me.dgSubjectSetUp.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.dgSubjectSetUp.DefaultCellStyle = DataGridViewCellStyle2
        Me.dgSubjectSetUp.Location = New System.Drawing.Point(5, 148)
        Me.dgSubjectSetUp.Name = "dgSubjectSetUp"
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle3.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle3.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle3.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgSubjectSetUp.RowHeadersDefaultCellStyle = DataGridViewCellStyle3
        Me.dgSubjectSetUp.RowHeadersWidth = 4
        Me.dgSubjectSetUp.ShowCellErrors = False
        Me.dgSubjectSetUp.ShowCellToolTips = False
        Me.dgSubjectSetUp.ShowEditingIcon = False
        Me.dgSubjectSetUp.ShowRowErrors = False
        Me.dgSubjectSetUp.Size = New System.Drawing.Size(412, 225)
        Me.dgSubjectSetUp.TabIndex = 8
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(292, 14)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(32, 13)
        Me.Label7.TabIndex = 5
        Me.Label7.Text = "Year:"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(15, 14)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(78, 13)
        Me.Label1.TabIndex = 5
        Me.Label1.Text = "Subject Group:"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(15, 42)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(77, 13)
        Me.Label2.TabIndex = 6
        Me.Label2.Text = "Subject Name:"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(15, 70)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(74, 13)
        Me.Label3.TabIndex = 7
        Me.Label3.Text = "Subject Code:"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(206, 70)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(69, 13)
        Me.Label4.TabIndex = 9
        Me.Label4.Text = "Abbreviation:"
        '
        'StatusStrip5
        '
        Me.StatusStrip5.AutoSize = False
        Me.StatusStrip5.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip5.Location = New System.Drawing.Point(5, 383)
        Me.StatusStrip5.Name = "StatusStrip5"
        Me.StatusStrip5.Size = New System.Drawing.Size(424, 5)
        Me.StatusStrip5.SizingGrip = False
        Me.StatusStrip5.TabIndex = 80
        Me.StatusStrip5.Text = "StatusStrip5"
        '
        'StatusStrip8
        '
        Me.StatusStrip8.AutoSize = False
        Me.StatusStrip8.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip8.Dock = System.Windows.Forms.DockStyle.Top
        Me.StatusStrip8.Location = New System.Drawing.Point(5, 0)
        Me.StatusStrip8.Name = "StatusStrip8"
        Me.StatusStrip8.Size = New System.Drawing.Size(424, 5)
        Me.StatusStrip8.SizingGrip = False
        Me.StatusStrip8.TabIndex = 79
        Me.StatusStrip8.Text = "StatusStrip8"
        '
        'StatusStrip9
        '
        Me.StatusStrip9.AutoSize = False
        Me.StatusStrip9.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip9.Dock = System.Windows.Forms.DockStyle.Right
        Me.StatusStrip9.Location = New System.Drawing.Point(429, 0)
        Me.StatusStrip9.Name = "StatusStrip9"
        Me.StatusStrip9.Size = New System.Drawing.Size(5, 388)
        Me.StatusStrip9.SizingGrip = False
        Me.StatusStrip9.TabIndex = 78
        Me.StatusStrip9.Text = "StatusStrip9"
        '
        'StatusStrip10
        '
        Me.StatusStrip10.AutoSize = False
        Me.StatusStrip10.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip10.Dock = System.Windows.Forms.DockStyle.Left
        Me.StatusStrip10.Location = New System.Drawing.Point(0, 0)
        Me.StatusStrip10.Name = "StatusStrip10"
        Me.StatusStrip10.Size = New System.Drawing.Size(5, 388)
        Me.StatusStrip10.SizingGrip = False
        Me.StatusStrip10.TabIndex = 82
        Me.StatusStrip10.Text = "StatusStrip10"
        '
        'Panel1
        '
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel1.Controls.Add(Me.btnClose)
        Me.Panel1.Controls.Add(Me.btnUpdate)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel1.Location = New System.Drawing.Point(5, 5)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(424, 47)
        Me.Panel1.TabIndex = 77
        '
        'btnClose
        '
        Me.btnClose.Location = New System.Drawing.Point(330, 11)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(75, 23)
        Me.btnClose.TabIndex = 58
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'btnUpdate
        '
        Me.btnUpdate.Location = New System.Drawing.Point(99, 11)
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
        Me.StatusStrip4.Size = New System.Drawing.Size(429, 5)
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
        Me.StatusStrip6.Size = New System.Drawing.Size(429, 5)
        Me.StatusStrip6.SizingGrip = False
        Me.StatusStrip6.TabIndex = 74
        Me.StatusStrip6.Text = "StatusStrip6"
        '
        'StatusStrip7
        '
        Me.StatusStrip7.AutoSize = False
        Me.StatusStrip7.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip7.Dock = System.Windows.Forms.DockStyle.Right
        Me.StatusStrip7.Location = New System.Drawing.Point(429, 0)
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
        Me.StatusStrip2.Size = New System.Drawing.Size(436, 25)
        Me.StatusStrip2.SizingGrip = False
        Me.StatusStrip2.TabIndex = 16
        Me.StatusStrip2.Text = "StatusStrip2"
        '
        'ToolStripStatusLabel1
        '
        Me.ToolStripStatusLabel1.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ToolStripStatusLabel1.Name = "ToolStripStatusLabel1"
        Me.ToolStripStatusLabel1.Size = New System.Drawing.Size(177, 20)
        Me.ToolStripStatusLabel1.Text = "TIMETABLE SUBJECT SETUP"
        '
        'StatusStrip1
        '
        Me.StatusStrip1.AutoSize = False
        Me.StatusStrip1.BackColor = System.Drawing.Color.LightSeaGreen
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 475)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(436, 25)
        Me.StatusStrip1.SizingGrip = False
        Me.StatusStrip1.TabIndex = 15
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'frmTTSubjectSetup
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.LightSkyBlue
        Me.ClientSize = New System.Drawing.Size(460, 526)
        Me.Controls.Add(Me.pnlSubjectSetup)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Name = "frmTTSubjectSetup"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Subject Setup"
        Me.pnlSubjectSetup.ResumeLayout(False)
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.ResumeLayout(False)
        Me.Panel2.ResumeLayout(False)
        Me.Panel2.PerformLayout()
        CType(Me.dgSubjectSetUp, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel1.ResumeLayout(False)
        Me.StatusStrip2.ResumeLayout(False)
        Me.StatusStrip2.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents pnlSubjectSetup As System.Windows.Forms.Panel
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents StatusStrip3 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip4 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip6 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip7 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip2 As System.Windows.Forms.StatusStrip
    Friend WithEvents ToolStripStatusLabel1 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents txtSubColor As System.Windows.Forms.TextBox
    Friend WithEvents cboSubjectColor As System.Windows.Forms.ComboBox
    Friend WithEvents cboSubjectName As System.Windows.Forms.ComboBox
    Friend WithEvents cboSubjectGroup As System.Windows.Forms.ComboBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents dgSubjectSetUp As System.Windows.Forms.DataGridView
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents StatusStrip5 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip8 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip9 As System.Windows.Forms.StatusStrip
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents btnUpdate As System.Windows.Forms.Button
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents txtAbbreviation As System.Windows.Forms.TextBox
    Friend WithEvents txtSubjectCode As System.Windows.Forms.TextBox
    Friend WithEvents cboYear As System.Windows.Forms.ComboBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents StatusStrip10 As System.Windows.Forms.StatusStrip
End Class
