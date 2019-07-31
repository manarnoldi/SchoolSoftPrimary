<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmInvIssueApproval
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
        Me.pnlIssueApproval = New System.Windows.Forms.Panel()
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
        Me.SplitContainer2 = New System.Windows.Forms.SplitContainer()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.txtApprName = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.cboApprNo = New System.Windows.Forms.ComboBox()
        Me.btnLoad = New System.Windows.Forms.Button()
        Me.dtpReqTo = New System.Windows.Forms.DateTimePicker()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.dtpReqFrom = New System.Windows.Forms.DateTimePicker()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.StatusStrip7 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip8 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip9 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip10 = New System.Windows.Forms.StatusStrip()
        Me.dgVwIssueAppr = New System.Windows.Forms.DataGridView()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.btnUpdate = New System.Windows.Forms.Button()
        Me.btnClose = New System.Windows.Forms.Button()
        Me.StatusStrip3 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip4 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip5 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip6 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip2 = New System.Windows.Forms.StatusStrip()
        Me.ToolStripStatusLabel1 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip()
        Me.Column1 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column3 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column2 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column4 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column5 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column6 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column7 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column8 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column9 = New System.Windows.Forms.DataGridViewCheckBoxColumn()
        Me.Column10 = New System.Windows.Forms.DataGridViewCheckBoxColumn()
        Me.pnlIssueApproval.SuspendLayout()
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        CType(Me.SplitContainer2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer2.Panel1.SuspendLayout()
        Me.SplitContainer2.Panel2.SuspendLayout()
        Me.SplitContainer2.SuspendLayout()
        Me.Panel1.SuspendLayout()
        CType(Me.dgVwIssueAppr, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel2.SuspendLayout()
        Me.StatusStrip2.SuspendLayout()
        Me.SuspendLayout()
        '
        'pnlIssueApproval
        '
        Me.pnlIssueApproval.BackColor = System.Drawing.SystemColors.Control
        Me.pnlIssueApproval.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pnlIssueApproval.Controls.Add(Me.SplitContainer1)
        Me.pnlIssueApproval.Controls.Add(Me.StatusStrip2)
        Me.pnlIssueApproval.Controls.Add(Me.StatusStrip1)
        Me.pnlIssueApproval.Location = New System.Drawing.Point(12, 12)
        Me.pnlIssueApproval.Name = "pnlIssueApproval"
        Me.pnlIssueApproval.Size = New System.Drawing.Size(1000, 539)
        Me.pnlIssueApproval.TabIndex = 0
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
        Me.SplitContainer1.Panel1.Controls.Add(Me.SplitContainer2)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.Panel2)
        Me.SplitContainer1.Panel2.Controls.Add(Me.StatusStrip3)
        Me.SplitContainer1.Panel2.Controls.Add(Me.StatusStrip4)
        Me.SplitContainer1.Panel2.Controls.Add(Me.StatusStrip5)
        Me.SplitContainer1.Panel2.Controls.Add(Me.StatusStrip6)
        Me.SplitContainer1.Size = New System.Drawing.Size(998, 487)
        Me.SplitContainer1.SplitterDistance = 432
        Me.SplitContainer1.SplitterWidth = 1
        Me.SplitContainer1.TabIndex = 37
        '
        'SplitContainer2
        '
        Me.SplitContainer2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.SplitContainer2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer2.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer2.Name = "SplitContainer2"
        Me.SplitContainer2.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'SplitContainer2.Panel1
        '
        Me.SplitContainer2.Panel1.Controls.Add(Me.Panel1)
        Me.SplitContainer2.Panel1.Controls.Add(Me.StatusStrip7)
        Me.SplitContainer2.Panel1.Controls.Add(Me.StatusStrip8)
        Me.SplitContainer2.Panel1.Controls.Add(Me.StatusStrip9)
        Me.SplitContainer2.Panel1.Controls.Add(Me.StatusStrip10)
        '
        'SplitContainer2.Panel2
        '
        Me.SplitContainer2.Panel2.Controls.Add(Me.dgVwIssueAppr)
        Me.SplitContainer2.Size = New System.Drawing.Size(998, 432)
        Me.SplitContainer2.SplitterDistance = 65
        Me.SplitContainer2.SplitterWidth = 1
        Me.SplitContainer2.TabIndex = 0
        '
        'Panel1
        '
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel1.Controls.Add(Me.txtApprName)
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Controls.Add(Me.Label3)
        Me.Panel1.Controls.Add(Me.cboApprNo)
        Me.Panel1.Controls.Add(Me.btnLoad)
        Me.Panel1.Controls.Add(Me.dtpReqTo)
        Me.Panel1.Controls.Add(Me.Label13)
        Me.Panel1.Controls.Add(Me.dtpReqFrom)
        Me.Panel1.Controls.Add(Me.Label12)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel1.Location = New System.Drawing.Point(5, 5)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(986, 53)
        Me.Panel1.TabIndex = 78
        '
        'txtApprName
        '
        Me.txtApprName.Location = New System.Drawing.Point(748, 15)
        Me.txtApprName.Name = "txtApprName"
        Me.txtApprName.ReadOnly = True
        Me.txtApprName.Size = New System.Drawing.Size(226, 20)
        Me.txtApprName.TabIndex = 100
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(701, 18)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(41, 13)
        Me.Label1.TabIndex = 99
        Me.Label1.Text = "Name :"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(501, 18)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(91, 13)
        Me.Label3.TabIndex = 97
        Me.Label3.Text = "Approved By No :"
        '
        'cboApprNo
        '
        Me.cboApprNo.DropDownHeight = 100
        Me.cboApprNo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboApprNo.FormattingEnabled = True
        Me.cboApprNo.IntegralHeight = False
        Me.cboApprNo.Location = New System.Drawing.Point(598, 15)
        Me.cboApprNo.Name = "cboApprNo"
        Me.cboApprNo.Size = New System.Drawing.Size(97, 21)
        Me.cboApprNo.TabIndex = 98
        '
        'btnLoad
        '
        Me.btnLoad.Location = New System.Drawing.Point(402, 13)
        Me.btnLoad.Name = "btnLoad"
        Me.btnLoad.Size = New System.Drawing.Size(75, 23)
        Me.btnLoad.TabIndex = 96
        Me.btnLoad.Text = "Load"
        Me.btnLoad.UseVisualStyleBackColor = True
        '
        'dtpReqTo
        '
        Me.dtpReqTo.CustomFormat = " dd - MM - yyyy"
        Me.dtpReqTo.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpReqTo.Location = New System.Drawing.Point(264, 15)
        Me.dtpReqTo.Name = "dtpReqTo"
        Me.dtpReqTo.Size = New System.Drawing.Size(119, 20)
        Me.dtpReqTo.TabIndex = 95
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Location = New System.Drawing.Point(206, 18)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(52, 13)
        Me.Label13.TabIndex = 94
        Me.Label13.Text = "To Date :"
        '
        'dtpReqFrom
        '
        Me.dtpReqFrom.CustomFormat = " dd - MM - yyyy"
        Me.dtpReqFrom.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpReqFrom.Location = New System.Drawing.Point(81, 15)
        Me.dtpReqFrom.Name = "dtpReqFrom"
        Me.dtpReqFrom.Size = New System.Drawing.Size(119, 20)
        Me.dtpReqFrom.TabIndex = 93
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Location = New System.Drawing.Point(13, 18)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(62, 13)
        Me.Label12.TabIndex = 92
        Me.Label12.Text = "From Date :"
        '
        'StatusStrip7
        '
        Me.StatusStrip7.AutoSize = False
        Me.StatusStrip7.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip7.Dock = System.Windows.Forms.DockStyle.Right
        Me.StatusStrip7.Location = New System.Drawing.Point(991, 5)
        Me.StatusStrip7.Name = "StatusStrip7"
        Me.StatusStrip7.Size = New System.Drawing.Size(5, 53)
        Me.StatusStrip7.SizingGrip = False
        Me.StatusStrip7.TabIndex = 77
        Me.StatusStrip7.Text = "StatusStrip7"
        '
        'StatusStrip8
        '
        Me.StatusStrip8.AutoSize = False
        Me.StatusStrip8.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip8.Dock = System.Windows.Forms.DockStyle.Left
        Me.StatusStrip8.Location = New System.Drawing.Point(0, 5)
        Me.StatusStrip8.Name = "StatusStrip8"
        Me.StatusStrip8.Size = New System.Drawing.Size(5, 53)
        Me.StatusStrip8.SizingGrip = False
        Me.StatusStrip8.TabIndex = 76
        Me.StatusStrip8.Text = "StatusStrip8"
        '
        'StatusStrip9
        '
        Me.StatusStrip9.AutoSize = False
        Me.StatusStrip9.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip9.Location = New System.Drawing.Point(0, 58)
        Me.StatusStrip9.Name = "StatusStrip9"
        Me.StatusStrip9.Size = New System.Drawing.Size(996, 5)
        Me.StatusStrip9.SizingGrip = False
        Me.StatusStrip9.TabIndex = 75
        Me.StatusStrip9.Text = "StatusStrip9"
        '
        'StatusStrip10
        '
        Me.StatusStrip10.AutoSize = False
        Me.StatusStrip10.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip10.Dock = System.Windows.Forms.DockStyle.Top
        Me.StatusStrip10.Location = New System.Drawing.Point(0, 0)
        Me.StatusStrip10.Name = "StatusStrip10"
        Me.StatusStrip10.Size = New System.Drawing.Size(996, 5)
        Me.StatusStrip10.SizingGrip = False
        Me.StatusStrip10.TabIndex = 74
        Me.StatusStrip10.Text = "StatusStrip10"
        '
        'dgVwIssueAppr
        '
        Me.dgVwIssueAppr.AllowUserToAddRows = False
        Me.dgVwIssueAppr.AllowUserToDeleteRows = False
        Me.dgVwIssueAppr.AllowUserToResizeColumns = False
        Me.dgVwIssueAppr.AllowUserToResizeRows = False
        Me.dgVwIssueAppr.BackgroundColor = System.Drawing.SystemColors.ControlLightLight
        Me.dgVwIssueAppr.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.dgVwIssueAppr.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgVwIssueAppr.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Column1, Me.Column3, Me.Column2, Me.Column4, Me.Column5, Me.Column6, Me.Column7, Me.Column8, Me.Column9, Me.Column10})
        Me.dgVwIssueAppr.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgVwIssueAppr.Location = New System.Drawing.Point(0, 0)
        Me.dgVwIssueAppr.Name = "dgVwIssueAppr"
        Me.dgVwIssueAppr.RowHeadersWidth = 4
        Me.dgVwIssueAppr.ShowCellErrors = False
        Me.dgVwIssueAppr.ShowCellToolTips = False
        Me.dgVwIssueAppr.ShowEditingIcon = False
        Me.dgVwIssueAppr.ShowRowErrors = False
        Me.dgVwIssueAppr.Size = New System.Drawing.Size(996, 364)
        Me.dgVwIssueAppr.TabIndex = 1
        '
        'Panel2
        '
        Me.Panel2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel2.Controls.Add(Me.btnUpdate)
        Me.Panel2.Controls.Add(Me.btnClose)
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel2.Location = New System.Drawing.Point(5, 5)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(986, 42)
        Me.Panel2.TabIndex = 73
        '
        'btnUpdate
        '
        Me.btnUpdate.Location = New System.Drawing.Point(678, 8)
        Me.btnUpdate.Name = "btnUpdate"
        Me.btnUpdate.Size = New System.Drawing.Size(75, 23)
        Me.btnUpdate.TabIndex = 84
        Me.btnUpdate.Text = "Update"
        Me.btnUpdate.UseVisualStyleBackColor = True
        '
        'btnClose
        '
        Me.btnClose.Location = New System.Drawing.Point(899, 8)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(75, 23)
        Me.btnClose.TabIndex = 83
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'StatusStrip3
        '
        Me.StatusStrip3.AutoSize = False
        Me.StatusStrip3.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip3.Dock = System.Windows.Forms.DockStyle.Right
        Me.StatusStrip3.Location = New System.Drawing.Point(991, 5)
        Me.StatusStrip3.Name = "StatusStrip3"
        Me.StatusStrip3.Size = New System.Drawing.Size(5, 42)
        Me.StatusStrip3.SizingGrip = False
        Me.StatusStrip3.TabIndex = 72
        Me.StatusStrip3.Text = "StatusStrip3"
        '
        'StatusStrip4
        '
        Me.StatusStrip4.AutoSize = False
        Me.StatusStrip4.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip4.Dock = System.Windows.Forms.DockStyle.Left
        Me.StatusStrip4.Location = New System.Drawing.Point(0, 5)
        Me.StatusStrip4.Name = "StatusStrip4"
        Me.StatusStrip4.Size = New System.Drawing.Size(5, 42)
        Me.StatusStrip4.SizingGrip = False
        Me.StatusStrip4.TabIndex = 71
        Me.StatusStrip4.Text = "StatusStrip4"
        '
        'StatusStrip5
        '
        Me.StatusStrip5.AutoSize = False
        Me.StatusStrip5.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip5.Location = New System.Drawing.Point(0, 47)
        Me.StatusStrip5.Name = "StatusStrip5"
        Me.StatusStrip5.Size = New System.Drawing.Size(996, 5)
        Me.StatusStrip5.SizingGrip = False
        Me.StatusStrip5.TabIndex = 70
        Me.StatusStrip5.Text = "StatusStrip5"
        '
        'StatusStrip6
        '
        Me.StatusStrip6.AutoSize = False
        Me.StatusStrip6.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip6.Dock = System.Windows.Forms.DockStyle.Top
        Me.StatusStrip6.Location = New System.Drawing.Point(0, 0)
        Me.StatusStrip6.Name = "StatusStrip6"
        Me.StatusStrip6.Size = New System.Drawing.Size(996, 5)
        Me.StatusStrip6.SizingGrip = False
        Me.StatusStrip6.TabIndex = 69
        Me.StatusStrip6.Text = "StatusStrip6"
        '
        'StatusStrip2
        '
        Me.StatusStrip2.AutoSize = False
        Me.StatusStrip2.BackColor = System.Drawing.Color.LightSeaGreen
        Me.StatusStrip2.Dock = System.Windows.Forms.DockStyle.Top
        Me.StatusStrip2.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripStatusLabel1})
        Me.StatusStrip2.Location = New System.Drawing.Point(0, 0)
        Me.StatusStrip2.Name = "StatusStrip2"
        Me.StatusStrip2.Size = New System.Drawing.Size(998, 25)
        Me.StatusStrip2.SizingGrip = False
        Me.StatusStrip2.TabIndex = 36
        Me.StatusStrip2.Text = "StatusStrip2"
        '
        'ToolStripStatusLabel1
        '
        Me.ToolStripStatusLabel1.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ToolStripStatusLabel1.Name = "ToolStripStatusLabel1"
        Me.ToolStripStatusLabel1.Size = New System.Drawing.Size(191, 20)
        Me.ToolStripStatusLabel1.Text = "INVENTORY ISSUE APPROVAL"
        '
        'StatusStrip1
        '
        Me.StatusStrip1.AutoSize = False
        Me.StatusStrip1.BackColor = System.Drawing.Color.LightSeaGreen
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 512)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(998, 25)
        Me.StatusStrip1.SizingGrip = False
        Me.StatusStrip1.TabIndex = 35
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'Column1
        '
        Me.Column1.HeaderText = "IssueNote No"
        Me.Column1.Name = "Column1"
        Me.Column1.ReadOnly = True
        Me.Column1.Width = 120
        '
        'Column3
        '
        Me.Column3.HeaderText = "Item Name"
        Me.Column3.Name = "Column3"
        Me.Column3.ReadOnly = True
        Me.Column3.Width = 170
        '
        'Column2
        '
        Me.Column2.HeaderText = "UOM"
        Me.Column2.Name = "Column2"
        Me.Column2.ReadOnly = True
        Me.Column2.Width = 60
        '
        'Column4
        '
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        Me.Column4.DefaultCellStyle = DataGridViewCellStyle1
        Me.Column4.HeaderText = "Quantity"
        Me.Column4.Name = "Column4"
        Me.Column4.ReadOnly = True
        Me.Column4.Width = 60
        '
        'Column5
        '
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        Me.Column5.DefaultCellStyle = DataGridViewCellStyle2
        Me.Column5.HeaderText = "Unit Cost"
        Me.Column5.Name = "Column5"
        Me.Column5.ReadOnly = True
        Me.Column5.Width = 60
        '
        'Column6
        '
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        Me.Column6.DefaultCellStyle = DataGridViewCellStyle3
        Me.Column6.HeaderText = "Total Cost"
        Me.Column6.Name = "Column6"
        Me.Column6.ReadOnly = True
        Me.Column6.Width = 80
        '
        'Column7
        '
        Me.Column7.HeaderText = "RequestBy"
        Me.Column7.Name = "Column7"
        Me.Column7.ReadOnly = True
        '
        'Column8
        '
        Me.Column8.HeaderText = "IssueTo"
        Me.Column8.Name = "Column8"
        Me.Column8.ReadOnly = True
        Me.Column8.Width = 200
        '
        'Column9
        '
        Me.Column9.HeaderText = "Approve"
        Me.Column9.Name = "Column9"
        Me.Column9.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.Column9.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.Column9.Width = 60
        '
        'Column10
        '
        Me.Column10.HeaderText = "Reject"
        Me.Column10.Name = "Column10"
        Me.Column10.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.Column10.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.Column10.Width = 60
        '
        'frmInvIssueApproval
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.LightSkyBlue
        Me.ClientSize = New System.Drawing.Size(1024, 563)
        Me.Controls.Add(Me.pnlIssueApproval)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Name = "frmInvIssueApproval"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Issue Approval"
        Me.pnlIssueApproval.ResumeLayout(False)
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.ResumeLayout(False)
        Me.SplitContainer2.Panel1.ResumeLayout(False)
        Me.SplitContainer2.Panel2.ResumeLayout(False)
        CType(Me.SplitContainer2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer2.ResumeLayout(False)
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        CType(Me.dgVwIssueAppr, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel2.ResumeLayout(False)
        Me.StatusStrip2.ResumeLayout(False)
        Me.StatusStrip2.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents pnlIssueApproval As System.Windows.Forms.Panel
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents StatusStrip2 As System.Windows.Forms.StatusStrip
    Friend WithEvents ToolStripStatusLabel1 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents btnUpdate As System.Windows.Forms.Button
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents StatusStrip3 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip4 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip5 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip6 As System.Windows.Forms.StatusStrip
    Friend WithEvents SplitContainer2 As System.Windows.Forms.SplitContainer
    Friend WithEvents StatusStrip7 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip8 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip9 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip10 As System.Windows.Forms.StatusStrip
    Friend WithEvents dgVwIssueAppr As System.Windows.Forms.DataGridView
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents btnLoad As System.Windows.Forms.Button
    Friend WithEvents dtpReqTo As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents dtpReqFrom As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents cboApprNo As System.Windows.Forms.ComboBox
    Friend WithEvents txtApprName As System.Windows.Forms.TextBox
    Friend WithEvents Column1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column3 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column2 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column4 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column5 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column6 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column7 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column8 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column9 As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents Column10 As System.Windows.Forms.DataGridViewCheckBoxColumn
End Class
