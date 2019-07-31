<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmFinPayApproval
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
        Me.pnlPayApprove = New System.Windows.Forms.Panel()
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.btnLoad = New System.Windows.Forms.Button()
        Me.dtpApprovalDate = New System.Windows.Forms.DateTimePicker()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.txtEmpName = New System.Windows.Forms.TextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.cboEmpNo = New System.Windows.Forms.ComboBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.cboVoucherNo = New System.Windows.Forms.ComboBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.cboYear = New System.Windows.Forms.ComboBox()
        Me.cboTerm = New System.Windows.Forms.ComboBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.StatusStrip3 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip4 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip6 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip7 = New System.Windows.Forms.StatusStrip()
        Me.SplitContainer2 = New System.Windows.Forms.SplitContainer()
        Me.lstPayApprovals = New System.Windows.Forms.ListView()
        Me.ColumnHeader1 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader2 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader3 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader4 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader5 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader6 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.btnApprove = New System.Windows.Forms.Button()
        Me.btnClose = New System.Windows.Forms.Button()
        Me.StatusStrip5 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip8 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip9 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip10 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip2 = New System.Windows.Forms.StatusStrip()
        Me.ToolStripStatusLabel1 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip()
        Me.pnlPayApprove.SuspendLayout()
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.Panel1.SuspendLayout()
        CType(Me.SplitContainer2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer2.Panel1.SuspendLayout()
        Me.SplitContainer2.Panel2.SuspendLayout()
        Me.SplitContainer2.SuspendLayout()
        Me.Panel3.SuspendLayout()
        Me.StatusStrip2.SuspendLayout()
        Me.SuspendLayout()
        '
        'pnlPayApprove
        '
        Me.pnlPayApprove.BackColor = System.Drawing.SystemColors.Control
        Me.pnlPayApprove.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pnlPayApprove.Controls.Add(Me.SplitContainer1)
        Me.pnlPayApprove.Controls.Add(Me.StatusStrip2)
        Me.pnlPayApprove.Controls.Add(Me.StatusStrip1)
        Me.pnlPayApprove.Location = New System.Drawing.Point(12, 12)
        Me.pnlPayApprove.Name = "pnlPayApprove"
        Me.pnlPayApprove.Size = New System.Drawing.Size(913, 432)
        Me.pnlPayApprove.TabIndex = 0
        '
        'SplitContainer1
        '
        Me.SplitContainer1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.SplitContainer1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer1.Location = New System.Drawing.Point(0, 25)
        Me.SplitContainer1.Name = "SplitContainer1"
        Me.SplitContainer1.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.Panel1)
        Me.SplitContainer1.Panel1.Controls.Add(Me.StatusStrip3)
        Me.SplitContainer1.Panel1.Controls.Add(Me.StatusStrip4)
        Me.SplitContainer1.Panel1.Controls.Add(Me.StatusStrip6)
        Me.SplitContainer1.Panel1.Controls.Add(Me.StatusStrip7)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.SplitContainer2)
        Me.SplitContainer1.Size = New System.Drawing.Size(911, 380)
        Me.SplitContainer1.SplitterDistance = 84
        Me.SplitContainer1.SplitterWidth = 1
        Me.SplitContainer1.TabIndex = 35
        '
        'Panel1
        '
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel1.Controls.Add(Me.btnLoad)
        Me.Panel1.Controls.Add(Me.dtpApprovalDate)
        Me.Panel1.Controls.Add(Me.Label9)
        Me.Panel1.Controls.Add(Me.txtEmpName)
        Me.Panel1.Controls.Add(Me.Label6)
        Me.Panel1.Controls.Add(Me.cboEmpNo)
        Me.Panel1.Controls.Add(Me.Label5)
        Me.Panel1.Controls.Add(Me.Label4)
        Me.Panel1.Controls.Add(Me.cboVoucherNo)
        Me.Panel1.Controls.Add(Me.Label3)
        Me.Panel1.Controls.Add(Me.cboYear)
        Me.Panel1.Controls.Add(Me.cboTerm)
        Me.Panel1.Controls.Add(Me.Label2)
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel1.Location = New System.Drawing.Point(5, 5)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(897, 70)
        Me.Panel1.TabIndex = 93
        '
        'btnLoad
        '
        Me.btnLoad.Location = New System.Drawing.Point(805, 7)
        Me.btnLoad.Name = "btnLoad"
        Me.btnLoad.Size = New System.Drawing.Size(75, 23)
        Me.btnLoad.TabIndex = 44
        Me.btnLoad.Text = "Load"
        Me.btnLoad.UseVisualStyleBackColor = True
        '
        'dtpApprovalDate
        '
        Me.dtpApprovalDate.CustomFormat = "  dd - MM - yyyy"
        Me.dtpApprovalDate.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtpApprovalDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpApprovalDate.Location = New System.Drawing.Point(664, 9)
        Me.dtpApprovalDate.Name = "dtpApprovalDate"
        Me.dtpApprovalDate.Size = New System.Drawing.Size(110, 20)
        Me.dtpApprovalDate.TabIndex = 43
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(580, 12)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(78, 13)
        Me.Label9.TabIndex = 42
        Me.Label9.Text = "Approval Date:"
        '
        'txtEmpName
        '
        Me.txtEmpName.Location = New System.Drawing.Point(433, 39)
        Me.txtEmpName.Name = "txtEmpName"
        Me.txtEmpName.ReadOnly = True
        Me.txtEmpName.Size = New System.Drawing.Size(447, 20)
        Me.txtEmpName.TabIndex = 37
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(340, 42)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(87, 13)
        Me.Label6.TabIndex = 36
        Me.Label6.Text = "Employee Name:"
        '
        'cboEmpNo
        '
        Me.cboEmpNo.DropDownHeight = 80
        Me.cboEmpNo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboEmpNo.FormattingEnabled = True
        Me.cboEmpNo.IntegralHeight = False
        Me.cboEmpNo.Location = New System.Drawing.Point(192, 39)
        Me.cboEmpNo.Name = "cboEmpNo"
        Me.cboEmpNo.Size = New System.Drawing.Size(131, 21)
        Me.cboEmpNo.TabIndex = 35
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(90, 42)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(96, 13)
        Me.Label5.TabIndex = 34
        Me.Label5.Text = "Employee Number:"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(13, 42)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(71, 13)
        Me.Label4.TabIndex = 33
        Me.Label4.Text = "Approved By:"
        '
        'cboVoucherNo
        '
        Me.cboVoucherNo.DropDownHeight = 80
        Me.cboVoucherNo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboVoucherNo.FormattingEnabled = True
        Me.cboVoucherNo.IntegralHeight = False
        Me.cboVoucherNo.Location = New System.Drawing.Point(413, 9)
        Me.cboVoucherNo.Name = "cboVoucherNo"
        Me.cboVoucherNo.Size = New System.Drawing.Size(155, 21)
        Me.cboVoucherNo.TabIndex = 32
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(340, 12)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(67, 13)
        Me.Label3.TabIndex = 31
        Me.Label3.Text = "Voucher No:"
        '
        'cboYear
        '
        Me.cboYear.DropDownHeight = 80
        Me.cboYear.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboYear.FormattingEnabled = True
        Me.cboYear.IntegralHeight = False
        Me.cboYear.Location = New System.Drawing.Point(227, 9)
        Me.cboYear.Name = "cboYear"
        Me.cboYear.Size = New System.Drawing.Size(96, 21)
        Me.cboYear.TabIndex = 30
        '
        'cboTerm
        '
        Me.cboTerm.DropDownHeight = 80
        Me.cboTerm.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboTerm.FormattingEnabled = True
        Me.cboTerm.IntegralHeight = False
        Me.cboTerm.Location = New System.Drawing.Point(53, 9)
        Me.cboTerm.Name = "cboTerm"
        Me.cboTerm.Size = New System.Drawing.Size(103, 21)
        Me.cboTerm.TabIndex = 29
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(189, 12)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(32, 13)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "Year:"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(13, 12)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(34, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Term:"
        '
        'StatusStrip3
        '
        Me.StatusStrip3.AutoSize = False
        Me.StatusStrip3.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip3.Dock = System.Windows.Forms.DockStyle.Left
        Me.StatusStrip3.Location = New System.Drawing.Point(0, 5)
        Me.StatusStrip3.Name = "StatusStrip3"
        Me.StatusStrip3.Size = New System.Drawing.Size(5, 70)
        Me.StatusStrip3.SizingGrip = False
        Me.StatusStrip3.TabIndex = 92
        Me.StatusStrip3.Text = "StatusStrip3"
        '
        'StatusStrip4
        '
        Me.StatusStrip4.AutoSize = False
        Me.StatusStrip4.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip4.Location = New System.Drawing.Point(0, 75)
        Me.StatusStrip4.Name = "StatusStrip4"
        Me.StatusStrip4.Size = New System.Drawing.Size(902, 5)
        Me.StatusStrip4.SizingGrip = False
        Me.StatusStrip4.TabIndex = 91
        Me.StatusStrip4.Text = "StatusStrip4"
        '
        'StatusStrip6
        '
        Me.StatusStrip6.AutoSize = False
        Me.StatusStrip6.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip6.Dock = System.Windows.Forms.DockStyle.Top
        Me.StatusStrip6.Location = New System.Drawing.Point(0, 0)
        Me.StatusStrip6.Name = "StatusStrip6"
        Me.StatusStrip6.Size = New System.Drawing.Size(902, 5)
        Me.StatusStrip6.SizingGrip = False
        Me.StatusStrip6.TabIndex = 90
        Me.StatusStrip6.Text = "StatusStrip6"
        '
        'StatusStrip7
        '
        Me.StatusStrip7.AutoSize = False
        Me.StatusStrip7.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip7.Dock = System.Windows.Forms.DockStyle.Right
        Me.StatusStrip7.Location = New System.Drawing.Point(902, 0)
        Me.StatusStrip7.Name = "StatusStrip7"
        Me.StatusStrip7.Size = New System.Drawing.Size(5, 80)
        Me.StatusStrip7.SizingGrip = False
        Me.StatusStrip7.TabIndex = 89
        Me.StatusStrip7.Text = "StatusStrip7"
        '
        'SplitContainer2
        '
        Me.SplitContainer2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.SplitContainer2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer2.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer2.Name = "SplitContainer2"
        Me.SplitContainer2.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'SplitContainer2.Panel1
        '
        Me.SplitContainer2.Panel1.Controls.Add(Me.lstPayApprovals)
        '
        'SplitContainer2.Panel2
        '
        Me.SplitContainer2.Panel2.Controls.Add(Me.Panel3)
        Me.SplitContainer2.Panel2.Controls.Add(Me.StatusStrip5)
        Me.SplitContainer2.Panel2.Controls.Add(Me.StatusStrip8)
        Me.SplitContainer2.Panel2.Controls.Add(Me.StatusStrip9)
        Me.SplitContainer2.Panel2.Controls.Add(Me.StatusStrip10)
        Me.SplitContainer2.Size = New System.Drawing.Size(911, 295)
        Me.SplitContainer2.SplitterDistance = 239
        Me.SplitContainer2.SplitterWidth = 1
        Me.SplitContainer2.TabIndex = 36
        '
        'lstPayApprovals
        '
        Me.lstPayApprovals.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColumnHeader1, Me.ColumnHeader2, Me.ColumnHeader3, Me.ColumnHeader4, Me.ColumnHeader5, Me.ColumnHeader6})
        Me.lstPayApprovals.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lstPayApprovals.FullRowSelect = True
        Me.lstPayApprovals.GridLines = True
        Me.lstPayApprovals.Location = New System.Drawing.Point(0, 0)
        Me.lstPayApprovals.Name = "lstPayApprovals"
        Me.lstPayApprovals.Size = New System.Drawing.Size(907, 235)
        Me.lstPayApprovals.TabIndex = 94
        Me.lstPayApprovals.UseCompatibleStateImageBehavior = False
        Me.lstPayApprovals.View = System.Windows.Forms.View.Details
        '
        'ColumnHeader1
        '
        Me.ColumnHeader1.Text = "Request Date"
        Me.ColumnHeader1.Width = 100
        '
        'ColumnHeader2
        '
        Me.ColumnHeader2.Text = "Request By"
        Me.ColumnHeader2.Width = 190
        '
        'ColumnHeader3
        '
        Me.ColumnHeader3.Text = "Payment To"
        Me.ColumnHeader3.Width = 190
        '
        'ColumnHeader4
        '
        Me.ColumnHeader4.Text = "Payment For"
        Me.ColumnHeader4.Width = 170
        '
        'ColumnHeader5
        '
        Me.ColumnHeader5.Text = "Account No"
        Me.ColumnHeader5.Width = 150
        '
        'ColumnHeader6
        '
        Me.ColumnHeader6.Text = "Amount"
        Me.ColumnHeader6.Width = 100
        '
        'Panel3
        '
        Me.Panel3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel3.Controls.Add(Me.btnApprove)
        Me.Panel3.Controls.Add(Me.btnClose)
        Me.Panel3.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel3.Location = New System.Drawing.Point(5, 5)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(897, 41)
        Me.Panel3.TabIndex = 89
        '
        'btnApprove
        '
        Me.btnApprove.Location = New System.Drawing.Point(583, 7)
        Me.btnApprove.Name = "btnApprove"
        Me.btnApprove.Size = New System.Drawing.Size(75, 23)
        Me.btnApprove.TabIndex = 27
        Me.btnApprove.Text = "Approve"
        Me.btnApprove.UseVisualStyleBackColor = True
        '
        'btnClose
        '
        Me.btnClose.Location = New System.Drawing.Point(805, 7)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(75, 23)
        Me.btnClose.TabIndex = 23
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'StatusStrip5
        '
        Me.StatusStrip5.AutoSize = False
        Me.StatusStrip5.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip5.Dock = System.Windows.Forms.DockStyle.Left
        Me.StatusStrip5.Location = New System.Drawing.Point(0, 5)
        Me.StatusStrip5.Name = "StatusStrip5"
        Me.StatusStrip5.Size = New System.Drawing.Size(5, 41)
        Me.StatusStrip5.SizingGrip = False
        Me.StatusStrip5.TabIndex = 88
        Me.StatusStrip5.Text = "StatusStrip5"
        '
        'StatusStrip8
        '
        Me.StatusStrip8.AutoSize = False
        Me.StatusStrip8.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip8.Location = New System.Drawing.Point(0, 46)
        Me.StatusStrip8.Name = "StatusStrip8"
        Me.StatusStrip8.Size = New System.Drawing.Size(902, 5)
        Me.StatusStrip8.SizingGrip = False
        Me.StatusStrip8.TabIndex = 87
        Me.StatusStrip8.Text = "StatusStrip8"
        '
        'StatusStrip9
        '
        Me.StatusStrip9.AutoSize = False
        Me.StatusStrip9.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip9.Dock = System.Windows.Forms.DockStyle.Top
        Me.StatusStrip9.Location = New System.Drawing.Point(0, 0)
        Me.StatusStrip9.Name = "StatusStrip9"
        Me.StatusStrip9.Size = New System.Drawing.Size(902, 5)
        Me.StatusStrip9.SizingGrip = False
        Me.StatusStrip9.TabIndex = 86
        Me.StatusStrip9.Text = "StatusStrip9"
        '
        'StatusStrip10
        '
        Me.StatusStrip10.AutoSize = False
        Me.StatusStrip10.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip10.Dock = System.Windows.Forms.DockStyle.Right
        Me.StatusStrip10.Location = New System.Drawing.Point(902, 0)
        Me.StatusStrip10.Name = "StatusStrip10"
        Me.StatusStrip10.Size = New System.Drawing.Size(5, 51)
        Me.StatusStrip10.SizingGrip = False
        Me.StatusStrip10.TabIndex = 85
        Me.StatusStrip10.Text = "StatusStrip10"
        '
        'StatusStrip2
        '
        Me.StatusStrip2.AutoSize = False
        Me.StatusStrip2.BackColor = System.Drawing.Color.LightSeaGreen
        Me.StatusStrip2.Dock = System.Windows.Forms.DockStyle.Top
        Me.StatusStrip2.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripStatusLabel1})
        Me.StatusStrip2.Location = New System.Drawing.Point(0, 0)
        Me.StatusStrip2.Name = "StatusStrip2"
        Me.StatusStrip2.Size = New System.Drawing.Size(911, 25)
        Me.StatusStrip2.SizingGrip = False
        Me.StatusStrip2.TabIndex = 34
        Me.StatusStrip2.Text = "StatusStrip2"
        '
        'ToolStripStatusLabel1
        '
        Me.ToolStripStatusLabel1.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ToolStripStatusLabel1.Name = "ToolStripStatusLabel1"
        Me.ToolStripStatusLabel1.Size = New System.Drawing.Size(139, 20)
        Me.ToolStripStatusLabel1.Text = "PAYMENT APPROVAL"
        '
        'StatusStrip1
        '
        Me.StatusStrip1.AutoSize = False
        Me.StatusStrip1.BackColor = System.Drawing.Color.LightSeaGreen
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 405)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(911, 25)
        Me.StatusStrip1.SizingGrip = False
        Me.StatusStrip1.TabIndex = 33
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'frmFinPayApproval
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.LightSkyBlue
        Me.ClientSize = New System.Drawing.Size(937, 456)
        Me.Controls.Add(Me.pnlPayApprove)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Name = "frmFinPayApproval"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Payment Approval"
        Me.pnlPayApprove.ResumeLayout(False)
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.ResumeLayout(False)
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.SplitContainer2.Panel1.ResumeLayout(False)
        Me.SplitContainer2.Panel2.ResumeLayout(False)
        CType(Me.SplitContainer2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer2.ResumeLayout(False)
        Me.Panel3.ResumeLayout(False)
        Me.StatusStrip2.ResumeLayout(False)
        Me.StatusStrip2.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents pnlPayApprove As System.Windows.Forms.Panel
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents SplitContainer2 As System.Windows.Forms.SplitContainer
    Friend WithEvents StatusStrip2 As System.Windows.Forms.StatusStrip
    Friend WithEvents ToolStripStatusLabel1 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents Panel3 As System.Windows.Forms.Panel
    Friend WithEvents btnApprove As System.Windows.Forms.Button
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents StatusStrip5 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip8 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip9 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip10 As System.Windows.Forms.StatusStrip
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents StatusStrip3 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip4 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip6 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip7 As System.Windows.Forms.StatusStrip
    Friend WithEvents lstPayApprovals As System.Windows.Forms.ListView
    Friend WithEvents cboVoucherNo As System.Windows.Forms.ComboBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents cboYear As System.Windows.Forms.ComboBox
    Friend WithEvents cboTerm As System.Windows.Forms.ComboBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents cboEmpNo As System.Windows.Forms.ComboBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents txtEmpName As System.Windows.Forms.TextBox
    Friend WithEvents dtpApprovalDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents btnLoad As System.Windows.Forms.Button
    Friend WithEvents ColumnHeader1 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader2 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader3 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader4 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader5 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader6 As System.Windows.Forms.ColumnHeader
End Class
