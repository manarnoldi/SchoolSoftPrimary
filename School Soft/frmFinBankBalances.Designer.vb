<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmFinBankBalances
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
        Me.pnlBankBalances = New System.Windows.Forms.Panel()
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.cboAccountNo = New System.Windows.Forms.ComboBox()
        Me.cboAccountName = New System.Windows.Forms.ComboBox()
        Me.cboPhoneNo = New System.Windows.Forms.ComboBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.txtAmount = New System.Windows.Forms.TextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.cboTransferTo = New System.Windows.Forms.ComboBox()
        Me.btnClose = New System.Windows.Forms.Button()
        Me.btnUpdate = New System.Windows.Forms.Button()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.txtBalance = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.cboBankName = New System.Windows.Forms.ComboBox()
        Me.StatusStrip12 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip13 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip14 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip5 = New System.Windows.Forms.StatusStrip()
        Me.lstBankBalances = New System.Windows.Forms.ListView()
        Me.ColumnHeader1 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader4 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader2 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader3 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.StatusStrip2 = New System.Windows.Forms.StatusStrip()
        Me.ToolStripStatusLabel1 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip()
        Me.pnlBankBalances.SuspendLayout()
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.StatusStrip2.SuspendLayout()
        Me.SuspendLayout()
        '
        'pnlBankBalances
        '
        Me.pnlBankBalances.BackColor = System.Drawing.SystemColors.Control
        Me.pnlBankBalances.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pnlBankBalances.Controls.Add(Me.SplitContainer1)
        Me.pnlBankBalances.Controls.Add(Me.StatusStrip2)
        Me.pnlBankBalances.Controls.Add(Me.StatusStrip1)
        Me.pnlBankBalances.Location = New System.Drawing.Point(12, 12)
        Me.pnlBankBalances.Name = "pnlBankBalances"
        Me.pnlBankBalances.Size = New System.Drawing.Size(702, 485)
        Me.pnlBankBalances.TabIndex = 3
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer1.Location = New System.Drawing.Point(0, 25)
        Me.SplitContainer1.Name = "SplitContainer1"
        Me.SplitContainer1.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.Panel2)
        Me.SplitContainer1.Panel1.Controls.Add(Me.StatusStrip12)
        Me.SplitContainer1.Panel1.Controls.Add(Me.StatusStrip13)
        Me.SplitContainer1.Panel1.Controls.Add(Me.StatusStrip14)
        Me.SplitContainer1.Panel1.Controls.Add(Me.StatusStrip5)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.lstBankBalances)
        Me.SplitContainer1.Size = New System.Drawing.Size(700, 433)
        Me.SplitContainer1.SplitterDistance = 161
        Me.SplitContainer1.SplitterWidth = 1
        Me.SplitContainer1.TabIndex = 23
        '
        'Panel2
        '
        Me.Panel2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel2.Controls.Add(Me.cboAccountNo)
        Me.Panel2.Controls.Add(Me.cboAccountName)
        Me.Panel2.Controls.Add(Me.cboPhoneNo)
        Me.Panel2.Controls.Add(Me.Label7)
        Me.Panel2.Controls.Add(Me.txtAmount)
        Me.Panel2.Controls.Add(Me.Label6)
        Me.Panel2.Controls.Add(Me.cboTransferTo)
        Me.Panel2.Controls.Add(Me.btnClose)
        Me.Panel2.Controls.Add(Me.btnUpdate)
        Me.Panel2.Controls.Add(Me.Label5)
        Me.Panel2.Controls.Add(Me.txtBalance)
        Me.Panel2.Controls.Add(Me.Label4)
        Me.Panel2.Controls.Add(Me.Label3)
        Me.Panel2.Controls.Add(Me.Label2)
        Me.Panel2.Controls.Add(Me.Label1)
        Me.Panel2.Controls.Add(Me.cboBankName)
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel2.Location = New System.Drawing.Point(5, 5)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(690, 151)
        Me.Panel2.TabIndex = 61
        '
        'cboAccountNo
        '
        Me.cboAccountNo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboAccountNo.FormattingEnabled = True
        Me.cboAccountNo.Location = New System.Drawing.Point(109, 79)
        Me.cboAccountNo.Name = "cboAccountNo"
        Me.cboAccountNo.Size = New System.Drawing.Size(294, 21)
        Me.cboAccountNo.TabIndex = 17
        '
        'cboAccountName
        '
        Me.cboAccountName.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboAccountName.FormattingEnabled = True
        Me.cboAccountName.Location = New System.Drawing.Point(109, 46)
        Me.cboAccountName.Name = "cboAccountName"
        Me.cboAccountName.Size = New System.Drawing.Size(294, 21)
        Me.cboAccountName.TabIndex = 16
        '
        'cboPhoneNo
        '
        Me.cboPhoneNo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboPhoneNo.Enabled = False
        Me.cboPhoneNo.FormattingEnabled = True
        Me.cboPhoneNo.Location = New System.Drawing.Point(508, 79)
        Me.cboPhoneNo.Name = "cboPhoneNo"
        Me.cboPhoneNo.Size = New System.Drawing.Size(167, 21)
        Me.cboPhoneNo.TabIndex = 15
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(420, 82)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(81, 13)
        Me.Label7.TabIndex = 14
        Me.Label7.Text = "Phone Number:"
        '
        'txtAmount
        '
        Me.txtAmount.Location = New System.Drawing.Point(508, 114)
        Me.txtAmount.Name = "txtAmount"
        Me.txtAmount.Size = New System.Drawing.Size(167, 20)
        Me.txtAmount.TabIndex = 13
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(420, 117)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(82, 13)
        Me.Label6.TabIndex = 12
        Me.Label6.Text = "Amount ( Ksh ) :"
        '
        'cboTransferTo
        '
        Me.cboTransferTo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboTransferTo.FormattingEnabled = True
        Me.cboTransferTo.Location = New System.Drawing.Point(508, 46)
        Me.cboTransferTo.Name = "cboTransferTo"
        Me.cboTransferTo.Size = New System.Drawing.Size(167, 21)
        Me.cboTransferTo.TabIndex = 11
        '
        'btnClose
        '
        Me.btnClose.Location = New System.Drawing.Point(328, 112)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(75, 23)
        Me.btnClose.TabIndex = 10
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'btnUpdate
        '
        Me.btnUpdate.Location = New System.Drawing.Point(109, 112)
        Me.btnUpdate.Name = "btnUpdate"
        Me.btnUpdate.Size = New System.Drawing.Size(75, 23)
        Me.btnUpdate.TabIndex = 4
        Me.btnUpdate.Text = "Update"
        Me.btnUpdate.UseVisualStyleBackColor = True
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(420, 49)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(68, 13)
        Me.Label5.TabIndex = 8
        Me.Label5.Text = "Transfer To :"
        '
        'txtBalance
        '
        Me.txtBalance.Location = New System.Drawing.Point(508, 12)
        Me.txtBalance.Name = "txtBalance"
        Me.txtBalance.ReadOnly = True
        Me.txtBalance.Size = New System.Drawing.Size(167, 20)
        Me.txtBalance.TabIndex = 7
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(420, 15)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(82, 13)
        Me.Label4.TabIndex = 6
        Me.Label4.Text = "Balance ( Ksh ):"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(13, 82)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(90, 13)
        Me.Label3.TabIndex = 3
        Me.Label3.Text = "Account Number:"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(13, 49)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(81, 13)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "Account Name:"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(13, 15)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(66, 13)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Bank Name:"
        '
        'cboBankName
        '
        Me.cboBankName.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboBankName.FormattingEnabled = True
        Me.cboBankName.Location = New System.Drawing.Point(109, 12)
        Me.cboBankName.Name = "cboBankName"
        Me.cboBankName.Size = New System.Drawing.Size(294, 21)
        Me.cboBankName.TabIndex = 0
        '
        'StatusStrip12
        '
        Me.StatusStrip12.AutoSize = False
        Me.StatusStrip12.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip12.Dock = System.Windows.Forms.DockStyle.Left
        Me.StatusStrip12.Location = New System.Drawing.Point(0, 5)
        Me.StatusStrip12.Name = "StatusStrip12"
        Me.StatusStrip12.Size = New System.Drawing.Size(5, 151)
        Me.StatusStrip12.SizingGrip = False
        Me.StatusStrip12.TabIndex = 60
        Me.StatusStrip12.Text = "StatusStrip12"
        '
        'StatusStrip13
        '
        Me.StatusStrip13.AutoSize = False
        Me.StatusStrip13.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip13.Location = New System.Drawing.Point(0, 156)
        Me.StatusStrip13.Name = "StatusStrip13"
        Me.StatusStrip13.Size = New System.Drawing.Size(695, 5)
        Me.StatusStrip13.SizingGrip = False
        Me.StatusStrip13.TabIndex = 59
        Me.StatusStrip13.Text = "StatusStrip13"
        '
        'StatusStrip14
        '
        Me.StatusStrip14.AutoSize = False
        Me.StatusStrip14.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip14.Dock = System.Windows.Forms.DockStyle.Top
        Me.StatusStrip14.Location = New System.Drawing.Point(0, 0)
        Me.StatusStrip14.Name = "StatusStrip14"
        Me.StatusStrip14.Size = New System.Drawing.Size(695, 5)
        Me.StatusStrip14.SizingGrip = False
        Me.StatusStrip14.TabIndex = 58
        Me.StatusStrip14.Text = "StatusStrip14"
        '
        'StatusStrip5
        '
        Me.StatusStrip5.AutoSize = False
        Me.StatusStrip5.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip5.Dock = System.Windows.Forms.DockStyle.Right
        Me.StatusStrip5.Location = New System.Drawing.Point(695, 0)
        Me.StatusStrip5.Name = "StatusStrip5"
        Me.StatusStrip5.Size = New System.Drawing.Size(5, 161)
        Me.StatusStrip5.SizingGrip = False
        Me.StatusStrip5.TabIndex = 57
        Me.StatusStrip5.Text = "StatusStrip5"
        '
        'lstBankBalances
        '
        Me.lstBankBalances.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColumnHeader1, Me.ColumnHeader4, Me.ColumnHeader2, Me.ColumnHeader3})
        Me.lstBankBalances.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lstBankBalances.FullRowSelect = True
        Me.lstBankBalances.GridLines = True
        Me.lstBankBalances.Location = New System.Drawing.Point(0, 0)
        Me.lstBankBalances.Name = "lstBankBalances"
        Me.lstBankBalances.Size = New System.Drawing.Size(700, 271)
        Me.lstBankBalances.TabIndex = 27
        Me.lstBankBalances.UseCompatibleStateImageBehavior = False
        Me.lstBankBalances.View = System.Windows.Forms.View.Details
        '
        'ColumnHeader1
        '
        Me.ColumnHeader1.Text = "Bank Name"
        Me.ColumnHeader1.Width = 200
        '
        'ColumnHeader4
        '
        Me.ColumnHeader4.Text = "Account Name"
        Me.ColumnHeader4.Width = 200
        '
        'ColumnHeader2
        '
        Me.ColumnHeader2.Text = "Account Number"
        Me.ColumnHeader2.Width = 170
        '
        'ColumnHeader3
        '
        Me.ColumnHeader3.Text = "Balance"
        Me.ColumnHeader3.Width = 120
        '
        'StatusStrip2
        '
        Me.StatusStrip2.AutoSize = False
        Me.StatusStrip2.BackColor = System.Drawing.Color.LightSeaGreen
        Me.StatusStrip2.Dock = System.Windows.Forms.DockStyle.Top
        Me.StatusStrip2.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripStatusLabel1})
        Me.StatusStrip2.Location = New System.Drawing.Point(0, 0)
        Me.StatusStrip2.Name = "StatusStrip2"
        Me.StatusStrip2.Size = New System.Drawing.Size(700, 25)
        Me.StatusStrip2.SizingGrip = False
        Me.StatusStrip2.TabIndex = 22
        Me.StatusStrip2.Text = "StatusStrip2"
        '
        'ToolStripStatusLabel1
        '
        Me.ToolStripStatusLabel1.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ToolStripStatusLabel1.Name = "ToolStripStatusLabel1"
        Me.ToolStripStatusLabel1.Size = New System.Drawing.Size(112, 20)
        Me.ToolStripStatusLabel1.Text = "BANK BALANCES"
        '
        'StatusStrip1
        '
        Me.StatusStrip1.AutoSize = False
        Me.StatusStrip1.BackColor = System.Drawing.Color.LightSeaGreen
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 458)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(700, 25)
        Me.StatusStrip1.SizingGrip = False
        Me.StatusStrip1.TabIndex = 21
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'frmFinBankBalances
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.SkyBlue
        Me.ClientSize = New System.Drawing.Size(725, 508)
        Me.Controls.Add(Me.pnlBankBalances)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Name = "frmFinBankBalances"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Bank Balances"
        Me.pnlBankBalances.ResumeLayout(False)
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.ResumeLayout(False)
        Me.Panel2.ResumeLayout(False)
        Me.Panel2.PerformLayout()
        Me.StatusStrip2.ResumeLayout(False)
        Me.StatusStrip2.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents pnlBankBalances As System.Windows.Forms.Panel
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents cboBankName As System.Windows.Forms.ComboBox
    Friend WithEvents StatusStrip12 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip13 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip14 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip5 As System.Windows.Forms.StatusStrip
    Friend WithEvents btnUpdate As System.Windows.Forms.Button
    Friend WithEvents StatusStrip2 As System.Windows.Forms.StatusStrip
    Friend WithEvents ToolStripStatusLabel1 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents txtBalance As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents cboTransferTo As System.Windows.Forms.ComboBox
    Friend WithEvents cboPhoneNo As System.Windows.Forms.ComboBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents txtAmount As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents lstBankBalances As System.Windows.Forms.ListView
    Friend WithEvents ColumnHeader1 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader4 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader2 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader3 As System.Windows.Forms.ColumnHeader
    Friend WithEvents cboAccountNo As System.Windows.Forms.ComboBox
    Friend WithEvents cboAccountName As System.Windows.Forms.ComboBox
End Class
