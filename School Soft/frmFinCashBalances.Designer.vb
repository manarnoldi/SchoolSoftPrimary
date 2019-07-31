<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmFinCashBalances
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
        Me.pnlCashBalances = New System.Windows.Forms.Panel()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.txtBalance = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.cboTransferTo = New System.Windows.Forms.ComboBox()
        Me.cboAccountNo = New System.Windows.Forms.ComboBox()
        Me.cboAccountName = New System.Windows.Forms.ComboBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.txtAmount = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.btnUpdate = New System.Windows.Forms.Button()
        Me.btnClose = New System.Windows.Forms.Button()
        Me.StatusStrip12 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip13 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip14 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip5 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip2 = New System.Windows.Forms.StatusStrip()
        Me.ToolStripStatusLabel1 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip()
        Me.txtVendorName = New System.Windows.Forms.TextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.pnlCashBalances.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.StatusStrip2.SuspendLayout()
        Me.SuspendLayout()
        '
        'pnlCashBalances
        '
        Me.pnlCashBalances.BackColor = System.Drawing.SystemColors.Control
        Me.pnlCashBalances.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pnlCashBalances.Controls.Add(Me.Panel2)
        Me.pnlCashBalances.Controls.Add(Me.StatusStrip2)
        Me.pnlCashBalances.Controls.Add(Me.StatusStrip1)
        Me.pnlCashBalances.Location = New System.Drawing.Point(11, 12)
        Me.pnlCashBalances.Name = "pnlCashBalances"
        Me.pnlCashBalances.Size = New System.Drawing.Size(438, 402)
        Me.pnlCashBalances.TabIndex = 4
        '
        'Panel2
        '
        Me.Panel2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel2.Controls.Add(Me.txtBalance)
        Me.Panel2.Controls.Add(Me.Label4)
        Me.Panel2.Controls.Add(Me.Panel1)
        Me.Panel2.Controls.Add(Me.StatusStrip12)
        Me.Panel2.Controls.Add(Me.StatusStrip13)
        Me.Panel2.Controls.Add(Me.StatusStrip14)
        Me.Panel2.Controls.Add(Me.StatusStrip5)
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel2.Location = New System.Drawing.Point(0, 25)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(436, 350)
        Me.Panel2.TabIndex = 62
        '
        'txtBalance
        '
        Me.txtBalance.Location = New System.Drawing.Point(111, 24)
        Me.txtBalance.Name = "txtBalance"
        Me.txtBalance.ReadOnly = True
        Me.txtBalance.Size = New System.Drawing.Size(296, 20)
        Me.txtBalance.TabIndex = 7
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(23, 27)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(82, 13)
        Me.Label4.TabIndex = 6
        Me.Label4.Text = "Balance ( Ksh ):"
        '
        'Panel1
        '
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel1.Controls.Add(Me.txtVendorName)
        Me.Panel1.Controls.Add(Me.Label6)
        Me.Panel1.Controls.Add(Me.Label5)
        Me.Panel1.Controls.Add(Me.cboTransferTo)
        Me.Panel1.Controls.Add(Me.cboAccountNo)
        Me.Panel1.Controls.Add(Me.cboAccountName)
        Me.Panel1.Controls.Add(Me.Label3)
        Me.Panel1.Controls.Add(Me.Label2)
        Me.Panel1.Controls.Add(Me.txtAmount)
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Controls.Add(Me.btnUpdate)
        Me.Panel1.Controls.Add(Me.btnClose)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel1.Location = New System.Drawing.Point(5, 5)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(424, 338)
        Me.Panel1.TabIndex = 65
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(17, 67)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(68, 13)
        Me.Label5.TabIndex = 8
        Me.Label5.Text = "Transfer To :"
        '
        'cboTransferTo
        '
        Me.cboTransferTo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboTransferTo.FormattingEnabled = True
        Me.cboTransferTo.Location = New System.Drawing.Point(105, 64)
        Me.cboTransferTo.Name = "cboTransferTo"
        Me.cboTransferTo.Size = New System.Drawing.Size(296, 21)
        Me.cboTransferTo.TabIndex = 11
        '
        'cboAccountNo
        '
        Me.cboAccountNo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboAccountNo.FormattingEnabled = True
        Me.cboAccountNo.Location = New System.Drawing.Point(105, 158)
        Me.cboAccountNo.Name = "cboAccountNo"
        Me.cboAccountNo.Size = New System.Drawing.Size(296, 21)
        Me.cboAccountNo.TabIndex = 71
        '
        'cboAccountName
        '
        Me.cboAccountName.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboAccountName.FormattingEnabled = True
        Me.cboAccountName.Location = New System.Drawing.Point(105, 110)
        Me.cboAccountName.Name = "cboAccountName"
        Me.cboAccountName.Size = New System.Drawing.Size(296, 21)
        Me.cboAccountName.TabIndex = 66
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(17, 161)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(67, 13)
        Me.Label3.TabIndex = 70
        Me.Label3.Text = "Account No:"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(17, 113)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(81, 13)
        Me.Label2.TabIndex = 68
        Me.Label2.Text = "Account Name:"
        '
        'txtAmount
        '
        Me.txtAmount.Location = New System.Drawing.Point(105, 250)
        Me.txtAmount.Name = "txtAmount"
        Me.txtAmount.Size = New System.Drawing.Size(296, 20)
        Me.txtAmount.TabIndex = 67
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(17, 253)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(79, 13)
        Me.Label1.TabIndex = 66
        Me.Label1.Text = "Amount ( Ksh ):"
        '
        'btnUpdate
        '
        Me.btnUpdate.Location = New System.Drawing.Point(105, 292)
        Me.btnUpdate.Name = "btnUpdate"
        Me.btnUpdate.Size = New System.Drawing.Size(75, 23)
        Me.btnUpdate.TabIndex = 4
        Me.btnUpdate.Text = "Update"
        Me.btnUpdate.UseVisualStyleBackColor = True
        '
        'btnClose
        '
        Me.btnClose.Location = New System.Drawing.Point(326, 292)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(75, 23)
        Me.btnClose.TabIndex = 10
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'StatusStrip12
        '
        Me.StatusStrip12.AutoSize = False
        Me.StatusStrip12.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip12.Dock = System.Windows.Forms.DockStyle.Left
        Me.StatusStrip12.Location = New System.Drawing.Point(0, 5)
        Me.StatusStrip12.Name = "StatusStrip12"
        Me.StatusStrip12.Size = New System.Drawing.Size(5, 338)
        Me.StatusStrip12.SizingGrip = False
        Me.StatusStrip12.TabIndex = 64
        Me.StatusStrip12.Text = "StatusStrip12"
        '
        'StatusStrip13
        '
        Me.StatusStrip13.AutoSize = False
        Me.StatusStrip13.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip13.Location = New System.Drawing.Point(0, 343)
        Me.StatusStrip13.Name = "StatusStrip13"
        Me.StatusStrip13.Size = New System.Drawing.Size(429, 5)
        Me.StatusStrip13.SizingGrip = False
        Me.StatusStrip13.TabIndex = 63
        Me.StatusStrip13.Text = "StatusStrip13"
        '
        'StatusStrip14
        '
        Me.StatusStrip14.AutoSize = False
        Me.StatusStrip14.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip14.Dock = System.Windows.Forms.DockStyle.Top
        Me.StatusStrip14.Location = New System.Drawing.Point(0, 0)
        Me.StatusStrip14.Name = "StatusStrip14"
        Me.StatusStrip14.Size = New System.Drawing.Size(429, 5)
        Me.StatusStrip14.SizingGrip = False
        Me.StatusStrip14.TabIndex = 62
        Me.StatusStrip14.Text = "StatusStrip14"
        '
        'StatusStrip5
        '
        Me.StatusStrip5.AutoSize = False
        Me.StatusStrip5.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip5.Dock = System.Windows.Forms.DockStyle.Right
        Me.StatusStrip5.Location = New System.Drawing.Point(429, 0)
        Me.StatusStrip5.Name = "StatusStrip5"
        Me.StatusStrip5.Size = New System.Drawing.Size(5, 348)
        Me.StatusStrip5.SizingGrip = False
        Me.StatusStrip5.TabIndex = 61
        Me.StatusStrip5.Text = "StatusStrip5"
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
        Me.StatusStrip2.TabIndex = 22
        Me.StatusStrip2.Text = "StatusStrip2"
        '
        'ToolStripStatusLabel1
        '
        Me.ToolStripStatusLabel1.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ToolStripStatusLabel1.Name = "ToolStripStatusLabel1"
        Me.ToolStripStatusLabel1.Size = New System.Drawing.Size(111, 20)
        Me.ToolStripStatusLabel1.Text = "CASH BALANCES"
        '
        'StatusStrip1
        '
        Me.StatusStrip1.AutoSize = False
        Me.StatusStrip1.BackColor = System.Drawing.Color.LightSeaGreen
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 375)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(436, 25)
        Me.StatusStrip1.SizingGrip = False
        Me.StatusStrip1.TabIndex = 21
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'txtVendorName
        '
        Me.txtVendorName.Location = New System.Drawing.Point(105, 205)
        Me.txtVendorName.Name = "txtVendorName"
        Me.txtVendorName.ReadOnly = True
        Me.txtVendorName.Size = New System.Drawing.Size(296, 20)
        Me.txtVendorName.TabIndex = 73
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(17, 208)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(75, 13)
        Me.Label6.TabIndex = 72
        Me.Label6.Text = "Vendor Name:"
        '
        'frmFinCashBalances
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.LightSkyBlue
        Me.ClientSize = New System.Drawing.Size(463, 426)
        Me.Controls.Add(Me.pnlCashBalances)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Name = "frmFinCashBalances"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Cash Balances"
        Me.pnlCashBalances.ResumeLayout(False)
        Me.Panel2.ResumeLayout(False)
        Me.Panel2.PerformLayout()
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.StatusStrip2.ResumeLayout(False)
        Me.StatusStrip2.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents pnlCashBalances As System.Windows.Forms.Panel
    Friend WithEvents StatusStrip2 As System.Windows.Forms.StatusStrip
    Friend WithEvents ToolStripStatusLabel1 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents cboTransferTo As System.Windows.Forms.ComboBox
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents btnUpdate As System.Windows.Forms.Button
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents txtBalance As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents StatusStrip12 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip13 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip14 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip5 As System.Windows.Forms.StatusStrip
    Friend WithEvents cboAccountNo As System.Windows.Forms.ComboBox
    Friend WithEvents cboAccountName As System.Windows.Forms.ComboBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtAmount As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtVendorName As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
End Class
