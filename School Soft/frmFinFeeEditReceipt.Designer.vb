<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmFinFeeEditReceipt
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
        Me.pnlEditReceipt = New System.Windows.Forms.Panel()
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
        Me.SplitContainer2 = New System.Windows.Forms.SplitContainer()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.btnLoad = New System.Windows.Forms.Button()
        Me.cboReceiptNo = New System.Windows.Forms.ComboBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.cboYear = New System.Windows.Forms.ComboBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.cboTermName = New System.Windows.Forms.ComboBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.cboMonth = New System.Windows.Forms.ComboBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.StatusStrip3 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip4 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip6 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip7 = New System.Windows.Forms.StatusStrip()
        Me.crtViewResultsSummary = New CrystalDecisions.Windows.Forms.CrystalReportViewer()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.txtReverseReason = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.btnClose = New System.Windows.Forms.Button()
        Me.btnReverse = New System.Windows.Forms.Button()
        Me.StatusStrip5 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip8 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip9 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip10 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip2 = New System.Windows.Forms.StatusStrip()
        Me.ToolStripStatusLabel1 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip()
        Me.pnlEditReceipt.SuspendLayout()
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        CType(Me.SplitContainer2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer2.Panel1.SuspendLayout()
        Me.SplitContainer2.Panel2.SuspendLayout()
        Me.SplitContainer2.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.StatusStrip2.SuspendLayout()
        Me.SuspendLayout()
        '
        'pnlEditReceipt
        '
        Me.pnlEditReceipt.BackColor = System.Drawing.SystemColors.Control
        Me.pnlEditReceipt.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pnlEditReceipt.Controls.Add(Me.SplitContainer1)
        Me.pnlEditReceipt.Controls.Add(Me.StatusStrip2)
        Me.pnlEditReceipt.Controls.Add(Me.StatusStrip1)
        Me.pnlEditReceipt.Location = New System.Drawing.Point(12, 12)
        Me.pnlEditReceipt.Name = "pnlEditReceipt"
        Me.pnlEditReceipt.Size = New System.Drawing.Size(974, 577)
        Me.pnlEditReceipt.TabIndex = 0
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
        Me.SplitContainer1.Panel1.Controls.Add(Me.SplitContainer2)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.Panel2)
        Me.SplitContainer1.Panel2.Controls.Add(Me.StatusStrip5)
        Me.SplitContainer1.Panel2.Controls.Add(Me.StatusStrip8)
        Me.SplitContainer1.Panel2.Controls.Add(Me.StatusStrip9)
        Me.SplitContainer1.Panel2.Controls.Add(Me.StatusStrip10)
        Me.SplitContainer1.Size = New System.Drawing.Size(972, 525)
        Me.SplitContainer1.SplitterDistance = 470
        Me.SplitContainer1.SplitterWidth = 1
        Me.SplitContainer1.TabIndex = 40
        '
        'SplitContainer2
        '
        Me.SplitContainer2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer2.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer2.Name = "SplitContainer2"
        Me.SplitContainer2.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'SplitContainer2.Panel1
        '
        Me.SplitContainer2.Panel1.Controls.Add(Me.Panel1)
        Me.SplitContainer2.Panel1.Controls.Add(Me.StatusStrip3)
        Me.SplitContainer2.Panel1.Controls.Add(Me.StatusStrip4)
        Me.SplitContainer2.Panel1.Controls.Add(Me.StatusStrip6)
        Me.SplitContainer2.Panel1.Controls.Add(Me.StatusStrip7)
        '
        'SplitContainer2.Panel2
        '
        Me.SplitContainer2.Panel2.Controls.Add(Me.crtViewResultsSummary)
        Me.SplitContainer2.Size = New System.Drawing.Size(972, 470)
        Me.SplitContainer2.SplitterDistance = 52
        Me.SplitContainer2.SplitterWidth = 1
        Me.SplitContainer2.TabIndex = 41
        '
        'Panel1
        '
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel1.Controls.Add(Me.btnLoad)
        Me.Panel1.Controls.Add(Me.cboReceiptNo)
        Me.Panel1.Controls.Add(Me.Label2)
        Me.Panel1.Controls.Add(Me.cboYear)
        Me.Panel1.Controls.Add(Me.Label3)
        Me.Panel1.Controls.Add(Me.cboTermName)
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Controls.Add(Me.cboMonth)
        Me.Panel1.Controls.Add(Me.Label4)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel1.Location = New System.Drawing.Point(5, 5)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(962, 42)
        Me.Panel1.TabIndex = 73
        '
        'btnLoad
        '
        Me.btnLoad.Location = New System.Drawing.Point(875, 8)
        Me.btnLoad.Name = "btnLoad"
        Me.btnLoad.Size = New System.Drawing.Size(75, 23)
        Me.btnLoad.TabIndex = 16
        Me.btnLoad.Text = "View"
        Me.btnLoad.UseVisualStyleBackColor = True
        '
        'cboReceiptNo
        '
        Me.cboReceiptNo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboReceiptNo.FormattingEnabled = True
        Me.cboReceiptNo.Location = New System.Drawing.Point(641, 10)
        Me.cboReceiptNo.Name = "cboReceiptNo"
        Me.cboReceiptNo.Size = New System.Drawing.Size(178, 21)
        Me.cboReceiptNo.TabIndex = 15
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(571, 13)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(64, 13)
        Me.Label2.TabIndex = 14
        Me.Label2.Text = "Receipt No:"
        '
        'cboYear
        '
        Me.cboYear.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboYear.FormattingEnabled = True
        Me.cboYear.Location = New System.Drawing.Point(48, 10)
        Me.cboYear.Name = "cboYear"
        Me.cboYear.Size = New System.Drawing.Size(98, 21)
        Me.cboYear.TabIndex = 13
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(7, 13)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(35, 13)
        Me.Label3.TabIndex = 12
        Me.Label3.Text = "Year :"
        '
        'cboTermName
        '
        Me.cboTermName.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboTermName.FormattingEnabled = True
        Me.cboTermName.Location = New System.Drawing.Point(245, 10)
        Me.cboTermName.Name = "cboTermName"
        Me.cboTermName.Size = New System.Drawing.Size(102, 21)
        Me.cboTermName.TabIndex = 11
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(174, 13)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(65, 13)
        Me.Label1.TabIndex = 10
        Me.Label1.Text = "Term Name:"
        '
        'cboMonth
        '
        Me.cboMonth.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboMonth.FormattingEnabled = True
        Me.cboMonth.Location = New System.Drawing.Point(417, 10)
        Me.cboMonth.Name = "cboMonth"
        Me.cboMonth.Size = New System.Drawing.Size(139, 21)
        Me.cboMonth.TabIndex = 9
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(371, 13)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(40, 13)
        Me.Label4.TabIndex = 8
        Me.Label4.Text = "Month:"
        '
        'StatusStrip3
        '
        Me.StatusStrip3.AutoSize = False
        Me.StatusStrip3.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip3.Dock = System.Windows.Forms.DockStyle.Left
        Me.StatusStrip3.Location = New System.Drawing.Point(0, 5)
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
        Me.StatusStrip4.Location = New System.Drawing.Point(0, 47)
        Me.StatusStrip4.Name = "StatusStrip4"
        Me.StatusStrip4.Size = New System.Drawing.Size(967, 5)
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
        Me.StatusStrip6.Size = New System.Drawing.Size(967, 5)
        Me.StatusStrip6.SizingGrip = False
        Me.StatusStrip6.TabIndex = 70
        Me.StatusStrip6.Text = "StatusStrip6"
        '
        'StatusStrip7
        '
        Me.StatusStrip7.AutoSize = False
        Me.StatusStrip7.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip7.Dock = System.Windows.Forms.DockStyle.Right
        Me.StatusStrip7.Location = New System.Drawing.Point(967, 0)
        Me.StatusStrip7.Name = "StatusStrip7"
        Me.StatusStrip7.Size = New System.Drawing.Size(5, 52)
        Me.StatusStrip7.SizingGrip = False
        Me.StatusStrip7.TabIndex = 69
        Me.StatusStrip7.Text = "StatusStrip7"
        '
        'crtViewResultsSummary
        '
        Me.crtViewResultsSummary.ActiveViewIndex = -1
        Me.crtViewResultsSummary.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.crtViewResultsSummary.Cursor = System.Windows.Forms.Cursors.Default
        Me.crtViewResultsSummary.DisplayStatusBar = False
        Me.crtViewResultsSummary.Dock = System.Windows.Forms.DockStyle.Fill
        Me.crtViewResultsSummary.Location = New System.Drawing.Point(0, 0)
        Me.crtViewResultsSummary.Name = "crtViewResultsSummary"
        Me.crtViewResultsSummary.ShowGroupTreeButton = False
        Me.crtViewResultsSummary.ShowLogo = False
        Me.crtViewResultsSummary.ShowParameterPanelButton = False
        Me.crtViewResultsSummary.ShowTextSearchButton = False
        Me.crtViewResultsSummary.Size = New System.Drawing.Size(972, 417)
        Me.crtViewResultsSummary.TabIndex = 36
        Me.crtViewResultsSummary.ToolPanelView = CrystalDecisions.Windows.Forms.ToolPanelViewType.None
        '
        'Panel2
        '
        Me.Panel2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel2.Controls.Add(Me.txtReverseReason)
        Me.Panel2.Controls.Add(Me.Label5)
        Me.Panel2.Controls.Add(Me.btnClose)
        Me.Panel2.Controls.Add(Me.btnReverse)
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel2.Location = New System.Drawing.Point(5, 5)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(962, 44)
        Me.Panel2.TabIndex = 74
        '
        'txtReverseReason
        '
        Me.txtReverseReason.Location = New System.Drawing.Point(103, 11)
        Me.txtReverseReason.Name = "txtReverseReason"
        Me.txtReverseReason.Size = New System.Drawing.Size(480, 20)
        Me.txtReverseReason.TabIndex = 22
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(7, 14)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(90, 13)
        Me.Label5.TabIndex = 21
        Me.Label5.Text = "Reverse Reason:"
        '
        'btnClose
        '
        Me.btnClose.Location = New System.Drawing.Point(875, 9)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(75, 23)
        Me.btnClose.TabIndex = 20
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'btnReverse
        '
        Me.btnReverse.Location = New System.Drawing.Point(626, 9)
        Me.btnReverse.Name = "btnReverse"
        Me.btnReverse.Size = New System.Drawing.Size(75, 23)
        Me.btnReverse.TabIndex = 19
        Me.btnReverse.Text = "Reverse"
        Me.btnReverse.UseVisualStyleBackColor = True
        '
        'StatusStrip5
        '
        Me.StatusStrip5.AutoSize = False
        Me.StatusStrip5.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip5.Dock = System.Windows.Forms.DockStyle.Left
        Me.StatusStrip5.Location = New System.Drawing.Point(0, 5)
        Me.StatusStrip5.Name = "StatusStrip5"
        Me.StatusStrip5.Size = New System.Drawing.Size(5, 44)
        Me.StatusStrip5.SizingGrip = False
        Me.StatusStrip5.TabIndex = 72
        Me.StatusStrip5.Text = "StatusStrip5"
        '
        'StatusStrip8
        '
        Me.StatusStrip8.AutoSize = False
        Me.StatusStrip8.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip8.Location = New System.Drawing.Point(0, 49)
        Me.StatusStrip8.Name = "StatusStrip8"
        Me.StatusStrip8.Size = New System.Drawing.Size(967, 5)
        Me.StatusStrip8.SizingGrip = False
        Me.StatusStrip8.TabIndex = 71
        Me.StatusStrip8.Text = "StatusStrip8"
        '
        'StatusStrip9
        '
        Me.StatusStrip9.AutoSize = False
        Me.StatusStrip9.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip9.Dock = System.Windows.Forms.DockStyle.Top
        Me.StatusStrip9.Location = New System.Drawing.Point(0, 0)
        Me.StatusStrip9.Name = "StatusStrip9"
        Me.StatusStrip9.Size = New System.Drawing.Size(967, 5)
        Me.StatusStrip9.SizingGrip = False
        Me.StatusStrip9.TabIndex = 70
        Me.StatusStrip9.Text = "StatusStrip9"
        '
        'StatusStrip10
        '
        Me.StatusStrip10.AutoSize = False
        Me.StatusStrip10.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip10.Dock = System.Windows.Forms.DockStyle.Right
        Me.StatusStrip10.Location = New System.Drawing.Point(967, 0)
        Me.StatusStrip10.Name = "StatusStrip10"
        Me.StatusStrip10.Size = New System.Drawing.Size(5, 54)
        Me.StatusStrip10.SizingGrip = False
        Me.StatusStrip10.TabIndex = 69
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
        Me.StatusStrip2.Size = New System.Drawing.Size(972, 25)
        Me.StatusStrip2.SizingGrip = False
        Me.StatusStrip2.TabIndex = 38
        Me.StatusStrip2.Text = "StatusStrip2"
        '
        'ToolStripStatusLabel1
        '
        Me.ToolStripStatusLabel1.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ToolStripStatusLabel1.Name = "ToolStripStatusLabel1"
        Me.ToolStripStatusLabel1.Size = New System.Drawing.Size(118, 20)
        Me.ToolStripStatusLabel1.Text = "FIND FEE RECEIPT"
        '
        'StatusStrip1
        '
        Me.StatusStrip1.AutoSize = False
        Me.StatusStrip1.BackColor = System.Drawing.Color.LightSeaGreen
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 550)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(972, 25)
        Me.StatusStrip1.SizingGrip = False
        Me.StatusStrip1.TabIndex = 39
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'frmFinFeeEditReceipt
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.LightSkyBlue
        Me.ClientSize = New System.Drawing.Size(998, 600)
        Me.Controls.Add(Me.pnlEditReceipt)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Name = "frmFinFeeEditReceipt"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Find Receipt"
        Me.pnlEditReceipt.ResumeLayout(False)
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
        Me.Panel2.ResumeLayout(False)
        Me.Panel2.PerformLayout()
        Me.StatusStrip2.ResumeLayout(False)
        Me.StatusStrip2.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents pnlEditReceipt As System.Windows.Forms.Panel
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents SplitContainer2 As System.Windows.Forms.SplitContainer
    Friend WithEvents StatusStrip2 As System.Windows.Forms.StatusStrip
    Friend WithEvents ToolStripStatusLabel1 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip3 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip4 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip6 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip7 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip5 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip8 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip9 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip10 As System.Windows.Forms.StatusStrip
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents crtViewResultsSummary As CrystalDecisions.Windows.Forms.CrystalReportViewer
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents btnLoad As System.Windows.Forms.Button
    Friend WithEvents cboReceiptNo As System.Windows.Forms.ComboBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents cboYear As System.Windows.Forms.ComboBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents cboTermName As System.Windows.Forms.ComboBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents cboMonth As System.Windows.Forms.ComboBox
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents btnReverse As System.Windows.Forms.Button
    Friend WithEvents txtReverseReason As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
End Class
