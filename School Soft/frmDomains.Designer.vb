<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmDomains
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
        Me.btnView = New System.Windows.Forms.Button()
        Me.pnlDomains = New System.Windows.Forms.Panel()
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
        Me.btnSave = New System.Windows.Forms.Button()
        Me.txtDomName = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.StatusStrip5 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip4 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip3 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip10 = New System.Windows.Forms.StatusStrip()
        Me.SplitContainer2 = New System.Windows.Forms.SplitContainer()
        Me.lstDomains = New System.Windows.Forms.ListView()
        Me.ColumnHeader1 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.btnDelete = New System.Windows.Forms.Button()
        Me.btnUpdate = New System.Windows.Forms.Button()
        Me.btnExit = New System.Windows.Forms.Button()
        Me.StatusStrip6 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip7 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip8 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip9 = New System.Windows.Forms.StatusStrip()
        Me.ToolStripStatusLabel2 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.StatusStrip2 = New System.Windows.Forms.StatusStrip()
        Me.ToolStripStatusLabel1 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip()
        Me.pnlDomains.SuspendLayout()
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        CType(Me.SplitContainer2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer2.Panel1.SuspendLayout()
        Me.SplitContainer2.Panel2.SuspendLayout()
        Me.SplitContainer2.SuspendLayout()
        Me.Panel3.SuspendLayout()
        Me.StatusStrip9.SuspendLayout()
        Me.StatusStrip2.SuspendLayout()
        Me.SuspendLayout()
        '
        'btnView
        '
        Me.btnView.Location = New System.Drawing.Point(5, 13)
        Me.btnView.Name = "btnView"
        Me.btnView.Size = New System.Drawing.Size(75, 23)
        Me.btnView.TabIndex = 3
        Me.btnView.Text = "View"
        Me.btnView.UseVisualStyleBackColor = True
        '
        'pnlDomains
        '
        Me.pnlDomains.BackColor = System.Drawing.SystemColors.Control
        Me.pnlDomains.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pnlDomains.Controls.Add(Me.SplitContainer1)
        Me.pnlDomains.Controls.Add(Me.StatusStrip2)
        Me.pnlDomains.Controls.Add(Me.StatusStrip1)
        Me.pnlDomains.Location = New System.Drawing.Point(12, 12)
        Me.pnlDomains.Name = "pnlDomains"
        Me.pnlDomains.Size = New System.Drawing.Size(365, 477)
        Me.pnlDomains.TabIndex = 1
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
        Me.SplitContainer1.Panel1.Controls.Add(Me.btnSave)
        Me.SplitContainer1.Panel1.Controls.Add(Me.txtDomName)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label1)
        Me.SplitContainer1.Panel1.Controls.Add(Me.StatusStrip5)
        Me.SplitContainer1.Panel1.Controls.Add(Me.StatusStrip4)
        Me.SplitContainer1.Panel1.Controls.Add(Me.StatusStrip3)
        Me.SplitContainer1.Panel1.Controls.Add(Me.StatusStrip10)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.SplitContainer2)
        Me.SplitContainer1.Size = New System.Drawing.Size(363, 425)
        Me.SplitContainer1.SplitterDistance = 85
        Me.SplitContainer1.TabIndex = 11
        '
        'btnSave
        '
        Me.btnSave.Location = New System.Drawing.Point(268, 41)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(75, 21)
        Me.btnSave.TabIndex = 2
        Me.btnSave.Text = "Save"
        Me.btnSave.UseVisualStyleBackColor = True
        '
        'txtDomName
        '
        Me.txtDomName.Location = New System.Drawing.Point(16, 41)
        Me.txtDomName.Name = "txtDomName"
        Me.txtDomName.Size = New System.Drawing.Size(243, 20)
        Me.txtDomName.TabIndex = 15
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(13, 20)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(77, 13)
        Me.Label1.TabIndex = 14
        Me.Label1.Text = "Domain Name:"
        '
        'StatusStrip5
        '
        Me.StatusStrip5.AutoSize = False
        Me.StatusStrip5.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip5.Dock = System.Windows.Forms.DockStyle.Right
        Me.StatusStrip5.Location = New System.Drawing.Point(353, 10)
        Me.StatusStrip5.Name = "StatusStrip5"
        Me.StatusStrip5.Size = New System.Drawing.Size(10, 65)
        Me.StatusStrip5.SizingGrip = False
        Me.StatusStrip5.TabIndex = 13
        Me.StatusStrip5.Text = "StatusStrip5"
        '
        'StatusStrip4
        '
        Me.StatusStrip4.AutoSize = False
        Me.StatusStrip4.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip4.Dock = System.Windows.Forms.DockStyle.Left
        Me.StatusStrip4.Location = New System.Drawing.Point(0, 10)
        Me.StatusStrip4.Name = "StatusStrip4"
        Me.StatusStrip4.Size = New System.Drawing.Size(10, 65)
        Me.StatusStrip4.SizingGrip = False
        Me.StatusStrip4.TabIndex = 12
        Me.StatusStrip4.Text = "StatusStrip4"
        '
        'StatusStrip3
        '
        Me.StatusStrip3.AutoSize = False
        Me.StatusStrip3.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip3.Location = New System.Drawing.Point(0, 75)
        Me.StatusStrip3.Name = "StatusStrip3"
        Me.StatusStrip3.Size = New System.Drawing.Size(363, 10)
        Me.StatusStrip3.SizingGrip = False
        Me.StatusStrip3.TabIndex = 11
        Me.StatusStrip3.Text = "StatusStrip3"
        '
        'StatusStrip10
        '
        Me.StatusStrip10.AutoSize = False
        Me.StatusStrip10.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip10.Dock = System.Windows.Forms.DockStyle.Top
        Me.StatusStrip10.Location = New System.Drawing.Point(0, 0)
        Me.StatusStrip10.Name = "StatusStrip10"
        Me.StatusStrip10.Size = New System.Drawing.Size(363, 10)
        Me.StatusStrip10.SizingGrip = False
        Me.StatusStrip10.TabIndex = 10
        Me.StatusStrip10.Text = "StatusStrip10"
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
        Me.SplitContainer2.Panel1.Controls.Add(Me.lstDomains)
        '
        'SplitContainer2.Panel2
        '
        Me.SplitContainer2.Panel2.Controls.Add(Me.Panel3)
        Me.SplitContainer2.Panel2.Controls.Add(Me.StatusStrip6)
        Me.SplitContainer2.Panel2.Controls.Add(Me.StatusStrip7)
        Me.SplitContainer2.Panel2.Controls.Add(Me.StatusStrip8)
        Me.SplitContainer2.Panel2.Controls.Add(Me.StatusStrip9)
        Me.SplitContainer2.Size = New System.Drawing.Size(363, 336)
        Me.SplitContainer2.SplitterDistance = 252
        Me.SplitContainer2.TabIndex = 12
        '
        'lstDomains
        '
        Me.lstDomains.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColumnHeader1})
        Me.lstDomains.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lstDomains.FullRowSelect = True
        Me.lstDomains.GridLines = True
        Me.lstDomains.Location = New System.Drawing.Point(0, 0)
        Me.lstDomains.Name = "lstDomains"
        Me.lstDomains.Size = New System.Drawing.Size(363, 252)
        Me.lstDomains.TabIndex = 0
        Me.lstDomains.UseCompatibleStateImageBehavior = False
        Me.lstDomains.View = System.Windows.Forms.View.Details
        '
        'ColumnHeader1
        '
        Me.ColumnHeader1.Text = "Domain Name"
        Me.ColumnHeader1.Width = 355
        '
        'Panel3
        '
        Me.Panel3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel3.Controls.Add(Me.btnView)
        Me.Panel3.Controls.Add(Me.btnDelete)
        Me.Panel3.Controls.Add(Me.btnUpdate)
        Me.Panel3.Controls.Add(Me.btnExit)
        Me.Panel3.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel3.Location = New System.Drawing.Point(10, 20)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(343, 50)
        Me.Panel3.TabIndex = 16
        '
        'btnDelete
        '
        Me.btnDelete.Location = New System.Drawing.Point(173, 13)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(75, 23)
        Me.btnDelete.TabIndex = 2
        Me.btnDelete.Text = "Delete"
        Me.btnDelete.UseVisualStyleBackColor = True
        '
        'btnUpdate
        '
        Me.btnUpdate.Location = New System.Drawing.Point(89, 13)
        Me.btnUpdate.Name = "btnUpdate"
        Me.btnUpdate.Size = New System.Drawing.Size(75, 23)
        Me.btnUpdate.TabIndex = 1
        Me.btnUpdate.Text = "Update"
        Me.btnUpdate.UseVisualStyleBackColor = True
        '
        'btnExit
        '
        Me.btnExit.Location = New System.Drawing.Point(257, 13)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(75, 23)
        Me.btnExit.TabIndex = 0
        Me.btnExit.Text = "Exit"
        Me.btnExit.UseVisualStyleBackColor = True
        '
        'StatusStrip6
        '
        Me.StatusStrip6.AutoSize = False
        Me.StatusStrip6.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip6.Dock = System.Windows.Forms.DockStyle.Right
        Me.StatusStrip6.Location = New System.Drawing.Point(353, 20)
        Me.StatusStrip6.Name = "StatusStrip6"
        Me.StatusStrip6.Size = New System.Drawing.Size(10, 50)
        Me.StatusStrip6.SizingGrip = False
        Me.StatusStrip6.TabIndex = 15
        Me.StatusStrip6.Text = "StatusStrip6"
        '
        'StatusStrip7
        '
        Me.StatusStrip7.AutoSize = False
        Me.StatusStrip7.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip7.Dock = System.Windows.Forms.DockStyle.Left
        Me.StatusStrip7.Location = New System.Drawing.Point(0, 20)
        Me.StatusStrip7.Name = "StatusStrip7"
        Me.StatusStrip7.Size = New System.Drawing.Size(10, 50)
        Me.StatusStrip7.SizingGrip = False
        Me.StatusStrip7.TabIndex = 14
        Me.StatusStrip7.Text = "StatusStrip7"
        '
        'StatusStrip8
        '
        Me.StatusStrip8.AutoSize = False
        Me.StatusStrip8.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip8.Location = New System.Drawing.Point(0, 70)
        Me.StatusStrip8.Name = "StatusStrip8"
        Me.StatusStrip8.Size = New System.Drawing.Size(363, 10)
        Me.StatusStrip8.SizingGrip = False
        Me.StatusStrip8.TabIndex = 13
        Me.StatusStrip8.Text = "StatusStrip8"
        '
        'StatusStrip9
        '
        Me.StatusStrip9.AutoSize = False
        Me.StatusStrip9.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip9.Dock = System.Windows.Forms.DockStyle.Top
        Me.StatusStrip9.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripStatusLabel2})
        Me.StatusStrip9.Location = New System.Drawing.Point(0, 0)
        Me.StatusStrip9.Name = "StatusStrip9"
        Me.StatusStrip9.Size = New System.Drawing.Size(363, 20)
        Me.StatusStrip9.SizingGrip = False
        Me.StatusStrip9.TabIndex = 12
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
        Me.StatusStrip2.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripStatusLabel1})
        Me.StatusStrip2.Location = New System.Drawing.Point(0, 0)
        Me.StatusStrip2.Name = "StatusStrip2"
        Me.StatusStrip2.Size = New System.Drawing.Size(363, 25)
        Me.StatusStrip2.SizingGrip = False
        Me.StatusStrip2.TabIndex = 10
        Me.StatusStrip2.Text = "StatusStrip2"
        '
        'ToolStripStatusLabel1
        '
        Me.ToolStripStatusLabel1.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ToolStripStatusLabel1.Name = "ToolStripStatusLabel1"
        Me.ToolStripStatusLabel1.Size = New System.Drawing.Size(118, 20)
        Me.ToolStripStatusLabel1.Text = "DOMAIN MASTER"
        '
        'StatusStrip1
        '
        Me.StatusStrip1.AutoSize = False
        Me.StatusStrip1.BackColor = System.Drawing.Color.LightSeaGreen
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 450)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(363, 25)
        Me.StatusStrip1.SizingGrip = False
        Me.StatusStrip1.TabIndex = 9
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'frmDomains
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.LightSkyBlue
        Me.ClientSize = New System.Drawing.Size(388, 501)
        Me.Controls.Add(Me.pnlDomains)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Name = "frmDomains"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Domains"
        Me.pnlDomains.ResumeLayout(False)
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel1.PerformLayout()
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.ResumeLayout(False)
        Me.SplitContainer2.Panel1.ResumeLayout(False)
        Me.SplitContainer2.Panel2.ResumeLayout(False)
        CType(Me.SplitContainer2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer2.ResumeLayout(False)
        Me.Panel3.ResumeLayout(False)
        Me.StatusStrip9.ResumeLayout(False)
        Me.StatusStrip9.PerformLayout()
        Me.StatusStrip2.ResumeLayout(False)
        Me.StatusStrip2.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents btnView As System.Windows.Forms.Button
    Friend WithEvents pnlDomains As System.Windows.Forms.Panel
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents txtDomName As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents StatusStrip5 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip4 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip3 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip10 As System.Windows.Forms.StatusStrip
    Friend WithEvents SplitContainer2 As System.Windows.Forms.SplitContainer
    Friend WithEvents lstDomains As System.Windows.Forms.ListView
    Friend WithEvents ColumnHeader1 As System.Windows.Forms.ColumnHeader
    Friend WithEvents Panel3 As System.Windows.Forms.Panel
    Friend WithEvents btnDelete As System.Windows.Forms.Button
    Friend WithEvents btnUpdate As System.Windows.Forms.Button
    Friend WithEvents btnExit As System.Windows.Forms.Button
    Friend WithEvents StatusStrip6 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip7 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip8 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip9 As System.Windows.Forms.StatusStrip
    Friend WithEvents ToolStripStatusLabel2 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents StatusStrip2 As System.Windows.Forms.StatusStrip
    Friend WithEvents ToolStripStatusLabel1 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
End Class
