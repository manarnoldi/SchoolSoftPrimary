<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmDomainRights
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
        Me.pnlDomRights = New System.Windows.Forms.Panel()
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.lstModuleAlreday = New System.Windows.Forms.ListView()
        Me.ColumnHeader2 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.btnLeftAll = New System.Windows.Forms.Button()
        Me.btnRightAll = New System.Windows.Forms.Button()
        Me.btnLeftOne = New System.Windows.Forms.Button()
        Me.btnRightOne = New System.Windows.Forms.Button()
        Me.lstModules = New System.Windows.Forms.ListView()
        Me.ColumnHeader1 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.Label1 = New System.Windows.Forms.Label()
        Me.cboDomains = New System.Windows.Forms.ComboBox()
        Me.StatusStrip7 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip4 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip9 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip6 = New System.Windows.Forms.StatusStrip()
        Me.ToolStripStatusLabel2 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.btnReload = New System.Windows.Forms.Button()
        Me.btnClose = New System.Windows.Forms.Button()
        Me.btnUpdate = New System.Windows.Forms.Button()
        Me.StatusStrip5 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip8 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip10 = New System.Windows.Forms.StatusStrip()
        Me.ToolStripStatusLabel3 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.StatusStrip3 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip2 = New System.Windows.Forms.StatusStrip()
        Me.ToolStripStatusLabel1 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip()
        Me.pnlDomRights.SuspendLayout()
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.StatusStrip6.SuspendLayout()
        Me.StatusStrip10.SuspendLayout()
        Me.StatusStrip2.SuspendLayout()
        Me.SuspendLayout()
        '
        'pnlDomRights
        '
        Me.pnlDomRights.BackColor = System.Drawing.SystemColors.Control
        Me.pnlDomRights.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pnlDomRights.Controls.Add(Me.SplitContainer1)
        Me.pnlDomRights.Controls.Add(Me.StatusStrip2)
        Me.pnlDomRights.Controls.Add(Me.StatusStrip1)
        Me.pnlDomRights.Location = New System.Drawing.Point(12, 12)
        Me.pnlDomRights.Name = "pnlDomRights"
        Me.pnlDomRights.Size = New System.Drawing.Size(679, 525)
        Me.pnlDomRights.TabIndex = 0
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
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label2)
        Me.SplitContainer1.Panel1.Controls.Add(Me.lstModuleAlreday)
        Me.SplitContainer1.Panel1.Controls.Add(Me.btnLeftAll)
        Me.SplitContainer1.Panel1.Controls.Add(Me.btnRightAll)
        Me.SplitContainer1.Panel1.Controls.Add(Me.btnLeftOne)
        Me.SplitContainer1.Panel1.Controls.Add(Me.btnRightOne)
        Me.SplitContainer1.Panel1.Controls.Add(Me.lstModules)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label1)
        Me.SplitContainer1.Panel1.Controls.Add(Me.cboDomains)
        Me.SplitContainer1.Panel1.Controls.Add(Me.StatusStrip7)
        Me.SplitContainer1.Panel1.Controls.Add(Me.StatusStrip4)
        Me.SplitContainer1.Panel1.Controls.Add(Me.StatusStrip9)
        Me.SplitContainer1.Panel1.Controls.Add(Me.StatusStrip6)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.btnReload)
        Me.SplitContainer1.Panel2.Controls.Add(Me.btnClose)
        Me.SplitContainer1.Panel2.Controls.Add(Me.btnUpdate)
        Me.SplitContainer1.Panel2.Controls.Add(Me.StatusStrip5)
        Me.SplitContainer1.Panel2.Controls.Add(Me.StatusStrip8)
        Me.SplitContainer1.Panel2.Controls.Add(Me.StatusStrip10)
        Me.SplitContainer1.Panel2.Controls.Add(Me.StatusStrip3)
        Me.SplitContainer1.Size = New System.Drawing.Size(677, 473)
        Me.SplitContainer1.SplitterDistance = 393
        Me.SplitContainer1.SplitterWidth = 1
        Me.SplitContainer1.TabIndex = 13
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(361, 31)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(289, 16)
        Me.Label2.TabIndex = 31
        Me.Label2.Text = "Domain Already Accessing Following Modules:"
        '
        'lstModuleAlreday
        '
        Me.lstModuleAlreday.CheckBoxes = True
        Me.lstModuleAlreday.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColumnHeader2})
        Me.lstModuleAlreday.FullRowSelect = True
        Me.lstModuleAlreday.GridLines = True
        Me.lstModuleAlreday.Location = New System.Drawing.Point(364, 59)
        Me.lstModuleAlreday.Name = "lstModuleAlreday"
        Me.lstModuleAlreday.Size = New System.Drawing.Size(282, 315)
        Me.lstModuleAlreday.TabIndex = 30
        Me.lstModuleAlreday.UseCompatibleStateImageBehavior = False
        Me.lstModuleAlreday.View = System.Windows.Forms.View.Details
        '
        'ColumnHeader2
        '
        Me.ColumnHeader2.Text = "ModuleName"
        Me.ColumnHeader2.Width = 278
        '
        'btnLeftAll
        '
        Me.btnLeftAll.Location = New System.Drawing.Point(319, 209)
        Me.btnLeftAll.Name = "btnLeftAll"
        Me.btnLeftAll.Size = New System.Drawing.Size(39, 23)
        Me.btnLeftAll.TabIndex = 29
        Me.btnLeftAll.Text = "<<"
        Me.btnLeftAll.UseVisualStyleBackColor = True
        '
        'btnRightAll
        '
        Me.btnRightAll.Location = New System.Drawing.Point(319, 123)
        Me.btnRightAll.Name = "btnRightAll"
        Me.btnRightAll.Size = New System.Drawing.Size(39, 23)
        Me.btnRightAll.TabIndex = 28
        Me.btnRightAll.Text = ">>"
        Me.btnRightAll.UseVisualStyleBackColor = True
        '
        'btnLeftOne
        '
        Me.btnLeftOne.Location = New System.Drawing.Point(319, 180)
        Me.btnLeftOne.Name = "btnLeftOne"
        Me.btnLeftOne.Size = New System.Drawing.Size(39, 23)
        Me.btnLeftOne.TabIndex = 27
        Me.btnLeftOne.Text = "<"
        Me.btnLeftOne.UseVisualStyleBackColor = True
        '
        'btnRightOne
        '
        Me.btnRightOne.Location = New System.Drawing.Point(319, 152)
        Me.btnRightOne.Name = "btnRightOne"
        Me.btnRightOne.Size = New System.Drawing.Size(39, 23)
        Me.btnRightOne.TabIndex = 26
        Me.btnRightOne.Text = ">"
        Me.btnRightOne.UseVisualStyleBackColor = True
        '
        'lstModules
        '
        Me.lstModules.CheckBoxes = True
        Me.lstModules.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColumnHeader1})
        Me.lstModules.FullRowSelect = True
        Me.lstModules.GridLines = True
        Me.lstModules.Location = New System.Drawing.Point(31, 59)
        Me.lstModules.Name = "lstModules"
        Me.lstModules.Size = New System.Drawing.Size(282, 315)
        Me.lstModules.TabIndex = 25
        Me.lstModules.UseCompatibleStateImageBehavior = False
        Me.lstModules.View = System.Windows.Forms.View.Details
        '
        'ColumnHeader1
        '
        Me.ColumnHeader1.Text = "Module Name"
        Me.ColumnHeader1.Width = 278
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(28, 31)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(79, 13)
        Me.Label1.TabIndex = 23
        Me.Label1.Text = "Select Domain:"
        '
        'cboDomains
        '
        Me.cboDomains.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboDomains.FormattingEnabled = True
        Me.cboDomains.Location = New System.Drawing.Point(113, 28)
        Me.cboDomains.Name = "cboDomains"
        Me.cboDomains.Size = New System.Drawing.Size(200, 21)
        Me.cboDomains.TabIndex = 22
        '
        'StatusStrip7
        '
        Me.StatusStrip7.AutoSize = False
        Me.StatusStrip7.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip7.Dock = System.Windows.Forms.DockStyle.Right
        Me.StatusStrip7.Location = New System.Drawing.Point(665, 20)
        Me.StatusStrip7.Name = "StatusStrip7"
        Me.StatusStrip7.Size = New System.Drawing.Size(10, 361)
        Me.StatusStrip7.SizingGrip = False
        Me.StatusStrip7.TabIndex = 21
        Me.StatusStrip7.Text = "StatusStrip7"
        '
        'StatusStrip4
        '
        Me.StatusStrip4.AutoSize = False
        Me.StatusStrip4.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip4.Dock = System.Windows.Forms.DockStyle.Left
        Me.StatusStrip4.Location = New System.Drawing.Point(0, 20)
        Me.StatusStrip4.Name = "StatusStrip4"
        Me.StatusStrip4.Size = New System.Drawing.Size(10, 361)
        Me.StatusStrip4.SizingGrip = False
        Me.StatusStrip4.TabIndex = 19
        Me.StatusStrip4.Text = "StatusStrip4"
        '
        'StatusStrip9
        '
        Me.StatusStrip9.AutoSize = False
        Me.StatusStrip9.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip9.Location = New System.Drawing.Point(0, 381)
        Me.StatusStrip9.Name = "StatusStrip9"
        Me.StatusStrip9.Size = New System.Drawing.Size(675, 10)
        Me.StatusStrip9.SizingGrip = False
        Me.StatusStrip9.TabIndex = 16
        Me.StatusStrip9.Text = "StatusStrip9"
        '
        'StatusStrip6
        '
        Me.StatusStrip6.AutoSize = False
        Me.StatusStrip6.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip6.Dock = System.Windows.Forms.DockStyle.Top
        Me.StatusStrip6.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripStatusLabel2})
        Me.StatusStrip6.Location = New System.Drawing.Point(0, 0)
        Me.StatusStrip6.Name = "StatusStrip6"
        Me.StatusStrip6.Size = New System.Drawing.Size(675, 20)
        Me.StatusStrip6.SizingGrip = False
        Me.StatusStrip6.TabIndex = 15
        Me.StatusStrip6.Text = "StatusStrip6"
        '
        'ToolStripStatusLabel2
        '
        Me.ToolStripStatusLabel2.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ToolStripStatusLabel2.Name = "ToolStripStatusLabel2"
        Me.ToolStripStatusLabel2.Size = New System.Drawing.Size(78, 15)
        Me.ToolStripStatusLabel2.Text = "Enter Details"
        '
        'btnReload
        '
        Me.btnReload.Location = New System.Drawing.Point(32, 32)
        Me.btnReload.Name = "btnReload"
        Me.btnReload.Size = New System.Drawing.Size(75, 23)
        Me.btnReload.TabIndex = 31
        Me.btnReload.Text = "Reload"
        Me.btnReload.UseVisualStyleBackColor = True
        '
        'btnClose
        '
        Me.btnClose.Location = New System.Drawing.Point(575, 32)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(75, 23)
        Me.btnClose.TabIndex = 30
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'btnUpdate
        '
        Me.btnUpdate.Location = New System.Drawing.Point(364, 32)
        Me.btnUpdate.Name = "btnUpdate"
        Me.btnUpdate.Size = New System.Drawing.Size(75, 23)
        Me.btnUpdate.TabIndex = 29
        Me.btnUpdate.Text = "Update"
        Me.btnUpdate.UseVisualStyleBackColor = True
        '
        'StatusStrip5
        '
        Me.StatusStrip5.AutoSize = False
        Me.StatusStrip5.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip5.Dock = System.Windows.Forms.DockStyle.Right
        Me.StatusStrip5.Location = New System.Drawing.Point(665, 20)
        Me.StatusStrip5.Name = "StatusStrip5"
        Me.StatusStrip5.Size = New System.Drawing.Size(10, 47)
        Me.StatusStrip5.SizingGrip = False
        Me.StatusStrip5.TabIndex = 24
        Me.StatusStrip5.Text = "StatusStrip5"
        '
        'StatusStrip8
        '
        Me.StatusStrip8.AutoSize = False
        Me.StatusStrip8.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip8.Dock = System.Windows.Forms.DockStyle.Left
        Me.StatusStrip8.Location = New System.Drawing.Point(0, 20)
        Me.StatusStrip8.Name = "StatusStrip8"
        Me.StatusStrip8.Size = New System.Drawing.Size(10, 47)
        Me.StatusStrip8.SizingGrip = False
        Me.StatusStrip8.TabIndex = 23
        Me.StatusStrip8.Text = "StatusStrip8"
        '
        'StatusStrip10
        '
        Me.StatusStrip10.AutoSize = False
        Me.StatusStrip10.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip10.Dock = System.Windows.Forms.DockStyle.Top
        Me.StatusStrip10.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripStatusLabel3})
        Me.StatusStrip10.Location = New System.Drawing.Point(0, 0)
        Me.StatusStrip10.Name = "StatusStrip10"
        Me.StatusStrip10.Size = New System.Drawing.Size(675, 20)
        Me.StatusStrip10.SizingGrip = False
        Me.StatusStrip10.TabIndex = 21
        Me.StatusStrip10.Text = "StatusStrip10"
        '
        'ToolStripStatusLabel3
        '
        Me.ToolStripStatusLabel3.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ToolStripStatusLabel3.Name = "ToolStripStatusLabel3"
        Me.ToolStripStatusLabel3.Size = New System.Drawing.Size(80, 15)
        Me.ToolStripStatusLabel3.Text = "Carry Actions"
        '
        'StatusStrip3
        '
        Me.StatusStrip3.AutoSize = False
        Me.StatusStrip3.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip3.Location = New System.Drawing.Point(0, 67)
        Me.StatusStrip3.Name = "StatusStrip3"
        Me.StatusStrip3.Size = New System.Drawing.Size(675, 10)
        Me.StatusStrip3.SizingGrip = False
        Me.StatusStrip3.TabIndex = 22
        Me.StatusStrip3.Text = "StatusStrip3"
        '
        'StatusStrip2
        '
        Me.StatusStrip2.AutoSize = False
        Me.StatusStrip2.BackColor = System.Drawing.Color.LightSeaGreen
        Me.StatusStrip2.Dock = System.Windows.Forms.DockStyle.Top
        Me.StatusStrip2.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripStatusLabel1})
        Me.StatusStrip2.Location = New System.Drawing.Point(0, 0)
        Me.StatusStrip2.Name = "StatusStrip2"
        Me.StatusStrip2.Size = New System.Drawing.Size(677, 25)
        Me.StatusStrip2.SizingGrip = False
        Me.StatusStrip2.TabIndex = 12
        Me.StatusStrip2.Text = "StatusStrip2"
        '
        'ToolStripStatusLabel1
        '
        Me.ToolStripStatusLabel1.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ToolStripStatusLabel1.Name = "ToolStripStatusLabel1"
        Me.ToolStripStatusLabel1.Size = New System.Drawing.Size(163, 20)
        Me.ToolStripStatusLabel1.Text = "ASSIGN DOMAIN RIGHTS"
        '
        'StatusStrip1
        '
        Me.StatusStrip1.AutoSize = False
        Me.StatusStrip1.BackColor = System.Drawing.Color.LightSeaGreen
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 498)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(677, 25)
        Me.StatusStrip1.SizingGrip = False
        Me.StatusStrip1.TabIndex = 11
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'frmDomainRights
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.LightSkyBlue
        Me.ClientSize = New System.Drawing.Size(701, 555)
        Me.Controls.Add(Me.pnlDomRights)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.MaximizeBox = False
        Me.Name = "frmDomainRights"
        Me.ShowIcon = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Domain Rights"
        Me.pnlDomRights.ResumeLayout(False)
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel1.PerformLayout()
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.ResumeLayout(False)
        Me.StatusStrip6.ResumeLayout(False)
        Me.StatusStrip6.PerformLayout()
        Me.StatusStrip10.ResumeLayout(False)
        Me.StatusStrip10.PerformLayout()
        Me.StatusStrip2.ResumeLayout(False)
        Me.StatusStrip2.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents pnlDomRights As System.Windows.Forms.Panel
    Friend WithEvents StatusStrip2 As System.Windows.Forms.StatusStrip
    Friend WithEvents ToolStripStatusLabel1 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents StatusStrip7 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip4 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip9 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip6 As System.Windows.Forms.StatusStrip
    Friend WithEvents ToolStripStatusLabel2 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents StatusStrip5 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip8 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip10 As System.Windows.Forms.StatusStrip
    Friend WithEvents ToolStripStatusLabel3 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents StatusStrip3 As System.Windows.Forms.StatusStrip
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents cboDomains As System.Windows.Forms.ComboBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents lstModuleAlreday As System.Windows.Forms.ListView
    Friend WithEvents ColumnHeader2 As System.Windows.Forms.ColumnHeader
    Friend WithEvents btnLeftAll As System.Windows.Forms.Button
    Friend WithEvents btnRightAll As System.Windows.Forms.Button
    Friend WithEvents btnLeftOne As System.Windows.Forms.Button
    Friend WithEvents btnRightOne As System.Windows.Forms.Button
    Friend WithEvents lstModules As System.Windows.Forms.ListView
    Friend WithEvents ColumnHeader1 As System.Windows.Forms.ColumnHeader
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents btnUpdate As System.Windows.Forms.Button
    Friend WithEvents btnReload As System.Windows.Forms.Button
End Class
