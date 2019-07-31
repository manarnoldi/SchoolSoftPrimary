<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmTTGeneration
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
        Me.components = New System.ComponentModel.Container()
        Me.pnlTTGeneration = New System.Windows.Forms.Panel()
        Me.ContextMenuStripTT = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.SETMAXIMUMLESSONSPERDAYToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.SETMAXIMUMDOUBLESPERDAYToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.btnView = New System.Windows.Forms.Button()
        Me.cboYear = New System.Windows.Forms.ComboBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.btnGentt = New System.Windows.Forms.Button()
        Me.btnClose = New System.Windows.Forms.Button()
        Me.btnSave = New System.Windows.Forms.Button()
        Me.btngenStr = New System.Windows.Forms.Button()
        Me.StatusStrip6 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip7 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip8 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip9 = New System.Windows.Forms.StatusStrip()
        Me.dgTTGeneration = New System.Windows.Forms.DataGridView()
        Me.StatusStrip2 = New System.Windows.Forms.StatusStrip()
        Me.ToolStripStatusLabel1 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip()
        Me.ToolStripStatusLabel2 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.genTimer = New System.Windows.Forms.Timer(Me.components)
        Me.bgWorkerGen = New System.ComponentModel.BackgroundWorker()
        Me.pnlTTGeneration.SuspendLayout()
        Me.ContextMenuStripTT.SuspendLayout()
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.Panel1.SuspendLayout()
        CType(Me.dgTTGeneration, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.StatusStrip2.SuspendLayout()
        Me.StatusStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'pnlTTGeneration
        '
        Me.pnlTTGeneration.AutoSize = True
        Me.pnlTTGeneration.BackColor = System.Drawing.SystemColors.Control
        Me.pnlTTGeneration.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pnlTTGeneration.ContextMenuStrip = Me.ContextMenuStripTT
        Me.pnlTTGeneration.Controls.Add(Me.SplitContainer1)
        Me.pnlTTGeneration.Controls.Add(Me.StatusStrip2)
        Me.pnlTTGeneration.Controls.Add(Me.StatusStrip1)
        Me.pnlTTGeneration.Location = New System.Drawing.Point(12, 12)
        Me.pnlTTGeneration.Name = "pnlTTGeneration"
        Me.pnlTTGeneration.Size = New System.Drawing.Size(981, 517)
        Me.pnlTTGeneration.TabIndex = 0
        '
        'ContextMenuStripTT
        '
        Me.ContextMenuStripTT.BackColor = System.Drawing.Color.LightSkyBlue
        Me.ContextMenuStripTT.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ContextMenuStripTT.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.SETMAXIMUMLESSONSPERDAYToolStripMenuItem, Me.SETMAXIMUMDOUBLESPERDAYToolStripMenuItem})
        Me.ContextMenuStripTT.Name = "ContextMenuStripTT"
        Me.ContextMenuStripTT.Size = New System.Drawing.Size(289, 48)
        '
        'SETMAXIMUMLESSONSPERDAYToolStripMenuItem
        '
        Me.SETMAXIMUMLESSONSPERDAYToolStripMenuItem.Name = "SETMAXIMUMLESSONSPERDAYToolStripMenuItem"
        Me.SETMAXIMUMLESSONSPERDAYToolStripMenuItem.Size = New System.Drawing.Size(288, 22)
        Me.SETMAXIMUMLESSONSPERDAYToolStripMenuItem.Text = "SET MAXIMUM LESSONS PER DAY"
        '
        'SETMAXIMUMDOUBLESPERDAYToolStripMenuItem
        '
        Me.SETMAXIMUMDOUBLESPERDAYToolStripMenuItem.Name = "SETMAXIMUMDOUBLESPERDAYToolStripMenuItem"
        Me.SETMAXIMUMDOUBLESPERDAYToolStripMenuItem.Size = New System.Drawing.Size(288, 22)
        Me.SETMAXIMUMDOUBLESPERDAYToolStripMenuItem.Text = "SET MAXIMUM DOUBLES PER DAY"
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
        Me.SplitContainer1.Panel1.Controls.Add(Me.StatusStrip6)
        Me.SplitContainer1.Panel1.Controls.Add(Me.StatusStrip7)
        Me.SplitContainer1.Panel1.Controls.Add(Me.StatusStrip8)
        Me.SplitContainer1.Panel1.Controls.Add(Me.StatusStrip9)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.dgTTGeneration)
        Me.SplitContainer1.Size = New System.Drawing.Size(979, 465)
        Me.SplitContainer1.SplitterDistance = 46
        Me.SplitContainer1.SplitterWidth = 1
        Me.SplitContainer1.TabIndex = 17
        '
        'Panel1
        '
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel1.Controls.Add(Me.btnView)
        Me.Panel1.Controls.Add(Me.cboYear)
        Me.Panel1.Controls.Add(Me.Label7)
        Me.Panel1.Controls.Add(Me.btnGentt)
        Me.Panel1.Controls.Add(Me.btnClose)
        Me.Panel1.Controls.Add(Me.btnSave)
        Me.Panel1.Controls.Add(Me.btngenStr)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel1.Location = New System.Drawing.Point(5, 5)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(967, 34)
        Me.Panel1.TabIndex = 51
        '
        'btnView
        '
        Me.btnView.Location = New System.Drawing.Point(760, 4)
        Me.btnView.Name = "btnView"
        Me.btnView.Size = New System.Drawing.Size(75, 23)
        Me.btnView.TabIndex = 14
        Me.btnView.Text = "View"
        Me.btnView.UseVisualStyleBackColor = True
        '
        'cboYear
        '
        Me.cboYear.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboYear.FormattingEnabled = True
        Me.cboYear.Location = New System.Drawing.Point(48, 6)
        Me.cboYear.Name = "cboYear"
        Me.cboYear.Size = New System.Drawing.Size(86, 21)
        Me.cboYear.TabIndex = 13
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(10, 9)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(32, 13)
        Me.Label7.TabIndex = 12
        Me.Label7.Text = "Year:"
        '
        'btnGentt
        '
        Me.btnGentt.Location = New System.Drawing.Point(326, 4)
        Me.btnGentt.Name = "btnGentt"
        Me.btnGentt.Size = New System.Drawing.Size(146, 23)
        Me.btnGentt.TabIndex = 4
        Me.btnGentt.Text = "Generate TimeTable"
        Me.btnGentt.UseVisualStyleBackColor = True
        '
        'btnClose
        '
        Me.btnClose.Location = New System.Drawing.Point(885, 4)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(74, 23)
        Me.btnClose.TabIndex = 2
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'btnSave
        '
        Me.btnSave.Location = New System.Drawing.Point(636, 4)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(75, 23)
        Me.btnSave.TabIndex = 1
        Me.btnSave.Text = "Save"
        Me.btnSave.UseVisualStyleBackColor = True
        '
        'btngenStr
        '
        Me.btngenStr.Location = New System.Drawing.Point(160, 4)
        Me.btngenStr.Name = "btngenStr"
        Me.btngenStr.Size = New System.Drawing.Size(146, 23)
        Me.btngenStr.TabIndex = 0
        Me.btngenStr.Text = "Generate Structure"
        Me.btngenStr.UseVisualStyleBackColor = True
        '
        'StatusStrip6
        '
        Me.StatusStrip6.AutoSize = False
        Me.StatusStrip6.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip6.Dock = System.Windows.Forms.DockStyle.Right
        Me.StatusStrip6.Location = New System.Drawing.Point(972, 5)
        Me.StatusStrip6.Name = "StatusStrip6"
        Me.StatusStrip6.Size = New System.Drawing.Size(5, 34)
        Me.StatusStrip6.SizingGrip = False
        Me.StatusStrip6.TabIndex = 50
        Me.StatusStrip6.Text = "StatusStrip6"
        '
        'StatusStrip7
        '
        Me.StatusStrip7.AutoSize = False
        Me.StatusStrip7.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip7.Dock = System.Windows.Forms.DockStyle.Left
        Me.StatusStrip7.Location = New System.Drawing.Point(0, 5)
        Me.StatusStrip7.Name = "StatusStrip7"
        Me.StatusStrip7.Size = New System.Drawing.Size(5, 34)
        Me.StatusStrip7.SizingGrip = False
        Me.StatusStrip7.TabIndex = 49
        Me.StatusStrip7.Text = "StatusStrip7"
        '
        'StatusStrip8
        '
        Me.StatusStrip8.AutoSize = False
        Me.StatusStrip8.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip8.Location = New System.Drawing.Point(0, 39)
        Me.StatusStrip8.Name = "StatusStrip8"
        Me.StatusStrip8.Size = New System.Drawing.Size(977, 5)
        Me.StatusStrip8.SizingGrip = False
        Me.StatusStrip8.TabIndex = 48
        Me.StatusStrip8.Text = "StatusStrip8"
        '
        'StatusStrip9
        '
        Me.StatusStrip9.AutoSize = False
        Me.StatusStrip9.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip9.Dock = System.Windows.Forms.DockStyle.Top
        Me.StatusStrip9.Location = New System.Drawing.Point(0, 0)
        Me.StatusStrip9.Name = "StatusStrip9"
        Me.StatusStrip9.Size = New System.Drawing.Size(977, 5)
        Me.StatusStrip9.SizingGrip = False
        Me.StatusStrip9.TabIndex = 47
        Me.StatusStrip9.Text = "StatusStrip9"
        '
        'dgTTGeneration
        '
        Me.dgTTGeneration.AllowUserToAddRows = False
        Me.dgTTGeneration.AllowUserToDeleteRows = False
        Me.dgTTGeneration.AllowUserToResizeColumns = False
        Me.dgTTGeneration.AllowUserToResizeRows = False
        Me.dgTTGeneration.BackgroundColor = System.Drawing.Color.White
        Me.dgTTGeneration.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgTTGeneration.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgTTGeneration.Location = New System.Drawing.Point(0, 0)
        Me.dgTTGeneration.Name = "dgTTGeneration"
        Me.dgTTGeneration.RowHeadersWidth = 4
        Me.dgTTGeneration.ShowCellErrors = False
        Me.dgTTGeneration.ShowCellToolTips = False
        Me.dgTTGeneration.ShowEditingIcon = False
        Me.dgTTGeneration.ShowRowErrors = False
        Me.dgTTGeneration.Size = New System.Drawing.Size(977, 416)
        Me.dgTTGeneration.TabIndex = 1
        '
        'StatusStrip2
        '
        Me.StatusStrip2.AutoSize = False
        Me.StatusStrip2.BackColor = System.Drawing.Color.LightSeaGreen
        Me.StatusStrip2.Dock = System.Windows.Forms.DockStyle.Top
        Me.StatusStrip2.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripStatusLabel1})
        Me.StatusStrip2.Location = New System.Drawing.Point(0, 0)
        Me.StatusStrip2.Name = "StatusStrip2"
        Me.StatusStrip2.Size = New System.Drawing.Size(979, 25)
        Me.StatusStrip2.SizingGrip = False
        Me.StatusStrip2.TabIndex = 16
        Me.StatusStrip2.Text = "StatusStrip2"
        '
        'ToolStripStatusLabel1
        '
        Me.ToolStripStatusLabel1.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ToolStripStatusLabel1.Name = "ToolStripStatusLabel1"
        Me.ToolStripStatusLabel1.Size = New System.Drawing.Size(162, 20)
        Me.ToolStripStatusLabel1.Text = "TIMETABLE GENERATION"
        '
        'StatusStrip1
        '
        Me.StatusStrip1.AutoSize = False
        Me.StatusStrip1.BackColor = System.Drawing.Color.LightSeaGreen
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripStatusLabel2})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 490)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(979, 25)
        Me.StatusStrip1.SizingGrip = False
        Me.StatusStrip1.TabIndex = 15
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'ToolStripStatusLabel2
        '
        Me.ToolStripStatusLabel2.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ToolStripStatusLabel2.Name = "ToolStripStatusLabel2"
        Me.ToolStripStatusLabel2.Size = New System.Drawing.Size(179, 20)
        Me.ToolStripStatusLabel2.Text = "Timetable Generation Status ..."
        '
        'genTimer
        '
        Me.genTimer.Enabled = True
        Me.genTimer.Interval = 1000
        '
        'bgWorkerGen
        '
        Me.bgWorkerGen.WorkerReportsProgress = True
        Me.bgWorkerGen.WorkerSupportsCancellation = True
        '
        'frmTTGeneration
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.LightSkyBlue
        Me.ClientSize = New System.Drawing.Size(1005, 541)
        Me.Controls.Add(Me.pnlTTGeneration)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Name = "frmTTGeneration"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Generation"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.pnlTTGeneration.ResumeLayout(False)
        Me.ContextMenuStripTT.ResumeLayout(False)
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.ResumeLayout(False)
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        CType(Me.dgTTGeneration, System.ComponentModel.ISupportInitialize).EndInit()
        Me.StatusStrip2.ResumeLayout(False)
        Me.StatusStrip2.PerformLayout()
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents pnlTTGeneration As System.Windows.Forms.Panel
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents StatusStrip2 As System.Windows.Forms.StatusStrip
    Friend WithEvents ToolStripStatusLabel1 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents StatusStrip6 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip7 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip8 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip9 As System.Windows.Forms.StatusStrip
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents btngenStr As System.Windows.Forms.Button
    Friend WithEvents dgTTGeneration As System.Windows.Forms.DataGridView
    Friend WithEvents btnGentt As System.Windows.Forms.Button
    Friend WithEvents cboYear As System.Windows.Forms.ComboBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents btnView As System.Windows.Forms.Button
    Friend WithEvents ContextMenuStripTT As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents SETMAXIMUMLESSONSPERDAYToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents SETMAXIMUMDOUBLESPERDAYToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents genTimer As System.Windows.Forms.Timer
    Friend WithEvents ToolStripStatusLabel2 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents bgWorkerGen As System.ComponentModel.BackgroundWorker
End Class
