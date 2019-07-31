<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmSchClasses
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
        Me.pnlSchClasses = New System.Windows.Forms.Panel()
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
        Me.cboStream = New System.Windows.Forms.ComboBox()
        Me.cboClassName = New System.Windows.Forms.ComboBox()
        Me.cboYear = New System.Windows.Forms.ComboBox()
        Me.btnSave = New System.Windows.Forms.Button()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.StatusStrip5 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip4 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip3 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip10 = New System.Windows.Forms.StatusStrip()
        Me.SplitContainer2 = New System.Windows.Forms.SplitContainer()
        Me.lstClasses = New System.Windows.Forms.ListView()
        Me.ColumnHeader1 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader2 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader3 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.StatusStrip7 = New System.Windows.Forms.StatusStrip()
        Me.btnView = New System.Windows.Forms.Button()
        Me.btnDelete = New System.Windows.Forms.Button()
        Me.btnUpdate = New System.Windows.Forms.Button()
        Me.btnClose = New System.Windows.Forms.Button()
        Me.StatusStrip6 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip8 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip9 = New System.Windows.Forms.StatusStrip()
        Me.ToolStripStatusLabel2 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.StatusStrip2 = New System.Windows.Forms.StatusStrip()
        Me.ToolStripStatusLabel1 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip()
        Me.pnlSchClasses.SuspendLayout()
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        CType(Me.SplitContainer2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer2.Panel1.SuspendLayout()
        Me.SplitContainer2.Panel2.SuspendLayout()
        Me.SplitContainer2.SuspendLayout()
        Me.StatusStrip9.SuspendLayout()
        Me.StatusStrip2.SuspendLayout()
        Me.SuspendLayout()
        '
        'pnlSchClasses
        '
        Me.pnlSchClasses.BackColor = System.Drawing.SystemColors.Control
        Me.pnlSchClasses.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pnlSchClasses.Controls.Add(Me.SplitContainer1)
        Me.pnlSchClasses.Controls.Add(Me.StatusStrip2)
        Me.pnlSchClasses.Controls.Add(Me.StatusStrip1)
        Me.pnlSchClasses.Location = New System.Drawing.Point(12, 12)
        Me.pnlSchClasses.Name = "pnlSchClasses"
        Me.pnlSchClasses.Size = New System.Drawing.Size(620, 450)
        Me.pnlSchClasses.TabIndex = 0
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
        Me.SplitContainer1.Panel1.Controls.Add(Me.cboStream)
        Me.SplitContainer1.Panel1.Controls.Add(Me.cboClassName)
        Me.SplitContainer1.Panel1.Controls.Add(Me.cboYear)
        Me.SplitContainer1.Panel1.Controls.Add(Me.btnSave)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label3)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label2)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label1)
        Me.SplitContainer1.Panel1.Controls.Add(Me.StatusStrip5)
        Me.SplitContainer1.Panel1.Controls.Add(Me.StatusStrip4)
        Me.SplitContainer1.Panel1.Controls.Add(Me.StatusStrip3)
        Me.SplitContainer1.Panel1.Controls.Add(Me.StatusStrip10)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.SplitContainer2)
        Me.SplitContainer1.Size = New System.Drawing.Size(618, 398)
        Me.SplitContainer1.SplitterDistance = 73
        Me.SplitContainer1.SplitterWidth = 1
        Me.SplitContainer1.TabIndex = 21
        '
        'cboStream
        '
        Me.cboStream.FormattingEnabled = True
        Me.cboStream.Location = New System.Drawing.Point(426, 26)
        Me.cboStream.Name = "cboStream"
        Me.cboStream.Size = New System.Drawing.Size(81, 21)
        Me.cboStream.TabIndex = 35
        '
        'cboClassName
        '
        Me.cboClassName.FormattingEnabled = True
        Me.cboClassName.Location = New System.Drawing.Point(239, 26)
        Me.cboClassName.Name = "cboClassName"
        Me.cboClassName.Size = New System.Drawing.Size(125, 21)
        Me.cboClassName.TabIndex = 34
        '
        'cboYear
        '
        Me.cboYear.FormattingEnabled = True
        Me.cboYear.Location = New System.Drawing.Point(56, 26)
        Me.cboYear.Name = "cboYear"
        Me.cboYear.Size = New System.Drawing.Size(101, 21)
        Me.cboYear.TabIndex = 33
        '
        'btnSave
        '
        Me.btnSave.Location = New System.Drawing.Point(523, 24)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(75, 23)
        Me.btnSave.TabIndex = 39
        Me.btnSave.Text = "Save"
        Me.btnSave.UseVisualStyleBackColor = True
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(376, 29)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(43, 13)
        Me.Label3.TabIndex = 32
        Me.Label3.Text = "Stream:"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(167, 29)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(66, 13)
        Me.Label2.TabIndex = 31
        Me.Label2.Text = "Class Name:"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(18, 29)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(32, 13)
        Me.Label1.TabIndex = 30
        Me.Label1.Text = "Year:"
        '
        'StatusStrip5
        '
        Me.StatusStrip5.AutoSize = False
        Me.StatusStrip5.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip5.Dock = System.Windows.Forms.DockStyle.Right
        Me.StatusStrip5.Location = New System.Drawing.Point(608, 10)
        Me.StatusStrip5.Name = "StatusStrip5"
        Me.StatusStrip5.Size = New System.Drawing.Size(10, 53)
        Me.StatusStrip5.SizingGrip = False
        Me.StatusStrip5.TabIndex = 29
        Me.StatusStrip5.Text = "StatusStrip5"
        '
        'StatusStrip4
        '
        Me.StatusStrip4.AutoSize = False
        Me.StatusStrip4.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip4.Dock = System.Windows.Forms.DockStyle.Left
        Me.StatusStrip4.Location = New System.Drawing.Point(0, 10)
        Me.StatusStrip4.Name = "StatusStrip4"
        Me.StatusStrip4.Size = New System.Drawing.Size(10, 53)
        Me.StatusStrip4.SizingGrip = False
        Me.StatusStrip4.TabIndex = 28
        Me.StatusStrip4.Text = "StatusStrip4"
        '
        'StatusStrip3
        '
        Me.StatusStrip3.AutoSize = False
        Me.StatusStrip3.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip3.Location = New System.Drawing.Point(0, 63)
        Me.StatusStrip3.Name = "StatusStrip3"
        Me.StatusStrip3.Size = New System.Drawing.Size(618, 10)
        Me.StatusStrip3.SizingGrip = False
        Me.StatusStrip3.TabIndex = 27
        Me.StatusStrip3.Text = "StatusStrip3"
        '
        'StatusStrip10
        '
        Me.StatusStrip10.AutoSize = False
        Me.StatusStrip10.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip10.Dock = System.Windows.Forms.DockStyle.Top
        Me.StatusStrip10.Location = New System.Drawing.Point(0, 0)
        Me.StatusStrip10.Name = "StatusStrip10"
        Me.StatusStrip10.Size = New System.Drawing.Size(618, 10)
        Me.StatusStrip10.SizingGrip = False
        Me.StatusStrip10.TabIndex = 26
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
        Me.SplitContainer2.Panel1.Controls.Add(Me.lstClasses)
        '
        'SplitContainer2.Panel2
        '
        Me.SplitContainer2.Panel2.Controls.Add(Me.StatusStrip7)
        Me.SplitContainer2.Panel2.Controls.Add(Me.btnView)
        Me.SplitContainer2.Panel2.Controls.Add(Me.btnDelete)
        Me.SplitContainer2.Panel2.Controls.Add(Me.btnUpdate)
        Me.SplitContainer2.Panel2.Controls.Add(Me.btnClose)
        Me.SplitContainer2.Panel2.Controls.Add(Me.StatusStrip6)
        Me.SplitContainer2.Panel2.Controls.Add(Me.StatusStrip8)
        Me.SplitContainer2.Panel2.Controls.Add(Me.StatusStrip9)
        Me.SplitContainer2.Size = New System.Drawing.Size(618, 324)
        Me.SplitContainer2.SplitterDistance = 255
        Me.SplitContainer2.SplitterWidth = 1
        Me.SplitContainer2.TabIndex = 22
        '
        'lstClasses
        '
        Me.lstClasses.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColumnHeader1, Me.ColumnHeader2, Me.ColumnHeader3})
        Me.lstClasses.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lstClasses.FullRowSelect = True
        Me.lstClasses.GridLines = True
        Me.lstClasses.Location = New System.Drawing.Point(0, 0)
        Me.lstClasses.Name = "lstClasses"
        Me.lstClasses.Size = New System.Drawing.Size(618, 255)
        Me.lstClasses.TabIndex = 24
        Me.lstClasses.UseCompatibleStateImageBehavior = False
        Me.lstClasses.View = System.Windows.Forms.View.Details
        '
        'ColumnHeader1
        '
        Me.ColumnHeader1.Text = "YEAR"
        Me.ColumnHeader1.Width = 170
        '
        'ColumnHeader2
        '
        Me.ColumnHeader2.Text = "CLASS NAME"
        Me.ColumnHeader2.Width = 275
        '
        'ColumnHeader3
        '
        Me.ColumnHeader3.Text = "STREAM"
        Me.ColumnHeader3.Width = 140
        '
        'StatusStrip7
        '
        Me.StatusStrip7.AutoSize = False
        Me.StatusStrip7.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip7.Dock = System.Windows.Forms.DockStyle.Left
        Me.StatusStrip7.Location = New System.Drawing.Point(0, 20)
        Me.StatusStrip7.Name = "StatusStrip7"
        Me.StatusStrip7.Size = New System.Drawing.Size(10, 38)
        Me.StatusStrip7.SizingGrip = False
        Me.StatusStrip7.TabIndex = 43
        Me.StatusStrip7.Text = "StatusStrip7"
        '
        'btnView
        '
        Me.btnView.Location = New System.Drawing.Point(163, 26)
        Me.btnView.Name = "btnView"
        Me.btnView.Size = New System.Drawing.Size(75, 23)
        Me.btnView.TabIndex = 42
        Me.btnView.Text = "View"
        Me.btnView.UseVisualStyleBackColor = True
        '
        'btnDelete
        '
        Me.btnDelete.Enabled = False
        Me.btnDelete.Location = New System.Drawing.Point(404, 26)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(75, 23)
        Me.btnDelete.TabIndex = 41
        Me.btnDelete.Text = "Delete"
        Me.btnDelete.UseVisualStyleBackColor = True
        '
        'btnUpdate
        '
        Me.btnUpdate.Enabled = False
        Me.btnUpdate.Location = New System.Drawing.Point(282, 26)
        Me.btnUpdate.Name = "btnUpdate"
        Me.btnUpdate.Size = New System.Drawing.Size(75, 23)
        Me.btnUpdate.TabIndex = 40
        Me.btnUpdate.Text = "Update"
        Me.btnUpdate.UseVisualStyleBackColor = True
        '
        'btnClose
        '
        Me.btnClose.Location = New System.Drawing.Point(525, 26)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(75, 23)
        Me.btnClose.TabIndex = 38
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'StatusStrip6
        '
        Me.StatusStrip6.AutoSize = False
        Me.StatusStrip6.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip6.Dock = System.Windows.Forms.DockStyle.Right
        Me.StatusStrip6.Location = New System.Drawing.Point(608, 20)
        Me.StatusStrip6.Name = "StatusStrip6"
        Me.StatusStrip6.Size = New System.Drawing.Size(10, 38)
        Me.StatusStrip6.SizingGrip = False
        Me.StatusStrip6.TabIndex = 34
        Me.StatusStrip6.Text = "StatusStrip6"
        '
        'StatusStrip8
        '
        Me.StatusStrip8.AutoSize = False
        Me.StatusStrip8.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip8.Location = New System.Drawing.Point(0, 58)
        Me.StatusStrip8.Name = "StatusStrip8"
        Me.StatusStrip8.Size = New System.Drawing.Size(618, 10)
        Me.StatusStrip8.SizingGrip = False
        Me.StatusStrip8.TabIndex = 33
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
        Me.StatusStrip9.Size = New System.Drawing.Size(618, 20)
        Me.StatusStrip9.SizingGrip = False
        Me.StatusStrip9.TabIndex = 32
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
        Me.StatusStrip2.Size = New System.Drawing.Size(618, 25)
        Me.StatusStrip2.SizingGrip = False
        Me.StatusStrip2.TabIndex = 20
        Me.StatusStrip2.Text = "StatusStrip2"
        '
        'ToolStripStatusLabel1
        '
        Me.ToolStripStatusLabel1.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ToolStripStatusLabel1.Name = "ToolStripStatusLabel1"
        Me.ToolStripStatusLabel1.Size = New System.Drawing.Size(116, 20)
        Me.ToolStripStatusLabel1.Text = "SCHOOL CLASSES"
        '
        'StatusStrip1
        '
        Me.StatusStrip1.AutoSize = False
        Me.StatusStrip1.BackColor = System.Drawing.Color.LightSeaGreen
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 423)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(618, 25)
        Me.StatusStrip1.SizingGrip = False
        Me.StatusStrip1.TabIndex = 19
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'frmSchClasses
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.LightBlue
        Me.ClientSize = New System.Drawing.Size(642, 474)
        Me.Controls.Add(Me.pnlSchClasses)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Name = "frmSchClasses"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.Text = "Classes"
        Me.pnlSchClasses.ResumeLayout(False)
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel1.PerformLayout()
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.ResumeLayout(False)
        Me.SplitContainer2.Panel1.ResumeLayout(False)
        Me.SplitContainer2.Panel2.ResumeLayout(False)
        CType(Me.SplitContainer2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer2.ResumeLayout(False)
        Me.StatusStrip9.ResumeLayout(False)
        Me.StatusStrip9.PerformLayout()
        Me.StatusStrip2.ResumeLayout(False)
        Me.StatusStrip2.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents pnlSchClasses As System.Windows.Forms.Panel
    Friend WithEvents StatusStrip2 As System.Windows.Forms.StatusStrip
    Friend WithEvents ToolStripStatusLabel1 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents SplitContainer2 As System.Windows.Forms.SplitContainer
    Friend WithEvents StatusStrip5 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip4 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip3 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip10 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip6 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip8 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip9 As System.Windows.Forms.StatusStrip
    Friend WithEvents ToolStripStatusLabel2 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents lstClasses As System.Windows.Forms.ListView
    Friend WithEvents ColumnHeader1 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader2 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader3 As System.Windows.Forms.ColumnHeader
    Friend WithEvents cboStream As System.Windows.Forms.ComboBox
    Friend WithEvents cboClassName As System.Windows.Forms.ComboBox
    Friend WithEvents cboYear As System.Windows.Forms.ComboBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents StatusStrip7 As System.Windows.Forms.StatusStrip
    Friend WithEvents btnView As System.Windows.Forms.Button
    Friend WithEvents btnDelete As System.Windows.Forms.Button
    Friend WithEvents btnUpdate As System.Windows.Forms.Button
    Friend WithEvents btnClose As System.Windows.Forms.Button
End Class
