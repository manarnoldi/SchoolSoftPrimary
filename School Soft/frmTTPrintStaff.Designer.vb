<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmTTPrintStaff
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
        Me.pnlPrintStaff = New System.Windows.Forms.Panel()
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.txtSearchTeacher = New System.Windows.Forms.TextBox()
        Me.StatusStrip6 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip7 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip8 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip9 = New System.Windows.Forms.StatusStrip()
        Me.SplitContainer2 = New System.Windows.Forms.SplitContainer()
        Me.lstTTTeacher = New System.Windows.Forms.ListView()
        Me.ColumnHeader1 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader2 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.btnPrint = New System.Windows.Forms.Button()
        Me.btnClose = New System.Windows.Forms.Button()
        Me.btnPreview = New System.Windows.Forms.Button()
        Me.StatusStrip3 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip4 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip5 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip10 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip2 = New System.Windows.Forms.StatusStrip()
        Me.ToolStripStatusLabel1 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip()
        Me.ToolStripStatusLabel2 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.btnLoad = New System.Windows.Forms.Button()
        Me.pnlPrintStaff.SuspendLayout()
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.Panel2.SuspendLayout()
        CType(Me.SplitContainer2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer2.Panel1.SuspendLayout()
        Me.SplitContainer2.Panel2.SuspendLayout()
        Me.SplitContainer2.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.StatusStrip2.SuspendLayout()
        Me.StatusStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'pnlPrintStaff
        '
        Me.pnlPrintStaff.BackColor = System.Drawing.SystemColors.Control
        Me.pnlPrintStaff.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pnlPrintStaff.Controls.Add(Me.SplitContainer1)
        Me.pnlPrintStaff.Controls.Add(Me.StatusStrip2)
        Me.pnlPrintStaff.Controls.Add(Me.StatusStrip1)
        Me.pnlPrintStaff.Location = New System.Drawing.Point(12, 12)
        Me.pnlPrintStaff.Name = "pnlPrintStaff"
        Me.pnlPrintStaff.Size = New System.Drawing.Size(466, 508)
        Me.pnlPrintStaff.TabIndex = 0
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
        Me.SplitContainer1.Panel1.Controls.Add(Me.StatusStrip6)
        Me.SplitContainer1.Panel1.Controls.Add(Me.StatusStrip7)
        Me.SplitContainer1.Panel1.Controls.Add(Me.StatusStrip8)
        Me.SplitContainer1.Panel1.Controls.Add(Me.StatusStrip9)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.SplitContainer2)
        Me.SplitContainer1.Size = New System.Drawing.Size(464, 456)
        Me.SplitContainer1.SplitterDistance = 57
        Me.SplitContainer1.SplitterWidth = 1
        Me.SplitContainer1.TabIndex = 21
        '
        'Panel2
        '
        Me.Panel2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel2.Controls.Add(Me.Label1)
        Me.Panel2.Controls.Add(Me.txtSearchTeacher)
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel2.Location = New System.Drawing.Point(5, 5)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(452, 45)
        Me.Panel2.TabIndex = 59
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(13, 15)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(118, 13)
        Me.Label1.TabIndex = 11
        Me.Label1.Text = "Quick Search Teacher:"
        '
        'txtSearchTeacher
        '
        Me.txtSearchTeacher.Location = New System.Drawing.Point(137, 12)
        Me.txtSearchTeacher.Name = "txtSearchTeacher"
        Me.txtSearchTeacher.Size = New System.Drawing.Size(300, 20)
        Me.txtSearchTeacher.TabIndex = 10
        '
        'StatusStrip6
        '
        Me.StatusStrip6.AutoSize = False
        Me.StatusStrip6.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip6.Dock = System.Windows.Forms.DockStyle.Right
        Me.StatusStrip6.Location = New System.Drawing.Point(457, 5)
        Me.StatusStrip6.Name = "StatusStrip6"
        Me.StatusStrip6.Size = New System.Drawing.Size(5, 45)
        Me.StatusStrip6.SizingGrip = False
        Me.StatusStrip6.TabIndex = 58
        Me.StatusStrip6.Text = "StatusStrip6"
        '
        'StatusStrip7
        '
        Me.StatusStrip7.AutoSize = False
        Me.StatusStrip7.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip7.Dock = System.Windows.Forms.DockStyle.Left
        Me.StatusStrip7.Location = New System.Drawing.Point(0, 5)
        Me.StatusStrip7.Name = "StatusStrip7"
        Me.StatusStrip7.Size = New System.Drawing.Size(5, 45)
        Me.StatusStrip7.SizingGrip = False
        Me.StatusStrip7.TabIndex = 57
        Me.StatusStrip7.Text = "StatusStrip7"
        '
        'StatusStrip8
        '
        Me.StatusStrip8.AutoSize = False
        Me.StatusStrip8.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip8.Location = New System.Drawing.Point(0, 50)
        Me.StatusStrip8.Name = "StatusStrip8"
        Me.StatusStrip8.Size = New System.Drawing.Size(462, 5)
        Me.StatusStrip8.SizingGrip = False
        Me.StatusStrip8.TabIndex = 56
        Me.StatusStrip8.Text = "StatusStrip8"
        '
        'StatusStrip9
        '
        Me.StatusStrip9.AutoSize = False
        Me.StatusStrip9.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip9.Dock = System.Windows.Forms.DockStyle.Top
        Me.StatusStrip9.Location = New System.Drawing.Point(0, 0)
        Me.StatusStrip9.Name = "StatusStrip9"
        Me.StatusStrip9.Size = New System.Drawing.Size(462, 5)
        Me.StatusStrip9.SizingGrip = False
        Me.StatusStrip9.TabIndex = 55
        Me.StatusStrip9.Text = "StatusStrip9"
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
        Me.SplitContainer2.Panel1.Controls.Add(Me.lstTTTeacher)
        '
        'SplitContainer2.Panel2
        '
        Me.SplitContainer2.Panel2.Controls.Add(Me.Panel1)
        Me.SplitContainer2.Panel2.Controls.Add(Me.StatusStrip3)
        Me.SplitContainer2.Panel2.Controls.Add(Me.StatusStrip4)
        Me.SplitContainer2.Panel2.Controls.Add(Me.StatusStrip5)
        Me.SplitContainer2.Panel2.Controls.Add(Me.StatusStrip10)
        Me.SplitContainer2.Size = New System.Drawing.Size(462, 396)
        Me.SplitContainer2.SplitterDistance = 341
        Me.SplitContainer2.SplitterWidth = 1
        Me.SplitContainer2.TabIndex = 0
        '
        'lstTTTeacher
        '
        Me.lstTTTeacher.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.lstTTTeacher.CheckBoxes = True
        Me.lstTTTeacher.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColumnHeader1, Me.ColumnHeader2})
        Me.lstTTTeacher.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lstTTTeacher.GridLines = True
        Me.lstTTTeacher.Location = New System.Drawing.Point(0, 0)
        Me.lstTTTeacher.Name = "lstTTTeacher"
        Me.lstTTTeacher.Size = New System.Drawing.Size(462, 341)
        Me.lstTTTeacher.TabIndex = 1
        Me.lstTTTeacher.UseCompatibleStateImageBehavior = False
        Me.lstTTTeacher.View = System.Windows.Forms.View.Details
        '
        'ColumnHeader1
        '
        Me.ColumnHeader1.Text = "Emp No"
        Me.ColumnHeader1.Width = 80
        '
        'ColumnHeader2
        '
        Me.ColumnHeader2.Text = "Teacher Name"
        Me.ColumnHeader2.Width = 350
        '
        'Panel1
        '
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel1.Controls.Add(Me.btnLoad)
        Me.Panel1.Controls.Add(Me.btnPrint)
        Me.Panel1.Controls.Add(Me.btnClose)
        Me.Panel1.Controls.Add(Me.btnPreview)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel1.Location = New System.Drawing.Point(5, 5)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(452, 44)
        Me.Panel1.TabIndex = 64
        '
        'btnPrint
        '
        Me.btnPrint.Location = New System.Drawing.Point(245, 9)
        Me.btnPrint.Name = "btnPrint"
        Me.btnPrint.Size = New System.Drawing.Size(75, 23)
        Me.btnPrint.TabIndex = 4
        Me.btnPrint.Text = "Print"
        Me.btnPrint.UseVisualStyleBackColor = True
        '
        'btnClose
        '
        Me.btnClose.Location = New System.Drawing.Point(362, 9)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(75, 23)
        Me.btnClose.TabIndex = 3
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'btnPreview
        '
        Me.btnPreview.Location = New System.Drawing.Point(128, 9)
        Me.btnPreview.Name = "btnPreview"
        Me.btnPreview.Size = New System.Drawing.Size(75, 23)
        Me.btnPreview.TabIndex = 2
        Me.btnPreview.Text = "Preview"
        Me.btnPreview.UseVisualStyleBackColor = True
        '
        'StatusStrip3
        '
        Me.StatusStrip3.AutoSize = False
        Me.StatusStrip3.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip3.Dock = System.Windows.Forms.DockStyle.Right
        Me.StatusStrip3.Location = New System.Drawing.Point(457, 5)
        Me.StatusStrip3.Name = "StatusStrip3"
        Me.StatusStrip3.Size = New System.Drawing.Size(5, 44)
        Me.StatusStrip3.SizingGrip = False
        Me.StatusStrip3.TabIndex = 63
        Me.StatusStrip3.Text = "StatusStrip3"
        '
        'StatusStrip4
        '
        Me.StatusStrip4.AutoSize = False
        Me.StatusStrip4.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip4.Dock = System.Windows.Forms.DockStyle.Left
        Me.StatusStrip4.Location = New System.Drawing.Point(0, 5)
        Me.StatusStrip4.Name = "StatusStrip4"
        Me.StatusStrip4.Size = New System.Drawing.Size(5, 44)
        Me.StatusStrip4.SizingGrip = False
        Me.StatusStrip4.TabIndex = 62
        Me.StatusStrip4.Text = "StatusStrip4"
        '
        'StatusStrip5
        '
        Me.StatusStrip5.AutoSize = False
        Me.StatusStrip5.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip5.Location = New System.Drawing.Point(0, 49)
        Me.StatusStrip5.Name = "StatusStrip5"
        Me.StatusStrip5.Size = New System.Drawing.Size(462, 5)
        Me.StatusStrip5.SizingGrip = False
        Me.StatusStrip5.TabIndex = 61
        Me.StatusStrip5.Text = "StatusStrip5"
        '
        'StatusStrip10
        '
        Me.StatusStrip10.AutoSize = False
        Me.StatusStrip10.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip10.Dock = System.Windows.Forms.DockStyle.Top
        Me.StatusStrip10.Location = New System.Drawing.Point(0, 0)
        Me.StatusStrip10.Name = "StatusStrip10"
        Me.StatusStrip10.Size = New System.Drawing.Size(462, 5)
        Me.StatusStrip10.SizingGrip = False
        Me.StatusStrip10.TabIndex = 60
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
        Me.StatusStrip2.Size = New System.Drawing.Size(464, 25)
        Me.StatusStrip2.SizingGrip = False
        Me.StatusStrip2.TabIndex = 20
        Me.StatusStrip2.Text = "StatusStrip2"
        '
        'ToolStripStatusLabel1
        '
        Me.ToolStripStatusLabel1.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ToolStripStatusLabel1.Name = "ToolStripStatusLabel1"
        Me.ToolStripStatusLabel1.Size = New System.Drawing.Size(218, 20)
        Me.ToolStripStatusLabel1.Text = "TIMETABLE PRINT TEACHER COPY"
        '
        'StatusStrip1
        '
        Me.StatusStrip1.AutoSize = False
        Me.StatusStrip1.BackColor = System.Drawing.Color.LightSeaGreen
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripStatusLabel2})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 481)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(464, 25)
        Me.StatusStrip1.SizingGrip = False
        Me.StatusStrip1.TabIndex = 19
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'ToolStripStatusLabel2
        '
        Me.ToolStripStatusLabel2.Name = "ToolStripStatusLabel2"
        Me.ToolStripStatusLabel2.Size = New System.Drawing.Size(84, 20)
        Me.ToolStripStatusLabel2.Text = "Printing Status"
        '
        'btnLoad
        '
        Me.btnLoad.Location = New System.Drawing.Point(16, 9)
        Me.btnLoad.Name = "btnLoad"
        Me.btnLoad.Size = New System.Drawing.Size(75, 23)
        Me.btnLoad.TabIndex = 5
        Me.btnLoad.Text = "Load"
        Me.btnLoad.UseVisualStyleBackColor = True
        '
        'frmTTPrintStaff
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.LightSkyBlue
        Me.ClientSize = New System.Drawing.Size(487, 532)
        Me.Controls.Add(Me.pnlPrintStaff)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Name = "frmTTPrintStaff"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Print Staff"
        Me.pnlPrintStaff.ResumeLayout(False)
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.ResumeLayout(False)
        Me.Panel2.ResumeLayout(False)
        Me.Panel2.PerformLayout()
        Me.SplitContainer2.Panel1.ResumeLayout(False)
        Me.SplitContainer2.Panel2.ResumeLayout(False)
        CType(Me.SplitContainer2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer2.ResumeLayout(False)
        Me.Panel1.ResumeLayout(False)
        Me.StatusStrip2.ResumeLayout(False)
        Me.StatusStrip2.PerformLayout()
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents pnlPrintStaff As System.Windows.Forms.Panel
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents StatusStrip2 As System.Windows.Forms.StatusStrip
    Friend WithEvents ToolStripStatusLabel1 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents StatusStrip6 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip7 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip8 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip9 As System.Windows.Forms.StatusStrip
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtSearchTeacher As System.Windows.Forms.TextBox
    Friend WithEvents SplitContainer2 As System.Windows.Forms.SplitContainer
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents btnPrint As System.Windows.Forms.Button
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents btnPreview As System.Windows.Forms.Button
    Friend WithEvents StatusStrip3 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip4 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip5 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip10 As System.Windows.Forms.StatusStrip
    Friend WithEvents lstTTTeacher As System.Windows.Forms.ListView
    Friend WithEvents ColumnHeader1 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader2 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ToolStripStatusLabel2 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents btnLoad As System.Windows.Forms.Button
End Class
