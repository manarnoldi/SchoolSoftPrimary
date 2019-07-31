<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmAcadExamResultsView
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
        Me.pnlAcadExamResultsView = New System.Windows.Forms.Panel()
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.cboExamType = New System.Windows.Forms.ComboBox()
        Me.btnClose = New System.Windows.Forms.Button()
        Me.btnLoad = New System.Windows.Forms.Button()
        Me.btnExport = New System.Windows.Forms.Button()
        Me.cboSubject = New System.Windows.Forms.ComboBox()
        Me.cboExam = New System.Windows.Forms.ComboBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.cboTerm = New System.Windows.Forms.ComboBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.cboStream = New System.Windows.Forms.ComboBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.cboClass = New System.Windows.Forms.ComboBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.cboYear = New System.Windows.Forms.ComboBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.StatusStrip11 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip12 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip13 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip14 = New System.Windows.Forms.StatusStrip()
        Me.lstStaffSubject = New System.Windows.Forms.ListView()
        Me.ColumnHeader1 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader2 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader3 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader4 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader5 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader6 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader9 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader7 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.StatusStrip2 = New System.Windows.Forms.StatusStrip()
        Me.ToolStripStatusLabel1 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.pnlAcadExamResultsView.SuspendLayout()
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.StatusStrip2.SuspendLayout()
        Me.SuspendLayout()
        '
        'pnlAcadExamResultsView
        '
        Me.pnlAcadExamResultsView.BackColor = System.Drawing.SystemColors.Control
        Me.pnlAcadExamResultsView.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pnlAcadExamResultsView.Controls.Add(Me.SplitContainer1)
        Me.pnlAcadExamResultsView.Controls.Add(Me.StatusStrip2)
        Me.pnlAcadExamResultsView.Controls.Add(Me.StatusStrip1)
        Me.pnlAcadExamResultsView.Location = New System.Drawing.Point(12, 12)
        Me.pnlAcadExamResultsView.Name = "pnlAcadExamResultsView"
        Me.pnlAcadExamResultsView.Size = New System.Drawing.Size(904, 526)
        Me.pnlAcadExamResultsView.TabIndex = 0
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
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label7)
        Me.SplitContainer1.Panel1.Controls.Add(Me.cboExamType)
        Me.SplitContainer1.Panel1.Controls.Add(Me.btnClose)
        Me.SplitContainer1.Panel1.Controls.Add(Me.btnLoad)
        Me.SplitContainer1.Panel1.Controls.Add(Me.btnExport)
        Me.SplitContainer1.Panel1.Controls.Add(Me.cboSubject)
        Me.SplitContainer1.Panel1.Controls.Add(Me.cboExam)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label5)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label4)
        Me.SplitContainer1.Panel1.Controls.Add(Me.cboTerm)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label6)
        Me.SplitContainer1.Panel1.Controls.Add(Me.cboStream)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label3)
        Me.SplitContainer1.Panel1.Controls.Add(Me.cboClass)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label2)
        Me.SplitContainer1.Panel1.Controls.Add(Me.cboYear)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label1)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Panel1)
        Me.SplitContainer1.Panel1.Controls.Add(Me.StatusStrip11)
        Me.SplitContainer1.Panel1.Controls.Add(Me.StatusStrip12)
        Me.SplitContainer1.Panel1.Controls.Add(Me.StatusStrip13)
        Me.SplitContainer1.Panel1.Controls.Add(Me.StatusStrip14)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.lstStaffSubject)
        Me.SplitContainer1.Size = New System.Drawing.Size(902, 474)
        Me.SplitContainer1.SplitterDistance = 85
        Me.SplitContainer1.SplitterWidth = 1
        Me.SplitContainer1.TabIndex = 21
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(489, 51)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(67, 13)
        Me.Label7.TabIndex = 123
        Me.Label7.Text = "Exam Name:"
        '
        'cboExamType
        '
        Me.cboExamType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboExamType.FormattingEnabled = True
        Me.cboExamType.Location = New System.Drawing.Point(372, 48)
        Me.cboExamType.Name = "cboExamType"
        Me.cboExamType.Size = New System.Drawing.Size(111, 21)
        Me.cboExamType.TabIndex = 122
        '
        'btnClose
        '
        Me.btnClose.Location = New System.Drawing.Point(813, 13)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(75, 23)
        Me.btnClose.TabIndex = 121
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'btnLoad
        '
        Me.btnLoad.Location = New System.Drawing.Point(690, 46)
        Me.btnLoad.Name = "btnLoad"
        Me.btnLoad.Size = New System.Drawing.Size(75, 23)
        Me.btnLoad.TabIndex = 120
        Me.btnLoad.Text = "Load"
        Me.btnLoad.UseVisualStyleBackColor = True
        '
        'btnExport
        '
        Me.btnExport.Location = New System.Drawing.Point(779, 46)
        Me.btnExport.Name = "btnExport"
        Me.btnExport.Size = New System.Drawing.Size(109, 23)
        Me.btnExport.TabIndex = 119
        Me.btnExport.Text = "Export To Excel"
        Me.btnExport.UseVisualStyleBackColor = True
        '
        'cboSubject
        '
        Me.cboSubject.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboSubject.FormattingEnabled = True
        Me.cboSubject.Location = New System.Drawing.Point(98, 48)
        Me.cboSubject.Name = "cboSubject"
        Me.cboSubject.Size = New System.Drawing.Size(200, 21)
        Me.cboSubject.TabIndex = 118
        '
        'cboExam
        '
        Me.cboExam.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboExam.FormattingEnabled = True
        Me.cboExam.Location = New System.Drawing.Point(562, 48)
        Me.cboExam.Name = "cboExam"
        Me.cboExam.Size = New System.Drawing.Size(122, 21)
        Me.cboExam.TabIndex = 118
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(15, 51)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(77, 13)
        Me.Label5.TabIndex = 117
        Me.Label5.Text = "Subject Name:"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(304, 51)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(63, 13)
        Me.Label4.TabIndex = 117
        Me.Label4.Text = "Exam Type:"
        '
        'cboTerm
        '
        Me.cboTerm.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboTerm.FormattingEnabled = True
        Me.cboTerm.Location = New System.Drawing.Point(250, 15)
        Me.cboTerm.Name = "cboTerm"
        Me.cboTerm.Size = New System.Drawing.Size(125, 21)
        Me.cboTerm.TabIndex = 116
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(208, 18)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(34, 13)
        Me.Label6.TabIndex = 115
        Me.Label6.Text = "Term:"
        '
        'cboStream
        '
        Me.cboStream.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboStream.FormattingEnabled = True
        Me.cboStream.Location = New System.Drawing.Point(690, 15)
        Me.cboStream.Name = "cboStream"
        Me.cboStream.Size = New System.Drawing.Size(99, 21)
        Me.cboStream.TabIndex = 114
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(613, 18)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(71, 13)
        Me.Label3.TabIndex = 113
        Me.Label3.Text = "Class Stream:"
        '
        'cboClass
        '
        Me.cboClass.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboClass.FormattingEnabled = True
        Me.cboClass.Location = New System.Drawing.Point(467, 15)
        Me.cboClass.Name = "cboClass"
        Me.cboClass.Size = New System.Drawing.Size(124, 21)
        Me.cboClass.TabIndex = 112
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(395, 18)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(66, 13)
        Me.Label2.TabIndex = 111
        Me.Label2.Text = "Class Name:"
        '
        'cboYear
        '
        Me.cboYear.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboYear.FormattingEnabled = True
        Me.cboYear.Location = New System.Drawing.Point(83, 15)
        Me.cboYear.Name = "cboYear"
        Me.cboYear.Size = New System.Drawing.Size(100, 21)
        Me.cboYear.TabIndex = 110
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(15, 18)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(63, 13)
        Me.Label1.TabIndex = 109
        Me.Label1.Text = "Year Name:"
        '
        'StatusStrip11
        '
        Me.StatusStrip11.AutoSize = False
        Me.StatusStrip11.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip11.Dock = System.Windows.Forms.DockStyle.Right
        Me.StatusStrip11.Location = New System.Drawing.Point(895, 5)
        Me.StatusStrip11.Name = "StatusStrip11"
        Me.StatusStrip11.Size = New System.Drawing.Size(5, 73)
        Me.StatusStrip11.SizingGrip = False
        Me.StatusStrip11.TabIndex = 41
        Me.StatusStrip11.Text = "StatusStrip11"
        '
        'StatusStrip12
        '
        Me.StatusStrip12.AutoSize = False
        Me.StatusStrip12.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip12.Dock = System.Windows.Forms.DockStyle.Left
        Me.StatusStrip12.Location = New System.Drawing.Point(0, 5)
        Me.StatusStrip12.Name = "StatusStrip12"
        Me.StatusStrip12.Size = New System.Drawing.Size(5, 73)
        Me.StatusStrip12.SizingGrip = False
        Me.StatusStrip12.TabIndex = 40
        Me.StatusStrip12.Text = "StatusStrip12"
        '
        'StatusStrip13
        '
        Me.StatusStrip13.AutoSize = False
        Me.StatusStrip13.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip13.Location = New System.Drawing.Point(0, 78)
        Me.StatusStrip13.Name = "StatusStrip13"
        Me.StatusStrip13.Size = New System.Drawing.Size(900, 5)
        Me.StatusStrip13.SizingGrip = False
        Me.StatusStrip13.TabIndex = 39
        Me.StatusStrip13.Text = "StatusStrip13"
        '
        'StatusStrip14
        '
        Me.StatusStrip14.AutoSize = False
        Me.StatusStrip14.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip14.Dock = System.Windows.Forms.DockStyle.Top
        Me.StatusStrip14.Location = New System.Drawing.Point(0, 0)
        Me.StatusStrip14.Name = "StatusStrip14"
        Me.StatusStrip14.Size = New System.Drawing.Size(900, 5)
        Me.StatusStrip14.SizingGrip = False
        Me.StatusStrip14.TabIndex = 38
        Me.StatusStrip14.Text = "StatusStrip14"
        '
        'lstStaffSubject
        '
        Me.lstStaffSubject.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColumnHeader1, Me.ColumnHeader2, Me.ColumnHeader3, Me.ColumnHeader4, Me.ColumnHeader5, Me.ColumnHeader6, Me.ColumnHeader9, Me.ColumnHeader7})
        Me.lstStaffSubject.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lstStaffSubject.FullRowSelect = True
        Me.lstStaffSubject.GridLines = True
        Me.lstStaffSubject.Location = New System.Drawing.Point(0, 0)
        Me.lstStaffSubject.Name = "lstStaffSubject"
        Me.lstStaffSubject.Size = New System.Drawing.Size(900, 386)
        Me.lstStaffSubject.TabIndex = 55
        Me.lstStaffSubject.UseCompatibleStateImageBehavior = False
        Me.lstStaffSubject.View = System.Windows.Forms.View.Details
        '
        'ColumnHeader1
        '
        Me.ColumnHeader1.Text = "Adm No"
        Me.ColumnHeader1.Width = 70
        '
        'ColumnHeader2
        '
        Me.ColumnHeader2.Text = "Student Name"
        Me.ColumnHeader2.Width = 260
        '
        'ColumnHeader3
        '
        Me.ColumnHeader3.Text = "Subject Name"
        Me.ColumnHeader3.Width = 140
        '
        'ColumnHeader4
        '
        Me.ColumnHeader4.Text = "Year"
        Me.ColumnHeader4.Width = 70
        '
        'ColumnHeader5
        '
        Me.ColumnHeader5.Text = "Examination"
        Me.ColumnHeader5.Width = 140
        '
        'ColumnHeader6
        '
        Me.ColumnHeader6.Text = "Score"
        Me.ColumnHeader6.Width = 70
        '
        'ColumnHeader9
        '
        Me.ColumnHeader9.Text = "Out Of"
        Me.ColumnHeader9.Width = 70
        '
        'ColumnHeader7
        '
        Me.ColumnHeader7.Text = "Rank"
        Me.ColumnHeader7.Width = 70
        '
        'StatusStrip2
        '
        Me.StatusStrip2.AutoSize = False
        Me.StatusStrip2.BackColor = System.Drawing.Color.LightSeaGreen
        Me.StatusStrip2.Dock = System.Windows.Forms.DockStyle.Top
        Me.StatusStrip2.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripStatusLabel1})
        Me.StatusStrip2.Location = New System.Drawing.Point(0, 0)
        Me.StatusStrip2.Name = "StatusStrip2"
        Me.StatusStrip2.Size = New System.Drawing.Size(902, 25)
        Me.StatusStrip2.SizingGrip = False
        Me.StatusStrip2.TabIndex = 20
        Me.StatusStrip2.Text = "StatusStrip2"
        '
        'ToolStripStatusLabel1
        '
        Me.ToolStripStatusLabel1.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ToolStripStatusLabel1.Name = "ToolStripStatusLabel1"
        Me.ToolStripStatusLabel1.Size = New System.Drawing.Size(138, 20)
        Me.ToolStripStatusLabel1.Text = "EXAM RESULTS VIEW"
        '
        'StatusStrip1
        '
        Me.StatusStrip1.AutoSize = False
        Me.StatusStrip1.BackColor = System.Drawing.Color.LightSeaGreen
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 499)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(902, 25)
        Me.StatusStrip1.SizingGrip = False
        Me.StatusStrip1.TabIndex = 19
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'Panel1
        '
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel1.Location = New System.Drawing.Point(5, 5)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(890, 73)
        Me.Panel1.TabIndex = 124
        '
        'frmAcadExamResultsView
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.LightSkyBlue
        Me.ClientSize = New System.Drawing.Size(928, 550)
        Me.Controls.Add(Me.pnlAcadExamResultsView)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Name = "frmAcadExamResultsView"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Results View"
        Me.pnlAcadExamResultsView.ResumeLayout(False)
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel1.PerformLayout()
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.ResumeLayout(False)
        Me.StatusStrip2.ResumeLayout(False)
        Me.StatusStrip2.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents pnlAcadExamResultsView As System.Windows.Forms.Panel
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents StatusStrip2 As System.Windows.Forms.StatusStrip
    Friend WithEvents ToolStripStatusLabel1 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip11 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip12 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip13 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip14 As System.Windows.Forms.StatusStrip
    Friend WithEvents lstStaffSubject As System.Windows.Forms.ListView
    Friend WithEvents btnExport As System.Windows.Forms.Button
    Friend WithEvents cboSubject As System.Windows.Forms.ComboBox
    Friend WithEvents cboExam As System.Windows.Forms.ComboBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents cboTerm As System.Windows.Forms.ComboBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents cboStream As System.Windows.Forms.ComboBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents cboClass As System.Windows.Forms.ComboBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents cboYear As System.Windows.Forms.ComboBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents ColumnHeader1 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader2 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader3 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader4 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader5 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader6 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader9 As System.Windows.Forms.ColumnHeader
    Friend WithEvents btnLoad As System.Windows.Forms.Button
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents ColumnHeader7 As System.Windows.Forms.ColumnHeader
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents cboExamType As System.Windows.Forms.ComboBox
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
End Class
