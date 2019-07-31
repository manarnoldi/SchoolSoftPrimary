<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmSchMaxSubjectSetUp
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
        Me.pnlStudSubject = New System.Windows.Forms.Panel()
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.txtMaxNo = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.StatusStrip12 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip13 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip14 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip5 = New System.Windows.Forms.StatusStrip()
        Me.SplitContainer2 = New System.Windows.Forms.SplitContainer()
        Me.lstMaxSubjects = New System.Windows.Forms.ListView()
        Me.ColumnHeader1 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader2 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader3 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ContextMenuStrip1 = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.UPDATEToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.DELETEToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.CLOSEToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.btnView = New System.Windows.Forms.Button()
        Me.StatusStrip3 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip4 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip6 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip7 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip2 = New System.Windows.Forms.StatusStrip()
        Me.ToolStripStatusLabel1 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip()
        Me.btnClear = New System.Windows.Forms.Button()
        Me.btnSaveChanges = New System.Windows.Forms.Button()
        Me.btnClose = New System.Windows.Forms.Button()
        Me.txtParameterName = New System.Windows.Forms.TextBox()
        Me.txtDescription = New System.Windows.Forms.TextBox()
        Me.pnlStudSubject.SuspendLayout()
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.Panel1.SuspendLayout()
        CType(Me.SplitContainer2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer2.Panel1.SuspendLayout()
        Me.SplitContainer2.Panel2.SuspendLayout()
        Me.SplitContainer2.SuspendLayout()
        Me.ContextMenuStrip1.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.StatusStrip2.SuspendLayout()
        Me.SuspendLayout()
        '
        'pnlStudSubject
        '
        Me.pnlStudSubject.BackColor = System.Drawing.SystemColors.Control
        Me.pnlStudSubject.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pnlStudSubject.Controls.Add(Me.SplitContainer1)
        Me.pnlStudSubject.Controls.Add(Me.StatusStrip2)
        Me.pnlStudSubject.Controls.Add(Me.StatusStrip1)
        Me.pnlStudSubject.Location = New System.Drawing.Point(12, 12)
        Me.pnlStudSubject.Name = "pnlStudSubject"
        Me.pnlStudSubject.Size = New System.Drawing.Size(584, 446)
        Me.pnlStudSubject.TabIndex = 0
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
        Me.SplitContainer1.Panel1.Controls.Add(Me.StatusStrip12)
        Me.SplitContainer1.Panel1.Controls.Add(Me.StatusStrip13)
        Me.SplitContainer1.Panel1.Controls.Add(Me.StatusStrip14)
        Me.SplitContainer1.Panel1.Controls.Add(Me.StatusStrip5)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.SplitContainer2)
        Me.SplitContainer1.Size = New System.Drawing.Size(582, 394)
        Me.SplitContainer1.SplitterDistance = 74
        Me.SplitContainer1.SplitterWidth = 1
        Me.SplitContainer1.TabIndex = 31
        '
        'Panel1
        '
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel1.Controls.Add(Me.txtDescription)
        Me.Panel1.Controls.Add(Me.txtParameterName)
        Me.Panel1.Controls.Add(Me.txtMaxNo)
        Me.Panel1.Controls.Add(Me.Label3)
        Me.Panel1.Controls.Add(Me.Label2)
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel1.Location = New System.Drawing.Point(5, 5)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(570, 62)
        Me.Panel1.TabIndex = 53
        '
        'txtMaxNo
        '
        Me.txtMaxNo.Location = New System.Drawing.Point(116, 34)
        Me.txtMaxNo.Name = "txtMaxNo"
        Me.txtMaxNo.Size = New System.Drawing.Size(112, 20)
        Me.txtMaxNo.TabIndex = 59
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(339, 9)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(63, 13)
        Me.Label3.TabIndex = 8
        Me.Label3.Text = "Description:"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(16, 37)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(94, 13)
        Me.Label2.TabIndex = 7
        Me.Label2.Text = "Maximum Number:"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(16, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(89, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Parameter Name:"
        '
        'StatusStrip12
        '
        Me.StatusStrip12.AutoSize = False
        Me.StatusStrip12.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip12.Dock = System.Windows.Forms.DockStyle.Left
        Me.StatusStrip12.Location = New System.Drawing.Point(0, 5)
        Me.StatusStrip12.Name = "StatusStrip12"
        Me.StatusStrip12.Size = New System.Drawing.Size(5, 62)
        Me.StatusStrip12.SizingGrip = False
        Me.StatusStrip12.TabIndex = 52
        Me.StatusStrip12.Text = "StatusStrip12"
        '
        'StatusStrip13
        '
        Me.StatusStrip13.AutoSize = False
        Me.StatusStrip13.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip13.Location = New System.Drawing.Point(0, 67)
        Me.StatusStrip13.Name = "StatusStrip13"
        Me.StatusStrip13.Size = New System.Drawing.Size(575, 5)
        Me.StatusStrip13.SizingGrip = False
        Me.StatusStrip13.TabIndex = 51
        Me.StatusStrip13.Text = "StatusStrip13"
        '
        'StatusStrip14
        '
        Me.StatusStrip14.AutoSize = False
        Me.StatusStrip14.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip14.Dock = System.Windows.Forms.DockStyle.Top
        Me.StatusStrip14.Location = New System.Drawing.Point(0, 0)
        Me.StatusStrip14.Name = "StatusStrip14"
        Me.StatusStrip14.Size = New System.Drawing.Size(575, 5)
        Me.StatusStrip14.SizingGrip = False
        Me.StatusStrip14.TabIndex = 50
        Me.StatusStrip14.Text = "StatusStrip14"
        '
        'StatusStrip5
        '
        Me.StatusStrip5.AutoSize = False
        Me.StatusStrip5.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip5.Dock = System.Windows.Forms.DockStyle.Right
        Me.StatusStrip5.Location = New System.Drawing.Point(575, 0)
        Me.StatusStrip5.Name = "StatusStrip5"
        Me.StatusStrip5.Size = New System.Drawing.Size(5, 72)
        Me.StatusStrip5.SizingGrip = False
        Me.StatusStrip5.TabIndex = 49
        Me.StatusStrip5.Text = "StatusStrip5"
        '
        'SplitContainer2
        '
        Me.SplitContainer2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.SplitContainer2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer2.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer2.Name = "SplitContainer2"
        Me.SplitContainer2.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'SplitContainer2.Panel1
        '
        Me.SplitContainer2.Panel1.Controls.Add(Me.lstMaxSubjects)
        '
        'SplitContainer2.Panel2
        '
        Me.SplitContainer2.Panel2.Controls.Add(Me.Panel2)
        Me.SplitContainer2.Panel2.Controls.Add(Me.StatusStrip3)
        Me.SplitContainer2.Panel2.Controls.Add(Me.StatusStrip4)
        Me.SplitContainer2.Panel2.Controls.Add(Me.StatusStrip6)
        Me.SplitContainer2.Panel2.Controls.Add(Me.StatusStrip7)
        Me.SplitContainer2.Size = New System.Drawing.Size(582, 319)
        Me.SplitContainer2.SplitterDistance = 271
        Me.SplitContainer2.SplitterWidth = 1
        Me.SplitContainer2.TabIndex = 0
        '
        'lstMaxSubjects
        '
        Me.lstMaxSubjects.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColumnHeader1, Me.ColumnHeader2, Me.ColumnHeader3})
        Me.lstMaxSubjects.ContextMenuStrip = Me.ContextMenuStrip1
        Me.lstMaxSubjects.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lstMaxSubjects.FullRowSelect = True
        Me.lstMaxSubjects.GridLines = True
        Me.lstMaxSubjects.Location = New System.Drawing.Point(0, 0)
        Me.lstMaxSubjects.Name = "lstMaxSubjects"
        Me.lstMaxSubjects.Size = New System.Drawing.Size(580, 269)
        Me.lstMaxSubjects.TabIndex = 1
        Me.lstMaxSubjects.UseCompatibleStateImageBehavior = False
        Me.lstMaxSubjects.View = System.Windows.Forms.View.Details
        '
        'ColumnHeader1
        '
        Me.ColumnHeader1.Text = "Parameter Name"
        Me.ColumnHeader1.Width = 250
        '
        'ColumnHeader2
        '
        Me.ColumnHeader2.Text = "Description"
        Me.ColumnHeader2.Width = 150
        '
        'ColumnHeader3
        '
        Me.ColumnHeader3.Text = "Max Number"
        Me.ColumnHeader3.Width = 150
        '
        'ContextMenuStrip1
        '
        Me.ContextMenuStrip1.BackColor = System.Drawing.Color.LightSeaGreen
        Me.ContextMenuStrip1.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ContextMenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.UPDATEToolStripMenuItem, Me.DELETEToolStripMenuItem, Me.CLOSEToolStripMenuItem})
        Me.ContextMenuStrip1.Name = "ContextMenuStrip1"
        Me.ContextMenuStrip1.Size = New System.Drawing.Size(127, 70)
        '
        'UPDATEToolStripMenuItem
        '
        Me.UPDATEToolStripMenuItem.Name = "UPDATEToolStripMenuItem"
        Me.UPDATEToolStripMenuItem.Size = New System.Drawing.Size(126, 22)
        Me.UPDATEToolStripMenuItem.Text = "UPDATE"
        '
        'DELETEToolStripMenuItem
        '
        Me.DELETEToolStripMenuItem.Name = "DELETEToolStripMenuItem"
        Me.DELETEToolStripMenuItem.Size = New System.Drawing.Size(126, 22)
        Me.DELETEToolStripMenuItem.Text = "DELETE"
        '
        'CLOSEToolStripMenuItem
        '
        Me.CLOSEToolStripMenuItem.Name = "CLOSEToolStripMenuItem"
        Me.CLOSEToolStripMenuItem.Size = New System.Drawing.Size(126, 22)
        Me.CLOSEToolStripMenuItem.Text = "CLOSE"
        '
        'Panel2
        '
        Me.Panel2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel2.Controls.Add(Me.btnClear)
        Me.Panel2.Controls.Add(Me.btnSaveChanges)
        Me.Panel2.Controls.Add(Me.btnClose)
        Me.Panel2.Controls.Add(Me.btnView)
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel2.Location = New System.Drawing.Point(5, 5)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(570, 35)
        Me.Panel2.TabIndex = 54
        '
        'btnView
        '
        Me.btnView.Location = New System.Drawing.Point(3, 5)
        Me.btnView.Name = "btnView"
        Me.btnView.Size = New System.Drawing.Size(75, 23)
        Me.btnView.TabIndex = 0
        Me.btnView.Text = "View"
        Me.btnView.UseVisualStyleBackColor = True
        '
        'StatusStrip3
        '
        Me.StatusStrip3.AutoSize = False
        Me.StatusStrip3.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip3.Dock = System.Windows.Forms.DockStyle.Left
        Me.StatusStrip3.Location = New System.Drawing.Point(0, 5)
        Me.StatusStrip3.Name = "StatusStrip3"
        Me.StatusStrip3.Size = New System.Drawing.Size(5, 35)
        Me.StatusStrip3.SizingGrip = False
        Me.StatusStrip3.TabIndex = 52
        Me.StatusStrip3.Text = "StatusStrip3"
        '
        'StatusStrip4
        '
        Me.StatusStrip4.AutoSize = False
        Me.StatusStrip4.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip4.Location = New System.Drawing.Point(0, 40)
        Me.StatusStrip4.Name = "StatusStrip4"
        Me.StatusStrip4.Size = New System.Drawing.Size(575, 5)
        Me.StatusStrip4.SizingGrip = False
        Me.StatusStrip4.TabIndex = 51
        Me.StatusStrip4.Text = "StatusStrip4"
        '
        'StatusStrip6
        '
        Me.StatusStrip6.AutoSize = False
        Me.StatusStrip6.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip6.Dock = System.Windows.Forms.DockStyle.Top
        Me.StatusStrip6.Location = New System.Drawing.Point(0, 0)
        Me.StatusStrip6.Name = "StatusStrip6"
        Me.StatusStrip6.Size = New System.Drawing.Size(575, 5)
        Me.StatusStrip6.SizingGrip = False
        Me.StatusStrip6.TabIndex = 50
        Me.StatusStrip6.Text = "StatusStrip6"
        '
        'StatusStrip7
        '
        Me.StatusStrip7.AutoSize = False
        Me.StatusStrip7.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip7.Dock = System.Windows.Forms.DockStyle.Right
        Me.StatusStrip7.Location = New System.Drawing.Point(575, 0)
        Me.StatusStrip7.Name = "StatusStrip7"
        Me.StatusStrip7.Size = New System.Drawing.Size(5, 45)
        Me.StatusStrip7.SizingGrip = False
        Me.StatusStrip7.TabIndex = 49
        Me.StatusStrip7.Text = "StatusStrip7"
        '
        'StatusStrip2
        '
        Me.StatusStrip2.AutoSize = False
        Me.StatusStrip2.BackColor = System.Drawing.Color.LightSeaGreen
        Me.StatusStrip2.Dock = System.Windows.Forms.DockStyle.Top
        Me.StatusStrip2.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripStatusLabel1})
        Me.StatusStrip2.Location = New System.Drawing.Point(0, 0)
        Me.StatusStrip2.Name = "StatusStrip2"
        Me.StatusStrip2.Size = New System.Drawing.Size(582, 25)
        Me.StatusStrip2.SizingGrip = False
        Me.StatusStrip2.TabIndex = 30
        Me.StatusStrip2.Text = "StatusStrip2"
        '
        'ToolStripStatusLabel1
        '
        Me.ToolStripStatusLabel1.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ToolStripStatusLabel1.Name = "ToolStripStatusLabel1"
        Me.ToolStripStatusLabel1.Size = New System.Drawing.Size(182, 20)
        Me.ToolStripStatusLabel1.Text = "MAXIMUM SUBJECTS SETUP"
        '
        'StatusStrip1
        '
        Me.StatusStrip1.AutoSize = False
        Me.StatusStrip1.BackColor = System.Drawing.Color.LightSeaGreen
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 419)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(582, 25)
        Me.StatusStrip1.SizingGrip = False
        Me.StatusStrip1.TabIndex = 29
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'btnClear
        '
        Me.btnClear.Location = New System.Drawing.Point(153, 5)
        Me.btnClear.Name = "btnClear"
        Me.btnClear.Size = New System.Drawing.Size(75, 23)
        Me.btnClear.TabIndex = 12
        Me.btnClear.Text = "Clear"
        Me.btnClear.UseVisualStyleBackColor = True
        '
        'btnSaveChanges
        '
        Me.btnSaveChanges.Location = New System.Drawing.Point(300, 5)
        Me.btnSaveChanges.Name = "btnSaveChanges"
        Me.btnSaveChanges.Size = New System.Drawing.Size(109, 23)
        Me.btnSaveChanges.TabIndex = 11
        Me.btnSaveChanges.Text = "Update Changes"
        Me.btnSaveChanges.UseVisualStyleBackColor = True
        '
        'btnClose
        '
        Me.btnClose.Location = New System.Drawing.Point(481, 5)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(75, 23)
        Me.btnClose.TabIndex = 10
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'txtParameterName
        '
        Me.txtParameterName.Location = New System.Drawing.Point(116, 6)
        Me.txtParameterName.Name = "txtParameterName"
        Me.txtParameterName.ReadOnly = True
        Me.txtParameterName.Size = New System.Drawing.Size(217, 20)
        Me.txtParameterName.TabIndex = 60
        '
        'txtDescription
        '
        Me.txtDescription.Location = New System.Drawing.Point(408, 6)
        Me.txtDescription.Name = "txtDescription"
        Me.txtDescription.ReadOnly = True
        Me.txtDescription.Size = New System.Drawing.Size(148, 20)
        Me.txtDescription.TabIndex = 61
        '
        'frmSchMaxSubjectSetUp
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.LightSkyBlue
        Me.ClientSize = New System.Drawing.Size(607, 470)
        Me.Controls.Add(Me.pnlStudSubject)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Name = "frmSchMaxSubjectSetUp"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Max Subjects"
        Me.pnlStudSubject.ResumeLayout(False)
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
        Me.ContextMenuStrip1.ResumeLayout(False)
        Me.Panel2.ResumeLayout(False)
        Me.StatusStrip2.ResumeLayout(False)
        Me.StatusStrip2.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents pnlStudSubject As System.Windows.Forms.Panel
    Friend WithEvents StatusStrip2 As System.Windows.Forms.StatusStrip
    Friend WithEvents ToolStripStatusLabel1 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents SplitContainer2 As System.Windows.Forms.SplitContainer
    Friend WithEvents lstMaxSubjects As System.Windows.Forms.ListView
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents txtMaxNo As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents StatusStrip12 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip13 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip14 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip5 As System.Windows.Forms.StatusStrip
    Friend WithEvents ColumnHeader1 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader2 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader3 As System.Windows.Forms.ColumnHeader
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents btnView As System.Windows.Forms.Button
    Friend WithEvents StatusStrip3 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip4 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip6 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip7 As System.Windows.Forms.StatusStrip
    Friend WithEvents ContextMenuStrip1 As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents UPDATEToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents DELETEToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents CLOSEToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents btnClear As System.Windows.Forms.Button
    Friend WithEvents btnSaveChanges As System.Windows.Forms.Button
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents txtDescription As System.Windows.Forms.TextBox
    Friend WithEvents txtParameterName As System.Windows.Forms.TextBox
End Class
