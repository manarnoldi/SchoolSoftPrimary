<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmStudImages
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
        Me.pnlStudImages = New System.Windows.Forms.Panel()
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
        Me.btnRotate = New System.Windows.Forms.Button()
        Me.btnClear = New System.Windows.Forms.Button()
        Me.btnSelect = New System.Windows.Forms.Button()
        Me.pbStudImage = New System.Windows.Forms.PictureBox()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.StatusStrip6 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip7 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip8 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip9 = New System.Windows.Forms.StatusStrip()
        Me.txtPicName = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.cboStudentNo = New System.Windows.Forms.ComboBox()
        Me.cboClassYear = New System.Windows.Forms.ComboBox()
        Me.btnClose = New System.Windows.Forms.Button()
        Me.cboClassName = New System.Windows.Forms.ComboBox()
        Me.btnUpdate = New System.Windows.Forms.Button()
        Me.btnDelete = New System.Windows.Forms.Button()
        Me.btnView = New System.Windows.Forms.Button()
        Me.btnSave = New System.Windows.Forms.Button()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.txtStudName = New System.Windows.Forms.TextBox()
        Me.cboClassStream = New System.Windows.Forms.ComboBox()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.StatusStrip11 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip12 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip13 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip14 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip2 = New System.Windows.Forms.StatusStrip()
        Me.ToolStripStatusLabel1 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip()
        Me.fdlg = New System.Windows.Forms.OpenFileDialog()
        Me.pnlStudImages.SuspendLayout()
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        CType(Me.pbStudImage, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.StatusStrip2.SuspendLayout()
        Me.SuspendLayout()
        '
        'pnlStudImages
        '
        Me.pnlStudImages.BackColor = System.Drawing.SystemColors.Control
        Me.pnlStudImages.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pnlStudImages.Controls.Add(Me.SplitContainer1)
        Me.pnlStudImages.Controls.Add(Me.StatusStrip2)
        Me.pnlStudImages.Controls.Add(Me.StatusStrip1)
        Me.pnlStudImages.Location = New System.Drawing.Point(12, 12)
        Me.pnlStudImages.Name = "pnlStudImages"
        Me.pnlStudImages.Size = New System.Drawing.Size(770, 399)
        Me.pnlStudImages.TabIndex = 0
        '
        'SplitContainer1
        '
        Me.SplitContainer1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.SplitContainer1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer1.Location = New System.Drawing.Point(0, 25)
        Me.SplitContainer1.Name = "SplitContainer1"
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.btnRotate)
        Me.SplitContainer1.Panel1.Controls.Add(Me.btnClear)
        Me.SplitContainer1.Panel1.Controls.Add(Me.btnSelect)
        Me.SplitContainer1.Panel1.Controls.Add(Me.pbStudImage)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Panel2)
        Me.SplitContainer1.Panel1.Controls.Add(Me.StatusStrip6)
        Me.SplitContainer1.Panel1.Controls.Add(Me.StatusStrip7)
        Me.SplitContainer1.Panel1.Controls.Add(Me.StatusStrip8)
        Me.SplitContainer1.Panel1.Controls.Add(Me.StatusStrip9)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.txtPicName)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label1)
        Me.SplitContainer1.Panel2.Controls.Add(Me.cboStudentNo)
        Me.SplitContainer1.Panel2.Controls.Add(Me.cboClassYear)
        Me.SplitContainer1.Panel2.Controls.Add(Me.btnClose)
        Me.SplitContainer1.Panel2.Controls.Add(Me.cboClassName)
        Me.SplitContainer1.Panel2.Controls.Add(Me.btnUpdate)
        Me.SplitContainer1.Panel2.Controls.Add(Me.btnDelete)
        Me.SplitContainer1.Panel2.Controls.Add(Me.btnView)
        Me.SplitContainer1.Panel2.Controls.Add(Me.btnSave)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label2)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label6)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label7)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label8)
        Me.SplitContainer1.Panel2.Controls.Add(Me.txtStudName)
        Me.SplitContainer1.Panel2.Controls.Add(Me.cboClassStream)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label10)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Panel1)
        Me.SplitContainer1.Panel2.Controls.Add(Me.StatusStrip11)
        Me.SplitContainer1.Panel2.Controls.Add(Me.StatusStrip12)
        Me.SplitContainer1.Panel2.Controls.Add(Me.StatusStrip13)
        Me.SplitContainer1.Panel2.Controls.Add(Me.StatusStrip14)
        Me.SplitContainer1.Size = New System.Drawing.Size(768, 347)
        Me.SplitContainer1.SplitterDistance = 329
        Me.SplitContainer1.SplitterWidth = 1
        Me.SplitContainer1.TabIndex = 19
        '
        'btnRotate
        '
        Me.btnRotate.Location = New System.Drawing.Point(236, 304)
        Me.btnRotate.Name = "btnRotate"
        Me.btnRotate.Size = New System.Drawing.Size(75, 23)
        Me.btnRotate.TabIndex = 58
        Me.btnRotate.Text = "Rotate"
        Me.btnRotate.UseVisualStyleBackColor = True
        '
        'btnClear
        '
        Me.btnClear.Location = New System.Drawing.Point(126, 304)
        Me.btnClear.Name = "btnClear"
        Me.btnClear.Size = New System.Drawing.Size(75, 23)
        Me.btnClear.TabIndex = 53
        Me.btnClear.Text = "Clear"
        Me.btnClear.UseVisualStyleBackColor = True
        '
        'btnSelect
        '
        Me.btnSelect.Location = New System.Drawing.Point(16, 304)
        Me.btnSelect.Name = "btnSelect"
        Me.btnSelect.Size = New System.Drawing.Size(75, 23)
        Me.btnSelect.TabIndex = 38
        Me.btnSelect.Text = "Select"
        Me.btnSelect.UseVisualStyleBackColor = True
        '
        'pbStudImage
        '
        Me.pbStudImage.BackColor = System.Drawing.SystemColors.ButtonHighlight
        Me.pbStudImage.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pbStudImage.Location = New System.Drawing.Point(16, 15)
        Me.pbStudImage.Name = "pbStudImage"
        Me.pbStudImage.Size = New System.Drawing.Size(295, 272)
        Me.pbStudImage.TabIndex = 0
        Me.pbStudImage.TabStop = False
        '
        'Panel2
        '
        Me.Panel2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel2.Location = New System.Drawing.Point(5, 5)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(317, 335)
        Me.Panel2.TabIndex = 59
        '
        'StatusStrip6
        '
        Me.StatusStrip6.AutoSize = False
        Me.StatusStrip6.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip6.Dock = System.Windows.Forms.DockStyle.Right
        Me.StatusStrip6.Location = New System.Drawing.Point(322, 5)
        Me.StatusStrip6.Name = "StatusStrip6"
        Me.StatusStrip6.Size = New System.Drawing.Size(5, 335)
        Me.StatusStrip6.SizingGrip = False
        Me.StatusStrip6.TabIndex = 37
        Me.StatusStrip6.Text = "StatusStrip6"
        '
        'StatusStrip7
        '
        Me.StatusStrip7.AutoSize = False
        Me.StatusStrip7.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip7.Dock = System.Windows.Forms.DockStyle.Left
        Me.StatusStrip7.Location = New System.Drawing.Point(0, 5)
        Me.StatusStrip7.Name = "StatusStrip7"
        Me.StatusStrip7.Size = New System.Drawing.Size(5, 335)
        Me.StatusStrip7.SizingGrip = False
        Me.StatusStrip7.TabIndex = 36
        Me.StatusStrip7.Text = "StatusStrip7"
        '
        'StatusStrip8
        '
        Me.StatusStrip8.AutoSize = False
        Me.StatusStrip8.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip8.Location = New System.Drawing.Point(0, 340)
        Me.StatusStrip8.Name = "StatusStrip8"
        Me.StatusStrip8.Size = New System.Drawing.Size(327, 5)
        Me.StatusStrip8.SizingGrip = False
        Me.StatusStrip8.TabIndex = 35
        Me.StatusStrip8.Text = "StatusStrip8"
        '
        'StatusStrip9
        '
        Me.StatusStrip9.AutoSize = False
        Me.StatusStrip9.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip9.Dock = System.Windows.Forms.DockStyle.Top
        Me.StatusStrip9.Location = New System.Drawing.Point(0, 0)
        Me.StatusStrip9.Name = "StatusStrip9"
        Me.StatusStrip9.Size = New System.Drawing.Size(327, 5)
        Me.StatusStrip9.SizingGrip = False
        Me.StatusStrip9.TabIndex = 34
        Me.StatusStrip9.Text = "StatusStrip9"
        '
        'txtPicName
        '
        Me.txtPicName.Location = New System.Drawing.Point(99, 259)
        Me.txtPicName.Name = "txtPicName"
        Me.txtPicName.ReadOnly = True
        Me.txtPicName.Size = New System.Drawing.Size(324, 20)
        Me.txtPicName.TabIndex = 57
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(19, 262)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(74, 13)
        Me.Label1.TabIndex = 56
        Me.Label1.Text = "Picture Name:"
        '
        'cboStudentNo
        '
        Me.cboStudentNo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboStudentNo.FormattingEnabled = True
        Me.cboStudentNo.Location = New System.Drawing.Point(99, 137)
        Me.cboStudentNo.Name = "cboStudentNo"
        Me.cboStudentNo.Size = New System.Drawing.Size(216, 21)
        Me.cboStudentNo.TabIndex = 55
        '
        'cboClassYear
        '
        Me.cboClassYear.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboClassYear.FormattingEnabled = True
        Me.cboClassYear.Location = New System.Drawing.Point(99, 99)
        Me.cboClassYear.Name = "cboClassYear"
        Me.cboClassYear.Size = New System.Drawing.Size(216, 21)
        Me.cboClassYear.TabIndex = 54
        '
        'btnClose
        '
        Me.btnClose.Location = New System.Drawing.Point(348, 22)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(75, 23)
        Me.btnClose.TabIndex = 49
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'cboClassName
        '
        Me.cboClassName.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboClassName.FormattingEnabled = True
        Me.cboClassName.Location = New System.Drawing.Point(99, 24)
        Me.cboClassName.Name = "cboClassName"
        Me.cboClassName.Size = New System.Drawing.Size(216, 21)
        Me.cboClassName.TabIndex = 53
        '
        'btnUpdate
        '
        Me.btnUpdate.Location = New System.Drawing.Point(242, 211)
        Me.btnUpdate.Name = "btnUpdate"
        Me.btnUpdate.Size = New System.Drawing.Size(75, 23)
        Me.btnUpdate.TabIndex = 52
        Me.btnUpdate.Text = "Update"
        Me.btnUpdate.UseVisualStyleBackColor = True
        '
        'btnDelete
        '
        Me.btnDelete.Location = New System.Drawing.Point(348, 211)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(75, 23)
        Me.btnDelete.TabIndex = 51
        Me.btnDelete.Text = "Delete"
        Me.btnDelete.UseVisualStyleBackColor = True
        '
        'btnView
        '
        Me.btnView.Location = New System.Drawing.Point(131, 211)
        Me.btnView.Name = "btnView"
        Me.btnView.Size = New System.Drawing.Size(75, 23)
        Me.btnView.TabIndex = 50
        Me.btnView.Text = "View"
        Me.btnView.UseVisualStyleBackColor = True
        '
        'btnSave
        '
        Me.btnSave.Location = New System.Drawing.Point(22, 211)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(75, 23)
        Me.btnSave.TabIndex = 48
        Me.btnSave.Text = "Save"
        Me.btnSave.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(19, 177)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(78, 13)
        Me.Label2.TabIndex = 40
        Me.Label2.Text = "Student Name:"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(19, 102)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(60, 13)
        Me.Label6.TabIndex = 44
        Me.Label6.Text = "Class Year:"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(19, 64)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(71, 13)
        Me.Label7.TabIndex = 43
        Me.Label7.Text = "Class Stream:"
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(19, 27)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(66, 13)
        Me.Label8.TabIndex = 42
        Me.Label8.Text = "Class Name:"
        '
        'txtStudName
        '
        Me.txtStudName.Location = New System.Drawing.Point(99, 174)
        Me.txtStudName.Name = "txtStudName"
        Me.txtStudName.ReadOnly = True
        Me.txtStudName.Size = New System.Drawing.Size(324, 20)
        Me.txtStudName.TabIndex = 41
        '
        'cboClassStream
        '
        Me.cboClassStream.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboClassStream.FormattingEnabled = True
        Me.cboClassStream.Location = New System.Drawing.Point(99, 61)
        Me.cboClassStream.Name = "cboClassStream"
        Me.cboClassStream.Size = New System.Drawing.Size(216, 21)
        Me.cboClassStream.TabIndex = 39
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(19, 140)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(64, 13)
        Me.Label10.TabIndex = 38
        Me.Label10.Text = "Student No:"
        '
        'Panel1
        '
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel1.Location = New System.Drawing.Point(5, 5)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(426, 335)
        Me.Panel1.TabIndex = 58
        '
        'StatusStrip11
        '
        Me.StatusStrip11.AutoSize = False
        Me.StatusStrip11.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip11.Dock = System.Windows.Forms.DockStyle.Right
        Me.StatusStrip11.Location = New System.Drawing.Point(431, 5)
        Me.StatusStrip11.Name = "StatusStrip11"
        Me.StatusStrip11.Size = New System.Drawing.Size(5, 335)
        Me.StatusStrip11.SizingGrip = False
        Me.StatusStrip11.TabIndex = 37
        Me.StatusStrip11.Text = "StatusStrip11"
        '
        'StatusStrip12
        '
        Me.StatusStrip12.AutoSize = False
        Me.StatusStrip12.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip12.Dock = System.Windows.Forms.DockStyle.Left
        Me.StatusStrip12.Location = New System.Drawing.Point(0, 5)
        Me.StatusStrip12.Name = "StatusStrip12"
        Me.StatusStrip12.Size = New System.Drawing.Size(5, 335)
        Me.StatusStrip12.SizingGrip = False
        Me.StatusStrip12.TabIndex = 36
        Me.StatusStrip12.Text = "StatusStrip12"
        '
        'StatusStrip13
        '
        Me.StatusStrip13.AutoSize = False
        Me.StatusStrip13.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip13.Location = New System.Drawing.Point(0, 340)
        Me.StatusStrip13.Name = "StatusStrip13"
        Me.StatusStrip13.Size = New System.Drawing.Size(436, 5)
        Me.StatusStrip13.SizingGrip = False
        Me.StatusStrip13.TabIndex = 35
        Me.StatusStrip13.Text = "StatusStrip13"
        '
        'StatusStrip14
        '
        Me.StatusStrip14.AutoSize = False
        Me.StatusStrip14.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip14.Dock = System.Windows.Forms.DockStyle.Top
        Me.StatusStrip14.Location = New System.Drawing.Point(0, 0)
        Me.StatusStrip14.Name = "StatusStrip14"
        Me.StatusStrip14.Size = New System.Drawing.Size(436, 5)
        Me.StatusStrip14.SizingGrip = False
        Me.StatusStrip14.TabIndex = 34
        Me.StatusStrip14.Text = "StatusStrip14"
        '
        'StatusStrip2
        '
        Me.StatusStrip2.AutoSize = False
        Me.StatusStrip2.BackColor = System.Drawing.Color.LightSeaGreen
        Me.StatusStrip2.Dock = System.Windows.Forms.DockStyle.Top
        Me.StatusStrip2.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripStatusLabel1})
        Me.StatusStrip2.Location = New System.Drawing.Point(0, 0)
        Me.StatusStrip2.Name = "StatusStrip2"
        Me.StatusStrip2.Size = New System.Drawing.Size(768, 25)
        Me.StatusStrip2.SizingGrip = False
        Me.StatusStrip2.TabIndex = 18
        Me.StatusStrip2.Text = "StatusStrip2"
        '
        'ToolStripStatusLabel1
        '
        Me.ToolStripStatusLabel1.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ToolStripStatusLabel1.Name = "ToolStripStatusLabel1"
        Me.ToolStripStatusLabel1.Size = New System.Drawing.Size(119, 20)
        Me.ToolStripStatusLabel1.Text = "STUDENT IMAGES"
        '
        'StatusStrip1
        '
        Me.StatusStrip1.AutoSize = False
        Me.StatusStrip1.BackColor = System.Drawing.Color.LightSeaGreen
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 372)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(768, 25)
        Me.StatusStrip1.SizingGrip = False
        Me.StatusStrip1.TabIndex = 17
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'fdlg
        '
        Me.fdlg.FileName = "Select Photo"
        '
        'frmStudImages
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.LightSkyBlue
        Me.ClientSize = New System.Drawing.Size(795, 423)
        Me.Controls.Add(Me.pnlStudImages)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Name = "frmStudImages"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Student Images"
        Me.pnlStudImages.ResumeLayout(False)
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.Panel2.PerformLayout()
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.ResumeLayout(False)
        CType(Me.pbStudImage, System.ComponentModel.ISupportInitialize).EndInit()
        Me.StatusStrip2.ResumeLayout(False)
        Me.StatusStrip2.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents pnlStudImages As System.Windows.Forms.Panel
    Friend WithEvents StatusStrip2 As System.Windows.Forms.StatusStrip
    Friend WithEvents ToolStripStatusLabel1 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents pbStudImage As System.Windows.Forms.PictureBox
    Friend WithEvents btnSelect As System.Windows.Forms.Button
    Friend WithEvents StatusStrip6 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip7 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip8 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip9 As System.Windows.Forms.StatusStrip
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents btnUpdate As System.Windows.Forms.Button
    Friend WithEvents btnDelete As System.Windows.Forms.Button
    Friend WithEvents btnView As System.Windows.Forms.Button
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents btnClear As System.Windows.Forms.Button
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents cboClassStream As System.Windows.Forms.ComboBox
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents StatusStrip11 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip12 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip13 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusStrip14 As System.Windows.Forms.StatusStrip
    Friend WithEvents cboStudentNo As System.Windows.Forms.ComboBox
    Friend WithEvents cboClassYear As System.Windows.Forms.ComboBox
    Friend WithEvents cboClassName As System.Windows.Forms.ComboBox
    Friend WithEvents txtStudName As System.Windows.Forms.TextBox
    Friend WithEvents fdlg As System.Windows.Forms.OpenFileDialog
    Friend WithEvents txtPicName As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents btnRotate As System.Windows.Forms.Button
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
End Class
