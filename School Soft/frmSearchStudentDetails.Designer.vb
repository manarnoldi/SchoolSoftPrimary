<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmSearchStudentDetails
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
        Me.pnlSearchDetails = New System.Windows.Forms.Panel()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel()
        Me.lstStudDetails = New System.Windows.Forms.ListView()
        Me.ColumnHeader1 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader2 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader3 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader4 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader5 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.txtStudentName = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.txtStudentNo = New System.Windows.Forms.TextBox()
        Me.StatusStrip15 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip16 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip17 = New System.Windows.Forms.StatusStrip()
        Me.ToolStripStatusLabel2 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.StatusStrip18 = New System.Windows.Forms.StatusStrip()
        Me.StatusStrip2 = New System.Windows.Forms.StatusStrip()
        Me.ToolStripStatusLabel1 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip()
        Me.pnlSearchDetails.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.TableLayoutPanel1.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.Panel3.SuspendLayout()
        Me.StatusStrip17.SuspendLayout()
        Me.StatusStrip2.SuspendLayout()
        Me.SuspendLayout()
        '
        'pnlSearchDetails
        '
        Me.pnlSearchDetails.BackColor = System.Drawing.SystemColors.Control
        Me.pnlSearchDetails.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pnlSearchDetails.Controls.Add(Me.Panel1)
        Me.pnlSearchDetails.Controls.Add(Me.StatusStrip1)
        Me.pnlSearchDetails.Location = New System.Drawing.Point(12, 12)
        Me.pnlSearchDetails.Name = "pnlSearchDetails"
        Me.pnlSearchDetails.Size = New System.Drawing.Size(1036, 682)
        Me.pnlSearchDetails.TabIndex = 0
        '
        'Panel1
        '
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel1.Controls.Add(Me.TableLayoutPanel1)
        Me.Panel1.Controls.Add(Me.StatusStrip2)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel1.Location = New System.Drawing.Point(0, 0)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(1034, 649)
        Me.Panel1.TabIndex = 5
        '
        'TableLayoutPanel1
        '
        Me.TableLayoutPanel1.ColumnCount = 1
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.Controls.Add(Me.lstStudDetails, 0, 1)
        Me.TableLayoutPanel1.Controls.Add(Me.Panel2, 0, 0)
        Me.TableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(0, 31)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        Me.TableLayoutPanel1.RowCount = 2
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 16.72078!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 83.27922!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20.0!))
        Me.TableLayoutPanel1.Size = New System.Drawing.Size(1032, 616)
        Me.TableLayoutPanel1.TabIndex = 4
        '
        'lstStudDetails
        '
        Me.lstStudDetails.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColumnHeader1, Me.ColumnHeader2, Me.ColumnHeader3, Me.ColumnHeader4, Me.ColumnHeader5})
        Me.lstStudDetails.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lstStudDetails.FullRowSelect = True
        Me.lstStudDetails.GridLines = True
        Me.lstStudDetails.Location = New System.Drawing.Point(4, 107)
        Me.lstStudDetails.Margin = New System.Windows.Forms.Padding(4)
        Me.lstStudDetails.Name = "lstStudDetails"
        Me.lstStudDetails.Size = New System.Drawing.Size(1024, 505)
        Me.lstStudDetails.TabIndex = 26
        Me.lstStudDetails.UseCompatibleStateImageBehavior = False
        Me.lstStudDetails.View = System.Windows.Forms.View.Details
        '
        'ColumnHeader1
        '
        Me.ColumnHeader1.Text = "Adm Number"
        Me.ColumnHeader1.Width = 100
        '
        'ColumnHeader2
        '
        Me.ColumnHeader2.Text = "Student Name"
        Me.ColumnHeader2.Width = 400
        '
        'ColumnHeader3
        '
        Me.ColumnHeader3.Text = "Class Name"
        Me.ColumnHeader3.Width = 100
        '
        'ColumnHeader4
        '
        Me.ColumnHeader4.Text = "Stream"
        Me.ColumnHeader4.Width = 70
        '
        'ColumnHeader5
        '
        Me.ColumnHeader5.Text = "Year"
        Me.ColumnHeader5.Width = 70
        '
        'Panel2
        '
        Me.Panel2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel2.Controls.Add(Me.Panel3)
        Me.Panel2.Controls.Add(Me.StatusStrip15)
        Me.Panel2.Controls.Add(Me.StatusStrip16)
        Me.Panel2.Controls.Add(Me.StatusStrip17)
        Me.Panel2.Controls.Add(Me.StatusStrip18)
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel2.Location = New System.Drawing.Point(3, 3)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(1026, 97)
        Me.Panel2.TabIndex = 0
        '
        'Panel3
        '
        Me.Panel3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel3.Controls.Add(Me.Label2)
        Me.Panel3.Controls.Add(Me.txtStudentName)
        Me.Panel3.Controls.Add(Me.Label1)
        Me.Panel3.Controls.Add(Me.txtStudentNo)
        Me.Panel3.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel3.Location = New System.Drawing.Point(7, 25)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(1010, 64)
        Me.Panel3.TabIndex = 73
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(412, 26)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(151, 17)
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "Search Student Name:"
        '
        'txtStudentName
        '
        Me.txtStudentName.Location = New System.Drawing.Point(569, 23)
        Me.txtStudentName.Name = "txtStudentName"
        Me.txtStudentName.Size = New System.Drawing.Size(254, 22)
        Me.txtStudentName.TabIndex = 2
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(14, 26)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(111, 17)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Search Adm No:"
        '
        'txtStudentNo
        '
        Me.txtStudentNo.Location = New System.Drawing.Point(131, 23)
        Me.txtStudentNo.Name = "txtStudentNo"
        Me.txtStudentNo.Size = New System.Drawing.Size(236, 22)
        Me.txtStudentNo.TabIndex = 0
        '
        'StatusStrip15
        '
        Me.StatusStrip15.AutoSize = False
        Me.StatusStrip15.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip15.Dock = System.Windows.Forms.DockStyle.Left
        Me.StatusStrip15.ImageScalingSize = New System.Drawing.Size(20, 20)
        Me.StatusStrip15.Location = New System.Drawing.Point(0, 25)
        Me.StatusStrip15.Name = "StatusStrip15"
        Me.StatusStrip15.Padding = New System.Windows.Forms.Padding(1, 4, 1, 27)
        Me.StatusStrip15.Size = New System.Drawing.Size(7, 64)
        Me.StatusStrip15.SizingGrip = False
        Me.StatusStrip15.TabIndex = 72
        Me.StatusStrip15.Text = "StatusStrip15"
        '
        'StatusStrip16
        '
        Me.StatusStrip16.AutoSize = False
        Me.StatusStrip16.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip16.ImageScalingSize = New System.Drawing.Size(20, 20)
        Me.StatusStrip16.Location = New System.Drawing.Point(0, 89)
        Me.StatusStrip16.Name = "StatusStrip16"
        Me.StatusStrip16.Padding = New System.Windows.Forms.Padding(1, 0, 19, 0)
        Me.StatusStrip16.Size = New System.Drawing.Size(1017, 6)
        Me.StatusStrip16.SizingGrip = False
        Me.StatusStrip16.TabIndex = 71
        Me.StatusStrip16.Text = "StatusStrip16"
        '
        'StatusStrip17
        '
        Me.StatusStrip17.AutoSize = False
        Me.StatusStrip17.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip17.Dock = System.Windows.Forms.DockStyle.Top
        Me.StatusStrip17.ImageScalingSize = New System.Drawing.Size(20, 20)
        Me.StatusStrip17.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripStatusLabel2})
        Me.StatusStrip17.Location = New System.Drawing.Point(0, 0)
        Me.StatusStrip17.Name = "StatusStrip17"
        Me.StatusStrip17.Padding = New System.Windows.Forms.Padding(1, 0, 19, 0)
        Me.StatusStrip17.Size = New System.Drawing.Size(1017, 25)
        Me.StatusStrip17.SizingGrip = False
        Me.StatusStrip17.TabIndex = 70
        Me.StatusStrip17.Text = "StatusStrip17"
        '
        'ToolStripStatusLabel2
        '
        Me.ToolStripStatusLabel2.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ToolStripStatusLabel2.Name = "ToolStripStatusLabel2"
        Me.ToolStripStatusLabel2.Size = New System.Drawing.Size(161, 20)
        Me.ToolStripStatusLabel2.Text = "Select Student Details"
        '
        'StatusStrip18
        '
        Me.StatusStrip18.AutoSize = False
        Me.StatusStrip18.BackColor = System.Drawing.Color.MediumAquamarine
        Me.StatusStrip18.Dock = System.Windows.Forms.DockStyle.Right
        Me.StatusStrip18.ImageScalingSize = New System.Drawing.Size(20, 20)
        Me.StatusStrip18.Location = New System.Drawing.Point(1017, 0)
        Me.StatusStrip18.Name = "StatusStrip18"
        Me.StatusStrip18.Padding = New System.Windows.Forms.Padding(1, 4, 1, 27)
        Me.StatusStrip18.Size = New System.Drawing.Size(7, 95)
        Me.StatusStrip18.SizingGrip = False
        Me.StatusStrip18.TabIndex = 69
        Me.StatusStrip18.Text = "StatusStrip18"
        '
        'StatusStrip2
        '
        Me.StatusStrip2.AutoSize = False
        Me.StatusStrip2.BackColor = System.Drawing.Color.LightSeaGreen
        Me.StatusStrip2.Dock = System.Windows.Forms.DockStyle.Top
        Me.StatusStrip2.ImageScalingSize = New System.Drawing.Size(20, 20)
        Me.StatusStrip2.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripStatusLabel1})
        Me.StatusStrip2.Location = New System.Drawing.Point(0, 0)
        Me.StatusStrip2.Name = "StatusStrip2"
        Me.StatusStrip2.Padding = New System.Windows.Forms.Padding(1, 0, 19, 0)
        Me.StatusStrip2.Size = New System.Drawing.Size(1032, 31)
        Me.StatusStrip2.SizingGrip = False
        Me.StatusStrip2.TabIndex = 37
        Me.StatusStrip2.Text = "StatusStrip2"
        '
        'ToolStripStatusLabel1
        '
        Me.ToolStripStatusLabel1.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ToolStripStatusLabel1.Name = "ToolStripStatusLabel1"
        Me.ToolStripStatusLabel1.Size = New System.Drawing.Size(287, 26)
        Me.ToolStripStatusLabel1.Text = "SEARCH STUDENT CLASS DETAILS"
        '
        'StatusStrip1
        '
        Me.StatusStrip1.AutoSize = False
        Me.StatusStrip1.BackColor = System.Drawing.Color.LightSeaGreen
        Me.StatusStrip1.ImageScalingSize = New System.Drawing.Size(20, 20)
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 649)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Padding = New System.Windows.Forms.Padding(1, 0, 19, 0)
        Me.StatusStrip1.Size = New System.Drawing.Size(1034, 31)
        Me.StatusStrip1.SizingGrip = False
        Me.StatusStrip1.TabIndex = 3
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'frmSearchStudentDetails
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.LightSkyBlue
        Me.ClientSize = New System.Drawing.Size(1060, 706)
        Me.Controls.Add(Me.pnlSearchDetails)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Name = "frmSearchStudentDetails"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Student Details"
        Me.pnlSearchDetails.ResumeLayout(False)
        Me.Panel1.ResumeLayout(False)
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.Panel2.ResumeLayout(False)
        Me.Panel3.ResumeLayout(False)
        Me.Panel3.PerformLayout()
        Me.StatusStrip17.ResumeLayout(False)
        Me.StatusStrip17.PerformLayout()
        Me.StatusStrip2.ResumeLayout(False)
        Me.StatusStrip2.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents pnlSearchDetails As Panel
    Friend WithEvents StatusStrip1 As StatusStrip
    Friend WithEvents Panel1 As Panel
    Friend WithEvents TableLayoutPanel1 As TableLayoutPanel
    Friend WithEvents Panel2 As Panel
    Friend WithEvents Panel3 As Panel
    Friend WithEvents StatusStrip15 As StatusStrip
    Friend WithEvents StatusStrip16 As StatusStrip
    Friend WithEvents StatusStrip17 As StatusStrip
    Friend WithEvents ToolStripStatusLabel2 As ToolStripStatusLabel
    Friend WithEvents StatusStrip18 As StatusStrip
    Friend WithEvents lstStudDetails As ListView
    Friend WithEvents ColumnHeader1 As ColumnHeader
    Friend WithEvents ColumnHeader2 As ColumnHeader
    Friend WithEvents ColumnHeader3 As ColumnHeader
    Friend WithEvents ColumnHeader4 As ColumnHeader
    Friend WithEvents ColumnHeader5 As ColumnHeader
    Friend WithEvents Label1 As Label
    Friend WithEvents txtStudentNo As TextBox
    Friend WithEvents Label2 As Label
    Friend WithEvents txtStudentName As TextBox
    Friend WithEvents StatusStrip2 As StatusStrip
    Friend WithEvents ToolStripStatusLabel1 As ToolStripStatusLabel
End Class
