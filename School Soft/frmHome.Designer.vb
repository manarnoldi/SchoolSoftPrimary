<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmHome
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
        Dim TreeNode1 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("School Details")
        Dim TreeNode2 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Term Dates")
        Dim TreeNode3 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Departments")
        Dim TreeNode4 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Teaching Rooms")
        Dim TreeNode5 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Classes")
        Dim TreeNode6 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Staff Details")
        Dim TreeNode7 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Class Heads")
        Dim TreeNode8 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("School Periods")
        Dim TreeNode9 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Max Subject SetUp")
        Dim TreeNode10 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("SCHOOL SETUP", New System.Windows.Forms.TreeNode() {TreeNode1, TreeNode2, TreeNode3, TreeNode4, TreeNode5, TreeNode6, TreeNode7, TreeNode8, TreeNode9})
        Dim TreeNode11 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Class Lists")
        Dim TreeNode12 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Teacher Subject")
        Dim TreeNode13 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("ADMINISTRATION", New System.Windows.Forms.TreeNode() {TreeNode11, TreeNode12})
        Dim TreeNode24 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Payment Modes")
        Dim TreeNode25 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Payment Accounts")
        Dim TreeNode26 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Bank Balances")
        Dim TreeNode27 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Cash Balances")
        Dim TreeNode28 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Mobile Balances")
        Dim TreeNode29 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Adjust Balances")
        Dim TreeNode30 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("SETTINGS", New System.Windows.Forms.TreeNode() {TreeNode24, TreeNode25, TreeNode26, TreeNode27, TreeNode28, TreeNode29})
        Dim TreeNode31 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Fee Category")
        Dim TreeNode32 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("VoteHeads")
        Dim TreeNode33 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Set Student Fee")
        Dim TreeNode34 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Fee Receipts")
        Dim TreeNode35 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Find Receipt")
        Dim TreeNode36 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("SCHOOL FEE", New System.Windows.Forms.TreeNode() {TreeNode31, TreeNode32, TreeNode33, TreeNode34, TreeNode35})
        Dim TreeNode37 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("OTHERS")
        Dim TreeNode38 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("INCOME", New System.Windows.Forms.TreeNode() {TreeNode36, TreeNode37})
        Dim TreeNode39 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Expense Category")
        Dim TreeNode40 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Expense Master")
        Dim TreeNode41 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Payment Request")
        Dim TreeNode42 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Payment Approval")
        Dim TreeNode43 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Payment Voucher")
        Dim TreeNode44 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Payment Reversal")
        Dim TreeNode45 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("EXPENSES", New System.Windows.Forms.TreeNode() {TreeNode39, TreeNode40, TreeNode41, TreeNode42, TreeNode43, TreeNode44})
        Dim TreeNode46 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Account Transfers")
        Dim TreeNode47 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Account Adjustments")
        Dim TreeNode48 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Fee Payments")
        Dim TreeNode49 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Fee Balances")
        Dim TreeNode50 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Fee Statement")
        Dim TreeNode51 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Vote Summary")
        Dim TreeNode52 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Fee Expectation")
        Dim TreeNode53 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Other Income")
        Dim TreeNode54 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Expenses Summary")
        Dim TreeNode55 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Payment Approvals")
        Dim TreeNode56 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Payment Reversals")
        Dim TreeNode57 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Income Expenditure")
        Dim TreeNode58 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("REPORTS", New System.Windows.Forms.TreeNode() {TreeNode46, TreeNode47, TreeNode48, TreeNode49, TreeNode50, TreeNode51, TreeNode52, TreeNode53, TreeNode54, TreeNode55, TreeNode56, TreeNode57})
        Dim TreeNode59 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("REGISTER GRADES")
        Dim TreeNode60 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Register Subjects")
        Dim TreeNode61 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Student Subject")
        Dim TreeNode62 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Teacher Subject")
        Dim TreeNode63 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Exam Names")
        Dim TreeNode64 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Examinations")
        Dim TreeNode65 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("SUBJECTS", New System.Windows.Forms.TreeNode() {TreeNode60, TreeNode61, TreeNode62, TreeNode63, TreeNode64})
        Dim TreeNode66 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Mark Entry Class")
        Dim TreeNode67 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Mark Entry Subject")
        Dim TreeNode68 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Results Viewing")
        Dim TreeNode69 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Results Analysis")
        Dim TreeNode70 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Report Form")
        Dim TreeNode71 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("EXAM RESULTS", New System.Windows.Forms.TreeNode() {TreeNode66, TreeNode67, TreeNode68, TreeNode69, TreeNode70})
        Dim TreeNode14 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Details")
        Dim TreeNode15 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Assign Class")
        Dim TreeNode16 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Images")
        Dim TreeNode17 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Parents")
        Dim TreeNode18 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Former School")
        Dim TreeNode19 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Search Details")
        Dim TreeNode20 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("STUDENT DETAILS", New System.Windows.Forms.TreeNode() {TreeNode14, TreeNode15, TreeNode16, TreeNode17, TreeNode18, TreeNode19})
        Dim TreeNode21 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("PROMOTE STUDENT")
        Dim TreeNode22 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("ACCOMODATION")
        Dim TreeNode23 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("FEES SUMMARY")
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmHome))
        Me.ToolStripLabel1 = New System.Windows.Forms.ToolStripLabel()
        Me.ToolStrip1 = New System.Windows.Forms.ToolStrip()
        Me.ToolStripLabel2 = New System.Windows.Forms.ToolStripLabel()
        Me.ToolStripLabel3 = New System.Windows.Forms.ToolStripLabel()
        Me.MenuStrip = New System.Windows.Forms.MenuStrip()
        Me.LOGOUTToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.SECURITYToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.USERSToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.MODULESToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.DOMAINSToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.DOMAINRIGHTSToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.SYSTEMSETTINGSToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.EDITMYACCOUNTToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem()
        Me.REVERSEFEESToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.WindowManagerPanel1 = New MDIWindowManager.WindowManagerPanel()
        Me.TimerHome = New System.Windows.Forms.Timer(Me.components)
        Me.NaviBand1 = New Guifreaks.NavigationBar.NaviBand(Me.components)
        Me.tvSchool = New System.Windows.Forms.TreeView()
        Me.NaviBarHome = New Guifreaks.NavigationBar.NaviBar(Me.components)
        Me.naviFinance = New Guifreaks.NavigationBar.NaviBand(Me.components)
        Me.tvFinance = New System.Windows.Forms.TreeView()
        Me.naviSchool = New Guifreaks.NavigationBar.NaviBand(Me.components)
        Me.naviAcademics = New Guifreaks.NavigationBar.NaviBand(Me.components)
        Me.tvAcademics = New System.Windows.Forms.TreeView()
        Me.naviStudent = New Guifreaks.NavigationBar.NaviBand(Me.components)
        Me.tvStudent = New System.Windows.Forms.TreeView()
        Me.ToolStrip1.SuspendLayout()
        Me.MenuStrip.SuspendLayout()
        Me.NaviBand1.SuspendLayout()
        CType(Me.NaviBarHome, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.NaviBarHome.SuspendLayout()
        Me.naviFinance.ClientArea.SuspendLayout()
        Me.naviFinance.SuspendLayout()
        Me.naviSchool.ClientArea.SuspendLayout()
        Me.naviSchool.SuspendLayout()
        Me.naviAcademics.ClientArea.SuspendLayout()
        Me.naviAcademics.SuspendLayout()
        Me.naviStudent.ClientArea.SuspendLayout()
        Me.naviStudent.SuspendLayout()
        Me.SuspendLayout()
        '
        'ToolStripLabel1
        '
        Me.ToolStripLabel1.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right
        Me.ToolStripLabel1.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ToolStripLabel1.Name = "ToolStripLabel1"
        Me.ToolStripLabel1.Size = New System.Drawing.Size(159, 28)
        Me.ToolStripLabel1.Text = "ToolStripLabel1"
        '
        'ToolStrip1
        '
        Me.ToolStrip1.BackColor = System.Drawing.Color.LightSeaGreen
        Me.ToolStrip1.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.ToolStrip1.GripStyle = System.Windows.Forms.ToolStripGripStyle.Hidden
        Me.ToolStrip1.ImageScalingSize = New System.Drawing.Size(20, 20)
        Me.ToolStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripLabel1, Me.ToolStripLabel2, Me.ToolStripLabel3})
        Me.ToolStrip1.Location = New System.Drawing.Point(0, 691)
        Me.ToolStrip1.Name = "ToolStrip1"
        Me.ToolStrip1.Size = New System.Drawing.Size(1352, 31)
        Me.ToolStrip1.TabIndex = 25
        Me.ToolStrip1.Text = "ToolStrip1"
        '
        'ToolStripLabel2
        '
        Me.ToolStripLabel2.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ToolStripLabel2.Name = "ToolStripLabel2"
        Me.ToolStripLabel2.Size = New System.Drawing.Size(159, 28)
        Me.ToolStripLabel2.Text = "ToolStripLabel2"
        '
        'ToolStripLabel3
        '
        Me.ToolStripLabel3.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right
        Me.ToolStripLabel3.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ToolStripLabel3.Name = "ToolStripLabel3"
        Me.ToolStripLabel3.Size = New System.Drawing.Size(159, 28)
        Me.ToolStripLabel3.Text = "ToolStripLabel3"
        '
        'MenuStrip
        '
        Me.MenuStrip.BackColor = System.Drawing.Color.LightSeaGreen
        Me.MenuStrip.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MenuStrip.ImageScalingSize = New System.Drawing.Size(20, 20)
        Me.MenuStrip.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.LOGOUTToolStripMenuItem, Me.SECURITYToolStripMenuItem, Me.EDITMYACCOUNTToolStripMenuItem1, Me.REVERSEFEESToolStripMenuItem})
        Me.MenuStrip.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip.Name = "MenuStrip"
        Me.MenuStrip.Padding = New System.Windows.Forms.Padding(8, 2, 0, 2)
        Me.MenuStrip.RenderMode = System.Windows.Forms.ToolStripRenderMode.Professional
        Me.MenuStrip.Size = New System.Drawing.Size(1352, 31)
        Me.MenuStrip.TabIndex = 21
        Me.MenuStrip.Text = "MenuStrip"
        '
        'LOGOUTToolStripMenuItem
        '
        Me.LOGOUTToolStripMenuItem.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LOGOUTToolStripMenuItem.Name = "LOGOUTToolStripMenuItem"
        Me.LOGOUTToolStripMenuItem.Size = New System.Drawing.Size(96, 27)
        Me.LOGOUTToolStripMenuItem.Text = "LOG OUT"
        '
        'SECURITYToolStripMenuItem
        '
        Me.SECURITYToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.USERSToolStripMenuItem, Me.MODULESToolStripMenuItem, Me.DOMAINSToolStripMenuItem, Me.DOMAINRIGHTSToolStripMenuItem, Me.SYSTEMSETTINGSToolStripMenuItem})
        Me.SECURITYToolStripMenuItem.Name = "SECURITYToolStripMenuItem"
        Me.SECURITYToolStripMenuItem.Size = New System.Drawing.Size(100, 27)
        Me.SECURITYToolStripMenuItem.Text = "SECURITY"
        '
        'USERSToolStripMenuItem
        '
        Me.USERSToolStripMenuItem.Name = "USERSToolStripMenuItem"
        Me.USERSToolStripMenuItem.Size = New System.Drawing.Size(235, 28)
        Me.USERSToolStripMenuItem.Text = "USERS"
        '
        'MODULESToolStripMenuItem
        '
        Me.MODULESToolStripMenuItem.Name = "MODULESToolStripMenuItem"
        Me.MODULESToolStripMenuItem.Size = New System.Drawing.Size(235, 28)
        Me.MODULESToolStripMenuItem.Text = "MODULES"
        '
        'DOMAINSToolStripMenuItem
        '
        Me.DOMAINSToolStripMenuItem.Name = "DOMAINSToolStripMenuItem"
        Me.DOMAINSToolStripMenuItem.Size = New System.Drawing.Size(235, 28)
        Me.DOMAINSToolStripMenuItem.Text = "DOMAINS"
        '
        'DOMAINRIGHTSToolStripMenuItem
        '
        Me.DOMAINRIGHTSToolStripMenuItem.Name = "DOMAINRIGHTSToolStripMenuItem"
        Me.DOMAINRIGHTSToolStripMenuItem.Size = New System.Drawing.Size(235, 28)
        Me.DOMAINRIGHTSToolStripMenuItem.Text = "DOMAIN RIGHTS"
        '
        'SYSTEMSETTINGSToolStripMenuItem
        '
        Me.SYSTEMSETTINGSToolStripMenuItem.Name = "SYSTEMSETTINGSToolStripMenuItem"
        Me.SYSTEMSETTINGSToolStripMenuItem.Size = New System.Drawing.Size(235, 28)
        Me.SYSTEMSETTINGSToolStripMenuItem.Text = "SYSTEM SETTINGS"
        '
        'EDITMYACCOUNTToolStripMenuItem1
        '
        Me.EDITMYACCOUNTToolStripMenuItem1.Name = "EDITMYACCOUNTToolStripMenuItem1"
        Me.EDITMYACCOUNTToolStripMenuItem1.Size = New System.Drawing.Size(176, 27)
        Me.EDITMYACCOUNTToolStripMenuItem1.Text = "EDIT MY ACCOUNT"
        '
        'REVERSEFEESToolStripMenuItem
        '
        Me.REVERSEFEESToolStripMenuItem.Name = "REVERSEFEESToolStripMenuItem"
        Me.REVERSEFEESToolStripMenuItem.Size = New System.Drawing.Size(134, 27)
        Me.REVERSEFEESToolStripMenuItem.Text = "REVERSE FEES"
        '
        'WindowManagerPanel1
        '
        Me.WindowManagerPanel1.AllowUserVerticalRepositioning = False
        Me.WindowManagerPanel1.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.WindowManagerPanel1.AutoDetectMdiChildWindows = True
        Me.WindowManagerPanel1.AutoHide = False
        Me.WindowManagerPanel1.BackColor = System.Drawing.Color.LightBlue
        Me.WindowManagerPanel1.ButtonRenderMode = MDIWindowManager.ButtonRenderMode.Standard
        Me.WindowManagerPanel1.DisableCloseAction = False
        Me.WindowManagerPanel1.DisableHTileAction = False
        Me.WindowManagerPanel1.DisablePopoutAction = False
        Me.WindowManagerPanel1.DisableTileAction = False
        Me.WindowManagerPanel1.EnableTabPaintEvent = False
        Me.WindowManagerPanel1.Location = New System.Drawing.Point(362, 33)
        Me.WindowManagerPanel1.Margin = New System.Windows.Forms.Padding(4)
        Me.WindowManagerPanel1.MinMode = False
        Me.WindowManagerPanel1.Name = "WindowManagerPanel1"
        Me.WindowManagerPanel1.Orientation = MDIWindowManager.WindowManagerOrientation.Top
        Me.WindowManagerPanel1.ShowCloseButton = True
        Me.WindowManagerPanel1.ShowIcons = False
        Me.WindowManagerPanel1.ShowLayoutButtons = False
        Me.WindowManagerPanel1.ShowTitle = False
        Me.WindowManagerPanel1.Size = New System.Drawing.Size(988, 36)
        Me.WindowManagerPanel1.Style = MDIWindowManager.TabStyle.AngledHiliteTabs
        Me.WindowManagerPanel1.TabIndex = 26
        Me.WindowManagerPanel1.TabRenderMode = MDIWindowManager.TabsProvider.Standard
        Me.WindowManagerPanel1.TitleBackColor = System.Drawing.SystemColors.ControlDark
        Me.WindowManagerPanel1.TitleForeColor = System.Drawing.SystemColors.ControlLightLight
        '
        'TimerHome
        '
        Me.TimerHome.Enabled = True
        Me.TimerHome.Interval = 1000
        '
        'NaviBand1
        '
        '
        'NaviBand1.ClientArea
        '
        Me.NaviBand1.ClientArea.Location = New System.Drawing.Point(0, 0)
        Me.NaviBand1.ClientArea.Name = "ClientArea"
        Me.NaviBand1.ClientArea.Size = New System.Drawing.Size(281, 374)
        Me.NaviBand1.ClientArea.TabIndex = 0
        Me.NaviBand1.Location = New System.Drawing.Point(1, 27)
        Me.NaviBand1.Name = "NaviBand1"
        Me.NaviBand1.Size = New System.Drawing.Size(281, 374)
        Me.NaviBand1.TabIndex = 17
        '
        'tvSchool
        '
        Me.tvSchool.BackColor = System.Drawing.Color.Azure
        Me.tvSchool.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.tvSchool.Dock = System.Windows.Forms.DockStyle.Fill
        Me.tvSchool.Font = New System.Drawing.Font("Century Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tvSchool.Indent = 35
        Me.tvSchool.ItemHeight = 30
        Me.tvSchool.Location = New System.Drawing.Point(0, 0)
        Me.tvSchool.Margin = New System.Windows.Forms.Padding(4)
        Me.tvSchool.Name = "tvSchool"
        TreeNode1.Name = "nodeSchDetails"
        TreeNode1.Text = "School Details"
        TreeNode2.Name = "nodeTermDates"
        TreeNode2.Text = "Term Dates"
        TreeNode3.Name = "nodeDepartments"
        TreeNode3.Text = "Departments"
        TreeNode4.Name = "nodeTeachingRooms"
        TreeNode4.Text = "Teaching Rooms"
        TreeNode5.Name = "nodeClasses"
        TreeNode5.Text = "Classes"
        TreeNode6.Name = "nodeStaffDetails"
        TreeNode6.Text = "Staff Details"
        TreeNode7.Name = "nodeClassHeads"
        TreeNode7.Text = "Class Heads"
        TreeNode8.Name = "nodeSchoolPeriods"
        TreeNode8.Text = "School Periods"
        TreeNode9.Name = "nodeSubMaxSetUp"
        TreeNode9.Text = "Max Subject SetUp"
        TreeNode10.Name = "nodeSchSetUp"
        TreeNode10.Text = "SCHOOL SETUP"
        TreeNode11.Name = "nodeClassLists"
        TreeNode11.Text = "Class Lists"
        TreeNode12.Name = "nodeTeacherSubject"
        TreeNode12.Text = "Teacher Subject"
        TreeNode13.Name = "nodeAdministration"
        TreeNode13.Text = "ADMINISTRATION"
        Me.tvSchool.Nodes.AddRange(New System.Windows.Forms.TreeNode() {TreeNode10, TreeNode13})
        Me.tvSchool.Size = New System.Drawing.Size(358, 465)
        Me.tvSchool.TabIndex = 0
        '
        'NaviBarHome
        '
        Me.NaviBarHome.ActiveBand = Me.naviStudent
        Me.NaviBarHome.BackColor = System.Drawing.Color.Azure
        Me.NaviBarHome.Controls.Add(Me.naviStudent)
        Me.NaviBarHome.Controls.Add(Me.naviSchool)
        Me.NaviBarHome.Controls.Add(Me.naviFinance)
        Me.NaviBarHome.Controls.Add(Me.naviAcademics)
        Me.NaviBarHome.Dock = System.Windows.Forms.DockStyle.Left
        Me.NaviBarHome.Font = New System.Drawing.Font("Century Gothic", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.NaviBarHome.Location = New System.Drawing.Point(0, 31)
        Me.NaviBarHome.Margin = New System.Windows.Forms.Padding(4)
        Me.NaviBarHome.Name = "NaviBarHome"
        Me.NaviBarHome.ShowMoreOptionsButton = False
        Me.NaviBarHome.Size = New System.Drawing.Size(360, 660)
        Me.NaviBarHome.TabIndex = 30
        Me.NaviBarHome.Text = "NaviBarHome"
        Me.NaviBarHome.VisibleLargeButtons = 4
        '
        'naviFinance
        '
        '
        'naviFinance.ClientArea
        '
        Me.naviFinance.ClientArea.Controls.Add(Me.tvFinance)
        Me.naviFinance.ClientArea.Location = New System.Drawing.Point(0, 0)
        Me.naviFinance.ClientArea.Margin = New System.Windows.Forms.Padding(4)
        Me.naviFinance.ClientArea.Name = "ClientArea"
        Me.naviFinance.ClientArea.Size = New System.Drawing.Size(358, 465)
        Me.naviFinance.ClientArea.TabIndex = 0
        Me.naviFinance.Location = New System.Drawing.Point(1, 27)
        Me.naviFinance.Margin = New System.Windows.Forms.Padding(4)
        Me.naviFinance.Name = "naviFinance"
        Me.naviFinance.Size = New System.Drawing.Size(358, 465)
        Me.naviFinance.TabIndex = 5
        Me.naviFinance.Text = "FINANCE"
        '
        'tvFinance
        '
        Me.tvFinance.BackColor = System.Drawing.Color.Azure
        Me.tvFinance.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.tvFinance.Dock = System.Windows.Forms.DockStyle.Fill
        Me.tvFinance.Font = New System.Drawing.Font("Century Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tvFinance.Indent = 35
        Me.tvFinance.ItemHeight = 30
        Me.tvFinance.Location = New System.Drawing.Point(0, 0)
        Me.tvFinance.Margin = New System.Windows.Forms.Padding(4)
        Me.tvFinance.Name = "tvFinance"
        TreeNode24.Name = "nodePaymentModes"
        TreeNode24.Text = "Payment Modes"
        TreeNode25.Name = "nodePayAccounts"
        TreeNode25.Text = "Payment Accounts"
        TreeNode26.Name = "nodeBankBal"
        TreeNode26.Text = "Bank Balances"
        TreeNode27.Name = "nodeCashBal"
        TreeNode27.Text = "Cash Balances"
        TreeNode28.Name = "nodeMobileBal"
        TreeNode28.Text = "Mobile Balances"
        TreeNode29.Name = "nodeAdjustBal"
        TreeNode29.Text = "Adjust Balances"
        TreeNode30.Name = "nodeFinSetting"
        TreeNode30.Text = "SETTINGS"
        TreeNode31.Name = "nodeFeeCat"
        TreeNode31.Text = "Fee Category"
        TreeNode32.Name = "nodeVoteHeads"
        TreeNode32.Text = "VoteHeads"
        TreeNode33.Name = "nodeStudFee"
        TreeNode33.Text = "Set Student Fee"
        TreeNode34.Name = "nodeFeeReceipts"
        TreeNode34.Text = "Fee Receipts"
        TreeNode35.Name = "nodeFindReceipt"
        TreeNode35.Text = "Find Receipt"
        TreeNode36.Name = "nodeSchFee"
        TreeNode36.Text = "SCHOOL FEE"
        TreeNode37.Name = "nodeOthers"
        TreeNode37.Text = "OTHERS"
        TreeNode38.Name = "nodeIncome"
        TreeNode38.Text = "INCOME"
        TreeNode39.Name = "nodeExpCat"
        TreeNode39.Text = "Expense Category"
        TreeNode40.Name = "nodeExpMaster"
        TreeNode40.Text = "Expense Master"
        TreeNode41.Name = "nodePayRequest"
        TreeNode41.Text = "Payment Request"
        TreeNode42.Name = "nodePayApproval"
        TreeNode42.Text = "Payment Approval"
        TreeNode43.Name = "nodePaymentVoucher"
        TreeNode43.Text = "Payment Voucher"
        TreeNode44.Name = "nodePayReversal"
        TreeNode44.Text = "Payment Reversal"
        TreeNode45.Name = "nodeExpensesMain"
        TreeNode45.Text = "EXPENSES"
        TreeNode46.Name = "nodeAccountTransfers"
        TreeNode46.Text = "Account Transfers"
        TreeNode47.Name = "nodeAccountAdj"
        TreeNode47.Text = "Account Adjustments"
        TreeNode48.Name = "nodeFeePayments"
        TreeNode48.Text = "Fee Payments"
        TreeNode49.Name = "nodeFeeBalances"
        TreeNode49.Text = "Fee Balances"
        TreeNode50.Name = "nodeFeeStatement"
        TreeNode50.Text = "Fee Statement"
        TreeNode51.Name = "nodeVoteSummary"
        TreeNode51.Text = "Vote Summary"
        TreeNode52.Name = "nodeFeeExpectation"
        TreeNode52.Text = "Fee Expectation"
        TreeNode53.Name = "nodeOtherIncome"
        TreeNode53.Text = "Other Income"
        TreeNode54.Name = "nodeExpenses"
        TreeNode54.Text = "Expenses Summary"
        TreeNode55.Name = "nodePayApprovals"
        TreeNode55.Text = "Payment Approvals"
        TreeNode56.Name = "nodePayReversals"
        TreeNode56.Text = "Payment Reversals"
        TreeNode57.Name = "nodeIncomeExp"
        TreeNode57.Text = "Income Expenditure"
        TreeNode58.Name = "nodeIncReports"
        TreeNode58.Text = "REPORTS"
        Me.tvFinance.Nodes.AddRange(New System.Windows.Forms.TreeNode() {TreeNode30, TreeNode38, TreeNode45, TreeNode58})
        Me.tvFinance.Size = New System.Drawing.Size(358, 465)
        Me.tvFinance.TabIndex = 4
        '
        'naviSchool
        '
        '
        'naviSchool.ClientArea
        '
        Me.naviSchool.ClientArea.Controls.Add(Me.tvSchool)
        Me.naviSchool.ClientArea.Location = New System.Drawing.Point(0, 0)
        Me.naviSchool.ClientArea.Margin = New System.Windows.Forms.Padding(4)
        Me.naviSchool.ClientArea.Name = "ClientArea"
        Me.naviSchool.ClientArea.Size = New System.Drawing.Size(358, 465)
        Me.naviSchool.ClientArea.TabIndex = 0
        Me.naviSchool.Location = New System.Drawing.Point(1, 27)
        Me.naviSchool.Margin = New System.Windows.Forms.Padding(4)
        Me.naviSchool.Name = "naviSchool"
        Me.naviSchool.Size = New System.Drawing.Size(358, 465)
        Me.naviSchool.TabIndex = 1
        Me.naviSchool.Text = "SCHOOL"
        '
        'naviAcademics
        '
        '
        'naviAcademics.ClientArea
        '
        Me.naviAcademics.ClientArea.Controls.Add(Me.tvAcademics)
        Me.naviAcademics.ClientArea.Location = New System.Drawing.Point(0, 0)
        Me.naviAcademics.ClientArea.Margin = New System.Windows.Forms.Padding(4)
        Me.naviAcademics.ClientArea.Name = "ClientArea"
        Me.naviAcademics.ClientArea.Size = New System.Drawing.Size(358, 465)
        Me.naviAcademics.ClientArea.TabIndex = 0
        Me.naviAcademics.Location = New System.Drawing.Point(1, 27)
        Me.naviAcademics.Margin = New System.Windows.Forms.Padding(4)
        Me.naviAcademics.Name = "naviAcademics"
        Me.naviAcademics.Size = New System.Drawing.Size(358, 465)
        Me.naviAcademics.TabIndex = 3
        Me.naviAcademics.Text = "ACADEMICS"
        '
        'tvAcademics
        '
        Me.tvAcademics.BackColor = System.Drawing.Color.Azure
        Me.tvAcademics.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.tvAcademics.Dock = System.Windows.Forms.DockStyle.Fill
        Me.tvAcademics.Font = New System.Drawing.Font("Century Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tvAcademics.Indent = 35
        Me.tvAcademics.ItemHeight = 30
        Me.tvAcademics.Location = New System.Drawing.Point(0, 0)
        Me.tvAcademics.Margin = New System.Windows.Forms.Padding(4)
        Me.tvAcademics.Name = "tvAcademics"
        TreeNode59.Name = "nodeAcadGrades"
        TreeNode59.Text = "REGISTER GRADES"
        TreeNode60.Name = "nodeAcadSubjectsReg"
        TreeNode60.Text = "Register Subjects"
        TreeNode61.Name = "nodeAcadSubjectsStudent"
        TreeNode61.Text = "Student Subject"
        TreeNode62.Name = "nodeAcadSubjectsTeacher"
        TreeNode62.Text = "Teacher Subject"
        TreeNode63.Name = "nodeExamNames"
        TreeNode63.Text = "Exam Names"
        TreeNode64.Name = "nodeAcadExams"
        TreeNode64.Text = "Examinations"
        TreeNode65.Name = "nodeAcadSubjects"
        TreeNode65.Text = "SUBJECTS"
        TreeNode66.Name = "nodeAcadMarkEntryClass"
        TreeNode66.Text = "Mark Entry Class"
        TreeNode67.Name = "nodeAcadMarkEntrySubj"
        TreeNode67.Text = "Mark Entry Subject"
        TreeNode68.Name = "nodeResultsViewing"
        TreeNode68.Text = "Results Viewing"
        TreeNode69.Name = "nodeAcadResultAnalysis"
        TreeNode69.Text = "Results Analysis"
        TreeNode70.Name = "nodeAcadReportForm"
        TreeNode70.Text = "Report Form"
        TreeNode71.Name = "nodeAcadExamResults"
        TreeNode71.Text = "EXAM RESULTS"
        Me.tvAcademics.Nodes.AddRange(New System.Windows.Forms.TreeNode() {TreeNode59, TreeNode65, TreeNode71})
        Me.tvAcademics.Size = New System.Drawing.Size(358, 465)
        Me.tvAcademics.TabIndex = 1
        '
        'naviStudent
        '
        '
        'naviStudent.ClientArea
        '
        Me.naviStudent.ClientArea.Controls.Add(Me.tvStudent)
        Me.naviStudent.ClientArea.Location = New System.Drawing.Point(0, 0)
        Me.naviStudent.ClientArea.Margin = New System.Windows.Forms.Padding(4)
        Me.naviStudent.ClientArea.Name = "ClientArea"
        Me.naviStudent.ClientArea.Size = New System.Drawing.Size(358, 465)
        Me.naviStudent.ClientArea.TabIndex = 0
        Me.naviStudent.Location = New System.Drawing.Point(1, 27)
        Me.naviStudent.Margin = New System.Windows.Forms.Padding(4)
        Me.naviStudent.Name = "naviStudent"
        Me.naviStudent.Size = New System.Drawing.Size(358, 465)
        Me.naviStudent.TabIndex = 2
        Me.naviStudent.Text = "STUDENT"
        '
        'tvStudent
        '
        Me.tvStudent.BackColor = System.Drawing.Color.Azure
        Me.tvStudent.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.tvStudent.Dock = System.Windows.Forms.DockStyle.Fill
        Me.tvStudent.Font = New System.Drawing.Font("Century Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tvStudent.Indent = 35
        Me.tvStudent.ItemHeight = 30
        Me.tvStudent.Location = New System.Drawing.Point(0, 0)
        Me.tvStudent.Margin = New System.Windows.Forms.Padding(4)
        Me.tvStudent.Name = "tvStudent"
        TreeNode14.Name = "nodeStudDetails"
        TreeNode14.Text = "Details"
        TreeNode15.Name = "nodeAssigStudClass"
        TreeNode15.Text = "Assign Class"
        TreeNode16.Name = "nodeStudImages"
        TreeNode16.Text = "Images"
        TreeNode17.Name = "nodeStudParents"
        TreeNode17.Text = "Parents"
        TreeNode18.Name = "nodeFormerSchool"
        TreeNode18.Text = "Former School"
        TreeNode19.Name = "nodeSearchDetails"
        TreeNode19.Text = "Search Details"
        TreeNode20.Name = "Node0"
        TreeNode20.Text = "STUDENT DETAILS"
        TreeNode21.Name = "nodePromoteStudent"
        TreeNode21.Text = "PROMOTE STUDENT"
        TreeNode22.Name = "nodeStudAccomodation"
        TreeNode22.Text = "ACCOMODATION"
        TreeNode23.Name = "nodeStudFees"
        TreeNode23.Text = "FEES SUMMARY"
        Me.tvStudent.Nodes.AddRange(New System.Windows.Forms.TreeNode() {TreeNode20, TreeNode21, TreeNode22, TreeNode23})
        Me.tvStudent.Size = New System.Drawing.Size(358, 465)
        Me.tvStudent.TabIndex = 3
        '
        'frmHome
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Center
        Me.ClientSize = New System.Drawing.Size(1352, 722)
        Me.Controls.Add(Me.NaviBarHome)
        Me.Controls.Add(Me.WindowManagerPanel1)
        Me.Controls.Add(Me.ToolStrip1)
        Me.Controls.Add(Me.MenuStrip)
        Me.DoubleBuffered = True
        Me.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.IsMdiContainer = True
        Me.MainMenuStrip = Me.MenuStrip
        Me.Margin = New System.Windows.Forms.Padding(4)
        Me.Name = "frmHome"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "SCHOOL SOFT"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.ToolStrip1.ResumeLayout(False)
        Me.ToolStrip1.PerformLayout()
        Me.MenuStrip.ResumeLayout(False)
        Me.MenuStrip.PerformLayout()
        Me.NaviBand1.ResumeLayout(False)
        CType(Me.NaviBarHome, System.ComponentModel.ISupportInitialize).EndInit()
        Me.NaviBarHome.ResumeLayout(False)
        Me.naviFinance.ClientArea.ResumeLayout(False)
        Me.naviFinance.ResumeLayout(False)
        Me.naviSchool.ClientArea.ResumeLayout(False)
        Me.naviSchool.ResumeLayout(False)
        Me.naviAcademics.ClientArea.ResumeLayout(False)
        Me.naviAcademics.ResumeLayout(False)
        Me.naviStudent.ClientArea.ResumeLayout(False)
        Me.naviStudent.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents ToolStripLabel1 As System.Windows.Forms.ToolStripLabel
    Friend WithEvents ToolStrip1 As System.Windows.Forms.ToolStrip
    Friend WithEvents ToolStripLabel2 As System.Windows.Forms.ToolStripLabel
    Friend WithEvents MenuStrip As System.Windows.Forms.MenuStrip
    Friend WithEvents WindowManagerPanel1 As MDIWindowManager.WindowManagerPanel
    Friend WithEvents ToolStripLabel3 As System.Windows.Forms.ToolStripLabel
    Friend WithEvents TimerHome As System.Windows.Forms.Timer
    Friend WithEvents LOGOUTToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents SECURITYToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents USERSToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents MODULESToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents DOMAINSToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents DOMAINRIGHTSToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents NaviBand1 As Guifreaks.NavigationBar.NaviBand
    Friend WithEvents tvSchool As System.Windows.Forms.TreeView
    Friend WithEvents NaviBarHome As Guifreaks.NavigationBar.NaviBar
    Friend WithEvents naviStudent As Guifreaks.NavigationBar.NaviBand
    Friend WithEvents tvStudent As System.Windows.Forms.TreeView
    Friend WithEvents naviAcademics As Guifreaks.NavigationBar.NaviBand
    Friend WithEvents tvAcademics As System.Windows.Forms.TreeView
    Friend WithEvents naviSchool As Guifreaks.NavigationBar.NaviBand
    Friend WithEvents EDITMYACCOUNTToolStripMenuItem1 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents SYSTEMSETTINGSToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents naviFinance As Guifreaks.NavigationBar.NaviBand
    Friend WithEvents tvFinance As System.Windows.Forms.TreeView
    Friend WithEvents REVERSEFEESToolStripMenuItem As ToolStripMenuItem
End Class
