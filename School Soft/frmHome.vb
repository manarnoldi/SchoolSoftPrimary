Imports System.Windows.Forms
Imports CrystalDecisions.Shared
Imports CrystalDecisions.CrystalReports.Engine
Imports System.Data.SqlClient
Public Class frmHome
    Dim rightAccess As Boolean = False
    Dim reader As SqlDataReader
    Dim reader1 As SqlDataReader
    Dim cmdHome As New SqlCommand
    Dim cmdHome1 As New SqlCommand
    Private Sub frmHome_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        dbconnection()
        Me.NaviBarHome.ActiveBand = Nothing
        HideTreeViews()
        HideFeeReverseModule()
    End Sub

    Private Sub TimerHome_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TimerHome.Tick
        Me.ToolStripLabel1.Text = Space(10) & " TIME: " & TimeOfDay()
    End Sub

    Private Sub LOGOUTToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LOGOUTToolStripMenuItem.Click
        Dim result As MsgBoxResult = MsgBox("Are you sure you want to log out?", MsgBoxStyle.ApplicationModal + MsgBoxStyle.Question + MsgBoxStyle.YesNo, "Confirm Transaction")
        If result = MsgBoxResult.No Then
            Exit Sub
        End If
        frmLogin.Show()
        Me.Close()
    End Sub

    Private Sub USERSToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles USERSToolStripMenuItem.Click
        frmUsers.MdiParent = Me
        frmUsers.Show()
    End Sub

    Private Sub MODULESToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MODULESToolStripMenuItem.Click
        frmModules.MdiParent = Me
        frmModules.Show()
    End Sub

    Private Sub DOMAINSToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DOMAINSToolStripMenuItem.Click
        frmDomains.MdiParent = Me
        frmDomains.Show()
    End Sub

    Private Sub DOMAINRIGHTSToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DOMAINRIGHTSToolStripMenuItem.Click
        frmDomainRights.MdiParent = Me
        frmDomainRights.Show()
    End Sub

    Private Sub HideFeeReverseModule()
        Me.REVERSEFEESToolStripMenuItem.Visible = False

        If userName.Trim = "Arnold" Then
            Me.REVERSEFEESToolStripMenuItem.Visible = True
        End If

    End Sub

    Private Sub HideTreeViews()
        Me.tvSchool.Visible = False
        Me.tvStudent.Visible = False
        Me.tvAcademics.Visible = False
        'Me.tvTimeTable.Visible = False
        Me.tvFinance.Visible = False
        ' Me.tvInventory.Visible = False
        ' Me.tvAccommodation.Visible = False
        Me.NaviBarHome.VisibleLargeButtons = 7
    End Sub
    Public Sub NaviBarHome_ActiveBandChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles NaviBarHome.ActiveBandChanged
        Try
            activeband = Me.NaviBarHome.ActiveBand.Text.Trim
            HideTreeViews()
            dbconnection()
            cmdHome.Connection = conn
            cmdHome.CommandType = CommandType.Text
            cmdHome.CommandText = "SELECT * FROM vwUserRights WHERE (userName=@userName) And (domainName=@domainName)" &
                vbNewLine & " And (modName=@modName) And (userStatus='True') AND (domainStatus='True') AND " &
                vbNewLine & " (staffStatus='True') AND (modStatus='True') AND (rightAccess='True')"
            cmdHome.Parameters.Clear()
            cmdHome.Parameters.AddWithValue("@userName", userName.Trim)
            cmdHome.Parameters.AddWithValue("@domainName", domainName.Trim)
            cmdHome.Parameters.AddWithValue("@modName", activeband)
            reader = cmdHome.ExecuteReader
            If reader.HasRows Then
                rightAccess = True
            Else
                rightAccess = False
            End If

            If rightAccess = True Then

                Select Case activeband
                    Case "SCHOOL"
                        Me.tvSchool.Visible = True
                        Me.tvSchool.ExpandAll()
                    Case "STUDENT"
                        Me.tvStudent.Visible = True
                        Me.tvStudent.ExpandAll()
                    Case "ACADEMICS"
                        Me.tvAcademics.Visible = True
                        Me.tvAcademics.ExpandAll()
                    'Case "TIME TABLE"
                    '    Me.tvTimeTable.Visible = True
                    '    Me.tvTimeTable.ExpandAll()
                    Case "FINANCE"
                        Me.tvFinance.Visible = True
                        Me.tvFinance.ExpandAll()
                        'Case "INVENTORY"
                        '    Me.tvInventory.Visible = True
                        '    Me.tvInventory.ExpandAll()
                        'Case "ACCOMODATION"
                        '    Me.tvAccommodation.Visible = True
                        '    Me.tvAccommodation.ExpandAll()
                End Select
            End If
            activeband = Nothing
        Catch ex As Exception

        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub

    Private Sub EDITMYACCOUNTToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EDITMYACCOUNTToolStripMenuItem1.Click
        frmEditMyAccount.MdiParent = Me
        frmEditMyAccount.Show()
    End Sub

    Private Sub SECURITYToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SECURITYToolStripMenuItem.Click
        Try
            activeband = "SECURITY"
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            cmdHome.Connection = conn
            cmdHome.CommandType = CommandType.Text
            cmdHome.CommandText = "SELECT * FROM vwUserRights WHERE (userName=@userName) AND (domainName=@domainName)" &
                vbNewLine & " AND (modName='SECURITY') AND (userStatus='True') AND (domainStatus='True') AND " &
                vbNewLine & " (staffStatus='True') AND (modStatus='True') AND (rightAccess='True')"
            cmdHome.Parameters.Clear()
            cmdHome.Parameters.AddWithValue("@userName", userName.Trim)
            cmdHome.Parameters.AddWithValue("@domainName", domainName.Trim)
            reader = cmdHome.ExecuteReader
            If reader.HasRows Then
                rightAccess = True
            Else
                rightAccess = False
            End If
            If rightAccess = False Then
                MsgBox("You have no rights to access security modules", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
                Me.SECURITYToolStripMenuItem.Enabled = False
                Me.USERSToolStripMenuItem.Enabled = False
                Me.MODULESToolStripMenuItem.Enabled = False
                Me.DOMAINSToolStripMenuItem.Enabled = False
                Me.DOMAINRIGHTSToolStripMenuItem.Enabled = False
            Else
                Me.SECURITYToolStripMenuItem.Enabled = True
                Me.USERSToolStripMenuItem.Enabled = True
                Me.MODULESToolStripMenuItem.Enabled = True
                Me.DOMAINSToolStripMenuItem.Enabled = True
                Me.DOMAINRIGHTSToolStripMenuItem.Enabled = True
            End If
            reader.Close()
        Catch ex As Exception
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub

    Private Sub tvSchool_NodeMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.TreeNodeMouseClickEventArgs) Handles tvSchool.NodeMouseClick
        Try
            rightAccess = False
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            cmdHome.Connection = conn
            cmdHome.CommandType = CommandType.Text
            cmdHome.CommandText = "SELECT * FROM tblModules WHERE (modName=@modName) AND (status='True')"
            cmdHome.Parameters.Clear()
            cmdHome.Parameters.AddWithValue("@modName", e.Node.Text.Trim)
            reader = cmdHome.ExecuteReader
            If reader.HasRows Then
                Dim conn2 As New SqlConnection("Server=" & My.Settings.serverName.Trim & ";User ID=" & My.Settings.UserName.Trim & ";Database=" & My.Settings.dbName.Trim & ";Password=" & My.Settings.PassWord.Trim & "")
                conn2.Open()
                cmdHome1.CommandText = "SELECT * FROM vwUserRights WHERE (userName=@userName) AND (domainName=@domainName)" &
                vbNewLine & " AND (modName=@modName) AND (userStatus='True') AND (domainStatus='True') AND " &
                vbNewLine & " (staffStatus='True') AND (modStatus='True') AND (rightAccess='True')"
                cmdHome1.CommandType = CommandType.Text
                cmdHome1.Connection = conn2
                cmdHome1.Parameters.Clear()
                cmdHome1.Parameters.AddWithValue("@userName", userName)
                cmdHome1.Parameters.AddWithValue("@domainName", domainName)
                cmdHome1.Parameters.AddWithValue("@modName", e.Node.Text.Trim)
                reader1 = cmdHome1.ExecuteReader()
                If reader1.HasRows Then
                    rightAccess = True
                Else
                    rightAccess = False
                End If
                conn2.Close()
                reader1.Close()
            Else
                rightAccess = True
            End If
            reader.Close()
            If rightAccess = True Then
                Select Case e.Node.Name.Trim
                    Case "nodeSchDetails"
                        frmSchoolDetails.MdiParent = Me
                        frmSchoolDetails.Show()
                    Case "nodeTermDates"
                        frmSchTermDates.MdiParent = Me
                        frmSchTermDates.Show()
                    Case "nodeDepartments"
                        frmSchDepartments.MdiParent = Me
                        frmSchDepartments.Show()
                    Case "nodeTeachingRooms"
                        frmSchTeachingRooms.MdiParent = Me
                        frmSchTeachingRooms.Show()
                    Case "nodeClasses"
                        frmSchClasses.MdiParent = Me
                        frmSchClasses.Show()
                    Case "nodeStaffDetails"
                        frmSchoolStaff.MdiParent = Me
                        frmSchoolStaff.Show()
                    Case "nodeClassHeads"
                        frmClassHeads.MdiParent = Me
                        frmClassHeads.Show()
                    Case "nodeSchoolPeriods"
                        frmSchoolPeriods.MdiParent = Me
                        frmSchoolPeriods.Show()
                    Case "nodeSubMaxSetUp"
                        frmSchMaxSubjectSetUp.MdiParent = Me
                        frmSchMaxSubjectSetUp.Show()
                    Case "nodeTeacherSubject"
                        frmRptstaffsubj.MdiParent = Me
                        frmRptstaffsubj.Show()
                    Case "nodeClassLists"
                        frmRptClassLists.MdiParent = Me
                        frmRptClassLists.Show()
                End Select
            Else
                MsgBox("You have no rights to access this module!", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            End If
            rightAccess = False
        Catch ex As Exception
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub

    Private Sub tvStudent_NodeMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.TreeNodeMouseClickEventArgs) Handles tvStudent.NodeMouseClick
        Try
            rightAccess = False
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            cmdHome.Connection = conn
            cmdHome.CommandType = CommandType.Text
            cmdHome.CommandText = "SELECT * FROM tblModules WHERE (modName=@modName) AND (status='True')"
            cmdHome.Parameters.Clear()
            cmdHome.Parameters.AddWithValue("@modName", e.Node.Text.Trim)
            reader = cmdHome.ExecuteReader
            If reader.HasRows Then
                Dim conn2 As New SqlConnection("Server=" & My.Settings.serverName.Trim & ";User ID=" & My.Settings.UserName.Trim & ";Database=" & My.Settings.dbName.Trim & ";Password=" & My.Settings.PassWord.Trim & "")
                conn2.Open()
                cmdHome1.CommandText = "SELECT * FROM vwUserRights WHERE (userName=@userName) AND (domainName=@domainName)" &
                vbNewLine & " AND (modName=@modName) AND (userStatus='True') AND (domainStatus='True') AND " &
                vbNewLine & " (staffStatus='True') AND (modStatus='True') AND (rightAccess='True')"
                cmdHome1.CommandType = CommandType.Text
                cmdHome1.Connection = conn2
                cmdHome1.Parameters.Clear()
                cmdHome1.Parameters.AddWithValue("@userName", userName)
                cmdHome1.Parameters.AddWithValue("@domainName", domainName)
                cmdHome1.Parameters.AddWithValue("@modName", e.Node.Text.Trim)
                reader1 = cmdHome1.ExecuteReader()
                If reader1.HasRows Then
                    rightAccess = True
                Else
                    rightAccess = False
                End If
                conn2.Close()
                reader1.Close()
            Else
                rightAccess = True
            End If
            reader.Close()
            If rightAccess = True Then
                Select Case e.Node.Name.Trim
                    Case "nodeStudDetails"
                        frmStudentDetails.MdiParent = Me
                        frmStudentDetails.Show()
                    Case "nodeAssigStudClass"
                        frmStudClass.MdiParent = Me
                        frmStudClass.Show()
                    Case "nodeStudImages"
                        frmStudImages.MdiParent = Me
                        frmStudImages.Show()
                    Case "nodeStudParents"
                        frmStudParents.MdiParent = Me
                        frmStudParents.Show()
                    Case "nodeFormerSchool"
                        frmStudFormSchDetails.MdiParent = Me
                        frmStudFormSchDetails.Show()
                    Case "nodePromoteStudent"
                        frmPromoteStudent.MdiParent = Me
                        frmPromoteStudent.Show()
                    Case "nodeStudAccomodation"
                        frmStudAccomodation.MdiParent = Me
                        frmStudAccomodation.Show()
                    Case "nodeStudFees"
                        frmStudentFeeSummary.MdiParent = Me
                        frmStudentFeeSummary.Show()
                    Case "nodeSearchDetails"
                        frmSearchStudentDetails.MdiParent = Me
                        frmSearchStudentDetails.Show()
                End Select
            Else
                MsgBox("You have no rights to access this module!", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            End If
            rightAccess = False
        Catch ex As Exception
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub

    Private Sub SYSTEMSETTINGSToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SYSTEMSETTINGSToolStripMenuItem.Click
        frmSettings.ShowDialog()
        Me.Close()
    End Sub

    Private Sub tvAcademics_NodeMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.TreeNodeMouseClickEventArgs) Handles tvAcademics.NodeMouseClick
        Try
            rightAccess = False
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            cmdHome.Connection = conn
            cmdHome.CommandType = CommandType.Text
            cmdHome.CommandText = "SELECT * FROM tblModules WHERE (modName=@modName) AND (status='True')"
            cmdHome.Parameters.Clear()
            cmdHome.Parameters.AddWithValue("@modName", e.Node.Text.Trim)
            reader = cmdHome.ExecuteReader
            If reader.HasRows Then
                Dim conn2 As New SqlConnection("Server=" & My.Settings.serverName.Trim & ";User ID=" & My.Settings.UserName.Trim & ";Database=" & My.Settings.dbName.Trim & ";Password=" & My.Settings.PassWord.Trim & "")
                conn2.Open()
                cmdHome1.CommandText = "SELECT * FROM vwUserRights WHERE (userName=@userName) AND (domainName=@domainName)" &
                vbNewLine & " AND (modName=@modName) AND (userStatus='True') AND (domainStatus='True') AND " &
                vbNewLine & " (staffStatus='True') AND (modStatus='True') AND (rightAccess='True')"
                cmdHome1.CommandType = CommandType.Text
                cmdHome1.Connection = conn2
                cmdHome1.Parameters.Clear()
                cmdHome1.Parameters.AddWithValue("@userName", userName)
                cmdHome1.Parameters.AddWithValue("@domainName", domainName)
                cmdHome1.Parameters.AddWithValue("@modName", e.Node.Text.Trim)
                reader1 = cmdHome1.ExecuteReader()
                If reader1.HasRows Then
                    rightAccess = True
                Else
                    rightAccess = False
                End If
                conn2.Close()
                reader1.Close()
            Else
                rightAccess = True
            End If
            reader.Close()
            If rightAccess = True Then
                Select Case e.Node.Name.Trim
                    Case "nodeAcadGrades"
                        frmAcadGrades.MdiParent = Me
                        frmAcadGrades.Show()
                    Case "nodeAcadSubjectsReg"
                        frmAcadSubjects.MdiParent = Me
                        frmAcadSubjects.Show()
                    Case "nodeAcadExams"
                        frmAcadExaminations.MdiParent = Me
                        frmAcadExaminations.Show()
                    Case "nodeAcadSubjectsTeacher"
                        frmAcadStaffSubject.MdiParent = Me
                        frmAcadStaffSubject.Show()
                    Case "nodeAcadSubjectsStudent"
                        frmAcadStudentSubject.MdiParent = Me
                        frmAcadStudentSubject.Show()
                    Case "nodeAcadMarkEntryClass"
                        frmAcadMarkEntryClassWise.MdiParent = Me
                        frmAcadMarkEntryClassWise.Show()
                    Case "nodeAcadMarkEntrySubj"
                        frmAcadMarkEntrySubjectWise.MdiParent = Me
                        frmAcadMarkEntrySubjectWise.Show()
                    Case "nodeResultsViewing"
                        frmAcadExamResultsView.MdiParent = Me
                        frmAcadExamResultsView.Show()
                    Case "nodeAcadResultAnalysis"
                        frmAcadResultsSummary.MdiParent = Me
                        frmAcadResultsSummary.Show()
                    Case "nodeAcadReportForm"
                        frmAcadReportForms.MdiParent = Me
                        frmAcadReportForms.Show()
                    Case "nodeExamNames"
                        frmAcadExaminationNames.MdiParent = Me
                        frmAcadExaminationNames.Show()
                End Select
            Else
                MsgBox("You have no rights to access this module!", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            End If
            rightAccess = False
        Catch ex As Exception
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub

    Private Sub tvTimeTable_NodeMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.TreeNodeMouseClickEventArgs)
        Try
            rightAccess = False
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            cmdHome.Connection = conn
            cmdHome.CommandType = CommandType.Text
            cmdHome.CommandText = "SELECT * FROM tblModules WHERE (modName=@modName) AND (status='True')"
            cmdHome.Parameters.Clear()
            cmdHome.Parameters.AddWithValue("@modName", e.Node.Text.Trim)
            reader = cmdHome.ExecuteReader
            If reader.HasRows Then
                Dim conn2 As New SqlConnection("Server=" & My.Settings.serverName.Trim & ";User ID=" & My.Settings.UserName.Trim & ";Database=" & My.Settings.dbName.Trim & ";Password=" & My.Settings.PassWord.Trim & "")
                conn2.Open()
                cmdHome1.CommandText = "SELECT * FROM vwUserRights WHERE (userName=@userName) AND (domainName=@domainName)" &
                vbNewLine & " AND (modName=@modName) AND (userStatus='True') AND (domainStatus='True') AND " &
                vbNewLine & " (staffStatus='True') AND (modStatus='True') AND (rightAccess='True')"
                cmdHome1.CommandType = CommandType.Text
                cmdHome1.Connection = conn2
                cmdHome1.Parameters.Clear()
                cmdHome1.Parameters.AddWithValue("@userName", userName)
                cmdHome1.Parameters.AddWithValue("@domainName", domainName)
                cmdHome1.Parameters.AddWithValue("@modName", e.Node.Text.Trim)
                reader1 = cmdHome1.ExecuteReader()
                If reader1.HasRows Then
                    rightAccess = True
                Else
                    rightAccess = False
                End If
                conn2.Close()
                reader1.Close()
            Else
                rightAccess = True
            End If
            reader.Close()
            If rightAccess = True Then
                Select Case e.Node.Name.Trim
                    Case "nodeTimeTableSetUp"
                        frmTTGeneralSetUp.MdiParent = Me
                        frmTTGeneralSetUp.Show()
                    Case "nodePeriodSetUp"
                        frmTTPeriodSetUp.MdiParent = Me
                        frmTTPeriodSetUp.Show()
                    Case "nodeSubjectSetUp"
                        frmTTSubjectSetup.MdiParent = Me
                        frmTTSubjectSetup.Show()
                    Case "nodeTeacherSetUp"
                        frmTTStaffSetUp.MdiParent = Me
                        frmTTStaffSetUp.Show()
                    Case "nodeLessonSetUp"
                        frmTTLessonSetUp.MdiParent = Me
                        frmTTLessonSetUp.Show()
                    Case "nodeTimetableGeneration"
                        frmTTGeneration.MdiParent = Me
                        frmTTGeneration.Show()
                    Case "nodeStaffReport"
                        frmTTPrintStaff.MdiParent = Me
                        frmTTPrintStaff.Show()
                    Case "nodeClassReport"
                        frmTTPrintClass.MdiParent = Me
                        frmTTPrintClass.Show()
                    Case "nodeMasterTT"
                        frmTTPrintMaster.MdiParent = Me
                        frmTTPrintMaster.Show()
                    Case "nodeStaffWorkLoad"
                        frmTTStaffWorkLoad.MdiParent = Me
                        frmTTStaffWorkLoad.Show()
                End Select
            Else
                MsgBox("You have no rights to access this module!", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            End If
            rightAccess = False
        Catch ex As Exception
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub
    Private Sub tvFinance_NodeMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.TreeNodeMouseClickEventArgs) Handles tvFinance.NodeMouseClick
        Try
            rightAccess = False
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            cmdHome.Connection = conn
            cmdHome.CommandType = CommandType.Text
            cmdHome.CommandText = "SELECT * FROM tblModules WHERE (modName=@modName) AND (status='True')"
            cmdHome.Parameters.Clear()
            cmdHome.Parameters.AddWithValue("@modName", e.Node.Text.Trim)
            reader = cmdHome.ExecuteReader
            If reader.HasRows Then
                Dim conn2 As New SqlConnection("Server=" & My.Settings.serverName.Trim & ";User ID=" & My.Settings.UserName.Trim & ";Database=" & My.Settings.dbName.Trim & ";Password=" & My.Settings.PassWord.Trim & "")
                conn2.Open()
                cmdHome1.CommandText = "SELECT * FROM vwUserRights WHERE (userName=@userName) AND (domainName=@domainName)" &
                vbNewLine & " AND (modName=@modName) AND (userStatus='True') AND (domainStatus='True') AND " &
                vbNewLine & " (staffStatus='True') AND (modStatus='True') AND (rightAccess='True')"
                cmdHome1.CommandType = CommandType.Text
                cmdHome1.Connection = conn2
                cmdHome1.Parameters.Clear()
                cmdHome1.Parameters.AddWithValue("@userName", userName)
                cmdHome1.Parameters.AddWithValue("@domainName", domainName)
                cmdHome1.Parameters.AddWithValue("@modName", e.Node.Text.Trim)
                reader1 = cmdHome1.ExecuteReader()
                If reader1.HasRows Then
                    rightAccess = True
                Else
                    rightAccess = False
                End If
                conn2.Close()
                reader1.Close()
            Else
                rightAccess = True
            End If
            reader.Close()
            If rightAccess = True Then
                Select Case e.Node.Name.Trim
                    Case "nodeFeeCat"
                        frmFinFeeCategory.MdiParent = Me
                        frmFinFeeCategory.Show()
                    Case "nodePaymentModes"
                        frmFinPayModes.MdiParent = Me
                        frmFinPayModes.Show()
                    Case "nodePayAccounts"
                        frmFinPayAccounts.MdiParent = Me
                        frmFinPayAccounts.Show()
                    Case "nodeBankBal"
                        frmFinBankBalances.MdiParent = Me
                        frmFinBankBalances.Show()
                    Case "nodeCashBal"
                        frmFinCashBalances.MdiParent = Me
                        frmFinCashBalances.Show()
                    Case "nodeMobileBal"
                        frmFinMobileBalances.MdiParent = Me
                        frmFinMobileBalances.Show()
                    Case "nodeAdjustBal"
                        frmFinAdjustBalances.MdiParent = Me
                        frmFinAdjustBalances.Show()
                    Case "nodeVoteHeads"
                        frmFinFeeVoteHeads.MdiParent = Me
                        frmFinFeeVoteHeads.Show()
                    Case "nodeFeeSetting"
                        frmFinFeeSetUp.MdiParent = Me
                        frmFinFeeSetUp.Show()
                    Case "nodeStudFee"
                        frmFinFeeStudent.MdiParent = Me
                        frmFinFeeStudent.Show()
                    Case "nodeFeeReceipts"
                        'frmFinFeeReceipt.MdiParent = Me
                        'frmFinFeeReceipt.Show()
                        frmFinFeeReceipt2.MdiParent = Me
                        frmFinFeeReceipt2.Show()
                    Case "nodeOthers"
                        frmFinOtherIncome.MdiParent = Me
                        frmFinOtherIncome.Show()
                    Case "nodeExpCat"
                        frmFinExpCategory.MdiParent = Me
                        frmFinExpCategory.Show()
                    Case "nodeExpMaster"
                        frmFinExpenseMaster.MdiParent = Me
                        frmFinExpenseMaster.Show()
                    Case "nodePayRequest"
                        frmFinPayRequest.MdiParent = Me
                        frmFinPayRequest.Show()
                    Case "nodePayApproval"
                        frmFinPayApproval.MdiParent = Me
                        frmFinPayApproval.Show()
                    Case "nodePaymentVoucher"
                        frmFinPayVoucher.MdiParent = Me
                        frmFinPayVoucher.Show()
                    Case "nodePayReversal"
                        frmFinPayReverse.MdiParent = Me
                        frmFinPayReverse.Show()
                    Case "nodeFindReceipt"
                        frmFinFeeEditReceipt.MdiParent = Me
                        frmFinFeeEditReceipt.Show()
                    Case "nodeAccountTransfers"
                        frmFinRptAccountsTrans.MdiParent = Me
                        frmFinRptAccountsTrans.Show()
                    Case "nodeAccountAdj"
                        frmFinRptAccountsAdj.MdiParent = Me
                        frmFinRptAccountsAdj.Show()
                    Case "nodeFeePayments"
                        frmFinRptFeeReceipts.MdiParent = Me
                        frmFinRptFeeReceipts.Show()
                    Case "nodeFeeBalances"
                        frmFinRptFeeBalances.MdiParent = Me
                        frmFinRptFeeBalances.Show()
                    Case "nodeFeeStatement"
                        frmFinFeeStatement.MdiParent = Me
                        frmFinFeeStatement.Show()
                    Case "nodeVoteSummary"
                        frmFinRptFeeVotesSummary.MdiParent = Me
                        frmFinRptFeeVotesSummary.Show()
                    Case "nodeFeeExpectation"
                        frmFinRptFeeExpectation.MdiParent = Me
                        frmFinRptFeeExpectation.Show()
                    Case "nodeOtherIncome"
                        frmFinRptOtherIncome.MdiParent = Me
                        frmFinRptOtherIncome.Show()
                    Case "nodeExpenses"
                        frmFinRptExpenses.MdiParent = Me
                        frmFinRptExpenses.Show()
                    Case "nodePayApprovals"
                        frmFinRptExpensesApprovals.MdiParent = Me
                        frmFinRptExpensesApprovals.Show()
                    Case "nodePayReversals"
                        frmFinRptExpensesReversals.MdiParent = Me
                        frmFinRptExpensesReversals.Show()
                    Case "nodeIncomeExp"
                        frmFinRptIncomeExpenditure.MdiParent = Me
                        frmFinRptIncomeExpenditure.Show()
                End Select
            Else
                MsgBox("You have no rights to access this module!", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            End If
            rightAccess = False
        Catch ex As Exception
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub

    Private Sub tvInventory_NodeMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.TreeNodeMouseClickEventArgs)
        Try
            rightAccess = False
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            cmdHome.Connection = conn
            cmdHome.CommandType = CommandType.Text
            cmdHome.CommandText = "SELECT * FROM tblModules WHERE (modName=@modName) AND (status='True')"
            cmdHome.Parameters.Clear()
            cmdHome.Parameters.AddWithValue("@modName", e.Node.Text.Trim)
            reader = cmdHome.ExecuteReader
            If reader.HasRows Then
                Dim conn2 As New SqlConnection("Server=" & My.Settings.serverName.Trim & ";User ID=" & My.Settings.UserName.Trim & ";Database=" & My.Settings.dbName.Trim & ";Password=" & My.Settings.PassWord.Trim & "")
                conn2.Open()
                cmdHome1.CommandText = "SELECT * FROM vwUserRights WHERE (userName=@userName) AND (domainName=@domainName)" &
                vbNewLine & " AND (modName=@modName) AND (userStatus='True') AND (domainStatus='True') AND " &
                vbNewLine & " (staffStatus='True') AND (modStatus='True') AND (rightAccess='True')"
                cmdHome1.CommandType = CommandType.Text
                cmdHome1.Connection = conn2
                cmdHome1.Parameters.Clear()
                cmdHome1.Parameters.AddWithValue("@userName", userName)
                cmdHome1.Parameters.AddWithValue("@domainName", domainName)
                cmdHome1.Parameters.AddWithValue("@modName", e.Node.Text.Trim)
                reader1 = cmdHome1.ExecuteReader()
                If reader1.HasRows Then
                    rightAccess = True
                Else
                    rightAccess = False
                End If
                conn2.Close()
                reader1.Close()
            Else
                rightAccess = True
            End If
            reader.Close()
            If rightAccess = True Then
                Select Case e.Node.Name.Trim
                    Case "nodeInvMaster"
                        frmInvMaster.MdiParent = Me
                        frmInvMaster.Show()
                    Case "nodeInvVendors"
                        frmInvVendorMaster.MdiParent = Me
                        frmInvVendorMaster.Show()
                    Case "nodeInvReceipts"
                        frmInvReceipts.MdiParent = Me
                        frmInvReceipts.Show()
                    Case "nodeInvIssueReq"
                        frmInvIssueRequest.MdiParent = Me
                        frmInvIssueRequest.Show()
                    Case "nodeInvIssueApproval"
                        frmInvIssueApproval.MdiParent = Me
                        frmInvIssueApproval.Show()
                    Case "nodeIssReqNote"
                        frmInvIssueRequestNote.MdiParent = Me
                        frmInvIssueRequestNote.Show()
                    Case "nodeCategoryMaster"
                        frmInvMasterCategory.MdiParent = Me
                        frmInvMasterCategory.Show()
                    Case "nodeInvReport"
                        frmInvRptReceipts.MdiParent = Me
                        frmInvRptReceipts.Show()
                    Case "nodeIssuesReport"
                        frmInvRptIssues.MdiParent = Me
                        frmInvRptIssues.Show()
                    Case "nodeReturn"
                        frmInvItemReturns.MdiParent = Me
                        frmInvItemReturns.Show()
                    Case "nodePendingReorder"
                        frmInvRptPendingReorder.MdiParent = Me
                        frmInvRptPendingReorder.Show()
                    Case "nodeInvPending"
                        frmInvRptPendingReturn.MdiParent = Me
                        frmInvRptPendingReturn.Show()
                    Case "nodeStockAtHand"
                        frmInvRptStockAtHand.MdiParent = Me
                        frmInvRptStockAtHand.Show()
                    Case "nodeConsolidated"
                        frmInvRptConsolidated.MdiParent = Me
                        frmInvRptConsolidated.Show()
                End Select
            Else
                MsgBox("You have no rights to access this module!", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            End If
            rightAccess = False
        Catch ex As Exception
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub

    Private Sub tvAccommodation_NodeMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.TreeNodeMouseClickEventArgs)
        Try
            rightAccess = False
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            cmdHome.Connection = conn
            cmdHome.CommandType = CommandType.Text
            cmdHome.CommandText = "SELECT * FROM tblModules WHERE (modName=@modName) AND (status='True')"
            cmdHome.Parameters.Clear()
            cmdHome.Parameters.AddWithValue("@modName", e.Node.Text.Trim)
            reader = cmdHome.ExecuteReader
            If reader.HasRows Then
                Dim conn2 As New SqlConnection("Server=" & My.Settings.serverName.Trim & ";User ID=" & My.Settings.UserName.Trim & ";Database=" & My.Settings.dbName.Trim & ";Password=" & My.Settings.PassWord.Trim & "")
                conn2.Open()
                cmdHome1.CommandText = "SELECT * FROM vwUserRights WHERE (userName=@userName) AND (domainName=@domainName)" &
                vbNewLine & " AND (modName=@modName) AND (userStatus='True') AND (domainStatus='True') AND " &
                vbNewLine & " (staffStatus='True') AND (modStatus='True') AND (rightAccess='True')"
                cmdHome1.CommandType = CommandType.Text
                cmdHome1.Connection = conn2
                cmdHome1.Parameters.Clear()
                cmdHome1.Parameters.AddWithValue("@userName", userName)
                cmdHome1.Parameters.AddWithValue("@domainName", domainName)
                cmdHome1.Parameters.AddWithValue("@modName", e.Node.Text.Trim)
                reader1 = cmdHome1.ExecuteReader()
                If reader1.HasRows Then
                    rightAccess = True
                Else
                    rightAccess = False
                End If
                conn2.Close()
                reader1.Close()
            Else
                rightAccess = True
            End If
            reader.Close()
            If rightAccess = True Then
                Select Case e.Node.Name.Trim
                    Case "nodeIDormMaster"
                        frmAccRegisterDorms.MdiParent = Me
                        frmAccRegisterDorms.Show()
                    Case "nodeAssignDorms"
                        frmAccDormStudents.MdiParent = Me
                        frmAccDormStudents.Show()
                    Case "nodeAssignHeads"
                        frmAccDormitoryHeads.MdiParent = Me
                        frmAccDormitoryHeads.Show()
                    Case "nodeDormStatus"
                        frmAccDormRpt.MdiParent = Me
                        frmAccDormRpt.Show()
                    Case "nodeStudentsDorms"
                        frmAccDormStudentsRpt.MdiParent = Me
                        frmAccDormStudentsRpt.Show()
                End Select
            Else
                MsgBox("You have no rights to access this module!", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            End If
            rightAccess = False
        Catch ex As Exception
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try

    End Sub

    Private Sub REVERSEFEESToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles REVERSEFEESToolStripMenuItem.Click
        If userName.Trim = "Arnold" Then
            frmFinFeeBatchReversal.MdiParent = Me
            frmFinFeeBatchReversal.Show()
        End If
    End Sub

    Private Sub tvStudent_AfterSelect(sender As Object, e As TreeViewEventArgs) Handles tvStudent.AfterSelect

    End Sub
End Class

