Imports System.Data.SqlClient
Public Class frmFinRptExpensesApprovals
    Dim reader As SqlDataReader
    Dim rec As Integer = 0
    Dim cmdExpApprovals As New SqlCommand

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub rbMonthly_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbMonthly.CheckedChanged, rbDaily.CheckedChanged
        Me.crtVwFinExpApp.ReportSource = Nothing
        If Me.rbDaily.Checked = True Then
            Me.lblDateLabel.Text = "Select Date : "
            Me.dtpDate.Format = DateTimePickerFormat.Custom
            Me.dtpDate.CustomFormat = " dd - MM - yyyy"
            Me.dtpDate.ShowUpDown = False

        ElseIf Me.rbMonthly.Checked = True Then
            Me.lblDateLabel.Text = "Select Month : "
            Me.dtpDate.Format = DateTimePickerFormat.Custom
            Me.dtpDate.CustomFormat = " MMMM, yyyy"
            Me.dtpDate.ShowUpDown = True

        End If
    End Sub
    Private Sub frmFinRptExpensesApprovals_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint
        If Me.IsMdiChild Then
            Dim PnlLoc As New Point
            PnlLoc.X = CInt((Me.Width - Me.pnlFinExpApp.Width) / 2)
            PnlLoc.Y = CInt((Me.Height - Me.pnlFinExpApp.Height) / 2.5)
            Me.pnlFinExpApp.Location = PnlLoc
        Else
            Me.pnlFinExpApp.Dock = DockStyle.Fill
        End If
    End Sub

    Private Sub btnLoad_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLoad.Click
        If Me.rbDaily.Checked = False And Me.rbMonthly.Checked = False Then
            MsgBox("Check for either Daily or Monthly.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            Me.crtVwFinExpApp.ReportSource = Nothing
            Me.Cursor = Cursors.WaitCursor
            Me.cmdExpApprovals.Connection = conn
            Me.cmdExpApprovals.CommandType = CommandType.Text

            If Me.rbDaily.Checked = True Then

                Me.cmdExpApprovals.CommandText = "SELECT * FROM vwFinExpApprovalReport WHERE (approvalDateOnly=@approvalDateOnly)"
                Me.cmdExpApprovals.Parameters.Clear()
                Me.cmdExpApprovals.Parameters.AddWithValue("@approvalDateOnly", Me.dtpDate.Value.Date)
                reader = Me.cmdExpApprovals.ExecuteReader
                If reader.HasRows = True Then
                    Dim RptResultsView As New crtFinRptExpenseApprovalsDaily
                    SetReportLogOn(RptResultsView)
                    RptResultsView.SummaryInfo.ReportTitle = "DAILY FINANCE EXPENSE APPROVAL REPORT FOR " & _
                        Me.dtpDate.Value.Day.ToString("00") & "-" & Me.dtpDate.Value.Month.ToString("00") & "-" & _
                        Me.dtpDate.Value.Year.ToString("0000")
                    RptResultsView.SummaryInfo.ReportComments = fullName.Trim
                    RptResultsView.RecordSelectionFormula = "(CDATE({vwFinExpApprovalReport.approvalDateOnly})=#" & (Me.dtpDate.Value.ToShortDateString) & "#)"
                    Me.crtVwFinExpApp.ReportSource = RptResultsView
                    Me.crtVwFinExpApp.Zoom(100)
                    Me.crtVwFinExpApp.RefreshReport()
                Else
                    MsgBox("No Record Found!", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "System Reply")
                End If
                reader.Close()

            ElseIf Me.rbMonthly.Checked = True Then

                Me.cmdExpApprovals.CommandText = "SELECT * FROM vwFinExpApprovalReport WHERE (month=@month) AND (year=@year)"
                Me.cmdExpApprovals.Parameters.Clear()
                Me.cmdExpApprovals.Parameters.AddWithValue("@month", Me.dtpDate.Value.Month)
                Me.cmdExpApprovals.Parameters.AddWithValue("@year", Me.dtpDate.Value.Year)
                reader = Me.cmdExpApprovals.ExecuteReader
                If reader.HasRows = True Then
                    Dim RptResultsView As New crtFinRptExpenseApprovalsMonthly
                    SetReportLogOn(RptResultsView)
                    RptResultsView.SummaryInfo.ReportTitle = "MONTHLY FINANCE EXPENSE APPROVAL REPORT FOR  " & _
                        Me.dtpDate.Text.Trim.ToUpper

                    RptResultsView.SummaryInfo.ReportComments = fullName.Trim
                    RptResultsView.RecordSelectionFormula = "({vwFinExpApprovalReport.month}=" & Me.dtpDate.Value.Month & ")"
                    RptResultsView.RecordSelectionFormula += " AND ({vwFinExpApprovalReport.year}=" & Me.dtpDate.Value.Year & ")"
                    Me.crtVwFinExpApp.ReportSource = RptResultsView
                    Me.crtVwFinExpApp.Zoom(100)
                    Me.crtVwFinExpApp.RefreshReport()
                Else
                    MsgBox("No Record Found!", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "System Reply")
                End If
                reader.Close()
            End If
            Me.Cursor = Cursors.Arrow
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub

    Private Sub dtpDate_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtpDate.ValueChanged
        Me.crtVwFinExpApp.ReportSource = Nothing
    End Sub

End Class