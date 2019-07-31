Imports System.Data.SqlClient
Public Class frmInvRptIssues
    Dim reader As SqlDataReader
    Dim rec As Integer = 0
    Dim cmdInvReceRpt As New SqlCommand
    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub rbMonthly_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbMonthly.CheckedChanged, rbDaily.CheckedChanged
        Me.crtVwRptInvIssues.ReportSource = Nothing
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
            PnlLoc.X = CInt((Me.Width - Me.pnlInvIssuesRpt.Width) / 2)
            PnlLoc.Y = CInt((Me.Height - Me.pnlInvIssuesRpt.Height) / 2.5)
            Me.pnlInvIssuesRpt.Location = PnlLoc
        Else
            Me.pnlInvIssuesRpt.Dock = DockStyle.Fill
        End If
    End Sub

    Private Sub btnLoad_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLoad.Click
        If Me.rbDaily.Checked = False And Me.rbMonthly.Checked = False Then
            MsgBox("Check for either Daily or Monthly.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboStatus.Text.Trim.Length <= 0 Then
            MsgBox("Select status.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            Me.crtVwRptInvIssues.ReportSource = Nothing
            Me.Cursor = Cursors.WaitCursor
            Me.cmdInvReceRpt.Connection = conn
            Me.cmdInvReceRpt.CommandType = CommandType.Text

            If Me.rbDaily.Checked = True Then
                If Me.cboStatus.Text.Trim = "All" Then
                    Me.cmdInvReceRpt.CommandText = "SELECT * FROM vwInvIssuesDaily WHERE (issueReqDate=@issueReqDate)"
                    Me.cmdInvReceRpt.Parameters.Clear()
                    Me.cmdInvReceRpt.Parameters.AddWithValue("@issueReqDate", Me.dtpDate.Value.Date)
                    reader = Me.cmdInvReceRpt.ExecuteReader
                    If reader.HasRows = True Then
                        Dim RptResultsView As New crtInvRptIssuesDaily
                        SetReportLogOn(RptResultsView)
                        RptResultsView.SummaryInfo.ReportTitle = "DAILY INVENTORY ISSUES REPORT FOR " & _
                            Me.dtpDate.Value.Day.ToString("00") & "-" & Me.dtpDate.Value.Month.ToString("00") & "-" & _
                            Me.dtpDate.Value.Year.ToString("0000")
                        RptResultsView.SummaryInfo.ReportComments = fullName.Trim
                        RptResultsView.RecordSelectionFormula = "(CDATE({vwInvIssuesDaily.issueReqDate})=#" & (Me.dtpDate.Value.ToShortDateString) & "#)"
                        Me.crtVwRptInvIssues.ReportSource = RptResultsView
                        Me.crtVwRptInvIssues.Zoom(100)
                        Me.crtVwRptInvIssues.RefreshReport()
                    Else
                        MsgBox("No Record Found!", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "System Reply")
                    End If
                    reader.Close()
                Else
                    Me.cmdInvReceRpt.CommandText = "SELECT * FROM vwInvIssuesDaily WHERE (issueReqDate=@issueReqDate) AND (status=@status)"
                    Me.cmdInvReceRpt.Parameters.Clear()
                    Me.cmdInvReceRpt.Parameters.AddWithValue("@issueReqDate", Me.dtpDate.Value.Date)
                    Me.cmdInvReceRpt.Parameters.AddWithValue("@status", Me.cboStatus.Text.Trim)
                    reader = Me.cmdInvReceRpt.ExecuteReader
                    If reader.HasRows = True Then
                        Dim RptResultsView As New crtInvRptIssuesDaily
                        SetReportLogOn(RptResultsView)
                        RptResultsView.SummaryInfo.ReportTitle = "DAILY INVENTORY ISSUES REPORT FOR " & _
                            Me.dtpDate.Value.Day.ToString("00") & "-" & Me.dtpDate.Value.Month.ToString("00") & "-" & _
                            Me.dtpDate.Value.Year.ToString("0000")
                        RptResultsView.SummaryInfo.ReportComments = fullName.Trim
                        RptResultsView.RecordSelectionFormula = "(CDATE({vwInvIssuesDaily.issueReqDate})=#" & (Me.dtpDate.Value.ToShortDateString) & "#)"
                        RptResultsView.RecordSelectionFormula += "AND ({vwInvIssuesDaily.status}=" & Chr(34) & Me.cboStatus.Text.Trim & Chr(34) & ")"
                        Me.crtVwRptInvIssues.ReportSource = RptResultsView
                        Me.crtVwRptInvIssues.Zoom(100)
                        Me.crtVwRptInvIssues.RefreshReport()
                    Else
                        MsgBox("No Record Found!", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "System Reply")
                    End If
                    reader.Close()
                End If

            ElseIf Me.rbMonthly.Checked = True Then
                If Me.cboStatus.Text.Trim = "All" Then
                    Me.cmdInvReceRpt.CommandText = "SELECT * FROM vwInvIssuesMonthly WHERE (month=@month) AND (year=@year)"
                    Me.cmdInvReceRpt.Parameters.Clear()
                    Me.cmdInvReceRpt.Parameters.AddWithValue("@month", Me.dtpDate.Value.Month)
                    Me.cmdInvReceRpt.Parameters.AddWithValue("@year", Me.dtpDate.Value.Year)
                    reader = Me.cmdInvReceRpt.ExecuteReader
                    If reader.HasRows = True Then
                        Dim RptResultsView As New crtInvRptIssuesMonthly
                        SetReportLogOn(RptResultsView)
                        RptResultsView.SummaryInfo.ReportTitle = "MONTHLY INVENTORY ISSUES REPORT FOR  " & _
                            Me.dtpDate.Text.Trim.ToUpper

                        RptResultsView.SummaryInfo.ReportComments = fullName.Trim
                        RptResultsView.RecordSelectionFormula = "({vwInvIssuesMonthly.month}=" & Me.dtpDate.Value.Month & ")"
                        RptResultsView.RecordSelectionFormula += " AND ({vwInvIssuesMonthly.year}=" & Me.dtpDate.Value.Year & ")"
                        Me.crtVwRptInvIssues.ReportSource = RptResultsView
                        Me.crtVwRptInvIssues.Zoom(100)
                        Me.crtVwRptInvIssues.RefreshReport()
                    Else
                        MsgBox("No Record Found!", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "System Reply")
                    End If
                    reader.Close()
                Else
                    Me.cmdInvReceRpt.CommandText = "SELECT * FROM vwInvIssuesMonthly WHERE (month=@month) AND (year=@year) AND (status=@status)"
                    Me.cmdInvReceRpt.Parameters.Clear()
                    Me.cmdInvReceRpt.Parameters.AddWithValue("@month", Me.dtpDate.Value.Month)
                    Me.cmdInvReceRpt.Parameters.AddWithValue("@year", Me.dtpDate.Value.Year)
                    Me.cmdInvReceRpt.Parameters.AddWithValue("@status", Me.cboStatus.Text.Trim)
                    reader = Me.cmdInvReceRpt.ExecuteReader
                    If reader.HasRows = True Then
                        Dim RptResultsView As New crtInvRptIssuesMonthly
                        SetReportLogOn(RptResultsView)
                        RptResultsView.SummaryInfo.ReportTitle = "MONTHLY INVENTORY ISSUES REPORT FOR  " & _
                            Me.dtpDate.Text.Trim.ToUpper

                        RptResultsView.SummaryInfo.ReportComments = fullName.Trim
                        RptResultsView.RecordSelectionFormula = "({vwInvIssuesMonthly.month}=" & Me.dtpDate.Value.Month & ")"
                        RptResultsView.RecordSelectionFormula += " AND ({vwInvIssuesMonthly.year}=" & Me.dtpDate.Value.Year & ")"
                        RptResultsView.RecordSelectionFormula += "AND ({vwInvIssuesMonthly.status}=" & Chr(34) & Me.cboStatus.Text.Trim & Chr(34) & ")"
                        Me.crtVwRptInvIssues.ReportSource = RptResultsView
                        Me.crtVwRptInvIssues.Zoom(100)
                        Me.crtVwRptInvIssues.RefreshReport()
                    Else
                        MsgBox("No Record Found!", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "System Reply")
                    End If
                    reader.Close()
                End If
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
        Me.crtVwRptInvIssues.ReportSource = Nothing
    End Sub

    Private Sub cboStatus_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboStatus.SelectedIndexChanged
        Me.crtVwRptInvIssues.ReportSource = Nothing
    End Sub

    Private Sub frmInvRptIssues_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub
End Class