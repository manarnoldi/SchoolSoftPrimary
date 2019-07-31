Imports System.Data.SqlClient
Public Class frmFinRptExpenses
    Dim reader As SqlDataReader
    Dim rec As Integer
    Dim cmdRptVoteSummary As New SqlCommand
    Private Sub frmFinRptExpenses_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            loadCombos()
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub
    Private Sub loadCombos()
        Me.cmdRptVoteSummary.Connection = conn
        Me.cmdRptVoteSummary.CommandType = CommandType.Text
        Me.cmdRptVoteSummary.CommandText = "SELECT DISTINCT termName FROM tblSchoolcalendar WHERE (status=1) ORDER BY termName"
        Me.cmdRptVoteSummary.Parameters.Clear()
        reader = Me.cmdRptVoteSummary.ExecuteReader
        While reader.Read
            Me.cboTerm.Items.Add(IIf(DBNull.Value.Equals(reader!termName), "", reader!termName))
        End While
        reader.Close()
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub rbDaily_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbDaily.CheckedChanged, rbMonthly.CheckedChanged, rbTermly.CheckedChanged, rbYearly.CheckedChanged
        Me.crtVwExpenseReport.ReportSource = Nothing
        If Me.rbDaily.Checked = True Then
            Me.lblDateLabel.Text = "Select Date : "
            Me.dtpDate.Format = DateTimePickerFormat.Custom
            Me.dtpDate.CustomFormat = " dd - MM - yyyy"
            Me.dtpDate.ShowUpDown = False

            Me.cboTerm.Text = ""
            Me.cboTerm.SelectedIndex = -1
            Me.cboTerm.Enabled = False

        ElseIf Me.rbMonthly.Checked = True Then
            Me.lblDateLabel.Text = "Select Month : "
            Me.dtpDate.Format = DateTimePickerFormat.Custom
            Me.dtpDate.CustomFormat = " MMMM, yyyy"
            Me.dtpDate.ShowUpDown = True

            Me.cboTerm.Text = ""
            Me.cboTerm.SelectedIndex = -1
            Me.cboTerm.Enabled = False

        ElseIf Me.rbTermly.Checked = True Then
            Me.lblDateLabel.Text = "Select Year : "
            Me.dtpDate.Format = DateTimePickerFormat.Custom
            Me.dtpDate.CustomFormat = "    yyyy"
            Me.dtpDate.ShowUpDown = True

            Me.cboTerm.Text = ""
            Me.cboTerm.SelectedIndex = -1
            Me.cboTerm.Enabled = True

        ElseIf Me.rbYearly.Checked = True Then
            Me.lblDateLabel.Text = "Select Year : "
            Me.dtpDate.Format = DateTimePickerFormat.Custom
            Me.dtpDate.CustomFormat = "    yyyy"
            Me.dtpDate.ShowUpDown = True

            Me.cboTerm.Text = ""
            Me.cboTerm.SelectedIndex = -1
            Me.cboTerm.Enabled = False

        End If
    End Sub

    Private Sub frmFinRptFeeVotesSummary_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint
        If Me.IsMdiChild Then
            Dim PnlLoc As New Point
            PnlLoc.X = CInt((Me.Width - Me.pnlExpenseRpt.Width) / 2)
            PnlLoc.Y = CInt((Me.Height - Me.pnlExpenseRpt.Height) / 2.5)
            Me.pnlExpenseRpt.Location = PnlLoc
        Else
            Me.pnlExpenseRpt.Dock = DockStyle.Fill
        End If
    End Sub

    Private Sub btnLoad_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLoad.Click
        If Me.rbDaily.Checked = False And Me.rbMonthly.Checked = False And Me.rbTermly.Checked = False And Me.rbYearly.Checked = False Then
            MsgBox("Check for either Daily,Monthly,Termly Or Yearly.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.rbTermly.Checked = True Then
            If Me.cboTerm.Text.Trim.Length <= 0 Then
                MsgBox("Term Name Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
                Exit Sub
            End If
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            Me.crtVwExpenseReport.ReportSource = Nothing
            Me.Cursor = Cursors.WaitCursor
            Me.cmdRptVoteSummary.Connection = conn
            Me.cmdRptVoteSummary.CommandType = CommandType.Text

            If Me.rbDaily.Checked = True Then

                Me.cmdRptVoteSummary.CommandText = "SELECT * FROM vwFinExpensesDaily WHERE (voucherDate=@voucherDate)"
                Me.cmdRptVoteSummary.Parameters.Clear()
                Me.cmdRptVoteSummary.Parameters.AddWithValue("@voucherDate", Me.dtpDate.Value.Date)
                reader = Me.cmdRptVoteSummary.ExecuteReader
                If reader.HasRows = True Then
                    Dim RptResultsView As New crtFinRptExpensesDaily
                    SetReportLogOn(RptResultsView)
                    RptResultsView.SummaryInfo.ReportTitle = "DAILY EXPENSES REPORT FOR " & _
                        Me.dtpDate.Value.Day.ToString("00") & "-" & Me.dtpDate.Value.Month.ToString("00") & "-" & _
                        Me.dtpDate.Value.Year.ToString("0000")
                    RptResultsView.SummaryInfo.ReportComments = fullName.Trim
                    RptResultsView.RecordSelectionFormula = "(CDATE({vwFinExpensesDaily.voucherDate})=#" & (Me.dtpDate.Value.ToShortDateString) & "#)"
                    Me.crtVwExpenseReport.ReportSource = RptResultsView
                    Me.crtVwExpenseReport.Zoom(100)
                    Me.crtVwExpenseReport.RefreshReport()
                Else
                    MsgBox("No Record Found!", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "System Reply")
                End If
                reader.Close()

            ElseIf Me.rbMonthly.Checked = True Then

                Me.cmdRptVoteSummary.CommandText = "SELECT * FROM vwFinExpensesMonthly WHERE (month=@month) AND (year=@year)"
                Me.cmdRptVoteSummary.Parameters.Clear()
                Me.cmdRptVoteSummary.Parameters.AddWithValue("@month", Me.dtpDate.Value.Month)
                Me.cmdRptVoteSummary.Parameters.AddWithValue("@year", Me.dtpDate.Value.Year)
                reader = Me.cmdRptVoteSummary.ExecuteReader
                If reader.HasRows = True Then
                    Dim RptResultsView As New crtFinRptExpensesMonthly
                    SetReportLogOn(RptResultsView)
                    RptResultsView.SummaryInfo.ReportTitle = "MONTHLY EXPENSES REPORT FOR  " & _
                        Me.dtpDate.Text.Trim.ToUpper

                    RptResultsView.SummaryInfo.ReportComments = fullName.Trim
                    RptResultsView.RecordSelectionFormula = "({vwFinExpensesMonthly.month}=" & Me.dtpDate.Value.Month & ")"
                    RptResultsView.RecordSelectionFormula += " AND ({vwFinExpensesMonthly.year}=" & Me.dtpDate.Value.Year & ")"
                    Me.crtVwExpenseReport.ReportSource = RptResultsView
                    Me.crtVwExpenseReport.Zoom(100)
                    Me.crtVwExpenseReport.RefreshReport()
                Else
                    MsgBox("No Record Found!", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "System Reply")
                End If
                reader.Close()

            ElseIf Me.rbTermly.Checked = True Then

                Me.cmdRptVoteSummary.CommandText = "SELECT * FROM vwFinExpensesTermly WHERE (year=@year) AND (termName=@termName)"
                Me.cmdRptVoteSummary.Parameters.Clear()
                Me.cmdRptVoteSummary.Parameters.AddWithValue("@year", Me.dtpDate.Value.Year)
                Me.cmdRptVoteSummary.Parameters.AddWithValue("@termName", Me.cboTerm.Text.Trim)
                reader = Me.cmdRptVoteSummary.ExecuteReader
                If reader.HasRows = True Then
                    Dim RptResultsView As New crtFinRptExpensesTermly
                    SetReportLogOn(RptResultsView)
                    RptResultsView.SummaryInfo.ReportTitle = "TERMLY EXPENSES REPORT FOR " & Me.cboTerm.Text.Trim & " " & _
                        Me.dtpDate.Value.Year.ToString("0000")

                    RptResultsView.SummaryInfo.ReportComments = fullName.Trim
                    RptResultsView.RecordSelectionFormula = "({vwFinExpensesTermly.year}=" & Me.dtpDate.Value.Year & ")"
                    RptResultsView.RecordSelectionFormula += "AND ({vwFinExpensesTermly.termName}=" & Chr(34) & Me.cboTerm.Text.Trim & Chr(34) & ")"
                    Me.crtVwExpenseReport.ReportSource = RptResultsView
                    Me.crtVwExpenseReport.Zoom(100)
                    Me.crtVwExpenseReport.RefreshReport()
                Else
                    MsgBox("No Record Found!", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "System Reply")
                End If
                reader.Close()

            ElseIf Me.rbYearly.Checked = True Then

                Me.cmdRptVoteSummary.CommandText = "SELECT * FROM vwFinExpensesYearly WHERE (year=@year)"
                Me.cmdRptVoteSummary.Parameters.Clear()
                Me.cmdRptVoteSummary.Parameters.AddWithValue("@year", Me.dtpDate.Value.Year)
                reader = Me.cmdRptVoteSummary.ExecuteReader
                If reader.HasRows = True Then
                    Dim RptResultsView As New crtFinRptExpensesYearly
                    SetReportLogOn(RptResultsView)
                    RptResultsView.SummaryInfo.ReportTitle = "YEARLY EXPENSES REPORT FOR  " & _
                        Me.dtpDate.Value.Year.ToString("0000")

                    RptResultsView.SummaryInfo.ReportComments = fullName.Trim
                    RptResultsView.RecordSelectionFormula = "({vwFinExpensesYearly.year}=" & Me.dtpDate.Value.Year & ")"
                    Me.crtVwExpenseReport.ReportSource = RptResultsView
                    Me.crtVwExpenseReport.Zoom(100)
                    Me.crtVwExpenseReport.RefreshReport()
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

    Private Sub cboTerm_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboTerm.SelectedIndexChanged
        Me.crtVwExpenseReport.ReportSource = Nothing
    End Sub

    Private Sub dtpDate_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtpDate.ValueChanged
        Me.crtVwExpenseReport.ReportSource = Nothing
    End Sub
End Class