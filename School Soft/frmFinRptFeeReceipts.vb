Imports System.Data.SqlClient
Public Class frmFinRptFeeReceipts
    Dim reader As SqlDataReader
    Dim rec As Integer
    Dim cmdRptFee As New SqlCommand
    Private Sub frmFinRptFeeReceipts_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
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
        Me.cboTerm.Items.Clear()
        Me.cboTerm.Text = ""
        Me.cboTerm.SelectedIndex = -1

        Me.cmdRptFee.Connection = conn
        Me.cmdRptFee.CommandType = CommandType.Text
        Me.cmdRptFee.CommandText = "SELECT DISTINCT termName FROM tblSchoolcalendar WHERE (status=1) ORDER BY termName"
        Me.cmdRptFee.Parameters.Clear()
        reader = Me.cmdRptFee.ExecuteReader
        While reader.Read
            Me.cboTerm.Items.Add(IIf(DBNull.Value.Equals(reader!termName), "", reader!termName))
        End While
        reader.Close()
    End Sub
    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub rbDaily_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbDaily.CheckedChanged, rbMonthly.CheckedChanged, rbYearly.CheckedChanged, rbTermly.CheckedChanged
        Me.crtVwFeeReceiptsRpt.ReportSource = Nothing
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

    Private Sub frmFinRptFeeReceipts_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint
        If Me.IsMdiChild Then
            Dim PnlLoc As New Point
            PnlLoc.X = CInt((Me.Width - Me.pnlFeePayRpt.Width) / 2)
            PnlLoc.Y = CInt((Me.Height - Me.pnlFeePayRpt.Height) / 2.5)
            Me.pnlFeePayRpt.Location = PnlLoc
        Else
            Me.pnlFeePayRpt.Dock = DockStyle.Fill
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
            Me.crtVwFeeReceiptsRpt.ReportSource = Nothing
            Me.Cursor = Cursors.WaitCursor
            Me.cmdRptFee.Connection = conn
            Me.cmdRptFee.CommandType = CommandType.Text

            If Me.rbDaily.Checked = True Then

                Me.cmdRptFee.CommandText = "SELECT * FROM vwFinFeeReceiptsDailyRpt WHERE (actualPayDate=@actualPayDate)"
                Me.cmdRptFee.Parameters.Clear()
                Me.cmdRptFee.Parameters.AddWithValue("@actualPayDate", Me.dtpDate.Value.Date)
                reader = Me.cmdRptFee.ExecuteReader
                If reader.HasRows = True Then
                    Dim RptResultsView As New crtFinFeeReceiptsDailyRpt
                    SetReportLogOn(RptResultsView)
                    RptResultsView.SummaryInfo.ReportTitle = "DAILY FINANCE FEE RECEIPTS REPORT FOR " & _
                        Me.dtpDate.Value.Day.ToString("00") & "-" & Me.dtpDate.Value.Month.ToString("00") & "-" & _
                        Me.dtpDate.Value.Year.ToString("0000")
                    RptResultsView.SummaryInfo.ReportComments = fullName.Trim
                    RptResultsView.RecordSelectionFormula = "(CDATE({vwFinFeeReceiptsDailyRpt.actualPayDate})=#" & (Me.dtpDate.Value.ToShortDateString) & "#)"
                    Me.crtVwFeeReceiptsRpt.ReportSource = RptResultsView
                    Me.crtVwFeeReceiptsRpt.Zoom(100)
                    Me.crtVwFeeReceiptsRpt.RefreshReport()
                Else
                    MsgBox("No Record Found!", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "System Reply")
                End If
                reader.Close()

            ElseIf Me.rbMonthly.Checked = True Then

                Me.cmdRptFee.CommandText = "SELECT * FROM vwFinFeeReceiptsMonthlyRpt WHERE (monthInt=@monthInt) AND (year=@year)"
                Me.cmdRptFee.Parameters.Clear()
                Me.cmdRptFee.Parameters.AddWithValue("@monthInt", Me.dtpDate.Value.Month)
                Me.cmdRptFee.Parameters.AddWithValue("@year", Me.dtpDate.Value.Year)
                reader = Me.cmdRptFee.ExecuteReader
                If reader.HasRows = True Then
                    Dim RptResultsView As New crtFinFeeReceiptsMonthlyRpt
                    SetReportLogOn(RptResultsView)
                    RptResultsView.SummaryInfo.ReportTitle = "MONTHLY FINANCE FEE RECEIPTS REPORT FOR  " & _
                        Me.dtpDate.Text.Trim.ToUpper

                    RptResultsView.SummaryInfo.ReportComments = fullName.Trim
                    RptResultsView.RecordSelectionFormula = "({vwFinFeeReceiptsMonthlyRpt.monthInt}=" & Me.dtpDate.Value.Month & ")"
                    RptResultsView.RecordSelectionFormula += " AND ({vwFinFeeReceiptsMonthlyRpt.year}=" & Me.dtpDate.Value.Year & ")"
                    Me.crtVwFeeReceiptsRpt.ReportSource = RptResultsView
                    Me.crtVwFeeReceiptsRpt.Zoom(100)
                    Me.crtVwFeeReceiptsRpt.RefreshReport()
                Else
                    MsgBox("No Record Found!", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "System Reply")
                End If
                reader.Close()

            ElseIf Me.rbTermly.Checked = True Then

                Me.cmdRptFee.CommandText = "SELECT * FROM vwFinFeeReceiptsTermlyRpt WHERE (year=@year) AND (termName=@termName)"
                Me.cmdRptFee.Parameters.Clear()
                Me.cmdRptFee.Parameters.AddWithValue("@year", Me.dtpDate.Value.Year)
                Me.cmdRptFee.Parameters.AddWithValue("@termName", Me.cboTerm.Text.Trim)
                reader = Me.cmdRptFee.ExecuteReader
                If reader.HasRows = True Then
                    Dim RptResultsView As New crtFinFeeReceiptsTermly
                    SetReportLogOn(RptResultsView)
                    RptResultsView.SummaryInfo.ReportTitle = "TERMLY FINANCE FEE RECEIPTS REPORT FOR TERM " & Me.cboTerm.Text.Trim & " AND YEAR " & _
                        Me.dtpDate.Value.Year.ToString("0000")

                    RptResultsView.SummaryInfo.ReportComments = fullName.Trim
                    RptResultsView.RecordSelectionFormula = "({vwFinFeeReceiptsTermlyRpt.year}=" & Me.dtpDate.Value.Year & ")"
                    RptResultsView.RecordSelectionFormula += "AND ({vwFinFeeReceiptsTermlyRpt.termName}=" & Chr(34) & Me.cboTerm.Text.Trim & Chr(34) & ")"
                    Me.crtVwFeeReceiptsRpt.ReportSource = RptResultsView
                    Me.crtVwFeeReceiptsRpt.Zoom(100)
                    Me.crtVwFeeReceiptsRpt.RefreshReport()
                Else
                    MsgBox("No Record Found!", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "System Reply")
                End If
                reader.Close()

            ElseIf Me.rbYearly.Checked = True Then

                Me.cmdRptFee.CommandText = "SELECT * FROM vwFinFeeReceiptsYearlyRpt WHERE (year=@year)"
                Me.cmdRptFee.Parameters.Clear()
                Me.cmdRptFee.Parameters.AddWithValue("@year", Me.dtpDate.Value.Year)
                reader = Me.cmdRptFee.ExecuteReader
                If reader.HasRows = True Then
                    Dim RptResultsView As New crtFinFeeReceiptsYearly
                    SetReportLogOn(RptResultsView)
                    RptResultsView.SummaryInfo.ReportTitle = "YEARLY FINANCE FEE RECEIPTS REPORT FOR  " & _
                        Me.dtpDate.Value.Year.ToString("0000")

                    RptResultsView.SummaryInfo.ReportComments = fullName.Trim
                    RptResultsView.RecordSelectionFormula = "({vwFinFeeReceiptsYearlyRpt.year}=" & Me.dtpDate.Value.Year & ")"
                    Me.crtVwFeeReceiptsRpt.ReportSource = RptResultsView
                    Me.crtVwFeeReceiptsRpt.Zoom(100)
                    Me.crtVwFeeReceiptsRpt.RefreshReport()
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
        Me.crtVwFeeReceiptsRpt.ReportSource = Nothing
    End Sub

    Private Sub dtpDate_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtpDate.ValueChanged
        Me.crtVwFeeReceiptsRpt.ReportSource = Nothing
    End Sub
End Class