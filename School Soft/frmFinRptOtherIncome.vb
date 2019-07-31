Imports System.Data.SqlClient
Public Class frmFinRptOtherIncome
    Dim reader As SqlDataReader
    Dim rec As Integer = 0
    Dim cmdOtherInc As New SqlCommand
    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub frmFinRptOtherIncome_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
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
        Me.cmdOtherInc.Connection = conn
        Me.cmdOtherInc.CommandType = CommandType.Text
        Me.cmdOtherInc.CommandText = "SELECT DISTINCT termName FROM tblSchoolcalendar WHERE (status=1) ORDER BY termName"
        Me.cmdOtherInc.Parameters.Clear()
        reader = Me.cmdOtherInc.ExecuteReader
        While reader.Read
            Me.cboTerm.Items.Add(IIf(DBNull.Value.Equals(reader!termName), "", reader!termName))
        End While
        reader.Close()
    End Sub

    Private Sub frmFinRptOtherIncome_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint
        If Me.IsMdiChild Then
            Dim PnlLoc As New Point
            PnlLoc.X = CInt((Me.Width - Me.pnlOtherIncome.Width) / 2)
            PnlLoc.Y = CInt((Me.Height - Me.pnlOtherIncome.Height) / 2.5)
            Me.pnlOtherIncome.Location = PnlLoc
        Else
            Me.pnlOtherIncome.Dock = DockStyle.Fill
        End If
    End Sub

    Private Sub btnLoad_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLoad.Click
        If Me.rbMonthly.Checked = False And Me.rbTermly.Checked = False And Me.rbYearly.Checked = False Then
            MsgBox("Check for either Monthly,Termly Or Yearly.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
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
            Me.crtVwOtherInc.ReportSource = Nothing
            Me.Cursor = Cursors.WaitCursor
            Me.cmdOtherInc.Connection = conn
            Me.cmdOtherInc.CommandType = CommandType.Text

            If Me.rbMonthly.Checked = True Then

                Me.cmdOtherInc.CommandText = "SELECT * FROM vwFinOtherIncMonthlyRpt WHERE (month=@month) AND (year=@year)"
                Me.cmdOtherInc.Parameters.Clear()
                Me.cmdOtherInc.Parameters.AddWithValue("@month", Me.dtpDate.Value.Month)
                Me.cmdOtherInc.Parameters.AddWithValue("@year", Me.dtpDate.Value.Year)
                reader = Me.cmdOtherInc.ExecuteReader
                If reader.HasRows = True Then
                    Dim RptResultsView As New crtFinRptOtherIncMonthly
                    SetReportLogOn(RptResultsView)
                    RptResultsView.SummaryInfo.ReportTitle = "FINANCE MONTHLY OTHER INCOME REPORT FOR  " & _
                        Me.dtpDate.Text.Trim.ToUpper

                    RptResultsView.SummaryInfo.ReportComments = fullName.Trim
                    RptResultsView.RecordSelectionFormula = "({vwFinOtherIncMonthlyRpt.month}=" & Me.dtpDate.Value.Month & ")"
                    RptResultsView.RecordSelectionFormula += " AND ({vwFinOtherIncMonthlyRpt.year}=" & Me.dtpDate.Value.Year & ")"
                    Me.crtVwOtherInc.ReportSource = RptResultsView
                    Me.crtVwOtherInc.Zoom(100)
                    Me.crtVwOtherInc.RefreshReport()
                Else
                    MsgBox("No Record Found!", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "System Reply")
                End If
                reader.Close()

            ElseIf Me.rbTermly.Checked = True Then

                Me.cmdOtherInc.CommandText = "SELECT * FROM vwFinOtherIncTermlyRpt WHERE (year=@year) AND (termName=@termName)"
                Me.cmdOtherInc.Parameters.Clear()
                Me.cmdOtherInc.Parameters.AddWithValue("@year", Me.dtpDate.Value.Year)
                Me.cmdOtherInc.Parameters.AddWithValue("@termName", Me.cboTerm.Text.Trim)
                reader = Me.cmdOtherInc.ExecuteReader
                If reader.HasRows = True Then
                    Dim RptResultsView As New crtFinRptOtherIncTermly
                    SetReportLogOn(RptResultsView)
                    RptResultsView.SummaryInfo.ReportTitle = "FINANCE TERMLY OTHER INCOME REPORT FOR " & Me.cboTerm.Text.Trim & " " & _
                        Me.dtpDate.Value.Year.ToString("0000")

                    RptResultsView.SummaryInfo.ReportComments = fullName.Trim
                    RptResultsView.RecordSelectionFormula = "({vwFinOtherIncTermlyRpt.year}=" & Me.dtpDate.Value.Year & ")"
                    RptResultsView.RecordSelectionFormula += "AND ({vwFinOtherIncTermlyRpt.termName}=" & Chr(34) & Me.cboTerm.Text.Trim & Chr(34) & ")"
                    Me.crtVwOtherInc.ReportSource = RptResultsView
                    Me.crtVwOtherInc.Zoom(100)
                    Me.crtVwOtherInc.RefreshReport()
                Else
                    MsgBox("No Record Found!", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "System Reply")
                End If
                reader.Close()

            ElseIf Me.rbYearly.Checked = True Then

                Me.cmdOtherInc.CommandText = "SELECT * FROM vwFinOtherIncYearlyRpt WHERE (year=@year)"
                Me.cmdOtherInc.Parameters.Clear()
                Me.cmdOtherInc.Parameters.AddWithValue("@year", Me.dtpDate.Value.Year)
                reader = Me.cmdOtherInc.ExecuteReader
                If reader.HasRows = True Then
                    Dim RptResultsView As New crtFinRptOtherIncYearly
                    SetReportLogOn(RptResultsView)
                    RptResultsView.SummaryInfo.ReportTitle = "FINANCE YEAR OTHER INCOME REPORT FOR  " & _
                        Me.dtpDate.Value.Year.ToString("0000")

                    RptResultsView.SummaryInfo.ReportComments = fullName.Trim
                    RptResultsView.RecordSelectionFormula = "({vwFinOtherIncYearlyRpt.year}=" & Me.dtpDate.Value.Year & ")"
                    Me.crtVwOtherInc.ReportSource = RptResultsView
                    Me.crtVwOtherInc.Zoom(100)
                    Me.crtVwOtherInc.RefreshReport()
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

    Private Sub rbMonthly_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbMonthly.CheckedChanged, rbTermly.CheckedChanged, rbYearly.CheckedChanged
        Me.crtVwOtherInc.ReportSource = Nothing
        If Me.rbMonthly.Checked = True Then
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
End Class