Imports System.Data.SqlClient
Public Class frmFinRptExpensesReversals
    Dim reader As SqlDataReader
    Dim rec As Integer = 0
    Dim cmdExpReversals As New SqlCommand

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub rbMonthly_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbMonthly.CheckedChanged, rbDaily.CheckedChanged
        Me.crtVwFinExpRev.ReportSource = Nothing
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
            PnlLoc.X = CInt((Me.Width - Me.pnlFinExpRev.Width) / 2)
            PnlLoc.Y = CInt((Me.Height - Me.pnlFinExpRev.Height) / 2.5)
            Me.pnlFinExpRev.Location = PnlLoc
        Else
            Me.pnlFinExpRev.Dock = DockStyle.Fill
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
            Me.crtVwFinExpRev.ReportSource = Nothing
            Me.Cursor = Cursors.WaitCursor
            Me.cmdExpReversals.Connection = conn
            Me.cmdExpReversals.CommandType = CommandType.Text

            If Me.rbDaily.Checked = True Then

                Me.cmdExpReversals.CommandText = "SELECT * FROM vwFinExpReversalReport WHERE (reverseDateOnly=@reverseDateOnly)"
                Me.cmdExpReversals.Parameters.Clear()
                Me.cmdExpReversals.Parameters.AddWithValue("@reverseDateOnly", Me.dtpDate.Value.Date)
                reader = Me.cmdExpReversals.ExecuteReader
                If reader.HasRows = True Then
                    Dim RptResultsView As New crtFinRptExpenseReversalsDaily
                    SetReportLogOn(RptResultsView)
                    RptResultsView.SummaryInfo.ReportTitle = "DAILY FINANCE EXPENSE REVERSAL REPORT FOR " & _
                        Me.dtpDate.Value.Day.ToString("00") & "-" & Me.dtpDate.Value.Month.ToString("00") & "-" & _
                        Me.dtpDate.Value.Year.ToString("0000")
                    RptResultsView.SummaryInfo.ReportComments = fullName.Trim
                    RptResultsView.RecordSelectionFormula = "(CDATE({vwFinExpReversalReport.reverseDateOnly})=#" & (Me.dtpDate.Value.ToShortDateString) & "#)"
                    Me.crtVwFinExpRev.ReportSource = RptResultsView
                    Me.crtVwFinExpRev.Zoom(100)
                    Me.crtVwFinExpRev.RefreshReport()
                Else
                    MsgBox("No Record Found!", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "System Reply")
                End If
                reader.Close()

            ElseIf Me.rbMonthly.Checked = True Then

                Me.cmdExpReversals.CommandText = "SELECT * FROM vwFinExpReversalReport WHERE (month=@month) AND (year=@year)"
                Me.cmdExpReversals.Parameters.Clear()
                Me.cmdExpReversals.Parameters.AddWithValue("@month", Me.dtpDate.Value.Month)
                Me.cmdExpReversals.Parameters.AddWithValue("@year", Me.dtpDate.Value.Year)
                reader = Me.cmdExpReversals.ExecuteReader
                If reader.HasRows = True Then
                    Dim RptResultsView As New crtFinRptExpenseReversalsMonthly
                    SetReportLogOn(RptResultsView)
                    RptResultsView.SummaryInfo.ReportTitle = "MONTHLY FINANCE EXPENSE REVERSAL REPORT FOR  " & _
                        Me.dtpDate.Text.Trim.ToUpper

                    RptResultsView.SummaryInfo.ReportComments = fullName.Trim
                    RptResultsView.RecordSelectionFormula = "({vwFinExpReversalReport.month}=" & Me.dtpDate.Value.Month & ")"
                    RptResultsView.RecordSelectionFormula += " AND ({vwFinExpReversalReport.year}=" & Me.dtpDate.Value.Year & ")"
                    Me.crtVwFinExpRev.ReportSource = RptResultsView
                    Me.crtVwFinExpRev.Zoom(100)
                    Me.crtVwFinExpRev.RefreshReport()
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

    Private Sub dtpDate_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Me.crtVwFinExpRev.ReportSource = Nothing
    End Sub
End Class