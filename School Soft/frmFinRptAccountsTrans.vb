Imports System.Data.SqlClient
Public Class frmFinRptAccountsTrans
    Dim reader As SqlDataReader
    Dim cmdAccountsTrans As New SqlCommand
    Dim rec As Integer = 0
    Private Sub rbDaily_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbDaily.CheckedChanged, rbMonthly.CheckedChanged, rbYearly.CheckedChanged
        Me.crtVwFinAccTransRpt.ReportSource = Nothing
        If Me.rbDaily.Checked = True And Me.rbMonthly.Checked = False And Me.rbYearly.Checked = False Then
            Me.lblDateLabel.Text = "Select Date : "
            Me.dtpDate.Format = DateTimePickerFormat.Custom
            Me.dtpDate.CustomFormat = "         dd - MM - yyyy"
            Me.dtpDate.ShowUpDown = False

        ElseIf Me.rbDaily.Checked = False And Me.rbMonthly.Checked = True And Me.rbYearly.Checked = False Then
            Me.lblDateLabel.Text = "Select Month : "
            Me.dtpDate.Format = DateTimePickerFormat.Custom
            Me.dtpDate.CustomFormat = "    MMMM, yyyy"
            Me.dtpDate.ShowUpDown = True

        ElseIf Me.rbDaily.Checked = False And Me.rbMonthly.Checked = False And Me.rbYearly.Checked = True Then
            Me.lblDateLabel.Text = "Select Year : "
            Me.dtpDate.Format = DateTimePickerFormat.Custom
            Me.dtpDate.CustomFormat = "           yyyy"
            Me.dtpDate.ShowUpDown = True

        End If
    End Sub

    Private Sub frmFinRptAccountsTrans_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint
        If Me.IsMdiChild Then
            Dim PnlLoc As New Point
            PnlLoc.X = CInt((Me.Width - Me.pnlAccountTrans.Width) / 2)
            PnlLoc.Y = CInt((Me.Height - Me.pnlAccountTrans.Height) / 2.5)
            Me.pnlAccountTrans.Location = PnlLoc
        Else
            Me.pnlAccountTrans.Dock = DockStyle.Fill
        End If
    End Sub

    Private Sub btnLoad_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLoad.Click
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            Me.crtVwFinAccTransRpt.ReportSource = Nothing
            Me.Cursor = Cursors.WaitCursor
            Me.cmdAccountsTrans.Connection = conn
            Me.cmdAccountsTrans.CommandType = CommandType.Text

            If Me.rbDaily.Checked = True Then

                Me.cmdAccountsTrans.CommandText = "SELECT * FROM vwFinAccountTransDaily WHERE (dateDoneOnly=@dateDoneOnly)"
                Me.cmdAccountsTrans.Parameters.Clear()
                Me.cmdAccountsTrans.Parameters.AddWithValue("@dateDoneOnly", Me.dtpDate.Value.Date)
                reader = Me.cmdAccountsTrans.ExecuteReader
                If reader.HasRows = True Then
                    Dim RptResultsView As New crtFinRptAccountTransDaily
                    SetReportLogOn(RptResultsView)
                    RptResultsView.SummaryInfo.ReportTitle = "DAILY FINANCE INTER-ACCOUNTS TRANSFER REPORT FOR " & _
                        Me.dtpDate.Value.Day.ToString("00") & "-" & Me.dtpDate.Value.Month.ToString("00") & "-" & _
                        Me.dtpDate.Value.Year.ToString("0000")
                    RptResultsView.SummaryInfo.ReportComments = fullName.Trim
                    RptResultsView.RecordSelectionFormula = "(CDATE({vwFinAccountTransDaily.dateDoneOnly})=#" & (Me.dtpDate.Value.ToShortDateString) & "#)"
                    Me.crtVwFinAccTransRpt.ReportSource = RptResultsView
                    Me.crtVwFinAccTransRpt.Zoom(100)
                    Me.crtVwFinAccTransRpt.RefreshReport()
                Else
                    MsgBox("No Record Found!", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "System Reply")
                End If
                reader.Close()

            ElseIf Me.rbMonthly.Checked = True Then

                Me.cmdAccountsTrans.CommandText = "SELECT * FROM vwFinAccountTransMonthly WHERE (month=@month) AND (year=@year)"
                Me.cmdAccountsTrans.Parameters.Clear()
                Me.cmdAccountsTrans.Parameters.AddWithValue("@month", Me.dtpDate.Value.Month)
                Me.cmdAccountsTrans.Parameters.AddWithValue("@year", Me.dtpDate.Value.Year)
                reader = Me.cmdAccountsTrans.ExecuteReader
                If reader.HasRows = True Then
                    Dim RptResultsView As New crtFinRptAccountTransMonthly
                    SetReportLogOn(RptResultsView)
                    RptResultsView.SummaryInfo.ReportTitle = "MONTHLY FINANCE INTER-ACCOUNTS TRANSFER REPORT FOR  " & _
                        Me.dtpDate.Text.Trim.ToUpper

                    RptResultsView.SummaryInfo.ReportComments = fullName.Trim
                    RptResultsView.RecordSelectionFormula = "({vwFinAccountTransMonthly.month}=" & Me.dtpDate.Value.Month & ")"
                    RptResultsView.RecordSelectionFormula += " AND ({vwFinAccountTransMonthly.year}=" & Me.dtpDate.Value.Year & ")"
                    Me.crtVwFinAccTransRpt.ReportSource = RptResultsView
                    Me.crtVwFinAccTransRpt.Zoom(100)
                    Me.crtVwFinAccTransRpt.RefreshReport()
                Else
                    MsgBox("No Record Found!", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "System Reply")
                End If
                reader.Close()

            ElseIf Me.rbYearly.Checked = True Then

                Me.cmdAccountsTrans.CommandText = "SELECT * FROM vwFinAccountTransYearly WHERE (year=@year)"
                Me.cmdAccountsTrans.Parameters.Clear()
                Me.cmdAccountsTrans.Parameters.AddWithValue("@year", Me.dtpDate.Value.Year)
                reader = Me.cmdAccountsTrans.ExecuteReader
                If reader.HasRows = True Then
                    Dim RptResultsView As New crtFinRptAccountTransYearly
                    SetReportLogOn(RptResultsView)
                    RptResultsView.SummaryInfo.ReportTitle = "YEARLY FINANCE INTER-ACCOUNTS TRANSFER REPORT FOR  " & _
                        Me.dtpDate.Value.Year.ToString("0000")

                    RptResultsView.SummaryInfo.ReportComments = fullName.Trim
                    RptResultsView.RecordSelectionFormula = "({vwFinAccountTransYearly.year}=" & Me.dtpDate.Value.Year & ")"
                    Me.crtVwFinAccTransRpt.ReportSource = RptResultsView
                    Me.crtVwFinAccTransRpt.Zoom(100)
                    Me.crtVwFinAccTransRpt.RefreshReport()
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
        Me.crtVwFinAccTransRpt.ReportSource = Nothing
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub frmFinRptAccountsTrans_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub
End Class