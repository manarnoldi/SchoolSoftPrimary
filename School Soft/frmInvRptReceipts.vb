Imports System.Data.SqlClient
Public Class frmInvRptReceipts
    Dim reader As SqlDataReader
    Dim rec As Integer = 0
    Dim cmdInvReceRpt As New SqlCommand
    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub rbMonthly_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbMonthly.CheckedChanged, rbDaily.CheckedChanged
        Me.crtVwRptInvReceipts.ReportSource = Nothing
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
            PnlLoc.X = CInt((Me.Width - Me.pnlInvReceiptsRpt.Width) / 2)
            PnlLoc.Y = CInt((Me.Height - Me.pnlInvReceiptsRpt.Height) / 2.5)
            Me.pnlInvReceiptsRpt.Location = PnlLoc
        Else
            Me.pnlInvReceiptsRpt.Dock = DockStyle.Fill
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
            Me.crtVwRptInvReceipts.ReportSource = Nothing
            Me.Cursor = Cursors.WaitCursor
            Me.cmdInvReceRpt.Connection = conn
            Me.cmdInvReceRpt.CommandType = CommandType.Text
            Me.Cursor = Cursors.WaitCursor
            If Me.rbDaily.Checked = True Then

                Me.cmdInvReceRpt.CommandText = "SELECT * FROM vwInvReceiptsDaily WHERE (dateOfReceipt=@dateOfReceipt)"
                Me.cmdInvReceRpt.Parameters.Clear()
                Me.cmdInvReceRpt.Parameters.AddWithValue("@dateOfReceipt", Me.dtpDate.Value.Date)
                reader = Me.cmdInvReceRpt.ExecuteReader
                If reader.HasRows = True Then
                    Dim RptResultsView As New crtInvRptReceiptsDaily
                    SetReportLogOn(RptResultsView)
                    RptResultsView.SummaryInfo.ReportTitle = "DAILY INVENTORY RECEIPTS REPORT FOR " & _
                        Me.dtpDate.Value.Day.ToString("00") & "-" & Me.dtpDate.Value.Month.ToString("00") & "-" & _
                        Me.dtpDate.Value.Year.ToString("0000")
                    RptResultsView.SummaryInfo.ReportComments = fullName.Trim
                    RptResultsView.RecordSelectionFormula = "(CDATE({vwInvReceiptsDaily.dateOfReceipt})=#" & (Me.dtpDate.Value.ToShortDateString) & "#)"
                    Me.crtVwRptInvReceipts.ReportSource = RptResultsView
                    Me.crtVwRptInvReceipts.Zoom(100)
                    Me.crtVwRptInvReceipts.RefreshReport()
                Else
                    MsgBox("No Record Found!", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "System Reply")
                End If
                reader.Close()

            ElseIf Me.rbMonthly.Checked = True Then

                Me.cmdInvReceRpt.CommandText = "SELECT * FROM vwInvReceiptsMonthly WHERE (month=@month) AND (year=@year)"
                Me.cmdInvReceRpt.Parameters.Clear()
                Me.cmdInvReceRpt.Parameters.AddWithValue("@month", Me.dtpDate.Value.Month)
                Me.cmdInvReceRpt.Parameters.AddWithValue("@year", Me.dtpDate.Value.Year)
                reader = Me.cmdInvReceRpt.ExecuteReader
                If reader.HasRows = True Then
                    Dim RptResultsView As New crtInvRptReceiptsMonthly
                    SetReportLogOn(RptResultsView)
                    RptResultsView.SummaryInfo.ReportTitle = "MONTHLY INVENTORY RECEIPTS REPORT FOR  " & _
                        Me.dtpDate.Text.Trim.ToUpper

                    RptResultsView.SummaryInfo.ReportComments = fullName.Trim
                    RptResultsView.RecordSelectionFormula = "({vwInvReceiptsMonthly.month}=" & Me.dtpDate.Value.Month & ")"
                    RptResultsView.RecordSelectionFormula += " AND ({vwInvReceiptsMonthly.year}=" & Me.dtpDate.Value.Year & ")"
                    Me.crtVwRptInvReceipts.ReportSource = RptResultsView
                    Me.crtVwRptInvReceipts.Zoom(100)
                    Me.crtVwRptInvReceipts.RefreshReport()
                Else
                    MsgBox("No Record Found!", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "System Reply")
                End If
                reader.Close()
            End If
            Me.Cursor = Cursors.Arrow
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            Me.Cursor = Cursors.Arrow
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub

    Private Sub dtpDate_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtpDate.ValueChanged
        Me.crtVwRptInvReceipts.ReportSource = Nothing
    End Sub

    Private Sub frmInvRptReceipts_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub
End Class