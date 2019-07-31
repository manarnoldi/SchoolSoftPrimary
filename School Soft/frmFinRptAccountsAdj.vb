Imports System.Data.SqlClient
Public Class frmFinRptAccountsAdj
    Dim reader As SqlDataReader
    Dim rec As Integer
    Dim cmdAccountAdj As New SqlCommand

    Private Sub rbYearly_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbYearly.CheckedChanged
        Me.crtVwFinAccAdjRpt.ReportSource = Nothing
    End Sub

    Private Sub dtpDate_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtpDate.ValueChanged
        Me.crtVwFinAccAdjRpt.ReportSource = Nothing
    End Sub

    Private Sub frmFinRptAccountsAdj_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.rbYearly.Checked = True
    End Sub

    Private Sub frmFinRptAccountsAdj_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint
        If Me.IsMdiChild Then
            Dim PnlLoc As New Point
            PnlLoc.X = CInt((Me.Width - Me.pnlAccountAdj.Width) / 2)
            PnlLoc.Y = CInt((Me.Height - Me.pnlAccountAdj.Height) / 2.5)
            Me.pnlAccountAdj.Location = PnlLoc
        Else
            Me.pnlAccountAdj.Dock = DockStyle.Fill
        End If
    End Sub

    Private Sub btnLoad_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLoad.Click
        If Me.rbYearly.Checked = False Then
            MsgBox("Check Yearly Button.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            Me.crtVwFinAccAdjRpt.ReportSource = Nothing
            Me.Cursor = Cursors.WaitCursor
            Me.cmdAccountAdj.Connection = conn
            Me.cmdAccountAdj.CommandType = CommandType.Text

            If Me.rbYearly.Checked = True Then

                Me.cmdAccountAdj.CommandText = "SELECT * FROM vwFinAccountAdjYearly WHERE (year=@year)"
                Me.cmdAccountAdj.Parameters.Clear()
                Me.cmdAccountAdj.Parameters.AddWithValue("@year", Me.dtpDate.Value.Year)
                reader = Me.cmdAccountAdj.ExecuteReader
                If reader.HasRows = True Then
                    Dim RptResultsView As New crtFinRptAccountAdjYearly
                    SetReportLogOn(RptResultsView)
                    RptResultsView.SummaryInfo.ReportTitle = "YEARLY FINANCE ACCOUNTS ADJUSTMENT REPORT FOR  " & _
                        Me.dtpDate.Value.Year.ToString("0000")

                    RptResultsView.SummaryInfo.ReportComments = fullName.Trim
                    RptResultsView.RecordSelectionFormula = "({vwFinAccountAdjYearly.year}=" & Me.dtpDate.Value.Year & ")"
                    Me.crtVwFinAccAdjRpt.ReportSource = RptResultsView
                    Me.crtVwFinAccAdjRpt.Zoom(100)
                    Me.crtVwFinAccAdjRpt.RefreshReport()
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

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub
End Class