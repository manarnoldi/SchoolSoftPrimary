Imports System.Data.SqlClient
Public Class frmInvRptPendingReorder
    Dim reader As SqlDataReader
    Dim rec As Integer
    Dim cmdPendingReorder As New SqlCommand
    Private Sub frmInvRptPendingReorder_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
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
        Me.cboStoreName.Items.Clear()
        Me.cboStoreName.Text = ""
        Me.cboStoreName.SelectedIndex = -1

        Me.cmdPendingReorder.Connection = conn
        Me.cmdPendingReorder.CommandType = CommandType.Text
        Me.cmdPendingReorder.CommandText = "SELECT DISTINCT categoryName FROM tblInvMasterCategory ORDER BY categoryName"
        Me.cmdPendingReorder.Parameters.Clear()
        reader = Me.cmdPendingReorder.ExecuteReader
        While reader.Read
            Me.cboStoreName.Items.Add(IIf(DBNull.Value.Equals(reader!categoryName), "", reader!categoryName))
        End While
        reader.Close()
    End Sub
    Private Sub frmInvRptPendingReorder_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint
        If Me.IsMdiChild Then
            Dim PnlLoc As New Point
            PnlLoc.X = CInt((Me.Width - Me.pnlInvReorderRpt.Width) / 2)
            PnlLoc.Y = CInt((Me.Height - Me.pnlInvReorderRpt.Height) / 2.5)
            Me.pnlInvReorderRpt.Location = PnlLoc
        Else
            Me.pnlInvReorderRpt.Dock = DockStyle.Fill
        End If
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub cboStoreName_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboStoreName.SelectedIndexChanged
        Me.crtVwRptInvIssues.ReportSource = Nothing
    End Sub

    Private Sub btnLoad_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLoad.Click
        Me.crtVwRptInvIssues.ReportSource = Nothing
        If Me.cboStoreName.Text.Trim.Length <= 0 Then
            MsgBox("Store Name Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            Me.cmdPendingReorder.Connection = conn
            Me.cmdPendingReorder.CommandType = CommandType.Text
            Me.cmdPendingReorder.CommandText = "SELECT * FROM vwInvPendingReorder WHERE (categoryName=@categoryName)"
            Me.cmdPendingReorder.Parameters.Clear()
            Me.cmdPendingReorder.Parameters.AddWithValue("@categoryName", Me.cboStoreName.Text.Trim)
            reader = Me.cmdPendingReorder.ExecuteReader
            If reader.HasRows = True Then
                Dim RptResultsView As New crtInvRptPendingReorder
                SetReportLogOn(RptResultsView)
                RptResultsView.SummaryInfo.ReportTitle = "INVENTORY REORDER REPORT FOR " & Me.cboStoreName.Text.Trim.ToUpper & " AS OF " & _
                    Date.Now.Day.ToString("00") & "-" & Date.Now.Month.ToString("00") & "-" & _
                    Date.Now.Year.ToString("0000")
                RptResultsView.SummaryInfo.ReportComments = fullName.Trim
                RptResultsView.RecordSelectionFormula = "({vwInvPendingReorder.categoryName}=" & Chr(34) & Me.cboStoreName.Text.Trim & Chr(34) & ")"
                Me.crtVwRptInvIssues.ReportSource = RptResultsView
                Me.crtVwRptInvIssues.Zoom(100)
                Me.crtVwRptInvIssues.RefreshReport()
            Else
                MsgBox("No Record Found!", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "System Reply")
            End If
            reader.Close()
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub
End Class