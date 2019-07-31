Imports System.Data.SqlClient
Public Class frmInvRptStockAtHand
    Dim reader As SqlDataReader
    Dim rec As Integer
    Dim cmdStockAtHand As New SqlCommand
    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub frmInvRptStockAtHand_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub

    Private Sub frmInvRptStockAtHand_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint
        If Me.IsMdiChild Then
            Dim PnlLoc As New Point
            PnlLoc.X = CInt((Me.Width - Me.pnlInvStockAtHandRpt.Width) / 2)
            PnlLoc.Y = CInt((Me.Height - Me.pnlInvStockAtHandRpt.Height) / 2.5)
            Me.pnlInvStockAtHandRpt.Location = PnlLoc
        Else
            Me.pnlInvStockAtHandRpt.Dock = DockStyle.Fill
        End If
    End Sub

    Private Sub btnLoad_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLoad.Click
        
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            Me.cmdStockAtHand.Connection = conn
            Me.cmdStockAtHand.CommandType = CommandType.Text
            Me.cmdStockAtHand.CommandText = "SELECT * FROM vwInvStockAtHand"
            Me.cmdStockAtHand.Parameters.Clear()
            reader = Me.cmdStockAtHand.ExecuteReader
            If reader.HasRows = True Then
                Me.Cursor = Cursors.WaitCursor
                Dim RptResultsView As New crtInvRptStockAtHand
                SetReportLogOn(RptResultsView)
                RptResultsView.SummaryInfo.ReportTitle = " STOCK ON HAND AS AT NOW "
                RptResultsView.SummaryInfo.ReportComments = fullName.Trim
                Me.crtVwRptInvStockAtHand.ReportSource = RptResultsView
                Me.crtVwRptInvStockAtHand.Zoom(100)
                Me.crtVwRptInvStockAtHand.RefreshReport()
            Else
                MsgBox("No Record Found!", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "System Reply")
            End If
            reader.Close()
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            Me.Cursor = Cursors.Arrow
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub
End Class