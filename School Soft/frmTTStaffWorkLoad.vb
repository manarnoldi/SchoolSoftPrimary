Imports System.Data.SqlClient
Public Class frmTTStaffWorkLoad
    Dim cmdStaffWorkload As New SqlCommand
    Dim reader As SqlDataReader
    Dim rec As Integer
    Private Sub frmTTStaffWorkLoad_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
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
        Me.cboYear.Items.Clear()
        Me.cboYear.SelectedIndex = -1
        Me.cboYear.Text = ""

        Me.cmdStaffWorkload.Connection = conn
        Me.cmdStaffWorkload.CommandType = CommandType.Text
        Me.cmdStaffWorkload.CommandText = "SELECT DISTINCT year FROM tblClasses WHERE (status=1) ORDER BY year"
        Me.cmdStaffWorkload.Parameters.Clear()
        reader = Me.cmdStaffWorkload.ExecuteReader
        While reader.Read
            Me.cboYear.Items.Add(IIf(DBNull.Value.Equals(reader!year), "", reader!year))
        End While
        reader.Close()
    End Sub
    Private Sub frmTTStaffWorkLoad_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint
        If Me.IsMdiChild Then
            Dim PnlLoc As New Point
            PnlLoc.X = CInt((Me.Width - Me.pnlStaffWorkLoad.Width) / 2)
            PnlLoc.Y = CInt((Me.Height - Me.pnlStaffWorkLoad.Height) / 2.5)
            Me.pnlStaffWorkLoad.Location = PnlLoc
        Else
            Me.pnlStaffWorkLoad.Dock = DockStyle.Fill
        End If
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub btnLoad_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLoad.Click
        Dim RptHeader As String = Nothing
        If Me.cboYear.Text.Trim.Length <= 0 Then
            MsgBox("Year is Missing", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly + MsgBoxStyle.ApplicationModal, "Missing Details")
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()

            Me.Cursor = Cursors.WaitCursor
            Dim RptResultsView As New crtTTStaffWorkload
            RptHeader = "STAFF WORKLOAD REPORT FOR " & Me.cboYear.Text.Trim
            My.Settings.Reload()
            RptResultsView.DataSourceConnections(0).SetConnection("ARNOLD-PC", "SchoolSoft", False)
            RptResultsView.DataSourceConnections(0).SetLogon("sa", "123")
            'SetReportLogOn(RptResultsView)
            RptResultsView.SummaryInfo.ReportComments = fullName.Trim
            RptResultsView.SummaryInfo.ReportTitle = RptHeader
            RptResultsView.RecordSelectionFormula = "({vwTTRptStaffWorkLoad.year}=" & Me.cboYear.Text.Trim & ")"
            Me.crtVwStaffWorkLoad.ReportSource = RptResultsView
            Me.crtVwStaffWorkLoad.Zoom(100)
            Me.crtVwStaffWorkLoad.RefreshReport()
            Me.Cursor = Cursors.Arrow
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub
End Class