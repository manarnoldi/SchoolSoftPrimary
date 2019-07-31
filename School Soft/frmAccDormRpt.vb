Imports System.Data.SqlClient
Public Class frmAccDormRpt
    Dim reader As SqlDataReader
    Dim rec As Integer
    Dim cmdConsolidated As New SqlCommand

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub frmInvRptConsolidated_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
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
        Me.cboYear.Text = ""
        Me.cboYear.SelectedIndex = -1

        Me.cmdConsolidated.Connection = conn
        Me.cmdConsolidated.CommandType = CommandType.Text
        Me.cmdConsolidated.CommandText = "SELECT DISTINCT year FROM tblSchoolCalendar WHERE (status=1) ORDER BY year"
        Me.cmdConsolidated.Parameters.Clear()
        reader = Me.cmdConsolidated.ExecuteReader
        While reader.Read
            Me.cboYear.Items.Add(IIf(DBNull.Value.Equals(reader!year), "", reader!year))
        End While
        reader.Close()
    End Sub

    Private Sub frmInvRptConsolidated_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint
        If Me.IsMdiChild Then
            Dim PnlLoc As New Point
            PnlLoc.X = CInt((Me.Width - Me.pnlAccDormDetails.Width) / 2)
            PnlLoc.Y = CInt((Me.Height - Me.pnlAccDormDetails.Height) / 2.5)
            Me.pnlAccDormDetails.Location = PnlLoc
        Else
            Me.pnlAccDormDetails.Dock = DockStyle.Fill
        End If
    End Sub

    Private Sub btnLoad_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLoad.Click
        Me.crtVwRptAccDormDetails.ReportSource = Nothing
        If Me.cboYear.Text.Trim.Length <= 0 Then
            MsgBox("Year Is Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            Me.cmdConsolidated.Connection = conn
            Me.cmdConsolidated.CommandType = CommandType.Text
            Me.cmdConsolidated.CommandText = "SELECT * FROM vwAccDormDetails WHERE (academicYear=@academicYear)"
            Me.cmdConsolidated.Parameters.Clear()
            Me.cmdConsolidated.Parameters.AddWithValue("@academicYear", Me.cboYear.Text.Trim)
            reader = Me.cmdConsolidated.ExecuteReader
            If reader.HasRows = True Then
                Me.Cursor = Cursors.WaitCursor
                Dim RptResultsView As New crtAccDormRpt
                SetReportLogOn(RptResultsView)
                RptResultsView.SummaryInfo.ReportComments = fullName.Trim
                RptResultsView.RecordSelectionFormula = "({vwAccDormDetails.academicYear}=" & Me.cboYear.Text.Trim & ")"
                Me.crtVwRptAccDormDetails.ReportSource = RptResultsView
                Me.crtVwRptAccDormDetails.Zoom(100)
                Me.crtVwRptAccDormDetails.RefreshReport()
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

    Private Sub cboYear_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboYear.SelectedIndexChanged
        Me.crtVwRptAccDormDetails.ReportSource = Nothing
    End Sub
End Class