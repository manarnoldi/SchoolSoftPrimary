Imports System.Data.SqlClient
Public Class frmFinRptFeeExpectation
    Dim reader As SqlDataReader
    Dim rec As Integer = 0
    Dim cmdFeeExpectation As New SqlCommand

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub frmFinRptFeeExpectation_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            loadcombos()
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub
    Private Sub loadcombos()
        Me.cboTerm.Items.Clear()
        Me.cboTerm.Text = ""
        Me.cboTerm.SelectedIndex = -1

        Me.cboYear.Items.Clear()
        Me.cboYear.Text = ""
        Me.cboYear.SelectedIndex = -1

        Me.cmdFeeExpectation.CommandType = CommandType.Text
        Me.cmdFeeExpectation.Connection = conn
        Me.cmdFeeExpectation.CommandText = "SELECT DISTINCT termName FROM tblSchoolCalendar WHERE (status=1) ORDER BY termName"
        Me.cmdFeeExpectation.Parameters.Clear()
        reader = Me.cmdFeeExpectation.ExecuteReader
        While reader.Read
            Me.cboTerm.Items.Add(IIf(DBNull.Value.Equals(reader!termName), "", reader!termName))
        End While
        reader.Close()

        Me.cmdFeeExpectation.CommandText = "SELECT DISTINCT year FROM tblSchoolCalendar WHERE (status=1) ORDER BY year"
        Me.cmdFeeExpectation.Parameters.Clear()
        reader = Me.cmdFeeExpectation.ExecuteReader
        While reader.Read
            Me.cboYear.Items.Add(IIf(DBNull.Value.Equals(reader!year), "", reader!year))
        End While
        reader.Close()
    End Sub

    Private Sub frmFinRptFeeExpectation_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint
        If Me.IsMdiChild Then
            Dim PnlLoc As New Point
            PnlLoc.X = CInt((Me.Width - Me.pnlFeeExpectation.Width) / 2)
            PnlLoc.Y = CInt((Me.Height - Me.pnlFeeExpectation.Height) / 2.5)
            Me.pnlFeeExpectation.Location = PnlLoc
        Else
            Me.pnlFeeExpectation.Dock = DockStyle.Fill
        End If
    End Sub

    Private Sub btnLoad_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLoad.Click
        If Me.cboTerm.Text.Trim.Length <= 0 Then
            MsgBox("Term Is Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboYear.Text.Trim.Length <= 0 Then
            MsgBox("Year Is Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            Me.crtVwFeeExpectation.ReportSource = Nothing
            Me.Cursor = Cursors.WaitCursor
            Me.cmdFeeExpectation.Connection = conn
            Me.cmdFeeExpectation.CommandType = CommandType.Text
            Me.cmdFeeExpectation.CommandText = "SELECT * FROM vwFinFeeExpectations WHERE (year=@year) AND (termName=@termName)"
            Me.cmdFeeExpectation.Parameters.Clear()
            Me.cmdFeeExpectation.Parameters.AddWithValue("@year", Me.cboYear.Text.Trim)
            Me.cmdFeeExpectation.Parameters.AddWithValue("@termName", Me.cboTerm.Text.Trim)
            reader = Me.cmdFeeExpectation.ExecuteReader
            If reader.HasRows = True Then
                Dim RptResultsView As New crtFinRptFeeExpectation
                SetReportLogOn(RptResultsView)
                RptResultsView.SummaryInfo.ReportTitle = "TERMLY FINANCE FEE EXPECTATION SUMMARY FOR " & Me.cboTerm.Text.Trim & " " & _
                    Me.cboYear.Text.Trim

                RptResultsView.SummaryInfo.ReportComments = fullName.Trim
                RptResultsView.RecordSelectionFormula = "({vwFinFeeExpectations.year}=" & Me.cboYear.Text.Trim & ")"
                RptResultsView.RecordSelectionFormula += "AND ({vwFinFeeExpectations.termName}=" & Chr(34) & Me.cboTerm.Text.Trim & Chr(34) & ")"
                Me.crtVwFeeExpectation.ReportSource = RptResultsView
                Me.crtVwFeeExpectation.Zoom(100)
                Me.crtVwFeeExpectation.RefreshReport()
            Else
                MsgBox("No Record Found!", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "System Reply")
            End If
            reader.Close()
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