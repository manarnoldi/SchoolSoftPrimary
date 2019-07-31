Imports System.Data.SqlClient
Public Class frmFinRptIncomeExpenditure
    Dim reader As SqlDataReader
    Dim rec As Integer
    Dim cmdInExp As New SqlCommand
    Private Sub frmFinRptIncomeExpenditure_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
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
        Me.cmdInExp.Connection = conn
        Me.cmdInExp.CommandType = CommandType.Text
        Me.cmdInExp.CommandText = "SELECT DISTINCT termName FROM tblSchoolcalendar WHERE (status=1) ORDER BY termName"
        Me.cmdInExp.Parameters.Clear()
        reader = Me.cmdInExp.ExecuteReader
        While reader.Read
            Me.cboTerm.Items.Add(IIf(DBNull.Value.Equals(reader!termName), "", reader!termName))
        End While
        reader.Close()

        Me.cmdInExp.CommandText = "SELECT DISTINCT year FROM tblSchoolcalendar WHERE (status=1) ORDER BY year"
        Me.cmdInExp.Parameters.Clear()
        reader = Me.cmdInExp.ExecuteReader
        While reader.Read
            Me.cboYear.Items.Add(IIf(DBNull.Value.Equals(reader!year), "", reader!year))
        End While
        reader.Close()
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub frmFinRptIncomeExpenditure_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint
        If Me.IsMdiChild Then
            Dim PnlLoc As New Point
            PnlLoc.X = CInt((Me.Width - Me.pnlIncomeExpense.Width) / 2)
            PnlLoc.Y = CInt((Me.Height - Me.pnlIncomeExpense.Height) / 2.5)
            Me.pnlIncomeExpense.Location = PnlLoc
        Else
            Me.pnlIncomeExpense.Dock = DockStyle.Fill
        End If
    End Sub

    Private Sub btnLoad_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLoad.Click
        If Me.rbTermly.Checked = False And Me.rbYearly.Checked = False Then
            MsgBox("Check for either Termly Or Yearly.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboYear.Text.Trim.Length <= 0 Then
            MsgBox("Year is missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
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
            Me.crtVwIncExp.ReportSource = Nothing
            Me.Cursor = Cursors.WaitCursor
            Me.cmdInExp.Connection = conn
            Me.cmdInExp.CommandType = CommandType.Text

            If Me.rbTermly.Checked = True Then

                Me.cmdInExp.CommandText = "SELECT * FROM vwFinIncExpStatement WHERE (year=@year) AND (termName=@termName)"
                Me.cmdInExp.Parameters.Clear()
                Me.cmdInExp.Parameters.AddWithValue("@year", Me.cboYear.Text.Trim)
                Me.cmdInExp.Parameters.AddWithValue("@termName", Me.cboTerm.Text.Trim)
                reader = Me.cmdInExp.ExecuteReader
                If reader.HasRows = True Then
                    Dim RptResultsView As New crtFinRptIncomeExpenseTermly
                    SetReportLogOn(RptResultsView)
                    RptResultsView.SummaryInfo.ReportTitle = "TERMLY INCOME EXPENDITURE STATEMENT FOR " & Me.cboTerm.Text.Trim & " " & _
                        Me.cboYear.Text.Trim

                    RptResultsView.SummaryInfo.ReportComments = fullName.Trim
                    RptResultsView.RecordSelectionFormula = "({vwFinIncExpStatement.year}=" & Me.cboYear.Text.Trim & ")"
                    RptResultsView.RecordSelectionFormula += "AND ({vwFinIncExpStatement.termName}=" & Chr(34) & Me.cboTerm.Text.Trim & Chr(34) & ")"
                    Me.crtVwIncExp.ReportSource = RptResultsView
                    Me.crtVwIncExp.Zoom(100)
                    Me.crtVwIncExp.RefreshReport()
                Else
                    MsgBox("No Record Found!", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "System Reply")
                End If
                reader.Close()

            ElseIf Me.rbYearly.Checked = True Then

                Me.cmdInExp.CommandText = "SELECT * FROM vwFinIncExpStatementYear WHERE (year=@year)"
                Me.cmdInExp.Parameters.Clear()
                Me.cmdInExp.Parameters.AddWithValue("@year", Me.cboYear.Text.Trim)
                reader = Me.cmdInExp.ExecuteReader
                If reader.HasRows = True Then
                    Dim RptResultsView As New crtFinRptIncomeExpenseYearly
                    SetReportLogOn(RptResultsView)
                    RptResultsView.SummaryInfo.ReportTitle = "YEARLY INCOME EXPENDITURE STATEMENT FOR  " & _
                        Me.cboYear.Text.Trim

                    RptResultsView.SummaryInfo.ReportComments = fullName.Trim
                    RptResultsView.RecordSelectionFormula = "({vwFinIncExpStatementYear.year}=" & Me.cboYear.Text.Trim & ")"
                    Me.crtVwIncExp.ReportSource = RptResultsView
                    Me.crtVwIncExp.Zoom(100)
                    Me.crtVwIncExp.RefreshReport()
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

    Private Sub cboTerm_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboTerm.SelectedIndexChanged
        Me.crtVwIncExp.ReportSource = Nothing
    End Sub

    Private Sub cboYear_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboYear.SelectedIndexChanged
        Me.crtVwIncExp.ReportSource = Nothing
    End Sub

    Private Sub rbTermly_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbTermly.CheckedChanged, rbYearly.CheckedChanged
        Me.crtVwIncExp.ReportSource = Nothing
        If Me.rbTermly.Checked = True Then
            Me.cboTerm.Text = ""
            Me.cboTerm.SelectedIndex = -1
            Me.cboTerm.Enabled = True

        ElseIf Me.rbYearly.Checked = True Then
            Me.cboTerm.Text = ""
            Me.cboTerm.SelectedIndex = -1
            Me.cboTerm.Enabled = False
        End If
    End Sub
End Class