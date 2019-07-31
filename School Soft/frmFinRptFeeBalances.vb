Imports System.Data.SqlClient
Public Class frmFinRptFeeBalances
    Dim reader As sqldatareader
    Dim cmdFeeBalance As New sqlcommand
    Dim rec As Integer = 0
    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub frmFinRptFeeBalances_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
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

    Private Sub frmFinRptFeeBalances_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint
        If Me.IsMdiChild Then
            Dim PnlLoc As New Point
            PnlLoc.X = CInt((Me.Width - Me.pnlFeeBalances.Width) / 2)
            PnlLoc.Y = CInt((Me.Height - Me.pnlFeeBalances.Height) / 2.5)
            Me.pnlFeeBalances.Location = PnlLoc
        Else
            Me.pnlFeeBalances.Dock = DockStyle.Fill
        End If
    End Sub
    Private Sub loadCombos()
        Me.cboClass.Items.Clear()
        Me.cboClass.Text = ""
        Me.cboClass.SelectedIndex = -1

        Me.cmdFeeBalance.CommandType = CommandType.Text
        Me.cmdFeeBalance.Connection = conn
        Me.cmdFeeBalance.CommandText = "SELECT DISTINCT termName FROM tblSchoolCalendar WHERE (status=1) ORDER BY termName"
        Me.cmdFeeBalance.Parameters.Clear()
        reader = Me.cmdFeeBalance.ExecuteReader
        While reader.Read
            Me.cboTerm.Items.Add(IIf(DBNull.Value.Equals(reader!termName), "", reader!termName))
        End While
        reader.Close()

        Me.cmdFeeBalance.CommandText = "SELECT DISTINCT className FROM  tblClasses WHERE (status=1) ORDER BY className"
        Me.cmdFeeBalance.Parameters.Clear()
        reader = Me.cmdFeeBalance.ExecuteReader
        While reader.Read
            Me.cboClass.Items.Add(IIf(DBNull.Value.Equals(reader!className), "", reader!className))
        End While
        reader.Close()

        Me.cmdFeeBalance.CommandText = "SELECT DISTINCT year FROM  tblSchoolCalendar WHERE (status=1) ORDER BY year"
        Me.cmdFeeBalance.Parameters.Clear()
        reader = Me.cmdFeeBalance.ExecuteReader
        While reader.Read
            Me.cboYear.Items.Add(IIf(DBNull.Value.Equals(reader!year), "", reader!year))
        End While
        reader.Close()
    End Sub

    Private Sub cboReportType_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboReportType.SelectedIndexChanged
        Me.crtVwFinFeeBalRpt.ReportSource = Nothing
        If Me.cboReportType.Text = "Detailed" Then
            Me.cboClass.Text = ""
            Me.cboClass.SelectedIndex = -1
            Me.cboClass.Enabled = True
        ElseIf Me.cboReportType.Text = "Summary" Then
            Me.cboClass.Text = ""
            Me.cboClass.SelectedIndex = -1
            Me.cboClass.Enabled = False
        End If
    End Sub

    Private Sub btnLoad_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLoad.Click
        Me.crtVwFinFeeBalRpt.ReportSource = Nothing
        If Me.cboReportType.Text = "Detailed" Then
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
            Me.crtVwFinFeeBalRpt.ReportSource = Nothing
            Me.Cursor = Cursors.WaitCursor
            Me.cmdFeeBalance.Connection = conn
            Me.cmdFeeBalance.CommandType = CommandType.Text

            If Me.cboReportType.Text.Trim = "Detailed" Then

                Me.cmdFeeBalance.CommandText = "SELECT * FROM vwFinFeeStudBalance WHERE (termName=@termName) AND (year=@year) AND " & _
                    vbNewLine & " (className=@className)"
                Me.cmdFeeBalance.Parameters.Clear()
                Me.cmdFeeBalance.Parameters.AddWithValue("@termName", Me.cboTerm.Text.Trim)
                Me.cmdFeeBalance.Parameters.AddWithValue("@year", Me.cboYear.Text.Trim)
                Me.cmdFeeBalance.Parameters.AddWithValue("@className", Me.cboClass.Text.Trim)
                reader = Me.cmdFeeBalance.ExecuteReader
                If reader.HasRows = True Then
                    Dim RptResultsView As New crtFinRptFeeBalancesDetailed
                    SetReportLogOn(RptResultsView)
                    RptResultsView.SummaryInfo.ReportTitle = "FEE BALANCE REPORT FOR " & Me.cboTerm.Text.Trim & " " & _
                        Me.cboClass.Text.Trim & " " & Me.cboYear.Text.Trim
                    RptResultsView.SummaryInfo.ReportComments = fullName.Trim
                    RptResultsView.RecordSelectionFormula = "({vwFinFeeStudBalance.year}=" & Me.cboYear.Text.Trim & ")"
                    RptResultsView.RecordSelectionFormula += "AND ({vwFinFeeStudBalance.termName}=" & Chr(34) & Me.cboTerm.Text.Trim & Chr(34) & ")"
                    RptResultsView.RecordSelectionFormula += "AND ({vwFinFeeStudBalance.className}=" & Chr(34) & Me.cboClass.Text.Trim & Chr(34) & ")"
                    Me.crtVwFinFeeBalRpt.ReportSource = RptResultsView
                    Me.crtVwFinFeeBalRpt.Zoom(100)
                    Me.crtVwFinFeeBalRpt.RefreshReport()
                Else
                    MsgBox("No Record Found!", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "System Reply")
                End If
                reader.Close()

            ElseIf Me.cboReportType.Text.Trim = "Summary" Then

                Me.cmdFeeBalance.CommandText = "SELECT * FROM vwFinFeeStudBalanceSummary WHERE (year=@year) AND (termName=@termName)"
                Me.cmdFeeBalance.Parameters.Clear()
                Me.cmdFeeBalance.Parameters.AddWithValue("@year", Me.cboYear.Text.Trim)
                Me.cmdFeeBalance.Parameters.AddWithValue("@termName", Me.cboTerm.Text.Trim)
                reader = Me.cmdFeeBalance.ExecuteReader
                If reader.HasRows = True Then
                    Dim RptResultsView As New crtFinRptFeeBalancesSummary
                    SetReportLogOn(RptResultsView)
                    RptResultsView.SummaryInfo.ReportTitle = "FEE BALANCE REPORT FOR " & Me.cboTerm.Text.Trim & "" & Me.cboYear.Text.Trim 
                    RptResultsView.SummaryInfo.ReportComments = fullName.Trim
                    RptResultsView.RecordSelectionFormula = "({vwFinFeeStudBalanceSummary.year}=" & Me.cboYear.Text.Trim & ")"
                    RptResultsView.RecordSelectionFormula += "AND ({vwFinFeeStudBalanceSummary.termName}=" & Chr(34) & Me.cboTerm.Text.Trim & Chr(34) & ")"
                    Me.crtVwFinFeeBalRpt.ReportSource = RptResultsView
                    Me.crtVwFinFeeBalRpt.Zoom(100)
                    Me.crtVwFinFeeBalRpt.RefreshReport()
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
End Class