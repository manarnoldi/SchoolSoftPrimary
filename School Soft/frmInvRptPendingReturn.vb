Imports System.Data.SqlClient
Public Class frmInvRptPendingReturn
    Dim rec As Integer = 0
    Dim cmdpendingReturn As New SqlCommand
    Dim reader As SqlDataReader

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub frmInvRptPendingReturn_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
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

        Me.cmdpendingReturn.Connection = conn
        Me.cmdpendingReturn.CommandType = CommandType.Text
        Me.cmdpendingReturn.CommandText = "SELECT DISTINCT year FROM tblSchoolCalendar WHERE (status=1) ORDER BY year"
        Me.cmdpendingReturn.Parameters.Clear()
        reader = Me.cmdpendingReturn.ExecuteReader
        While reader.Read
            Me.cboYear.Items.Add(IIf(DBNull.Value.Equals(reader!year), "", reader!year))
        End While
        reader.Close()
    End Sub

    Private Sub frmInvRptPendingReturn_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint
        If Me.IsMdiChild Then
            Dim PnlLoc As New Point
            PnlLoc.X = CInt((Me.Width - Me.pnlInvPendRetRpt.Width) / 2)
            PnlLoc.Y = CInt((Me.Height - Me.pnlInvPendRetRpt.Height) / 2.5)
            Me.pnlInvPendRetRpt.Location = PnlLoc
        Else
            Me.pnlInvPendRetRpt.Dock = DockStyle.Fill
        End If
    End Sub

    Private Sub cboMonth_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboMonth.SelectedIndexChanged
        Me.crtVwRptInvIssues.ReportSource = Nothing
    End Sub

    Private Sub cboYear_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboYear.SelectedIndexChanged
        Me.crtVwRptInvIssues.ReportSource = Nothing
    End Sub

    Private Sub btnLoad_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLoad.Click
        Me.crtVwRptInvIssues.ReportSource = Nothing
        If Me.cboMonth.Text.Trim.Length <= 0 Then
            MsgBox("Month is Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
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
            Me.cmdpendingReturn.Connection = conn
            Me.cmdpendingReturn.CommandType = CommandType.Text
            Me.cmdpendingReturn.CommandText = "SELECT * FROM vwInvPendingReturnRpt WHERE (month=@month) AND (year=@year)"
            Me.cmdpendingReturn.Parameters.Clear()
            Me.cmdpendingReturn.Parameters.AddWithValue("@month", monthToInt(Me.cboMonth.Text.Trim))
            Me.cmdpendingReturn.Parameters.AddWithValue("@year", Me.cboYear.Text.Trim)
            reader = Me.cmdpendingReturn.ExecuteReader
            If reader.HasRows = True Then
                Dim RptResultsView As New crtInvRptPendingReturn
                SetReportLogOn(RptResultsView)
                RptResultsView.SummaryInfo.ReportTitle = "INVENTORY PENDING RETURN REPORT FOR " & Me.cboMonth.Text.Trim.ToUpper & " " & _
                    Me.cboYear.Text.Trim

                RptResultsView.SummaryInfo.ReportComments = fullName.Trim
                RptResultsView.RecordSelectionFormula = "({vwInvPendingReturnRpt.month}=" & monthToInt(Me.cboMonth.Text.Trim) & ")"
                RptResultsView.RecordSelectionFormula += " AND ({vwInvPendingReturnRpt.year}=" & Me.cboYear.Text.Trim & ")"
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
    Private Function monthToInt(ByVal monthName As String)
        Dim monthInt As Integer = 0
        If monthName = "January" Then
            monthInt = 1
        ElseIf monthName = "February" Then
            monthInt = 2
        ElseIf monthName = "March" Then
            monthInt = 3
        ElseIf monthName = "April" Then
            monthInt = 4
        ElseIf monthName = "May" Then
            monthInt = 5
        ElseIf monthName = "June" Then
            monthInt = 6
        ElseIf monthName = "July" Then
            monthInt = 7
        ElseIf monthName = "August" Then
            monthInt = 8
        ElseIf monthName = "September" Then
            monthInt = 9
        ElseIf monthName = "October" Then
            monthInt = 10
        ElseIf monthName = "November" Then
            monthInt = 11
        ElseIf monthName = "December" Then
            monthInt = 12
        End If
        Return monthInt
    End Function
End Class