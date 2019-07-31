Imports System.Data.SqlClient
Public Class frmInvIssueRequestNote
    Dim reader As SqlDataReader
    Dim rec As Integer
    Dim cmdIssueReqNote As New SqlCommand
    Private Sub frmInvIssueRequestNote_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
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

        Me.cmdIssueReqNote.Connection = conn
        Me.cmdIssueReqNote.CommandType = CommandType.Text
        Me.cmdIssueReqNote.CommandText = "SELECT DISTINCT year FROM tblSchoolCalendar WHERE (status=1) ORDER BY year"
        Me.cmdIssueReqNote.Parameters.Clear()
        reader = Me.cmdIssueReqNote.ExecuteReader
        While reader.Read
            Me.cboYear.Items.Add(IIf(DBNull.Value.Equals(reader!year), "", reader!year))
        End While
        reader.Close()
    End Sub
    Private Function monthToInt(ByVal monthName As String)
        Dim monthId As Integer = 0
        If monthName = "January" Then
            monthId = 1
        ElseIf monthName = "February" Then
            monthId = 2
        ElseIf monthName = "March" Then
            monthId = 3
        ElseIf monthName = "April" Then
            monthId = 4
        ElseIf monthName = "May" Then
            monthId = 5
        ElseIf monthName = "June" Then
            monthId = 6
        ElseIf monthName = "July" Then
            monthId = 7
        ElseIf monthName = "August" Then
            monthId = 8
        ElseIf monthName = "September" Then
            monthId = 9
        ElseIf monthName = "October" Then
            monthId = 10
        ElseIf monthName = "November" Then
            monthId = 11
        ElseIf monthName = "December" Then
            monthId = 12
        End If
        Return monthId
    End Function
    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub frmInvIssueRequestNote_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint
        If Me.IsMdiChild Then
            Dim PnlLoc As New Point
            PnlLoc.X = CInt((Me.Width - Me.pnlIssueReqNote.Width) / 2)
            PnlLoc.Y = CInt((Me.Height - Me.pnlIssueReqNote.Height) / 2.5)
            Me.pnlIssueReqNote.Location = PnlLoc
        Else
            Me.pnlIssueReqNote.Dock = DockStyle.Fill
        End If
    End Sub

    Private Sub cboYear_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboYear.SelectedIndexChanged, cboMonth.SelectedIndexChanged
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            If Me.cboMonth.Text.Trim.Length > 0 And Me.cboYear.Text.Trim.Length > 0 Then

                Me.cboIssueNoteNo.Items.Clear()
                Me.cboIssueNoteNo.Text = ""
                Me.cboIssueNoteNo.SelectedIndex = -1

                Me.cmdIssueReqNote.Connection = conn
                Me.cmdIssueReqNote.CommandType = CommandType.Text
                Me.cmdIssueReqNote.CommandText = "SELECT DISTINCT issueNumber FROM tblInvIssues WHERE (month=@month) AND (year=@year) " & _
                    vbNewLine & " ORDER BY issueNumber"
                Me.cmdIssueReqNote.Parameters.Clear()
                Me.cmdIssueReqNote.Parameters.AddWithValue("@month", monthToInt(Me.cboMonth.Text.Trim))
                Me.cmdIssueReqNote.Parameters.AddWithValue("@year", Me.cboYear.Text.Trim)
                reader = Me.cmdIssueReqNote.ExecuteReader
                If reader.HasRows = True Then
                    While reader.Read
                        Me.cboIssueNoteNo.Items.Add(IIf(DBNull.Value.Equals(reader!issueNumber), "", reader!issueNumber))
                    End While
                ElseIf reader.HasRows = False Then
                    MsgBox("No Issue Note Number Found For the Month and Year.", MsgBoxStyle.Information + _
                       MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "System Reply")
                End If
                reader.Close()
            End If
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub

    Private Sub btnLoad_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLoad.Click
        If Me.cboMonth.Text.Trim.Length <= 0 Then
            MsgBox("Month is Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboIssueNoteNo.Text.Trim.Length <= 0 Then
            MsgBox("Issue Note Number Is Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
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

            Me.crtVwIssueReqNote.ReportSource = Nothing

            Me.Cursor = Cursors.WaitCursor

            Dim RptResultsView As New crtInvIssueNote
            SetReportLogOn(RptResultsView)
            RptResultsView.SummaryInfo.ReportComments = fullName.Trim
            RptResultsView.RecordSelectionFormula = "({vwInvIssueRequest.issueNumber}=" & Chr(34) & Me.cboIssueNoteNo.Text.Trim & Chr(34) & ")"
            RptResultsView.RecordSelectionFormula += "AND ({vwInvIssueRequest.month}=" & monthToInt(Me.cboMonth.Text.Trim) & ")"
            RptResultsView.RecordSelectionFormula += "AND ({vwInvIssueRequest.year}=" & Me.cboYear.Text.Trim & ")"

            Me.crtVwIssueReqNote.ReportSource = RptResultsView
            Me.crtVwIssueReqNote.Zoom(100)
            Me.crtVwIssueReqNote.RefreshReport()

            'loadCombos()
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