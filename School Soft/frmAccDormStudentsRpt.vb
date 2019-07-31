Imports System.Data.SqlClient
Public Class frmAccDormStudentsRpt
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
        Me.cboDormName.Items.Clear()
        Me.cboDormName.Text = ""
        Me.cboDormName.SelectedIndex = -1

        Me.cmdConsolidated.Connection = conn
        Me.cmdConsolidated.CommandType = CommandType.Text
        Me.cmdConsolidated.CommandText = "SELECT DISTINCT dormName FROM tblAccdormRegister ORDER BY dormName"
        Me.cmdConsolidated.Parameters.Clear()
        reader = Me.cmdConsolidated.ExecuteReader
        Me.cboDormName.Items.Add("ALL")
        While reader.Read
            Me.cboDormName.Items.Add(IIf(DBNull.Value.Equals(reader!dormName), "", reader!dormName))
        End While
        reader.Close()
    End Sub

    Private Sub frmInvRptConsolidated_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint
        If Me.IsMdiChild Then
            Dim PnlLoc As New Point
            PnlLoc.X = CInt((Me.Width - Me.pnlAccDormStudentsRpt.Width) / 2)
            PnlLoc.Y = CInt((Me.Height - Me.pnlAccDormStudentsRpt.Height) / 2.5)
            Me.pnlAccDormStudentsRpt.Location = PnlLoc
        Else
            Me.pnlAccDormStudentsRpt.Dock = DockStyle.Fill
        End If
    End Sub

    Private Sub btnLoad_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLoad.Click
        Me.crtVwRptDormStud.ReportSource = Nothing
        If Me.cboDormName.Text.Trim.Length <= 0 Then
            MsgBox("Dorm Name Is Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            Me.cmdConsolidated.Connection = conn
            Me.cmdConsolidated.CommandType = CommandType.Text

            If Me.cboDormName.Text.Trim = "ALL" Then
                Me.cmdConsolidated.CommandText = "SELECT * FROM vwAccDormStudentDetails"
                Me.cmdConsolidated.Parameters.Clear()
                reader = Me.cmdConsolidated.ExecuteReader
                If reader.HasRows = True Then
                    Me.Cursor = Cursors.WaitCursor
                    Dim RptResultsView As New crtAccDormStudentsRpt
                    SetReportLogOn(RptResultsView)
                    RptResultsView.SummaryInfo.ReportTitle = "STUDENT DORMITORY REPORT FOR DORM " & Me.cboDormName.Text.Trim
                    RptResultsView.SummaryInfo.ReportComments = fullName.Trim
                    Me.crtVwRptDormStud.ReportSource = RptResultsView
                    Me.crtVwRptDormStud.Zoom(100)
                    Me.crtVwRptDormStud.RefreshReport()
                Else
                    MsgBox("No Record Found!", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "System Reply")
                End If
                reader.Close()
            Else
                Me.cmdConsolidated.CommandText = "SELECT * FROM vwAccDormStudentDetails WHERE (dormName=@dormName)"
                Me.cmdConsolidated.Parameters.Clear()
                Me.cmdConsolidated.Parameters.AddWithValue("@dormName", Me.cboDormName.Text.Trim)
                reader = Me.cmdConsolidated.ExecuteReader
                If reader.HasRows = True Then
                    Me.Cursor = Cursors.WaitCursor
                    Dim RptResultsView As New crtAccDormStudentsRpt
                    SetReportLogOn(RptResultsView)
                    RptResultsView.SummaryInfo.ReportTitle = "STUDENT DORMITORY REPORT FOR DORM " & Me.cboDormName.Text.Trim
                    RptResultsView.SummaryInfo.ReportComments = fullName.Trim
                    RptResultsView.RecordSelectionFormula = "({vwAccDormStudentDetails.dormName}=" & Chr(34) & Me.cboDormName.Text.Trim & Chr(34) & ")"
                    Me.crtVwRptDormStud.ReportSource = RptResultsView
                    Me.crtVwRptDormStud.Zoom(100)
                    Me.crtVwRptDormStud.RefreshReport()
                Else
                    MsgBox("No Record Found!", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "System Reply")
                End If
                reader.Close()
            End If
            

        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            Me.Cursor = Cursors.Arrow
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub

    Private Sub cboYear_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboDormName.SelectedIndexChanged
        Me.crtVwRptDormStud.ReportSource = Nothing
    End Sub
End Class