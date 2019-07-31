Imports System.Data.SqlClient
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports System.Drawing.Printing
Public Class frmRptClassLists
    Dim reader As SqlDataReader
    Dim rec As Integer = 0
    Dim cmdClassList As New SqlCommand
    Private Sub frmRptClassLists_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
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
        Me.cboClassName.Items.Clear()
        Me.cboStreamName.Items.Clear()
        Me.cboYear.Items.Clear()

        Me.cboClassName.SelectedIndex = -1
        Me.cboStreamName.SelectedIndex = -1
        Me.cboYear.SelectedIndex = -1

        Me.cboClassName.Text = ""
        Me.cboStreamName.Text = ""
        Me.cboYear.Text = ""

        Me.cmdClassList.Connection = conn
        Me.cmdClassList.CommandType = CommandType.Text
        Me.cmdClassList.CommandText = "SELECT DISTINCT className FROM tblClasses WHERE (status=1) ORDER BY className"
        Me.cmdClassList.Parameters.Clear()
        reader = Me.cmdClassList.ExecuteReader
        While reader.Read
            Me.cboClassName.Items.Add(IIf(DBNull.Value.Equals(reader!className), "", reader!className))
        End While
        reader.Close()

        Me.cmdClassList.CommandText = "SELECT DISTINCT stream FROM tblClasses WHERE (status=1) ORDER BY stream"
        Me.cmdClassList.Parameters.Clear()
        reader = Me.cmdClassList.ExecuteReader
        While reader.Read
            Me.cboStreamName.Items.Add(IIf(DBNull.Value.Equals(reader!stream), "", reader!stream))
        End While
        reader.Close()

        Me.cmdClassList.CommandText = "SELECT DISTINCT year FROM tblClasses WHERE (status=1) ORDER BY year"
        Me.cmdClassList.Parameters.Clear()
        reader = Me.cmdClassList.ExecuteReader
        While reader.Read
            Me.cboYear.Items.Add(IIf(DBNull.Value.Equals(reader!year), "", reader!year))
        End While
        reader.Close()
    End Sub
    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub frmRptClassLists_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint
        If Me.IsMdiChild Then
            Dim PnlLoc As New Point
            PnlLoc.X = CInt((Me.Width - Me.pnlClassLists.Width) / 2)
            PnlLoc.Y = CInt((Me.Height - Me.pnlClassLists.Height) / 2.5)
            Me.pnlClassLists.Location = PnlLoc
        Else
            Me.pnlClassLists.Dock = DockStyle.Fill
        End If
    End Sub

    Private Sub btnLoad_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLoad.Click
        If Me.cboYear.Text.Trim.Length <= 0 Then
            MsgBox("Year is Missing", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly + MsgBoxStyle.ApplicationModal, "Missing Details")
            Exit Sub
        ElseIf Me.cboClassName.Text.Trim.Length <= 0 Then
            MsgBox("Class Name is Missing", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly + MsgBoxStyle.ApplicationModal, "Missing Details")
            Exit Sub
        ElseIf Me.cboStreamName.Text.Trim.Length <= 0 Then
            MsgBox("Stream is Missing", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly + MsgBoxStyle.ApplicationModal, "Missing Details")
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()

            Me.Cursor = Cursors.WaitCursor
            Dim RptResultsView As New crtRptClassLists
            SetReportLogOn(RptResultsView)
            RptResultsView.SummaryInfo.ReportTitle = ("CLASS LIST FOR " & Me.cboClassName.Text.Trim & " " & _
                Me.cboStreamName.Text.Trim & " YEAR " & Me.cboYear.Text.Trim & "").ToUpper
            RptResultsView.SummaryInfo.ReportComments = fullName.Trim
            RptResultsView.RecordSelectionFormula = "({vwRptClassLists.stream}=" & Chr(34) & Me.cboStreamName.Text.Trim & Chr(34) & ")"
            RptResultsView.RecordSelectionFormula += "AND ({vwRptClassLists.year}=" & Me.cboYear.Text.Trim & ")"
            RptResultsView.RecordSelectionFormula += "AND ({vwRptClassLists.className}=" & Chr(34) & Me.cboClassName.Text.Trim & Chr(34) & ")"
            Me.crtVwClassLists.ReportSource = RptResultsView
            Me.crtVwClassLists.Zoom(100)
            Me.crtVwClassLists.RefreshReport()
            Me.Cursor = Cursors.Arrow
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub

    Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        If Me.cboYear.Text.Trim.Length <= 0 Then
            MsgBox("Year is Missing", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly + MsgBoxStyle.ApplicationModal, "Missing Details")
            Exit Sub
        ElseIf Me.cboClassName.Text.Trim.Length <= 0 Then
            MsgBox("Class Name is Missing", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly + MsgBoxStyle.ApplicationModal, "Missing Details")
            Exit Sub
        ElseIf Me.cboStreamName.Text.Trim.Length <= 0 Then
            MsgBox("Stream is Missing", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly + MsgBoxStyle.ApplicationModal, "Missing Details")
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()

            Me.Cursor = Cursors.WaitCursor
            Dim RptResultsView As New crtRptClassLists
            SetReportLogOn(RptResultsView)
            RptResultsView.SummaryInfo.ReportTitle = ("CLASS LIST FOR " & Me.cboClassName.Text.Trim & " " & _
                Me.cboStreamName.Text.Trim & " YEAR " & Me.cboYear.Text.Trim & "").ToUpper
            RptResultsView.SummaryInfo.ReportComments = fullName.Trim
            RptResultsView.RecordSelectionFormula = "({vwRptClassLists.stream}=" & Chr(34) & Me.cboStreamName.Text.Trim & Chr(34) & ")"
            RptResultsView.RecordSelectionFormula += "AND ({vwRptClassLists.year}=" & Me.cboYear.Text.Trim & ")"
            RptResultsView.RecordSelectionFormula += "AND ({vwRptClassLists.className}=" & Chr(34) & Me.cboClassName.Text.Trim & Chr(34) & ")"
            RptResultsView.PrintToPrinter(1, True, 1, RptResultsView.FormatEngine.GetLastPageNumber(New CrystalDecisions.Shared.ReportPageRequestContext()))
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub
End Class