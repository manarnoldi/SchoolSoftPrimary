Imports System.Data.SqlClient
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports System.Drawing.Printing
Public Class frmTTPrintMaster
    Dim reader As SqlDataReader
    Dim cmdPrintMaster As New SqlCommand
    Dim rec As Integer = 0
    Private Sub frmTTPrintMaster_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
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

        Me.cmdPrintMaster.Connection = conn
        Me.cmdPrintMaster.CommandType = CommandType.Text
        Me.cmdPrintMaster.CommandText = "SELECT DISTINCT year FROM tblSchoolCalendar WHERE (status=1) ORDER BY year"
        Me.cmdPrintMaster.Parameters.Clear()
        reader = Me.cmdPrintMaster.ExecuteReader
        While reader.Read
            Me.cboYear.Items.Add(IIf(DBNull.Value.Equals(reader!year), "", reader!year))
        End While
        reader.Close()
    End Sub
    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub btnPreview_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPreview.Click
        If Me.cboYear.Text.Trim.Count <= 0 Then
            MsgBox("Select year first.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            Me.Cursor = Cursors.WaitCursor
            Dim RptResultsView As New crtTTPrintMaster
            SetReportLogOn(RptResultsView)
            RptResultsView.SummaryInfo.ReportTitle = "SCHOOL MASTER TIMETABLE FOR YEAR " & Me.cboYear.Text.Trim
            RptResultsView.RecordSelectionFormula = "({vwTTPrintClass.year}=" & Me.cboYear.Text.Trim & ")"


            RptResultsView.RecordSelectionFormula = "(date({vwGood_receipt_note.DATE}) IN #" & Format(Me.cboYear.Text, "yyyy-MM-dd") & "# TO #" _
            & Format(Me.cboYear.Text, "yyyy-MM-dd") & "# "


            Me.crtVwPrintMasterCopy.ReportSource = RptResultsView
            Me.crtVwPrintMasterCopy.Zoom(100)
            Me.crtVwPrintMasterCopy.RefreshReport()
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
        If Me.cboYear.Text.Trim.Count <= 0 Then
            MsgBox("Select year first.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If
        Dim result As MsgBoxResult = MsgBox("Print Report?", MsgBoxStyle.Question + MsgBoxStyle.ApplicationModal + MsgBoxStyle.YesNo, "Confirm Transaction")
        If result = MsgBoxResult.No Then
            Exit Sub
        End If
        Me.Cursor = Cursors.WaitCursor
        Dim RptResultsView As New crtTTPrintMaster
        SetReportLogOn(RptResultsView)
        RptResultsView.SummaryInfo.ReportTitle = "SCHOOL MASTER TIMETABLE FOR YEAR " & Me.cboYear.Text.Trim
        RptResultsView.RecordSelectionFormula = "({vwTTPrintClass.year}=" & Me.cboYear.Text.Trim & ")"

        Me.crtVwPrintMasterCopy.ReportSource = RptResultsView
        Me.crtVwPrintMasterCopy.Zoom(100)
        Me.crtVwPrintMasterCopy.RefreshReport()
        RptResultsView.PrintToPrinter(1, True, 1, RptResultsView.FormatEngine.GetLastPageNumber(New CrystalDecisions.Shared.ReportPageRequestContext()))
    End Sub

    Private Sub frmTTPrintMaster_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint
        If Me.IsMdiChild Then
            Dim PnlLoc As New Point
            PnlLoc.X = CInt((Me.Width - Me.pnlPrintMaster.Width) / 2)
            PnlLoc.Y = CInt((Me.Height - Me.pnlPrintMaster.Height) / 2.5)
            Me.pnlPrintMaster.Location = PnlLoc
        Else
            Me.pnlPrintMaster.Dock = DockStyle.Fill
        End If
    End Sub
End Class