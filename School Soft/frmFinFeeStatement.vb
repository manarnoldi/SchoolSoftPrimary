Imports System.Data.SqlClient
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports System.Drawing.Printing
Public Class frmFinFeeStatement
    Dim reader As SqlDataReader
    Dim rec As Integer = 0
    Dim cmdStatement As New SqlCommand
    Private Sub frmFinFeeStatement_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
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
        Me.cmdStatement.CommandType = CommandType.Text
        Me.cmdStatement.Connection = conn
        Me.cmdStatement.CommandText = "SELECT DISTINCT className FROM  tblClasses WHERE (status=1) ORDER BY className"
        Me.cmdStatement.Parameters.Clear()
        reader = Me.cmdStatement.ExecuteReader
        While reader.Read
            Me.cboClass.Items.Add(IIf(DBNull.Value.Equals(reader!className), "", reader!className))
        End While
        reader.Close()

        Me.cmdStatement.CommandText = "SELECT DISTINCT stream FROM  tblClasses WHERE (status=1) ORDER BY stream"
        Me.cmdStatement.Parameters.Clear()
        reader = Me.cmdStatement.ExecuteReader
        While reader.Read
            Me.cboStream.Items.Add(IIf(DBNull.Value.Equals(reader!stream), "", reader!stream))
        End While
        reader.Close()

        Me.cmdStatement.CommandText = "SELECT DISTINCT year FROM  tblSchoolCalendar WHERE (status=1) ORDER BY year"
        Me.cmdStatement.Parameters.Clear()
        reader = Me.cmdStatement.ExecuteReader
        While reader.Read
            Me.cboYear.Items.Add(IIf(DBNull.Value.Equals(reader!year), "", reader!year))
        End While
        reader.Close()
    End Sub
    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub frmFinFeeStatement_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint
        If Me.IsMdiChild Then
            Dim PnlLoc As New Point
            PnlLoc.X = CInt((Me.Width - Me.pnlFeeStatement.Width) / 2)
            PnlLoc.Y = CInt((Me.Height - Me.pnlFeeStatement.Height) / 2.5)
            Me.pnlFeeStatement.Location = PnlLoc
        Else
            Me.pnlFeeStatement.Dock = DockStyle.Fill
        End If
    End Sub

    Private Sub btnLoad_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLoad.Click
        Me.lstFeeBalances.Items.Clear()
        If Me.cboClass.Text.Trim.Length <= 0 Then
            MsgBox("Class Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboStream.Text.Trim.Length <= 0 Then
            MsgBox("Stream Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboYear.Text.Trim.Length <= 0 Then
            MsgBox("Year Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            Me.cmdStatement.Connection = conn
            Me.cmdStatement.CommandType = CommandType.Text
            Me.cmdStatement.CommandText = "SELECT * FROM vwStudClass WHERE (classStudStatus=1) AND (studStatus=1) AND (classStatus=1) " & _
                vbNewLine & " AND (className=@className) AND (stream=@stream) AND (year=@year)"
            Me.cmdStatement.Parameters.Clear()
            Me.cmdStatement.Parameters.AddWithValue("@className", Me.cboClass.Text.Trim)
            Me.cmdStatement.Parameters.AddWithValue("@stream", Me.cboStream.Text.Trim)
            Me.cmdStatement.Parameters.AddWithValue("@year", Me.cboYear.Text.Trim)
            reader = Me.cmdStatement.ExecuteReader
            While reader.Read
                li = Me.lstFeeBalances.Items.Add(IIf(DBNull.Value.Equals(reader!admNo), "", reader!admNo))
                li.SubItems.Add(IIf(DBNull.Value.Equals(reader!FullName), "", reader!FullName))
            End While
            reader.Close()
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub

    Private Sub cboTerm_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboClass.SelectedIndexChanged, cboStream.SelectedIndexChanged, cboYear.SelectedIndexChanged
        Me.lstFeeBalances.Items.Clear()
    End Sub

    Private Sub lstFeeBalances_ColumnClick(ByVal sender As Object, ByVal e As System.Windows.Forms.ColumnClickEventArgs) Handles lstFeeBalances.ColumnClick
        If (e.Column() = 0) And (Me.lstFeeBalances.Items.Count > 0) Then
            For Each Li As ListViewItem In Me.lstFeeBalances.Items
                Li.Checked = Not (Li.Checked)
            Next
        End If
    End Sub

    Private Sub btnPreview_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPreview.Click
        If Me.lstFeeBalances.Items.Count <= 0 Then
            MsgBox("Load items in the list before preview.", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly + MsgBoxStyle.ApplicationModal, "System Reply")
            Exit Sub
        ElseIf Me.lstFeeBalances.CheckedItems.Count <= 0 Then
            MsgBox("Check the item to preview.", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly + MsgBoxStyle.ApplicationModal, "System Reply")
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            If Me.lstFeeBalances.CheckedItems.Count <> 1 Then
                MsgBox("Preview one item at a time.", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly + MsgBoxStyle.ApplicationModal, "System Reply")
                Exit Sub
            Else
                Me.Cursor = Cursors.WaitCursor
                Dim RptResultsView As New crtFinRptFeeStatement
                SetReportLogOn(RptResultsView)
                RptResultsView.SummaryInfo.ReportComments = fullName.Trim
                RptResultsView.RecordSelectionFormula = "({vwFinFeeStatement.admNo}=" & Me.lstFeeBalances.CheckedItems(0).Text.Trim & ")"
                RptResultsView.RecordSelectionFormula += "AND ({vwFinFeeStatement.year}=" & Me.cboYear.Text.Trim & ")"
                RptResultsView.RecordSelectionFormula += "AND ({vwFinFeeStatement.className}=" & Chr(34) & Me.cboClass.Text.Trim & Chr(34) & ")"
                RptResultsView.RecordSelectionFormula += "AND ({vwFinFeeStatement.stream}=" & Chr(34) & Me.cboStream.Text.Trim & Chr(34) & ")"

                frmResultsViewing.crtViewResultsSummary.ReportSource = RptResultsView
                frmResultsViewing.crtViewResultsSummary.Zoom(100)
                frmResultsViewing.crtViewResultsSummary.RefreshReport()
                frmResultsViewing.MdiParent = frmHome
                frmResultsViewing.Show()
                Me.Cursor = Cursors.Arrow
            End If
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub

    Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        If Me.lstFeeBalances.Items.Count <= 0 Then
            MsgBox("Load items in the list before printing.", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly + MsgBoxStyle.ApplicationModal, "System Reply")
            Exit Sub
        ElseIf Me.lstFeeBalances.CheckedItems.Count <= 0 Then
            MsgBox("Check the items to print.", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly + MsgBoxStyle.ApplicationModal, "System Reply")
            Exit Sub
        Else
            Dim result As MsgBoxResult = MsgBox("Print Seleced Items?", MsgBoxStyle.Question + MsgBoxStyle.ApplicationModal + MsgBoxStyle.YesNo, "Confirm Transaction")
            If result = MsgBoxResult.No Then
                Exit Sub
            End If
            Me.Cursor = Cursors.WaitCursor
            j = 1
            For i = 0 To Me.lstFeeBalances.CheckedItems.Count - 1
                Dim RptResultsView As New crtFinRptFeeStatement
                Dim ReportPrintOptions As PrintOptions = RptResultsView.PrintOptions
                Dim ReportPrinter As New PrinterSettings
                ReportPrintOptions.PrinterName = ReportPrinter.PrinterName
                ReportPrintOptions.PaperOrientation = PaperOrientation.DefaultPaperOrientation
                ReportPrintOptions.PaperSize = CrystalDecisions.Shared.PaperSize.PaperA4
                ReportPrintOptions.PrinterDuplex = PrinterDuplex.Default
                SetReportLogOn(RptResultsView)
                RptResultsView.RecordSelectionFormula = "({vwFinFeeStatement.admNo}=" & Me.lstFeeBalances.CheckedItems(0).Text.Trim & ")"
                RptResultsView.RecordSelectionFormula += "AND ({vwFinFeeStatement.year}=" & Me.cboYear.Text.Trim & ")"
                RptResultsView.RecordSelectionFormula += "AND ({vwFinFeeStatement.className}=" & Chr(34) & Me.cboClass.Text.Trim & Chr(34) & ")"
                RptResultsView.RecordSelectionFormula += "AND ({vwFinFeeStatement.stream}=" & Chr(34) & Me.cboStream.Text.Trim & Chr(34) & ")"

                RptResultsView.PrintToPrinter(1, True, 1, RptResultsView.FormatEngine.GetLastPageNumber(New CrystalDecisions.Shared.ReportPageRequestContext()))
                Me.ToolStripStatusLabel2.Text = "Printing...." & j & " Of " & Me.lstFeeBalances.CheckedItems.Count & " Reports."
                j = j + 1
            Next
            Me.ToolStripStatusLabel2.Text = "Printing Complete"
            Me.Cursor = Cursors.Arrow
        End If
        Me.lstFeeBalances.Items.Clear()
        Me.ToolStripStatusLabel2.Text = "Printing Status"
    End Sub
End Class