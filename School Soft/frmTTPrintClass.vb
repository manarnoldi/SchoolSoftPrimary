Imports System.Data.SqlClient
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports System.Drawing.Printing
Public Class frmTTPrintClass
    Dim reader As SqlDataReader
    Dim cmdPrintTTClass As New SqlCommand
    Dim rec As Integer = 0
    Private Sub frmTTPrintClass_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
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

    Private Sub frmTTPrintClass_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint
        If Me.IsMdiChild Then
            Dim PnlLoc As New Point
            PnlLoc.X = CInt((Me.Width - Me.pnlPrintClass.Width) / 2)
            PnlLoc.Y = CInt((Me.Height - Me.pnlPrintClass.Height) / 2.5)
            Me.pnlPrintClass.Location = PnlLoc
        Else
            Me.pnlPrintClass.Dock = DockStyle.Fill
        End If
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub
    Private Sub loadCombos()
        Me.cboYear.Items.Clear()
        Me.cboYear.Text = ""
        Me.cboYear.SelectedIndex = -1

        Me.cmdPrintTTClass.Connection = conn
        Me.cmdPrintTTClass.CommandType = CommandType.Text
        Me.cmdPrintTTClass.CommandText = "SELECT DISTINCT year FROM tblSchoolCalendar WHERE (status=1) ORDER BY year"
        Me.cmdPrintTTClass.Parameters.Clear()
        reader = Me.cmdPrintTTClass.ExecuteReader
        While reader.Read
            Me.cboYear.Items.Add(IIf(DBNull.Value.Equals(reader!year), "", reader!year))
        End While
        reader.Close()
    End Sub

    Private Sub btnLoad_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLoad.Click
        If Me.cboYear.Text.Trim.Count <= 0 Then
            MsgBox("Select year first.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If
        Me.lstTTClasses.Items.Clear()
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            Me.cmdPrintTTClass.Connection = conn
            Me.cmdPrintTTClass.CommandType = CommandType.Text
            Me.cmdPrintTTClass.CommandText = "SELECT * FROM tblClasses WHERE (status=1) AND (year=@year) ORDER BY year,className,stream"
            Me.cmdPrintTTClass.Parameters.Clear()
            Me.cmdPrintTTClass.Parameters.AddWithValue("@year", Me.cboYear.Text.Trim)
            reader = Me.cmdPrintTTClass.ExecuteReader
            While reader.Read
                li = Me.lstTTClasses.Items.Add(IIf(DBNull.Value.Equals(reader!className), "", reader!className))
                li.SubItems.Add(IIf(DBNull.Value.Equals(reader!stream), "", reader!stream))
                li.SubItems.Add(IIf(DBNull.Value.Equals(reader!year), "", reader!year))
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

    Private Sub txtSeachClass_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtSeachClass.TextChanged
        Me.lstTTClasses.Items.Clear()
        If Me.cboYear.Text.Trim.Count <= 0 Then
            MsgBox("Select year first.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            If Me.txtSeachClass.Text.Length < 5 Then
                Exit Sub
            End If
            Me.cmdPrintTTClass.Connection = conn
            Me.cmdPrintTTClass.CommandType = CommandType.Text
            Me.cmdPrintTTClass.CommandText = "SELECT * FROM tblClasses WHERE (status=1) AND (year=@year) AND " & _
                vbNewLine & " (className LIKE @className) ORDER BY year,className,stream"
            Me.cmdPrintTTClass.Parameters.Clear()
            Me.cmdPrintTTClass.Parameters.AddWithValue("@year", Me.cboYear.Text.Trim)
            Me.cmdPrintTTClass.Parameters.AddWithValue("@className", String.Format("%{0}%", TryCast(Me.txtSeachClass.Text.Trim, String).Trim))
            reader = Me.cmdPrintTTClass.ExecuteReader
            While reader.Read
                li = Me.lstTTClasses.Items.Add(IIf(DBNull.Value.Equals(reader!className), "", reader!className))
                li.SubItems.Add(IIf(DBNull.Value.Equals(reader!stream), "", reader!stream))
                li.SubItems.Add(IIf(DBNull.Value.Equals(reader!year), "", reader!year))
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

    Private Sub btnPreview_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPreview.Click
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            If Me.lstTTClasses.CheckedItems.Count <> 1 Then
                MsgBox("Preview one item at a time.", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly + MsgBoxStyle.ApplicationModal, "System Reply")
                Exit Sub
            Else
                Me.Cursor = Cursors.WaitCursor
                Dim RptResultsView As New crtTTPrintClass
                SetReportLogOn(RptResultsView)
                RptResultsView.SummaryInfo.ReportTitle = "SCHOOL TIME TABLE FOR " & Me.lstTTClasses.CheckedItems(0).Text.Trim & " " & _
                    Me.lstTTClasses.CheckedItems(0).SubItems(1).Text.Trim & " " & Me.cboYear.Text.Trim
                RptResultsView.RecordSelectionFormula = "({vwTTPrintClass.className}=" & Chr(34) & Me.lstTTClasses.CheckedItems(0).Text.Trim & Chr(34) & ")"
                RptResultsView.RecordSelectionFormula += "AND ({vwTTPrintClass.year}=" & Me.cboYear.Text.Trim & ")"
                RptResultsView.RecordSelectionFormula += "AND ({vwTTPrintClass.stream}=" & Chr(34) & Me.lstTTClasses.CheckedItems(0).SubItems(1).Text.Trim & Chr(34) & ")"
                
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

    Private Sub cboYear_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboYear.SelectedIndexChanged
        Me.lstTTClasses.Items.Clear()
    End Sub

    Private Sub lstTTClasses_ColumnClick(ByVal sender As Object, ByVal e As System.Windows.Forms.ColumnClickEventArgs) Handles lstTTClasses.ColumnClick
        If (e.Column() = 0) And (Me.lstTTClasses.Items.Count > 0) Then
            For Each Li As ListViewItem In Me.lstTTClasses.Items
                Li.Checked = Not (Li.Checked)
            Next
        End If
    End Sub

    Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        If Me.lstTTClasses.Items.Count <= 0 Then
            MsgBox("Load items in the list before printing.", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly + MsgBoxStyle.ApplicationModal, "System Reply")
            Exit Sub
        ElseIf Me.lstTTClasses.CheckedItems.Count <= 0 Then
            MsgBox("Check the items to print.", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly + MsgBoxStyle.ApplicationModal, "System Reply")
            Exit Sub
        Else
            Dim result As MsgBoxResult = MsgBox("Print Seleced Items?", MsgBoxStyle.Question + MsgBoxStyle.ApplicationModal + MsgBoxStyle.YesNo, "Confirm Transaction")
            If result = MsgBoxResult.No Then
                Exit Sub
            End If
            Me.Cursor = Cursors.WaitCursor
            j = 1
            For i = 0 To Me.lstTTClasses.CheckedItems.Count - 1
                Dim RptResultsView As New crtTTPrintClass
                SetReportLogOn(RptResultsView)
                RptResultsView.SummaryInfo.ReportTitle = "SCHOOL TIME TABLE FOR " & Me.lstTTClasses.CheckedItems(0).Text.Trim & " " & _
                    Me.lstTTClasses.CheckedItems(0).SubItems(1).Text.Trim & " " & Me.cboYear.Text.Trim
                RptResultsView.RecordSelectionFormula = "({vwTTPrintClass.className}=" & Chr(34) & Me.lstTTClasses.CheckedItems(0).Text.Trim & Chr(34) & ")"
                RptResultsView.RecordSelectionFormula += "AND ({vwTTPrintClass.year}=" & Me.cboYear.Text.Trim & ")"
                RptResultsView.RecordSelectionFormula += "AND ({vwTTPrintClass.stream}=" & Chr(34) & Me.lstTTClasses.CheckedItems(0).SubItems(1).Text.Trim & Chr(34) & ")"

                RptResultsView.PrintToPrinter(1, True, 1, RptResultsView.FormatEngine.GetLastPageNumber(New CrystalDecisions.Shared.ReportPageRequestContext()))
                Me.ToolStripStatusLabel2.Text = "Printing...." & j & " Of " & Me.lstTTClasses.CheckedItems.Count & " Reports."
                j = j + 1
            Next
            Me.ToolStripStatusLabel2.Text = "Printing Complete"
            Me.Cursor = Cursors.Arrow
        End If
        Me.lstTTClasses.Items.Clear()
        Me.ToolStripStatusLabel2.Text = "Printing Status"
    End Sub
End Class