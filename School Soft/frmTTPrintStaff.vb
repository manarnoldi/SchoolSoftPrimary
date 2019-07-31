Imports System.Data.SqlClient
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports System.Drawing.Printing
Public Class frmTTPrintStaff
    Dim cmdPrintStaff As New sqlcommand
    Dim rec As Integer = 0
    Dim reader As SqlDataReader
    Private Sub frmTTPrintStaff_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint
        If Me.IsMdiChild Then
            Dim PnlLoc As New Point
            PnlLoc.X = CInt((Me.Width - Me.pnlPrintStaff.Width) / 2)
            PnlLoc.Y = CInt((Me.Height - Me.pnlPrintStaff.Height) / 2.5)
            Me.pnlPrintStaff.Location = PnlLoc
        Else
            Me.pnlPrintStaff.Dock = DockStyle.Fill
        End If
    End Sub

    Private Sub txtSearchTeacher_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtSearchTeacher.TextChanged
        Me.lstTTTeacher.Items.Clear()
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            If Me.txtSearchTeacher.Text.Length < 1 Then
                Exit Sub
            End If
            Me.cmdPrintStaff.Connection = conn
            Me.cmdPrintStaff.CommandType = CommandType.Text
            Me.cmdPrintStaff.CommandText = "SELECT * FROM  tblSchoolStaff WHERE (status=1) AND " & _
                vbNewLine & " (FullName LIKE @FullName) ORDER BY empNo"
            Me.cmdPrintStaff.Parameters.Clear()
            Me.cmdPrintStaff.Parameters.AddWithValue("@FullName", String.Format("%{0}%", TryCast(Me.txtSearchTeacher.Text.Trim, String).Trim))
            reader = Me.cmdPrintStaff.ExecuteReader
            While reader.Read
                li = Me.lstTTTeacher.Items.Add(IIf(DBNull.Value.Equals(reader!empNo), "", reader!empNo))
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

    Private Sub btnPreview_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPreview.Click
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            If Me.lstTTTeacher.CheckedItems.Count <> 1 Then
                MsgBox("Preview one item at a time.", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly + MsgBoxStyle.ApplicationModal, "System Reply")
                Exit Sub
            Else
                Me.Cursor = Cursors.WaitCursor
                Dim RptResultsView As New crtTTPrintStaff
                SetReportLogOn(RptResultsView)
                RptResultsView.SummaryInfo.ReportTitle = "SCHOOL TIME TABLE FOR TEACHERS"
                RptResultsView.RecordSelectionFormula = "({vwTTPrintTeacher.empNo}=" & Chr(34) & Me.lstTTTeacher.CheckedItems(0).Text.Trim & Chr(34) & ")"

                frmResultsViewing.crtViewResultsSummary.ReportSource = RptResultsView
                frmResultsViewing.crtViewResultsSummary.Zoom(100)
                frmResultsViewing.crtViewResultsSummary.RefreshReport()
                frmResultsViewing.MdiParent = frmHome
                frmResultsViewing.Show()
                frmResultsViewing.Text = "Print Staff"
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

    Private Sub lstTTTeacher_ColumnClick(ByVal sender As Object, ByVal e As System.Windows.Forms.ColumnClickEventArgs) Handles lstTTTeacher.ColumnClick
        If (e.Column() = 0) And (Me.lstTTTeacher.Items.Count > 0) Then
            For Each Li As ListViewItem In Me.lstTTTeacher.Items
                Li.Checked = Not (Li.Checked)
            Next
        End If
    End Sub

    Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        If Me.lstTTTeacher.Items.Count <= 0 Then
            MsgBox("Load items in the list before printing.", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly + MsgBoxStyle.ApplicationModal, "System Reply")
            Exit Sub
        ElseIf Me.lstTTTeacher.CheckedItems.Count <= 0 Then
            MsgBox("Check the items to print.", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly + MsgBoxStyle.ApplicationModal, "System Reply")
            Exit Sub
        Else
            Dim result As MsgBoxResult = MsgBox("Print Seleced Items?", MsgBoxStyle.Question + MsgBoxStyle.ApplicationModal + MsgBoxStyle.YesNo, "Confirm Transaction")
            If result = MsgBoxResult.No Then
                Exit Sub
            End If
            Me.Cursor = Cursors.WaitCursor
            j = 1
            For i = 0 To Me.lstTTTeacher.CheckedItems.Count - 1
                 Dim RptResultsView As New crtTTPrintStaff
                SetReportLogOn(RptResultsView)
                RptResultsView.SummaryInfo.ReportTitle = "SCHOOL TIME TABLE FOR TEACHERS"
                RptResultsView.RecordSelectionFormula = "({vwTTPrintTeacher.empNo}=" & Chr(34) & Me.lstTTTeacher.CheckedItems(0).Text.Trim & Chr(34) & ")"

                RptResultsView.PrintToPrinter(1, True, 1, RptResultsView.FormatEngine.GetLastPageNumber(New CrystalDecisions.Shared.ReportPageRequestContext()))
                Me.ToolStripStatusLabel2.Text = "Printing...." & j & " Of " & Me.lstTTTeacher.CheckedItems.Count & " Reports."
                j = j + 1
            Next
            Me.ToolStripStatusLabel2.Text = "Printing Complete"
            Me.Cursor = Cursors.Arrow
        End If
        Me.lstTTTeacher.Items.Clear()
        Me.ToolStripStatusLabel2.Text = "Printing Status"
    End Sub


    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub btnLoad_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLoad.Click
        Me.lstTTTeacher.Items.Clear()
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            Me.cmdPrintStaff.Connection = conn
            Me.cmdPrintStaff.CommandType = CommandType.Text
            Me.cmdPrintStaff.CommandText = "SELECT * FROM tblSchoolStaff WHERE (status=1) ORDER BY empNo"
            Me.cmdPrintStaff.Parameters.Clear()
            reader = Me.cmdPrintStaff.ExecuteReader
            While reader.Read
                li = Me.lstTTTeacher.Items.Add(IIf(DBNull.Value.Equals(reader!empNo), "", reader!empNo))
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
End Class