Imports System.Data.SqlClient
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports System.Drawing.Printing
Public Class frmAcadReportForms
    Dim resultsFound As Boolean
    Dim reader As SqlDataReader
    Dim cmdRptCard As New SqlCommand
    Dim rec As Integer
    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub frmAcadReportForms_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
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
        Me.cboClass.Items.Clear()
        Me.cboYear.Items.Clear()
        Me.cboStream.Items.Clear()

        Me.cboClass.Text = ""
        Me.cboYear.Text = ""
        Me.cboStream.Text = ""

        Me.cboClass.SelectedIndex = -1
        Me.cboYear.SelectedIndex = -1
        Me.cboStream.SelectedIndex = -1

        Me.cmdRptCard.Connection = conn
        Me.cmdRptCard.CommandType = CommandType.Text
        Me.cmdRptCard.CommandText = "SELECT DISTINCT year FROM  tblClasses WHERE (year IS NOT NULL) AND " & _
            vbNewLine & " (status=1) ORDER BY Year"
        Me.cmdRptCard.Parameters.Clear()
        reader = Me.cmdRptCard.ExecuteReader
        While reader.Read
            Me.cboYear.Items.Add(IIf(DBNull.Value.Equals(reader!year), "", reader!year))
        End While
        reader.Close()

        Me.cmdRptCard.CommandText = "SELECT DISTINCT className FROM  tblClasses WHERE (className IS NOT NULL) AND " & _
            vbNewLine & " (status=1) ORDER BY className"
        Me.cmdRptCard.Parameters.Clear()
        reader = Me.cmdRptCard.ExecuteReader
        While reader.Read
            Me.cboClass.Items.Add(IIf(DBNull.Value.Equals(reader!className), "", reader!className))
        End While
        reader.Close()

        Me.cmdRptCard.CommandText = "SELECT DISTINCT stream FROM  tblClasses WHERE (stream IS NOT NULL) AND " & _
            vbNewLine & " (status=1) ORDER BY stream"
        Me.cmdRptCard.Parameters.Clear()
        reader = Me.cmdRptCard.ExecuteReader
        While reader.Read
            Me.cboStream.Items.Add(IIf(DBNull.Value.Equals(reader!stream), "", reader!stream))
        End While
        reader.Close()
    End Sub
    Private Sub frmAcadReportForms_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint
        If Me.IsMdiChild Then
            Dim PnlLoc As New Point
            PnlLoc.X = CInt((Me.Width - Me.pnlReportForms.Width) / 2)
            PnlLoc.Y = CInt((Me.Height - Me.pnlReportForms.Height) / 2.5)
            Me.pnlReportForms.Location = PnlLoc
        Else
            Me.pnlReportForms.Dock = DockStyle.Fill
        End If
    End Sub

    Private Sub cboYear_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboYear.SelectedIndexChanged
        Me.lstStudentSubject.Items.Clear()
        If Me.cboYear.Text.Trim.Length <= 0 Then
            Exit Sub
        End If

        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            Me.cboTerm.Items.Clear()
            Me.cboTerm.Text = ""
            Me.cboTerm.SelectedIndex = -1

            Me.cmdRptCard.Connection = conn
            Me.cmdRptCard.CommandType = CommandType.Text
            Me.cmdRptCard.CommandText = "SELECT DISTINCT termName FROM tblSchoolCalendar WHERE (termName IS NOT NULL) AND " & _
                vbNewLine & " (status=1) AND (year=@year) ORDER BY termName"
            Me.cmdRptCard.Parameters.Clear()
            Me.cmdRptCard.Parameters.AddWithValue("@year", Me.cboYear.Text.Trim)
            reader = Me.cmdRptCard.ExecuteReader
            While reader.Read
                Me.cboTerm.Items.Add(IIf(DBNull.Value.Equals(reader!termName), "", reader!termName))
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

    Private Sub btnLoad_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLoad.Click
        Me.lstStudentSubject.Items.Clear()
        resultsFound = True
        If Me.cboYear.Text.Trim.Length <= 0 Then
            MsgBox("Year is Missing", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly + MsgBoxStyle.ApplicationModal, "Missing Details")
            Exit Sub
        ElseIf Me.cboClass.Text.Trim.Length <= 0 Then
            MsgBox("Class Name is Missing", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly + MsgBoxStyle.ApplicationModal, "Missing Details")
            Exit Sub
        ElseIf Me.cboStream.Text.Trim.Length <= 0 Then
            MsgBox("Strem Name is Missing", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly + MsgBoxStyle.ApplicationModal, "Missing Details")
            Exit Sub
        ElseIf Me.cboTerm.Text.Trim.Length <= 0 Then
            MsgBox("Term is Missing", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly + MsgBoxStyle.ApplicationModal, "Missing Details")
            Exit Sub
        End If

        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()

            Me.cmdRptCard.Connection = conn
            Me.cmdRptCard.CommandType = CommandType.StoredProcedure
            Me.cmdRptCard.CommandText = "sprocReportFormCheckIfOk"
            Me.cmdRptCard.Parameters.Clear()
            Me.cmdRptCard.Parameters.AddWithValue("@class", Me.cboClass.Text.Trim)
            Me.cmdRptCard.Parameters.AddWithValue("@stream", Me.cboStream.Text.Trim)
            Me.cmdRptCard.Parameters.AddWithValue("@year", Me.cboYear.Text.Trim)
            Me.cmdRptCard.Parameters.AddWithValue("@term", Me.cboTerm.Text.Trim)
            reader = Me.cmdRptCard.ExecuteReader
            While reader.Read
                resultsFound = IIf(DBNull.Value.Equals(reader!resultsFound), "", reader!resultsFound)
            End While
            reader.Close()

            If resultsFound = False Then
                MsgBox("Right Click and Run Results First!", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly + MsgBoxStyle.ApplicationModal, "Missing Details")
                Exit Sub
            End If

            Me.cmdRptCard.Connection = conn
            Me.cmdRptCard.CommandType = CommandType.Text
            Me.cmdRptCard.CommandText = "SELECT * FROM vwStudClass WHERE (classStatus=1) " & _
                vbNewLine & " AND (studStatus=1) AND (className=@className) AND (year=@year) AND (stream=@stream) " & _
                vbNewLine & " ORDER BY admNo"
            Me.cmdRptCard.Parameters.Clear()
            Me.cmdRptCard.Parameters.AddWithValue("@year", Me.cboYear.Text.Trim)
            Me.cmdRptCard.Parameters.AddWithValue("@className", Me.cboClass.Text.Trim)
            Me.cmdRptCard.Parameters.AddWithValue("@stream", Me.cboStream.Text.Trim)
            reader = Me.cmdRptCard.ExecuteReader
            While reader.Read
                li = Me.lstStudentSubject.Items.Add(IIf(DBNull.Value.Equals(reader!admNo), "", reader!admNo))
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

    Private Sub cboClass_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboClass.SelectedIndexChanged
        Me.lstStudentSubject.Items.Clear()
    End Sub

    Private Sub cboTerm_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboTerm.SelectedIndexChanged
        Me.lstStudentSubject.Items.Clear()
    End Sub
    Private Sub lstStudentSubject_ColumnClick(ByVal sender As Object, ByVal e As System.Windows.Forms.ColumnClickEventArgs) Handles lstStudentSubject.ColumnClick
        If (e.Column() = 0) And (Me.lstStudentSubject.Items.Count > 0) Then
            For Each Li As ListViewItem In Me.lstStudentSubject.Items
                Li.Checked = Not (Li.Checked)
            Next
        End If
    End Sub
    Private Sub cboStream_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboStream.SelectedIndexChanged
        Me.lstStudentSubject.Items.Clear()
    End Sub

    Private Sub btnPreview_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPreview.Click
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            If Me.lstStudentSubject.CheckedItems.Count <> 1 Then
                MsgBox("Preview one item at a time.", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly + MsgBoxStyle.ApplicationModal, "System Reply")
                Exit Sub
            Else
                Me.Cursor = Cursors.WaitCursor

                Me.cmdRptCard.Connection = conn
                Me.cmdRptCard.CommandType = CommandType.StoredProcedure
                Me.cmdRptCard.CommandText = "sprocTempReportForm"
                Me.cmdRptCard.Parameters.Clear()
                Me.cmdRptCard.Parameters.AddWithValue("@year1", Me.cboYear.Text.Trim)
                Me.cmdRptCard.Parameters.AddWithValue("@term1", Me.cboTerm.Text.Trim)
                Me.cmdRptCard.Parameters.AddWithValue("@stream1", Me.cboStream.Text.Trim)
                Me.cmdRptCard.Parameters.AddWithValue("@admNo1", Me.lstStudentSubject.CheckedItems(0).Text.Trim)
                Me.cmdRptCard.Parameters.AddWithValue("@className1", Me.cboClass.Text.Trim)
                Me.cmdRptCard.ExecuteNonQuery()

                Dim RptResultsView As New crtAcadResultsReportForm
                SetReportLogOn(RptResultsView)
                RptResultsView.SummaryInfo.ReportComments = fullName.Trim
                RptResultsView.RecordSelectionFormula = "({tblTempReportForm.admNo}=" & Me.lstStudentSubject.CheckedItems(0).Text.Trim & ")"
                RptResultsView.RecordSelectionFormula += "AND ({tblTempReportForm.year}=" & Me.cboYear.Text.Trim & ")"
                RptResultsView.RecordSelectionFormula += "AND ({tblTempReportForm.className}=" & Chr(34) & Me.cboClass.Text.Trim & Chr(34) & ")"
                RptResultsView.RecordSelectionFormula += "AND ({tblTempReportForm.termName}=" & Chr(34) & Me.cboTerm.Text.Trim & Chr(34) & ")"
                RptResultsView.RecordSelectionFormula += "AND ({tblTempReportForm.stream}=" & Chr(34) & Me.cboStream.Text.Trim & Chr(34) & ")"

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
        If Me.lstStudentSubject.Items.Count <= 0 Then
            MsgBox("Load items in the list before printing.", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly + MsgBoxStyle.ApplicationModal, "System Reply")
            Exit Sub
        ElseIf Me.lstStudentSubject.CheckedItems.Count <= 0 Then
            MsgBox("Check the items to print.", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly + MsgBoxStyle.ApplicationModal, "System Reply")
            Exit Sub
        Else
            Dim result As MsgBoxResult = MsgBox("Print Seleced Items?", MsgBoxStyle.Question + MsgBoxStyle.ApplicationModal + MsgBoxStyle.YesNo, "Confirm Transaction")
            If result = MsgBoxResult.No Then
                Exit Sub
            End If
            Try
                If conn.State = ConnectionState.Closed Then
                    conn.Open()
                End If
                dbconnection()
                Me.Cursor = Cursors.WaitCursor

                j = 1
                

                For i = 0 To Me.lstStudentSubject.CheckedItems.Count - 1
                    Me.cmdRptCard.Connection = conn
                    Me.cmdRptCard.CommandType = CommandType.StoredProcedure
                    Me.cmdRptCard.CommandText = "sprocTempReportForm"
                    Me.cmdRptCard.Parameters.Clear()
                    Me.cmdRptCard.Parameters.AddWithValue("@year1", Me.cboYear.Text.Trim)
                    Me.cmdRptCard.Parameters.AddWithValue("@term1", Me.cboTerm.Text.Trim)
                    Me.cmdRptCard.Parameters.AddWithValue("@stream1", Me.cboStream.Text.Trim)
                    Me.cmdRptCard.Parameters.AddWithValue("@admNo1", Me.lstStudentSubject.CheckedItems(i).Text.Trim)
                    Me.cmdRptCard.Parameters.AddWithValue("@className1", Me.cboClass.Text.Trim)
                    Me.cmdRptCard.ExecuteNonQuery()

                    Dim RptResultsView As New crtAcadResultsReportForm
                    SetReportLogOn(RptResultsView)
                    Dim ReportPrintOptions As PrintOptions = RptResultsView.PrintOptions
                    Dim ReportPrinter As New PrinterSettings
                    ReportPrintOptions.PrinterName = ReportPrinter.PrinterName
                    ReportPrintOptions.PaperOrientation = PaperOrientation.DefaultPaperOrientation
                    ReportPrintOptions.PaperSize = CrystalDecisions.Shared.PaperSize.PaperA4
                    ReportPrintOptions.PrinterDuplex = PrinterDuplex.Default
                    RptResultsView.SummaryInfo.ReportComments = fullName.Trim
                    RptResultsView.RecordSelectionFormula = "({tblTempReportForm.admNo}=" & Me.lstStudentSubject.CheckedItems(i).Text.Trim & ")"
                    RptResultsView.RecordSelectionFormula += "AND ({tblTempReportForm.year}=" & Me.cboYear.Text.Trim & ")"
                    RptResultsView.RecordSelectionFormula += "AND ({tblTempReportForm.className}=" & Chr(34) & Me.cboClass.Text.Trim & Chr(34) & ")"
                    RptResultsView.RecordSelectionFormula += "AND ({tblTempReportForm.termName}=" & Chr(34) & Me.cboTerm.Text.Trim & Chr(34) & ")"
                    RptResultsView.RecordSelectionFormula += "AND ({tblTempReportForm.stream}=" & Chr(34) & Me.cboStream.Text.Trim & Chr(34) & ")"

                    RptResultsView.PrintToPrinter(1, True, 1, 1)
                    Me.ToolStripStatusLabel2.Text = "Printing...." & j & " Of " & Me.lstStudentSubject.CheckedItems.Count & " Reports."
                    j = j + 1
                Next
                Me.ToolStripStatusLabel2.Text = "Printing Complete"
                Me.Cursor = Cursors.Arrow

                Me.lstStudentSubject.Items.Clear()
                Me.ToolStripStatusLabel2.Text = "Printing Status"

            Catch ex As Exception
                MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Finally
                If conn.State = ConnectionState.Open Then
                    conn.Close()
                End If
            End Try
        End If
    End Sub

    Private Sub txtSearchAdmNo_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtSearchAdmNo.TextChanged
        Me.lstStudentSubject.Items.Clear()
        If Me.txtSearchAdmNo.Text.Trim.Length <= 0 Then
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            Me.cmdRptCard.Connection = conn
            Me.cmdRptCard.CommandType = CommandType.Text
            Me.cmdRptCard.CommandText = "SELECT * FROM vwStudClass WHERE (classStatus=1) AND (classStudStatus=1) " & _
                vbNewLine & " AND (studStatus=1) AND  (admNo LIKE @admNo) " & _
                vbNewLine & " ORDER BY FullName"
            Me.cmdRptCard.Parameters.Clear()
            Me.cmdRptCard.Parameters.AddWithValue("@admNo", String.Format("%{0}%", TryCast(Me.txtSearchAdmNo.Text.Trim, String).Trim))
            reader = Me.cmdRptCard.ExecuteReader
            While reader.Read
                li = Me.lstStudentSubject.Items.Add(IIf(DBNull.Value.Equals(reader!admNo), "", reader!admNo))
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

    Private Sub lstStudentSubject_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lstStudentSubject.SelectedIndexChanged

    End Sub

    Private Sub CLOSEToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CLOSEToolStripMenuItem.Click
        Me.Close()
    End Sub

    Private Sub UPDATEToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UPDATEToolStripMenuItem.Click
        If Me.cboYear.Text.Trim.Length <= 0 Then
            MsgBox("Year is Missing", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly + MsgBoxStyle.ApplicationModal, "Missing Details")
            Exit Sub
        ElseIf Me.cboTerm.Text.Trim.Length <= 0 Then
            MsgBox("Term is Missing", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly + MsgBoxStyle.ApplicationModal, "Missing Details")
            Exit Sub
        ElseIf Me.cboClass.Text.Trim.Length <= 0 Then
            MsgBox("Class is Missing", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly + MsgBoxStyle.ApplicationModal, "Missing Details")
            Exit Sub
        ElseIf Me.cboStream.Text.Trim.Length <= 0 Then
            MsgBox("Stream is Missing", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly + MsgBoxStyle.ApplicationModal, "Missing Details")
            Exit Sub
        End If
        Dim result As MsgBoxResult = MsgBox("Are You Sure You Want To Run Summaries?", MsgBoxStyle.Question + MsgBoxStyle.ApplicationModal + MsgBoxStyle.YesNo, "Confirm Transaction")
        If result = MsgBoxResult.No Then
            Exit Sub
        End If


        Try
            Me.Cursor = Cursors.WaitCursor
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            Me.cmdRptCard.Connection = conn
            Me.cmdRptCard.CommandType = CommandType.StoredProcedure
            Me.cmdRptCard.CommandText = "sprocTempTermSummary"
            Me.cmdRptCard.Parameters.Clear()
            Me.cmdRptCard.Parameters.AddWithValue("@termYear", Me.cboYear.Text.Trim)
            Me.cmdRptCard.Parameters.AddWithValue("@termName", Me.cboTerm.Text.Trim)
            Me.cmdRptCard.Parameters.AddWithValue("@class", Me.cboClass.Text.Trim)
            Me.cmdRptCard.Parameters.AddWithValue("@stream", Me.cboStream.Text.Trim)
            Me.cmdRptCard.Parameters.AddWithValue("@runBy", userName)
            Me.cmdRptCard.ExecuteNonQuery()

            Me.cmdRptCard.Connection = conn
            Me.cmdRptCard.CommandType = CommandType.StoredProcedure
            Me.cmdRptCard.CommandText = "sprocTempExamResultsSubjects"
            Me.cmdRptCard.Parameters.Clear()
            Me.cmdRptCard.Parameters.AddWithValue("@term", Me.cboTerm.Text.Trim)
            Me.cmdRptCard.Parameters.AddWithValue("@class", Me.cboClass.Text.Trim)
            Me.cmdRptCard.Parameters.AddWithValue("@stream", Me.cboStream.Text.Trim)
            Me.cmdRptCard.Parameters.AddWithValue("@year", Me.cboYear.Text.Trim)
            Me.cmdRptCard.ExecuteNonQuery()


            MsgBox("Run Is SuccessFully Completed!", MsgBoxStyle.Information + MsgBoxStyle.OkOnly + MsgBoxStyle.ApplicationModal, "System Reply")
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            Me.Cursor = Cursors.Arrow
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub

    'Private Sub RUNSTREAMRESULTSToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RUNSTREAMRESULTSToolStripMenuItem.Click
    '    If Me.cboTerm.Text.Trim.Length <= 0 Then
    '        MsgBox("Term is Missing", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly + MsgBoxStyle.ApplicationModal, "Missing Details")
    '        Exit Sub
    '    ElseIf Me.cboClass.Text.Trim.Length <= 0 Then
    '        MsgBox("Class is Missing", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly + MsgBoxStyle.ApplicationModal, "Missing Details")
    '        Exit Sub
    '    ElseIf Me.cboYear.Text.Trim.Length <= 0 Then
    '        MsgBox("Year is Missing", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly + MsgBoxStyle.ApplicationModal, "Missing Details")
    '        Exit Sub
    '    ElseIf Me.cboStream.Text.Trim.Length <= 0 Then
    '        MsgBox("Stream is Missing", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly + MsgBoxStyle.ApplicationModal, "Missing Details")
    '        Exit Sub
    '    End If
    '    Dim result As MsgBoxResult = MsgBox("Are You Sure You Want To Run Summaries?", MsgBoxStyle.Question + MsgBoxStyle.ApplicationModal + MsgBoxStyle.YesNo, "Confirm Transaction")
    '    If result = MsgBoxResult.No Then
    '        Exit Sub
    '    End If

    '    Try
    '        Me.Cursor = Cursors.WaitCursor
    '        If conn.State = ConnectionState.Closed Then
    '            conn.Open()
    '        End If
    '        dbconnection()
    '        Me.cmdRptCard.Connection = conn
    '        Me.cmdRptCard.CommandType = CommandType.StoredProcedure
    '        Me.cmdRptCard.CommandText = "sprocTempExamResultsSubjects"
    '        Me.cmdRptCard.Parameters.Clear()
    '        Me.cmdRptCard.Parameters.AddWithValue("@term", Me.cboTerm.Text.Trim)
    '        Me.cmdRptCard.Parameters.AddWithValue("@class", Me.cboClass.Text.Trim)
    '        Me.cmdRptCard.Parameters.AddWithValue("@stream", Me.cboStream.Text.Trim)
    '        Me.cmdRptCard.Parameters.AddWithValue("@year", Me.cboYear.Text.Trim)
    '        Me.cmdRptCard.ExecuteNonQuery()

    '        MsgBox("Run Is SuccessFully Completed!", MsgBoxStyle.Information + MsgBoxStyle.OkOnly + MsgBoxStyle.ApplicationModal, "System Reply")
    '    Catch ex As Exception
    '        MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
    '    Finally
    '        Me.Cursor = Cursors.Arrow
    '        If conn.State = ConnectionState.Open Then
    '            conn.Close()
    '        End If
    '    End Try
    'End Sub
End Class