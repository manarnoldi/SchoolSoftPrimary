Imports System.Data.SqlClient
Imports Excel = Microsoft.Office.Interop.Excel
Public Class frmAcadExamResultsView
    Dim cmdExamResults As New SqlCommand
    Dim reader As SqlDataReader
    Dim rec As Integer = 0
    Private Sub frmAcadExamResultsView_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
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

    Private Sub frmAcadExamResultsView_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint
        If Me.IsMdiChild Then
            Dim PnlLoc As New Point
            PnlLoc.X = CInt((Me.Width - Me.pnlAcadExamResultsView.Width) / 2)
            PnlLoc.Y = CInt((Me.Height - Me.pnlAcadExamResultsView.Height) / 2.5)
            Me.pnlAcadExamResultsView.Location = PnlLoc
        Else
            Me.pnlAcadExamResultsView.Dock = DockStyle.Fill
        End If
    End Sub
    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub
    Private Sub loadCombos()

        Me.cboYear.Items.Clear()
        Me.cboYear.Text = ""
        Me.cboYear.SelectedIndex = -1

        Me.cboClass.Items.Clear()
        Me.cboClass.Text = ""
        Me.cboClass.SelectedIndex = -1

        Me.cboExamType.Items.Clear()
        Me.cboExamType.Text = ""
        Me.cboExamType.SelectedIndex = -1

        Me.cboStream.Items.Clear()
        Me.cboStream.Text = ""
        Me.cboStream.SelectedIndex = -1

        Me.cmdExamResults.CommandType = CommandType.Text
        Me.cmdExamResults.Connection = conn
        Me.cmdExamResults.CommandText = "SELECT DISTINCT year FROM tblClasses WHERE (status=1) ORDER BY year"
        Me.cmdExamResults.Parameters.Clear()
        reader = Me.cmdExamResults.ExecuteReader
        While reader.Read
            Me.cboYear.Items.Add(IIf(DBNull.Value.Equals(reader!year), "", reader!year))
        End While
        reader.Close()

        Me.cmdExamResults.CommandText = "SELECT DISTINCT stream FROM tblClasses WHERE (status=1) ORDER BY stream"
        Me.cmdExamResults.Parameters.Clear()
        reader = Me.cmdExamResults.ExecuteReader
        While reader.Read
            Me.cboStream.Items.Add(IIf(DBNull.Value.Equals(reader!stream), "", reader!stream))
        End While
        reader.Close()

        Me.cmdExamResults.CommandText = "SELECT DISTINCT className FROM tblClasses WHERE (status=1) ORDER BY className"
        Me.cmdExamResults.Parameters.Clear()
        reader = Me.cmdExamResults.ExecuteReader
        While reader.Read
            Me.cboClass.Items.Add(IIf(DBNull.Value.Equals(reader!className), "", reader!className))
        End While
        reader.Close()

        Me.cmdExamResults.CommandText = "SELECT DISTINCT examType FROM tblExamNames WHERE (status=1) ORDER BY examType"
        Me.cmdExamResults.Parameters.Clear()
        reader = Me.cmdExamResults.ExecuteReader()
        If reader.HasRows = True Then
            While reader.Read
                Me.cboExamType.Items.Add(IIf(DBNull.Value.Equals(reader!examType), "", (reader!examType)))
            End While
        End If
        reader.Close()
    End Sub

    Private Sub cboYear_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboYear.SelectedIndexChanged
        Me.lstStaffSubject.Items.Clear()
        If Me.cboYear.Text.Trim.Length <= 0 Then
            Exit Sub
        End If

        Me.cboTerm.Items.Clear()
        Me.cboTerm.Text = ""
        Me.cboTerm.SelectedIndex = -1

        
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            Me.cmdExamResults.CommandType = CommandType.Text
            Me.cmdExamResults.Connection = conn
            Me.cmdExamResults.CommandText = "SELECT DISTINCT  termName FROM tblSchoolCalendar WHERE (status=1) AND (year=@year) " & _
                vbNewLine & "ORDER BY termName"
            Me.cmdExamResults.Parameters.Clear()
            Me.cmdExamResults.Parameters.AddWithValue("@year", Me.cboYear.Text.Trim)
            reader = Me.cmdExamResults.ExecuteReader
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

    Private Sub cboClass_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboClass.SelectedIndexChanged, cboStream.SelectedIndexChanged, cboTerm.SelectedIndexChanged
        Me.lstStaffSubject.Items.Clear()
        If Me.cboClass.Text.Trim.Length <= 0 Or Me.cboStream.Text.Trim.Length <= 0 Or Me.cboTerm.Text.Trim.Length <= 0 Or Me.cboYear.Text.Trim.Length <= 0 Then
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()

            Me.cboSubject.Items.Clear()
            Me.cboSubject.SelectedIndex = -1
            Me.cboSubject.Text = ""

            Me.cmdExamResults.CommandType = CommandType.Text
            Me.cmdExamResults.Connection = conn

            Me.cmdExamResults.Connection = conn
            Me.cmdExamResults.CommandType = CommandType.StoredProcedure
            Me.cmdExamResults.CommandText = "sprocExamEntrySubTermSubjects"
            Me.cmdExamResults.Parameters.Clear()
            Me.cmdExamResults.Parameters.AddWithValue("@Year", Me.cboYear.Text.Trim)
            Me.cmdExamResults.Parameters.AddWithValue("@stream", Me.cboStream.Text.Trim)
            Me.cmdExamResults.Parameters.AddWithValue("@class", Me.cboClass.Text.Trim)
            Me.cmdExamResults.Parameters.AddWithValue("@term", Me.cboTerm.Text.Trim)
            reader = Me.cmdExamResults.ExecuteReader
            While reader.Read
                Me.cboSubject.Items.Add(IIf(DBNull.Value.Equals(reader!subName), "", reader!subName))
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

    Private Sub cboSubject_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboSubject.SelectedIndexChanged
        Me.lstStaffSubject.Items.Clear()
    End Sub

    Private Sub btnLoad_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLoad.Click
        Me.lstStaffSubject.Items.Clear()
        If Me.cboClass.Text.Trim.Length <= 0 Then
            MsgBox("Class Name is missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboExam.Text.Trim.Length <= 0 Then
            MsgBox("Exam Name is missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboStream.Text.Trim.Length <= 0 Then
            MsgBox("stream Name is missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboSubject.Text.Trim.Length <= 0 Then
            MsgBox("Subject Name is missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboTerm.Text.Trim.Length <= 0 Then
            MsgBox("Term Name is missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboYear.Text.Trim.Length <= 0 Then
            MsgBox("Year is missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()

            Me.cmdExamResults.CommandType = CommandType.StoredProcedure
            Me.cmdExamResults.Connection = conn
            Me.cmdExamResults.CommandText = "sprocExamResultsView"
            Me.cmdExamResults.Parameters.Clear()
            Me.cmdExamResults.Parameters.AddWithValue("@year", Me.cboYear.Text.Trim)
            Me.cmdExamResults.Parameters.AddWithValue("@class", Me.cboClass.Text.Trim)
            Me.cmdExamResults.Parameters.AddWithValue("@stream", Me.cboStream.Text.Trim)
            Me.cmdExamResults.Parameters.AddWithValue("@termName", Me.cboTerm.Text.Trim)
            Me.cmdExamResults.Parameters.AddWithValue("@subject", Me.cboSubject.Text.Trim)
            Me.cmdExamResults.Parameters.AddWithValue("@examName", Me.cboExam.Text.Trim)
            reader = Me.cmdExamResults.ExecuteReader
            If reader.HasRows = True Then
                While reader.Read
                    li = Me.lstStaffSubject.Items.Add(IIf(DBNull.Value.Equals(reader!admNo), "", reader!admNo))
                    li.SubItems.Add(IIf(DBNull.Value.Equals(reader!FullName), "", reader!FullName))
                    li.SubItems.Add(IIf(DBNull.Value.Equals(reader!subName), "", reader!subName))
                    li.SubItems.Add(IIf(DBNull.Value.Equals(reader!year), "", reader!year))
                    li.SubItems.Add(IIf(DBNull.Value.Equals(reader!examName), "", reader!examName))
                    li.SubItems.Add(IIf(DBNull.Value.Equals(reader!marks), "", reader!marks))
                    li.SubItems.Add(IIf(DBNull.Value.Equals(reader!totalMarks), "", reader!totalMarks))
                    li.SubItems.Add(IIf(DBNull.Value.Equals(reader!examRank), "", reader!examRank))
                End While
            ElseIf reader.HasRows = False Then
                MsgBox("No Record Found!.", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "System Reply")
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

    Private Sub cboExam_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Me.lstStaffSubject.Items.Clear()
    End Sub

    Private Sub cboExam_SelectedIndexChanged_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboExam.SelectedIndexChanged
        Me.lstStaffSubject.Items.Clear()
    End Sub
    Private Sub releaseObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub
    Private Sub btnExport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExport.Click
        If Me.lstStaffSubject.Items.Count <= 0 Then
            MsgBox("No items in the list.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Missing Details")
            Exit Sub
        End If

        Me.Cursor = Cursors.WaitCursor

        Dim xlApp As Excel.Application
        Dim xlWorkBook As Excel.Workbook
        Dim xlWorkSheet As Excel.Worksheet
        Dim misValue As Object = System.Reflection.Missing.Value
        Dim i As Integer

        Dim year As String = Me.cboYear.Text.Trim
        Dim className As String = Me.cboClass.Text.Trim
        Dim stream As String = Me.cboStream.Text.Trim
        Dim term As String = Me.cboTerm.Text.Trim
        Dim subject As String = Me.cboSubject.Text.Trim
        Dim exam As String = Me.cboExam.Text.Trim
        Dim examType As String = Me.cboExamType.Text.Trim

        xlApp = New Excel.Application
        xlWorkBook = xlApp.Workbooks.Add(misValue)
        xlWorkSheet = CType(xlWorkBook.Sheets(1), Excel.Worksheet)

        xlWorkSheet.Cells(1, 1) = "ADM NO"
        xlWorkSheet.Cells(1, 2) = "FULLNAME"
        xlWorkSheet.Cells(1, 3) = "SUBJECT"
        xlWorkSheet.Cells(1, 4) = "EXAM TYPE"
        xlWorkSheet.Cells(1, 5) = "EXAM NAME"
        xlWorkSheet.Cells(1, 6) = "CLASS"
        xlWorkSheet.Cells(1, 7) = "STREAM"
        xlWorkSheet.Cells(1, 8) = "TERM"
        xlWorkSheet.Cells(1, 9) = "YEAR"
        xlWorkSheet.Cells(1, 10) = "SCORE"
        xlWorkSheet.Cells(1, 11) = "OUT OF"
        xlWorkSheet.Cells(1, 12) = "RANK"

        For i = 0 To Me.lstStaffSubject.Items.Count - 1
            xlWorkSheet.Cells(i + 2, 1) = Me.lstStaffSubject.Items(i).Text.Trim
            xlWorkSheet.Cells(i + 2, 2) = Me.lstStaffSubject.Items(i).SubItems(1).Text.Trim
            xlWorkSheet.Cells(i + 2, 3) = subject
            xlWorkSheet.Cells(i + 2, 4) = examType
            xlWorkSheet.Cells(i + 2, 5) = exam
            xlWorkSheet.Cells(i + 2, 6) = className
            xlWorkSheet.Cells(i + 2, 7) = stream
            xlWorkSheet.Cells(i + 2, 8) = term
            xlWorkSheet.Cells(i + 2, 9) = year
            xlWorkSheet.Cells(i + 2, 10) = Me.lstStaffSubject.Items(i).SubItems(5).Text.Trim
            xlWorkSheet.Cells(i + 2, 11) = Me.lstStaffSubject.Items(i).SubItems(6).Text.Trim
            xlWorkSheet.Cells(i + 2, 12) = Me.lstStaffSubject.Items(i).SubItems(7).Text.Trim
        Next
        Dim dlg As New SaveFileDialog
        dlg.Filter = "Excel Files (*.xlsx)|*.xlsx"
        dlg.FilterIndex = 1
        dlg.InitialDirectory = My.Application.Info.DirectoryPath & "\EXCEL\\EICHER\BILLS\"
        dlg.FileName = ""
        Dim ExcelFile As String = ""
        If dlg.ShowDialog = Windows.Forms.DialogResult.OK Then
            ExcelFile = dlg.FileName
            xlWorkSheet.SaveAs(ExcelFile)
        End If
        xlWorkBook.Close()

        xlApp.Quit()

        releaseObject(xlApp)
        releaseObject(xlWorkBook)
        releaseObject(xlWorkSheet)
        Me.Cursor = Cursors.Arrow
        MsgBox("Excel file created successfully", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Successful Transaction")
    End Sub

    Private Sub cboExamType_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboExamType.SelectedIndexChanged
        If Me.cboSubject.Text.Trim.Length <= 0 And Me.cboClass.Text.Trim.Length <= 0 And Me.cboStream.Text.Trim.Length <= 0 And _
            Me.cboTerm.Text.Trim.Length <= 0 And Me.cboYear.Text.Trim.Length <= 0 And Me.cboExamType.Text.Trim.Length <= 0 Then
            Exit Sub
        End If

        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()

            Me.cboExam.Items.Clear()
            Me.cboExam.Text = ""
            Me.cboExam.SelectedIndex = -1

            Me.lstStaffSubject.Items.Clear()
            Me.cmdExamResults.CommandType = CommandType.Text
            Me.cmdExamResults.Connection = conn
            Me.cmdExamResults.CommandText = "SELECT DISTINCT examName FROM vwAcadStudentExams WHERE (classYear=@year) AND " & _
                vbNewLine & " (termName=@termName) AND (className=@className) AND (stream=@stream) AND (subName=@subName) " & _
                vbNewLine & " AND (examType=@examType) ORDER BY examName"
            Me.cmdExamResults.Parameters.Clear()
            Me.cmdExamResults.Parameters.AddWithValue("@year", Me.cboYear.Text.Trim)
            Me.cmdExamResults.Parameters.AddWithValue("@className", Me.cboClass.Text.Trim)
            Me.cmdExamResults.Parameters.AddWithValue("@stream", Me.cboStream.Text.Trim)
            Me.cmdExamResults.Parameters.AddWithValue("@termName", Me.cboTerm.Text.Trim)
            Me.cmdExamResults.Parameters.AddWithValue("@subName", Me.cboSubject.Text.Trim)
            Me.cmdExamResults.Parameters.AddWithValue("@examType", Me.cboExamType.Text.Trim)

            reader = Me.cmdExamResults.ExecuteReader
            While reader.Read
                Me.cboExam.Items.Add(IIf(DBNull.Value.Equals(reader!examName), "", reader!examName))
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