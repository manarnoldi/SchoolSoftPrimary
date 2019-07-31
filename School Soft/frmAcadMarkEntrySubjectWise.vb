Imports System.Data.SqlClient
Public Class frmAcadMarkEntrySubjectWise
    Dim allowedAccess As Boolean = False
    Dim cmdMarkClassWise As New SqlCommand
    Dim reader As SqlDataReader
    Dim rec As Integer = 0
    Dim queryType As String = Nothing
    Dim teacher_sub_found As Boolean = False
    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub frmAcadMarkEntrySubjectWise_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint
        If Me.IsMdiChild Then
            Dim PnlLoc As New Point
            PnlLoc.X = CInt((Me.Width - Me.pnlAcadMarkEntryClass.Width) / 2)
            PnlLoc.Y = CInt((Me.Height - Me.pnlAcadMarkEntryClass.Height) / 2.5)
            Me.pnlAcadMarkEntryClass.Location = PnlLoc
        Else
            Me.pnlAcadMarkEntryClass.Dock = DockStyle.Fill
        End If
    End Sub

    Private Sub frmAcadMarkEntrySubjectWise_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
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
        Me.cboClass.Text = ""
        Me.cboClass.SelectedIndex = -1

        Me.cboStream.Items.Clear()
        Me.cboStream.Text = ""
        Me.cboStream.SelectedIndex = -1

        Me.cboYear.Items.Clear()
        Me.cboYear.Text = ""
        Me.cboYear.SelectedIndex = -1

        Me.cboExamType.Items.Clear()
        Me.cboExamType.Text = ""
        Me.cboExamType.SelectedIndex = -1

        Me.cmdMarkClassWise.Connection = conn
        Me.cmdMarkClassWise.CommandType = CommandType.Text
        Me.cmdMarkClassWise.CommandText = "SELECT DISTINCT className FROM tblClasses WHERE (status=1) ORDER BY className"
        Me.cmdMarkClassWise.Parameters.Clear()
        reader = Me.cmdMarkClassWise.ExecuteReader()
        If reader.HasRows = True Then
            While reader.Read
                Me.cboClass.Items.Add(IIf(DBNull.Value.Equals(reader!className), "", (reader!className)))
            End While
        End If
        reader.Close()

        Me.cmdMarkClassWise.CommandText = "SELECT DISTINCT stream FROM tblClasses WHERE (status=1) ORDER BY stream"
        Me.cmdMarkClassWise.Parameters.Clear()
        reader = Me.cmdMarkClassWise.ExecuteReader()
        If reader.HasRows = True Then
            While reader.Read
                Me.cboStream.Items.Add(IIf(DBNull.Value.Equals(reader!stream), "", (reader!stream)))
            End While
        End If
        reader.Close()

        Me.cmdMarkClassWise.CommandText = "SELECT DISTINCT year FROM tblClasses WHERE (status=1) ORDER BY year"
        Me.cmdMarkClassWise.Parameters.Clear()
        reader = Me.cmdMarkClassWise.ExecuteReader()
        If reader.HasRows = True Then
            While reader.Read
                Me.cboYear.Items.Add(IIf(DBNull.Value.Equals(reader!year), "", (reader!year)))
            End While
        End If
        reader.Close()

        Me.cmdMarkClassWise.CommandText = "SELECT DISTINCT examType FROM tblExamNames WHERE (status=1) ORDER BY examType"
        Me.cmdMarkClassWise.Parameters.Clear()
        reader = Me.cmdMarkClassWise.ExecuteReader()
        If reader.HasRows = True Then
            While reader.Read
                Me.cboExamType.Items.Add(IIf(DBNull.Value.Equals(reader!examType), "", (reader!examType)))
            End While
        End If
        reader.Close()
    End Sub
    Private Sub cboYear_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboYear.SelectedIndexChanged, cboClass.SelectedIndexChanged, cboStream.SelectedIndexChanged
        
        Me.cboTerm.Items.Clear()
        Me.cboTerm.Text = ""
        Me.cboTerm.SelectedIndex = -1

        Me.dgVwShettSubjectWise.Rows.Clear()
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            If Me.cboYear.Text.Trim.Length > 0 Then
                Me.cmdMarkClassWise.Connection = conn
                Me.cmdMarkClassWise.CommandType = CommandType.Text
                Me.cmdMarkClassWise.CommandText = "SELECT DISTINCT termName FROM tblSchoolCalendar WHERE (status=1) AND (year=@year)" & _
                    vbNewLine & " ORDER BY termName"
                Me.cmdMarkClassWise.Parameters.Clear()
                Me.cmdMarkClassWise.Parameters.AddWithValue("@year", Me.cboYear.Text.Trim)
                reader = Me.cmdMarkClassWise.ExecuteReader()
                If reader.HasRows = True Then
                    While reader.Read
                        Me.cboTerm.Items.Add(IIf(DBNull.Value.Equals(reader!termName), "", (reader!termName)))
                    End While
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

    Private Sub cboTerm_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboTerm.SelectedIndexChanged
        Me.cboExamName.Items.Clear()
        Me.cboExamName.Text = ""
        Me.cboExamName.SelectedIndex = -1

        Me.cboExamType.Text = ""
        Me.cboExamType.SelectedIndex = -1

        Me.cboSubject.Items.Clear()
        Me.cboSubject.Text = ""
        Me.cboSubject.SelectedIndex = -1

        Me.dgVwShettSubjectWise.Rows.Clear()
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            If Me.cboYear.Text.Trim.Length > 0 And Me.cboStream.Text.Trim.Length > 0 And Me.cboClass.Text.Trim.Length > 0 And Me.cboTerm.Text.Trim.Length > 0 Then
                Me.cmdMarkClassWise.Connection = conn
                Me.cmdMarkClassWise.CommandType = CommandType.StoredProcedure
                Me.cmdMarkClassWise.CommandText = "sprocExamEntrySubTermSubjects"
                Me.cmdMarkClassWise.Parameters.Clear()
                Me.cmdMarkClassWise.Parameters.AddWithValue("@Year", Me.cboYear.Text.Trim)
                Me.cmdMarkClassWise.Parameters.AddWithValue("@stream", Me.cboStream.Text.Trim)
                Me.cmdMarkClassWise.Parameters.AddWithValue("@class", Me.cboClass.Text.Trim)
                Me.cmdMarkClassWise.Parameters.AddWithValue("@term", Me.cboTerm.Text.Trim)
                reader = Me.cmdMarkClassWise.ExecuteReader()
                If reader.HasRows = True Then
                    While reader.Read
                        Me.cboSubject.Items.Add(IIf(DBNull.Value.Equals(reader!subName), "", (reader!subName)))
                    End While
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

    Private Sub cboSubject_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboSubject.SelectedIndexChanged

        Me.cboExamType.Text = ""
        Me.cboExamType.SelectedIndex = -1

        Me.cboExamName.Items.Clear()
        Me.cboExamName.Text = ""
        Me.cboExamName.SelectedIndex = -1
        Me.dgVwShettSubjectWise.Rows.Clear()

       
    End Sub
    Private Function checkIfSubTeacherRegistered(ByVal className As String, ByVal stream As String, ByVal year As Integer, ByVal subject As String)
        Dim teacherFound As Boolean = False

        Me.cmdMarkClassWise.Connection = conn
        Me.cmdMarkClassWise.CommandType = CommandType.Text
        Me.cmdMarkClassWise.CommandText = "SELECT * FROM vwAcadStaffSubjects WHERE (subName=@subName) AND (className=@className) " & _
            vbNewLine & " AND (stream=@stream) AND (year=@year)"
        Me.cmdMarkClassWise.Parameters.Clear()
        Me.cmdMarkClassWise.Parameters.AddWithValue("@subName", Me.cboSubject.Text.Trim)
        Me.cmdMarkClassWise.Parameters.AddWithValue("@className", Me.cboClass.Text.Trim)
        Me.cmdMarkClassWise.Parameters.AddWithValue("@stream", Me.cboStream.Text.Trim)
        Me.cmdMarkClassWise.Parameters.AddWithValue("@year", Me.cboYear.Text.Trim)
        reader = Me.cmdMarkClassWise.ExecuteReader()
        If reader.HasRows = True Then
            teacherFound = True
        ElseIf reader.HasRows = False Then
            teacherFound = False
        End If
        reader.Close()
        Return teacherFound
    End Function
    Private Sub btnLoad_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLoad.Click
        Me.dgVwShettSubjectWise.Rows.Clear()
        If Me.cboSubject.Text.Trim.Length <= 0 Then
            MsgBox("Missing Subject Name.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboClass.Text.Trim.Length <= 0 Then
            MsgBox("Missing Class Name.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboExamName.Text.Trim.Length <= 0 Then
            MsgBox("Missing Exam Name.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboStream.Text.Trim.Length <= 0 Then
            MsgBox("Missing Stream Name.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboTerm.Text.Trim.Length <= 0 Then
            MsgBox("Missing Term Name.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboYear.Text.Trim.Length <= 0 Then
            MsgBox("Missing Year.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If
        Me.Cursor = Cursors.WaitCursor
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            Me.dgVwShettSubjectWise.Rows.Clear()
            allowedAccess = False
            checkRights()
            If allowedAccess = False Then
                Exit Sub
            End If

            teacher_sub_found = checkIfSubTeacherRegistered(Me.cboClass.Text.Trim, Me.cboStream.Text.Trim, Me.cboYear.Text.Trim, Me.cboSubject.Text.Trim)
            If teacher_sub_found = False Then
                MsgBox("Subject Teacher Not Registered.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
                Exit Sub
            End If
            Me.cmdMarkClassWise.Connection = conn
            Me.cmdMarkClassWise.CommandType = CommandType.Text
            Me.cmdMarkClassWise.CommandText = "SELECT * FROM vwAcadStudentExams WHERE (className=@className) AND (stream=@stream) " & _
                vbNewLine & " AND (classYear=@classYear) AND (termName=@termName) AND (subName=@subName)  AND (examName=@examName) " & _
                vbNewLine & " AND (examType=@examType) ORDER BY admNo"
            Me.cmdMarkClassWise.Parameters.Clear()
            Me.cmdMarkClassWise.Parameters.AddWithValue("@className", Me.cboClass.Text.Trim)
            Me.cmdMarkClassWise.Parameters.AddWithValue("@stream", Me.cboStream.Text.Trim)
            Me.cmdMarkClassWise.Parameters.AddWithValue("@classYear", Me.cboYear.Text.Trim)
            Me.cmdMarkClassWise.Parameters.AddWithValue("@termName", Me.cboTerm.Text.Trim)
            Me.cmdMarkClassWise.Parameters.AddWithValue("@subName", Me.cboSubject.Text.Trim)
            Me.cmdMarkClassWise.Parameters.AddWithValue("@examName", Me.cboExamName.Text.Trim)
            Me.cmdMarkClassWise.Parameters.AddWithValue("@examType", Me.cboExamType.Text.Trim)
            reader = Me.cmdMarkClassWise.ExecuteReader()
            If reader.HasRows = True Then
                While reader.Read
                    Dim rowNum As Integer = Me.dgVwShettSubjectWise.Rows.Count
                    Dim row As String() = New String() {IIf(DBNull.Value.Equals(reader!admNo), "", (reader!admNo)), _
                                    IIf(DBNull.Value.Equals(reader!FullName), "", (reader!FullName)), _
                                    IIf(DBNull.Value.Equals(reader!subName), "", (reader!subName)), _
                                    IIf(DBNull.Value.Equals(reader!examName), "", (reader!examName)), _
                                    "", IIf(DBNull.Value.Equals(reader!totalMarks), "", (reader!totalMarks))}
                    Me.dgVwShettSubjectWise.Rows.Add(row)
                    Me.dgVwShettSubjectWise.Rows(rowNum).Cells(2).Tag = IIf(DBNull.Value.Equals(reader!examId), "", (reader!examId))
                    'Me.dgVwShettSubjectWise.Rows(rowNum).Tag = IIf(DBNull.Value.Equals(reader!marks), "", (reader!marks))
                End While
            ElseIf reader.HasRows = False Then
                MsgBox("No record Found For the Selection Done.", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "System Reply")
            End If
            reader.Close()


            i = 0

            For i = 0 To Me.dgVwShettSubjectWise.RowCount - 1
                Me.cmdMarkClassWise.CommandType = CommandType.StoredProcedure
                Me.cmdMarkClassWise.Connection = conn
                Me.cmdMarkClassWise.CommandText = "sprocAcadResultsClasswise"
                Me.cmdMarkClassWise.Parameters.Clear()
                Me.cmdMarkClassWise.Parameters.AddWithValue("@termName", Me.cboTerm.Text.Trim)
                Me.cmdMarkClassWise.Parameters.AddWithValue("@termYear", Me.cboYear.Text.Trim)
                Me.cmdMarkClassWise.Parameters.AddWithValue("@admNo", Me.dgVwShettSubjectWise.Rows(i).Cells(0).Value)
                Me.cmdMarkClassWise.Parameters.AddWithValue("@examId", Me.dgVwShettSubjectWise.Rows(i).Cells(2).Tag)

                If conn.State = ConnectionState.Closed Then
                    conn.Open()
                End If

                reader = Me.cmdMarkClassWise.ExecuteReader
                While reader.Read
                    Me.dgVwShettSubjectWise.Rows(i).Cells(4).Value = IIf(DBNull.Value.Equals(reader!marks), "", (reader!marks))
                    Me.dgVwShettSubjectWise.Rows(i).Tag = IIf(DBNull.Value.Equals(reader!marks), "", (reader!marks))
                End While
                reader.Close()
            Next
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            Me.Cursor = Cursors.Arrow
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub
    Private Sub checkRights()
        Me.cmdMarkClassWise.CommandType = CommandType.Text
        Me.cmdMarkClassWise.Connection = conn
        Me.cmdMarkClassWise.CommandText = "SELECT * FROM  vwUserRights WHERE  (userStatus=1) AND (domainStatus=1) AND " & _
            vbNewLine & " (staffStatus=1) AND (modStatus=1) AND (rightAccess=1) AND (empNo=@empNo) AND (domainName='ADMINISTRATOR')"
        Me.cmdMarkClassWise.Parameters.Clear()
        Me.cmdMarkClassWise.Parameters.AddWithValue("@empNo", empNo.Trim)
        reader = Me.cmdMarkClassWise.ExecuteReader
        If reader.HasRows = True Then
            allowedAccess = True
            reader.Close()
            Exit Sub
        ElseIf reader.HasRows = False Then
            allowedAccess = False
        End If
        reader.Close()

        Me.cmdMarkClassWise.CommandText = "SELECT * FROM vwAcadStaffSubjects WHERE (className=@className) AND (stream=@stream)" & _
            vbNewLine & " AND (year=@year) AND (empNo=@empNo) AND (subName=@subName)"
        Me.cmdMarkClassWise.Parameters.Clear()
        Me.cmdMarkClassWise.Parameters.AddWithValue("@className", Me.cboClass.Text.Trim)
        Me.cmdMarkClassWise.Parameters.AddWithValue("@stream", Me.cboStream.Text.Trim)
        Me.cmdMarkClassWise.Parameters.AddWithValue("@year", Me.cboYear.Text.Trim)
        Me.cmdMarkClassWise.Parameters.AddWithValue("@empNo", empNo.Trim)
        Me.cmdMarkClassWise.Parameters.AddWithValue("@subName", Me.cboSubject.Text.Trim)
        reader = Me.cmdMarkClassWise.ExecuteReader
        If reader.HasRows = True Then
            allowedAccess = True
        ElseIf reader.HasRows = False Then
            MsgBox("You are not the Subject Teacher of the class." & _
                   vbNewLine & "You cannot enter marks subjectWise.", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly + MsgBoxStyle.ApplicationModal, "System Reply")
            allowedAccess = False
            reader.Close()
            Exit Sub
        End If
        reader.Close()

        Me.cmdMarkClassWise.CommandText = "SELECT * FROM tblPeriods WHERE (periodName='EXAM RESULTS ENTRY') AND (status=1)"
        Me.cmdMarkClassWise.Parameters.Clear()
        reader = Me.cmdMarkClassWise.ExecuteReader
        If reader.HasRows = True Then
            While reader.Read
                Dim dateDiffBeg As Integer = DateDiff(DateInterval.Day, Date.Now.Date, IIf(DBNull.Value.Equals(reader!dateBeginning), "", (reader!dateBeginning)))
                Dim dateDiffEnd As Integer = DateDiff(DateInterval.Day, Date.Now.Date, IIf(DBNull.Value.Equals(reader!dateEnding), "", (reader!dateEnding)))
                If dateDiffBeg > 0 Then
                    MsgBox("Exam Results Entry Period Has Not Begun." & _
                  vbNewLine & "Contact the system Administrator.", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly + MsgBoxStyle.ApplicationModal, "System Reply")
                    allowedAccess = False
                ElseIf dateDiffEnd < 0 Then
                    MsgBox("Exam Results Entry Period Is Closed." & _
                 vbNewLine & "Contact the system Administrator.", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly + MsgBoxStyle.ApplicationModal, "System Reply")
                    allowedAccess = False
                Else
                    allowedAccess = True
                End If
            End While
        ElseIf reader.HasRows = False Then
            MsgBox("Exam Results entry period is not defined.", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly + MsgBoxStyle.ApplicationModal, "System Reply")
            allowedAccess = False
        End If
        reader.Close()
    End Sub

    Private Sub cboExamName_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboExamName.SelectedIndexChanged
        Me.dgVwShettSubjectWise.Rows.Clear()
    End Sub

    Private Sub dgVwShettSubjectWise_CellValueChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgVwShettSubjectWise.CellValueChanged
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            Dim mark As Double = 0
            Dim markDiff As Double = 0
            If e.ColumnIndex = 4 Then
                If Not (Me.dgVwShettSubjectWise.Rows(e.RowIndex).Cells(e.ColumnIndex).Value = "") Then
                    If IsNumeric(Me.dgVwShettSubjectWise.Rows(e.RowIndex).Cells(e.ColumnIndex).Value) = False Then
                        MsgBox("Non Numeric Values detected.", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly + MsgBoxStyle.ApplicationModal, "Error Detected")
                        Me.dgVwShettSubjectWise.Rows(e.RowIndex).Cells(e.ColumnIndex).Value = Me.dgVwShettSubjectWise.Rows(e.RowIndex).Tag
                        Exit Sub
                    End If
                    mark = Me.dgVwShettSubjectWise.Rows(e.RowIndex).Cells(e.ColumnIndex).Value
                    markDiff = Me.dgVwShettSubjectWise.Rows(e.RowIndex).Cells(e.ColumnIndex + 1).Value - mark
                    If markDiff < 0 Then
                        MsgBox("Mark Scored Cannot Exceed Total Mark.", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly + MsgBoxStyle.ApplicationModal, "Error Detected")
                        Me.dgVwShettSubjectWise.Rows(e.RowIndex).Cells(e.ColumnIndex).Value = Me.dgVwShettSubjectWise.Rows(e.RowIndex).Tag
                    End If
                    If mark < 0 Then
                        MsgBox("Mark Cannot be less than one.", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly + MsgBoxStyle.ApplicationModal, "Error Detected")
                        Me.dgVwShettSubjectWise.Rows(e.RowIndex).Cells(e.ColumnIndex).Value = Me.dgVwShettSubjectWise.Rows(e.RowIndex).Tag
                    End If
                End If
            End If
        Catch ex As Exception
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub

    Private Sub btnUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdate.Click
        Dim noRecord As Boolean = True
        If Me.cboYear.Text.Trim.Length <= 0 Then
            MsgBox("Missing Year", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly + MsgBoxStyle.ApplicationModal, "Error Detected")
            Exit Sub
        ElseIf Me.cboSubject.Text.Trim.Length <= 0 Then
            MsgBox("Missing Subject Name.", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly + MsgBoxStyle.ApplicationModal, "Error Detected")
            Exit Sub
        ElseIf Me.cboClass.Text.Trim.Length <= 0 Then
            MsgBox("Missing Class Name", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly + MsgBoxStyle.ApplicationModal, "Error Detected")
            Exit Sub
        ElseIf Me.cboExamName.Text.Trim.Length <= 0 Then
            MsgBox("Missing Exam Name", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly + MsgBoxStyle.ApplicationModal, "Error Detected")
            Exit Sub
        ElseIf Me.cboStream.Text.Trim.Length <= 0 Then
            MsgBox("Missing stream Name", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly + MsgBoxStyle.ApplicationModal, "Error Detected")
            Exit Sub
        ElseIf Me.cboTerm.Text.Trim.Length <= 0 Then
            MsgBox("Missing term Name", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly + MsgBoxStyle.ApplicationModal, "Error Detected")
            Exit Sub
        ElseIf Me.dgVwShettSubjectWise.Rows.Count <= 0 Then
            MsgBox("No records to Update", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly + MsgBoxStyle.ApplicationModal, "Error Detected")
            Exit Sub
        End If
        i = 0
        For i = 0 To Me.dgVwShettSubjectWise.Rows.Count - 1
            If Not (Me.dgVwShettSubjectWise.Rows(i).Cells(4).Value = Nothing) Then
                noRecord = False
            End If
        Next
        If noRecord = True Then
            MsgBox("No Marks to Update", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly + MsgBoxStyle.ApplicationModal, "Error Detected")
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            Dim result As MsgBoxResult = MsgBox("Update Record/s?", MsgBoxStyle.Question + MsgBoxStyle.ApplicationModal + MsgBoxStyle.YesNo, "Confirm Transaction")
            If result = MsgBoxResult.No Then
                Exit Sub
            End If
            i = 0
            rec = 0
            For i = 0 To Me.dgVwShettSubjectWise.Rows.Count - 1
                If Not (Me.dgVwShettSubjectWise.Rows(i).Cells(4).Value = Nothing) Then
                    Me.cmdMarkClassWise.Connection = conn
                    Me.cmdMarkClassWise.CommandText = "sprocExamResults"
                    Me.cmdMarkClassWise.CommandType = CommandType.StoredProcedure
                    Me.cmdMarkClassWise.Parameters.Clear()
                    Me.cmdMarkClassWise.Parameters.AddWithValue("@className", Me.cboClass.Text.Trim)
                    Me.cmdMarkClassWise.Parameters.AddWithValue("@stream", Me.cboStream.Text.Trim)
                    Me.cmdMarkClassWise.Parameters.AddWithValue("@year", Me.cboYear.Text.Trim)
                    Me.cmdMarkClassWise.Parameters.AddWithValue("@subName", Me.cboSubject.Text.Trim)
                    Me.cmdMarkClassWise.Parameters.AddWithValue("@termName", Me.cboTerm.Text.Trim)
                    Me.cmdMarkClassWise.Parameters.AddWithValue("@examName", Me.cboExamName.Text.Trim)
                    Me.cmdMarkClassWise.Parameters.AddWithValue("@examType", Me.cboExamType.Text.Trim)
                    Me.cmdMarkClassWise.Parameters.AddWithValue("@admNo", Me.dgVwShettSubjectWise.Rows(i).Cells(0).Value)
                    Me.cmdMarkClassWise.Parameters.AddWithValue("@marks", Me.dgVwShettSubjectWise.Rows(i).Cells(4).Value)
                    Me.cmdMarkClassWise.Parameters.AddWithValue("@fullName", Me.dgVwShettSubjectWise.Rows(i).Cells(1).Value)
                    Me.cmdMarkClassWise.Parameters.AddWithValue("@regBy", userName.Trim)
                    Me.cmdMarkClassWise.Parameters.AddWithValue("@dateOfReg", Date.Now)
                    Me.cmdMarkClassWise.Parameters.AddWithValue("@moduleName", "RESULTS ENTRY-SUBJECTWISE")
                    rec = rec + Me.cmdMarkClassWise.ExecuteNonQuery
                End If
            Next
            If rec > 0 Then
                MsgBox("Record/s Updated", MsgBoxStyle.Information + MsgBoxStyle.OkOnly + MsgBoxStyle.ApplicationModal, "System Reply")
            End If
            'clearTexts()
            'loadCombos()
            Me.dgVwShettSubjectWise.Rows.Clear()
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub
    Private Sub clearTexts()
        Me.cboSubject.Items.Clear()
        Me.cboSubject.Text = ""
        Me.cboSubject.SelectedIndex = -1

        Me.cboClass.Items.Clear()
        Me.cboClass.Text = ""
        Me.cboClass.SelectedIndex = -1

        Me.cboExamName.Items.Clear()
        Me.cboExamName.Text = ""
        Me.cboExamName.SelectedIndex = -1

        Me.cboExamType.Items.Clear()
        Me.cboExamType.Text = ""
        Me.cboExamType.SelectedIndex = -1

        Me.cboStream.Items.Clear()
        Me.cboStream.Text = ""
        Me.cboStream.SelectedIndex = -1

        Me.cboTerm.Items.Clear()
        Me.cboTerm.Text = ""
        Me.cboTerm.SelectedIndex = -1

        Me.cboYear.Items.Clear()
        Me.cboYear.Text = ""
        Me.cboYear.SelectedIndex = -1

        Me.dgVwShettSubjectWise.Rows.Clear()
    End Sub

    Private Sub cboExamType_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboExamType.SelectedIndexChanged
        Me.cboExamName.Items.Clear()
        Me.cboExamName.Text = ""
        Me.cboExamName.SelectedIndex = -1

        Me.dgVwShettSubjectWise.Rows.Clear()

        If Me.cboSubject.Text.Trim.Length > 0 And Me.cboYear.Text.Trim.Length > 0 And Me.cboStream.Text.Trim.Length > 0 And _
            Me.cboClass.Text.Trim.Length > 0 And Me.cboTerm.Text.Trim.Length > 0 And Me.cboExamType.Text.Trim.Length > 0 Then
            Try
                If conn.State = ConnectionState.Closed Then
                    conn.Open()
                End If
                dbconnection()
                Me.cmdMarkClassWise.Connection = conn
                Me.cmdMarkClassWise.CommandType = CommandType.Text
                Me.cmdMarkClassWise.CommandText = "SELECT DISTINCT examName FROM vwAcadClassExams WHERE (className=@className) AND " & _
                    vbNewLine & " (stream=@stream) AND (classYear=@classYear) AND (termName=@termName) AND (subName=@subName) " & _
                    vbNewLine & " AND (examType=@examType) ORDER BY examName"
                Me.cmdMarkClassWise.Parameters.Clear()
                Me.cmdMarkClassWise.Parameters.AddWithValue("@termName", Me.cboTerm.Text.Trim)
                Me.cmdMarkClassWise.Parameters.AddWithValue("@classYear", Me.cboYear.Text.Trim)
                Me.cmdMarkClassWise.Parameters.AddWithValue("@stream", Me.cboStream.Text.Trim)
                Me.cmdMarkClassWise.Parameters.AddWithValue("@className", Me.cboClass.Text.Trim)
                Me.cmdMarkClassWise.Parameters.AddWithValue("@examType", Me.cboExamType.Text.Trim)
                Me.cmdMarkClassWise.Parameters.AddWithValue("@subName", Me.cboSubject.Text.Trim)
                reader = Me.cmdMarkClassWise.ExecuteReader()
                If reader.HasRows = True Then
                    While reader.Read
                        Me.cboExamName.Items.Add(IIf(DBNull.Value.Equals(reader!examName), "", (reader!examName)))
                    End While
                End If
                reader.Close()
            Catch ex As Exception
                MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Finally
                If conn.State = ConnectionState.Open Then
                    conn.Close()
                End If
            End Try
        End If
    End Sub

    Private Sub dgVwShettSubjectWise_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgVwShettSubjectWise.CellContentClick

    End Sub
End Class