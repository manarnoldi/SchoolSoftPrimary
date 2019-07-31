Imports System.Data.SqlClient
Public Class frmAcadMarkEntryClassWise
    Dim allowedAccess As Boolean = False
    Dim cmdMarkClassWise As New SqlCommand
    Dim reader As SqlDataReader
    Dim rec As Integer = 0
    Dim queryType As String = Nothing

    Private Sub frmAcadMarkEntryClassWise_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
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

        Me.cboExamType.Items.Clear()
        Me.cboExamType.Text = ""
        Me.cboExamType.SelectedIndex = -1

        Me.cboYear.Items.Clear()
        Me.cboYear.Text = ""
        Me.cboYear.SelectedIndex = -1

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
    Private Sub frmAcadMarkEntryClassWise_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint
        If Me.IsMdiChild Then
            Dim PnlLoc As New Point
            PnlLoc.X = CInt((Me.Width - Me.pnlAcadMarkEntryClass.Width) / 2)
            PnlLoc.Y = CInt((Me.Height - Me.pnlAcadMarkEntryClass.Height) / 2.5)
            Me.pnlAcadMarkEntryClass.Location = PnlLoc
        Else
            Me.pnlAcadMarkEntryClass.Dock = DockStyle.Fill
        End If
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub cboYear_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboYear.SelectedIndexChanged, cboClass.SelectedIndexChanged, cboStream.SelectedIndexChanged
        Me.cboAdmNo.Items.Clear()
        Me.cboAdmNo.Text = ""
        Me.cboAdmNo.SelectedIndex = -1

        Me.cboTerm.Items.Clear()
        Me.cboTerm.Text = ""
        Me.cboTerm.SelectedIndex = -1

        Me.cboExam.Items.Clear()
        Me.cboExam.Text = ""
        Me.cboExam.SelectedIndex = -1

        Me.dgVwSheetClasswise.Rows.Clear()
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
            If Me.cboYear.Text.Trim.Length > 0 And Me.cboStream.Text.Trim.Length > 0 And Me.cboClass.Text.Trim.Length > 0 Then
                Me.cmdMarkClassWise.Connection = conn
                Me.cmdMarkClassWise.CommandType = CommandType.Text
                Me.cmdMarkClassWise.CommandText = "SELECT admNo FROM  vwStudClass WHERE (className=@className) AND " & _
                    vbNewLine & " (stream=@stream) AND (year=@year) AND (classStatus=1) AND (classStudStatus=1) AND " & _
                    vbNewLine & " (studStatus=1) ORDER BY admNo"
                Me.cmdMarkClassWise.Parameters.Clear()
                Me.cmdMarkClassWise.Parameters.AddWithValue("@year", Me.cboYear.Text.Trim)
                Me.cmdMarkClassWise.Parameters.AddWithValue("@stream", Me.cboStream.Text.Trim)
                Me.cmdMarkClassWise.Parameters.AddWithValue("@className", Me.cboClass.Text.Trim)
                reader = Me.cmdMarkClassWise.ExecuteReader()
                If reader.HasRows = True Then
                    While reader.Read
                        Me.cboAdmNo.Items.Add(IIf(DBNull.Value.Equals(reader!admNo), "", (reader!admNo)))
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

        Me.cmdMarkClassWise.CommandText = "SELECT * FROM vwClassHeads WHERE  (classStatus=1) AND (classHeadStatus=1) AND " & _
            vbNewLine & " (schoolStaffStatus=1) AND (studentStatus=1) AND (className=@className) AND (stream=@stream) AND " & _
            vbNewLine & " (year=@year) AND (empNo=@empNo)"
        Me.cmdMarkClassWise.Parameters.Clear()
        Me.cmdMarkClassWise.Parameters.AddWithValue("@className", Me.cboClass.Text.Trim)
        Me.cmdMarkClassWise.Parameters.AddWithValue("@stream", Me.cboStream.Text.Trim)
        Me.cmdMarkClassWise.Parameters.AddWithValue("@year", Me.cboYear.Text.Trim)
        Me.cmdMarkClassWise.Parameters.AddWithValue("@empNo", empNo.Trim)
        reader = Me.cmdMarkClassWise.ExecuteReader
        If reader.HasRows = True Then
            allowedAccess = True
        ElseIf reader.HasRows = False Then
            MsgBox("You are not the class teacher." & _
                   vbNewLine & "You cannot enter marks classwise.", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly + MsgBoxStyle.ApplicationModal, "System Reply")
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

    Private Sub btnView_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnView.Click
        If Me.cboAdmNo.Text.Trim.Length <= 0 Then
            MsgBox("Missing Admission Number.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboClass.Text.Trim.Length <= 0 Then
            MsgBox("Missing Class Name.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboExam.Text.Trim.Length <= 0 Then
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

        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            Me.dgVwSheetClasswise.Rows.Clear()
            allowedAccess = False
            checkRights()
            If allowedAccess = False Then
                Exit Sub
            End If
            Me.cmdMarkClassWise.Connection = conn
            Me.cmdMarkClassWise.CommandType = CommandType.Text
            Me.cmdMarkClassWise.CommandText = "SELECT * FROM vwAcadStudentExams WHERE (className=@className) AND (stream=@stream) " & _
                vbNewLine & " AND (classYear=@classYear) AND (termName=@termName) AND (admNo=@admNo)  AND (examName=@examName) AND " & _
                vbNewLine & " (examType=@examType) ORDER BY subCode"
            Me.cmdMarkClassWise.Parameters.Clear()
            Me.cmdMarkClassWise.Parameters.AddWithValue("@className", Me.cboClass.Text.Trim)
            Me.cmdMarkClassWise.Parameters.AddWithValue("@stream", Me.cboStream.Text.Trim)
            Me.cmdMarkClassWise.Parameters.AddWithValue("@classYear", Me.cboYear.Text.Trim)
            Me.cmdMarkClassWise.Parameters.AddWithValue("@termName", Me.cboTerm.Text.Trim)
            Me.cmdMarkClassWise.Parameters.AddWithValue("@admNo", Me.cboAdmNo.Text.Trim)
            Me.cmdMarkClassWise.Parameters.AddWithValue("@examName", Me.cboExam.Text.Trim)
            Me.cmdMarkClassWise.Parameters.AddWithValue("@examType", Me.cboExamType.Text.Trim)
            reader = Me.cmdMarkClassWise.ExecuteReader()
            If reader.HasRows = True Then
                While reader.Read
                    Dim rowNum As Integer = Me.dgVwSheetClasswise.Rows.Count
                    Dim row As String() = New String() {IIf(DBNull.Value.Equals(reader!admNo), "", (reader!admNo)), _
                                    IIf(DBNull.Value.Equals(reader!FullName), "", (reader!FullName)), _
                                    IIf(DBNull.Value.Equals(reader!subName), "", (reader!subName)), _
                                    IIf(DBNull.Value.Equals(reader!examName), "", (reader!examName)), _
                                    "", IIf(DBNull.Value.Equals(reader!totalMarks), "", (reader!totalMarks))}
                    Me.dgVwSheetClasswise.Rows.Add(row)
                    Me.dgVwSheetClasswise.Rows(rowNum).Cells(2).Tag = IIf(DBNull.Value.Equals(reader!examId), "", (reader!examId))

                End While
            ElseIf reader.HasRows = False Then
                MsgBox("No record Found For the Selection Done.", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "System Reply")
            End If
            reader.Close()

            i = 0
            
            For i = 0 To Me.dgVwSheetClasswise.RowCount - 1
                Me.cmdMarkClassWise.CommandType = CommandType.StoredProcedure
                Me.cmdMarkClassWise.Connection = conn
                Me.cmdMarkClassWise.CommandText = "sprocAcadResultsClasswise"
                Me.cmdMarkClassWise.Parameters.Clear()
                Me.cmdMarkClassWise.Parameters.AddWithValue("@termName", Me.cboTerm.Text.Trim)
                Me.cmdMarkClassWise.Parameters.AddWithValue("@termYear", Me.cboYear.Text.Trim)
                Me.cmdMarkClassWise.Parameters.AddWithValue("@admNo", Me.cboAdmNo.Text.Trim)
                Me.cmdMarkClassWise.Parameters.AddWithValue("@examId", Me.dgVwSheetClasswise.Rows(i).Cells(2).Tag)

                If conn.State = ConnectionState.Closed Then
                    conn.Open()
                End If

                reader = Me.cmdMarkClassWise.ExecuteReader
                While reader.Read
                    Me.dgVwSheetClasswise.Rows(i).Cells(4).Value = IIf(DBNull.Value.Equals(reader!marks), "", (reader!marks))
                    Me.dgVwSheetClasswise.Rows(i).Tag = IIf(DBNull.Value.Equals(reader!marks), "", (reader!marks))
                End While
                reader.Close()
            Next
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub

    Private Sub cboExam_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboExam.SelectedIndexChanged
        Me.dgVwSheetClasswise.Rows.Clear()
    End Sub

    Private Sub cboAdmNo_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboAdmNo.SelectedIndexChanged
        Me.dgVwSheetClasswise.Rows.Clear()
        Me.cboExam.Items.Clear()
        Me.cboExam.Text = ""
        Me.cboExam.SelectedIndex = -1

        Me.cboExamType.Text = ""
        Me.cboExamType.SelectedIndex = -1
    End Sub

    Private Sub btnUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdate.Click
        Dim noRecord As Boolean = True
        If Me.cboYear.Text.Trim.Length <= 0 Then
            MsgBox("Missing Year", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly + MsgBoxStyle.ApplicationModal, "Error Detected")
            Exit Sub
        ElseIf Me.cboAdmNo.Text.Trim.Length <= 0 Then
            MsgBox("Missing Student Number", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly + MsgBoxStyle.ApplicationModal, "Error Detected")
            Exit Sub
        ElseIf Me.cboClass.Text.Trim.Length <= 0 Then
            MsgBox("Missing Class Name", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly + MsgBoxStyle.ApplicationModal, "Error Detected")
            Exit Sub
        ElseIf Me.cboExam.Text.Trim.Length <= 0 Then
            MsgBox("Missing Exam Name", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly + MsgBoxStyle.ApplicationModal, "Error Detected")
            Exit Sub
        ElseIf Me.cboStream.Text.Trim.Length <= 0 Then
            MsgBox("Missing stream Name", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly + MsgBoxStyle.ApplicationModal, "Error Detected")
            Exit Sub
        ElseIf Me.cboTerm.Text.Trim.Length <= 0 Then
            MsgBox("Missing term Name", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly + MsgBoxStyle.ApplicationModal, "Error Detected")
            Exit Sub
        ElseIf Me.dgVwSheetClasswise.Rows.Count <= 0 Then
            MsgBox("No records to Update", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly + MsgBoxStyle.ApplicationModal, "Error Detected")
            Exit Sub
        End If
        i = 0
        For i = 0 To Me.dgVwSheetClasswise.Rows.Count - 1
            If Not (Me.dgVwSheetClasswise.Rows(i).Cells(4).Value = Nothing) Then
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
            For i = 0 To Me.dgVwSheetClasswise.Rows.Count - 1
                If Not (Me.dgVwSheetClasswise.Rows(i).Cells(4).Value = Nothing) Then
                    Me.cmdMarkClassWise.Connection = conn
                    Me.cmdMarkClassWise.CommandText = "sprocExamResults"
                    Me.cmdMarkClassWise.CommandType = CommandType.StoredProcedure
                    Me.cmdMarkClassWise.Parameters.Clear()
                    Me.cmdMarkClassWise.Parameters.AddWithValue("@className", Me.cboClass.Text.Trim)
                    Me.cmdMarkClassWise.Parameters.AddWithValue("@stream", Me.cboStream.Text.Trim)
                    Me.cmdMarkClassWise.Parameters.AddWithValue("@year", Me.cboYear.Text.Trim)
                    Me.cmdMarkClassWise.Parameters.AddWithValue("@subName", Me.dgVwSheetClasswise.Rows(i).Cells(2).Value)
                    Me.cmdMarkClassWise.Parameters.AddWithValue("@termName", Me.cboTerm.Text.Trim)
                    Me.cmdMarkClassWise.Parameters.AddWithValue("@examName", Me.cboExam.Text.Trim)
                    Me.cmdMarkClassWise.Parameters.AddWithValue("@examType", Me.cboExamType.Text.Trim)
                    Me.cmdMarkClassWise.Parameters.AddWithValue("@admNo", Me.cboAdmNo.Text.Trim)
                    Me.cmdMarkClassWise.Parameters.AddWithValue("@marks", Me.dgVwSheetClasswise.Rows(i).Cells(4).Value)
                    Me.cmdMarkClassWise.Parameters.AddWithValue("@fullName", Me.dgVwSheetClasswise.Rows(i).Cells(1).Value)
                    Me.cmdMarkClassWise.Parameters.AddWithValue("@regBy", userName.Trim)
                    Me.cmdMarkClassWise.Parameters.AddWithValue("@dateOfReg", Date.Now)
                    Me.cmdMarkClassWise.Parameters.AddWithValue("@moduleName", "RESULTS ENTRY-CLASSWISE")
                    rec = rec + Me.cmdMarkClassWise.ExecuteNonQuery
                End If
            Next
            If rec > 0 Then
                MsgBox("Record/s Updated", MsgBoxStyle.Information + MsgBoxStyle.OkOnly + MsgBoxStyle.ApplicationModal, "System Reply")
            End If
            'clearTexts()
            'loadCombos()
            Me.dgVwSheetClasswise.Rows.Clear()
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub

    Private Sub dgVwSheetClasswise_CellValueChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgVwSheetClasswise.CellValueChanged
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            Dim mark As Double = 0
            Dim markDiff As Double = 0
            If e.ColumnIndex = 4 Then
                If Not (Me.dgVwSheetClasswise.Rows(e.RowIndex).Cells(e.ColumnIndex).Value = "") Then
                    If IsNumeric(Me.dgVwSheetClasswise.Rows(e.RowIndex).Cells(e.ColumnIndex).Value) = False Then
                        MsgBox("Non Numeric Values detected.", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly + MsgBoxStyle.ApplicationModal, "Error Detected")
                        Me.dgVwSheetClasswise.Rows(e.RowIndex).Cells(e.ColumnIndex).Value = Me.dgVwSheetClasswise.Rows(e.RowIndex).Tag
                        Exit Sub
                    End If
                    mark = Me.dgVwSheetClasswise.Rows(e.RowIndex).Cells(e.ColumnIndex).Value
                    markDiff = Me.dgVwSheetClasswise.Rows(e.RowIndex).Cells(e.ColumnIndex + 1).Value - mark
                    If markDiff < 0 Then
                        MsgBox("Mark Scored Cannot Exceed Total Mark.", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly + MsgBoxStyle.ApplicationModal, "Error Detected")
                        Me.dgVwSheetClasswise.Rows(e.RowIndex).Cells(e.ColumnIndex).Value = Me.dgVwSheetClasswise.Rows(e.RowIndex).Tag
                    End If
                    If mark < 0 Then
                        MsgBox("Mark Cannot be less than one.", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly + MsgBoxStyle.ApplicationModal, "Error Detected")
                        Me.dgVwSheetClasswise.Rows(e.RowIndex).Cells(e.ColumnIndex).Value = Me.dgVwSheetClasswise.Rows(e.RowIndex).Tag
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
    Private Sub clearTexts()
        Me.cboAdmNo.Items.Clear()
        Me.cboAdmNo.Text = ""
        Me.cboAdmNo.SelectedIndex = -1

        Me.cboClass.Items.Clear()
        Me.cboClass.Text = ""
        Me.cboClass.SelectedIndex = -1

        Me.cboExamType.Items.Clear()
        Me.cboExamType.Text = ""
        Me.cboExamType.SelectedIndex = -1

        Me.cboExam.Items.Clear()
        Me.cboExam.Text = ""
        Me.cboExam.SelectedIndex = -1

        Me.cboStream.Items.Clear()
        Me.cboStream.Text = ""
        Me.cboStream.SelectedIndex = -1

        Me.cboTerm.Items.Clear()
        Me.cboTerm.Text = ""
        Me.cboTerm.SelectedIndex = -1

        Me.cboYear.Items.Clear()
        Me.cboYear.Text = ""
        Me.cboYear.SelectedIndex = -1

        Me.dgVwSheetClasswise.Rows.Clear()
    End Sub

    Private Sub cboExamType_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboExamType.SelectedIndexChanged
        Me.dgVwSheetClasswise.Rows.Clear()
        If Me.cboExamType.Text.Trim.Length <= 0 Then
            Exit Sub
        End If

        Me.cboExam.Items.Clear()
        Me.cboExam.Text = ""
        Me.cboExam.SelectedIndex = -1
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            If Me.cboTerm.Text.Trim.Length > 0 And Me.cboClass.Text.Trim.Length > 0 And Me.cboStream.Text.Trim.Length > 0 And _
                Me.cboYear.Text.Trim.Length > 0 And Me.cboExamType.Text.Trim.Length > 0 Then
                Me.cmdMarkClassWise.Connection = conn
                Me.cmdMarkClassWise.CommandType = CommandType.Text
                Me.cmdMarkClassWise.CommandText = "SELECT DISTINCT examName FROM vwAcadClassExams WHERE (className=@className) AND " & _
                    vbNewLine & " (stream=@stream) AND (classYear=@classYear) AND (termName=@termName) AND (examType=@examType) " & _
                    vbNewLine & " ORDER BY examName"
                Me.cmdMarkClassWise.Parameters.Clear()
                Me.cmdMarkClassWise.Parameters.AddWithValue("@termName", Me.cboTerm.Text.Trim)
                Me.cmdMarkClassWise.Parameters.AddWithValue("@classYear", Me.cboYear.Text.Trim)
                Me.cmdMarkClassWise.Parameters.AddWithValue("@stream", Me.cboStream.Text.Trim)
                Me.cmdMarkClassWise.Parameters.AddWithValue("@className", Me.cboClass.Text.Trim)
                Me.cmdMarkClassWise.Parameters.AddWithValue("@examType", Me.cboExamType.Text.Trim)
                reader = Me.cmdMarkClassWise.ExecuteReader()
                If reader.HasRows = True Then
                    While reader.Read
                        Me.cboExam.Items.Add(IIf(DBNull.Value.Equals(reader!examName), "", (reader!examName)))
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

    End Sub
End Class