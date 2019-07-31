Imports System.Data.SqlClient
Public Class frmAcadExaminations
    Dim rec As Integer = 0
    Dim reader As SqlDataReader
    Dim cmdAcadExams As New SqlCommand
    Dim queryType As String = Nothing
    Dim recordExists As Boolean = True
    Dim recordOk As Boolean = False
    Private Sub frmAcadExaminations_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            loadCombos()
            loadLists()
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub
    Private Sub loadLists()
        Me.lstClassDetails.Items.Clear()
        Me.lstVwSubjectDetails.Items.Clear()
        Me.lstViewExams.Items.Clear()

        Me.cmdAcadExams.Connection = conn
        Me.cmdAcadExams.CommandText = "SELECT * FROM tblSubjects WHERE (subStatus='True') ORDER BY subName"
        Me.cmdAcadExams.CommandType = CommandType.Text
        Me.cmdAcadExams.Parameters.Clear()
        reader = Me.cmdAcadExams.ExecuteReader
        If reader.HasRows Then
            While reader.Read
                li = Me.lstVwSubjectDetails.Items.Add(IIf(DBNull.Value.Equals(reader!subName), "", reader!subName))
            End While
        End If
        reader.Close()
    End Sub
    Private Sub frmAcadExaminations_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint
        If Me.IsMdiChild Then
            Dim PnlLoc As New Point
            PnlLoc.X = CInt((Me.Width - Me.pnlExaminations.Width) / 2)
            PnlLoc.Y = CInt((Me.Height - Me.pnlExaminations.Height) / 2.5)
            Me.pnlExaminations.Location = PnlLoc
        Else
            Me.pnlExaminations.Dock = DockStyle.Fill
        End If
    End Sub
    Private Sub loadCombos()
        Me.cboExamType.Items.Clear()
        Me.cboExamType.Text = ""
        Me.cboExamType.SelectedIndex = -1

        Me.cboExamName.Items.Clear()
        Me.cboExamName.Text = ""
        Me.cboExamName.SelectedIndex = -1

        Me.cboTotalMarks.Items.Clear()
        Me.cboTotalMarks.Text = ""
        Me.cboTotalMarks.SelectedIndex = -1

        Me.cboExamContr.Items.Clear()
        Me.cboExamContr.Text = ""
        Me.cboExamContr.SelectedIndex = -1

        Me.cboTermUpdate.Items.Clear()
        Me.cboTermUpdate.Text = ""
        Me.cboTermUpdate.SelectedIndex = -1

        Me.cboTermView.Items.Clear()
        Me.cboTermView.Text = ""
        Me.cboTermView.SelectedIndex = -1

        Me.cboYearUpdate.Items.Clear()
        Me.cboYearUpdate.Text = ""
        Me.cboYearUpdate.SelectedIndex = -1

        Me.cboYearView.Items.Clear()
        Me.cboYearView.Text = ""
        Me.cboYearView.SelectedIndex = -1

        Me.cmdAcadExams.Connection = conn
        Me.cmdAcadExams.CommandText = "SELECT DISTINCT examType FROM tblExamNames WHERE (status=1) ORDER BY examType"
        Me.cmdAcadExams.CommandType = CommandType.Text
        Me.cmdAcadExams.Parameters.Clear()
        reader = Me.cmdAcadExams.ExecuteReader
        If reader.HasRows Then
            While reader.Read
                Me.cboExamType.Items.Add(IIf(DBNull.Value.Equals(reader!examType), "", reader!examType))
            End While
        End If
        reader.Close()

        Me.cmdAcadExams.CommandText = "SELECT DISTINCT totalMarks FROM  tblClassExams WHERE (Status=1) ORDER BY totalMarks"
        Me.cmdAcadExams.Parameters.Clear()
        reader = Me.cmdAcadExams.ExecuteReader
        If reader.HasRows Then
            While reader.Read
                Me.cboTotalMarks.Items.Add(IIf(DBNull.Value.Equals(reader!totalMarks), "", reader!totalMarks))
            End While
        End If
        reader.Close()

        Me.cmdAcadExams.CommandText = "SELECT DISTINCT partTypeMark FROM  tblClassExams WHERE (status=1) ORDER BY partTypeMark"
        Me.cmdAcadExams.Parameters.Clear()
        reader = Me.cmdAcadExams.ExecuteReader
        If reader.HasRows Then
            While reader.Read
                Me.cboExamContr.Items.Add(IIf(DBNull.Value.Equals(reader!partTypeMark), "", reader!partTypeMark))
            End While
        End If
        reader.Close()

        Me.cmdAcadExams.CommandText = "SELECT DISTINCT termName FROM tblSchoolCalendar WHERE (Status='True') "
        Me.cmdAcadExams.Parameters.Clear()
        reader = Me.cmdAcadExams.ExecuteReader
        If reader.HasRows Then
            While reader.Read
                Me.cboTermUpdate.Items.Add(IIf(DBNull.Value.Equals(reader!termName), "", reader!termName))
            End While
        End If
        reader.Close()

        Me.cmdAcadExams.CommandText = "SELECT DISTINCT termName FROM tblSchoolCalendar WHERE (Status='True') "
        Me.cmdAcadExams.Parameters.Clear()
        reader = Me.cmdAcadExams.ExecuteReader
        If reader.HasRows Then
            While reader.Read
                Me.cboTermView.Items.Add(IIf(DBNull.Value.Equals(reader!termName), "", reader!termName))
            End While
        End If
        reader.Close()

        Me.cmdAcadExams.CommandText = "SELECT DISTINCT year FROM tblClasses WHERE (Status='True')"
        Me.cmdAcadExams.Parameters.Clear()
        reader = Me.cmdAcadExams.ExecuteReader
        If reader.HasRows Then
            While reader.Read
                Me.cboYearUpdate.Items.Add(IIf(DBNull.Value.Equals(reader!year), "", reader!year))
            End While
        End If
        reader.Close()

        Me.cmdAcadExams.CommandText = "SELECT DISTINCT year FROM tblClasses WHERE (Status='True')"
        Me.cmdAcadExams.Parameters.Clear()
        reader = Me.cmdAcadExams.ExecuteReader
        If reader.HasRows Then
            While reader.Read
                Me.cboYearView.Items.Add(IIf(DBNull.Value.Equals(reader!year), "", reader!year))
            End While
        End If
        reader.Close()
    End Sub
    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub cboYearUpdate_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboYearUpdate.SelectedIndexChanged, cboTermUpdate.SelectedIndexChanged
        If (Me.cboTermUpdate.Text.Trim.Length <= 0) Or (Me.cboYearUpdate.Text.Trim.Length <= 0) Then
            Exit Sub
        End If
        Try
            Me.lstClassDetails.Items.Clear()
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            Me.cmdAcadExams.CommandText = "SELECT * FROM tblClasses WHERE (Status='True') AND (year=@year) ORDER BY className,stream"
            Me.cmdAcadExams.Parameters.Clear()
            Me.cmdAcadExams.Parameters.AddWithValue("@year", Me.cboYearUpdate.Text.Trim)
            reader = Me.cmdAcadExams.ExecuteReader
            If reader.HasRows Then
                While reader.Read
                    li = Me.lstClassDetails.Items.Add(IIf(DBNull.Value.Equals(reader!className), "", reader!className))
                    li.SubItems.Add(IIf(DBNull.Value.Equals(reader!stream), "", reader!stream))
                    li.SubItems.Add(IIf(DBNull.Value.Equals(reader!year), "", reader!year))
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
    End Sub

    Private Sub cboYearView_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboYearView.SelectedIndexChanged
        Me.lstViewExams.Items.Clear()
        Try

            Me.cboClass.Items.Clear()
            Me.cboClass.Text = ""
            Me.cboClass.SelectedIndex = -1

            Me.cboStream.Items.Clear()
            Me.cboStream.Text = ""
            Me.cboStream.SelectedIndex = -1

            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()

            Me.cmdAcadExams.Connection = conn
            Me.cmdAcadExams.CommandText = "SELECT DISTINCT className FROM tblClasses WHERE (Status='True') AND (Year=@year)"
            Me.cmdAcadExams.CommandType = CommandType.Text
            Me.cmdAcadExams.Parameters.Clear()
            Me.cmdAcadExams.Parameters.AddWithValue("@year", Me.cboYearView.Text.Trim)
            reader = Me.cmdAcadExams.ExecuteReader
            If reader.HasRows Then
                While reader.Read
                    Me.cboClass.Items.Add(IIf(DBNull.Value.Equals(reader!className), "", reader!className))
                End While
            End If
            reader.Close()

            Me.cmdAcadExams.CommandText = "SELECT DISTINCT stream FROM tblClasses WHERE (Status='True') AND (Year=@Year)"
            Me.cmdAcadExams.CommandType = CommandType.Text
            Me.cmdAcadExams.Parameters.Clear()
            Me.cmdAcadExams.Parameters.AddWithValue("@year", Me.cboYearView.Text.Trim)
            reader = Me.cmdAcadExams.ExecuteReader
            If reader.HasRows Then
                While reader.Read
                    Me.cboStream.Items.Add(IIf(DBNull.Value.Equals(reader!stream), "", reader!stream))
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
    End Sub

    Private Sub lstVwSubjectDetails_ColumnClick(ByVal sender As Object, ByVal e As System.Windows.Forms.ColumnClickEventArgs) Handles lstVwSubjectDetails.ColumnClick
        If (e.Column() = 0) And (Me.lstVwSubjectDetails.Items.Count > 0) Then
            For Each Li As ListViewItem In Me.lstVwSubjectDetails.Items
                Li.Checked = Not (Li.Checked)
            Next
        End If
    End Sub

    Private Sub lstClassDetails_ColumnClick(ByVal sender As Object, ByVal e As System.Windows.Forms.ColumnClickEventArgs) Handles lstClassDetails.ColumnClick
        If (e.Column() = 0) And (Me.lstClassDetails.Items.Count > 0) Then
            For Each Li As ListViewItem In Me.lstClassDetails.Items
                Li.Checked = Not (Li.Checked)
            Next
        End If
    End Sub

    Private Sub btnLoad_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLoad.Click
        If Me.cboTermView.Text.Trim.Length <= 0 Then
            MsgBox("Missing Term Name", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboYearView.Text.Trim.Length <= 0 Then
            MsgBox("Missing class year", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboClass.Text.Trim.Length <= 0 Then
            MsgBox("Missing class name", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboStream.Text.Trim.Length <= 0 Then
            MsgBox("Missing class stream", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            Me.lstViewExams.Items.Clear()
            dbconnection()
            Me.cmdAcadExams.Connection = conn
            Me.cmdAcadExams.CommandText = "SELECT * FROM vwAcadClassExams WHERE (termName=@termName) AND (classYear=@classYear) " & _
                vbNewLine & " AND (className=@className) AND (stream=@stream) " & _
                vbNewLine & " ORDER BY subName,examType,examName,totalMarks,totalTypeMark,partTypeMark"
            Me.cmdAcadExams.CommandType = CommandType.Text
            Me.cmdAcadExams.Parameters.Clear()
            Me.cmdAcadExams.Parameters.AddWithValue("@termName", Me.cboTermView.Text.Trim)
            Me.cmdAcadExams.Parameters.AddWithValue("@classYear", Me.cboYearView.Text.Trim)
            Me.cmdAcadExams.Parameters.AddWithValue("@className", Me.cboClass.Text.Trim)
            Me.cmdAcadExams.Parameters.AddWithValue("@stream", Me.cboStream.Text.Trim)
            reader = Me.cmdAcadExams.ExecuteReader
            If reader.HasRows Then
                While reader.Read
                    li = Me.lstViewExams.Items.Add(IIf(DBNull.Value.Equals(reader!subName), "", reader!subName))
                    li.SubItems.Add(IIf(DBNull.Value.Equals(reader!examType), "", reader!examType))
                    li.SubItems.Add(IIf(DBNull.Value.Equals(reader!totalTypeMark), "", reader!totalTypeMark))
                    li.SubItems.Add(IIf(DBNull.Value.Equals(reader!examName), "", reader!examName))
                    li.SubItems.Add(IIf(DBNull.Value.Equals(reader!totalMarks), "", reader!totalMarks))
                    li.SubItems.Add(IIf(DBNull.Value.Equals(reader!partTypeMark), "", reader!partTypeMark))
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
    End Sub

    Private Sub txtSearchSubject_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtSearchSubject.TextChanged
        Me.lstViewExams.Items.Clear()
        If Me.cboTermView.Text.Trim.Length <= 0 Then
            MsgBox("Missing Term Name.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboYearView.Text.Trim.Length <= 0 Then
            MsgBox("Missing class Year.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboClass.Text.Trim.Length <= 0 Then
            MsgBox("Missing class Name.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboStream.Text.Trim.Length <= 0 Then
            MsgBox("Missing class Stream.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If
        If Me.txtSearchSubject.Text.Trim.Length <= 3 Then
            Me.lstViewExams.Items.Clear()
            Exit Sub
        End If
        
        Me.lstViewExams.Items.Clear()
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            Me.cmdAcadExams.Connection = conn
            Me.cmdAcadExams.CommandText = "SELECT * FROM vwAcadClassExams WHERE (termName=@termName) AND (classYear=@classYear) " & _
                vbNewLine & " AND (className=@className) AND  (stream=@stream) AND (subName LIKE @subName) " & _
                vbNewLine & " ORDER BY subName,examType,examName,totalMarks,totalTypeMark,partTypeMark"
            Me.cmdAcadExams.CommandType = CommandType.Text
            Me.cmdAcadExams.Parameters.Clear()
            Me.cmdAcadExams.Parameters.AddWithValue("@termName", Me.cboTermView.Text.Trim)
            Me.cmdAcadExams.Parameters.AddWithValue("@classYear", Me.cboYearView.Text.Trim)
            Me.cmdAcadExams.Parameters.AddWithValue("@stream", Me.cboStream.Text.Trim)
            Me.cmdAcadExams.Parameters.AddWithValue("@className", Me.cboClass.Text.Trim)
            Me.cmdAcadExams.Parameters.AddWithValue("@subName", String.Format("%{0}%", TryCast(Me.txtSearchSubject.Text.Trim, String).Trim))

            reader = Me.cmdAcadExams.ExecuteReader
            If reader.HasRows Then
                While reader.Read
                   li = Me.lstViewExams.Items.Add(IIf(DBNull.Value.Equals(reader!subName), "", reader!subName))
                    li.SubItems.Add(IIf(DBNull.Value.Equals(reader!examType), "", reader!examType))
                    li.SubItems.Add(IIf(DBNull.Value.Equals(reader!totalTypeMark), "", reader!totalTypeMark))
                    li.SubItems.Add(IIf(DBNull.Value.Equals(reader!examName), "", reader!examName))
                    li.SubItems.Add(IIf(DBNull.Value.Equals(reader!totalMarks), "", reader!totalMarks))
                    li.SubItems.Add(IIf(DBNull.Value.Equals(reader!partTypeMark), "", reader!partTypeMark))
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
    End Sub

    Private Sub CLOSEToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CLOSEToolStripMenuItem.Click
        Me.Close()
    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        If Me.cboExamName.Text.Trim.Length <= 0 Then
            MsgBox("Exam Name Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboTotalMarks.Text.Trim.Length <= 0 Then
            MsgBox("Exam Total Marks Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboExamContr.Text.Trim.Length <= 0 Then
            MsgBox("End Term part Type Mark Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboTotalTypeMarks.Text.Trim.Length <= 0 Then
            MsgBox("Exam Total Type Marks Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboTermUpdate.Text.Trim.Length <= 0 Then
            MsgBox("Term Name Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboYearUpdate.Text.Trim.Length <= 0 Then
            MsgBox("Year Name Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.lstVwSubjectDetails.Items.Count <= 0 Then
            MsgBox("No subjects registered." & vbNewLine & "Contact the system administrator.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.lstClassDetails.Items.Count <= 0 Then
            MsgBox("No classes registered." & vbNewLine & "Contact the system administrator.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.lstVwSubjectDetails.CheckedItems.Count <= 0 Then
            MsgBox("No subjects checked.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.lstClassDetails.CheckedItems.Count <= 0 Then
            MsgBox("No classes checked.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            Dim result As MsgBoxResult = MsgBox("Save Record?", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.YesNo, "Confirm Transaction")
            If result = MsgBoxResult.No Then
                Exit Sub
            End If
            rec = 0
            i = 0
            j = 0
            maxrec = Me.lstVwSubjectDetails.CheckedItems.Count - 1
            maxrec1 = Me.lstClassDetails.CheckedItems.Count - 1
            For i = 0 To maxrec
                j = 0
                For j = 0 To maxrec1
                    queryType = "INSERT"
                    Me.recordExists = True
                    checkIfExists()
                    If Me.recordExists = True Then
                        MsgBox("Entry For Subject " & Me.lstVwSubjectDetails.CheckedItems(i).Text.Trim & " CLASS " & _
                               Me.lstClassDetails.CheckedItems(j).Text & " STREAM " & Me.lstClassDetails.CheckedItems(j).SubItems(1).Text.Trim & " Not Saved." & _
                               vbNewLine & "The Record Already Exists in Database" & vbNewLine & "Click Ok To Continue Saving Record/s", _
                               MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "System Reply")
                    End If
                    If recordExists = False Then
                        
                        Me.cmdAcadExams.Connection = conn
                        Me.cmdAcadExams.CommandText = "sprocClassExams"
                        Me.cmdAcadExams.CommandType = CommandType.StoredProcedure
                        Me.cmdAcadExams.Parameters.Clear()
                        Me.cmdAcadExams.Parameters.AddWithValue("@className", Me.lstClassDetails.CheckedItems(j).Text.Trim)
                        Me.cmdAcadExams.Parameters.AddWithValue("@stream", Me.lstClassDetails.CheckedItems(j).SubItems(1).Text.Trim)
                        Me.cmdAcadExams.Parameters.AddWithValue("@year", Me.cboYearUpdate.Text.Trim)
                        Me.cmdAcadExams.Parameters.AddWithValue("@termName", Me.cboTermUpdate.Text.Trim)
                        Me.cmdAcadExams.Parameters.AddWithValue("@subName", Me.lstVwSubjectDetails.CheckedItems(i).Text.Trim)
                        Me.cmdAcadExams.Parameters.AddWithValue("@queryType", Me.queryType.Trim)
                        Me.cmdAcadExams.Parameters.AddWithValue("@dateOfReg", Date.Now)
                        Me.cmdAcadExams.Parameters.AddWithValue("@regBy", userName.Trim)
                        Me.cmdAcadExams.Parameters.AddWithValue("@examName", Me.cboExamName.Text.Trim)
                        Me.cmdAcadExams.Parameters.AddWithValue("@totalMarks", Me.cboTotalMarks.Text.Trim)
                        Me.cmdAcadExams.Parameters.AddWithValue("@endTermMark", Me.cboExamContr.Text.Trim)
                        Me.cmdAcadExams.Parameters.AddWithValue("@totalTypeMark", Me.cboTotalTypeMarks.Text.Trim)
                        Me.cmdAcadExams.Parameters.AddWithValue("@examType", Me.cboExamType.Text.Trim)
                        rec = rec + Me.cmdAcadExams.ExecuteNonQuery
                    End If
                Next
            Next
            If rec > 0 Then
                MsgBox("Record/s Saved", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "SuccessFull Transaction")
                clearTexts()
                loadCombos()
                loadLists()
                Me.lstVwSubjectDetails.Enabled = True
                Me.lstClassDetails.Enabled = True
                Me.cboExamName.Enabled = True
                Me.cboTermUpdate.Enabled = True
                Me.cboYearUpdate.Enabled = True
                Me.btnSave.Enabled = True
                Me.btnUpdate.Enabled = False
            End If
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub
    Private Sub clearTexts()
        Me.lstClassDetails.Items.Clear()
        Me.lstViewExams.Items.Clear()
        Me.lstVwSubjectDetails.Items.Clear()

        Me.txtSearchSubject.Text = ""
        Me.cboClass.Items.Clear()
        Me.cboExamContr.Items.Clear()
        Me.cboTotalMarks.Items.Clear()
        Me.cboTotalTypeMarks.Items.Clear()
        Me.cboExamName.Items.Clear()
        Me.cboStream.Items.Clear()
        Me.cboTermUpdate.Items.Clear()
        Me.cboTermView.Items.Clear()
        Me.cboYearUpdate.Items.Clear()
        Me.cboYearView.Items.Clear()

        Me.cboClass.Text = ""
        Me.cboExamContr.Text = ""
        Me.cboExamName.Text = ""
        Me.cboStream.Text = ""
        Me.cboTermUpdate.Text = ""
        Me.cboTermView.Text = ""
        Me.cboTotalMarks.Text = ""
        Me.cboYearUpdate.Text = ""
        Me.cboYearView.Text = ""
        Me.cboTotalTypeMarks.Text = ""

        Me.cboClass.SelectedIndex = -1
        Me.cboExamContr.SelectedIndex = -1
        Me.cboExamName.SelectedIndex = -1
        Me.cboStream.SelectedIndex = -1
        Me.cboTermUpdate.SelectedIndex = -1
        Me.cboTermView.SelectedIndex = -1
        Me.cboTotalMarks.SelectedIndex = -1
        Me.cboYearUpdate.SelectedIndex = -1
        Me.cboYearView.SelectedIndex = -1
        Me.cboTotalTypeMarks.SelectedIndex = -1
    End Sub
    Private Sub checkIfExists()
        Me.cmdAcadExams.Connection = conn
        Me.cmdAcadExams.CommandText = "SELECT * FROM vwAcadClassExams WHERE  (termName=@termName) AND (classYear=@classYear) AND " & _
                vbNewLine & " (className=@className) AND (stream=@stream) AND (subName=@subName) AND (examName=@examName)"
        Me.cmdAcadExams.CommandType = CommandType.Text
        Me.cmdAcadExams.Parameters.Clear()
        Me.cmdAcadExams.Parameters.AddWithValue("@className", Me.lstClassDetails.CheckedItems(j).Text.Trim)
        Me.cmdAcadExams.Parameters.AddWithValue("@stream", Me.lstClassDetails.CheckedItems(j).SubItems(1).Text.Trim)
        Me.cmdAcadExams.Parameters.AddWithValue("@classYear", Me.cboYearUpdate.Text.Trim)
        Me.cmdAcadExams.Parameters.AddWithValue("@termName", Me.cboTermUpdate.Text.Trim)
        Me.cmdAcadExams.Parameters.AddWithValue("@subName", Me.lstVwSubjectDetails.CheckedItems(i).Text.Trim)
        Me.cmdAcadExams.Parameters.AddWithValue("@examName", Me.cboExamName.Text.Trim)
        reader = Me.cmdAcadExams.ExecuteReader
        If reader.HasRows = True Then
            Me.recordExists = True
        Else
            Me.recordExists = False
        End If
        reader.Close()
    End Sub

    Private Sub UPDATEToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UPDATEToolStripMenuItem.Click
        If Me.lstViewExams.Items.Count <= 0 Then
            MsgBox("No items in the list.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.lstViewExams.SelectedItems.Count <= 0 Then
            MsgBox("Select the item to update.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.lstViewExams.SelectedItems.Count > 1 Then
            MsgBox("Select One item at a time.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If
        If Me.lstViewExams.SelectedItems.Count = 1 Then
            Me.cboExamType.Text = Me.lstViewExams.SelectedItems(0).SubItems(1).Text.Trim
            Me.cboTotalTypeMarks.Text = Me.lstViewExams.SelectedItems(0).SubItems(2).Text.Trim
            Me.cboExamName.Text = Me.lstViewExams.SelectedItems(0).SubItems(3).Text.Trim
            Me.cboTotalMarks.Text = Me.lstViewExams.SelectedItems(0).SubItems(4).Text.Trim
            Me.cboExamContr.Text = Me.lstViewExams.SelectedItems(0).SubItems(5).Text.Trim
            Me.cboTermUpdate.Text = Me.cboTermView.Text.Trim
            Me.cboYearUpdate.Text = Me.cboYearView.Text.Trim

            i = 0
            recordExists = False
            For i = 0 To Me.lstVwSubjectDetails.Items.Count - 1
                If Me.lstVwSubjectDetails.Items(i).Text.Trim = Me.lstViewExams.SelectedItems(0).Text.Trim Then
                    recordExists = True
                End If
            Next
            If recordExists = True Then
                Me.lstVwSubjectDetails.Items.Clear()
                li = Me.lstVwSubjectDetails.Items.Add(Me.lstViewExams.SelectedItems(0).Text.Trim)
                Me.lstVwSubjectDetails.Items(0).Checked = True
            End If

            recordExists = False
            i = 0
            For i = 0 To Me.lstClassDetails.Items.Count - 1
                If (Me.lstClassDetails.Items(i).Text.Trim = Me.cboClass.Text.Trim) And ( _
                    Me.lstClassDetails.Items(i).SubItems(1).Text.Trim = Me.cboStream.Text.Trim) Then
                    recordExists = True
                End If
            Next
            If recordExists = True Then
                Me.lstClassDetails.Items.Clear()
                li = Me.lstClassDetails.Items.Add(Me.cboClass.Text.Trim)
                li.SubItems.Add(Me.cboStream.Text.Trim)
                Me.lstClassDetails.Items(0).Checked = True
            End If
            Me.lstVwSubjectDetails.Enabled = False
            Me.lstClassDetails.Enabled = False
            Me.cboExamType.Enabled = False
            Me.cboExamName.Enabled = False
            Me.cboTermUpdate.Enabled = False
            Me.cboYearUpdate.Enabled = False
            Me.btnSave.Enabled = False
            Me.btnUpdate.Enabled = True
            Me.lstViewExams.Items.Clear()
        End If
    End Sub

    Private Sub btnUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdate.Click
        If Me.cboExamName.Text.Trim.Length <= 0 Then
            MsgBox("Exam Name Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboTotalMarks.Text.Trim.Length <= 0 Then
            MsgBox("Total Marks Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboExamContr.Text.Trim.Length <= 0 Then
            MsgBox("End Term Mark contribution Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboTermUpdate.Text.Trim.Length <= 0 Then
            MsgBox("Term Name Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboYearUpdate.Text.Trim.Length <= 0 Then
            MsgBox("Year Name Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.lstVwSubjectDetails.Items.Count <= 0 Then
            MsgBox("No subjects registered." & vbNewLine & "Contact the system administrator.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.lstClassDetails.Items.Count <= 0 Then
            MsgBox("No classes registered." & vbNewLine & "Contact the system administrator.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.lstVwSubjectDetails.CheckedItems.Count <= 0 Then
            MsgBox("No subjects checked.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.lstClassDetails.CheckedItems.Count <= 0 Then
            MsgBox("No classes checked.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            Me.recordExists = True
            checkIfExistsOne()
            If Me.recordExists = True Then
                MsgBox("Record Exists.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Duplicate Detected")
                Exit Sub
            End If
            Dim result As MsgBoxResult = MsgBox("Update Record?", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.YesNo, "Confirm Transaction")
            If result = MsgBoxResult.No Then
                Exit Sub
            End If
            queryType = "UPDATE"
            Me.cmdAcadExams.Connection = conn
            Me.cmdAcadExams.CommandText = "sprocClassExams"
            Me.cmdAcadExams.CommandType = CommandType.StoredProcedure
            Me.cmdAcadExams.Parameters.Clear()
            Me.cmdAcadExams.Parameters.AddWithValue("@className", Me.lstClassDetails.CheckedItems(0).Text.Trim)
            Me.cmdAcadExams.Parameters.AddWithValue("@stream", Me.lstClassDetails.CheckedItems(0).SubItems(1).Text.Trim)
            Me.cmdAcadExams.Parameters.AddWithValue("@year", Me.cboYearUpdate.Text.Trim)
            Me.cmdAcadExams.Parameters.AddWithValue("@termName", Me.cboTermUpdate.Text.Trim)
            Me.cmdAcadExams.Parameters.AddWithValue("@subName", Me.lstVwSubjectDetails.CheckedItems(0).Text.Trim)
            Me.cmdAcadExams.Parameters.AddWithValue("@queryType", Me.queryType.Trim)
            Me.cmdAcadExams.Parameters.AddWithValue("@dateOfReg", Date.Now)
            Me.cmdAcadExams.Parameters.AddWithValue("@regBy", userName.Trim)
            Me.cmdAcadExams.Parameters.AddWithValue("@examName", Me.cboExamName.Text.Trim)
            Me.cmdAcadExams.Parameters.AddWithValue("@totalMarks", Me.cboTotalMarks.Text.Trim)
            Me.cmdAcadExams.Parameters.AddWithValue("@endTermMark", Me.cboExamContr.Text.Trim)
            Me.cmdAcadExams.Parameters.AddWithValue("@totalTypeMark", Me.cboTotalTypeMarks.Text.Trim)
            Me.cmdAcadExams.Parameters.AddWithValue("@examType", Me.cboExamType.Text.Trim)
            rec = Me.cmdAcadExams.ExecuteNonQuery
            MsgBox("Record Updated", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "SuccessFull Transaction")
            Me.lstVwSubjectDetails.Enabled = True
            Me.lstClassDetails.Enabled = True
            Me.cboExamName.Enabled = True
            Me.cboExamType.Enabled = True
            Me.cboTermUpdate.Enabled = True
            Me.cboYearUpdate.Enabled = True
            Me.btnSave.Enabled = True
            Me.btnUpdate.Enabled = False
            clearTexts()
            loadCombos()
            loadLists()
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub
    Private Sub checkIfExistsOne()
        Me.cmdAcadExams.Connection = conn
        Me.cmdAcadExams.CommandText = "SELECT * FROM vwAcadClassExams WHERE (termName=@termName) AND (classYear=@classYear) AND " & _
                vbNewLine & " (className=@className) AND (stream=@stream) AND (subName=@subName) AND (totalMarks=@totalMarks) AND " & _
                vbNewLine & " (totalTypeMark=@totalTypeMark) AND (partTypeMark=@partTypeMark) AND (examName=@examName)"
        Me.cmdAcadExams.CommandType = CommandType.Text
        Me.cmdAcadExams.Parameters.Clear()
        Me.cmdAcadExams.Parameters.AddWithValue("@className", Me.lstClassDetails.CheckedItems(0).Text.Trim)
        Me.cmdAcadExams.Parameters.AddWithValue("@stream", Me.lstClassDetails.CheckedItems(0).SubItems(1).Text.Trim)
        Me.cmdAcadExams.Parameters.AddWithValue("@classYear", Me.cboYearUpdate.Text.Trim)
        Me.cmdAcadExams.Parameters.AddWithValue("@termName", Me.cboTermUpdate.Text.Trim)
        Me.cmdAcadExams.Parameters.AddWithValue("@subName", Me.lstVwSubjectDetails.CheckedItems(0).Text.Trim)
        Me.cmdAcadExams.Parameters.AddWithValue("@examName", Me.cboExamName.Text.Trim)
        Me.cmdAcadExams.Parameters.AddWithValue("@totalMarks", Me.cboTotalMarks.Text.Trim)
        Me.cmdAcadExams.Parameters.AddWithValue("@totalTypeMark", Me.cboTotalTypeMarks.Text.Trim)
        Me.cmdAcadExams.Parameters.AddWithValue("@partTypeMark", Me.cboExamContr.Text.Trim)
        reader = Me.cmdAcadExams.ExecuteReader
        If reader.HasRows Then
            Me.recordExists = True
        Else
            Me.recordExists = False
        End If
        reader.Close()
    End Sub

    Private Sub btnClear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClear.Click
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            Me.lstVwSubjectDetails.Enabled = True
            Me.lstClassDetails.Enabled = True
            Me.cboExamType.Enabled = True
            Me.cboExamName.Enabled = True
            Me.cboTermUpdate.Enabled = True
            Me.cboYearUpdate.Enabled = True
            Me.btnSave.Enabled = True
            Me.btnUpdate.Enabled = False
            clearTexts()
            loadCombos()
            loadLists()
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub

    Private Sub DELETEToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DELETEToolStripMenuItem.Click
        If Me.lstViewExams.Items.Count <= 0 Then
            MsgBox("No items in the list.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.lstViewExams.SelectedItems.Count <= 0 Then
            MsgBox("Select the item to update.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.lstViewExams.SelectedItems.Count > 1 Then
            MsgBox("Select One item at a time.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If
        If Me.lstViewExams.SelectedItems.Count = 1 Then
            Try
                If conn.State = ConnectionState.Closed Then
                    conn.Open()
                End If
                dbconnection()
                Dim result As MsgBoxResult = MsgBox("Delete Record?", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.YesNo, "Confirm Transaction")
                If result = MsgBoxResult.No Then
                    Exit Sub
                End If
                queryType = "DELETE"
                Me.cmdAcadExams.Connection = conn
                Me.cmdAcadExams.CommandText = "sprocClassExams"
                Me.cmdAcadExams.CommandType = CommandType.StoredProcedure
                Me.cmdAcadExams.Parameters.Clear()
                Me.cmdAcadExams.Parameters.AddWithValue("@className", Me.cboClass.Text.Trim)
                Me.cmdAcadExams.Parameters.AddWithValue("@stream", Me.cboStream.Text.Trim)
                Me.cmdAcadExams.Parameters.AddWithValue("@year", Me.cboYearView.Text.Trim)
                Me.cmdAcadExams.Parameters.AddWithValue("@termName", Me.cboTermView.Text.Trim)
                Me.cmdAcadExams.Parameters.AddWithValue("@subName", Me.lstViewExams.SelectedItems(0).Text.Trim)
                Me.cmdAcadExams.Parameters.AddWithValue("@queryType", Me.queryType.Trim)
                Me.cmdAcadExams.Parameters.AddWithValue("@dateOfReg", Date.Now)
                Me.cmdAcadExams.Parameters.AddWithValue("@regBy", userName.Trim)
                Me.cmdAcadExams.Parameters.AddWithValue("@examName", Me.lstViewExams.SelectedItems(0).SubItems(3).Text.Trim)
                Me.cmdAcadExams.Parameters.AddWithValue("@totalMarks", Me.lstViewExams.SelectedItems(0).SubItems(4).Text.Trim)
                Me.cmdAcadExams.Parameters.AddWithValue("@endTermMark", Me.lstViewExams.SelectedItems(0).SubItems(5).Text.Trim)
                Me.cmdAcadExams.Parameters.AddWithValue("@totalTypeMark", Me.lstViewExams.SelectedItems(0).SubItems(2).Text.Trim)
                Me.cmdAcadExams.Parameters.AddWithValue("@examType", Me.lstViewExams.SelectedItems(0).SubItems(1).Text.Trim)
                rec = Me.cmdAcadExams.ExecuteNonQuery
                MsgBox("Record Deleted", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "SuccessFull Transaction")
                Me.lstVwSubjectDetails.Enabled = True
                Me.lstClassDetails.Enabled = True
                Me.cboExamName.Enabled = True
                Me.cboTermUpdate.Enabled = True
                Me.cboYearUpdate.Enabled = True
                Me.btnSave.Enabled = True
                Me.btnUpdate.Enabled = False
                clearTexts()
                loadCombos()
                loadLists()
            Catch ex As Exception
                MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Finally
                If conn.State = ConnectionState.Open Then
                    conn.Close()
                End If
            End Try
        End If
    End Sub

    Private Sub cboExamType_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboExamType.SelectedIndexChanged
        If Me.cboExamType.Text.Trim.Length <= 0 Then
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            Me.cboExamName.Items.Clear()
            Me.cboExamName.Text = ""
            Me.cboExamName.SelectedIndex = -1

            Me.cboTotalTypeMarks.Items.Clear()
            Me.cboTotalTypeMarks.Text = ""
            Me.cboTotalTypeMarks.SelectedIndex = -1

            Me.cmdAcadExams.Connection = conn
            Me.cmdAcadExams.CommandType = CommandType.Text
            Me.cmdAcadExams.CommandText = "SELECT examName FROM tblExamNames WHERE (status=1) AND (examType=@examType) ORDER BY examName"
            Me.cmdAcadExams.Parameters.Clear()
            Me.cmdAcadExams.Parameters.AddWithValue("@examType", Me.cboExamType.Text.Trim)
            reader = Me.cmdAcadExams.ExecuteReader
            While reader.Read
                Me.cboExamName.Items.Add(IIf(DBNull.Value.Equals(reader!examName), "", (reader!examName)))
            End While
            reader.Close()

            Me.cmdAcadExams.CommandText = "SELECT DISTINCT totalTypeMark FROM vwAcadClassExams WHERE (examType=@examType) " & _
                vbNewLine & " ORDER BY totalTypeMark"
            Me.cmdAcadExams.Parameters.Clear()
            Me.cmdAcadExams.Parameters.AddWithValue("@examType", Me.cboExamType.Text.Trim)
            reader = Me.cmdAcadExams.ExecuteReader
            While reader.Read
                Me.cboTotalTypeMarks.Items.Add(IIf(DBNull.Value.Equals(reader!totalTypeMark), "", (reader!totalTypeMark)))
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

    Private Sub cboStream_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboStream.SelectedIndexChanged
        Me.lstViewExams.Items.Clear()
    End Sub

    Private Sub cboTermView_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboTermView.SelectedIndexChanged
        Me.lstViewExams.Items.Clear()
    End Sub

    Private Sub cboClass_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboClass.SelectedIndexChanged
        Me.lstViewExams.Items.Clear()
    End Sub

End Class