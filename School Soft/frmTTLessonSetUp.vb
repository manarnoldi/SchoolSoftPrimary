Imports System.Data.SqlClient
Imports Excel = Microsoft.Office.Interop.Excel
Public Class frmTTLessonSetUp
    Dim xlApp As Excel.Application
    Dim xlWorkBook As Excel.Workbook
    Dim xlWorkSheet As Excel.Worksheet
    Dim range As Excel.Range
    Dim fileName As String = Nothing
    Dim sheetName As String = Nothing
    Dim cmdTTLesson As New SqlCommand
    Dim rec As Integer
    Dim lessonTypeint As Integer = 0
    Dim reader As SqlDataReader
    Private Sub frmTTLessonSetUp_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
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

        Me.cboStream.Items.Clear()
        Me.cboStream.Text = ""
        Me.cboStream.SelectedIndex = -1

        Me.cboClassName.Items.Clear()
        Me.cboClassName.Text = ""
        Me.cboClassName.SelectedIndex = -1

        Me.cboSubjectName.Items.Clear()
        Me.cboSubjectName.Text = ""
        Me.cboSubjectName.SelectedIndex = -1

        Me.cboRoomType.Items.Clear()
        Me.cboRoomType.Text = ""
        Me.cboRoomType.SelectedIndex = -1

        Me.cmdTTLesson.Connection = conn
        Me.cmdTTLesson.CommandType = CommandType.Text
        Me.cmdTTLesson.CommandText = "SELECT DISTINCT year FROM tblClasses WHERE (status=1) ORDER BY year"
        Me.cmdTTLesson.Parameters.Clear()
        reader = Me.cmdTTLesson.ExecuteReader
        While reader.Read
            Me.cboYear.Items.Add(IIf(DBNull.Value.Equals(reader!year), "", reader!year))
        End While
        reader.Close()

        Me.cmdTTLesson.CommandText = "SELECT DISTINCT stream FROM tblClasses WHERE (status=1) ORDER BY stream"
        Me.cmdTTLesson.Parameters.Clear()
        reader = Me.cmdTTLesson.ExecuteReader
        While reader.Read
            Me.cboStream.Items.Add(IIf(DBNull.Value.Equals(reader!stream), "", reader!stream))
        End While
        reader.Close()

        Me.cmdTTLesson.CommandText = "SELECT DISTINCT className FROM tblClasses WHERE (status=1) ORDER BY className"
        Me.cmdTTLesson.Parameters.Clear()
        reader = Me.cmdTTLesson.ExecuteReader
        While reader.Read
            Me.cboClassName.Items.Add(IIf(DBNull.Value.Equals(reader!className), "", reader!className))
        End While
        reader.Close()

        Me.cmdTTLesson.CommandText = "SELECT DISTINCT subName,subCode FROM tblSubjects WHERE (subStatus=1) ORDER BY subCode"
        Me.cmdTTLesson.Parameters.Clear()
        reader = Me.cmdTTLesson.ExecuteReader
        While reader.Read
            Me.cboSubjectName.Items.Add(IIf(DBNull.Value.Equals(reader!subName), "", reader!subName))
        End While
        reader.Close()

        Me.cmdTTLesson.CommandText = "SELECT DISTINCT roomType FROM  tblSchoolRooms WHERE (status=1) ORDER BY roomType"
        Me.cmdTTLesson.Parameters.Clear()
        reader = Me.cmdTTLesson.ExecuteReader
        While reader.Read
            Me.cboRoomType.Items.Add(IIf(DBNull.Value.Equals(reader!roomType), "", reader!roomType))
        End While
        reader.Close()

    End Sub
    Private Sub frmTTLessonSetUp_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint
        If Me.IsMdiChild Then
            Dim PnlLoc As New Point
            PnlLoc.X = CInt((Me.Width - Me.pnlLessonSetUp.Width) / 2)
            PnlLoc.Y = CInt((Me.Height - Me.pnlLessonSetUp.Height) / 2.5)
            Me.pnlLessonSetUp.Location = PnlLoc
        Else
            Me.pnlLessonSetUp.Dock = DockStyle.Fill
        End If
    End Sub


    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub


    Private Sub cboStaffNo_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboStaffNo.SelectedIndexChanged
        Me.txtStaffName.Text = ""
        If Me.cboStaffNo.Text.Trim.Length <= 0 Then
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()

            Me.cmdTTLesson.Connection = conn
            Me.cmdTTLesson.CommandType = CommandType.Text
            Me.cmdTTLesson.CommandText = "SELECT FullName FROM tblSchoolStaff WHERE (status=1) and (empNo=@empNo)"
            Me.cmdTTLesson.Parameters.Clear()
            Me.cmdTTLesson.Parameters.AddWithValue("empNo", Me.cboStaffNo.Text.Trim)
            reader = Me.cmdTTLesson.ExecuteReader
            While reader.Read
                Me.txtStaffName.Text = (IIf(DBNull.Value.Equals(reader!FullName), "", reader!FullName))
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

    Private Sub cboSubjectName_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles _
        cboSubjectName.SelectedIndexChanged, cboClassName.SelectedIndexChanged, cboStream.SelectedIndexChanged, cboYear.SelectedIndexChanged
        Me.txtSubAbbr.Text = ""
        Me.txtSubCode.Text = ""
        Me.cboStaffNo.Items.Clear()
        Me.cboStaffNo.Text = ""
        Me.cboStaffNo.SelectedIndex = -1
        Me.txtStaffName.Text = ""

        If Me.cboSubjectName.Text.Trim.Length <= 0 Then
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()

            Me.cmdTTLesson.Connection = conn
            Me.cmdTTLesson.CommandType = CommandType.Text
            Me.cmdTTLesson.CommandText = "SELECT subCode,abbr FROM tblSubjects WHERE (subStatus=1) and (subName=@subName)"
            Me.cmdTTLesson.Parameters.Clear()
            Me.cmdTTLesson.Parameters.AddWithValue("@subName", Me.cboSubjectName.Text.Trim)
            reader = Me.cmdTTLesson.ExecuteReader
            While reader.Read
                Me.txtSubCode.Text = (IIf(DBNull.Value.Equals(reader!subCode), "", reader!subCode))
                Me.txtSubAbbr.Text = (IIf(DBNull.Value.Equals(reader!abbr), "", reader!abbr))
            End While
            reader.Close()
            If Me.cboYear.Text.Trim.Length > 0 And Me.cboClassName.Text.Trim.Length > 0 And Me.cboStream.Text.Trim.Length > 0 Then
                Me.cmdTTLesson.CommandText = "SELECT DISTINCT empNo FROM vwAcadStaffSubjects WHERE (className=@className) AND (stream=@stream) " & _
                    vbNewLine & " AND (year=@year) AND (subName=@subName)"
                Me.cmdTTLesson.Parameters.Clear()
                Me.cmdTTLesson.Parameters.AddWithValue("@className", Me.cboClassName.Text.Trim)
                Me.cmdTTLesson.Parameters.AddWithValue("@stream", Me.cboStream.Text.Trim)
                Me.cmdTTLesson.Parameters.AddWithValue("@year", Me.cboYear.Text.Trim)
                Me.cmdTTLesson.Parameters.AddWithValue("@subName", Me.cboSubjectName.Text.Trim)
                reader = Me.cmdTTLesson.ExecuteReader
                While reader.Read
                    Me.cboStaffNo.Items.Add(IIf(DBNull.Value.Equals(reader!empNo), "", reader!empNo))
                End While
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

    Private Sub cboRoomType_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboRoomType.SelectedIndexChanged
        Me.cboRoomName.Items.Clear()
        Me.cboRoomName.Text = ""
        Me.cboRoomName.SelectedIndex = -1

        If Me.cboRoomType.Text.Trim.Length <= 0 Then
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()

            Me.cmdTTLesson.Connection = conn
            Me.cmdTTLesson.CommandType = CommandType.Text
            Me.cmdTTLesson.CommandText = "SELECT roomName FROM tblSchoolRooms WHERE (status=1) AND (roomType=@roomType)"
            Me.cmdTTLesson.Parameters.Clear()
            Me.cmdTTLesson.Parameters.AddWithValue("@roomType", Me.cboRoomType.Text.Trim)
            reader = Me.cmdTTLesson.ExecuteReader
            While reader.Read
                Me.cboRoomName.Items.Add(IIf(DBNull.Value.Equals(reader!roomName), "", reader!roomName))
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

    Private Sub btnAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAdd.Click
        If Me.cboYear.Text.Trim.Length <= 0 Then
            MsgBox("Missing year.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboClassName.Text.Trim.Length <= 0 Then
            MsgBox("Missing Class Name.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboStream.Text.Trim.Length <= 0 Then
            MsgBox("Missing Stream.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboStaffNo.Text.Trim.Length <= 0 Then
            MsgBox("Missing Staff Number.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.txtStaffName.Text.Trim.Length <= 0 Then
            MsgBox("Missing Staff Name.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub

        ElseIf Me.cboSubjectName.Text.Trim.Length <= 0 Then
            MsgBox("Missing Subject Name.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboLessonType.Text.Trim.Length <= 0 Then
            MsgBox("Missing Lesson Type.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboTypeCount.Text.Trim.Length <= 0 Then
            MsgBox("Missing Lesson Type Count.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboRoomName.Text.Trim.Length <= 0 Then
            MsgBox("Missing Lesson Room Name.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If
        If Me.lstLessonSetUp.Items.Count >= 1 Then
            For i = 0 To Me.lstLessonSetUp.Items.Count - 1
                If Me.lstLessonSetUp.Items(i).Text.Trim = Me.cboYear.Text.Trim And _
                    Me.lstLessonSetUp.Items(i).SubItems(1).Text.Trim = Me.cboClassName.Text.Trim And _
                    Me.lstLessonSetUp.Items(i).SubItems(2).Text.Trim = Me.cboStream.Text.Trim And _
                    Me.lstLessonSetUp.Items(i).SubItems(5).Text.Trim = Me.cboSubjectName.Text.Trim And _
                    Me.lstLessonSetUp.Items(i).SubItems(3).Text.Trim = Me.cboStaffNo.Text.Trim And _
                    Me.lstLessonSetUp.Items(i).SubItems(6).Text.Trim = Me.cboLessonType.Text.Trim And _
                    Me.lstLessonSetUp.Items(i).SubItems(8).Text.Trim = Me.cboRoomName.Text.Trim Then
                    MsgBox("Duplicate record found in the list.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
                    Exit Sub
                End If
            Next
        End If

        li = Me.lstLessonSetUp.Items.Add(Me.cboYear.Text.Trim)
        li.SubItems.Add(Me.cboClassName.Text.Trim)
        li.SubItems.Add(Me.cboStream.Text.Trim)
        li.SubItems.Add(Me.cboStaffNo.Text.Trim)
        li.SubItems.Add(Me.txtStaffName.Text.Trim)

        li.SubItems.Add(Me.cboSubjectName.Text.Trim)
        li.SubItems.Add(Me.cboLessonType.Text.Trim)
        li.SubItems.Add(Me.cboTypeCount.Text.Trim)
        li.SubItems.Add(Me.cboRoomName.Text.Trim)
    End Sub

    Private Sub btnRemove_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRemove.Click
        If Me.lstLessonSetUp.Items.Count <= 0 Then
            MsgBox("Item list is empty.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.lstLessonSetUp.SelectedItems.Count <= 0 Then
            MsgBox("No item selected from the list.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If
        For i = 0 To Me.lstLessonSetUp.SelectedItems.Count - 1
            Me.lstLessonSetUp.SelectedItems(0).Remove()
        Next
    End Sub
    Private Function checkifStaffiSAllocated(ByVal empNo As String, ByVal subName As String, ByVal className As String, ByVal stream As String, ByVal year As Integer)
        Me.cmdTTLesson.Connection = conn
        Me.cmdTTLesson.CommandType = CommandType.Text
        Me.cmdTTLesson.CommandText = "SELECT * FROM vwAcadStaffSubjects WHERE (className=@className) AND (stream=@stream) AND " & _
            vbNewLine & " (year=@year) AND (subName=@subName) AND (empNo=@empNo)"
        Me.cmdTTLesson.Parameters.Clear()
        Me.cmdTTLesson.Parameters.AddWithValue("@empNo", empNo.Trim)
        Me.cmdTTLesson.Parameters.AddWithValue("@className", className.Trim)
        Me.cmdTTLesson.Parameters.AddWithValue("@stream", stream.Trim)
        Me.cmdTTLesson.Parameters.AddWithValue("@year", year)
        Me.cmdTTLesson.Parameters.AddWithValue("@subName", subName)
        reader = Me.cmdTTLesson.ExecuteReader
        If reader.HasRows = True Then
            Return True
        ElseIf reader.HasRows = False Then
            Return False
        End If
        reader.Close()
        conn.Close()
    End Function
    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click

        If Me.lstLessonSetUp.Items.Count <= 0 Then
            MsgBox("Item list is empty.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If
        Try
            
            'For j = 0 To Me.lstLessonSetUp.Items.Count - 1
            '    If conn.State = ConnectionState.Closed Then
            '        conn.Open()
            '    End If
            '    dbconnection()
            '    Dim teacherAllocDone As Boolean = checkifStaffiSAllocated(Me.lstLessonSetUp.Items(j).SubItems(3).Text.Trim, Me.lstLessonSetUp.Items(j).SubItems(5).Text.Trim, _
            '                            Me.lstLessonSetUp.Items(j).SubItems(1).Text.Trim, Me.lstLessonSetUp.Items(j).SubItems(2).Text.Trim, _
            '                            Me.lstLessonSetUp.Items(j).Text.Trim)
            '    If teacherAllocDone = False Then
            '        MsgBox("SUBJECT-TEACHER ALLOCATION FOR YEAR " & Me.lstLessonSetUp.Items(j).Text.Trim & " STAFF NAME " & _
            '               Me.lstLessonSetUp.Items(j).SubItems(4).Text.Trim & " SUBJECT " & _
            '        Me.lstLessonSetUp.Items(j).SubItems(5).Text.Trim & " STREAM " & Me.lstLessonSetUp.Items(j).SubItems(2).Text.Trim _
            '        & " CLASS " & Me.lstLessonSetUp.Items(j).SubItems(1).Text.Trim & " IS NOT DONE." & _
            '        vbNewLine & "CANNOT PROCEED WITH SAVING.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly + _
            '        MsgBoxStyle.ApplicationModal, "MISSING LINK")
            '        Exit Sub
            '    End If
            'Next


            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            Dim result As MsgBoxResult = MsgBox("Save Record/s?", MsgBoxStyle.Question + MsgBoxStyle.YesNo + MsgBoxStyle.ApplicationModal, "Confirm Transaction")
            If result = MsgBoxResult.No Then
                Exit Sub
            End If
            For i = 0 To Me.lstLessonSetUp.Items.Count - 1
                If Me.lstLessonSetUp.Items(i).SubItems(6).Text.Trim = "Single" Then
                    lessonTypeint = 1
                ElseIf Me.lstLessonSetUp.Items(i).SubItems(6).Text.Trim = "Double" Then
                    lessonTypeint = 2
                ElseIf Me.lstLessonSetUp.Items(i).SubItems(6).Text.Trim = "Triple" Then
                    lessonTypeint = 3
                End If
                Dim recordExists As Boolean = checkIfRecordExists(Me.lstLessonSetUp.Items(i).Text.Trim, _
                                                                  Me.lstLessonSetUp.Items(i).SubItems(1).Text.Trim, _
                                                                  Me.lstLessonSetUp.Items(i).SubItems(2).Text.Trim, _
                                                                  Me.lstLessonSetUp.Items(i).SubItems(5).Text.Trim, _
                                                                  Me.lstLessonSetUp.Items(i).SubItems(3).Text.Trim,
                                                                  lessonTypeint, Me.lstLessonSetUp.Items(i).SubItems(8).Text.Trim)
                If conn.State = ConnectionState.Closed Then
                    conn.Open()
                End If
                dbconnection()
                If recordExists = True Then
                    MsgBox("RECORD FOR YEAR " & Me.lstLessonSetUp.Items(i).Text.Trim & " " & Me.lstLessonSetUp.Items(i).SubItems(1).Text.Trim & " " & _
                    Me.lstLessonSetUp.Items(i).SubItems(2).Text.Trim & " FOR STAFF " & Me.lstLessonSetUp.Items(i).SubItems(3).Text.Trim & " SUBJECT " & _
                    Me.lstLessonSetUp.Items(i).SubItems(5).Text.Trim & " LESSON TYPE " & Me.lstLessonSetUp.Items(i).SubItems(6).Text.Trim _
                    & " CLASS ROOM " & Me.lstLessonSetUp.Items(i).SubItems(8).Text.Trim & " WILL NOT BE SAVED." & _
                    vbNewLine & "DUPLICATE FOUND IN THE SYSTEM", MsgBoxStyle.Information + MsgBoxStyle.OkOnly + _
                    MsgBoxStyle.ApplicationModal, "DUPLICATE FOUND")
                Else
                    Me.cmdTTLesson.Connection = conn
                    Me.cmdTTLesson.CommandType = CommandType.StoredProcedure
                    Me.cmdTTLesson.CommandText = "sprocTTLessonSetUp"
                    Me.cmdTTLesson.Parameters.Clear()
                    Me.cmdTTLesson.Parameters.AddWithValue("@className", Me.lstLessonSetUp.Items(i).SubItems(1).Text.Trim)
                    Me.cmdTTLesson.Parameters.AddWithValue("@stream", Me.lstLessonSetUp.Items(i).SubItems(2).Text.Trim)
                    Me.cmdTTLesson.Parameters.AddWithValue("@year", Me.lstLessonSetUp.Items(i).Text.Trim)
                    Me.cmdTTLesson.Parameters.AddWithValue("@subjectName", Me.lstLessonSetUp.Items(i).SubItems(5).Text.Trim)
                    Me.cmdTTLesson.Parameters.AddWithValue("@roomName", Me.lstLessonSetUp.Items(i).SubItems(8).Text.Trim)
                    Me.cmdTTLesson.Parameters.AddWithValue("@empNo", Me.lstLessonSetUp.Items(i).SubItems(3).Text.Trim)
                    Me.cmdTTLesson.Parameters.AddWithValue("@lessonType", lessonTypeint)
                    Me.cmdTTLesson.Parameters.AddWithValue("@lessonTypeCount", Me.lstLessonSetUp.Items(i).SubItems(7).Text.Trim)
                    Me.cmdTTLesson.Parameters.AddWithValue("@userName", userName.Trim)
                    Me.cmdTTLesson.Parameters.AddWithValue("@queryType", "INSERT")
                    rec = rec + Me.cmdTTLesson.ExecuteNonQuery
                End If
            Next
            If rec > 0 Then
                MsgBox("Record/s Saved.", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "System Reply")
            End If
            Me.lstLessonSetUp.Items.Clear()
            loadCombos()
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub
    Private Function checkIfRecordExists(ByVal year As Integer, ByVal className As String, ByVal stream As String, _
                                         ByVal subname As String, ByVal staffno As String, ByVal lessontype As Integer, _
                                         ByVal roomName As String)
        Me.cmdTTLesson.Connection = conn
        Me.cmdTTLesson.CommandType = CommandType.Text
        Me.cmdTTLesson.CommandText = "SELECT * FROM vwTTLessonSetUp WHERE (className=@className) AND (stream=@stream) " & _
            vbNewLine & " AND (year=@year) AND (empNo=@empNo) AND (subName=@subName) AND (lessonType=@lessonType) AND (roomName=@roomName)"
        Me.cmdTTLesson.Parameters.Clear()
        Me.cmdTTLesson.Parameters.AddWithValue("@className", className.Trim)
        Me.cmdTTLesson.Parameters.AddWithValue("@stream", stream.Trim)
        Me.cmdTTLesson.Parameters.AddWithValue("@year", year)
        Me.cmdTTLesson.Parameters.AddWithValue("@empNo", staffno.Trim)
        Me.cmdTTLesson.Parameters.AddWithValue("@subName", subname.Trim)
        Me.cmdTTLesson.Parameters.AddWithValue("@lessonType", lessontype)
        Me.cmdTTLesson.Parameters.AddWithValue("@roomName", roomName.Trim)
        reader = Me.cmdTTLesson.ExecuteReader
        If reader.HasRows = True Then
            Return True
        ElseIf reader.HasRows = False Then
            Return False
        End If
        reader.Close()
        conn.Close()
    End Function
    Private Sub loadlist()
        Dim lessonType As String = Nothing
        Me.lstLessonSetUp.Items.Clear()
        Me.cmdTTLesson.Connection = conn
        Me.cmdTTLesson.CommandType = CommandType.Text
        Me.cmdTTLesson.CommandText = "SELECT * FROM vwTTLessonSetUp WHERE (year=@year)"
        Me.cmdTTLesson.Parameters.Clear()
        Me.cmdTTLesson.Parameters.AddWithValue("@year", Me.cboYear.Text.Trim)
        reader = Me.cmdTTLesson.ExecuteReader
        While reader.Read
            If IIf(DBNull.Value.Equals(reader!lessonType), "", reader!lessonType) = 1 Then
                lessonType = "Single"
            ElseIf IIf(DBNull.Value.Equals(reader!lessonType), "", reader!lessonType) = 2 Then
                lessonType = "Double"
            ElseIf IIf(DBNull.Value.Equals(reader!lessonType), "", reader!lessonType) = 3 Then
                lessonType = "Triple"
            End If
            li = Me.lstLessonSetUp.Items.Add(IIf(DBNull.Value.Equals(reader!year), "", reader!year))
            li.SubItems.Add(IIf(DBNull.Value.Equals(reader!className), "", reader!className))
            li.SubItems.Add(IIf(DBNull.Value.Equals(reader!stream), "", reader!stream))
            li.SubItems.Add(IIf(DBNull.Value.Equals(reader!empNo), "", reader!empNo))
            li.SubItems.Add(IIf(DBNull.Value.Equals(reader!FullName), "", reader!FullName))
            li.SubItems.Add(IIf(DBNull.Value.Equals(reader!subName), "", reader!subName))
            li.SubItems.Add(lessonType.Trim)
            li.SubItems.Add(IIf(DBNull.Value.Equals(reader!lessonTypeCount), "", reader!lessonTypeCount))
            li.SubItems.Add(IIf(DBNull.Value.Equals(reader!roomName), "", reader!roomName))
        End While
        reader.Close()
    End Sub

    Private Sub btnView_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnView.Click
        If Me.cboYear.Text.Trim.Length <= 0 Then
            MsgBox("Select academic year.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            loadlist()
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

    Private Sub DELETEToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DELETEToolStripMenuItem.Click
        If Me.lstLessonSetUp.Items.Count <= 0 Then
            MsgBox("Item list is empty.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.lstLessonSetUp.SelectedItems.Count <= 0 Then
            MsgBox("Select items to delete.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            Dim result As MsgBoxResult = MsgBox("Delete Record/s?", MsgBoxStyle.Question + MsgBoxStyle.YesNo + MsgBoxStyle.ApplicationModal, "Confirm Transaction")
            If result = MsgBoxResult.No Then
                Exit Sub
            End If
            For i = 0 To Me.lstLessonSetUp.SelectedItems.Count - 1
                If Me.lstLessonSetUp.SelectedItems(i).SubItems(6).Text.Trim = "Single" Then
                    lessonTypeint = 1
                ElseIf Me.lstLessonSetUp.SelectedItems(i).SubItems(6).Text.Trim = "Double" Then
                    lessonTypeint = 2
                ElseIf Me.lstLessonSetUp.SelectedItems(i).SubItems(6).Text.Trim = "Triple" Then
                    lessonTypeint = 3
                End If
                Me.cmdTTLesson.Connection = conn
                Me.cmdTTLesson.CommandType = CommandType.StoredProcedure
                Me.cmdTTLesson.CommandText = "sprocTTLessonSetUp"
                Me.cmdTTLesson.Parameters.Clear()
                Me.cmdTTLesson.Parameters.AddWithValue("@className", Me.lstLessonSetUp.SelectedItems(i).SubItems(1).Text.Trim)
                Me.cmdTTLesson.Parameters.AddWithValue("@stream", Me.lstLessonSetUp.SelectedItems(i).SubItems(2).Text.Trim)
                Me.cmdTTLesson.Parameters.AddWithValue("@year", Me.lstLessonSetUp.SelectedItems(i).Text.Trim)
                Me.cmdTTLesson.Parameters.AddWithValue("@subjectName", Me.lstLessonSetUp.SelectedItems(i).SubItems(5).Text.Trim)
                Me.cmdTTLesson.Parameters.AddWithValue("@roomName", Me.lstLessonSetUp.SelectedItems(i).SubItems(8).Text.Trim)
                Me.cmdTTLesson.Parameters.AddWithValue("@empNo", Me.lstLessonSetUp.SelectedItems(i).SubItems(3).Text.Trim)
                Me.cmdTTLesson.Parameters.AddWithValue("@lessonType", lessonTypeint)
                Me.cmdTTLesson.Parameters.AddWithValue("@lessonTypeCount", Me.lstLessonSetUp.SelectedItems(i).SubItems(7).Text.Trim)
                Me.cmdTTLesson.Parameters.AddWithValue("@userName", userName.Trim)
                Me.cmdTTLesson.Parameters.AddWithValue("@queryType", "DELETE")
                rec = rec + Me.cmdTTLesson.ExecuteNonQuery
            Next
            If rec > 0 Then
                MsgBox("Records Deleted.", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "System Reply")
            End If
            Me.lstLessonSetUp.Items.Clear()
            loadCombos()
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
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
        If Me.lstLessonSetUp.Items.Count <= 0 Then
            MsgBox("No items in the list.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If
        Dim result As MsgBoxResult = MsgBox("Export Records To Excel?", MsgBoxStyle.Question + MsgBoxStyle.ApplicationModal + MsgBoxStyle.YesNo, "Confirm Transaction")
        If result = MsgBoxResult.No Then
            Exit Sub
        End If
        Me.Cursor = Cursors.WaitCursor

        Dim xlApp As Excel.Application
        Dim xlWorkBook As Excel.Workbook
        Dim xlWorkSheet As Excel.Worksheet
        Dim misValue As Object = System.Reflection.Missing.Value
        Dim i As Integer

        xlApp = New Excel.Application
        xlWorkBook = xlApp.Workbooks.Add(misValue)
        xlWorkSheet = xlWorkBook.Sheets("Sheet1")
        Dim col As Integer = 1
        For j As Integer = 0 To Me.lstLessonSetUp.Columns.Count - 1
            xlWorkSheet.Cells(1, col) = Me.lstLessonSetUp.Columns(j).Text.Trim
            col = col + 1
        Next


        For i = 0 To Me.lstLessonSetUp.Items.Count - 1
            xlWorkSheet.Cells(i + 2, 1) = Me.lstLessonSetUp.Items(i).Text.Trim
            xlWorkSheet.Cells(i + 2, 2) = Me.lstLessonSetUp.Items(i).SubItems(1).Text.Trim
            xlWorkSheet.Cells(i + 2, 3) = Me.lstLessonSetUp.Items(i).SubItems(2).Text.Trim
            xlWorkSheet.Cells(i + 2, 4) = Me.lstLessonSetUp.Items(i).SubItems(3).Text.Trim
            xlWorkSheet.Cells(i + 2, 5) = Me.lstLessonSetUp.Items(i).SubItems(4).Text.Trim
            xlWorkSheet.Cells(i + 2, 6) = Me.lstLessonSetUp.Items(i).SubItems(5).Text.Trim
            xlWorkSheet.Cells(i + 2, 7) = Me.lstLessonSetUp.Items(i).SubItems(6).Text.Trim
            xlWorkSheet.Cells(i + 2, 8) = Me.lstLessonSetUp.Items(i).SubItems(7).Text.Trim
            xlWorkSheet.Cells(i + 2, 9) = Me.lstLessonSetUp.Items(i).SubItems(8).Text.Trim
        Next
        Dim dlg As New SaveFileDialog
        dlg.Filter = "Excel Files (*.xlsx)|*.xlsx"
        dlg.FilterIndex = 1
        dlg.InitialDirectory = My.Application.Info.DirectoryPath & "\EXCEL\\EICHER\BILLS\"
        dlg.FileName = "TTLessonSetup"
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
        MsgBox(i & " Records Exported!", MsgBoxStyle.Information + MsgBoxStyle.OkOnly + MsgBoxStyle.ApplicationModal, "System Reply")
        Me.lstLessonSetUp.Items.Clear()
    End Sub

    Private Sub btnUpload_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpload.Click
        Try
            Dim result As MsgBoxResult = MsgBox("Import Excel?", MsgBoxStyle.Question + MsgBoxStyle.ApplicationModal + MsgBoxStyle.YesNo, "Confirm Transaction")
            If result = MsgBoxResult.No Then
                Exit Sub
            End If
            Me.OpenFileDialog.ShowDialog()
            fileName = Me.OpenFileDialog.FileName
            Me.lstLessonSetUp.Items.Clear()
            sheetName = System.IO.Path.GetFileNameWithoutExtension(fileName)

            If Me.fileName.Trim = Nothing Then
                MsgBox("Missing file Name", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
                Exit Sub
            End If

            If fileName.Contains(sheetName) Then
                xlApp = New Excel.Application
                xlWorkBook = xlApp.Workbooks.Open(fileName)
                xlWorkSheet = xlWorkBook.Worksheets(sheetName)
            Else
                MsgBox("SheetName not contained in the file Name", MsgBoxStyle.Information, "Error in selection")
                Exit Sub
            End If
            Me.Cursor = Windows.Forms.Cursors.WaitCursor
            MsgBox("Wait Until System Gives You A Reply" & _
            vbNewLine & "Click Ok To Continue.", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Please Wait")
            range = TryCast(xlWorkSheet, Excel.Worksheet).UsedRange
            k = range.Rows.Count - 1
            Dim OrderList As New ArrayList
            i = 0
            Me.lstLessonSetUp.Items.Clear()
            For i = 2 To range.Rows.Count
                Dim LstItem As New ListViewItem
                j = 1
                LstItem.Text = (IIf(IsNothing((CType(range.Cells(i, j), Excel.Range).Value)) = True, "", CType(range.Cells(i, j), Excel.Range).Value))
                j = j + 1
                LstItem.SubItems.Add(IIf(IsNothing((CType(range.Cells(i, j), Excel.Range).Value)) = True, "", CType(range.Cells(i, j), Excel.Range).Value))
                j = j + 1
                LstItem.SubItems.Add(IIf(IsNothing((CType(range.Cells(i, j), Excel.Range).Value)) = True, "", CType(range.Cells(i, j), Excel.Range).Value))
                j = j + 1
                LstItem.SubItems.Add(IIf(IsNothing((CType(range.Cells(i, j), Excel.Range).Value)) = True, "", CType(range.Cells(i, j), Excel.Range).Value))
                j = j + 1
                LstItem.SubItems.Add(IIf(IsNothing((CType(range.Cells(i, j), Excel.Range).Value)) = True, "", CType(range.Cells(i, j), Excel.Range).Value))
                j = j + 1
                LstItem.SubItems.Add(IIf(IsNothing((CType(range.Cells(i, j), Excel.Range).Value)) = True, "", CType(range.Cells(i, j), Excel.Range).Value))
                j = j + 1
                LstItem.SubItems.Add(IIf(IsNothing((CType(range.Cells(i, j), Excel.Range).Value)) = True, "", CType(range.Cells(i, j), Excel.Range).Value))
                j = j + 1
                LstItem.SubItems.Add(IIf(IsNothing((CType(range.Cells(i, j), Excel.Range).Value)) = True, "", CType(range.Cells(i, j), Excel.Range).Value))
                j = j + 1
                LstItem.SubItems.Add(IIf(IsNothing((CType(range.Cells(i, j), Excel.Range).Value)) = True, "", CType(range.Cells(i, j), Excel.Range).Value))
                j = 1
                OrderList.Add(LstItem)
                LstItem = Nothing
            Next

            For Each OrderItem As ListViewItem In TryCast(OrderList, ArrayList)
                Me.lstLessonSetUp.Items.Add(OrderItem)
            Next

            MsgBox("" & i - 2 & " -Records uploaded!", MsgBoxStyle.Information, "Records Uploaded")
            xlWorkBook.Close()
            xlApp.Quit()
            releaseObject(xlApp)
            releaseObject(xlWorkBook)
            releaseObject(xlWorkSheet)
        Catch
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            Me.Cursor = Windows.Forms.Cursors.Arrow
        End Try
    End Sub

    Private Sub btnClear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClear.Click
        Me.lstLessonSetUp.Items.Clear()
    End Sub
End Class