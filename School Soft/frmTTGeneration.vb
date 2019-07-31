Imports System.Data.SqlClient
Public Class frmTTGeneration
    Dim lessonAbbr As String = Nothing
    Dim className As String = Nothing
    Dim classStream As String = Nothing
    Dim classYear As Integer = 0
    Dim employeeNo As String = Nothing
    Dim weekDayNo As Integer = 0
    Dim lessonId As Integer = 0
    Dim subjColor As String = Nothing
    Dim lessonStaffCount As Integer = 0
    Dim maxDoubleLessPerDay As Integer = 0
    Dim minNo As Integer = 0
    Shared yearGlobal As Integer = 0
    Dim spaceCount As Integer = 0
    Dim lessonCount As Integer = 0
    Dim m As Integer = 0
    Dim selectedSpace() As String = Nothing
    Dim spaceDic As New Dictionary(Of Integer, Array)
    Dim freeSpaceCount As Integer = 0
    Dim rowNum As Integer = 0
    Dim reader As SqlDataReader
    Dim reader1 As SqlDataReader
    Shared cmdGeneration As New SqlCommand
    Dim rec As Integer
    Dim totalColumns As Integer
    Dim totalRows As Integer
    Dim daysperweek As Integer
    Dim dayoftheweek As String
    Private Sub frmTTGeneration_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint
        If Me.IsMdiChild Then
            Dim PnlLoc As New Point
            PnlLoc.X = CInt((Me.Width - Me.pnlTTGeneration.Width) / 2)
            PnlLoc.Y = CInt((Me.Height - Me.pnlTTGeneration.Height) / 2.5)
            Me.pnlTTGeneration.Location = PnlLoc
        Else
            Me.pnlTTGeneration.Dock = DockStyle.Fill
        End If
    End Sub
    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub
    Private Sub btngenStr_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btngenStr.Click
        rowNum = 0
        Me.dgTTGeneration.Rows.Clear()
        Me.dgTTGeneration.Columns.Clear()
        If Me.cboYear.Text.Trim.Length <= 0 Then
            MsgBox("Academic Year is missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()

            Dim periodsExist As Boolean = checkifperiodsExist()
            If periodsExist = False Then
                MsgBox("Register periods first.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
                Exit Sub
            End If

            Dim result As MsgBoxResult = MsgBox("Generate Timetable?", MsgBoxStyle.Question + MsgBoxStyle.YesNo + MsgBoxStyle.ApplicationModal, "Confirm Transaction")
            If result = MsgBoxResult.No Then
                Exit Sub
            End If

            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            cmdGeneration.CommandText = "SELECT (lessonsPerDay+breaksPerDay)+1 AS totalPeriods FROM tblTTSetUp WHERE (academicYear=@academicYear)"
            cmdGeneration.Parameters.Clear()
            cmdGeneration.Parameters.AddWithValue("@academicYear", Me.cboYear.Text.Trim)
            reader = cmdGeneration.ExecuteReader
            While reader.Read
                totalColumns = (IIf(DBNull.Value.Equals(reader!totalPeriods), "", reader!totalPeriods))
            End While
            reader.Close()

            cmdGeneration.CommandText = "SELECT (SELECT COUNT(DISTINCT ClassName)FROM tblClasses WHERE (year=@Year) AND " & _
                vbNewLine & " (status=1)) * (SELECT COUNT(DISTINCT Stream) FROM tblClasses WHERE (year=@Year) AND " & _
                vbNewLine & " (status=1)) * (SELECT daysPerWeek FROM tblTTSetUp WHERE (academicYear=@Year)) AS countRow"
            cmdGeneration.Parameters.Clear()
            cmdGeneration.Parameters.AddWithValue("@Year", Me.cboYear.Text.Trim)
            reader = cmdGeneration.ExecuteReader
            While reader.Read
                totalRows = (IIf(DBNull.Value.Equals(reader!countRow), "", reader!countRow))
            End While
            reader.Close()

            Me.dgTTGeneration.ColumnCount = totalColumns + 4
            Me.dgTTGeneration.RowCount = totalRows

            Me.dgTTGeneration.Columns(0).Name = "DAY"
            Me.dgTTGeneration.Columns(0).HeaderText = "DAY"
            Me.dgTTGeneration.Columns(0).Width = 100
            Me.dgTTGeneration.Columns(0).ReadOnly = True
            Me.dgTTGeneration.Columns(0).SortMode = DataGridViewColumnSortMode.NotSortable
            Me.dgTTGeneration.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

            Me.dgTTGeneration.Columns(1).Name = "CLASS"
            Me.dgTTGeneration.Columns(1).HeaderText = "CLASS"
            Me.dgTTGeneration.Columns(1).Width = 70
            Me.dgTTGeneration.Columns(1).ReadOnly = True
            Me.dgTTGeneration.Columns(1).SortMode = DataGridViewColumnSortMode.NotSortable
            Me.dgTTGeneration.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

            Me.dgTTGeneration.Columns(2).Name = "STREAM"
            Me.dgTTGeneration.Columns(2).HeaderText = "STREAM"
            Me.dgTTGeneration.Columns(2).Width = 70
            Me.dgTTGeneration.Columns(2).ReadOnly = True
            Me.dgTTGeneration.Columns(2).SortMode = DataGridViewColumnSortMode.NotSortable
            Me.dgTTGeneration.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

            Me.dgTTGeneration.Columns(3).Name = "YEAR"
            Me.dgTTGeneration.Columns(3).HeaderText = "YEAR"
            Me.dgTTGeneration.Columns(3).Width = 70
            Me.dgTTGeneration.Columns(3).ReadOnly = True
            Me.dgTTGeneration.Columns(3).SortMode = DataGridViewColumnSortMode.NotSortable
            Me.dgTTGeneration.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

            cmdGeneration.CommandText = "SELECT * FROM tblTTPeriods WHERE (academicYear=@academicYear)"
            cmdGeneration.Parameters.Clear()
            cmdGeneration.Parameters.AddWithValue("@academicYear", Me.cboYear.Text.Trim)
            reader = cmdGeneration.ExecuteReader
            j = 4
            While reader.Read
                Me.dgTTGeneration.Columns(j).Name = (IIf(DBNull.Value.Equals(reader!periodName), "", reader!periodName)).ToString.ToUpper
                Me.dgTTGeneration.Columns(j).HeaderText = (IIf(DBNull.Value.Equals(reader!periodName), "", reader!periodName)).ToString.ToUpper
                Me.dgTTGeneration.Columns(j).Width = 70
                Me.dgTTGeneration.Columns(j).ReadOnly = True
                Me.dgTTGeneration.Columns(j).SortMode = DataGridViewColumnSortMode.NotSortable
                Me.dgTTGeneration.Columns(j).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                j = j + 1
            End While
            reader.Close()

            cmdGeneration.CommandText = "SELECT * FROM tblTTSetUp WHERE (academicYear=@academicYear)"
            cmdGeneration.Parameters.Clear()
            cmdGeneration.Parameters.AddWithValue("@academicYear", Me.cboYear.Text.Trim)
            reader = cmdGeneration.ExecuteReader
            While reader.Read
                daysperweek = (IIf(DBNull.Value.Equals(reader!daysPerWeek), "", reader!daysPerWeek))
            End While
            reader.Close()

            i = 0
            j = 0
            k = 0
            For j = 1 To daysperweek
                If j = 1 Then
                    dayoftheweek = "MONDAY"
                ElseIf j = 2 Then
                    dayoftheweek = "TUESDAY"
                ElseIf j = 3 Then
                    dayoftheweek = "WEDNESDAY"
                ElseIf j = 4 Then
                    dayoftheweek = "THURSDAY"
                ElseIf j = 5 Then
                    dayoftheweek = "FRIDAY"
                ElseIf j = 6 Then
                    dayoftheweek = "SATURDAY"
                ElseIf j = 7 Then
                    dayoftheweek = "SUNDAY"
                End If
                loadClasses(dayoftheweek)
            Next

            i = 0
            j = 0
            For i = 4 To Me.dgTTGeneration.Columns.Count - 1
                If Not (Me.dgTTGeneration.Columns(i).HeaderText.Contains("LESSON")) Then
                    j = 0
                    For j = 0 To Me.dgTTGeneration.Rows.Count - 1
                        Me.dgTTGeneration.Item(i, j).Value = Me.dgTTGeneration.Columns(i).HeaderText
                    Next
                End If
            Next

        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub
    Private Function loadClasses(ByVal weekDayName As String)
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            cmdGeneration.Connection = conn
            cmdGeneration.CommandType = CommandType.Text
            cmdGeneration.CommandText = "SELECT DISTINCT className FROM tblClasses WHERE (Year=@Year) AND (status=1) ORDER BY className"
            cmdGeneration.Parameters.Clear()
            cmdGeneration.Parameters.AddWithValue("@Year", Me.cboYear.Text.Trim)
            reader = cmdGeneration.ExecuteReader
            While reader.Read
                loadStreams(weekDayName, (IIf(DBNull.Value.Equals(reader!className), "", reader!className)))
            End While
        Catch ex As Exception
        Finally
            reader.Close()
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Function
    Private Function loadStreams(ByVal weekDayName As String, ByVal className As String)
        Try
            dbconnection1()
            If conn1.State = ConnectionState.Closed Then
                conn1.Open()
            End If
            cmdGeneration.Connection = conn1
            cmdGeneration.CommandType = CommandType.Text
            cmdGeneration.CommandText = "SELECT DISTINCT Stream FROM tblClasses WHERE (Year=@Year) AND (status=1) ORDER BY stream"
            cmdGeneration.Parameters.Clear()
            cmdGeneration.Parameters.AddWithValue("@Year", Me.cboYear.Text.Trim)
            reader1 = cmdGeneration.ExecuteReader
            While reader1.Read
                Me.dgTTGeneration.Item(0, rowNum).Value = weekDayName
                Me.dgTTGeneration.Item(1, rowNum).Value = className
                Me.dgTTGeneration.Item(2, rowNum).Value = (IIf(DBNull.Value.Equals(reader1!Stream), "", reader1!Stream))
                Me.dgTTGeneration.Item(3, rowNum).Value = (Me.cboYear.Text.Trim)
                rowNum = rowNum + 1
            End While
        Catch ex As Exception
        Finally
            reader1.Close()
            If conn1.State = ConnectionState.Open Then
                conn1.Close()
            End If
        End Try
    End Function
    Private Sub loadCombos()
        Me.cboYear.Items.Clear()
        Me.cboYear.Text = ""
        Me.cboYear.SelectedIndex = -1
        cmdGeneration.Connection = conn
        cmdGeneration.CommandType = CommandType.Text
        cmdGeneration.CommandText = "SELECT DISTINCT year FROM tblSchoolCalendar WHERE (Status=1) ORDER BY year"
        cmdGeneration.Parameters.Clear()
        reader = cmdGeneration.ExecuteReader
        While reader.Read
            Me.cboYear.Items.Add(IIf(DBNull.Value.Equals(reader!year), "", reader!year))
        End While
        reader.Close()
    End Sub
    Private Sub frmTTGeneration_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
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
    Private Function checkifperiodsExist()
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            cmdGeneration.Connection = conn
            cmdGeneration.CommandType = CommandType.Text
            cmdGeneration.CommandText = "SELECT * FROM tblTTPeriods WHERE (academicYear=@academicYear)"
            cmdGeneration.Parameters.Clear()
            cmdGeneration.Parameters.AddWithValue("@academicYear", Me.cboYear.Text.Trim)
            reader = cmdGeneration.ExecuteReader
            If reader.HasRows = True Then
                Return True
            ElseIf reader.HasRows = False Then
                Return False
            End If
        Catch ex As Exception
        Finally
            reader.Close()
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Function

    Private Sub btnGentt_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnGentt.Click
        If Me.dgTTGeneration.Rows.Count = 0 Or Me.dgTTGeneration.Columns.Count = 0 Then
            MsgBox("Generate structure first.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If


        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            cmdGeneration.Connection = conn
            cmdGeneration.CommandType = CommandType.Text
            cmdGeneration.CommandText = "SELECT value FROM tblTTGenParameters " & _
                vbNewLine & " WHERE (acadYear=@acadYear) AND (parameterName='SET MAXIMUM LESSONS PER DAY')"
            cmdGeneration.Parameters.Clear()
            cmdGeneration.Parameters.AddWithValue("@acadYear", Me.cboYear.Text.Trim)
            reader = cmdGeneration.ExecuteReader
            While reader.Read
                lessonStaffCount = IIf(DBNull.Value.Equals(reader!value), "", reader!value)
            End While
            reader.Close()

            If lessonStaffCount <= 0 Then
                MsgBox("Right Click and SET MAXIMUM LESSONS PER DAY.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
                Exit Sub
            End If

            cmdGeneration.CommandText = "SELECT value FROM tblTTGenParameters " & _
               vbNewLine & " WHERE (acadYear=@acadYear) AND (parameterName='SET MAXIMUM DOUBLES PER DAY')"
            cmdGeneration.Parameters.Clear()
            cmdGeneration.Parameters.AddWithValue("@acadYear", Me.cboYear.Text.Trim)
            reader = cmdGeneration.ExecuteReader
            While reader.Read
                maxDoubleLessPerDay = IIf(DBNull.Value.Equals(reader!value), "", reader!value)
            End While
            reader.Close()

            If maxDoubleLessPerDay <= 0 Then
                MsgBox("Right Click and SET MAXIMUM DOUBLE LESSONS PER DAY.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
                Exit Sub
            End If

            Dim result As MsgBoxResult = MsgBox("Generate TimeTable?", MsgBoxStyle.Question + MsgBoxStyle.YesNo + MsgBoxStyle.ApplicationModal, "Confirm Transaction")
            If result = MsgBoxResult.No Then
                Exit Sub
            End If

            minNo = 0
            Me.Cursor = Cursors.WaitCursor
            For i = 4 To Me.dgTTGeneration.Columns.Count - 1
                If (Me.dgTTGeneration.Columns(i).HeaderText.Contains("LESSON")) Then
                    j = 0
                    For j = 0 To Me.dgTTGeneration.Rows.Count - 1
                        Me.dgTTGeneration.Item(i, j).Value = Nothing
                        Me.dgTTGeneration.Item(i, j).Style.BackColor = Color.White
                    Next
                End If
            Next

            cmdGeneration.CommandText = "UPDATE tblTTLessonSetUp SET lessonsInserted=0 WHERE (academicYear=@academicYear)"
            cmdGeneration.Parameters.Clear()
            cmdGeneration.Parameters.AddWithValue("@academicYear", Me.cboYear.Text.Trim)
            cmdGeneration.ExecuteNonQuery()
            spaceCount = countFreeSpaces(Me.dgTTGeneration)
            lessonCount = countFreeLessons()
            yearGlobal = Me.cboYear.Text.Trim
            If Not (bgWorkerGen.IsBusy) Then
                bgWorkerGen.RunWorkerAsync(Me.dgTTGeneration)
            End If

        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            Me.Cursor = Cursors.Arrow
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub
    Private Function selectFreeSpace(ByVal className As String, ByVal classStream As String, ByVal year As Integer, ByVal dgvw As DataGridView)
        Dim dictionaryKey As Integer = 0
        Dim s() As String = Nothing
        spaceDic.Clear()
        dictionaryKey = 1
        k = 0
        m = 0
        freeSpaceCount = 0
        For k = 0 To dgvw.Columns.Count - 1
            m = 0
            For m = 0 To dgvw.Rows.Count - 1
                If dgvw.Item(1, m).Value = className And dgvw.Item(2, m).Value = classStream And _
                    dgvw.Item(3, m).Value = year Then
                    If dgvw.Item(k, m).Value = "" Then
                        s = {k, m}
                        spaceDic.Add(dictionaryKey, s)
                        dictionaryKey = dictionaryKey + 1
                    End If
                End If
            Next
        Next
        Dim rnd As New Random
        Dim found As Integer = rnd.Next(1, dictionaryKey)
        If spaceDic.ContainsKey(found) Then
            selectedSpace = spaceDic.Item(found)
        End If
    End Function
    Private Function countFreeSpaces(ByVal dgvw As DataGridView)
        k = 0
        m = 0
        freeSpaceCount = 0
        For k = 0 To dgvw.Columns.Count - 1
            m = 0
            For m = 0 To dgvw.Rows.Count - 1
                If dgvw.Item(k, m).Value = "" Then
                    freeSpaceCount = freeSpaceCount + 1
                End If
            Next
        Next
        Return freeSpaceCount
    End Function
    Private Function countFreeLessons()
        Dim freelessonCount As Integer = 0
        Try
            dbconnection()
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            cmdGeneration.Connection = conn
            cmdGeneration.CommandType = CommandType.Text
            cmdGeneration.CommandText = "SELECT ISNULL(SUM(LESSONSREMAIN),0) AS lessonCount FROM tblttlessonsetup " & _
                vbNewLine & " WHERE (lessonsRemain>0)"
            cmdGeneration.Parameters.Clear()
            reader = cmdGeneration.ExecuteReader
            If reader.HasRows = True Then
                While reader.Read
                    freelessonCount = IIf(DBNull.Value.Equals(reader!lessonCount), "", reader!lessonCount)
                End While
                Return freelessonCount
            ElseIf reader.HasRows = False Then
                Return 0
            End If
        Catch ex As Exception
        Finally
            reader.Close()
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Function
    Private Sub pickLesson()
        lessonAbbr = Nothing
        className = Nothing
        classStream = Nothing
        classYear = 0
        employeeNo = Nothing
        lessonId = 0
        subjColor = Nothing
        Try
            dbconnection()
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            cmdGeneration.Connection = conn
            cmdGeneration.CommandType = CommandType.Text
            cmdGeneration.CommandText = "SELECT TOP 1 * FROM vwTTGenerationLessons WHERE (Year=@academicYear) AND " & _
                vbNewLine & " (lessonsRemain>0) ORDER BY NEWID()"
            cmdGeneration.Parameters.Clear()
            cmdGeneration.Parameters.AddWithValue("@academicYear", yearGlobal)
            reader = cmdGeneration.ExecuteReader
            While reader.Read
                lessonAbbr = IIf(DBNull.Value.Equals(reader!abbr), "", reader!abbr)
                className = IIf(DBNull.Value.Equals(reader!className), "", reader!className)
                classStream = IIf(DBNull.Value.Equals(reader!stream), "", reader!stream)
                classYear = IIf(DBNull.Value.Equals(reader!year), "", reader!year)
                employeeNo = IIf(DBNull.Value.Equals(reader!empNo), "", reader!empNo)
                lessonId = IIf(DBNull.Value.Equals(reader!lessonId), "", reader!lessonId)
                subjColor = IIf(DBNull.Value.Equals(reader!subColourName), "", reader!subColourName)
            End While
        Catch ex As Exception
        Finally
            reader.Close()
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub
    Private Function checkSubjectConstraint(ByVal subAbbr As String, ByVal lessonName As String, ByVal weekdayno As Integer)
        dbconnection()
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If
        Try
            cmdGeneration.Connection = conn
            cmdGeneration.CommandType = CommandType.Text
            cmdGeneration.CommandText = "SELECT * FROM vwTTSubjectSetUp WHERE (acadYear=@Year) AND (abbr=@abbr) AND " & _
                vbNewLine & " (lessonName=@lessonName) AND (weekDayNo=@weekDayNo) AND (allowed='YES')"
            cmdGeneration.Parameters.Clear()
            cmdGeneration.Parameters.AddWithValue("@Year", yearGlobal)
            cmdGeneration.Parameters.AddWithValue("@abbr", subAbbr.Trim)
            cmdGeneration.Parameters.AddWithValue("@lessonName", lessonName.Trim)
            cmdGeneration.Parameters.AddWithValue("@weekDayNo", weekdayno)
            reader = cmdGeneration.ExecuteReader
            If reader.HasRows = True Then
                Return True
            ElseIf reader.HasRows = False Then
                Return False
            End If
        Catch ex As Exception
        Finally
            reader.Close()
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Function
    Private Function checkStaffConstraint(ByVal employeeNumber As String, ByVal lessonName As String, ByVal weekdayno As Integer)
        Dim staffIsReg As Boolean = False
        dbconnection()
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If
        Try
            cmdGeneration.Connection = conn
            cmdGeneration.CommandType = CommandType.Text
            cmdGeneration.CommandText = "SELECT DISTINCT empNo,academicYear FROM vwTTStaffSetUp WHERE " & _
                vbNewLine & "(academicYear=@Year) AND (empNo=@empNo)"
            cmdGeneration.Parameters.Clear()
            cmdGeneration.Parameters.AddWithValue("@Year", yearGlobal)
            cmdGeneration.Parameters.AddWithValue("@empNo", employeeNumber.Trim)
            reader1 = cmdGeneration.ExecuteReader
            If reader1.HasRows = True Then
                staffIsReg = True
            ElseIf reader1.HasRows = False Then
                staffIsReg = False
            End If
            reader1.Close()

            If staffIsReg = False Then
                Return True
            ElseIf staffIsReg = True Then
                cmdGeneration.CommandText = "SELECT * FROM vwTTStaffSetUp WHERE (academicYear=@Year) AND (empNo=@empNo) AND " & _
                vbNewLine & " (lessonName=@lessonName) AND (weekDayNo=@weekDayNo) AND ((allowed='YES'))"
                cmdGeneration.Parameters.Clear()
                cmdGeneration.Parameters.AddWithValue("@Year", yearGlobal)
                cmdGeneration.Parameters.AddWithValue("@empNo", employeeNumber.Trim)
                cmdGeneration.Parameters.AddWithValue("@lessonName", lessonName.Trim)
                cmdGeneration.Parameters.AddWithValue("@weekDayNo", weekdayno)
                reader = cmdGeneration.ExecuteReader
                If reader.HasRows = True Then
                    Return True
                ElseIf reader.HasRows = False Then
                    Return False
                End If
            End If
        Catch ex As Exception
        Finally

            reader.Close()
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Function
    Private Function checkTeacherClashes(ByVal lessonName As String, ByVal staffno As String, ByVal weekDayName As String, ByVal dgvw As DataGridView)
        Dim staffRowCount As Integer = 0
        Dim staffFound As Boolean = False
        For k = 0 To dgvw.ColumnCount - 1
            If dgvw.Columns(k).HeaderText = lessonName Then
                For staffRowCount = 0 To dgvw.RowCount - 1
                    Dim cellvalue As String = (dgvw.Item(k, staffRowCount).Value)
                    If Not (cellvalue = "") Then
                        If cellvalue.Contains(staffno) = True And dgvw.Item(0, staffRowCount).Value = weekDayName.Trim Then
                            staffFound = True
                            Return staffFound
                            Exit Function
                        End If
                    End If
                Next
            End If
        Next
        Return staffFound
    End Function
    Private Function lessonNoPerDay(ByVal rowNum As Integer, ByVal staffSubj As String, ByVal dgvw As DataGridView)
        Dim lessonPerDay As Integer = 0
        Dim staffLessNoColumn As Integer = 0
        For staffLessNoColumn = 0 To dgvw.Columns.Count - 1
            If Not (dgvw.Rows(rowNum).Cells(staffLessNoColumn).Value = "") Then
                If dgvw.Item(staffLessNoColumn, rowNum).Value = staffSubj Then
                    lessonPerDay = lessonPerDay + 1
                End If
            End If
        Next
        Return lessonPerDay
    End Function
    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        dbconnection()
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If
        Try
            Dim result As MsgBoxResult = MsgBox("Save Record/s?", MsgBoxStyle.Question + MsgBoxStyle.YesNo + MsgBoxStyle.ApplicationModal, "Confirm Transaction")
            If result = MsgBoxResult.No Then
                Exit Sub
            End If

            rec = 0
            i = 0
            j = 0

            cmdGeneration.Connection = conn
            cmdGeneration.CommandType = CommandType.Text
            cmdGeneration.CommandText = "DELETE FROM tblTimeTable WHERE (acadYear=@classYear)"
            cmdGeneration.Parameters.Clear()
            cmdGeneration.Parameters.AddWithValue("@classYear", Me.cboYear.Text.Trim)
            cmdGeneration.ExecuteNonQuery()

            For i = 0 To Me.dgTTGeneration.ColumnCount - 1
                j = 0
                For j = 0 To Me.dgTTGeneration.RowCount - 1

                    cmdGeneration.Connection = conn
                    cmdGeneration.CommandType = CommandType.StoredProcedure
                    cmdGeneration.CommandText = "sprocTTGenSaveTT"
                    cmdGeneration.Parameters.Clear()
                    If Not (Me.dgTTGeneration.Item(i, j).Value = "") Then
                        If (Me.dgTTGeneration.Columns(i).HeaderText.Trim.Contains("LESSON")) Then
                            cmdGeneration.Parameters.AddWithValue("@employeeNo", Me.dgTTGeneration.Item(i, j).Value.ToString.Trim.Substring(4, 3))
                            cmdGeneration.Parameters.AddWithValue("@subjAbbr", Me.dgTTGeneration.Item(i, j).Value.ToString.Trim.Substring(0, 3))
                        End If
                    End If
                    cmdGeneration.Parameters.AddWithValue("@className", Me.dgTTGeneration.Item(1, j).Value)
                    cmdGeneration.Parameters.AddWithValue("@classYear", Me.dgTTGeneration.Item(3, j).Value)
                    cmdGeneration.Parameters.AddWithValue("@classStream", Me.dgTTGeneration.Item(2, j).Value)
                    cmdGeneration.Parameters.AddWithValue("@columnHeaderName", Me.dgTTGeneration.Columns(i).HeaderText.Trim)
                    cmdGeneration.Parameters.AddWithValue("@rowNum", j)
                    cmdGeneration.Parameters.AddWithValue("@columnNum", i)
                    cmdGeneration.Parameters.AddWithValue("@userName", userName.Trim)
                    cmdGeneration.Parameters.AddWithValue("@dayOfWeek", Me.dgTTGeneration.Item(0, j).Value)
                    cmdGeneration.Parameters.AddWithValue("@lessonId", Me.dgTTGeneration.Item(i, j).Tag)
                    cmdGeneration.ExecuteNonQuery()
                    rec = rec + 1
                    'End If
                    'End If
                Next
            Next
            If rec > 0 Then
                MsgBox("Record/s Saved.", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "System Reply")
            End If
            Me.dgTTGeneration.Rows.Clear()
            Me.dgTTGeneration.Columns.Clear()
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub

    Private Sub SETMAXIMUMLESSONSPERDAYToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SETMAXIMUMLESSONSPERDAYToolStripMenuItem.Click
        If Me.cboYear.Text.Trim.Length <= 0 Then
            MsgBox("Academic Year is missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If
        Dim UserInput As String
        UserInput = InputBox("Enter Lessons Per Day Per Class Per Staff", "TimeTable SetUp")
        If (UserInput.Trim.Length > 0) Then
            Try
                If conn.State = ConnectionState.Closed Then
                    conn.Open()
                End If
                cmdGeneration.Connection = conn
                cmdGeneration.CommandType = CommandType.Text
                cmdGeneration.CommandText = "UPDATE tblTTGenParameters SET value=@value WHERE " & _
                    vbNewLine & " (parameterName='SET MAXIMUM LESSONS PER DAY') AND (acadYear=@acadYear)"
                cmdGeneration.Parameters.Clear()
                cmdGeneration.Parameters.AddWithValue("@acadYear", Me.cboYear.Text.Trim)
                cmdGeneration.Parameters.AddWithValue("@value", UserInput.Trim.ToUpper)
                rec = rec + cmdGeneration.ExecuteNonQuery
                If rec > 0 Then
                    MsgBox("Record/s Saved.", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "System Reply")
                End If
            Catch ex As Exception
                MsgBox("Error Encountered", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly + MsgBoxStyle.ApplicationModal, "Application error")
            Finally
                If conn.State = ConnectionState.Open Then
                    conn.Close()
                End If
            End Try
        End If
    End Sub

    Private Sub SETMAXIMUMDOUBLESPERDAYToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SETMAXIMUMDOUBLESPERDAYToolStripMenuItem.Click
        If Me.cboYear.Text.Trim.Length <= 0 Then
            MsgBox("Academic Year is missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If
        Dim UserInput As String
        UserInput = InputBox("Enter Maximum Number Of Double Lessons Per Day per Class.", "TimeTable SetUp")
        If (UserInput.Trim.Length > 0) Then
            Try
                If conn.State = ConnectionState.Closed Then
                    conn.Open()
                End If
                cmdGeneration.Connection = conn
                cmdGeneration.CommandType = CommandType.Text
                cmdGeneration.CommandText = "UPDATE tblTTGenParameters SET value=@value WHERE " & _
                    vbNewLine & " (parameterName='SET MAXIMUM DOUBLES PER DAY') AND (acadYear=@acadYear)"
                cmdGeneration.Parameters.Clear()
                cmdGeneration.Parameters.AddWithValue("@acadYear", Me.cboYear.Text.Trim)
                cmdGeneration.Parameters.AddWithValue("@value", UserInput.Trim.ToUpper)
                rec = rec + cmdGeneration.ExecuteNonQuery
                If rec > 0 Then
                    MsgBox("Record/s Saved.", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "System Reply")
                End If
            Catch ex As Exception
                MsgBox("Error Encountered", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly + MsgBoxStyle.ApplicationModal, "Application error")
            Finally
                If conn.State = ConnectionState.Open Then
                    conn.Close()
                End If
            End Try
        End If
    End Sub

    Private Sub btnView_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnView.Click
        If Me.cboYear.Text.Trim.Length <= 0 Then
            MsgBox("Academic Year is missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            Me.dgTTGeneration.Rows.Clear()
            Me.dgTTGeneration.Columns.Clear()

            cmdGeneration.Connection = conn
            cmdGeneration.CommandText = "SELECT MAX(ISNULL(rowNum,0))+1 AS totalRows,MAX(ISNULL(columnNum,0))+1 AS totalColumns " & _
                vbNewLine & " FROM vwTTGenView WHERE (year=@year)"
            cmdGeneration.CommandType = CommandType.Text
            cmdGeneration.Parameters.Clear()
            cmdGeneration.Parameters.AddWithValue("@year", Me.cboYear.Text.Trim)
            reader = cmdGeneration.ExecuteReader
            While reader.Read
                totalColumns = IIf(DBNull.Value.Equals(reader!totalColumns), "", reader!totalColumns)
                totalRows = IIf(DBNull.Value.Equals(reader!totalRows), "", reader!totalRows)
            End While
            reader.Close()

            If totalColumns <= 1 Or totalRows <= 1 Then
                MsgBox("No timetable saved in the system.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
                Exit Sub
            End If
            Me.dgTTGeneration.ColumnCount = totalColumns
            Me.dgTTGeneration.RowCount = totalRows

            cmdGeneration.CommandText = "SELECT DISTINCT columnHeaderName,columnNum FROM vwTTGenView WHERE (year=@year) ORDER BY columnNum"
            cmdGeneration.CommandType = CommandType.Text
            cmdGeneration.Parameters.Clear()
            cmdGeneration.Parameters.AddWithValue("@year", Me.cboYear.Text.Trim)
            reader = cmdGeneration.ExecuteReader
            While reader.Read
                Me.dgTTGeneration.Columns(IIf(DBNull.Value.Equals(reader!columnNum), "", reader!columnNum)).HeaderText = _
                    IIf(DBNull.Value.Equals(reader!columnHeaderName), "", reader!columnHeaderName)
            End While
            reader.Close()

            cmdGeneration.CommandText = "SELECT * FROM vwTTGenView WHERE (year=@year)"
            cmdGeneration.Parameters.Clear()
            cmdGeneration.Parameters.AddWithValue("@year", Me.cboYear.Text.Trim)
            reader = cmdGeneration.ExecuteReader
            While reader.Read
                If IIf(DBNull.Value.Equals(reader!columnNum), "", reader!columnNum) = 0 Then
                    Me.dgTTGeneration.Columns(IIf(DBNull.Value.Equals(reader!columnNum), "", reader!columnNum)).Width = 100
                    Me.dgTTGeneration.Columns(IIf(DBNull.Value.Equals(reader!columnNum), "", reader!columnNum)).ReadOnly = True
                    Me.dgTTGeneration.Columns(IIf(DBNull.Value.Equals(reader!columnNum), "", reader!columnNum)).SortMode = DataGridViewColumnSortMode.NotSortable
                    Me.dgTTGeneration.Columns(IIf(DBNull.Value.Equals(reader!columnNum), "", reader!columnNum)).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                Else
                    Me.dgTTGeneration.Columns(IIf(DBNull.Value.Equals(reader!columnNum), "", reader!columnNum)).Width = 70
                    Me.dgTTGeneration.Columns(IIf(DBNull.Value.Equals(reader!columnNum), "", reader!columnNum)).ReadOnly = True
                    Me.dgTTGeneration.Columns(IIf(DBNull.Value.Equals(reader!columnNum), "", reader!columnNum)).SortMode = DataGridViewColumnSortMode.NotSortable
                    Me.dgTTGeneration.Columns(IIf(DBNull.Value.Equals(reader!columnNum), "", reader!columnNum)).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                End If

                Me.dgTTGeneration.Item(IIf(DBNull.Value.Equals(reader!columnNum), "", reader!columnNum), _
                                       IIf(DBNull.Value.Equals(reader!rowNum), "", reader!rowNum)).Value = _
                                   IIf(DBNull.Value.Equals(reader!cellValue), "", reader!cellValue)
                If Not (IIf(DBNull.Value.Equals(reader!cellBackColor), "", reader!cellBackColor) = "") Then
                    Me.dgTTGeneration.Item(IIf(DBNull.Value.Equals(reader!columnNum), "", reader!columnNum), _
                                       IIf(DBNull.Value.Equals(reader!rowNum), "", reader!rowNum)).Style.BackColor = _
                                   Color.FromName(IIf(DBNull.Value.Equals(reader!cellBackColor), "", reader!cellBackColor))
                End If
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

    Private Sub genTimer_Tick(ByVal sender As Object, ByVal e As System.EventArgs) Handles genTimer.Tick
        minNo = minNo + 1
    End Sub

    Private Sub bgWorkerGen_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles bgWorkerGen.DoWork
        Try
            Dim spaceNumber As New Dictionary(Of String, Integer)
            While spaceCount > 0 And lessonCount > 0
                pickLesson()
                selectFreeSpace(className.Trim, classStream.Trim, classYear, TryCast(e.Argument, DataGridView))
                i = selectedSpace(0)
                j = selectedSpace(1)

                If TryCast(e.Argument, DataGridView).Item(i, j).Value = "" Then
                    If TryCast(e.Argument, DataGridView).Item(0, j).Value = "MONDAY" Then
                        weekDayNo = 1
                    ElseIf TryCast(e.Argument, DataGridView).Item(0, j).Value = "TUESDAY" Then
                        weekDayNo = 2
                    ElseIf TryCast(e.Argument, DataGridView).Item(0, j).Value = "WEDNESDAY" Then
                        weekDayNo = 3
                    ElseIf TryCast(e.Argument, DataGridView).Item(0, j).Value = "THURSDAY" Then
                        weekDayNo = 4
                    ElseIf TryCast(e.Argument, DataGridView).Item(0, j).Value = "FRIDAY" Then
                        weekDayNo = 5
                    ElseIf TryCast(e.Argument, DataGridView).Item(0, j).Value = "SATURDAY" Then
                        weekDayNo = 6
                    ElseIf TryCast(e.Argument, DataGridView).Item(0, j).Value = "SUNDAY" Then
                        weekDayNo = 7
                    End If

                    Dim subjectAllowed As Boolean = checkSubjectConstraint(lessonAbbr.Trim, TryCast(e.Argument, DataGridView).Columns(i).HeaderText.Trim, weekDayNo)
                    Dim staffAllowed As Boolean = checkStaffConstraint(employeeNo.Trim, TryCast(e.Argument, DataGridView).Columns(i).HeaderText.Trim, weekDayNo)
                    Dim clashFound As Boolean = checkTeacherClashes(TryCast(e.Argument, DataGridView).Columns(i).HeaderText.Trim, employeeNo.Trim, TryCast(e.Argument, DataGridView).Item(0, j).Value, TryCast(e.Argument, DataGridView))
                    Dim lessonStaffCountFound As Integer = lessonNoPerDay(j, lessonAbbr.Trim & "-" & employeeNo.Trim, TryCast(e.Argument, DataGridView))
                    Dim maxDoubleLessPerDayFound As Integer = lessonNoPerDay(j, lessonAbbr.Trim & "-" & employeeNo.Trim, TryCast(e.Argument, DataGridView))

                    dbconnection()
                    If conn.State = ConnectionState.Closed Then
                        conn.Open()
                    End If

                    If minNo >= 0 And minNo <= 80 Then
                        If subjectAllowed = True And staffAllowed = True And clashFound = False And lessonStaffCountFound < lessonStaffCount Then
                            TryCast(e.Argument, DataGridView).Item(i, j).Value = lessonAbbr.Trim & "-" & employeeNo.Trim
                            TryCast(e.Argument, DataGridView).Item(i, j).Tag = lessonId
                            If Not (subjColor = Nothing) Then
                                TryCast(e.Argument, DataGridView).Item(i, j).Style.BackColor = Color.FromName(subjColor.Trim)
                            End If

                            cmdGeneration.Connection = conn
                            cmdGeneration.CommandType = CommandType.Text
                            cmdGeneration.CommandText = "UPDATE tblTTLessonSetUp SET lessonsInserted=lessonsInserted+1 WHERE (lessonId=@lessonId)"
                            cmdGeneration.Parameters.Clear()
                            cmdGeneration.Parameters.AddWithValue("@lessonId", lessonId)
                            cmdGeneration.ExecuteNonQuery()
                        End If

                    ElseIf minNo > 80 And minNo <= 85 Then
                        If clashFound = False Then
                            TryCast(e.Argument, DataGridView).Item(i, j).Value = lessonAbbr.Trim & "-" & employeeNo.Trim
                            TryCast(e.Argument, DataGridView).Item(i, j).Tag = lessonId
                            If Not (subjColor = Nothing) Then
                                TryCast(e.Argument, DataGridView).Item(i, j).Style.BackColor = Color.FromName(subjColor.Trim)
                            End If

                            cmdGeneration.Connection = conn
                            cmdGeneration.CommandType = CommandType.Text
                            cmdGeneration.CommandText = "UPDATE tblTTLessonSetUp SET lessonsInserted=lessonsInserted+1 WHERE (lessonId=@lessonId)"
                            cmdGeneration.Parameters.Clear()
                            cmdGeneration.Parameters.AddWithValue("@lessonId", lessonId)
                            cmdGeneration.ExecuteNonQuery()
                        End If
                    End If
                End If
                    spaceCount = countFreeSpaces(TryCast(e.Argument, DataGridView))
                    lessonCount = countFreeLessons()

                    spaceNumber.Clear()
                    spaceNumber.Add("lessonCount", lessonCount)
                    spaceNumber.Add("spaceCount", spaceCount)
                    bgWorkerGen.ReportProgress(minNo, spaceNumber)
                    e.Result = TryCast(e.Argument, DataGridView)
            End While
        Catch ex As Exception

        End Try
    End Sub

    Private Sub bgWorkerGen_ProgressChanged(ByVal sender As Object, ByVal e As System.ComponentModel.ProgressChangedEventArgs) Handles bgWorkerGen.ProgressChanged
        Try
            If Not (IsNothing(e.UserState)) Then

            End If
            ToolStripStatusLabel2.Text = "TIME ELAPSED IN SECONDS : " & minNo.ToString("000") & Space(1) & _
                String.Format("SPACES AVAILABLE : {0} LESSON TO INSERT : {1}", _
                              TryCast(e.UserState, Dictionary(Of String, Integer)).Item("spaceCount").ToString("000"), _
                              TryCast(e.UserState, Dictionary(Of String, Integer)).Item("lessonCount").ToString("000"))
        Catch ex As Exception

        End Try
    End Sub

    Private Sub bgWorkerGen_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles bgWorkerGen.RunWorkerCompleted
        Try
            Me.dgTTGeneration.Columns.Clear()
            Me.dgTTGeneration.Rows.Clear()
            Me.dgTTGeneration.DataSource = TryCast(e.Result, DataGridView).DataSource
        Catch ex As Exception

        End Try
    End Sub
End Class