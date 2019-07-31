Imports System.Data.SqlClient
Public Class frmAcadStudentSubject
    Dim numberOfSubjects As Integer = 0
    Dim maxNoOfSubjects As Integer = 0
    Dim recordExists As Boolean = True
    Dim rec As Integer = 0
    Dim queryType As String = Nothing
    Dim cmdStudSub As New SqlCommand
    Dim reader As SqlDataReader
    Private Sub frmAcadStudentSubject_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            loadCombos()
            loadList()
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub
    Private Sub frmAcadStudentSubject_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint
        If Me.IsMdiChild Then
            Dim PnlLoc As New Point
            PnlLoc.X = CInt((Me.Width - Me.pnlStudSubjects.Width) / 2)
            PnlLoc.Y = CInt((Me.Height - Me.pnlStudSubjects.Height) / 2.5)
            Me.pnlStudSubjects.Location = PnlLoc
        Else
            Me.pnlStudSubjects.Dock = DockStyle.Fill
        End If
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub
    Private Sub loadCombos()
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            Me.cboYearUpdate.Text = ""
            Me.cboYearUpdate.SelectedIndex = -1
            Me.cboYearUpdate.Items.Clear()

            Me.cboYearView.Text = ""
            Me.cboYearView.SelectedIndex = -1
            Me.cboYearView.Items.Clear()

            Me.cboStreamUpdate.Text = ""
            Me.cboStreamUpdate.SelectedIndex = -1
            Me.cboStreamUpdate.Items.Clear()

            Me.cboStreaView.Text = ""
            Me.cboStreaView.SelectedIndex = -1
            Me.cboStreaView.Items.Clear()

            Me.cboClassUpdate.Text = ""
            Me.cboClassUpdate.SelectedIndex = -1
            Me.cboClassUpdate.Items.Clear()

            Me.cboClassView.Text = ""
            Me.cboClassView.SelectedIndex = -1
            Me.cboClassView.Items.Clear()

            Me.cmdStudSub.CommandType = CommandType.Text
            Me.cmdStudSub.Connection = conn
            Me.cmdStudSub.CommandText = "SELECT DISTINCT year FROM tblClasses WHERE (status=1) ORDER BY year"
            Me.cmdStudSub.Parameters.Clear()
            reader = Me.cmdStudSub.ExecuteReader
            If reader.HasRows = True Then
                While reader.Read
                    Me.cboYearUpdate.Items.Add(IIf(DBNull.Value.Equals(reader!year), "", reader!year))
                    Me.cboYearView.Items.Add(IIf(DBNull.Value.Equals(reader!year), "", reader!year))
                End While
            End If
            reader.Close()

            Me.cmdStudSub.CommandText = "SELECT DISTINCT stream FROM tblClasses WHERE (status=1) ORDER BY stream"
            Me.cmdStudSub.Parameters.Clear()
            reader = Me.cmdStudSub.ExecuteReader
            If reader.HasRows = True Then
                While reader.Read
                    Me.cboStreamUpdate.Items.Add(IIf(DBNull.Value.Equals(reader!stream), "", reader!stream))
                    Me.cboStreaView.Items.Add(IIf(DBNull.Value.Equals(reader!stream), "", reader!stream))
                End While
            End If
            reader.Close()

            Me.cmdStudSub.CommandText = "SELECT DISTINCT className FROM tblClasses WHERE (status=1) ORDER BY className"
            Me.cmdStudSub.Parameters.Clear()
            reader = Me.cmdStudSub.ExecuteReader
            If reader.HasRows = True Then
                While reader.Read
                    Me.cboClassUpdate.Items.Add(IIf(DBNull.Value.Equals(reader!className), "", reader!className))
                    Me.cboClassView.Items.Add(IIf(DBNull.Value.Equals(reader!className), "", reader!className))
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
    Private Sub loadList()
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            Me.lstSubjectName.Items.Clear()

            Me.cmdStudSub.CommandType = CommandType.Text
            Me.cmdStudSub.Connection = conn
            Me.cmdStudSub.CommandText = "SELECT * FROM tblSubjects WHERE (subStatus=1) ORDER BY subCode"
            Me.cmdStudSub.Parameters.Clear()
            reader = Me.cmdStudSub.ExecuteReader
            If reader.HasRows = True Then
                While reader.Read
                    li = Me.lstSubjectName.Items.Add(IIf(DBNull.Value.Equals(reader!subName), "", reader!subName))
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
    Private Sub cboYearUpdate_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboYearUpdate.SelectedIndexChanged, cboStreamUpdate.SelectedIndexChanged, cboClassUpdate.SelectedIndexChanged
        If Me.cboClassUpdate.Text.Trim.Length <= 0 And Me.cboStreamUpdate.Text.Trim.Length <= 0 And Me.cboYearUpdate.Text.Trim.Length <= 0 Then
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            Me.lstStudent.Items.Clear()
            Me.cmdStudSub.CommandType = CommandType.Text
            Me.cmdStudSub.Connection = conn
            Me.cmdStudSub.CommandText = "SELECT * FROM vwStudClass WHERE (studStatus=1) AND (classStatus=1) AND (classStudStatus=1) AND " &
                vbNewLine & " (className=@className) AND (stream=@stream) AND (year=@year)"
            Me.cmdStudSub.Parameters.Clear()
            Me.cmdStudSub.Parameters.AddWithValue("@className", Me.cboClassUpdate.Text.Trim)
            Me.cmdStudSub.Parameters.AddWithValue("@stream", Me.cboStreamUpdate.Text.Trim)
            Me.cmdStudSub.Parameters.AddWithValue("@year", Me.cboYearUpdate.Text.Trim)
            reader = Me.cmdStudSub.ExecuteReader
            If reader.HasRows = True Then
                While reader.Read
                    li = Me.lstStudent.Items.Add(IIf(DBNull.Value.Equals(reader!admNo), "", reader!admNo))
                    li.SubItems.Add(IIf(DBNull.Value.Equals(reader!FullName), "", reader!FullName))
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

    Private Sub cboYearView_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboYearView.SelectedIndexChanged, cboClassView.SelectedIndexChanged, cboStreaView.SelectedIndexChanged
        Me.lstStudentSubjectView.Items.Clear()
        Me.txtTotalSubjStud.Text = ""
        If Me.cboClassView.Text.Trim.Length <= 0 And Me.cboStreaView.Text.Trim.Length <= 0 And Me.cboYearView.Text.Trim.Length <= 0 Then
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()

            Me.cboAdmNo.Text = ""
            Me.cboAdmNo.SelectedIndex = -1
            Me.cboAdmNo.Items.Clear()

            Me.cmdStudSub.CommandType = CommandType.Text
            Me.cmdStudSub.Connection = conn
            Me.cmdStudSub.CommandText = "SELECT * FROM vwStudClass WHERE (studStatus=1) AND (classStatus=1) AND (classStudStatus=1) AND " &
                vbNewLine & " (className=@className) AND (stream=@stream) AND (year=@year)"
            Me.cmdStudSub.Parameters.Clear()
            Me.cmdStudSub.Parameters.AddWithValue("@className", Me.cboClassView.Text.Trim)
            Me.cmdStudSub.Parameters.AddWithValue("@stream", Me.cboStreaView.Text.Trim)
            Me.cmdStudSub.Parameters.AddWithValue("@year", Me.cboYearView.Text.Trim)
            reader = Me.cmdStudSub.ExecuteReader
            If reader.HasRows = True Then
                While reader.Read
                    Me.cboAdmNo.Items.Add(IIf(DBNull.Value.Equals(reader!admNo), "", reader!admNo))
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

    Private Sub btnView_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnView.Click
        If Me.cboAdmNo.Text.Trim.Length <= 0 Then
            MsgBox("Admission Number is missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            Me.lstStudentSubjectView.Items.Clear()
            Me.txtTotalSubjStud.Text = ""
            Me.cmdStudSub.CommandType = CommandType.Text
            Me.cmdStudSub.Connection = conn
            Me.cmdStudSub.CommandText = "SELECT * FROM vwAcadStudentSubjects WHERE (className=@className)  AND  " &
                vbNewLine & "(stream=@stream) AND (year=@year) AND (admNo=@admNo)"
            Me.cmdStudSub.Parameters.Clear()
            Me.cmdStudSub.Parameters.AddWithValue("@className", Me.cboClassView.Text.Trim)
            Me.cmdStudSub.Parameters.AddWithValue("@stream", Me.cboStreaView.Text.Trim)
            Me.cmdStudSub.Parameters.AddWithValue("@year", Me.cboYearView.Text.Trim)
            Me.cmdStudSub.Parameters.AddWithValue("@admNo", Me.cboAdmNo.Text.Trim)
            reader = Me.cmdStudSub.ExecuteReader

            If reader.HasRows = True Then
                While reader.Read
                    li = Me.lstStudentSubjectView.Items.Add(IIf(DBNull.Value.Equals(reader!FullName), "", reader!FullName))
                    li.SubItems.Add(IIf(DBNull.Value.Equals(reader!subName), "", reader!subName))
                    li.SubItems.Add(IIf(DBNull.Value.Equals(reader!className), "", reader!className))
                    li.SubItems.Add(IIf(DBNull.Value.Equals(reader!stream), "", reader!stream))
                    li.SubItems.Add(IIf(DBNull.Value.Equals(reader!Year), "", reader!Year))
                End While
            End If
            reader.Close()
            Me.txtTotalSubjStud.Text = Me.lstStudentSubjectView.Items.Count
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub

    Private Sub cboAdmNo_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboAdmNo.SelectedIndexChanged
        If Me.cboAdmNo.Text.Trim.Length <= 0 Then
            Exit Sub
        End If
        Me.lstStudentSubjectView.Items.Clear()
        Me.txtTotalSubjStud.Text = ""
    End Sub

    Private Sub lstStudent_ColumnClick(ByVal sender As Object, ByVal e As System.Windows.Forms.ColumnClickEventArgs) Handles lstStudent.ColumnClick
        If (e.Column() = 0) And (Me.lstStudent.Items.Count > 0) Then
            For Each Li As ListViewItem In Me.lstStudent.Items
                Li.Checked = Not (Li.Checked)
            Next
        End If
    End Sub

    Private Sub txtQuickSearchStud_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtQuickSearchStud.TextChanged
        Me.txtTotalSubjStud.Text = ""
        If Me.txtQuickSearchStud.Text.Trim.Length <= 3 Then
            Me.lstStudentSubjectView.Items.Clear()
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            Me.lstStudentSubjectView.Items.Clear()

            Me.cmdStudSub.CommandType = CommandType.Text
            Me.cmdStudSub.Connection = conn
            Me.cmdStudSub.CommandText = "SELECT * FROM vwAcadStudentSubjects WHERE (FullName LIKE @FullName) " &
                vbNewLine & "ORDER BY FullName,year,stream,className,subName"
            Me.cmdStudSub.Parameters.Clear()
            Me.cmdStudSub.Parameters.AddWithValue("@FullName", String.Format("%{0}%", TryCast(Me.txtQuickSearchStud.Text.Trim, String).Trim))
            reader = Me.cmdStudSub.ExecuteReader

            If reader.HasRows = True Then
                While reader.Read
                    li = Me.lstStudentSubjectView.Items.Add(IIf(DBNull.Value.Equals(reader!FullName), "", reader!FullName))
                    li.SubItems.Add(IIf(DBNull.Value.Equals(reader!subName), "", reader!subName))
                    li.SubItems.Add(IIf(DBNull.Value.Equals(reader!className), "", reader!className))
                    li.SubItems.Add(IIf(DBNull.Value.Equals(reader!stream), "", reader!stream))
                    li.SubItems.Add(IIf(DBNull.Value.Equals(reader!Year), "", reader!Year))
                End While
            End If
            reader.Close()
            Me.txtTotalSubjStud.Text = Me.lstStudentSubjectView.Items.Count
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub

    Private Sub btnClear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClear.Click
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            clearTexts()
            loadCombos()
            loadList()
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub
    Private Sub clearTexts()
        Me.cboYearUpdate.Text = ""
        Me.cboYearUpdate.SelectedIndex = -1
        Me.cboYearUpdate.Items.Clear()

        Me.cboYearView.Text = ""
        Me.cboYearView.SelectedIndex = -1
        Me.cboYearView.Items.Clear()

        Me.cboStreamUpdate.Text = ""
        Me.cboStreamUpdate.SelectedIndex = -1
        Me.cboStreamUpdate.Items.Clear()

        Me.cboStreaView.Text = ""
        Me.cboStreaView.SelectedIndex = -1
        Me.cboStreaView.Items.Clear()

        Me.cboClassUpdate.Text = ""
        Me.cboClassUpdate.SelectedIndex = -1
        Me.cboClassUpdate.Items.Clear()

        Me.cboClassView.Text = ""
        Me.cboClassView.SelectedIndex = -1
        Me.cboClassView.Items.Clear()

        Me.cboAdmNo.Text = ""
        Me.cboAdmNo.SelectedIndex = -1
        Me.cboAdmNo.Items.Clear()

        Me.lstStudent.Items.Clear()
        Me.lstStudentSubjectView.Items.Clear()
        Me.lstSubjectName.Items.Clear()

        Me.txtQuickSearchStud.Text = ""
        Me.txtTotalSubjStud.Text = ""

    End Sub
    Private Function GetStudentSubjectNo(ByVal admNo As String, ByVal className As String, ByVal yearOf As Integer)
        Me.cmdStudSub.Connection = conn
        Me.cmdStudSub.CommandText = "SELECT COUNT(admNo) AS subNo FROM  vwAcadStudentSubjects WHERE (admNo=@admNo)  AND " &
            vbNewLine & " (year=@year)  AND (className=@className)"
        Me.cmdStudSub.CommandType = CommandType.Text
        Me.cmdStudSub.Parameters.Clear()
        Me.cmdStudSub.Parameters.AddWithValue("@admNo", admNo.Trim)
        Me.cmdStudSub.Parameters.AddWithValue("@className", className.Trim)
        Me.cmdStudSub.Parameters.AddWithValue("@year", yearOf)
        reader = Me.cmdStudSub.ExecuteReader
        If reader.HasRows = True Then
            While reader.Read
                numberOfSubjects = IIf(DBNull.Value.Equals(reader!subNo), "", reader!subNo)
            End While
        ElseIf reader.HasRows = False Then
            numberOfSubjects = 0
        End If
        reader.Close()
    End Function
    Private Sub getMaxNoOfSubjects()
        Me.cmdStudSub.Connection = conn
        Me.cmdStudSub.CommandType = CommandType.Text
        Me.cmdStudSub.CommandText = "SELECT maxNo FROM tblMaxNoOfSubjects WHERE (parameter='Student') AND (class=@class)"
        Me.cmdStudSub.Parameters.Clear()
        Me.cmdStudSub.Parameters.AddWithValue("@class", Me.cboClassUpdate.Text.Trim)
        reader = Me.cmdStudSub.ExecuteReader
        If reader.HasRows Then
            While reader.Read
                maxNoOfSubjects = IIf(DBNull.Value.Equals(reader!maxNo), "", (reader!maxNo))
            End While
        Else
            maxNoOfSubjects = 0
        End If
        reader.Close()
    End Sub
    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        Dim diffence As Integer = 0
        If Me.lstStudent.Items.Count <= 0 Then
            MsgBox("No student in the list to save.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.lstSubjectName.Items.Count <= 0 Then
            MsgBox("No subject in the list to save.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.lstStudent.CheckedItems.Count <= 0 Then
            MsgBox("No student checked in the list to save.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.lstSubjectName.CheckedItems.Count <= 0 Then
            MsgBox("No subject checked in the list to save.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboClassUpdate.Text.Trim.Length <= 0 Then
            MsgBox("Missing Class Name.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboStreamUpdate.Text.Trim.Length <= 0 Then
            MsgBox("Missing stream details.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboYearUpdate.Text.Trim.Length <= 0 Then
            MsgBox("Missing year Details.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If

        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()

            i = 0
            j = 0
            maxrec = Me.lstStudent.CheckedItems.Count - 1
            maxrec1 = Me.lstSubjectName.CheckedItems.Count - 1

            numberOfSubjects = 0
            GetStudentSubjectNo(Me.lstStudent.CheckedItems(i).Text.Trim, Me.cboClassUpdate.Text.Trim, Me.cboYearUpdate.Text.Trim)
            numberOfSubjects = numberOfSubjects + Me.lstSubjectName.CheckedItems.Count
            maxNoOfSubjects = 0
            getMaxNoOfSubjects()
            diffence = maxNoOfSubjects - numberOfSubjects
            If diffence < 0 Then
                MsgBox("If subjects added for student " & Me.lstStudent.CheckedItems(i).SubItems(1).Text.Trim & " will exceed maximum number of subjects per student." &
                       vbNewLine & "The student will have " & numberOfSubjects & " Subjects" &
                       vbNewLine & "Which exceed " & maxNoOfSubjects & " the maximum for class " &
                       Me.cboClassUpdate.Text.Trim & ".", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
                Exit Sub
            End If

            For i = 0 To maxrec
                maxrec1 = 0
                j = 0
                For j = 0 To Me.lstSubjectName.CheckedItems.Count - 1
                    recordExists = True
                    checkIfRecordsExist()
                    If recordExists = True Then
                        MsgBox("Entry For Student " & Me.lstStudent.CheckedItems(i).SubItems(1).Text.Trim & " ADM NO " &
                              Me.lstStudent.CheckedItems(i).Text.Trim & " SUBJECT " & Me.lstSubjectName.CheckedItems(j).Text.Trim & " Not Saved." &
                              vbNewLine & "The Record Already Exists in Database" & vbNewLine & "Click Ok To Continue Saving Record/s",
                              MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "System Reply")
                        Exit Sub
                    End If
                Next
            Next

            Dim result As MsgBoxResult = MsgBox("Save Record?", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.YesNo, "Confirm Transaction")
            If result = MsgBoxResult.No Then
                Exit Sub
            End If

            rec = 0
            i = 0
            j = 0
            maxrec = Me.lstStudent.CheckedItems.Count - 1
            maxrec1 = Me.lstSubjectName.CheckedItems.Count - 1
            For i = 0 To maxrec
                maxrec1 = 0
                j = 0
                For j = 0 To Me.lstSubjectName.CheckedItems.Count - 1
                    queryType = "INSERT"
                    Me.cmdStudSub.Connection = conn
                    Me.cmdStudSub.CommandType = CommandType.StoredProcedure
                    Me.cmdStudSub.CommandText = "sprocStudentSubject"
                    Me.cmdStudSub.Parameters.Clear()
                    Me.cmdStudSub.Parameters.AddWithValue("@subName", Me.lstSubjectName.CheckedItems(j).Text.Trim)
                    Me.cmdStudSub.Parameters.AddWithValue("@admNo", Me.lstStudent.CheckedItems(i).Text.Trim)
                    Me.cmdStudSub.Parameters.AddWithValue("@className", Me.cboClassUpdate.Text.Trim)
                    Me.cmdStudSub.Parameters.AddWithValue("@stream", Me.cboStreamUpdate.Text.Trim)
                    Me.cmdStudSub.Parameters.AddWithValue("@year", Me.cboYearUpdate.Text.Trim)
                    Me.cmdStudSub.Parameters.AddWithValue("@regBy", userName.Trim)
                    Me.cmdStudSub.Parameters.AddWithValue("@dateOfReg", Date.Now)
                    Me.cmdStudSub.Parameters.AddWithValue("@queryType", Me.queryType.Trim)
                    rec = Me.cmdStudSub.ExecuteNonQuery
                Next
            Next
            If rec > 0 Then
                MsgBox("Record/s Saved!", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "SuccessFull Transaction")
            End If
            clearTexts()
            loadCombos()
            loadList()
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            rec = 0
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub
    Private Sub checkIfRecordsExist()
        recordExists = True
        Me.cmdStudSub.Connection = conn
        Me.cmdStudSub.CommandText = "SELECT * FROM vwAcadStudentSubjects WHERE  (admNo=@admNo) AND  " &
            vbNewLine & " (className=@className)  AND (stream=@stream) AND (year=@year) AND (subName=@subName)"
        Me.cmdStudSub.CommandType = CommandType.Text
        Me.cmdStudSub.Parameters.Clear()
        Me.cmdStudSub.Parameters.AddWithValue("@subName", Me.lstSubjectName.CheckedItems(j).Text.Trim)
        Me.cmdStudSub.Parameters.AddWithValue("@admNo", Me.lstStudent.CheckedItems(i).Text.Trim)
        Me.cmdStudSub.Parameters.AddWithValue("@className", Me.cboClassUpdate.Text.Trim)
        Me.cmdStudSub.Parameters.AddWithValue("@stream", Me.cboStreamUpdate.Text.Trim)
        Me.cmdStudSub.Parameters.AddWithValue("@year", Me.cboYearUpdate.Text.Trim)
        reader = Me.cmdStudSub.ExecuteReader
        If reader.HasRows = True Then
            recordExists = True
        ElseIf reader.HasRows = False Then
            recordExists = False
        End If
        reader.Close()
    End Sub

    Private Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        If Me.lstStudentSubjectView.Items.Count <= 0 Then
            MsgBox("No items to delete.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.lstStudentSubjectView.CheckedItems.Count <= 0 Then
            MsgBox("Check the items to delete.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            Dim result As MsgBoxResult = MsgBox("Delete Record/s?", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.YesNo, "Confirm Transaction")
            If result = MsgBoxResult.No Then
                Exit Sub
            End If
            rec = 0
            i = 0
            For i = 0 To Me.lstStudentSubjectView.CheckedItems.Count - 1
                Me.queryType = "DELETE"
                Me.cmdStudSub.Connection = conn
                Me.cmdStudSub.CommandType = CommandType.StoredProcedure
                Me.cmdStudSub.CommandText = "sprocStudentSubject"
                Me.cmdStudSub.Parameters.Clear()
                Me.cmdStudSub.Parameters.AddWithValue("@subName", Me.lstStudentSubjectView.CheckedItems(i).SubItems(1).Text)
                Me.cmdStudSub.Parameters.AddWithValue("@admNo", Me.cboAdmNo.Text.Trim)
                Me.cmdStudSub.Parameters.AddWithValue("@className", Me.lstStudentSubjectView.CheckedItems(i).SubItems(2).Text)
                Me.cmdStudSub.Parameters.AddWithValue("@stream", Me.lstStudentSubjectView.CheckedItems(i).SubItems(3).Text)
                Me.cmdStudSub.Parameters.AddWithValue("@year", Me.lstStudentSubjectView.CheckedItems(i).SubItems(4).Text)
                Me.cmdStudSub.Parameters.AddWithValue("@regBy", userName.Trim)
                Me.cmdStudSub.Parameters.AddWithValue("@dateOfReg", Date.Now)
                Me.cmdStudSub.Parameters.AddWithValue("@queryType", Me.queryType.Trim)
                rec = rec + Me.cmdStudSub.ExecuteNonQuery
            Next
            If rec > 0 Then
                MsgBox("Record/s Deleted", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "SuccessFull Transactions")
            End If
            clearTexts()
            loadCombos()
            loadList()
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub

    Private Sub lstStudent_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lstStudent.SelectedIndexChanged

    End Sub
End Class