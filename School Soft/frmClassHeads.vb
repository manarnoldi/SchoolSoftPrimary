Imports System.Data.SqlClient
Public Class frmClassHeads
    Dim queryType As String = Nothing
    Dim entryOk As Boolean = False
    Dim recordExists As Boolean = True
    Dim cmdClassHeads As New SqlCommand
    Dim reader As SqlDataReader
    Dim rec As Integer = 0
    Private Sub frmClassHeads_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            loadCombos()
            loadDetails()
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub
    Private Sub frmClassHeads_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint
        If Me.IsMdiChild Then
            Dim PnlLoc As New Point
            PnlLoc.X = CInt((Me.Width - Me.pnlClassHeads.Width) / 2)
            PnlLoc.Y = CInt((Me.Height - Me.pnlClassHeads.Height) / 2.5)
            Me.pnlClassHeads.Location = PnlLoc
        Else
            Me.pnlClassHeads.Dock = DockStyle.Fill
        End If
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub
    Private Sub loadCombos()
        cboTeacherNo.Items.Clear()
        cboClassName.Items.Clear()
        cboClassStream.Items.Clear()
        cboClassYear.Items.Clear()

        cboTeacherNo.Text = ""
        cboClassName.Text = ""
        cboClassStream.Text = ""
        cboClassYear.Text = ""

        cmdClassHeads.Connection = conn
        cmdClassHeads.CommandType = CommandType.Text
        cmdClassHeads.CommandText = "SELECT DISTINCT className FROM tblClasses WHERE (status='True') ORDER BY className"
        cmdClassHeads.Parameters.Clear()
        reader = cmdClassHeads.ExecuteReader
        If reader.HasRows Then
            While reader.Read
                cboClassName.Items.Add(IIf(DBNull.Value.Equals(reader!className), "", reader!className))
            End While
        End If
        reader.Close()

        cmdClassHeads.CommandText = "SELECT DISTINCT stream FROM tblClasses WHERE (status='True') ORDER BY stream"
        cmdClassHeads.Parameters.Clear()
        reader = cmdClassHeads.ExecuteReader
        If reader.HasRows Then
            While reader.Read
                cboClassStream.Items.Add(IIf(DBNull.Value.Equals(reader!stream), "", reader!stream))
            End While
        End If
        reader.Close()

        cmdClassHeads.CommandText = "SELECT DISTINCT year FROM tblClasses WHERE (status='True') ORDER BY year"
        cmdClassHeads.Parameters.Clear()
        reader = cmdClassHeads.ExecuteReader
        If reader.HasRows Then
            While reader.Read
                cboClassYear.Items.Add(IIf(DBNull.Value.Equals(reader!year), "", reader!year))
            End While
        End If
        reader.Close()

        cmdClassHeads.CommandText = "SELECT DISTINCT empNo FROM tblSchoolStaff WHERE (status='True') AND (empType='Teaching') ORDER BY empNo"
        cmdClassHeads.Parameters.Clear()
        reader = cmdClassHeads.ExecuteReader
        If reader.HasRows Then
            While reader.Read
                cboTeacherNo.Items.Add(IIf(DBNull.Value.Equals(reader!empNo), "", reader!empNo))
            End While
        End If
        reader.Close()

    End Sub

    Private Sub cboClassName_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboClassName.SelectedIndexChanged, cboClassStream.SelectedIndexChanged, cboClassYear.SelectedIndexChanged
        Me.cboStudNo.Items.Clear()
        Me.cboStudNo.Text = ""
        Me.txtStudName.Text = ""
        Me.cboTeacherNo.Text = ""
        Me.txtTeacherName.Text = ""
        If Me.cboClassName.Text.Trim.Length <= 0 Or Me.cboClassStream.Text.Trim.Length <= 0 Or Me.cboClassYear.Text.Trim.Length <= 0 Then
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            cmdClassHeads.Connection = conn
            cmdClassHeads.CommandType = CommandType.Text
            cmdClassHeads.CommandText = "SELECT admNo FROM vwStudClass WHERE (className=@className) AND (stream=@stream) " & _
                vbNewLine & " AND (year=@year) AND (classStatus='True') AND (classStudStatus='True') AND (studStatus='True')  ORDER BY admNo"
            cmdClassHeads.Parameters.Clear()
            cmdClassHeads.Parameters.AddWithValue("@className", Me.cboClassName.Text)
            cmdClassHeads.Parameters.AddWithValue("@stream", Me.cboClassStream.Text)
            cmdClassHeads.Parameters.AddWithValue("@year", Me.cboClassYear.Text)
            reader = cmdClassHeads.ExecuteReader
            If reader.HasRows Then
                While reader.Read
                    cboStudNo.Items.Add(IIf(DBNull.Value.Equals(reader!admNo), "", reader!admNo))
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

    Private Sub cboTeacherNo_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboTeacherNo.SelectedIndexChanged
        If Me.cboTeacherNo.Text.Trim.Length <= 0 Then
            Exit Sub
        End If
        Me.txtTeacherName.Text = ""
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            cmdClassHeads.Connection = conn
            cmdClassHeads.CommandType = CommandType.Text
            cmdClassHeads.CommandText = "SELECT FullName FROM tblSchoolStaff WHERE (empNo=@empNo) AND (status='True') ORDER BY FullName"
            cmdClassHeads.Parameters.Clear()
            cmdClassHeads.Parameters.AddWithValue("@empNo", Me.cboTeacherNo.Text)
            reader = cmdClassHeads.ExecuteReader
            If reader.HasRows Then
                While reader.Read
                    Me.txtTeacherName.Text = (IIf(DBNull.Value.Equals(reader!FullName), "", reader!FullName))
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

    Private Sub cboStudNo_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboStudNo.SelectedIndexChanged
        If Me.cboTeacherNo.Text.Trim.Length <= 0 Then
            Exit Sub
        End If
        Me.txtStudName.Text = ""
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            cmdClassHeads.Connection = conn
            cmdClassHeads.CommandType = CommandType.Text
            cmdClassHeads.CommandText = "SELECT FullName FROM tblStudDetails WHERE (admNo=@admNo) AND (status='True') ORDER BY FullName"
            cmdClassHeads.Parameters.Clear()
            cmdClassHeads.Parameters.AddWithValue("@admNo", Me.cboStudNo.Text)
            reader = cmdClassHeads.ExecuteReader
            If reader.HasRows Then
                While reader.Read
                    Me.txtStudName.Text = (IIf(DBNull.Value.Equals(reader!FullName), "", reader!FullName))
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
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            loadDetails()
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
            Me.lstClassHeads.Enabled = True
            Me.cboClassName.Enabled = True
            Me.cboClassStream.Enabled = True
            Me.cboClassYear.Enabled = True
            Me.btnDelete.Enabled = False
            Me.btnUpdate.Enabled = False
            Me.btnSave.Enabled = True
            subClearTexts()
        End Try
    End Sub
    Private Sub loadDetails()
        Me.lstClassHeads.Items.Clear()
        cmdClassHeads.Connection = conn
        cmdClassHeads.CommandType = CommandType.Text
        cmdClassHeads.CommandText = "SELECT * FROM vwClassHeads WHERE (classHeadStatus='True') AND (schoolStaffStatus='True') AND " & _
            vbNewLine & " (studentStatus='True') AND (classStatus='True') ORDER BY year,className,Stream"
        cmdClassHeads.Parameters.Clear()
        reader = cmdClassHeads.ExecuteReader
        If reader.HasRows Then
            While reader.Read
                li = Me.lstClassHeads.Items.Add(IIf(DBNull.Value.Equals(reader!className), "", reader!className))
                li.SubItems.Add(IIf(DBNull.Value.Equals(reader!stream), "", reader!stream))
                li.SubItems.Add(IIf(DBNull.Value.Equals(reader!year), "", reader!year))
                li.SubItems.Add(IIf(DBNull.Value.Equals(reader!admNo), "", reader!admNo))
                li.SubItems.Add(IIf(DBNull.Value.Equals(reader!fullName), "", reader!fullName))
                li.SubItems.Add(IIf(DBNull.Value.Equals(reader!empNo), "", reader!empNo))
                li.SubItems.Add(IIf(DBNull.Value.Equals(reader!classTeacher), "", reader!classTeacher))
                li.Tag = IIf(DBNull.Value.Equals(reader!classHeadId), "", reader!classHeadId)
            End While
        End If
        reader.Close()
    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        entryOk = False
        checkEntry()
        If entryOk = False Then
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            GetClassID()
            If Not (Me.cboClassName.Tag = Nothing) Then
                checkIfExists(Me.cboClassName.Tag)
            End If
            If recordExists = True Then
                MsgBox("Record Data For the Class Saved Already.Try Update", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
                Exit Sub
            End If
            Dim result As MsgBoxResult = MsgBox("Save Record?", MsgBoxStyle.Question + MsgBoxStyle.ApplicationModal + MsgBoxStyle.YesNo, "Confirm Transaction")
            If result = MsgBoxResult.Yes Then
                queryType = "INSERT"
                cmdClassHeads.Connection = conn
                cmdClassHeads.CommandType = CommandType.StoredProcedure
                cmdClassHeads.CommandText = "sprocClassHeads"
                cmdClassHeads.Parameters.Clear()
                cmdClassHeads.Parameters.AddWithValue("@className", Me.cboClassName.Text.Trim)
                cmdClassHeads.Parameters.AddWithValue("@classYear", Me.cboClassYear.Text.Trim)
                cmdClassHeads.Parameters.AddWithValue("@classStream", Me.cboClassStream.Text.Trim)
                cmdClassHeads.Parameters.AddWithValue("@admNo", Me.cboStudNo.Text.Trim)
                cmdClassHeads.Parameters.AddWithValue("@queryType", queryType.Trim)
                cmdClassHeads.Parameters.AddWithValue("@regBy", userName.Trim)
                cmdClassHeads.Parameters.AddWithValue("@empNo", Me.cboTeacherNo.Text.Trim)
                cmdClassHeads.Parameters.AddWithValue("@dateOfReg", Date.Now)
                rec = cmdClassHeads.ExecuteNonQuery
                If rec > 0 Then
                    MsgBox("Record Saved!", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "SuccessFul Transaction")
                End If
            End If
            loadDetails()
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            subClearTexts()
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
            queryType = Nothing
            Me.btnDelete.Enabled = False
            Me.btnUpdate.Enabled = False
            Me.btnSave.Enabled = True
        End Try
    End Sub
    Private Sub checkEntry()
        entryOk = False
        If Me.cboClassName.Text.Trim.Length <= 0 Then
            MsgBox("Missing Class Name", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboClassStream.Text.Trim.Length <= 0 Then
            MsgBox("Missing Class stream", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboClassYear.Text.Trim.Length <= 0 Then
            MsgBox("Missing Class year", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboStudNo.Text.Trim.Length <= 0 Then
            MsgBox("Missing Student Number", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboTeacherNo.Text.Trim.Length <= 0 Then
            MsgBox("Missing Teacher Number", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.txtStudName.Text.Trim.Length <= 0 Then
            MsgBox("Missing Student Leader Name", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.txtTeacherName.Text.Trim.Length <= 0 Then
            MsgBox("Missing Class Teacher Name", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If
        entryOk = True
    End Sub
    Private Function checkIfExists(ByVal classId As Integer)
        recordExists = True
        cmdClassHeads.Connection = conn
        cmdClassHeads.CommandType = CommandType.Text
        cmdClassHeads.CommandText = "SELECT * FROM  tblClassHeads WHERE (classId=@classId) AND (status='True')"
        cmdClassHeads.Parameters.Clear()
        cmdClassHeads.Parameters.AddWithValue("@classId", classId)
        reader = cmdClassHeads.ExecuteReader
        If reader.HasRows Then
            recordExists = True
        Else
            recordExists = False
        End If
        reader.Close()
    End Function
    Private Sub checkIfExistsOne()
        recordExists = True
        cmdClassHeads.Connection = conn
        cmdClassHeads.CommandType = CommandType.Text
        cmdClassHeads.CommandText = "SELECT * FROM vwClassHeads WHERE (empNo=@empNo) AND (admNo=@admNo) AND (className=@className) " & _
            vbNewLine & " AND (stream=@stream) AND (year=@year) AND (classHeadStatus='True') AND (schoolStaffStatus='True') AND " & _
            vbNewLine & " (studentStatus='True') AND (classStatus='True')"
        cmdClassHeads.Parameters.Clear()
        cmdClassHeads.Parameters.AddWithValue("@empNo", Me.cboTeacherNo.Text.Trim)
        cmdClassHeads.Parameters.AddWithValue("@className", Me.cboClassName.Text.Trim)
        cmdClassHeads.Parameters.AddWithValue("@Year", Me.cboClassYear.Text.Trim)
        cmdClassHeads.Parameters.AddWithValue("@Stream", Me.cboClassStream.Text.Trim)
        cmdClassHeads.Parameters.AddWithValue("@admNo", Me.cboStudNo.Text.Trim)
        reader = cmdClassHeads.ExecuteReader
        If reader.HasRows Then
            recordExists = True
        Else
            recordExists = False
        End If
        reader.Close()
    End Sub
    Private Sub GetClassID()
        cmdClassHeads.Connection = conn
        cmdClassHeads.CommandType = CommandType.Text
        cmdClassHeads.CommandText = "SELECT * FROM  tblClasses WHERE (className=@className) AND (stream=@stream) AND (year=@year) AND (status='True')"
        cmdClassHeads.Parameters.Clear()
        cmdClassHeads.Parameters.AddWithValue("@className", Me.cboClassName.Text.Trim)
        cmdClassHeads.Parameters.AddWithValue("@stream", Me.cboClassStream.Text.Trim)
        cmdClassHeads.Parameters.AddWithValue("@year", Me.cboClassYear.Text.Trim)
        reader = cmdClassHeads.ExecuteReader
        If reader.HasRows Then
            While reader.Read
                Me.cboClassName.Tag = IIf(DBNull.Value.Equals(reader!classId), "", reader!classId)
            End While
        End If
        reader.Close()
    End Sub
    Private Sub subClearTexts()
        Me.cboClassName.Text = ""
        Me.cboClassName.SelectedIndex = -1
        Me.cboClassStream.Text = ""
        Me.cboClassStream.SelectedIndex = -1
        Me.cboClassYear.Text = ""
        Me.cboClassYear.SelectedIndex = -1
        Me.cboStudNo.Text = ""
        Me.cboStudNo.SelectedIndex = -1
        Me.cboTeacherNo.Text = ""
        Me.cboTeacherNo.SelectedIndex = -1
        Me.txtStudName.Text = ""
        Me.txtTeacherName.Text = ""
        Me.cboClassName.Tag = Nothing
        lstClassHeads.Tag = Nothing
    End Sub

    Private Sub btnUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdate.Click
        entryOk = False
        checkEntry()
        If entryOk = False Then
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            recordExists = True
            checkIfExistsOne()
            If recordExists = True Then
                MsgBox("Duplicate Data Found In The Database", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
                Exit Sub
            End If
            If recordExists = True Then
                MsgBox("Record Data For the Class Saved Already.Try Update", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
                Exit Sub
            End If
            Dim result As MsgBoxResult = MsgBox("Update Record?", MsgBoxStyle.Question + MsgBoxStyle.ApplicationModal + MsgBoxStyle.YesNo, "Confirm Transaction")
            If result = MsgBoxResult.Yes Then
                queryType = "UPDATE"
                cmdClassHeads.Connection = conn
                cmdClassHeads.CommandType = CommandType.StoredProcedure
                cmdClassHeads.CommandText = "sprocClassHeads"
                cmdClassHeads.Parameters.Clear()
                cmdClassHeads.Parameters.AddWithValue("@classHeadId", Me.lstClassHeads.Tag)
                cmdClassHeads.Parameters.AddWithValue("@className", Me.cboClassName.Text.Trim)
                cmdClassHeads.Parameters.AddWithValue("@classYear", Me.cboClassYear.Text.Trim)
                cmdClassHeads.Parameters.AddWithValue("@classStream", Me.cboClassStream.Text.Trim)
                cmdClassHeads.Parameters.AddWithValue("@admNo", Me.cboStudNo.Text.Trim)
                cmdClassHeads.Parameters.AddWithValue("@queryType", queryType.Trim)
                cmdClassHeads.Parameters.AddWithValue("@regBy", userName.Trim)
                cmdClassHeads.Parameters.AddWithValue("@empNo", Me.cboTeacherNo.Text.Trim)
                cmdClassHeads.Parameters.AddWithValue("@dateOfReg", Date.Now)
                rec = cmdClassHeads.ExecuteNonQuery
                If rec > 0 Then
                    MsgBox("Record Updated!", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "SuccessFul Transaction")
                End If
            End If
            loadDetails()
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            subClearTexts()
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
            queryType = Nothing
            Me.lstClassHeads.Enabled = True
            Me.cboClassName.Enabled = True
            Me.cboClassStream.Enabled = True
            Me.cboClassYear.Enabled = True
            Me.btnDelete.Enabled = False
            Me.btnUpdate.Enabled = False
            Me.btnSave.Enabled = True
        End Try
    End Sub

    Private Sub lstClassHeads_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lstClassHeads.Click
        If Me.lstClassHeads.SelectedItems.Count = 1 Then
            If Me.lstClassHeads.SelectedItems(0).SubItems(3).Text.Trim.Length <= 0 And Me.lstClassHeads.SelectedItems(0).SubItems(5).Text.Trim.Length <= 0 Then
                MsgBox("No Class Heads Data To Update", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
                Exit Sub
            End If
            Me.cboClassName.Text = Me.lstClassHeads.SelectedItems(0).Text
            Me.cboClassStream.Text = Me.lstClassHeads.SelectedItems(0).SubItems(1).Text
            Me.cboClassYear.Text = Me.lstClassHeads.SelectedItems(0).SubItems(2).Text
            Me.cboStudNo.Text = Me.lstClassHeads.SelectedItems(0).SubItems(3).Text
            Me.cboTeacherNo.Text = Me.lstClassHeads.SelectedItems(0).SubItems(5).Text
            Me.txtTeacherName.Text = Me.lstClassHeads.SelectedItems(0).SubItems(6).Text
            Me.txtStudName.Text = Me.lstClassHeads.SelectedItems(0).SubItems(4).Text
            lstClassHeads.Tag = Me.lstClassHeads.SelectedItems(0).Tag
            Me.lstClassHeads.Enabled = False
            Me.cboClassName.Enabled = False
            Me.cboClassStream.Enabled = False
            Me.cboClassYear.Enabled = False
            Me.btnDelete.Enabled = True
            Me.btnUpdate.Enabled = True
            Me.btnSave.Enabled = False
        End If
    End Sub

    Private Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete.Click

        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            Dim result As MsgBoxResult = MsgBox("DELETE Record?", MsgBoxStyle.Question + MsgBoxStyle.ApplicationModal + MsgBoxStyle.YesNo, "Confirm Transaction")
            If result = MsgBoxResult.Yes Then
                queryType = "DELETE"
                cmdClassHeads.Connection = conn
                cmdClassHeads.CommandType = CommandType.StoredProcedure
                cmdClassHeads.CommandText = "sprocClassHeads"
                cmdClassHeads.Parameters.Clear()
                cmdClassHeads.Parameters.AddWithValue("@classHeadId", Me.lstClassHeads.Tag)
                cmdClassHeads.Parameters.AddWithValue("@className", Me.cboClassName.Text.Trim)
                cmdClassHeads.Parameters.AddWithValue("@classYear", Me.cboClassYear.Text.Trim)
                cmdClassHeads.Parameters.AddWithValue("@classStream", Me.cboClassStream.Text.Trim)
                cmdClassHeads.Parameters.AddWithValue("@admNo", Me.cboStudNo.Text.Trim)
                cmdClassHeads.Parameters.AddWithValue("@queryType", queryType.Trim)
                cmdClassHeads.Parameters.AddWithValue("@regBy", userName.Trim)
                cmdClassHeads.Parameters.AddWithValue("@empNo", Me.cboTeacherNo.Text.Trim)
                cmdClassHeads.Parameters.AddWithValue("@dateOfReg", Date.Now)
                rec = cmdClassHeads.ExecuteNonQuery
                If rec > 0 Then
                    MsgBox("Record Deleted!", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "SuccessFul Transaction")
                End If
            End If
            loadDetails()
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            subClearTexts()
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
            queryType = Nothing
            Me.lstClassHeads.Enabled = True
            Me.cboClassName.Enabled = True
            Me.cboClassStream.Enabled = True
            Me.cboClassYear.Enabled = True
            Me.btnDelete.Enabled = False
            Me.btnUpdate.Enabled = False
            Me.btnSave.Enabled = True
        End Try
    End Sub
End Class