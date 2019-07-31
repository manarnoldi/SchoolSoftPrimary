Imports System.Data.SqlClient
Public Class frmStudClass
    Dim recordExist As Boolean = True
    Dim entryOk As Boolean = False
    Dim reader As SqlDataReader
    Dim cmdStudClass As New SqlCommand
    Dim rec As Integer = 0
    Dim queryType As String = Nothing

    Private Sub frmStudClass_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
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

    Private Sub frmStudClass_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint

        If Me.IsMdiChild Then
            Dim PnlLoc As New Point
            PnlLoc.X = CInt((Me.Width - Me.pnlStudClass.Width) / 2)
            PnlLoc.Y = CInt((Me.Height - Me.pnlStudClass.Height) / 2.5)
            Me.pnlStudClass.Location = PnlLoc
        Else
            Me.pnlStudClass.Dock = DockStyle.Fill
        End If
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub
    Private Sub loadCombos()
        Me.cboClassName.Text = ""
        Me.cboClassStream.Text = ""
        Me.cboClassYear.Text = ""

        Me.cboClassName.SelectedIndex = -1
        Me.cboClassStream.SelectedIndex = -1
        Me.cboClassYear.SelectedIndex = -1

        Me.cboClassName.Items.Clear()
        Me.cboClassStream.Items.Clear()
        Me.cboClassYear.Items.Clear()

        cmdStudClass.Connection = conn
        cmdStudClass.CommandType = CommandType.Text
        cmdStudClass.CommandText = "SELECT DISTINCT className FROM tblClasses WHERE (status='True') ORDER BY className"
        cmdStudClass.Parameters.Clear()
        reader = cmdStudClass.ExecuteReader
        If reader.HasRows Then
            While reader.Read
                Me.cboClassName.Items.Add(IIf(DBNull.Value.Equals(reader!className), "", reader!className))
            End While
        End If
        reader.Close()

        cmdStudClass.CommandText = "SELECT DISTINCT stream FROM tblClasses WHERE (status='True') ORDER BY stream"
        cmdStudClass.Parameters.Clear()
        reader = cmdStudClass.ExecuteReader
        If reader.HasRows Then
            While reader.Read
                Me.cboClassStream.Items.Add(IIf(DBNull.Value.Equals(reader!stream), "", reader!stream))
            End While
        End If
        reader.Close()

        cmdStudClass.CommandText = "SELECT DISTINCT year FROM tblClasses WHERE (status='True') ORDER BY year"
        cmdStudClass.Parameters.Clear()
        reader = cmdStudClass.ExecuteReader
        If reader.HasRows Then
            While reader.Read
                Me.cboClassYear.Items.Add(IIf(DBNull.Value.Equals(reader!year), "", reader!year))
            End While
        End If
        reader.Close()
    End Sub
    Private Sub btnLoad_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLoad.Click
        If Me.txtAdmNo.Text.Trim.Length <= 0 Then
            MsgBox("Enter Admission Number", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If
        Me.txtFullName.Text = ""
        Me.txtAdmNoVw.Text = ""
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            cmdStudClass.Connection = conn
            cmdStudClass.CommandType = CommandType.Text
            cmdStudClass.CommandText = "SELECT admNo,FullName FROM tblStudDetails WHERE (status='True') AND (admNo=@admNo)"
            cmdStudClass.Parameters.Clear()
            cmdStudClass.Parameters.AddWithValue("@admNo", Me.txtAdmNo.Text.Trim)
            reader = cmdStudClass.ExecuteReader
            If reader.HasRows Then
                While reader.Read
                    Me.txtAdmNoVw.Text = (IIf(DBNull.Value.Equals(reader!admNo), "", reader!admNo))
                    Me.txtFullName.Text = (IIf(DBNull.Value.Equals(reader!FullName), "", reader!FullName))
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
    Private Sub txtAdmNo_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtAdmNo.TextChanged
        If Me.txtAdmNo.Text = "" Then
            Me.txtAdmNoVw.Text = ""
            Me.txtFullName.Text = ""
            Me.lstStudentClass.Items.Clear()
        End If
    End Sub

    Private Sub cboClassName_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboClassName.SelectedIndexChanged, cboClassStream.SelectedIndexChanged, cboClassYear.SelectedIndexChanged
        If Me.cboClassName.Text.Trim.Length <= 0 Or Me.cboClassStream.Text.Trim.Length <= 0 Or Me.cboClassYear.Text.Trim.Length <= 0 Then
            Exit Sub
        End If
        Try
            Me.lstStudentClass.Items.Clear()
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            cmdStudClass.Connection = conn
            cmdStudClass.CommandType = CommandType.Text
            cmdStudClass.CommandText = "SELECT * FROM vwStudClass WHERE (studStatus='True') AND (classStatus='True') " & _
                vbNewLine & " AND (classStudStatus='True') AND (className=@className) AND (stream=@stream) AND (year=@year)"
            cmdStudClass.Parameters.Clear()
            cmdStudClass.Parameters.AddWithValue("@className", Me.cboClassName.Text.Trim)
            cmdStudClass.Parameters.AddWithValue("@stream", Me.cboClassStream.Text.Trim)
            cmdStudClass.Parameters.AddWithValue("@year", Me.cboClassYear.Text.Trim)
            reader = cmdStudClass.ExecuteReader
            If reader.HasRows Then
                While reader.Read
                    li = Me.lstStudentClass.Items.Add(IIf(DBNull.Value.Equals(reader!admNo), "", reader!admNo))
                    li.SubItems.Add(IIf(DBNull.Value.Equals(reader!FullName), "", reader!FullName))
                    li.SubItems.Add(IIf(DBNull.Value.Equals(reader!className), "", reader!className))
                    li.SubItems.Add(IIf(DBNull.Value.Equals(reader!stream), "", reader!stream))
                    li.SubItems.Add(IIf(DBNull.Value.Equals(reader!Year), "", reader!Year))
                    li.Tag = (IIf(DBNull.Value.Equals(reader!classStudId), "", reader!classStudId))
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
        If Me.txtAdmNoVw.Text.Trim.Length <= 0 Then
            Exit Sub
        End If
        Try
            Me.lstStudentClass.Items.Clear()
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            cmdStudClass.Connection = conn
            cmdStudClass.CommandType = CommandType.Text
            cmdStudClass.CommandText = "SELECT * FROM vwStudClass WHERE (studStatus='True') AND (classStatus='True') " & _
                vbNewLine & " AND (classStudStatus='True') AND (admNo=@admNo)"
            cmdStudClass.Parameters.Clear()
            cmdStudClass.Parameters.AddWithValue("@admNo", Me.txtAdmNo.Text.Trim)
            reader = cmdStudClass.ExecuteReader
            If reader.HasRows Then
                While reader.Read
                    li = Me.lstStudentClass.Items.Add(IIf(DBNull.Value.Equals(reader!admNo), "", reader!admNo))
                    li.SubItems.Add(IIf(DBNull.Value.Equals(reader!FullName), "", reader!FullName))
                    li.SubItems.Add(IIf(DBNull.Value.Equals(reader!className), "", reader!className))
                    li.SubItems.Add(IIf(DBNull.Value.Equals(reader!stream), "", reader!stream))
                    li.SubItems.Add(IIf(DBNull.Value.Equals(reader!Year), "", reader!Year))
                    li.Tag = (IIf(DBNull.Value.Equals(reader!classStudId), "", reader!classStudId))
                End While
            End If
            reader.Close()
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
            Me.btnSave.Enabled = True
            Me.btnUpdate.Enabled = False
            Me.btnDelete.Enabled = False
        End Try
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
            recordExist = True
            checkIfStudentAssigned()
            If recordExist = True Then
                MsgBox("Student Already Assigned To A different Class For The Year", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
                Exit Sub
            End If
            recordExist = True
            checkForExistence()
            If recordExist = True Then
                MsgBox("Record Exists", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
                Exit Sub
            End If
            Dim result As MsgBoxResult = MsgBox("Save Record?", MsgBoxStyle.Question + MsgBoxStyle.ApplicationModal + MsgBoxStyle.YesNo, "Confirm Transaction")
            If result = MsgBoxResult.Yes Then
                queryType = "INSERT"
                cmdStudClass.Connection = conn
                cmdStudClass.CommandType = CommandType.StoredProcedure
                cmdStudClass.CommandText = "sprocStudentClass"
                cmdStudClass.Parameters.Clear()
                cmdStudClass.Parameters.AddWithValue("@className", Me.cboClassName.Text.Trim)
                cmdStudClass.Parameters.AddWithValue("@classYear", Me.cboClassYear.Text.Trim)
                cmdStudClass.Parameters.AddWithValue("@classStream", Me.cboClassStream.Text.Trim)
                cmdStudClass.Parameters.AddWithValue("@admNo", Me.txtAdmNoVw.Text.Trim)
                cmdStudClass.Parameters.AddWithValue("@queryType", queryType.Trim)
                cmdStudClass.Parameters.AddWithValue("@regBy", userName.Trim)
                cmdStudClass.Parameters.AddWithValue("@dateOfReg", Date.Now)
                rec = cmdStudClass.ExecuteNonQuery
                If rec > 0 Then
                    MsgBox("Record Saved!", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "SuccessFul Transaction")
                End If
            End If
            clearTexts()
            'loadCombos()
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
            Me.txtAdmNo.Enabled = True
            Me.btnDelete.Enabled = False
            Me.btnUpdate.Enabled = False
            Me.btnSave.Enabled = True
            queryType = Nothing
        End Try
    End Sub
    Private Sub checkEntry()
        entryOk = False
        If Me.txtFullName.Text.Trim.Length <= 0 Or Me.txtAdmNoVw.Text.Trim.Length <= 0 Then
            MsgBox("Student Details Not Found.Reload Student", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
            entryOk = False
        ElseIf Me.txtAdmNo.Text.Trim <> Me.txtAdmNoVw.Text.Trim Then
            MsgBox("Load Student To Update", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
            entryOk = False
        ElseIf Me.cboClassName.Text.Length <= 0 Then
            MsgBox("Load Class Name", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
            entryOk = False
        ElseIf Me.cboClassStream.Text.Length <= 0 Then
            MsgBox("Load Class Stream", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
            entryOk = False
        ElseIf Me.cboClassYear.Text.Length <= 0 Then
            MsgBox("Load Class Year", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
            entryOk = False
        End If
        entryOk = True
    End Sub
    Private Sub checkForExistence()
        recordExist = True
        cmdStudClass.Connection = conn
        cmdStudClass.CommandType = CommandType.Text
        cmdStudClass.CommandText = "SELECT * FROM vwStudClass WHERE (studStatus='True') AND (classStatus='True') AND (admNo=@admNo)" & _
                vbNewLine & " AND (classStudStatus='True') AND (className=@className) AND (stream=@stream) AND (year=@year)"
        cmdStudClass.Parameters.Clear()
        cmdStudClass.Parameters.AddWithValue("@admNo", Me.txtAdmNoVw.Text.Trim)
        cmdStudClass.Parameters.AddWithValue("@className", Me.cboClassName.Text.Trim)
        cmdStudClass.Parameters.AddWithValue("@stream", Me.cboClassStream.Text.Trim)
        cmdStudClass.Parameters.AddWithValue("@year", Me.cboClassYear.Text.Trim)
        reader = cmdStudClass.ExecuteReader
        If reader.HasRows Then
            recordExist = True
        Else
            recordExist = False
        End If
        reader.Close()
    End Sub
    Private Sub clearTexts()
        Me.txtAdmNo.Text = ""
        Me.txtAdmNoVw.Text = ""
        Me.txtFullName.Text = ""
        'Me.cboClassYear.Text = ""
        'Me.cboClassStream.Text = ""
        'Me.cboClassName.Text = ""
        'Me.cboClassYear.SelectedIndex = -1
        'Me.cboClassStream.SelectedIndex = -1
        'Me.cboClassName.SelectedIndex = -1
        Me.lstStudentClass.Items.Clear()
        Me.lstStudentClass.Tag = Nothing
    End Sub
    Private Function checkIfStudentAssigned()
        recordExist = True
        cmdStudClass.Connection = conn
        cmdStudClass.CommandType = CommandType.Text
        cmdStudClass.CommandText = "SELECT * FROM vwStudClass WHERE (studStatus='True') AND (classStatus='True') AND (admNo=@admNo)" & _
                vbNewLine & " AND (classStudStatus='True') AND (year=@year)"
        cmdStudClass.Parameters.Clear()
        cmdStudClass.Parameters.AddWithValue("@admNo", Me.txtAdmNoVw.Text.Trim)
        cmdStudClass.Parameters.AddWithValue("@year", Me.cboClassYear.Text.Trim)
        reader = cmdStudClass.ExecuteReader
        If reader.HasRows Then
            recordExist = True
        Else
            recordExist = False
        End If
        reader.Close()
    End Function

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
            recordExist = True
            checkForExistence()
            If recordExist = True Then
                MsgBox("Record Exists", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
                Exit Sub
            End If
            Dim result As MsgBoxResult = MsgBox("Update Record?", MsgBoxStyle.Question + MsgBoxStyle.ApplicationModal + MsgBoxStyle.YesNo, "Confirm Transaction")
            If result = MsgBoxResult.Yes Then
                queryType = "UPDATE"
                cmdStudClass.Connection = conn
                cmdStudClass.CommandType = CommandType.StoredProcedure
                cmdStudClass.CommandText = "sprocStudentClass"
                cmdStudClass.Parameters.Clear()
                cmdStudClass.Parameters.AddWithValue("@classStudId", Me.lstStudentClass.Tag)
                cmdStudClass.Parameters.AddWithValue("@className", Me.cboClassName.Text.Trim)
                cmdStudClass.Parameters.AddWithValue("@classYear", Me.cboClassYear.Text.Trim)
                cmdStudClass.Parameters.AddWithValue("@classStream", Me.cboClassStream.Text.Trim)
                cmdStudClass.Parameters.AddWithValue("@admNo", Me.txtAdmNoVw.Text.Trim)
                cmdStudClass.Parameters.AddWithValue("@queryType", queryType.Trim)
                cmdStudClass.Parameters.AddWithValue("@regBy", userName.Trim)
                cmdStudClass.Parameters.AddWithValue("@dateOfReg", Date.Now)
                rec = cmdStudClass.ExecuteNonQuery
                If rec > 0 Then
                    MsgBox("Record Updated!", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "SuccessFul Transaction")
                End If
            End If
            clearTexts()
            loadCombos()
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
            Me.txtAdmNo.Enabled = True
            Me.btnDelete.Enabled = False
            Me.btnUpdate.Enabled = False
            Me.btnSave.Enabled = True
            queryType = Nothing
        End Try
    End Sub

    Private Sub lstStudentClass_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lstStudentClass.Click
        If Me.lstStudentClass.SelectedItems.Count = 1 Then
            Me.txtAdmNo.Text = Me.lstStudentClass.SelectedItems(0).Text.Trim
            Me.txtAdmNoVw.Text = Me.lstStudentClass.SelectedItems(0).Text.Trim
            Me.txtFullName.Text = Me.lstStudentClass.SelectedItems(0).SubItems(1).Text.Trim
            Me.lstStudentClass.Tag = Me.lstStudentClass.SelectedItems(0).Tag
            Me.cboClassName.Text = Me.lstStudentClass.SelectedItems(0).SubItems(2).Text.Trim
            Me.cboClassStream.Text = Me.lstStudentClass.SelectedItems(0).SubItems(3).Text.Trim
            Me.cboClassYear.Text = Me.lstStudentClass.SelectedItems(0).SubItems(4).Text.Trim
             Me.txtAdmNo.Enabled = False
            Me.btnSave.Enabled = False
            Me.btnUpdate.Enabled = True
            Me.btnDelete.Enabled = True
        End If
    End Sub

    Private Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            Dim result As MsgBoxResult = MsgBox("Delete Record?", MsgBoxStyle.Question + MsgBoxStyle.ApplicationModal + MsgBoxStyle.YesNo, "Confirm Transaction")
            If result = MsgBoxResult.Yes Then
                queryType = "DELETE"
                cmdStudClass.Connection = conn
                cmdStudClass.CommandType = CommandType.StoredProcedure
                cmdStudClass.CommandText = "sprocStudentClass"
                cmdStudClass.Parameters.Clear()
                cmdStudClass.Parameters.AddWithValue("@classStudId", Me.lstStudentClass.Tag)
                cmdStudClass.Parameters.AddWithValue("@className", Me.cboClassName.Text.Trim)
                cmdStudClass.Parameters.AddWithValue("@classYear", Me.cboClassYear.Text.Trim)
                cmdStudClass.Parameters.AddWithValue("@classStream", Me.cboClassStream.Text.Trim)
                cmdStudClass.Parameters.AddWithValue("@admNo", Me.txtAdmNoVw.Text.Trim)
                cmdStudClass.Parameters.AddWithValue("@queryType", queryType.Trim)
                cmdStudClass.Parameters.AddWithValue("@regBy", userName.Trim)
                cmdStudClass.Parameters.AddWithValue("@dateOfReg", Date.Now)
                rec = cmdStudClass.ExecuteNonQuery
                If rec > 0 Then
                    MsgBox("Record Deleted!", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "SuccessFul Transaction")
                End If
            End If
            clearTexts()
            loadCombos()
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
            Me.txtAdmNo.Enabled = True
            Me.btnDelete.Enabled = False
            Me.btnUpdate.Enabled = False
            Me.btnSave.Enabled = True
            queryType = Nothing
        End Try
    End Sub
End Class