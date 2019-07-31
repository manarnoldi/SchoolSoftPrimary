Imports System.Data.SqlClient
Public Class frmUsers
    Dim RecAffected As Integer = 0
    Dim RecAffected1 As Integer = 0
    Dim reader As SqlDataReader
    Dim reader1 As SqlDataReader
    Dim cmdUsers As New SqlCommand
    Dim cmdUsers1 As New SqlCommand
    Private Sub frmUsers_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        My.Settings.Reload()
        loadCombos()
        Me.btnDelete.Enabled = True
        Me.btnUpdate.Enabled = True
    End Sub
    Private Sub loadCombos()
        Try
            dbconnection()
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            cmdUsers.CommandText = "SELECT DISTINCT empType FROM  tblSchoolStaff WHERE (status='True') ORDER BY empType"
            cmdUsers.Connection = conn
            cmdUsers.CommandType = CommandType.Text
            cmdUsers.Parameters.Clear()
            reader = cmdUsers.ExecuteReader

            If reader.HasRows Then
                Me.cboStaffType.Items.Clear()
                Me.cboStaffType.Text = ""
                While reader.Read
                    Me.cboStaffType.Items.Add(IIf(DBNull.Value.Equals(reader!empType), "", reader!empType))
                End While
            End If
            reader.Close()

            cmdUsers.CommandText = "SELECT DISTINCT contractType FROM  tblSchoolStaff WHERE (status='True') ORDER BY contractType"
            cmdUsers.Parameters.Clear()
            reader = cmdUsers.ExecuteReader

            If reader.HasRows Then
                Me.cboContactType.Items.Clear()
                Me.cboContactType.Text = ""
                While reader.Read
                    Me.cboContactType.Items.Add(IIf(DBNull.Value.Equals(reader!contractType), "", reader!contractType))
                End While
            End If
            reader.Close()

            cmdUsers.CommandText = "SELECT DISTINCT domainName FROM tblDomains WHERE (status='True') ORDER BY domainName"
            cmdUsers.Parameters.Clear()
            reader = cmdUsers.ExecuteReader

            If reader.HasRows Then
                Me.cboDomain.Items.Clear()
                Me.cboDomain.Text = ""
                While reader.Read
                    Me.cboDomain.Items.Add(IIf(DBNull.Value.Equals(reader!domainName), "", reader!domainName))
                End While
            End If
            reader.Close()

        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Encountered")
        Finally
            reader.Close()
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub
    Private Sub frmUsers_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint
        If Me.IsMdiChild Then
            Dim PnlLoc As New Point
            PnlLoc.X = CInt((Me.Width - Me.pnlUsers.Width) / 2)
            PnlLoc.Y = CInt((Me.Height - Me.pnlUsers.Height) / 2.5)
            Me.pnlUsers.Location = PnlLoc
        Else
            Me.pnlUsers.Dock = DockStyle.Fill
        End If
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub cboContactType_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboContactType.SelectedIndexChanged, cboStaffType.SelectedIndexChanged
        If Me.cboStaffType.Text = "" Or Me.cboContactType.Text = "" Then
            Exit Sub
        End If
        Try
            dbconnection()
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            cmdUsers.CommandText = "SELECT empNo FROM  tblSchoolStaff WHERE (status='True') AND (contractType=@contractType) AND (empType=@empType) ORDER BY empNo"
            cmdUsers.Connection = conn
            cmdUsers.CommandType = CommandType.Text
            cmdUsers.Parameters.Clear()
            cmdUsers.Parameters.AddWithValue("@contractType", Me.cboContactType.Text.Trim)
            cmdUsers.Parameters.AddWithValue("@empType", Me.cboStaffType.Text.Trim)
            reader = cmdUsers.ExecuteReader

            If reader.HasRows Then
                Me.cboEmpNo.Items.Clear()
                Me.cboEmpNo.Text = ""
                While reader.Read
                    Me.cboEmpNo.Items.Add(IIf(DBNull.Value.Equals(reader!empNo), "", reader!empNo))
                End While
            End If
            reader.Close()
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Encountered")
        Finally
            reader.Close()
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
            Me.txtPassword.Enabled = True
            Me.txtUserName.Enabled = True
            Me.btnSave.Enabled = True
            Me.btnDelete.Enabled = False
            btnUpdate.Enabled = False
            Me.txtFullName.Text = ""
            Me.txtPassword.Text = ""
            Me.txtUserName.Text = ""
        End Try
    End Sub

    Private Sub cboEmpNo_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboEmpNo.SelectedIndexChanged
        If Me.cboEmpNo.Text = "" Then
            'MsgBox("Missing Employee Number", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If
        Try
            dbconnection()
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            cmdUsers.CommandText = "SELECT  FullName FROM  tblSchoolStaff WHERE (status='True') AND (empNo=@empNo)"
            cmdUsers.Connection = conn
            cmdUsers.CommandType = CommandType.Text
            cmdUsers.Parameters.Clear()
            cmdUsers.Parameters.AddWithValue("@empNo", Me.cboEmpNo.Text.Trim)
            reader = cmdUsers.ExecuteReader

            If reader.HasRows Then
                Me.txtFullName.Text = ""
                While reader.Read
                    Me.txtFullName.Text = IIf(DBNull.Value.Equals(reader!FullName), "NOT AVAILABLE", reader!FullName)
                End While
                reader.Close()
                cmdUsers1.Connection = conn
                cmdUsers1.CommandType = CommandType.Text
                cmdUsers1.CommandText = "SELECT  userId,userName,passWord,domainName FROM vwUsersDetails WHERE (empNo=@empNo) and (userStatus='True')"
                cmdUsers1.Parameters.Clear()
                cmdUsers1.Parameters.AddWithValue("@empNo", Me.cboEmpNo.Text.Trim)
                reader1 = cmdUsers1.ExecuteReader
                If reader1.HasRows Then
                    While reader1.Read
                        Me.cboDomain.Text = IIf(DBNull.Value.Equals(reader1!domainName), "", reader1!domainName)
                        Me.txtPassword.Text = IIf(DBNull.Value.Equals(reader1!passWord), "", reader1!passWord)
                        Me.txtUserName.Text = IIf(DBNull.Value.Equals(reader1!userName), "", reader1!userName)
                        Me.txtUserName.Tag = IIf(DBNull.Value.Equals(reader1!userId), "", reader1!userId)
                    End While
                    Me.txtPassword.Enabled = False
                    Me.txtUserName.Enabled = False
                    Me.btnSave.Enabled = False
                    Me.btnDelete.Enabled = True
                    btnUpdate.Enabled = True
                Else
                    Me.txtPassword.Enabled = True
                    Me.txtUserName.Enabled = True
                    Me.btnSave.Enabled = True
                    Me.btnDelete.Enabled = False
                    btnUpdate.Enabled = False
                    'Me.txtFullName.Text = ""
                    Me.txtPassword.Text = ""
                    Me.txtUserName.Text = ""
                    Me.cboDomain.SelectedIndex = -1
                End If
                reader1.Close()
            End If
            reader.Close()

        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Encountered")
        Finally
            reader.Close()
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub

    Private Sub cboDomain_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboDomain.SelectedIndexChanged
        If Me.cboDomain.Text = "" Then
            Exit Sub
        End If
        Try
            dbconnection()
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            cmdUsers.CommandText = "SELECT domainId FROM tblDomains WHERE (domainName=@domainName) AND (status='True')"
            cmdUsers.Connection = conn
            cmdUsers.CommandType = CommandType.Text
            cmdUsers.Parameters.Clear()
            cmdUsers.Parameters.AddWithValue("@domainName", Me.cboDomain.Text.Trim)
            reader = cmdUsers.ExecuteReader

            If reader.HasRows Then
                Me.cboDomain.Tag = Nothing
                While reader.Read
                    Me.cboDomain.Tag = (IIf(DBNull.Value.Equals(reader!domainId), "", reader!domainId))
                End While
            End If
            reader.Close()
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Encountered")
        Finally
            reader.Close()
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        If Me.cboDomain.Text.Trim.Length <= 0 Then
            MsgBox("Missing DomainName", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboEmpNo.Text.Trim.Length <= 0 Then
            MsgBox("Missing Employee Number", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.txtUserName.Text.Trim.Length <= 0 Then
            MsgBox("Missing UserName", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.txtPassword.Text.Trim.Length <= 0 Then
            MsgBox("Missing Password", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If
        Try
            dbconnection()
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            cmdUsers.CommandText = "SELECT * FROM tblUsers WHERE (userName=@userName) AND (status='True')"
            cmdUsers.Connection = conn
            cmdUsers.CommandType = CommandType.Text
            cmdUsers.Parameters.Clear()
            cmdUsers.Parameters.AddWithValue("@userName", Me.txtUserName.Text.Trim)
            reader = cmdUsers.ExecuteReader

            If (reader.HasRows) Then
                MsgBox("UserName Already Exits", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "UNSuccessFull Transaction")
                If conn.State = ConnectionState.Open Then
                    conn.Close()
                End If
                reader.Close()
                Exit Sub
            End If
            reader.Close()
            
            cmdUsers.CommandType = CommandType.StoredProcedure
            cmdUsers.Connection = conn
            cmdUsers.CommandText = "sprocUsersInsert"
            cmdUsers.Parameters.Clear()
            cmdUsers.Parameters.AddWithValue("@domainId", Me.cboDomain.Tag)
            cmdUsers.Parameters.AddWithValue("@userName1", Me.txtUserName.Text.Trim)
            cmdUsers.Parameters.AddWithValue("@passWord", Me.txtPassword.Text.Trim)
            cmdUsers.Parameters.AddWithValue("@dateOfReg", Date.Now)
            cmdUsers.Parameters.AddWithValue("@empNo", Me.cboEmpNo.Text)
            cmdUsers.Parameters.AddWithValue("@regBy", userName)
            RecAffected = cmdUsers.ExecuteNonQuery()
            If RecAffected > 0 Then
                MsgBox("Record Saved", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "SuccessFull Transaction")
            End If
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Encountered")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
            Me.txtPassword.Enabled = True
            Me.txtUserName.Enabled = True
            Me.btnSave.Enabled = True
            Me.btnDelete.Enabled = False
            btnUpdate.Enabled = False
            Me.cboDomain.SelectedIndex = -1
            Me.cboEmpNo.SelectedIndex = -1
            Me.cboContactType.SelectedIndex = -1
            Me.cboStaffType.SelectedIndex = -1
            Me.txtFullName.Text = ""
            Me.txtPassword.Text = ""
            Me.txtUserName.Text = ""
            Me.cboDomain.Tag = Nothing
        End Try
    End Sub

    Private Sub btnUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdate.Click
        If Not (Me.txtPassword.Enabled = True) Then
            If Me.cboDomain.Text.Trim.Length <= 0 Then
                MsgBox("Missing DomainName", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
                Exit Sub
            End If
            Dim result As MsgBoxResult = MsgBox("Update Record?", MsgBoxStyle.Question + MsgBoxStyle.YesNo + MsgBoxStyle.ApplicationModal, "Confirm Update")
            If result = MsgBoxResult.No Then
                Exit Sub
            End If
            Try
                dbconnection()
                If conn.State = ConnectionState.Closed Then
                    conn.Open()
                End If
                cmdUsers.CommandText = "SELECT * FROM tblUsers WHERE (userId=@userId) AND (domainId=@domainId)"
                cmdUsers.Connection = conn
                cmdUsers.CommandType = CommandType.Text
                cmdUsers.Parameters.Clear()
                cmdUsers.Parameters.AddWithValue("@domainId", Me.cboDomain.Tag)
                cmdUsers.Parameters.AddWithValue("@userId", Me.txtUserName.Tag)
                reader = cmdUsers.ExecuteReader

                If (reader.HasRows) Then
                    MsgBox("No changes In the Record", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "UNSuccessFull Transaction")
                    If conn.State = ConnectionState.Open Then
                        conn.Close()
                    End If
                    reader.Close()
                    Exit Sub
                End If
                reader.Close()
                
                cmdUsers.CommandType = CommandType.StoredProcedure
                cmdUsers.Connection = conn
                cmdUsers.CommandText = "sprocUsersUpdate"
                cmdUsers.Parameters.Clear()
                cmdUsers.Parameters.AddWithValue("@domainId", Me.cboDomain.Tag)
                cmdUsers.Parameters.AddWithValue("@domainName", Me.cboDomain.Text)
                cmdUsers.Parameters.AddWithValue("@userId", Me.txtUserName.Tag)
                cmdUsers.Parameters.AddWithValue("@userName", Me.txtUserName.Text)
                cmdUsers.Parameters.AddWithValue("@dateOfAction", Date.Now)
                cmdUsers.Parameters.AddWithValue("@doneBy", userName)
                RecAffected = cmdUsers.ExecuteNonQuery()

                If RecAffected > 0 Then
                    MsgBox("Record Updated", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "SuccessFull Transaction")
                End If
            Catch ex As Exception
                MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Encountered")
            Finally
                If conn.State = ConnectionState.Open Then
                    conn.Close()
                End If
                Me.txtPassword.Enabled = True
                Me.txtUserName.Enabled = True
                Me.btnSave.Enabled = True
                Me.btnDelete.Enabled = False
                btnUpdate.Enabled = False
                Me.cboDomain.SelectedIndex = -1
                Me.cboEmpNo.SelectedIndex = -1
                Me.cboContactType.SelectedIndex = -1
                Me.cboStaffType.SelectedIndex = -1
                Me.txtFullName.Text = ""
                Me.txtPassword.Text = ""
                Me.txtUserName.Text = ""
                Me.cboDomain.Tag = Nothing
            End Try
        End If
    End Sub

    Private Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        If Not (Me.txtPassword.Enabled = True) Then
            If Me.cboDomain.Text.Trim.Length <= 0 Then
                MsgBox("Missing DomainName", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
                Exit Sub
            End If
            Dim result As MsgBoxResult = MsgBox("Delete Record?", MsgBoxStyle.Question + MsgBoxStyle.YesNo + MsgBoxStyle.ApplicationModal, "Confirm Delete")
            If result = MsgBoxResult.No Then
                Exit Sub
            End If
            Try
                dbconnection()
                If conn.State = ConnectionState.Closed Then
                    conn.Open()
                End If

                cmdUsers.CommandType = CommandType.StoredProcedure
                cmdUsers.Connection = conn
                cmdUsers.CommandText = "sprocUsersDelete"
                cmdUsers.Parameters.Clear()
                cmdUsers.Parameters.AddWithValue("@userId", Me.txtUserName.Tag)
                cmdUsers.Parameters.AddWithValue("@userName", Me.txtUserName.Text)
                cmdUsers.Parameters.AddWithValue("@dateOfAction", Date.Now)
                cmdUsers.Parameters.AddWithValue("@doneBy", userName)
                RecAffected = cmdUsers.ExecuteNonQuery()

                If RecAffected > 0 Then
                    MsgBox("Record Deleted", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "SuccessFull Transaction")
                End If
            Catch ex As Exception
                MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Encountered")
            Finally
                If conn.State = ConnectionState.Open Then
                    conn.Close()
                End If
                Me.txtPassword.Enabled = True
                Me.txtUserName.Enabled = True
                Me.btnSave.Enabled = True
                Me.btnDelete.Enabled = False
                btnUpdate.Enabled = False
                Me.cboDomain.SelectedIndex = -1
                Me.cboEmpNo.SelectedIndex = -1
                Me.cboContactType.SelectedIndex = -1
                Me.cboStaffType.SelectedIndex = -1
                Me.txtFullName.Text = ""
                Me.txtPassword.Text = ""
                Me.txtUserName.Text = ""
                Me.cboDomain.Tag = Nothing
            End Try
        End If
    End Sub

    Private Sub txtUserName_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtUserName.LostFocus
        If Me.txtUserName.Enabled = True Then

        End If
    End Sub


End Class