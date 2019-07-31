Imports System.Data.SqlClient
Public Class frmAccDormStudents
    Dim reader As SqlDataReader
    Dim rec As Integer = 0
    Dim cmdDormStud As New SqlCommand
    Private Sub frmDormStudents_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()

            loadCombos()
            loadList()
            Me.btnUpdate.Enabled = False
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub
    Private Sub loadList()
        Me.lstDormDetails.Items.Clear()

        Me.cmdDormStud.Connection = conn
        Me.cmdDormStud.CommandType = CommandType.Text
        Me.cmdDormStud.CommandText = "SELECT * FROM  tblAccdormRegister ORDER BY dormName"
        Me.cmdDormStud.Parameters.Clear()
        reader = Me.cmdDormStud.ExecuteReader
        While reader.Read
            li = Me.lstDormDetails.Items.Add(IIf(DBNull.Value.Equals(reader!dormName), "", reader!dormName))
            li.SubItems.Add(IIf(DBNull.Value.Equals(reader!genderType), "", reader!genderType))
            li.SubItems.Add(IIf(DBNull.Value.Equals(reader!dormCapacity), "", reader!dormCapacity))
            li.SubItems.Add(IIf(DBNull.Value.Equals(reader!enrolled), "", reader!enrolled))
            li.SubItems.Add(IIf(DBNull.Value.Equals(reader!availableSpaces), "", reader!availableSpaces))
        End While
        reader.Close()
    End Sub

    Private Sub frmDormStudents_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint
        If Me.IsMdiChild Then
            Dim PnlLoc As New Point
            PnlLoc.X = CInt((Me.Width - Me.pnlRegisterDorms.Width) / 2)
            PnlLoc.Y = CInt((Me.Height - Me.pnlRegisterDorms.Height) / 2.5)
            Me.pnlRegisterDorms.Location = PnlLoc
        Else
            Me.pnlRegisterDorms.Dock = DockStyle.Fill
        End If
    End Sub

    Private Sub btnLoad_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLoad.Click
        Me.txtGender.Text = ""
        Me.txtStudentName.Text = ""
        Dim dormEnrolled As String = Nothing
        Dim gender As String = Nothing
        If Me.txtAdmNo.Text.Trim.Length <= 0 Then
            MsgBox("Missing Admission Number", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If

        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            Me.cmdDormStud.Connection = conn
            Me.cmdDormStud.CommandType = CommandType.Text
            Me.cmdDormStud.CommandText = "SELECT * FROM vwAccDormStudentDetails  WHERE (admNo=@admNo)"
            Me.cmdDormStud.Parameters.Clear()
            Me.cmdDormStud.Parameters.AddWithValue("@admNo", Me.txtAdmNo.Text.Trim)
            reader = Me.cmdDormStud.ExecuteReader
            If reader.HasRows = True Then
                While reader.Read
                    dormEnrolled = IIf(DBNull.Value.Equals(reader!dormName), "", reader!dormName)
                End While
                Dim result As MsgBoxResult = MsgBox("Student Already Enrolled IN " & dormEnrolled & _
                                                    vbNewLine & "Update Record?", MsgBoxStyle.Question + _
                                                    MsgBoxStyle.ApplicationModal + MsgBoxStyle.YesNo, "Record Found")
                If result = MsgBoxResult.No Then
                    Exit Sub
                End If
                Me.btnUpdate.Enabled = True
                Me.btnDelete.Enabled = True
                Me.btnSave.Enabled = False
                i = 0
                If Not (dormEnrolled = "") Then
                    For i = 0 To Me.lstDormDetails.Items.Count - 1
                        If Me.lstDormDetails.Items(i).Text = dormEnrolled Then
                            Me.lstDormDetails.Items(i).BackColor = Color.LightSeaGreen
                        End If
                    Next
                End If
                Me.cboHouse.Text = dormEnrolled
            ElseIf reader.HasRows = False Then
            End If
            reader.Close()

            Me.cmdDormStud.Connection = conn
            Me.cmdDormStud.CommandType = CommandType.Text
            Me.cmdDormStud.CommandText = "SELECT * FROM tblStudDetails  WHERE (admNo=@admNo) AND (status=1)"
            reader = Me.cmdDormStud.ExecuteReader
            Me.cmdDormStud.Parameters.Clear()
            Me.cmdDormStud.Parameters.AddWithValue("@admNo", Me.txtAdmNo.Text.Trim)
            If reader.HasRows = True Then
                While reader.Read
                    If IIf(DBNull.Value.Equals(reader!sex), "", reader!sex) = True Then
                        gender = "Male"
                    ElseIf IIf(DBNull.Value.Equals(reader!sex), "", reader!sex) = False Then
                        gender = "Female"
                    End If
                    Me.txtGender.Text = gender
                    Me.txtStudentName.Text = IIf(DBNull.Value.Equals(reader!FullName), "", reader!FullName)
                End While
            ElseIf reader.HasRows = False Then
                MsgBox("Student Not Found in the system!", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
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
    Private Sub loadCombos()
        Me.cboHouse.Items.Clear()
        Me.cboHouse.SelectedIndex = -1
        Me.cboHouse.Text = ""

        Me.cmdDormStud.Connection = conn
        Me.cmdDormStud.CommandType = CommandType.Text
        Me.cmdDormStud.CommandText = "SELECT DISTINCT dormName FROM  tblAccdormRegister ORDER BY dormName"
        Me.cmdDormStud.Parameters.Clear()
        reader = Me.cmdDormStud.ExecuteReader
        While reader.Read
            Me.cboHouse.Items.Add(IIf(DBNull.Value.Equals(reader!dormName), "", reader!dormName))
        End While
        reader.Close()
    End Sub

    Private Sub chckBoxAutoAssign_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chckBoxAutoAssign.CheckedChanged
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            If Me.chckBoxAutoAssign.Checked = True Then
                Dim genderType As String = Nothing
                If Me.txtStudentName.Text.Trim.Length <= 0 Then
                    MsgBox("Missing Student Name.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
                    Exit Sub
                ElseIf Me.txtGender.Text.Trim.Length <= 0 Then
                    MsgBox("Missing Dormitory Gender.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
                    Exit Sub
                End If
                Me.cmdDormStud.Connection = conn
                Me.cmdDormStud.CommandType = CommandType.Text
                Me.cmdDormStud.CommandText = "SELECT TOP 1 * FROM tblAccdormRegister WHERE (availableSpaces > 0) AND " & _
                    vbNewLine & " ((genderType='Mixed') or (genderType=@genderType))  ORDER BY availableSpaces DESC"
                Me.cmdDormStud.Parameters.Clear()
                Me.cmdDormStud.Parameters.AddWithValue("@genderType", Me.txtGender.Text.Trim)
                reader = Me.cmdDormStud.ExecuteReader
                If reader.HasRows = True Then
                    While reader.Read
                        Me.cboHouse.DropDownStyle = ComboBoxStyle.Simple
                        Me.cboHouse.Text = IIf(DBNull.Value.Equals(reader!dormName), "", reader!dormName)
                        Me.cboHouse.Enabled = False
                    End While
                ElseIf reader.HasRows = False Then
                    MsgBox("No Dorm Space Available.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
                End If
                reader.Close()
            Else
                Me.cboHouse.DropDownStyle = ComboBoxStyle.DropDownList
                Me.cboHouse.Text = ""
                Me.cboHouse.Enabled = True
            End If
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
            clearControls()

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
    Private Sub clearControls()
        Me.txtAdmNo.Text = ""
        Me.txtStudentName.Text = ""
        Me.chckBoxAutoAssign.Checked = False
        Me.cboHouse.Enabled = True
        Me.cboHouse.SelectedIndex = -1
        Me.cboHouse.DropDownStyle = ComboBoxStyle.DropDownList
        Me.txtGender.Text = ""
        Me.btnClose.Text = "Close"
        Me.btnSave.Enabled = True
        Me.btnUpdate.Enabled = False
        Me.btnDelete.Enabled = False
    End Sub
    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        If Me.cboHouse.Text = "" Then
            MsgBox("Missing House Name.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.txtAdmNo.Text = "" Then
            MsgBox("Missing admission number.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.txtStudentName.Text = "" Then
            MsgBox("Missing student name.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.txtGender.Text = "" Then
            MsgBox("Missing Gender.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()

            Dim result As MsgBoxResult = MsgBox("Assign Student?", MsgBoxStyle.Question + MsgBoxStyle.ApplicationModal + MsgBoxStyle.YesNo, "Confirm Transaction")
            If result = MsgBoxResult.No Then
                Exit Sub
            End If

            Me.cmdDormStud.Connection = conn
            Me.cmdDormStud.CommandType = CommandType.StoredProcedure
            Me.cmdDormStud.CommandText = "sprocAccDormStudDetails"
            Me.cmdDormStud.Parameters.Clear()
            Me.cmdDormStud.Parameters.AddWithValue("@admNo", Me.txtAdmNo.Text.Trim)
            Me.cmdDormStud.Parameters.AddWithValue("@dormName", Me.cboHouse.Text.Trim)
            Me.cmdDormStud.Parameters.AddWithValue("@userName", userName.Trim)
            Me.cmdDormStud.Parameters.AddWithValue("@queryType", 1)
            rec = Me.cmdDormStud.ExecuteNonQuery
            If rec > 0 Then
                MsgBox("Record Saved Successfully.", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "System Reply")
            End If
            clearControls()

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

    Private Sub txtAdmNo_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtAdmNo.TextChanged
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()

            Me.txtStudentName.Text = ""
            Me.chckBoxAutoAssign.Checked = False
            Me.cboHouse.Enabled = True
            Me.cboHouse.SelectedIndex = -1
            Me.cboHouse.DropDownStyle = ComboBoxStyle.DropDownList
            Me.txtGender.Text = ""
            Me.btnSave.Enabled = True
            Me.btnUpdate.Enabled = False
            Me.btnDelete.Enabled = False

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

    Private Sub btnUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdate.Click
        If Me.cboHouse.Text = "" Then
            MsgBox("Missing House Name.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.txtAdmNo.Text = "" Then
            MsgBox("Missing admission number.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.txtStudentName.Text = "" Then
            MsgBox("Missing student name.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.txtGender.Text = "" Then
            MsgBox("Missing Gender.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()

            Dim result As MsgBoxResult = MsgBox("Update Details?", MsgBoxStyle.Question + MsgBoxStyle.ApplicationModal + MsgBoxStyle.YesNo, "Confirm Transaction")
            If result = MsgBoxResult.No Then
                Exit Sub
            End If

            Me.cmdDormStud.Connection = conn
            Me.cmdDormStud.CommandType = CommandType.StoredProcedure
            Me.cmdDormStud.CommandText = "sprocAccDormStudDetails"
            Me.cmdDormStud.Parameters.Clear()
            Me.cmdDormStud.Parameters.AddWithValue("@admNo", Me.txtAdmNo.Text.Trim)
            Me.cmdDormStud.Parameters.AddWithValue("@dormName", Me.cboHouse.Text.Trim)
            Me.cmdDormStud.Parameters.AddWithValue("@userName", userName.Trim)
            Me.cmdDormStud.Parameters.AddWithValue("@queryType", 2)
            rec = Me.cmdDormStud.ExecuteNonQuery
            If rec > 0 Then
                MsgBox("Record Updated Successfully.", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "System Reply")
            End If
            clearControls()

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

    Private Sub CLOSEToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Me.Close()
    End Sub

    Private Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        If Me.cboHouse.Text = "" Then
            MsgBox("Missing House Name.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.txtAdmNo.Text = "" Then
            MsgBox("Missing admission number.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.txtStudentName.Text = "" Then
            MsgBox("Missing student name.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.txtGender.Text = "" Then
            MsgBox("Missing Gender.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()

            Dim result As MsgBoxResult = MsgBox("Delete Details?", MsgBoxStyle.Question + MsgBoxStyle.ApplicationModal + MsgBoxStyle.YesNo, "Confirm Transaction")
            If result = MsgBoxResult.No Then
                Exit Sub
            End If

            Me.cmdDormStud.Connection = conn
            Me.cmdDormStud.CommandType = CommandType.StoredProcedure
            Me.cmdDormStud.CommandText = "sprocAccDormStudDetails"
            Me.cmdDormStud.Parameters.Clear()
            Me.cmdDormStud.Parameters.AddWithValue("@admNo", Me.txtAdmNo.Text.Trim)
            Me.cmdDormStud.Parameters.AddWithValue("@dormName", Me.cboHouse.Text.Trim)
            Me.cmdDormStud.Parameters.AddWithValue("@userName", userName.Trim)
            Me.cmdDormStud.Parameters.AddWithValue("@queryType", 3)
            rec = Me.cmdDormStud.ExecuteNonQuery
            If rec > 0 Then
                MsgBox("Record Deleted Successfully.", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "System Reply")
            End If
            clearControls()

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

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub
End Class