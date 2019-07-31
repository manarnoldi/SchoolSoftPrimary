Imports System.Data.SqlClient
Public Class frmEditMyAccount
    Dim currentPass As String = Nothing
    Dim cmdEditAcc As New SqlCommand
    Dim reader As SqlDataReader
    Dim rec As Integer = 0
    Private Sub frmEditMyAccount_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            cmdEditAcc.CommandText = "SELECT * FROM  vwUsersDetails WHERE (userStatus='True') AND (userId=@userId)"
            cmdEditAcc.Connection = conn
            cmdEditAcc.CommandType = CommandType.Text
            cmdEditAcc.Parameters.Clear()
            cmdEditAcc.Parameters.AddWithValue("@userId", userId)
            reader = cmdEditAcc.ExecuteReader
            If reader.HasRows Then
                While reader.Read
                    Me.txtEmployeeName.Text = IIf(DBNull.Value.Equals(reader!FullName), "", reader!FullName)
                    Me.txtUserName.Text = IIf(DBNull.Value.Equals(reader!userName), "", reader!userName)
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

    Private Sub frmEditMyAccount_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint
        If Me.IsMdiChild Then
            Dim PnlLoc As New Point
            PnlLoc.X = CInt((Me.Width - Me.pnlEditMyAccount.Width) / 2)
            PnlLoc.Y = CInt((Me.Height - Me.pnlEditMyAccount.Height) / 2.5)
            Me.pnlEditMyAccount.Location = PnlLoc
        Else
            Me.pnlEditMyAccount.Dock = DockStyle.Fill
        End If
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub txtConfirmPassword_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtConfirmPassword.TextChanged
        If Me.txtConfirmPassword.Text = "" Then
            Exit Sub
        End If
        If Me.txtNewPassword.Text = "" Then
            MsgBox("Enter Password before confirming", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Missing Record")
            Me.txtConfirmPassword.Text = ""
            Exit Sub
        End If
    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            cmdEditAcc.CommandText = "SELECT * FROM  vwUsersDetails WHERE (userStatus='True') AND (userId=@userId)"
            cmdEditAcc.Connection = conn
            cmdEditAcc.CommandType = CommandType.Text
            cmdEditAcc.Parameters.Clear()
            cmdEditAcc.Parameters.AddWithValue("@userId", userId)
            reader = cmdEditAcc.ExecuteReader
            If reader.HasRows Then
                While reader.Read
                    currentPass = IIf(DBNull.Value.Equals(reader!passWord), "", reader!passWord)
                End While
            End If
            reader.Close()
            If Me.txtNewPassword.Text.Trim = "" Then
                MsgBox("Missing New password", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
                Exit Sub
            ElseIf Not (currentPass.Trim = Me.txtCurrentPassword.Text.Trim) Then
                MsgBox("Current password not correct", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
                Exit Sub
            ElseIf (Me.txtNewPassword.Text.Trim = Me.txtCurrentPassword.Text.Trim) Then
                MsgBox("Old and New password the same", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
                Exit Sub
            ElseIf Not (Me.txtNewPassword.Text.Trim = Me.txtConfirmPassword.Text.Trim) Then
                MsgBox("New And Old password dont match", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
                Exit Sub
            End If
            Dim result As MsgBoxResult = MsgBox("Save Record?", MsgBoxStyle.Question + MsgBoxStyle.ApplicationModal + MsgBoxStyle.YesNo, "Successful Transaction")
            If result = MsgBoxResult.No Then
                Exit Sub
            End If
            cmdEditAcc.CommandType = CommandType.StoredProcedure
            cmdEditAcc.Connection = conn
            cmdEditAcc.CommandText = "sprocEditAccount"
            cmdEditAcc.Parameters.Clear()
            cmdEditAcc.Parameters.AddWithValue("@userName", Me.txtUserName.Text.Trim)
            cmdEditAcc.Parameters.AddWithValue("@dateOfAction", Date.Now)
            cmdEditAcc.Parameters.AddWithValue("@doneBy", userName.Trim)
            cmdEditAcc.Parameters.AddWithValue("@userId", userId)
            cmdEditAcc.Parameters.AddWithValue("@newPass", Me.txtNewPassword.Text.Trim)
            rec = cmdEditAcc.ExecuteNonQuery
            If rec > 0 Then
                MsgBox("Record Saved", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Successful Transaction")
            End If
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
            Me.txtConfirmPassword.Text = ""
            Me.txtNewPassword.Text = ""
            Me.txtCurrentPassword.Text = ""
        End Try
    End Sub
End Class