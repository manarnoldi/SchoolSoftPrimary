Public Class frmSettings

    Private Sub frmSettings_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        frmLogin.Show()
    End Sub

    Private Sub frmSettings_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        My.Settings.Reload()
    End Sub

    Private Sub frmSettings_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.txtDatabaseName.Text = My.Settings.dbName
        Me.txtPassword.Text = My.Settings.PassWord
        Me.txtServerName.Text = My.Settings.serverName
        Me.txtUserName.Text = My.Settings.UserName
    End Sub


    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Me.Close()
    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        Try
            If Me.txtDatabaseName.Text.Trim.Length <= 0 Then
                MsgBox("Missing Database Name", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
                Exit Sub
            ElseIf Me.txtPassword.Text.Trim.Length <= 0 Then
                MsgBox("Missing Password", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
                Exit Sub
            ElseIf Me.txtServerName.Text.Trim.Length <= 0 Then
                MsgBox("Missing ServerName", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
                Exit Sub
            ElseIf Me.txtUserName.Text.Trim.Length <= 0 Then
                MsgBox("Missing UserName", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
                Exit Sub
            Else
                With My.Settings
                    .dbName = Me.txtDatabaseName.Text.Trim
                    .PassWord = Me.txtPassword.Text.Trim
                    .serverName = Me.txtServerName.Text.Trim
                    .UserName = Me.txtUserName.Text.Trim
                    .Save()
                    .Reload()
                    MsgBox("Settings updated successfully", MsgBoxStyle.Information + MsgBoxStyle.OkOnly + MsgBoxStyle.ApplicationModal, "SuccessFul Transaction")
                    Me.Close()
                    'frmLogin.Parent = frmHome
                    'frmLogin.Show()
                End With
            End If
        Catch ex As Exception
            MsgBox("Settings update failed", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly + MsgBoxStyle.ApplicationModal, "Error Detected")
        End Try
    End Sub
End Class