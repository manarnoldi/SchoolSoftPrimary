Imports System.Data.SqlClient
Public Class frmLogin
    Dim cmdLogin As New SqlCommand
    Dim LoginConn As SqlConnection
    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub btnLogin_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLogin.Click
        If Me.txtPassword.Text = "" Then
            MsgBox("Missing Password", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.txtUserName.Text = "" Then
            MsgBox("Missing UserName", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboDomain.Text = "" Then
            MsgBox("Missing Domain", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If
        Try
            If Me.LoginConn.State = ConnectionState.Closed Then
                Me.LoginConn.Open()
            End If

            cmdLogin.CommandText = "SELECT  * FROM vwUsersDetails WHERE (userName=@userName) AND (password=@password) AND " &
                vbNewLine & " (domainName=@domainName) AND (userStatus=1) AND (staffStatus=1)"
            cmdLogin.CommandType = CommandType.Text
            cmdLogin.Connection = LoginConn
            cmdLogin.Parameters.Clear()
            cmdLogin.Parameters.AddWithValue("@userName", Me.txtUserName.Text.Trim)
            cmdLogin.Parameters.AddWithValue("@password", Me.txtPassword.Text.Trim)
            cmdLogin.Parameters.AddWithValue("@domainName", Me.cboDomain.Text.Trim)
            Dim reader As SqlDataReader
            reader = cmdLogin.ExecuteReader
            Dim userCount As Integer = 0
            If reader.HasRows Then
                While (reader.Read)
                    domainName = TryCast(reader.Item("DomainName"), String)
                    userName = TryCast(reader.Item("Username"), String)
                    fullName = TryCast(reader.Item("FullName"), String)
                    empNo = TryCast(reader.Item("empNo"), String)
                    userId = CType(Val(reader!Userid), Integer)
                    frmHome.Show()
                    frmHome.ToolStripLabel2.Text = " LOGGED IN AS: " & UCase(fullName) & Space(5) & "DOMAIN : " & UCase(domainName) & Space(5)
                    frmHome.ToolStripLabel3.Text = " DATE: " & dateToday
                    frmHome.ToolStripLabel1.Text = Space(10) & " TIME: " & TimeOfDay()
                    Me.Close()
                End While
            Else
                MsgBox("Unknown user", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Login Error")
                Me.txtPassword.Clear()
                Me.txtUserName.Focus()
            End If
            reader.Close()
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            If Err.Number = 5 Then
                frmHome.Close()
                frmSettings.ShowDialog()
                Me.Close()
            End If
        Finally
            If LoginConn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub
    Private Sub loadDomains()

        If LoginConn.State = ConnectionState.Closed Then
            LoginConn.Open()
        End If
        cmdLogin.CommandText = "SELECT DISTINCT domainName FROM tblDomains WHERE (status='True')"
        cmdLogin.CommandType = CommandType.Text
        cmdLogin.Connection = LoginConn
        cmdLogin.Parameters.Clear()
        Dim reader As SqlDataReader = cmdLogin.ExecuteReader
        If reader.HasRows Then
            Me.cboDomain.Items.Clear()
            While reader.Read
                Me.cboDomain.Items.Add(IIf(DBNull.Value.Equals(reader!domainName), "", reader!domainName))
            End While
        End If
        reader.Close()
        If LoginConn.State = ConnectionState.Open Then
            LoginConn.Close()
        End If
    End Sub

    Private Sub frmLogin_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.txtUserName.Select()
            My.Settings.Reload()
            'My.Settings.Reset()
            If Date.Today >= My.Settings.EndDate Then
                My.Settings.closed = True
            End If
            If My.Settings.closed = True Then
                MsgBox("Fatal Error Occured." & vbNewLine & "Expiry Date Reached!" &
                      vbNewLine & "Contact System Developer on 0724-924465.",
                      MsgBoxStyle.Critical + MsgBoxStyle.OkOnly + MsgBoxStyle.ApplicationModal, "Contact Developer")
                Me.Close()
                Exit Sub
            End If
            LoginConn = New SqlConnection("Server=" & My.Settings.serverName.Trim & ";User ID=" & My.Settings.UserName.Trim & ";Database=" & My.Settings.dbName.Trim & ";Password=" & My.Settings.PassWord.Trim)
            If LoginConn.State = ConnectionState.Closed Then
                LoginConn.Open()
                loadDomains()
            End If
            Me.txtUserName.Focus()
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Server Error")
            frmSettings.ShowDialog()
            Me.Close()
        End Try
    End Sub

    Private Sub txtUserName_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtUserName.LostFocus
        Me.txtUserName.Text = StrConv(Me.txtUserName.Text, VbStrConv.ProperCase)
    End Sub

    Private Sub pnlLogin_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles pnlLogin.Paint

    End Sub
End Class