Imports System.Data.SqlClient
Public Class frmDomains
    Dim exists As Boolean = False
    Dim cmdModules As New SqlCommand
    Dim reader As SqlDataReader
    Dim rec As Integer = 0
    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Me.Close()
    End Sub
    Private Sub loadDomains()
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            Me.lstDomains.Items.Clear()
            dbconnection()
            cmdModules.CommandText = "SELECT domainId,domainName FROM tblDomains WHERE (status='True') ORDER BY domainName"
            cmdModules.CommandType = CommandType.Text
            cmdModules.Connection = conn
            cmdModules.Parameters.Clear()
            reader = cmdModules.ExecuteReader
            If reader.HasRows Then
                While reader.Read
                    li = Me.lstDomains.Items.Add(IIf(DBNull.Value.Equals(reader!domainName), "", reader!domainName))
                    li.Tag = IIf(DBNull.Value.Equals(reader!domainId), "", reader!domainId)
                End While
            End If
            reader.Close()
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly + MsgBoxStyle.ApplicationModal, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub
    Private Sub frmdomains_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        loadDomains()
        Me.btnDelete.Enabled = False
        Me.btnSave.Enabled = True
        Me.btnUpdate.Enabled = False
    End Sub

    Private Sub frmDomains_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint
        If Me.IsMdiChild Then
            Dim PnlLoc As New Point
            PnlLoc.X = CInt((Me.Width - Me.pnlDomains.Width) / 2)
            PnlLoc.Y = CInt((Me.Height - Me.pnlDomains.Height) / 2.5)
            Me.pnlDomains.Location = PnlLoc
        Else
            Me.pnlDomains.Dock = DockStyle.Fill
        End If
    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        If Me.txtDomName.Text.Trim.Length <= 0 Then
            MsgBox("Missing Domain Name", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly + MsgBoxStyle.ApplicationModal, "Error Detected")
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            checkExistence()
            If exists = True Then
                Exit Sub
            End If
            cmdModules.Connection = conn
            cmdModules.CommandType = CommandType.StoredProcedure
            cmdModules.CommandText = "sprocDomainsInsert"
            cmdModules.Parameters.Clear()
            cmdModules.Parameters.AddWithValue("@domainName", Me.txtDomName.Text.Trim)
            cmdModules.Parameters.AddWithValue("@userName", userName.Trim)
            cmdModules.Parameters.AddWithValue("@dateOfreg", Date.Now)
            rec = cmdModules.ExecuteNonQuery()
            If rec > 0 Then
                MsgBox("Record Saved", MsgBoxStyle.Information + MsgBoxStyle.OkOnly + MsgBoxStyle.ApplicationModal, "SuccessFul Transaction")
                loadDomains()
            End If
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly + MsgBoxStyle.ApplicationModal, "Error Detected")
        Finally
            Me.txtDomName.Tag = Nothing
            Me.txtDomName.Text = ""
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub

    Private Sub btnView_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnView.Click
        loadDomains()
        Me.btnDelete.Enabled = False
        Me.btnSave.Enabled = True
        Me.btnUpdate.Enabled = False
        Me.txtDomName.Text = ""
        Me.txtDomName.Tag = Nothing
    End Sub
    Private Sub checkExistence()
        cmdModules.Connection = conn
        cmdModules.CommandType = CommandType.Text
        cmdModules.CommandText = "SELECT * FROM tblDomains WHERE (status='True') and (domainName=@domainName)"
        cmdModules.Parameters.Clear()
        cmdModules.Parameters.AddWithValue("@domainName", Me.txtDomName.Text.Trim)
        reader = cmdModules.ExecuteReader
        If reader.HasRows Then
            MsgBox("Domain Name Exists", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly + MsgBoxStyle.ApplicationModal, "Error Detected")
            exists = True
        Else
            exists = False
        End If
        reader.Close()
    End Sub
    Private Sub btnUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdate.Click
        If Me.txtDomName.Tag = Nothing Or Me.txtDomName.Text.Trim.Length <= 0 Then
            MsgBox("Missing Domain Name to Edit." & vbNewLine & "Select From List", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly + MsgBoxStyle.ApplicationModal, "Error Detected")
            Exit Sub
        ElseIf Me.lstDomains.Items.Count <= 0 Then
            MsgBox("No record in the list to edit", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly + MsgBoxStyle.ApplicationModal, "Error Detected")
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            checkExistence()
            If exists = True Then
                Exit Sub
            End If
            Dim result As MsgBoxResult = MsgBox("Update Record?", MsgBoxStyle.Question + MsgBoxStyle.YesNo + MsgBoxStyle.ApplicationModal, "Confirm Delete")
            If result = MsgBoxResult.No Then
                Exit Sub
            End If
            cmdModules.Connection = conn
            cmdModules.CommandType = CommandType.StoredProcedure
            cmdModules.CommandText = "sprocDomainsUpdate"
            cmdModules.Parameters.Clear()
            cmdModules.Parameters.AddWithValue("@domainName", Me.txtDomName.Text.Trim)
            cmdModules.Parameters.AddWithValue("@userName", userName.Trim)
            cmdModules.Parameters.AddWithValue("@dateOfreg", Date.Now)
            cmdModules.Parameters.AddWithValue("@domainId", Me.txtDomName.Tag)
            rec = cmdModules.ExecuteNonQuery
            If rec > 0 Then
                MsgBox("Record Updated", MsgBoxStyle.Information + MsgBoxStyle.OkOnly + MsgBoxStyle.ApplicationModal, "Successfull Transaction")
                loadDomains()
            End If
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Information + MsgBoxStyle.OkOnly + MsgBoxStyle.ApplicationModal, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
            Me.txtDomName.Tag = Nothing
            Me.txtDomName.Text = ""
            Me.btnDelete.Enabled = False
            Me.btnSave.Enabled = True
            Me.btnUpdate.Enabled = False
        End Try
    End Sub

    Private Sub lstdomains_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lstDomains.Click
        If Me.lstDomains.SelectedItems.Count = 1 Then
            Me.txtDomName.Text = Me.lstDomains.SelectedItems(0).Text
            Me.txtDomName.Tag = Me.lstDomains.SelectedItems(0).Tag
            Me.btnDelete.Enabled = True
            Me.btnSave.Enabled = False
            Me.btnUpdate.Enabled = True
        End If
    End Sub

    Private Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        If Me.txtDomName.Tag = Nothing Or Me.txtDomName.Text.Trim.Length <= 0 Then
            MsgBox("Missing Domain Name to Edit." & vbNewLine & "Select From List", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly + MsgBoxStyle.ApplicationModal, "Error Detected")
            Exit Sub
        ElseIf Me.lstDomains.Items.Count <= 0 Then
            MsgBox("No record in the list to edit", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly + MsgBoxStyle.ApplicationModal, "Error Detected")
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            Dim result As MsgBoxResult = MsgBox("Delete Record?", MsgBoxStyle.Question + MsgBoxStyle.YesNo + MsgBoxStyle.ApplicationModal, "Confirm Delete")
            If result = MsgBoxResult.No Then
                Exit Sub
            End If
            cmdModules.Connection = conn
            cmdModules.CommandType = CommandType.StoredProcedure
            cmdModules.CommandText = "sprocDomainsDelete"
            cmdModules.Parameters.Clear()
            cmdModules.Parameters.AddWithValue("@domainName", Me.txtDomName.Text.Trim)
            cmdModules.Parameters.AddWithValue("@userName", userName.Trim)
            cmdModules.Parameters.AddWithValue("@dateOfreg", Date.Now)
            cmdModules.Parameters.AddWithValue("domainId", Me.txtDomName.Tag)
            rec = cmdModules.ExecuteNonQuery
            If rec > 0 Then
                MsgBox("Record Deleted", MsgBoxStyle.Information + MsgBoxStyle.OkOnly + MsgBoxStyle.ApplicationModal, "Successfull Transaction")
                loadDomains()
            End If
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Information + MsgBoxStyle.OkOnly + MsgBoxStyle.ApplicationModal, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
            Me.txtDomName.Tag = Nothing
            Me.txtDomName.Text = ""
            Me.btnDelete.Enabled = False
            Me.btnSave.Enabled = True
            Me.btnUpdate.Enabled = False
        End Try
    End Sub
End Class