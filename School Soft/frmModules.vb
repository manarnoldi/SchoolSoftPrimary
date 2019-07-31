Imports System.Data.SqlClient
Public Class frmModules
    Dim exists As Boolean = False
    Dim cmdModules As New SqlCommand
    Dim reader As SqlDataReader
    Dim rec As Integer = 0
    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Me.Close()
    End Sub
    Private Sub loadModules()
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            Me.lstModules.Items.Clear()
            dbconnection()
            cmdModules.CommandText = "SELECT modId,modName FROM tblModules WHERE (status='True') ORDER BY modName"
            cmdModules.CommandType = CommandType.Text
            cmdModules.Connection = conn
            cmdModules.Parameters.Clear()
            reader = cmdModules.ExecuteReader
            If reader.HasRows Then
                While reader.Read
                    li = Me.lstModules.Items.Add(IIf(DBNull.Value.Equals(reader!modName), "", reader!modName))
                    li.Tag = IIf(DBNull.Value.Equals(reader!modId), "", reader!modId)
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
    Private Sub frmModules_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        loadModules()
        Me.btnDelete.Enabled = False
        Me.btnSave.Enabled = True
        Me.btnUpdate.Enabled = False
    End Sub

    Private Sub frmModules_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint
        If Me.IsMdiChild Then
            Dim PnlLoc As New Point
            PnlLoc.X = CInt((Me.Width - Me.pnlModules.Width) / 2)
            PnlLoc.Y = CInt((Me.Height - Me.pnlModules.Height) / 2.5)
            Me.pnlModules.Location = PnlLoc
        Else
            Me.pnlModules.Dock = DockStyle.Fill
        End If
    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        If Me.txtModName.Text.Trim.Length <= 0 Then
            MsgBox("Missing Module Name", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly + MsgBoxStyle.ApplicationModal, "Error Detected")
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
            cmdModules.CommandText = "sprocModulesInsert"
            cmdModules.Parameters.Clear()
            cmdModules.Parameters.AddWithValue("@modName", Me.txtModName.Text.Trim)
            cmdModules.Parameters.AddWithValue("@userName", userName.Trim)
            cmdModules.Parameters.AddWithValue("@dateOfreg", Date.Now)
            rec = cmdModules.ExecuteNonQuery()
            If rec > 0 Then
                MsgBox("Record Saved", MsgBoxStyle.Information + MsgBoxStyle.OkOnly + MsgBoxStyle.ApplicationModal, "SuccessFul Transaction")
                loadModules()
            End If
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly + MsgBoxStyle.ApplicationModal, "Error Detected")
        Finally
            Me.txtModName.Tag = Nothing
            Me.txtModName.Text = ""
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub

    Private Sub btnView_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnView.Click
        loadModules()
        Me.btnDelete.Enabled = False
        Me.btnSave.Enabled = True
        Me.btnUpdate.Enabled = False
        Me.txtModName.Text = ""
        Me.txtModName.Tag = Nothing
    End Sub
    Private Sub checkExistence()
        cmdModules.Connection = conn
        cmdModules.CommandType = CommandType.Text
        cmdModules.CommandText = "SELECT * FROM tblModules WHERE (status='True') and (modName=@modName)"
        cmdModules.Parameters.Clear()
        cmdModules.Parameters.AddWithValue("@modName", Me.txtModName.Text.Trim)
        reader = cmdModules.ExecuteReader
        If reader.HasRows Then
            MsgBox("Module Name Exists", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly + MsgBoxStyle.ApplicationModal, "Error Detected")
            exists = True
        Else
            exists = False
        End If
        reader.Close()
    End Sub
    Private Sub btnUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdate.Click
        If Me.txtModName.Tag = Nothing Or Me.txtModName.Text.Trim.Length <= 0 Then
            MsgBox("Missing Module Name to Edit." & vbNewLine & "Select From List", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly + MsgBoxStyle.ApplicationModal, "Error Detected")
            Exit Sub
        ElseIf Me.lstModules.Items.Count <= 0 Then
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
            cmdModules.CommandText = "sprocModulesUpdate"
            cmdModules.Parameters.Clear()
            cmdModules.Parameters.AddWithValue("@modName", Me.txtModName.Text.Trim)
            cmdModules.Parameters.AddWithValue("@userName", userName.Trim)
            cmdModules.Parameters.AddWithValue("@dateOfreg", Date.Now)
            cmdModules.Parameters.AddWithValue("@modId", Me.txtModName.Tag)
            rec = cmdModules.ExecuteNonQuery
            If rec > 0 Then
                MsgBox("Record Updated", MsgBoxStyle.Information + MsgBoxStyle.OkOnly + MsgBoxStyle.ApplicationModal, "Successfull Transaction")
                loadModules()
            End If
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Information + MsgBoxStyle.OkOnly + MsgBoxStyle.ApplicationModal, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
            Me.txtModName.Tag = Nothing
            Me.txtModName.Text = ""
            Me.btnDelete.Enabled = False
            Me.btnSave.Enabled = True
            Me.btnUpdate.Enabled = False
        End Try
    End Sub

    Private Sub lstModules_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lstModules.Click
        If Me.lstModules.SelectedItems.Count = 1 Then
            Me.txtModName.Text = Me.lstModules.SelectedItems(0).Text
            Me.txtModName.Tag = Me.lstModules.SelectedItems(0).Tag
            Me.btnDelete.Enabled = True
            Me.btnSave.Enabled = False
            Me.btnUpdate.Enabled = True
        End If
    End Sub

    Private Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        If Me.txtModName.Tag = Nothing Or Me.txtModName.Text.Trim.Length <= 0 Then
            MsgBox("Missing Module Name to Edit." & vbNewLine & "Select From List", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly + MsgBoxStyle.ApplicationModal, "Error Detected")
            Exit Sub
        ElseIf Me.lstModules.Items.Count <= 0 Then
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
            cmdModules.CommandText = "sprocModulesDelete"
            cmdModules.Parameters.Clear()
            cmdModules.Parameters.AddWithValue("@modName", Me.txtModName.Text.Trim)
            cmdModules.Parameters.AddWithValue("@userName", userName.Trim)
            cmdModules.Parameters.AddWithValue("@dateOfreg", Date.Now)
            cmdModules.Parameters.AddWithValue("@modId", Me.txtModName.Tag)
            rec = cmdModules.ExecuteNonQuery
            If rec > 0 Then
                MsgBox("Record Deleted", MsgBoxStyle.Information + MsgBoxStyle.OkOnly + MsgBoxStyle.ApplicationModal, "Successfull Transaction")
                loadModules()
            End If
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Information + MsgBoxStyle.OkOnly + MsgBoxStyle.ApplicationModal, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
            Me.txtModName.Tag = Nothing
            Me.txtModName.Text = ""
            Me.btnDelete.Enabled = False
            Me.btnSave.Enabled = True
            Me.btnUpdate.Enabled = False
        End Try
    End Sub
End Class