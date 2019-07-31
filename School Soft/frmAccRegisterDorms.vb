Imports System.Data.SqlClient
Public Class frmAccRegisterDorms
    Dim rec As Integer = 0
    Dim reader As SqlDataReader
    Dim cmdRegDorms As New SqlCommand
    Private Sub frmRegisterDorms_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub
    Private Sub loadlist()
        Me.lstDormRegister.Items.Clear()

        Me.cmdRegDorms.Connection = conn
        Me.cmdRegDorms.CommandType = CommandType.Text
        Me.cmdRegDorms.CommandText = "SELECT * FROM tblAccdormRegister ORDER BY dormName"
        Me.cmdRegDorms.Parameters.Clear()
        reader = Me.cmdRegDorms.ExecuteReader
        While reader.Read
            li = Me.lstDormRegister.Items.Add(IIf(DBNull.Value.Equals(reader!dormName), "", reader!dormName))
            li.SubItems.Add(IIf(DBNull.Value.Equals(reader!dormCapacity), "", reader!dormCapacity))
            li.SubItems.Add(IIf(DBNull.Value.Equals(reader!genderType), "", reader!genderType))
            li.Tag = IIf(DBNull.Value.Equals(reader!dormDescription), "", reader!dormDescription)
        End While
        reader.Close()
    End Sub
    Private Sub frmRegisterDorms_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint
        If Me.IsMdiChild Then
            Dim PnlLoc As New Point
            PnlLoc.X = CInt((Me.Width - Me.pnlAccDormReg.Width) / 2)
            PnlLoc.Y = CInt((Me.Height - Me.pnlAccDormReg.Height) / 2.5)
            Me.pnlAccDormReg.Location = PnlLoc
        Else
            Me.pnlAccDormReg.Dock = DockStyle.Fill
        End If
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub btnView_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnView.Click
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            loadlist()
            Me.txtDormCapacity.Text = ""
            Me.txtDormDescription.Text = ""
            Me.txtDormName.Text = ""
            Me.cboGender.SelectedIndex = -1

            Me.txtDormName.Enabled = True
            Me.btnUpdate.Enabled = False
            Me.btnSave.Enabled = True
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        If Me.txtDormName.Text.Trim.Length <= 0 Then
            MsgBox("Dorm Name Missing.", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "System Reply")
            Exit Sub
        ElseIf Me.txtDormCapacity.Text.Trim.Length <= 0 Then
            MsgBox("Dorm Capacity Missing.", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "System Reply")
            Exit Sub
        ElseIf Me.cboGender.Text.Trim.Length <= 0 Then
            MsgBox("Dormitory Gender Missing.", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "System Reply")
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()

            Dim result As MsgBoxResult = MsgBox("Save Record?", MsgBoxStyle.Question + MsgBoxStyle.ApplicationModal + MsgBoxStyle.YesNo, "Confirm Transaction")
            If result = MsgBoxResult.No Then
                Exit Sub
            End If

            Me.cmdRegDorms.Connection = conn
            Me.cmdRegDorms.CommandType = CommandType.StoredProcedure
            Me.cmdRegDorms.CommandText = "sprocAccDormRegister"
            Me.cmdRegDorms.Parameters.Clear()
            Me.cmdRegDorms.Parameters.AddWithValue("@dormName", Me.txtDormName.Text.Trim)
            Me.cmdRegDorms.Parameters.AddWithValue("@dormDescription", Me.txtDormDescription.Text.Trim)
            Me.cmdRegDorms.Parameters.AddWithValue("@dormCapacity", Me.txtDormCapacity.Text.Trim)
            Me.cmdRegDorms.Parameters.AddWithValue("@regBy", userName.Trim)
            Me.cmdRegDorms.Parameters.AddWithValue("@genderType", Me.cboGender.Text.Trim)
            Me.cmdRegDorms.Parameters.AddWithValue("@queryType", 1)
            rec = Me.cmdRegDorms.ExecuteNonQuery
            If rec > 0 Then
                MsgBox("Record Saved Successfully.", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "System Reply")
            End If
            Me.lstDormRegister.Items.Clear()

            Me.txtDormCapacity.Text = ""
            Me.txtDormDescription.Text = ""
            Me.txtDormName.Text = ""
            Me.cboGender.SelectedIndex = -1

            Me.txtDormName.Enabled = True
            Me.btnUpdate.Enabled = False
            Me.btnSave.Enabled = True
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub

    Private Sub CLOSEToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CLOSEToolStripMenuItem.Click
        Me.Close()
    End Sub

    Private Sub UPDATEToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UPDATEToolStripMenuItem.Click
        If Me.lstDormRegister.Items.Count <= 0 Then
            MsgBox("View items and select to edit.", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "System Reply")
            Exit Sub
        ElseIf Me.lstDormRegister.CheckedItems.Count <= 0 Then
            MsgBox("Check item to Edit.", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "System Reply")
            Exit Sub
        ElseIf Me.lstDormRegister.CheckedItems.Count > 1 Then
            MsgBox("Edit One Item at A time.", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "System Reply")
            Exit Sub
        End If
        Me.txtDormCapacity.Text = ""
        Me.txtDormDescription.Text = ""
        Me.txtDormName.Text = ""
        Me.cboGender.SelectedIndex = -1

        Me.txtDormCapacity.Text = Me.lstDormRegister.CheckedItems(0).SubItems(1).Text.Trim
        Me.txtDormDescription.Text = Me.lstDormRegister.CheckedItems(0).Tag
        Me.txtDormName.Text = Me.lstDormRegister.CheckedItems(0).Text.Trim
        Me.cboGender.Text = Me.lstDormRegister.CheckedItems(0).SubItems(2).Text.Trim

        Me.txtDormName.Enabled = False
        Me.btnUpdate.Enabled = True
        Me.btnSave.Enabled = False
    End Sub

    Private Sub btnUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdate.Click
        If Me.txtDormName.Text.Trim.Length <= 0 Then
            MsgBox("Dorm Name Missing.", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "System Reply")
            Exit Sub
        ElseIf Me.txtDormCapacity.Text.Trim.Length <= 0 Then
            MsgBox("Dorm Capacity Missing.", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "System Reply")
            Exit Sub
        ElseIf Me.cboGender.Text.Trim.Length <= 0 Then
            MsgBox("Dormitory Gender Missing.", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "System Reply")
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()

            Dim result As MsgBoxResult = MsgBox("Update Record?", MsgBoxStyle.Question + MsgBoxStyle.ApplicationModal + MsgBoxStyle.YesNo, "Confirm Transaction")
            If result = MsgBoxResult.No Then
                Exit Sub
            End If

            Me.cmdRegDorms.Connection = conn
            Me.cmdRegDorms.CommandType = CommandType.StoredProcedure
            Me.cmdRegDorms.CommandText = "sprocAccDormRegister"
            Me.cmdRegDorms.Parameters.Clear()
            Me.cmdRegDorms.Parameters.AddWithValue("@dormName", Me.txtDormName.Text.Trim)
            Me.cmdRegDorms.Parameters.AddWithValue("@dormDescription", Me.txtDormDescription.Text.Trim)
            Me.cmdRegDorms.Parameters.AddWithValue("@dormCapacity", Me.txtDormCapacity.Text.Trim)
            Me.cmdRegDorms.Parameters.AddWithValue("@regBy", userName.Trim)
            Me.cmdRegDorms.Parameters.AddWithValue("@genderType", Me.cboGender.Text.Trim)
            Me.cmdRegDorms.Parameters.AddWithValue("@queryType", 2)
            rec = Me.cmdRegDorms.ExecuteNonQuery
            If rec > 0 Then
                MsgBox("Record Updated Successfully.", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "System Reply")
            End If
            Me.lstDormRegister.Items.Clear()

            Me.txtDormCapacity.Text = ""
            Me.txtDormDescription.Text = ""
            Me.txtDormName.Text = ""

            Me.cboGender.SelectedIndex = -1

            Me.txtDormName.Enabled = True
            Me.btnUpdate.Enabled = False
            Me.btnSave.Enabled = True
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub

    Private Sub DELETEToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DELETEToolStripMenuItem.Click
         If Me.lstDormRegister.Items.Count <= 0 Then
            MsgBox("View items and select to edit.", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "System Reply")
            Exit Sub
        ElseIf Me.lstDormRegister.CheckedItems.Count <= 0 Then
            MsgBox("Check item to Edit.", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "System Reply")
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()

            Dim result As MsgBoxResult = MsgBox("Delete Record/s?", MsgBoxStyle.Question + MsgBoxStyle.ApplicationModal + MsgBoxStyle.YesNo, "Confirm Transaction")
            If result = MsgBoxResult.No Then
                Exit Sub
            End If
            i = 0
            For i = 0 To Me.lstDormRegister.CheckedItems.Count - 1
                Me.cmdRegDorms.Connection = conn
                Me.cmdRegDorms.CommandType = CommandType.StoredProcedure
                Me.cmdRegDorms.CommandText = "sprocAccDormRegister"
                Me.cmdRegDorms.Parameters.Clear()
                Me.cmdRegDorms.Parameters.AddWithValue("@dormName", Me.lstDormRegister.CheckedItems(i).Text.Trim)
                Me.cmdRegDorms.Parameters.AddWithValue("@dormDescription", Me.lstDormRegister.CheckedItems(i).Tag)
                Me.cmdRegDorms.Parameters.AddWithValue("@dormCapacity", Me.lstDormRegister.CheckedItems(i).SubItems(1).Text.Trim)
                Me.cmdRegDorms.Parameters.AddWithValue("@regBy", userName.Trim)
                Me.cmdRegDorms.Parameters.AddWithValue("@genderType", Me.cboGender.Text.Trim)
                Me.cmdRegDorms.Parameters.AddWithValue("@queryType", 3)
                rec = rec + Me.cmdRegDorms.ExecuteNonQuery
            Next
            If rec > 0 Then
                MsgBox("Record/s Deleted Successfully.", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "System Reply")
            End If
            Me.lstDormRegister.Items.Clear()

            Me.txtDormCapacity.Text = ""
            Me.txtDormDescription.Text = ""
            Me.txtDormName.Text = ""

            Me.cboGender.SelectedIndex = -1

            Me.txtDormName.Enabled = True
            Me.btnUpdate.Enabled = False
            Me.btnSave.Enabled = True
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub
End Class