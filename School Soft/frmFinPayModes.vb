Imports System.Data.SqlClient
Public Class frmFinPayModes
    Dim cmdPayModes As New SqlCommand
    Dim reader As SqlDataReader
    Dim rec As Integer = 0
    Private Sub frmFinPaymentModes_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
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
    Private Sub loadList()
        Me.lstPayModes.Items.Clear()

        Me.cmdPayModes.Connection = conn
        Me.cmdPayModes.CommandType = CommandType.Text
        Me.cmdPayModes.CommandText = "SELECT * FROM tblFinPayModes ORDER BY modeName"
        Me.cmdPayModes.Parameters.Clear()
        reader = Me.cmdPayModes.ExecuteReader
        While reader.Read
            li = Me.lstPayModes.Items.Add(IIf(DBNull.Value.Equals(reader!modeName), "", reader!modeName))
            li.Tag = IIf(DBNull.Value.Equals(reader!modeId), "", reader!modeId)
        End While
        reader.Close()
    End Sub
    Private Sub loadCombos()
        Me.cboPayMode.Items.Clear()
        Me.cboPayMode.Text = ""
        Me.cboPayMode.SelectedIndex = -1

        Me.cmdPayModes.Connection = conn
        Me.cmdPayModes.CommandType = CommandType.Text
        Me.cmdPayModes.CommandText = "SELECT DISTINCT modeName FROM tblFinPayModes ORDER BY modeName"
        Me.cmdPayModes.Parameters.Clear()
        reader = Me.cmdPayModes.ExecuteReader
        While reader.Read
            Me.cboPayMode.Items.Add(IIf(DBNull.Value.Equals(reader!modeName), "", reader!modeName))
        End While
        reader.Close()
    End Sub
    Private Sub frmFinPaymentModes_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint
        If Me.IsMdiChild Then
            Dim PnlLoc As New Point
            PnlLoc.X = CInt((Me.Width - Me.pnlPayModes.Width) / 2)
            PnlLoc.Y = CInt((Me.Height - Me.pnlPayModes.Height) / 2.5)
            Me.pnlPayModes.Location = PnlLoc
        Else
            Me.pnlPayModes.Dock = DockStyle.Fill
        End If
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub CLOSEToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CLOSEToolStripMenuItem.Click
        Me.Close()
    End Sub

    Private Sub btnView_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnView.Click
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            loadCombos()
            loadList()
            Me.btnSave.Enabled = True
            Me.btnUpdate.Enabled = False
            Me.cboPayMode.Tag = ""
            Me.lstPayModes.Tag = ""
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        If Me.cboPayMode.Text.Trim.Length <= 0 Then
            MsgBox("Mode Name Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If
        
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()

            Dim recordExists As Boolean = checkIfRecordExists(Me.cboPayMode.Text.Trim)
            If recordExists = True Then
                MsgBox("Duplicate Found in the database!", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
                Exit Sub
            End If

            Dim result As MsgBoxResult = MsgBox("Save Record?", MsgBoxStyle.Question + MsgBoxStyle.ApplicationModal + MsgBoxStyle.YesNo, "Confirm Transaction")
            If result = MsgBoxResult.No Then
                Exit Sub
            End If

            Me.cmdPayModes.Connection = conn
            Me.cmdPayModes.CommandType = CommandType.StoredProcedure
            Me.cmdPayModes.CommandText = "sprocFinPayModes"
            Me.cmdPayModes.Parameters.Clear()
            Me.cmdPayModes.Parameters.AddWithValue("@modeName", Me.cboPayMode.Text.Trim)
            Me.cmdPayModes.Parameters.AddWithValue("@regBy", userName.Trim)
            Me.cmdPayModes.Parameters.AddWithValue("@queryType", "INSERT")
            rec = Me.cmdPayModes.ExecuteNonQuery
            If rec > 0 Then
                MsgBox("Record Saved Successfully.", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "System Reply")
            End If
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

    Private Sub UPDATEToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UPDATEToolStripMenuItem.Click
        If Me.lstPayModes.Items.Count <= 0 Then
            MsgBox("Missing Items In The List.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.lstPayModes.CheckedItems.Count <= 0 Then
            MsgBox("Missing Checked Items In The List.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.lstPayModes.CheckedItems.Count > 1 Then
            MsgBox("Edit one item At A Time.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            Me.cboPayMode.Tag = ""
            Me.lstPayModes.Tag = ""
            Me.cboPayMode.Text = Me.lstPayModes.CheckedItems(0).Text.Trim
            Me.cboPayMode.Tag = Me.lstPayModes.CheckedItems(0).Tag
            Me.lstPayModes.Tag = Me.lstPayModes.CheckedItems(0).Text.Trim
            Me.lstPayModes.Items.Clear()
            Me.btnUpdate.Enabled = True
            Me.btnSave.Enabled = False
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub
    Private Function checkIfRecordExists(ByVal modeName As String)
        Dim recordExists As Boolean = False
        Me.cmdPayModes.Connection = conn
        Me.cmdPayModes.CommandType = CommandType.Text
        Me.cmdPayModes.CommandText = "SELECT * FROM tblFinPayModes WHERE (modeName=@modeName)"
        Me.cmdPayModes.Parameters.Clear()
        Me.cmdPayModes.Parameters.AddWithValue("@modeName", modeName.Trim)
        reader = Me.cmdPayModes.ExecuteReader
        If reader.HasRows = True Then
            recordExists = True
        ElseIf reader.HasRows = False Then
            recordExists = False
        End If
        reader.Close()
        Return recordExists
    End Function

    Private Sub btnUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdate.Click
        If Me.cboPayMode.Text.Trim.Length <= 0 Then
            MsgBox("Mode Name Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()

            Dim recordExists As Boolean = checkIfRecordExists(Me.cboPayMode.Text.Trim)
            If recordExists = True Then
                MsgBox("Duplicate Found in the database!", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
                Exit Sub
            End If

            Dim result As MsgBoxResult = MsgBox("Update Record?", MsgBoxStyle.Question + MsgBoxStyle.ApplicationModal + MsgBoxStyle.YesNo, "Confirm Transaction")
            If result = MsgBoxResult.No Then
                Exit Sub
            End If

            Me.cmdPayModes.Connection = conn
            Me.cmdPayModes.CommandType = CommandType.StoredProcedure
            Me.cmdPayModes.CommandText = "sprocFinPayModes"
            Me.cmdPayModes.Parameters.Clear()
            Me.cmdPayModes.Parameters.AddWithValue("@modeName", Me.cboPayMode.Text.Trim)
            Me.cmdPayModes.Parameters.AddWithValue("@regBy", userName.Trim)
            Me.cmdPayModes.Parameters.AddWithValue("@queryType", "UPDATE")
            Me.cmdPayModes.Parameters.AddWithValue("@prevModeName", Me.lstPayModes.Tag)
            Me.cmdPayModes.Parameters.AddWithValue("@modeId", Me.cboPayMode.Tag)
            rec = Me.cmdPayModes.ExecuteNonQuery
            If rec > 0 Then
                MsgBox("Record Updated Successfully.", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "System Reply")
            End If
            loadCombos()
            loadList()
            Me.cboPayMode.Tag = ""
            Me.lstPayModes.Tag = ""
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
        If Me.lstPayModes.Items.Count <= 0 Then
            MsgBox("Missing Items In The List.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.lstPayModes.CheckedItems.Count <= 0 Then
            MsgBox("Missing Checked Items In The List.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
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

            For i = 0 To Me.lstPayModes.CheckedItems.Count - 1
                Me.cmdPayModes.Connection = conn
                Me.cmdPayModes.CommandType = CommandType.StoredProcedure
                Me.cmdPayModes.CommandText = "sprocFinPayModes"
                Me.cmdPayModes.Parameters.Clear()
                Me.cmdPayModes.Parameters.AddWithValue("@modeName", Me.lstPayModes.CheckedItems(i).Text.Trim)
                Me.cmdPayModes.Parameters.AddWithValue("@prevModeName", Me.lstPayModes.CheckedItems(i).Text.Trim)
                Me.cmdPayModes.Parameters.AddWithValue("@regBy", userName.Trim)
                Me.cmdPayModes.Parameters.AddWithValue("@queryType", "DELETE")
                Me.cmdPayModes.Parameters.AddWithValue("@modeId", Me.lstPayModes.CheckedItems(i).Tag)
                rec = Me.cmdPayModes.ExecuteNonQuery
            Next
            If rec > 0 Then
                MsgBox("Record/s Deleted Successfully.", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "System Reply")
            End If
            loadCombos()
            loadList()
            Me.cboPayMode.Tag = ""
            Me.lstPayModes.Tag = ""
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub
End Class