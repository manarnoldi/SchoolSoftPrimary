Imports System.Data.SqlClient
Public Class frmFinFeeCategory
    Dim cmdFeeCat As New SqlCommand
    Dim reader As SqlDataReader
    Dim rec As Integer = 0

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub
    Private Sub loadCombos()
        Me.cboFeeCategoryName.Items.Clear()
        Me.cboFeeCategoryName.Text = ""
        Me.cboFeeCategoryName.SelectedIndex = -1

        Me.cmdFeeCat.Connection = conn
        Me.cmdFeeCat.CommandType = CommandType.Text
        Me.cmdFeeCat.CommandText = "SELECT feeCatName FROM tblFinFeeCategory ORDER BY feeCatName"
        Me.cmdFeeCat.Parameters.Clear()
        reader = Me.cmdFeeCat.ExecuteReader
        While reader.Read
            Me.cboFeeCategoryName.Items.Add(IIf(DBNull.Value.Equals(reader!feeCatName), "", reader!feeCatName))
        End While
        reader.Close()
    End Sub
    Private Sub loadList()
        Me.lstFeeCategory.Items.Clear()

        Me.cmdFeeCat.Connection = conn
        Me.cmdFeeCat.CommandType = CommandType.Text
        Me.cmdFeeCat.CommandText = "SELECT * FROM tblFinFeeCategory ORDER BY feeCatName"
        Me.cmdFeeCat.Parameters.Clear()
        reader = Me.cmdFeeCat.ExecuteReader
        While reader.Read
            li = Me.lstFeeCategory.Items.Add(IIf(DBNull.Value.Equals(reader!feeCatName), "", reader!feeCatName))
            li.Tag = IIf(DBNull.Value.Equals(reader!feeCatId), "", reader!feeCatId)
        End While
        reader.Close()
    End Sub
    Private Sub frmFinFeeCategory_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
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

    Private Sub frmFinFeeCategory_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint
        If Me.IsMdiChild Then
            Dim PnlLoc As New Point
            PnlLoc.X = CInt((Me.Width - Me.pnlFeeCat.Width) / 2)
            PnlLoc.Y = CInt((Me.Height - Me.pnlFeeCat.Height) / 2.5)
            Me.pnlFeeCat.Location = PnlLoc
        Else
            Me.pnlFeeCat.Dock = DockStyle.Fill
        End If
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
            Me.cboFeeCategoryName.Tag = ""
            Me.lstFeeCategory.Tag = ""
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        If Me.cboFeeCategoryName.Text.Trim.Length <= 0 Then
            MsgBox("Fee Category Is Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()

            Dim recordExists As Boolean = checkIfRecordExists(Me.cboFeeCategoryName.Text.Trim)
            If recordExists = True Then
                MsgBox("Duplicate Found in the database!", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
                Exit Sub
            End If

            Dim result As MsgBoxResult = MsgBox("Save Record?", MsgBoxStyle.Question + MsgBoxStyle.ApplicationModal + MsgBoxStyle.YesNo, "Confirm Transaction")
            If result = MsgBoxResult.No Then
                Exit Sub
            End If

            Me.cmdFeeCat.Connection = conn
            Me.cmdFeeCat.CommandType = CommandType.StoredProcedure
            Me.cmdFeeCat.CommandText = "sprocFinFeeCategory"
            Me.cmdFeeCat.Parameters.Clear()
            Me.cmdFeeCat.Parameters.AddWithValue("@feeCatName", Me.cboFeeCategoryName.Text.Trim)
            Me.cmdFeeCat.Parameters.AddWithValue("@regBy", userName.Trim)
            Me.cmdFeeCat.Parameters.AddWithValue("@queryType", "INSERT")
            rec = Me.cmdFeeCat.ExecuteNonQuery
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
    Private Function checkIfRecordExists(ByVal feeCategoryName As String)
        Dim recordExists As Boolean = False
        Me.cmdFeeCat.Connection = conn
        Me.cmdFeeCat.CommandType = CommandType.Text
        Me.cmdFeeCat.CommandText = "SELECT * FROM tblFinFeeCategory WHERE (feeCatName=@feeCatName)"
        Me.cmdFeeCat.Parameters.Clear()
        Me.cmdFeeCat.Parameters.AddWithValue("@feeCatName", feeCategoryName.Trim)
        reader = Me.cmdFeeCat.ExecuteReader
        If reader.HasRows = True Then
            recordExists = True
        ElseIf reader.HasRows = False Then
            recordExists = False
        End If
        reader.Close()
        Return recordExists
    End Function

    Private Sub UPDATEToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UPDATEToolStripMenuItem.Click
        If Me.lstFeeCategory.Items.Count <= 0 Then
            MsgBox("Missing Items In The List.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.lstFeeCategory.CheckedItems.Count <= 0 Then
            MsgBox("Missing Checked Items In The List.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.lstFeeCategory.CheckedItems.Count > 1 Then
            MsgBox("Edit one item At A Time.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            Me.cboFeeCategoryName.Tag = ""
            Me.lstFeeCategory.Tag = ""
            Me.cboFeeCategoryName.Text = Me.lstFeeCategory.CheckedItems(0).Text.Trim
            Me.cboFeeCategoryName.Tag = Me.lstFeeCategory.CheckedItems(0).Tag
            Me.lstFeeCategory.Tag = Me.lstFeeCategory.CheckedItems(0).Text.Trim
            Me.lstFeeCategory.Items.Clear()
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

    Private Sub btnUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdate.Click
        If Me.cboFeeCategoryName.Text.Trim.Length <= 0 Then
            MsgBox("Fee Category Name Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()

            Dim recordExists As Boolean = checkIfRecordExists(Me.cboFeeCategoryName.Text.Trim)
            If recordExists = True Then
                MsgBox("Duplicate Found in the database!", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
                Exit Sub
            End If

            Dim result As MsgBoxResult = MsgBox("Update Record?", MsgBoxStyle.Question + MsgBoxStyle.ApplicationModal + MsgBoxStyle.YesNo, "Confirm Transaction")
            If result = MsgBoxResult.No Then
                Exit Sub
            End If

            Me.cmdFeeCat.Connection = conn
            Me.cmdFeeCat.CommandType = CommandType.StoredProcedure
            Me.cmdFeeCat.CommandText = "sprocFinFeeCategory"
            Me.cmdFeeCat.Parameters.Clear()
            Me.cmdFeeCat.Parameters.AddWithValue("@feeCatName", Me.cboFeeCategoryName.Text.Trim)
            Me.cmdFeeCat.Parameters.AddWithValue("@regBy", userName.Trim)
            Me.cmdFeeCat.Parameters.AddWithValue("@queryType", "UPDATE")
            Me.cmdFeeCat.Parameters.AddWithValue("@prevFeeCatName", Me.lstFeeCategory.Tag)
            Me.cmdFeeCat.Parameters.AddWithValue("@feeCatId", Me.cboFeeCategoryName.Tag)
            rec = Me.cmdFeeCat.ExecuteNonQuery
            If rec > 0 Then
                MsgBox("Record Updated Successfully.", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "System Reply")
            End If
            loadCombos()
            loadList()
            Me.cboFeeCategoryName.Tag = ""
            Me.lstFeeCategory.Tag = ""
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
        If Me.lstFeeCategory.Items.Count <= 0 Then
            MsgBox("Missing Items In The List.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.lstFeeCategory.CheckedItems.Count <= 0 Then
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

            For i = 0 To Me.lstFeeCategory.CheckedItems.Count - 1
                Me.cmdFeeCat.Connection = conn
                Me.cmdFeeCat.CommandType = CommandType.StoredProcedure
                Me.cmdFeeCat.CommandText = "sprocFinFeeCategory"
                Me.cmdFeeCat.Parameters.Clear()
                Me.cmdFeeCat.Parameters.AddWithValue("@feeCatName", Me.lstFeeCategory.CheckedItems(i).Text.Trim)
                Me.cmdFeeCat.Parameters.AddWithValue("@regBy", userName.Trim)
                Me.cmdFeeCat.Parameters.AddWithValue("@queryType", "DELETE")
                Me.cmdFeeCat.Parameters.AddWithValue("@feeCatId", Me.lstFeeCategory.CheckedItems(i).Tag)
                rec = Me.cmdFeeCat.ExecuteNonQuery
            Next
            If rec > 0 Then
                MsgBox("Record/s Deleted Successfully.", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "System Reply")
            End If
            loadCombos()
            loadList()
            Me.cboFeeCategoryName.Tag = ""
            Me.lstFeeCategory.Tag = ""
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub
End Class