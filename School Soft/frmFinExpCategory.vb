Imports System.Data.SqlClient
Public Class frmFinExpCategory
    Dim reader As SqlDataReader
    Dim rec As Integer
    Dim cmdExpCategory As New SqlCommand
    Private Sub frmFinExpCategory_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
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
        Me.txtCategoryName.Text = ""
        Me.txtCategoryName.Tag = Nothing

        Me.lstExpCategory.Enabled = True
        Me.btnSave.Enabled = True
        Me.btnUpdate.Enabled = False

        Me.lstExpCategory.Items.Clear()

        Me.btnUpdate.Tag = Nothing

        Me.cmdExpCategory.Connection = conn
        Me.cmdExpCategory.CommandType = CommandType.Text
        Me.cmdExpCategory.CommandText = "SELECT * FROM tblFinExpCategory ORDER BY expCategoryName"
        Me.cmdExpCategory.Parameters.Clear()
        reader = Me.cmdExpCategory.ExecuteReader
        While reader.Read
            li = Me.lstExpCategory.Items.Add(IIf(DBNull.Value.Equals(reader!expCategoryName), "", reader!expCategoryName))
            li.Tag = IIf(DBNull.Value.Equals(reader!expCatId), "", reader!expCatId)
        End While
        reader.Close()
    End Sub
    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub frmFinExpCategory_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint
        If Me.IsMdiChild Then
            Dim PnlLoc As New Point
            PnlLoc.X = CInt((Me.Width - Me.pnlExpCat.Width) / 2)
            PnlLoc.Y = CInt((Me.Height - Me.pnlExpCat.Height) / 2.5)
            Me.pnlExpCat.Location = PnlLoc
        Else
            Me.pnlExpCat.Dock = DockStyle.Fill
        End If
    End Sub

    Private Sub btnView_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnView.Click
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            loadList()
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        If Me.txtCategoryName.Text.Trim.Count <= 0 Then
            MsgBox("Expense Category Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()

            Dim recordExists As Boolean = checkIfRecordExists(Me.txtCategoryName.Text.Trim)
            If recordExists = True Then
                MsgBox("Duplicate Found in the database!", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
                Exit Sub
            End If

            Dim result As MsgBoxResult = MsgBox("Save Record?", MsgBoxStyle.Question + MsgBoxStyle.ApplicationModal + MsgBoxStyle.YesNo, "Confirm Transaction")
            If result = MsgBoxResult.No Then
                Exit Sub
            End If

            Me.cmdExpCategory.Connection = conn
            Me.cmdExpCategory.CommandType = CommandType.StoredProcedure
            Me.cmdExpCategory.CommandText = "sprocFinExpCategory"
            Me.cmdExpCategory.Parameters.Clear()
            Me.cmdExpCategory.Parameters.AddWithValue("@expCategoryName", Me.txtCategoryName.Text.Trim)
            Me.cmdExpCategory.Parameters.AddWithValue("@userName", userName.Trim)
            Me.cmdExpCategory.Parameters.AddWithValue("@queryType", "INSERT")
            rec = Me.cmdExpCategory.ExecuteNonQuery
            If rec > 0 Then
                MsgBox("Record/s Saved", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "SuccessFull Transaction")
            End If
            loadList()
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
        If Me.lstExpCategory.Items.Count <= 0 Then
            MsgBox("Items Missing in the List.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.lstExpCategory.CheckedItems.Count <= 0 Then
            MsgBox("Check item to Edit.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.lstExpCategory.CheckedItems.Count > 1 Then
            MsgBox("Edit one item at a time.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If

        Me.txtCategoryName.Text = Me.lstExpCategory.CheckedItems(0).Text.Trim
        Me.btnUpdate.Tag = Me.lstExpCategory.CheckedItems(0).Text.Trim
        Me.txtCategoryName.Tag = Me.lstExpCategory.CheckedItems(0).Tag
        Me.lstExpCategory.Items.Clear()
        Me.lstExpCategory.Enabled = False
        Me.btnSave.Enabled = False
        Me.btnUpdate.Enabled = True
    End Sub
    Private Function checkIfRecordExists(ByVal expCatName As String)
        Dim duplicateFound As Boolean = True

        Me.cmdExpCategory.CommandType = CommandType.Text
        Me.cmdExpCategory.Connection = conn
        Me.cmdExpCategory.CommandText = "SELECT * FROM tblFinExpCategory WHERE (expCategoryName=@expCategoryName)"
        Me.cmdExpCategory.Parameters.Clear()
        Me.cmdExpCategory.Parameters.AddWithValue("@expCategoryName", expCatName.Trim)
        reader = Me.cmdExpCategory.ExecuteReader
        If reader.HasRows = True Then
            duplicateFound = True
        ElseIf reader.HasRows = False Then
            duplicateFound = False
        End If
        reader.Close()
        Return duplicateFound
    End Function

    Private Sub Update_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdate.Click
        If Me.txtCategoryName.Text.Trim.Count <= 0 Then
            MsgBox("Expense Category Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.txtCategoryName.Tag = Nothing Then
            MsgBox("Important Details Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()

            Dim recordExists As Boolean = checkIfRecordExists(Me.txtCategoryName.Text.Trim)
            If recordExists = True Then
                MsgBox("Duplicate Found in the database!", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
                Exit Sub
            End If

            Dim result As MsgBoxResult = MsgBox("Update Record?", MsgBoxStyle.Question + MsgBoxStyle.ApplicationModal + MsgBoxStyle.YesNo, "Confirm Transaction")
            If result = MsgBoxResult.No Then
                Exit Sub
            End If

            Me.cmdExpCategory.Connection = conn
            Me.cmdExpCategory.CommandType = CommandType.StoredProcedure
            Me.cmdExpCategory.CommandText = "sprocFinExpCategory"
            Me.cmdExpCategory.Parameters.Clear()
            Me.cmdExpCategory.Parameters.AddWithValue("@expCategoryName", Me.txtCategoryName.Text.Trim)
            Me.cmdExpCategory.Parameters.AddWithValue("@userName", userName.Trim)
            Me.cmdExpCategory.Parameters.AddWithValue("@queryType", "UPDATE")
            Me.cmdExpCategory.Parameters.AddWithValue("@expCatId", Me.txtCategoryName.Tag)
            Me.cmdExpCategory.Parameters.AddWithValue("@prevExpCategoryName", Me.btnUpdate.Tag)
            rec = Me.cmdExpCategory.ExecuteNonQuery
            If rec > 0 Then
                MsgBox("Record/s Updated", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "SuccessFull Transaction")
            End If
            loadList()
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub

    Private Sub DELETEToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DELETEToolStripMenuItem.Click
        If Me.lstExpCategory.Items.Count <= 0 Then
            MsgBox("Items Missing in the List.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.lstExpCategory.CheckedItems.Count <= 0 Then
            MsgBox("Check item to Edit.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()

            Dim result As MsgBoxResult = MsgBox("Delete Record?", MsgBoxStyle.Question + MsgBoxStyle.ApplicationModal + MsgBoxStyle.YesNo, "Confirm Transaction")
            If result = MsgBoxResult.No Then
                Exit Sub
            End If

            For i = 0 To Me.lstExpCategory.CheckedItems.Count - 1
                Me.cmdExpCategory.Connection = conn
                Me.cmdExpCategory.CommandType = CommandType.StoredProcedure
                Me.cmdExpCategory.CommandText = "sprocFinExpCategory"
                Me.cmdExpCategory.Parameters.Clear()
                Me.cmdExpCategory.Parameters.AddWithValue("@expCategoryName", Me.lstExpCategory.CheckedItems(i).Text.Trim)
                Me.cmdExpCategory.Parameters.AddWithValue("@userName", userName.Trim)
                Me.cmdExpCategory.Parameters.AddWithValue("@queryType", "DELETE")
                Me.cmdExpCategory.Parameters.AddWithValue("@expCatId", Me.lstExpCategory.CheckedItems(i).Tag)
                rec = rec + Me.cmdExpCategory.ExecuteNonQuery
            Next
            If rec > 0 Then
                MsgBox("Record/s Deleted", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "SuccessFull Transaction")
            End If
            loadList()
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub
End Class