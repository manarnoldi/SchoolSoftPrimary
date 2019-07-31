Imports System.Data.SqlClient
Public Class frmFinExpenseMaster
    Dim reader As SqlDataReader
    Dim rec As Integer
    Dim cmdExpMaster As New SqlCommand
    Private Sub frmFinExpenseMaster_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
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

    Private Sub frmFinExpenseMaster_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint
        If Me.IsMdiChild Then
            Dim PnlLoc As New Point
            PnlLoc.X = CInt((Me.Width - Me.pnlExpMaster.Width) / 2)
            PnlLoc.Y = CInt((Me.Height - Me.pnlExpMaster.Height) / 2.5)
            Me.pnlExpMaster.Location = PnlLoc
        Else
            Me.pnlExpMaster.Dock = DockStyle.Fill
        End If
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub loadCombos()
        Me.cboCatName.Items.Clear()
        Me.cboCatName.Text = ""
        Me.cboCatName.SelectedIndex = -1

        Me.cmdExpMaster.Connection = conn
        Me.cmdExpMaster.CommandType = CommandType.Text
        Me.cmdExpMaster.CommandText = "SELECT DISTINCT expCategoryName FROM tblFinExpCategory ORDER BY expCategoryName"
        Me.cmdExpMaster.Parameters.Clear()
        reader = Me.cmdExpMaster.ExecuteReader
        While reader.Read
            Me.cboCatName.Items.Add(IIf(DBNull.Value.Equals(reader!expCategoryName), "", reader!expCategoryName))
        End While
        reader.Close()
    End Sub
    Private Sub loadList()
        Me.lstExpenseName.Items.Clear()

        Me.cmdExpMaster.Connection = conn
        Me.cmdExpMaster.CommandType = CommandType.Text
        Me.cmdExpMaster.CommandText = "SELECT * FROM vwFinExpName ORDER BY expName,expCategoryName"
        Me.cmdExpMaster.Parameters.Clear()
        reader = Me.cmdExpMaster.ExecuteReader
        While reader.Read
            li = Me.lstExpenseName.Items.Add(IIf(DBNull.Value.Equals(reader!expName), "", reader!expName))
            li.SubItems.Add(IIf(DBNull.Value.Equals(reader!expCategoryName), "", reader!expCategoryName))
            li.Tag = IIf(DBNull.Value.Equals(reader!expNameId), "", reader!expNameId)
        End While
        reader.Close()
    End Sub
    Private Sub clearTexts()
        Me.btnSave.Enabled = True
        Me.btnUpdate.Enabled = False
        Me.cboCatName.Enabled = True
        Me.txtExpName.Text = Nothing
        Me.txtExpName.Tag = Nothing
        Me.btnUpdate.Tag = Nothing
        Me.lstExpenseName.Enabled = True
    End Sub
    Private Function checkIfRecordExists(ByVal expenseName As String)
        Dim recordFound As Boolean = True

        Me.cmdExpMaster.Connection = conn
        Me.cmdExpMaster.CommandType = CommandType.Text
        Me.cmdExpMaster.CommandText = "SELECT * FROM vwFinExpName WHERE (expName=@expName)"
        Me.cmdExpMaster.Parameters.Clear()
        Me.cmdExpMaster.Parameters.AddWithValue("@expName", expenseName.Trim)
        reader = Me.cmdExpMaster.ExecuteReader
        If reader.HasRows = True Then
            recordFound = True
        ElseIf reader.HasRows = False Then
            recordFound = False
        End If
        reader.Close()
        Return recordFound
    End Function
    Private Function checkIfRecordExistsUpdate(ByVal expenseCategory As String, ByVal expenseName As String)
        Dim recordFound As Boolean = True

        Me.cmdExpMaster.Connection = conn
        Me.cmdExpMaster.CommandType = CommandType.Text
        Me.cmdExpMaster.CommandText = "SELECT * FROM vwFinExpName WHERE (expName=@expName) AND (expCategoryName=@expCategoryName)"
        Me.cmdExpMaster.Parameters.Clear()
        Me.cmdExpMaster.Parameters.AddWithValue("@expName", expenseName.Trim)
        Me.cmdExpMaster.Parameters.AddWithValue("@expCategoryName", expenseCategory.Trim)
        reader = Me.cmdExpMaster.ExecuteReader
        If reader.HasRows = True Then
            recordFound = True
        ElseIf reader.HasRows = False Then
            recordFound = False
        End If
        reader.Close()
        Return recordFound
    End Function

    Private Sub CLOSEToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CLOSEToolStripMenuItem.Click
        Me.Close()
    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        If Me.cboCatName.Text.Trim.Count <= 0 Then
            MsgBox("Expense Category Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.txtExpName.Text.Trim.Length <= 0 Then
            MsgBox("Expense Name Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()

            Dim recordExists As Boolean = checkIfRecordExists(Me.txtExpName.Text.Trim)
            If recordExists = True Then
                MsgBox("Duplicate Found in the database!", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
                Exit Sub
            End If

            Dim result As MsgBoxResult = MsgBox("Save Record?", MsgBoxStyle.Question + MsgBoxStyle.ApplicationModal + MsgBoxStyle.YesNo, "Confirm Transaction")
            If result = MsgBoxResult.No Then
                Exit Sub
            End If

            Me.cmdExpMaster.Connection = conn
            Me.cmdExpMaster.CommandType = CommandType.StoredProcedure
            Me.cmdExpMaster.CommandText = "sprocFinExpenseName"
            Me.cmdExpMaster.Parameters.Clear()
            Me.cmdExpMaster.Parameters.AddWithValue("@expCategoryName", Me.cboCatName.Text.Trim)
            Me.cmdExpMaster.Parameters.AddWithValue("@userName", userName.Trim)
            Me.cmdExpMaster.Parameters.AddWithValue("@queryType", "INSERT")
            Me.cmdExpMaster.Parameters.AddWithValue("@expName", Me.txtExpName.Text.Trim)
            rec = Me.cmdExpMaster.ExecuteNonQuery
            If rec > 0 Then
                MsgBox("Record/s Saved", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "SuccessFull Transaction")
            End If
            clearTexts()
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
        If Me.lstExpenseName.Items.Count <= 0 Then
            MsgBox("Items Missing in the List.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.lstExpenseName.CheckedItems.Count <= 0 Then
            MsgBox("Check item to Edit.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.lstExpenseName.CheckedItems.Count > 1 Then
            MsgBox("Edit one item at a time.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If

        Me.cboCatName.Text = Me.lstExpenseName.CheckedItems(0).SubItems(1).Text.Trim
        Me.cboCatName.Enabled = False

        Me.btnSave.Enabled = False
        Me.btnUpdate.Enabled = True

        Me.txtExpName.Text = Me.lstExpenseName.CheckedItems(0).Text.Trim
        Me.txtExpName.Tag = Me.lstExpenseName.CheckedItems(0).Tag

        Me.btnUpdate.Tag = Me.lstExpenseName.CheckedItems(0).Text.Trim

        Me.lstExpenseName.Items.Clear()
        Me.lstExpenseName.Enabled = False
    End Sub

    Private Sub btnView_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnView.Click
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            clearTexts()
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
        If Me.cboCatName.Text.Trim.Count <= 0 Then
            MsgBox("Expense Category Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.txtExpName.Text.Trim.Length <= 0 Then
            MsgBox("Expense Name Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()

            Dim recordExists As Boolean = checkIfRecordExistsUpdate(Me.cboCatName.Text.Trim, Me.txtExpName.Text.Trim)
            If recordExists = True Then
                MsgBox("Duplicate Found in the database!", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
                Exit Sub
            End If

            Dim result As MsgBoxResult = MsgBox("Update Record?", MsgBoxStyle.Question + MsgBoxStyle.ApplicationModal + MsgBoxStyle.YesNo, "Confirm Transaction")
            If result = MsgBoxResult.No Then
                Exit Sub
            End If

            Me.cmdExpMaster.Connection = conn
            Me.cmdExpMaster.CommandType = CommandType.StoredProcedure
            Me.cmdExpMaster.CommandText = "sprocFinExpenseName"
            Me.cmdExpMaster.Parameters.Clear()
            Me.cmdExpMaster.Parameters.AddWithValue("@expName", Me.txtExpName.Text.Trim)
            Me.cmdExpMaster.Parameters.AddWithValue("@expCategoryName", Me.cboCatName.Text.Trim)
            Me.cmdExpMaster.Parameters.AddWithValue("@userName", userName.Trim)
            Me.cmdExpMaster.Parameters.AddWithValue("@queryType", "UPDATE")
            Me.cmdExpMaster.Parameters.AddWithValue("@expNameId", Me.txtExpName.Tag)
            Me.cmdExpMaster.Parameters.AddWithValue("@prevExpName", Me.btnUpdate.Tag)
            rec = Me.cmdExpMaster.ExecuteNonQuery
            If rec > 0 Then
                MsgBox("Record/s Updated", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "SuccessFull Transaction")
            End If
            clearTexts()
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

    Private Sub DELETEToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DELETEToolStripMenuItem.Click
        If Me.lstExpenseName.Items.Count <= 0 Then
            MsgBox("Items Missing in the List.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.lstExpenseName.CheckedItems.Count <= 0 Then
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

            For i = 0 To Me.lstExpenseName.CheckedItems.Count - 1
                Me.cmdExpMaster.Connection = conn
                Me.cmdExpMaster.CommandType = CommandType.StoredProcedure
                Me.cmdExpMaster.CommandText = "sprocFinExpenseName"
                Me.cmdExpMaster.Parameters.Clear()
                Me.cmdExpMaster.Parameters.AddWithValue("@expName", Me.lstExpenseName.CheckedItems(i).Text.Trim)
                Me.cmdExpMaster.Parameters.AddWithValue("@expCategoryName", Me.lstExpenseName.CheckedItems(i).SubItems(1).Text.Trim)
                Me.cmdExpMaster.Parameters.AddWithValue("@userName", userName.Trim)
                Me.cmdExpMaster.Parameters.AddWithValue("@queryType", "DELETE")
                Me.cmdExpMaster.Parameters.AddWithValue("@expNameId", Me.lstExpenseName.CheckedItems(i).Tag)
                rec = rec + Me.cmdExpMaster.ExecuteNonQuery
            Next
            If rec > 0 Then
                MsgBox("Record/s Deleted", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "SuccessFull Transaction")
            End If
            clearTexts()
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
End Class