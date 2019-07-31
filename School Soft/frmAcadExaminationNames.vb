Imports System.Data.SqlClient
Public Class frmAcadExaminationNames
    Dim recordExists As Boolean = True
    Dim queryType As String = Nothing
    Dim reader As SqlDataReader
    Dim cmdExamNames As New SqlCommand
    Dim rec As Integer = 0
    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub frmAcadExaminationNames_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
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
    Private Sub loadCombos()
        Me.cboExamType.Items.Clear()
        Me.cboExamType.Text = ""
        Me.cboExamType.SelectedIndex = -1

        Me.cmdExamNames.Connection = conn
        Me.cmdExamNames.CommandType = CommandType.Text
        Me.cmdExamNames.CommandText = "SELECT DISTINCT examType FROM tblExamNames WHERE (status=1) AND (examType IS NOT NULL) ORDER BY examType"
        Me.cmdExamNames.Parameters.Clear()
        reader = Me.cmdExamNames.ExecuteReader
        While reader.Read
            Me.cboExamType.Items.Add(IIf(DBNull.Value.Equals(reader!examType), "", (reader!examType)))
        End While
        reader.Close()
    End Sub
    Private Sub loadList()
        Me.lstViewExamNames.Items.Clear()

        Me.cmdExamNames.Connection = conn
        Me.cmdExamNames.CommandType = CommandType.Text
        Me.cmdExamNames.CommandText = "SELECT * FROM tblExamNames WHERE (status=1) ORDER BY examType,ExamName"
        Me.cmdExamNames.Parameters.Clear()
        reader = Me.cmdExamNames.ExecuteReader
        While reader.Read
            li = Me.lstViewExamNames.Items.Add(IIf(DBNull.Value.Equals(reader!examType), "", (reader!examType)))
            li.SubItems.Add(IIf(DBNull.Value.Equals(reader!examName), "", (reader!examName)))
        End While
        reader.Close()
    End Sub

    Private Sub frmAcadExaminationNames_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint
        If Me.IsMdiChild Then
            Dim PnlLoc As New Point
            PnlLoc.X = CInt((Me.Width - Me.pnlExamNames.Width) / 2)
            PnlLoc.Y = CInt((Me.Height - Me.pnlExamNames.Height) / 2.5)
            Me.pnlExamNames.Location = PnlLoc
        Else
            Me.pnlExamNames.Dock = DockStyle.Fill
        End If
    End Sub

    Private Sub cboExamType_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboExamType.SelectedIndexChanged
        If Me.cboExamType.Text.Trim.Length <= 0 Then
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            Me.cboExamName.Items.Clear()
            Me.cboExamName.Text = ""
            Me.cboExamName.SelectedIndex = -1

            Me.cmdExamNames.Connection = conn
            Me.cmdExamNames.CommandType = CommandType.Text
            Me.cmdExamNames.CommandText = "SELECT examName FROM tblExamNames WHERE (status=1) AND (examType=@examType) ORDER BY examName"
            Me.cmdExamNames.Parameters.Clear()
            Me.cmdExamNames.Parameters.AddWithValue("@examType", Me.cboExamType.Text.Trim)
            reader = Me.cmdExamNames.ExecuteReader
            While reader.Read
                Me.cboExamName.Items.Add(IIf(DBNull.Value.Equals(reader!examName), "", (reader!examName)))
            End While
            reader.Close()
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

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        If Me.cboExamName.Enabled = False Then
            MsgBox("Record is for Update." & vbNewLine & "Update the Record.", MsgBoxStyle.Exclamation + _
                   MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Missing Details")
            Exit Sub
        ElseIf Me.cboExamType.Text.Trim.Length <= 0 Then
            MsgBox("Exam Type is Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Missing Details")
            Exit Sub
        ElseIf Me.cboExamName.Text.Trim.Length <= 0 Then
            MsgBox("Exam Name is Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Missing Details")
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            checkIfRecordExists()
            If recordExists = True Then
                MsgBox("Duplicate Already Exists In The System.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Duplicate Found")
                Exit Sub
            End If
            Dim result As MsgBoxResult = MsgBox("Save Record?", MsgBoxStyle.Question + MsgBoxStyle.ApplicationModal + MsgBoxStyle.YesNo, "Confirm Transaction")
            If result = MsgBoxResult.No Then
                Exit Sub
            End If
            Me.queryType = "INSERT"
            Me.cmdExamNames.Connection = conn
            Me.cmdExamNames.CommandType = CommandType.StoredProcedure
            Me.cmdExamNames.CommandText = "sprocExamNames"
            Me.cmdExamNames.Parameters.Clear()
            Me.cmdExamNames.Parameters.AddWithValue("@queryType", Me.queryType.Trim)
            Me.cmdExamNames.Parameters.AddWithValue("@examName", Me.cboExamName.Text.Trim)
            Me.cmdExamNames.Parameters.AddWithValue("@examType", Me.cboExamType.Text.Trim)
            Me.cmdExamNames.Parameters.AddWithValue("@regBy", userName.Trim)
            Me.cmdExamNames.Parameters.AddWithValue("@dateOfReg", Date.Now)
            rec = Me.cmdExamNames.ExecuteNonQuery
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
    Private Sub clearTexts()
        Me.cboExamName.Items.Clear()
        Me.cboExamName.Text = ""
        Me.cboExamName.SelectedIndex = -1

        Me.cboExamType.Items.Clear()
        Me.cboExamType.Text = ""
        Me.cboExamType.SelectedIndex = -1

        Me.lstViewExamNames.Items.Clear()
        Me.cboExamName.Enabled = True
    End Sub
    Private Sub btnView_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnView.Click
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            clearTexts()
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

    Private Sub UPDATEToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UPDATEToolStripMenuItem.Click
        If Me.lstViewExamNames.Items.Count <= 0 Then
            MsgBox("No items in the list to Update.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Missing Details")
            Exit Sub
        ElseIf Me.lstViewExamNames.SelectedItems.Count <= 0 Then
            MsgBox("No items selected in the list to Update.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Missing Details")
            Exit Sub
        ElseIf Me.lstViewExamNames.SelectedItems.Count > 1 Then
            MsgBox("Update one item at a time.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Missing Details")
            Exit Sub
        End If
        Me.cboExamType.Text = Me.lstViewExamNames.SelectedItems(0).Text.Trim
        Me.cboExamName.Enabled = False
        Me.cboExamName.Text = Me.lstViewExamNames.SelectedItems(0).SubItems(1).Text.Trim
        
        Me.lstViewExamNames.Items.Clear()
    End Sub

    Private Sub btnClear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClear.Click
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            clearTexts()
            loadList()
            loadCombos()
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub
    Private Sub checkIfRecordExists()
        Me.cmdExamNames.Connection = conn
        Me.cmdExamNames.CommandType = CommandType.Text
        Me.cmdExamNames.CommandText = "SELECT * FROM tblExamNames WHERE (examName=@examName) AND (status=1)"
        Me.cmdExamNames.Parameters.Clear()
        Me.cmdExamNames.Parameters.AddWithValue("@examName", Me.cboExamName.Text.Trim)
        reader = Me.cmdExamNames.ExecuteReader
        If reader.HasRows = True Then
            recordExists = True
        ElseIf reader.HasRows = False Then
            recordExists = False
        End If
        reader.Close()
    End Sub
    Private Sub checkIfRecordExistsOne()
        Me.cmdExamNames.Connection = conn
        Me.cmdExamNames.CommandType = CommandType.Text
        Me.cmdExamNames.CommandText = "SELECT * FROM tblExamNames WHERE (examName=@examName) AND (examType=@examType) AND (status=1)"
        Me.cmdExamNames.Parameters.Clear()
        Me.cmdExamNames.Parameters.AddWithValue("@examName", Me.cboExamName.Text.Trim)
        Me.cmdExamNames.Parameters.AddWithValue("@examType", Me.cboExamType.Text.Trim)
        reader = Me.cmdExamNames.ExecuteReader
        If reader.HasRows = True Then
            recordExists = True
        ElseIf reader.HasRows = False Then
            recordExists = False
        End If
        reader.Close()
    End Sub
    Private Sub btnUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdate.Click
        If Me.cboExamName.Enabled = True Then
            MsgBox("Select and Right click item to Update From the list.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Missing Details")
            Exit Sub
        ElseIf Me.cboExamType.Text.Trim.Length <= 0 Then
            MsgBox("Exam Type is Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Missing Details")
            Exit Sub
        ElseIf Me.cboExamName.Text.Trim.Length <= 0 Then
            MsgBox("Exam Name is Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Missing Details")
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            checkIfRecordExistsOne()
            If recordExists = True Then
                MsgBox("Duplicate Already Exists In The System.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Duplicate Found")
                Exit Sub
            End If
            Dim result As MsgBoxResult = MsgBox("Update Record?", MsgBoxStyle.Question + MsgBoxStyle.ApplicationModal + MsgBoxStyle.YesNo, "Confirm Transaction")
            If result = MsgBoxResult.No Then
                Exit Sub
            End If
            Me.queryType = "UPDATE"
            Me.cmdExamNames.Connection = conn
            Me.cmdExamNames.CommandType = CommandType.StoredProcedure
            Me.cmdExamNames.CommandText = "sprocExamNames"
            Me.cmdExamNames.Parameters.Clear()
            Me.cmdExamNames.Parameters.AddWithValue("@queryType", Me.queryType.Trim)
            Me.cmdExamNames.Parameters.AddWithValue("@examName", Me.cboExamName.Text.Trim)
            Me.cmdExamNames.Parameters.AddWithValue("@examType", Me.cboExamType.Text.Trim)
            Me.cmdExamNames.Parameters.AddWithValue("@regBy", userName.Trim)
            Me.cmdExamNames.Parameters.AddWithValue("@dateOfReg", Date.Now)
            rec = Me.cmdExamNames.ExecuteNonQuery
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
        If Me.lstViewExamNames.Items.Count <= 0 Then
            MsgBox("No items in the list to Update.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Missing Details")
            Exit Sub
        ElseIf Me.lstViewExamNames.SelectedItems.Count <= 0 Then
            MsgBox("No items selected in the list to Update.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Missing Details")
            Exit Sub
        ElseIf Me.lstViewExamNames.SelectedItems.Count > 1 Then
            MsgBox("Delete one item at a time.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Missing Details")
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
            Me.queryType = "DELETE"
            Me.cmdExamNames.Connection = conn
            Me.cmdExamNames.CommandType = CommandType.StoredProcedure
            Me.cmdExamNames.CommandText = "sprocExamNames"
            Me.cmdExamNames.Parameters.Clear()
            Me.cmdExamNames.Parameters.AddWithValue("@queryType", Me.queryType.Trim)
            Me.cmdExamNames.Parameters.AddWithValue("@examName", Me.lstViewExamNames.SelectedItems(0).SubItems(1).Text.Trim)
            Me.cmdExamNames.Parameters.AddWithValue("@examType", Me.lstViewExamNames.SelectedItems(0).Text.Trim)
            Me.cmdExamNames.Parameters.AddWithValue("@regBy", userName.Trim)
            Me.cmdExamNames.Parameters.AddWithValue("@dateOfReg", Date.Now)
            rec = Me.cmdExamNames.ExecuteNonQuery
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