Imports System.Data.SqlClient
Public Class frmInvMasterCategory
    Dim reader As SqlDataReader
    Dim rec As Integer
    Dim cmdCategory As New SqlCommand
    Private Sub frmInvCategory_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
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
        Dim serialNo As Integer = 1
        Me.lstInvCategoryMaster.Items.Clear()

        Me.cmdCategory.Connection = conn
        Me.cmdCategory.CommandType = CommandType.Text
        Me.cmdCategory.CommandText = "SELECT * FROM tblInvMasterCategory ORDER BY categoryName"
        Me.cmdCategory.Parameters.Clear()
        reader = Me.cmdCategory.ExecuteReader
        While reader.Read
            li = Me.lstInvCategoryMaster.Items.Add(serialNo)
            li.SubItems.Add(IIf(DBNull.Value.Equals(reader!categoryName), "", reader!categoryName))
            serialNo = serialNo + 1
        End While
        reader.Close()
    End Sub
    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub frmInvCategory_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint
        If Me.IsMdiChild Then
            Dim PnlLoc As New Point
            PnlLoc.X = CInt((Me.Width - Me.pnlInvCategories.Width) / 2)
            PnlLoc.Y = CInt((Me.Height - Me.pnlInvCategories.Height) / 2.5)
            Me.pnlInvCategories.Location = PnlLoc
        Else
            Me.pnlInvCategories.Dock = DockStyle.Fill
        End If
    End Sub
    Private Function checkIfRecordExists(ByVal categoryName As String)
        Dim recordExists As Boolean = False
        Me.cmdCategory.Connection = conn
        Me.cmdCategory.CommandType = CommandType.Text
        Me.cmdCategory.CommandText = "SELECT * FROM tblInvMasterCategory WHERE (categoryName=@categoryName)"
        Me.cmdCategory.Parameters.Clear()
        Me.cmdCategory.Parameters.AddWithValue("@categoryName", categoryName.Trim)
        reader = Me.cmdCategory.ExecuteReader
        If reader.HasRows = True Then
            recordExists = True
        ElseIf reader.HasRows = False Then
            recordExists = False
        End If
        reader.Close()
        Return recordExists
    End Function

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        If Me.txtCategoryName.Text.Trim.Length <= 0 Then
            MsgBox("Category Name Is Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
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

            Me.cmdCategory.Connection = conn
            Me.cmdCategory.CommandType = CommandType.StoredProcedure
            Me.cmdCategory.CommandText = "sprocInvCategoryMaster"
            Me.cmdCategory.Parameters.Clear()
            Me.cmdCategory.Parameters.AddWithValue("@categoryName", Me.txtCategoryName.Text.Trim)
            Me.cmdCategory.Parameters.AddWithValue("@regBy", userName.Trim)
            Me.cmdCategory.Parameters.AddWithValue("@dateOfReg", Date.Now)
            Me.cmdCategory.Parameters.AddWithValue("@queryType", "INSERT")
            rec = Me.cmdCategory.ExecuteNonQuery

            If rec > 0 Then
                MsgBox("Record Saved Successfully.", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "System Reply")
            End If
            loadList()
            Me.txtCategoryName.Text = ""
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub

    Private Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        If Me.lstInvCategoryMaster.Items.Count <= 0 Then
            MsgBox("No items in the List.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.lstInvCategoryMaster.CheckedItems.Count <= 0 Then
            MsgBox("No checked items in the List.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
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

            i = 0
            Me.cmdCategory.Connection = conn
            Me.cmdCategory.CommandType = CommandType.StoredProcedure
            For i = 0 To Me.lstInvCategoryMaster.CheckedItems.Count - 1
                Me.cmdCategory.CommandText = "sprocInvCategoryMaster"
                Me.cmdCategory.Parameters.Clear()
                Me.cmdCategory.Parameters.AddWithValue("@categoryName", Me.lstInvCategoryMaster.CheckedItems(i).SubItems(1).Text.Trim)
                Me.cmdCategory.Parameters.AddWithValue("@regBy", userName.Trim)
                Me.cmdCategory.Parameters.AddWithValue("@dateOfReg", Date.Now)
                Me.cmdCategory.Parameters.AddWithValue("@queryType", "DELETE")
                rec = Me.cmdCategory.ExecuteNonQuery
            Next
            If rec > 0 Then
                MsgBox("Record Deleted Successfully.", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "System Reply")
            End If
            loadList()
            Me.txtCategoryName.Text = ""
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub
End Class