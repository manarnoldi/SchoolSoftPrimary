Imports System.Threading
Imports System.Data.SqlClient
Public Class frmAcadSubjects
    Dim reader As SqlDataReader
    Dim cmdSubjects As New SqlCommand
    Dim queryType As String = Nothing
    Dim rec As Integer = 0
    Dim queryOk As Boolean = False
    Dim recordExists As Boolean = True
    Private Sub frmAcadSubjects_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint
        If Me.IsMdiChild Then
            Dim PnlLoc As New Point
            PnlLoc.X = CInt((Me.Width - Me.pnlAcadSubj.Width) / 2)
            PnlLoc.Y = CInt((Me.Height - Me.pnlAcadSubj.Height) / 2.5)
            Me.pnlAcadSubj.Location = PnlLoc
        Else
            Me.pnlAcadSubj.Dock = DockStyle.Fill
        End If
    End Sub

    Private Sub frmAcadSubjects_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
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
        cboSubjGroup.Items.Clear()
        cboSubjGroup.Text = ""
        cboSubjGroup.SelectedIndex = -1
        cmdSubjects.Connection = conn
        cmdSubjects.CommandType = CommandType.Text
        cmdSubjects.CommandText = "SELECT DISTINCT subGroup FROM  tblSubjects WHERE (subStatus='True') ORDER BY subGroup"
        cmdSubjects.Parameters.Clear()
        reader = cmdSubjects.ExecuteReader
        If reader.HasRows Then
            While reader.Read
                Me.cboSubjGroup.Items.Add(IIf(DBNull.Value.Equals(reader!subGroup), "", reader!subGroup))
            End While
        End If
        reader.Close()
    End Sub

    Private Sub btnView_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnView.Click
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            clearTexts()
            loadList()
            loadCombos()
            Me.txtSubjCode.Enabled = True
            Me.txtSubjName.Enabled = True
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
    Private Sub loadList()
        Me.lstSubjects.Items.Clear()
        cmdSubjects.Connection = conn
        cmdSubjects.CommandType = CommandType.Text
        cmdSubjects.CommandText = "SELECT * FROM  tblSubjects WHERE (subStatus='True') ORDER BY subCode"
        cmdSubjects.Parameters.Clear()
        reader = cmdSubjects.ExecuteReader
        If reader.HasRows Then
            While reader.Read
                li = Me.lstSubjects.Items.Add(IIf(DBNull.Value.Equals(reader!subCode), "", reader!subCode))
                li.SubItems.Add(IIf(DBNull.Value.Equals(reader!subName), "", reader!subName))
                li.SubItems.Add(IIf(DBNull.Value.Equals(reader!subGroup), "", reader!subGroup))
            End While
        End If
        reader.Close()
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        If Me.txtSubjCode.Text.Trim.Length <= 0 Then
            MsgBox("Missing Subject Code", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.txtSubjName.Text.Trim.Length <= 0 Then
            MsgBox("Missing Subject Name", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboSubjGroup.Text.Trim.Length <= 0 Then
            MsgBox("Missing Subject Group", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            recordExists = True
            checkIfRecordExists()
            If recordExists = True Then
                MsgBox("Duplicate Found In the Database", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
                Exit Sub
            End If
            Dim result As MsgBoxResult = MsgBox("Save Record?", MsgBoxStyle.Question + MsgBoxStyle.ApplicationModal + MsgBoxStyle.YesNo, "Confirm Transaction")
            If result = MsgBoxResult.No Then
                Exit Sub
            End If
            queryType = "INSERT"
            cmdSubjects.Connection = conn
            cmdSubjects.CommandType = CommandType.StoredProcedure
            cmdSubjects.CommandText = "sprocAcadSubjects"
            cmdSubjects.Parameters.Clear()
            cmdSubjects.Parameters.AddWithValue("@queryType", Me.queryType.Trim)
            cmdSubjects.Parameters.AddWithValue("@dateDone", Date.Now)
            cmdSubjects.Parameters.AddWithValue("@regBy", userName.Trim)
            cmdSubjects.Parameters.AddWithValue("@subCode", Me.txtSubjCode.Text.Trim)
            cmdSubjects.Parameters.AddWithValue("@subName", Me.txtSubjName.Text.Trim)
            cmdSubjects.Parameters.AddWithValue("@subGroup", Me.cboSubjGroup.Text.Trim)
            rec = cmdSubjects.ExecuteNonQuery
            If rec > 0 Then
                MsgBox("Record Saved!", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "SuccessFull Transaction")
            End If
            clearTexts()
            loadCombos()
            loadList()
            Me.btnUpdate.Enabled = False
            Me.btnSave.Enabled = True
            Me.txtSubjCode.Enabled = True
            Me.txtSubjName.Enabled = True
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            Me.queryType = Nothing
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub
    Private Sub clearTexts()
        Me.txtSubjCode.Text = ""
        Me.txtSubjName.Text = ""
        Me.cboSubjGroup.Text = ""
        Me.cboSubjGroup.Items.Clear()
        Me.cboSubjGroup.SelectedIndex = -1
        Me.lstSubjects.Items.Clear()
    End Sub
    Private Sub checkIfRecordExists()
        cmdSubjects.Connection = conn
        cmdSubjects.CommandType = CommandType.Text
        cmdSubjects.CommandText = "SELECT * FROM  tblSubjects WHERE (subStatus='True') AND (subName=@subName) ORDER BY subName"
        cmdSubjects.Parameters.Clear()
        cmdSubjects.Parameters.AddWithValue("@subName", Me.txtSubjName.Text.Trim)
        reader = cmdSubjects.ExecuteReader
        If reader.HasRows Then
            recordExists = True
        Else
            recordExists = False
        End If
        reader.Close()
    End Sub
    Private Sub checkIfRecordExistsOne()
        cmdSubjects.Connection = conn
        cmdSubjects.CommandType = CommandType.Text
        cmdSubjects.CommandText = "SELECT * FROM  tblSubjects WHERE (subStatus='True') AND (subName=@subName) AND (subCode=@subCode) " & _
            vbNewLine & " AND (subGroup=@subGroup) ORDER BY subName"
        cmdSubjects.Parameters.Clear()
        cmdSubjects.Parameters.AddWithValue("@subCode", Me.txtSubjCode.Text.Trim)
        cmdSubjects.Parameters.AddWithValue("@subName", Me.txtSubjName.Text.Trim)
        cmdSubjects.Parameters.AddWithValue("@subGroup", Me.cboSubjGroup.Text.Trim)
        reader = cmdSubjects.ExecuteReader
        If reader.HasRows Then
            recordExists = True
        Else
            recordExists = False
        End If
        reader.Close()
    End Sub
    Private Sub CLOSEToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CLOSEToolStripMenuItem1.Click
        Me.Close()
    End Sub

    Private Sub UPDATEToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UPDATEToolStripMenuItem.Click
        If Me.lstSubjects.Items.Count <= 0 Then
            MsgBox("No items in the list", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.lstSubjects.SelectedItems.Count <= 0 Then
            MsgBox("Select the item to update", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.lstSubjects.SelectedItems.Count > 1 Then
            MsgBox("Select One item at a time.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If
        If Me.lstSubjects.SelectedItems.Count = 1 Then
            Me.txtSubjCode.Text = Me.lstSubjects.SelectedItems(0).Text
            Me.txtSubjName.Text = Me.lstSubjects.SelectedItems(0).SubItems(1).Text
            Me.cboSubjGroup.Text = Me.lstSubjects.SelectedItems(0).SubItems(2).Text
            Me.txtSubjCode.Enabled = False
            Me.txtSubjName.Enabled = False
            Me.btnUpdate.Enabled = True
            Me.btnSave.Enabled = False
            Me.lstSubjects.Items.Clear()
        End If
    End Sub

    Private Sub btnUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdate.Click
        If Me.txtSubjCode.Text.Trim.Length <= 0 Then
            MsgBox("Missing Subject Code", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.txtSubjName.Text.Trim.Length <= 0 Then
            MsgBox("Missing Subject Name", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboSubjGroup.Text.Trim.Length <= 0 Then
            MsgBox("Missing Subject Group", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            recordExists = True
            checkIfRecordExistsOne()
            If recordExists = True Then
                MsgBox("Duplicate Found In the Database", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
                Exit Sub
            End If
            Dim result As MsgBoxResult = MsgBox("Update Record?", MsgBoxStyle.Question + MsgBoxStyle.ApplicationModal + MsgBoxStyle.YesNo, "Confirm Transaction")
            If result = MsgBoxResult.No Then
                Exit Sub
            End If
            queryType = "UPDATE"
            cmdSubjects.Connection = conn
            cmdSubjects.CommandType = CommandType.StoredProcedure
            cmdSubjects.CommandText = "sprocAcadSubjects"
            cmdSubjects.Parameters.Clear()
            cmdSubjects.Parameters.AddWithValue("@queryType", Me.queryType.Trim)
            cmdSubjects.Parameters.AddWithValue("@dateDone", Date.Now)
            cmdSubjects.Parameters.AddWithValue("@regBy", userName.Trim)
            cmdSubjects.Parameters.AddWithValue("@subCode", Me.txtSubjCode.Text.Trim)
            cmdSubjects.Parameters.AddWithValue("@subName", Me.txtSubjName.Text.Trim)
            cmdSubjects.Parameters.AddWithValue("@subGroup", Me.cboSubjGroup.Text.Trim)
            rec = cmdSubjects.ExecuteNonQuery
            If rec > 0 Then
                MsgBox("Record Updated!", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "SuccessFull Transaction")
            End If
            clearTexts()
            loadCombos()
            loadList()
            Me.btnUpdate.Enabled = False
            Me.btnSave.Enabled = True
            Me.txtSubjCode.Enabled = True
            Me.txtSubjName.Enabled = True
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            Me.queryType = Nothing
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub

    Private Sub DELETEToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DELETEToolStripMenuItem.Click
        If Me.lstSubjects.Items.Count <= 0 Then
            MsgBox("No items in the list", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.lstSubjects.SelectedItems.Count <= 0 Then
            MsgBox("Select the item to update", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.lstSubjects.SelectedItems.Count > 1 Then
            MsgBox("", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If
        If Me.lstSubjects.SelectedItems.Count = 1 Then
            Try
                If conn.State = ConnectionState.Closed Then
                    conn.Open()
                End If
                dbconnection()
                Dim result As MsgBoxResult = MsgBox("Delete Record?", MsgBoxStyle.Question + MsgBoxStyle.ApplicationModal + MsgBoxStyle.YesNo, "Confirm Transaction")
                If result = MsgBoxResult.No Then
                    Exit Sub
                End If
                queryType = "DELETE"
                cmdSubjects.Connection = conn
                cmdSubjects.CommandType = CommandType.StoredProcedure
                cmdSubjects.CommandText = "sprocAcadSubjects"
                cmdSubjects.Parameters.Clear()
                cmdSubjects.Parameters.AddWithValue("@queryType", Me.queryType.Trim)
                cmdSubjects.Parameters.AddWithValue("@dateDone", Date.Now)
                cmdSubjects.Parameters.AddWithValue("@regBy", userName.Trim)
                cmdSubjects.Parameters.AddWithValue("@subCode", Me.lstSubjects.SelectedItems(0).Text.Trim)
                cmdSubjects.Parameters.AddWithValue("@subName", Me.lstSubjects.SelectedItems(0).SubItems(1).Text.Trim)
                cmdSubjects.Parameters.AddWithValue("@subGroup", Me.lstSubjects.SelectedItems(0).SubItems(1).Text.Trim)
                rec = cmdSubjects.ExecuteNonQuery
                If rec > 0 Then
                    MsgBox("Record Deleted!", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "SuccessFull Transaction")
                End If
                clearTexts()
                loadCombos()
                loadList()
                Me.btnUpdate.Enabled = False
                Me.btnSave.Enabled = True
                Me.txtSubjCode.Enabled = True
                Me.txtSubjName.Enabled = True
            Catch ex As Exception
                MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Finally
                Me.queryType = Nothing
                If conn.State = ConnectionState.Open Then
                    conn.Close()
                End If
            End Try
        End If
    End Sub
End Class