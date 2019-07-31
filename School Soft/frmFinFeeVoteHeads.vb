Imports System.Data.SqlClient
Public Class frmFinFeeVoteHeads
    Dim cmdFeeVoteHeads As New SqlCommand
    Dim reader As SqlDataReader
    Dim rec As Integer = 0

    Private Sub frmFinFeeVoteHeads_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
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
        Me.cboVoteHeadName.Items.Clear()
        Me.cboVoteHeadName.Text = ""
        Me.cboVoteHeadName.SelectedIndex = -1

        Me.cmdFeeVoteHeads.Connection = conn
        Me.cmdFeeVoteHeads.CommandType = CommandType.Text
        Me.cmdFeeVoteHeads.CommandText = "SELECT DISTINCT voteHeadName FROM tblFinFeeVoteHeads ORDER BY voteHeadName"
        Me.cmdFeeVoteHeads.Parameters.Clear()
        reader = Me.cmdFeeVoteHeads.ExecuteReader
        While reader.Read
            Me.cboVoteHeadName.Items.Add(IIf(DBNull.Value.Equals(reader!voteHeadName), "", reader!voteHeadName))
        End While
        reader.Close()
    End Sub
    Private Sub loadList()
        Me.lstVoteHeads.Items.Clear()

        Me.cmdFeeVoteHeads.Connection = conn
        Me.cmdFeeVoteHeads.CommandType = CommandType.Text
        Me.cmdFeeVoteHeads.CommandText = "SELECT * FROM tblFinFeeVoteHeads ORDER BY priority"
        Me.cmdFeeVoteHeads.Parameters.Clear()
        reader = Me.cmdFeeVoteHeads.ExecuteReader
        While reader.Read
            li = Me.lstVoteHeads.Items.Add(IIf(DBNull.Value.Equals(reader!voteHeadName), "", reader!voteHeadName))
            li.SubItems.Add(IIf(DBNull.Value.Equals(reader!priority), "", reader!priority))
            li.Tag = IIf(DBNull.Value.Equals(reader!finVoteHeadId), "", reader!finVoteHeadId)
        End While
        reader.Close()
    End Sub
    Private Sub frmFinFeeVoteHeads_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint
        If Me.IsMdiChild Then
            Dim PnlLoc As New Point
            PnlLoc.X = CInt((Me.Width - Me.pnlFeeVoteHeads.Width) / 2)
            PnlLoc.Y = CInt((Me.Height - Me.pnlFeeVoteHeads.Height) / 2.5)
            Me.pnlFeeVoteHeads.Location = PnlLoc
        Else
            Me.pnlFeeVoteHeads.Dock = DockStyle.Fill
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
            loadCombos()
            loadList()
            Me.numVotePriority.Value = 1
            Me.btnSave.Enabled = True
            Me.btnUpdate.Enabled = False
            Me.cboVoteHeadName.Enabled = True
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
        If Me.cboVoteHeadName.Text.Trim.Length <= 0 Then
            MsgBox("VoteHead Name Is Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.numVotePriority.Text.Trim.Length <= 0 Then
            MsgBox("Priority Is Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()

            Dim voteHeadExists As Boolean = checkIfVoteHeadIsRegistered(Me.cboVoteHeadName.Text.Trim)
            If voteHeadExists = True Then
                MsgBox("Duplicate VoteHead Name Found in the database!", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
                Exit Sub
            End If

            Dim priorityExists As Boolean = checkIfPriorityIsUsed(Me.numVotePriority.Value)
            If priorityExists = True Then
                MsgBox("Priority is Already Used!", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
                Exit Sub
            End If

            Dim result As MsgBoxResult = MsgBox("Save Record?", MsgBoxStyle.Question + MsgBoxStyle.ApplicationModal + MsgBoxStyle.YesNo, "Confirm Transaction")
            If result = MsgBoxResult.No Then
                Exit Sub
            End If

            Me.cmdFeeVoteHeads.Connection = conn
            Me.cmdFeeVoteHeads.CommandType = CommandType.StoredProcedure
            Me.cmdFeeVoteHeads.CommandText = "sprocFinFeeVoteHeads"
            Me.cmdFeeVoteHeads.Parameters.Clear()
            Me.cmdFeeVoteHeads.Parameters.AddWithValue("@voteHeadName", Me.cboVoteHeadName.Text.Trim)
            Me.cmdFeeVoteHeads.Parameters.AddWithValue("@regBy", userName.Trim)
            Me.cmdFeeVoteHeads.Parameters.AddWithValue("@queryType", "INSERT")
            Me.cmdFeeVoteHeads.Parameters.AddWithValue("@priority", Me.numVotePriority.Value)
            rec = Me.cmdFeeVoteHeads.ExecuteNonQuery
            If rec > 0 Then
                MsgBox("Record Saved Successfully.", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "System Reply")
            End If
            Me.numVotePriority.Value = Me.numVotePriority.Value + 1
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
    Private Function checkIfPriorityIsUsedUpdate(ByVal priority As Integer, ByVal prevVoteHeadName As String)
        Dim priorityExists As Boolean = False
        Me.cmdFeeVoteHeads.Connection = conn
        Me.cmdFeeVoteHeads.CommandType = CommandType.Text
        Me.cmdFeeVoteHeads.CommandText = "SELECT * FROM tblFinFeeVoteHeads WHERE (priority=@priority) AND " & _
            vbNewLine & " (voteHeadName<>@voteHeadName)"
        Me.cmdFeeVoteHeads.Parameters.Clear()
        Me.cmdFeeVoteHeads.Parameters.AddWithValue("@priority", priority)
        Me.cmdFeeVoteHeads.Parameters.AddWithValue("@voteHeadName", prevVoteHeadName)
        reader = Me.cmdFeeVoteHeads.ExecuteReader
        If reader.HasRows = True Then
            priorityExists = True
        ElseIf reader.HasRows = False Then
            priorityExists = False
        End If
        reader.Close()
        Return priorityExists
    End Function
    Private Function checkIfPriorityIsUsed(ByVal priority As Integer)
        Dim priorityExists As Boolean = False
        Me.cmdFeeVoteHeads.Connection = conn
        Me.cmdFeeVoteHeads.CommandType = CommandType.Text
        Me.cmdFeeVoteHeads.CommandText = "SELECT * FROM tblFinFeeVoteHeads WHERE (priority=@priority)"
        Me.cmdFeeVoteHeads.Parameters.Clear()
        Me.cmdFeeVoteHeads.Parameters.AddWithValue("@priority", priority)
        reader = Me.cmdFeeVoteHeads.ExecuteReader
        If reader.HasRows = True Then
            priorityExists = True
        ElseIf reader.HasRows = False Then
            priorityExists = False
        End If
        reader.Close()
        Return priorityExists
    End Function
    Private Function checkIfVoteHeadIsRegistered(ByVal voteHeadName As String)
        Dim voteHeadExists As Boolean = False
        Me.cmdFeeVoteHeads.Connection = conn
        Me.cmdFeeVoteHeads.CommandType = CommandType.Text
        Me.cmdFeeVoteHeads.CommandText = "SELECT * FROM tblFinFeeVoteHeads WHERE (voteHeadName=@voteHeadName)"
        Me.cmdFeeVoteHeads.Parameters.Clear()
        Me.cmdFeeVoteHeads.Parameters.AddWithValue("@voteHeadName", voteHeadName.Trim)
        reader = Me.cmdFeeVoteHeads.ExecuteReader
        If reader.HasRows = True Then
            voteHeadExists = True
        ElseIf reader.HasRows = False Then
            voteHeadExists = False
        End If
        reader.Close()
        Return voteHeadExists
    End Function

    Private Sub UPDATEToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UPDATEToolStripMenuItem.Click
        If Me.lstVoteHeads.Items.Count <= 0 Then
            MsgBox("Missing Items In The List.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.lstVoteHeads.CheckedItems.Count <= 0 Then
            MsgBox("Missing Checked Items In The List.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.lstVoteHeads.CheckedItems.Count > 1 Then
            MsgBox("Edit one item At A Time.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            Me.cboVoteHeadName.Text = Me.lstVoteHeads.CheckedItems(0).Text.Trim
            Me.lstVoteHeads.Tag = Me.lstVoteHeads.CheckedItems(0).Text.Trim
            Me.numVotePriority.Value = Me.lstVoteHeads.CheckedItems(0).SubItems(1).Text.Trim
            Me.cboVoteHeadName.Tag = Me.lstVoteHeads.CheckedItems(0).Tag
            Me.lstVoteHeads.Items.Clear()
            Me.btnUpdate.Enabled = True
            Me.btnSave.Enabled = False
            'Me.cboVoteHeadName.Enabled = False
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub

    Private Sub btnUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdate.Click
        If Me.cboVoteHeadName.Text.Trim.Length <= 0 Then
            MsgBox("VoteHead Name Is Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.numVotePriority.Text.Trim.Length <= 0 Then
            MsgBox("Priority Is Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()

            'Dim voteHeadExists As Boolean = checkIfVoteHeadIsRegistered(Me.cboVoteHeadName.Text.Trim)
            'If voteHeadExists = True Then
            '    MsgBox("Duplicate VoteHead Name Found in the database!", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            '    Exit Sub
            'End If

            Dim priorityExists As Boolean = checkIfPriorityIsUsedUpdate(Me.numVotePriority.Value, Me.lstVoteHeads.Tag)
            If priorityExists = True Then
                MsgBox("Priority is Already Used!", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
                Exit Sub
            End If

            Dim result As MsgBoxResult = MsgBox("Update Record?", MsgBoxStyle.Question + MsgBoxStyle.ApplicationModal + MsgBoxStyle.YesNo, "Confirm Transaction")
            If result = MsgBoxResult.No Then
                Exit Sub
            End If

            Me.cmdFeeVoteHeads.Connection = conn
            Me.cmdFeeVoteHeads.CommandType = CommandType.StoredProcedure
            Me.cmdFeeVoteHeads.CommandText = "sprocFinFeeVoteHeads"
            Me.cmdFeeVoteHeads.Parameters.Clear()
            Me.cmdFeeVoteHeads.Parameters.AddWithValue("@voteHeadName", Me.cboVoteHeadName.Text.Trim)
            Me.cmdFeeVoteHeads.Parameters.AddWithValue("@prevvoteHeadName", Me.lstVoteHeads.Tag)
            Me.cmdFeeVoteHeads.Parameters.AddWithValue("@finVoteHeadId", Me.cboVoteHeadName.Tag)
            Me.cmdFeeVoteHeads.Parameters.AddWithValue("@regBy", userName.Trim)
            Me.cmdFeeVoteHeads.Parameters.AddWithValue("@queryType", "UPDATE")
            Me.cmdFeeVoteHeads.Parameters.AddWithValue("@priority", Me.numVotePriority.Value)
            rec = Me.cmdFeeVoteHeads.ExecuteNonQuery
            If rec > 0 Then
                MsgBox("Record Updated Successfully.", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "System Reply")
            End If
            loadCombos()
            loadList()
            Me.cboVoteHeadName.Tag = ""
            Me.lstVoteHeads.Tag = ""
            Me.numVotePriority.Value = 1
            Me.btnUpdate.Enabled = False
            Me.btnSave.Enabled = True
            Me.cboVoteHeadName.Enabled = True
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub

    Private Sub DELETEToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DELETEToolStripMenuItem.Click
        If Me.lstVoteHeads.Items.Count <= 0 Then
            MsgBox("Missing Items In The List.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.lstVoteHeads.CheckedItems.Count <= 0 Then
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

            For i = 0 To Me.lstVoteHeads.CheckedItems.Count - 1
                Me.cmdFeeVoteHeads.Connection = conn
                Me.cmdFeeVoteHeads.CommandType = CommandType.StoredProcedure
                Me.cmdFeeVoteHeads.CommandText = "sprocFinFeeVoteHeads"
                Me.cmdFeeVoteHeads.Parameters.Clear()
                Me.cmdFeeVoteHeads.Parameters.AddWithValue("@voteHeadName", Me.lstVoteHeads.CheckedItems(i).Text.Trim)
                Me.cmdFeeVoteHeads.Parameters.AddWithValue("@finVoteHeadId", Me.lstVoteHeads.CheckedItems(i).Tag)
                Me.cmdFeeVoteHeads.Parameters.AddWithValue("@regBy", userName.Trim)
                Me.cmdFeeVoteHeads.Parameters.AddWithValue("@queryType", "DELETE")
                Me.cmdFeeVoteHeads.Parameters.AddWithValue("@priority", Me.lstVoteHeads.CheckedItems(i).SubItems(1).Text.Trim)
                rec = Me.cmdFeeVoteHeads.ExecuteNonQuery
            Next
            If rec > 0 Then
                MsgBox("Record/s Deleted Successfully.", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "System Reply")
            End If
            loadCombos()
            loadList()
            Me.cboVoteHeadName.Tag = ""
            Me.lstVoteHeads.Tag = ""
            Me.numVotePriority.Value = 1
            Me.btnUpdate.Enabled = False
            Me.btnSave.Enabled = True
            Me.cboVoteHeadName.Enabled = True
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub
End Class