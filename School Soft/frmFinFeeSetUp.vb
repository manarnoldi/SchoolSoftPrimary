Imports System.Data.SqlClient
Public Class frmFinFeeSetUp
    Dim cmdFeeSetUp As New SqlCommand
    Dim reader As SqlDataReader
    Dim rec As Integer = 0
    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub
    Private Sub loadCombos()
        Me.cboTerm.Items.Clear()
        Me.cboTerm.Text = ""
        Me.cboTerm.SelectedIndex = -1

        Me.cboViewTerm.Items.Clear()
        Me.cboViewTerm.Text = ""
        Me.cboViewTerm.SelectedIndex = -1

        Me.cboYear.Items.Clear()
        Me.cboYear.Text = ""
        Me.cboYear.SelectedIndex = -1

        Me.cboViewYear.Items.Clear()
        Me.cboViewYear.Text = ""
        Me.cboViewYear.SelectedIndex = -1

        Me.cboFeeCat.Items.Clear()
        Me.cboFeeCat.Text = ""
        Me.cboFeeCat.SelectedIndex = -1

        Me.cboVwFeeCat.Items.Clear()
        Me.cboVwFeeCat.Text = ""
        Me.cboVwFeeCat.SelectedIndex = -1

        Me.cboVoteHeadName.Items.Clear()
        Me.cboVoteHeadName.Text = ""
        Me.cboVoteHeadName.SelectedIndex = -1

        Me.cmdFeeSetUp.Connection = conn
        Me.cmdFeeSetUp.CommandType = CommandType.Text
        Me.cmdFeeSetUp.CommandText = "SELECT DISTINCT termName FROM tblSchoolCalendar ORDER BY termName"
        Me.cmdFeeSetUp.Parameters.Clear()
        reader = Me.cmdFeeSetUp.ExecuteReader
        While reader.Read
            Me.cboTerm.Items.Add(IIf(DBNull.Value.Equals(reader!termName), "", reader!termName))
            Me.cboViewTerm.Items.Add(IIf(DBNull.Value.Equals(reader!termName), "", reader!termName))
        End While
        reader.Close()

        Me.cmdFeeSetUp.CommandText = "SELECT DISTINCT year FROM tblSchoolCalendar ORDER BY year"
        Me.cmdFeeSetUp.Parameters.Clear()
        reader = Me.cmdFeeSetUp.ExecuteReader
        While reader.Read
            Me.cboYear.Items.Add(IIf(DBNull.Value.Equals(reader!year), "", reader!year))
            Me.cboViewYear.Items.Add(IIf(DBNull.Value.Equals(reader!year), "", reader!year))
        End While
        reader.Close()

        Me.cmdFeeSetUp.CommandText = "SELECT DISTINCT feeCatName FROM tblFinFeeCategory ORDER BY feeCatName"
        Me.cmdFeeSetUp.Parameters.Clear()
        reader = Me.cmdFeeSetUp.ExecuteReader
        While reader.Read
            Me.cboFeeCat.Items.Add(IIf(DBNull.Value.Equals(reader!feeCatName), "", reader!feeCatName))
            Me.cboVwFeeCat.Items.Add(IIf(DBNull.Value.Equals(reader!feeCatName), "", reader!feeCatName))
        End While
        reader.Close()

        Me.cmdFeeSetUp.CommandText = "SELECT DISTINCT voteHeadName FROM tblFinFeeVoteHeads ORDER BY voteHeadName"
        Me.cmdFeeSetUp.Parameters.Clear()
        reader = Me.cmdFeeSetUp.ExecuteReader
        While reader.Read
            Me.cboVoteHeadName.Items.Add(IIf(DBNull.Value.Equals(reader!voteHeadName), "", reader!voteHeadName))
        End While
        reader.Close()
    End Sub
    Private Sub frmFinFeeSetUp_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            loadCombos()
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub

    Private Sub frmFinFeeSetUp_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint
        If Me.IsMdiChild Then
            Dim PnlLoc As New Point
            PnlLoc.X = CInt((Me.Width - Me.pnlFeeSetUp.Width) / 2)
            PnlLoc.Y = CInt((Me.Height - Me.pnlFeeSetUp.Height) / 2.5)
            Me.pnlFeeSetUp.Location = PnlLoc
        Else
            Me.pnlFeeSetUp.Dock = DockStyle.Fill
        End If
    End Sub

    Private Sub txtVoteAmount_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtVoteAmount.TextChanged
        If IsNumeric(Me.txtVoteAmount.Text.Trim) = False And Not (Me.txtVoteAmount.Text.Trim.Length <= 0) Then
            MsgBox("Non Numeric Values Detected.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Me.txtVoteAmount.Text = ""
            Exit Sub
        End If
    End Sub

    Private Sub cboYear_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboYear.SelectedIndexChanged
        
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            If Me.cboYear.Text.Trim.Length <= 0 Then
                Exit Sub
            ElseIf Me.cboYear.Text.Trim.Length > 0 Then
                Me.lstClasses.Items.Clear()

                Me.cmdFeeSetUp.Connection = conn
                Me.cmdFeeSetUp.CommandType = CommandType.Text
                Me.cmdFeeSetUp.CommandText = "SELECT * FROM  tblClasses WHERE (year=@year) ORDER BY className,stream"
                Me.cmdFeeSetUp.Parameters.Clear()
                Me.cmdFeeSetUp.Parameters.AddWithValue("@year", Me.cboYear.Text.Trim)
                reader = Me.cmdFeeSetUp.ExecuteReader
                While reader.Read
                    li = Me.lstClasses.Items.Add(IIf(DBNull.Value.Equals(reader!className), "", reader!className))
                    li.SubItems.Add(IIf(DBNull.Value.Equals(reader!stream), "", reader!stream))
                    li.SubItems.Add(IIf(DBNull.Value.Equals(reader!year), "", reader!year))
                End While
                reader.Close()
            End If
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub

    Private Sub lstClasses_ColumnClick(ByVal sender As Object, ByVal e As System.Windows.Forms.ColumnClickEventArgs) Handles lstClasses.ColumnClick
        If (e.Column() = 0) And (Me.lstClasses.Items.Count > 0) Then
            For Each Li As ListViewItem In Me.lstClasses.Items
                Li.Checked = Not (Li.Checked)
            Next
        End If
    End Sub

    'Private Sub lstVoteAmount_ColumnClick(ByVal sender As Object, ByVal e As System.Windows.Forms.ColumnClickEventArgs) Handles lstVoteAmount.ColumnClick
    '    If (e.Column() = 0) And (Me.lstVoteAmount.Items.Count > 0) Then
    '        For Each Li As ListViewItem In Me.lstVoteAmount.Items
    '            Li.Checked = Not (Li.Checked)
    '        Next
    '    End If
    'End Sub

    Private Sub cboViewYear_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboViewYear.SelectedIndexChanged
        If Me.cboViewYear.Text.Trim.Length <= 0 Then
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            Me.lstVwAssigned.Items.Clear()

            Me.cboVwClass.Items.Clear()
            Me.cboVwClass.Text = ""
            Me.cboVwClass.SelectedIndex = -1

            Me.cboVwStream.Items.Clear()
            Me.cboVwStream.Text = ""
            Me.cboVwStream.SelectedIndex = -1

            Me.cmdFeeSetUp.Connection = conn
            Me.cmdFeeSetUp.CommandType = CommandType.Text
            Me.cmdFeeSetUp.CommandText = "SELECT DISTINCT ClassName FROM  tblClasses WHERE (year=@year) ORDER BY className"
            Me.cmdFeeSetUp.Parameters.Clear()
            Me.cmdFeeSetUp.Parameters.AddWithValue("@year", Me.cboViewYear.Text.Trim)
            reader = Me.cmdFeeSetUp.ExecuteReader
            While reader.Read
                Me.cboVwClass.Items.Add(IIf(DBNull.Value.Equals(reader!className), "", reader!className))
            End While
            reader.Close()

            Me.cmdFeeSetUp.CommandText = "SELECT DISTINCT stream FROM  tblClasses WHERE (year=@year) ORDER BY stream"
            Me.cmdFeeSetUp.Parameters.Clear()
            Me.cmdFeeSetUp.Parameters.AddWithValue("@year", Me.cboViewYear.Text.Trim)
            reader = Me.cmdFeeSetUp.ExecuteReader
            While reader.Read
                Me.cboVwStream.Items.Add(IIf(DBNull.Value.Equals(reader!stream), "", reader!stream))
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

    Private Sub btnLoad_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLoad.Click
        Me.lstVwAssigned.Items.Clear()
        If Me.cboViewTerm.Text.Trim.Length <= 0 Then
            MsgBox("Select Term.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboViewYear.Text.Trim.Length <= 0 Then
            MsgBox("Select Year.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboVwClass.Text.Trim.Length <= 0 Then
            MsgBox("Select Class.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboVwStream.Text.Trim.Length <= 0 Then
            MsgBox("Select Stream.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboVwFeeCat.Text.Trim.Length <= 0 Then
            MsgBox("Select Fee Category.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()

            Me.cmdFeeSetUp.Connection = conn
            Me.cmdFeeSetUp.CommandType = CommandType.Text
            Me.cmdFeeSetUp.CommandText = "SELECT * FROM vwFinFeeSetUp WHERE (className=@className) AND (stream=@stream) " & _
                vbNewLine & " AND (year=@year) AND (feeCatName=@feeCatName) AND (termName=@termName) ORDER BY priority"
            Me.cmdFeeSetUp.Parameters.Clear()
            Me.cmdFeeSetUp.Parameters.AddWithValue("@className", Me.cboVwClass.Text.Trim)
            Me.cmdFeeSetUp.Parameters.AddWithValue("@stream", Me.cboVwStream.Text.Trim)
            Me.cmdFeeSetUp.Parameters.AddWithValue("@year", Me.cboViewYear.Text.Trim)
            Me.cmdFeeSetUp.Parameters.AddWithValue("@feeCatName", Me.cboVwFeeCat.Text.Trim)
            Me.cmdFeeSetUp.Parameters.AddWithValue("@termName", Me.cboViewTerm.Text.Trim)
            reader = Me.cmdFeeSetUp.ExecuteReader
            If reader.HasRows = True Then
                While reader.Read
                    li = Me.lstVwAssigned.Items.Add(IIf(DBNull.Value.Equals(reader!voteHeadName), "", reader!voteHeadName))
                    li.SubItems.Add(IIf(DBNull.Value.Equals(reader!priority), "", reader!priority))
                    li.SubItems.Add(IIf(DBNull.Value.Equals(reader!amount), "", reader!amount))
                    li.Tag = (IIf(DBNull.Value.Equals(reader!feeSetId), "", reader!feeSetId))
                End While
            ElseIf reader.HasRows = False Then
                MsgBox("No Record Found!", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "System Reply")
            End If
            reader.Close()
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub

    Private Sub btnAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAdd.Click
        If Me.cboVoteHeadName.Text.Trim.Length <= 0 Then
            MsgBox("Select VoteHead Name.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.txtVoteAmount.Text.Trim.Length <= 0 Then
            MsgBox("Enter VoteHead Amount.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            For i = 0 To Me.lstVoteAmount.Items.Count - 1
                If Me.cboVoteHeadName.Text.Trim = Me.lstVoteAmount.Items(i).Text.Trim Then
                    MsgBox("VoteHead Name Already In The List.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
                    Exit Sub
                End If
            Next
            li = Me.lstVoteAmount.Items.Add(Me.cboVoteHeadName.Text.Trim)
            li.SubItems.Add(Me.txtVoteAmount.Text.Trim)
            Me.txtVoteAmount.Text = ""
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub

    Private Sub ToolStripMenuItem3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem3.Click
        Me.Close()
    End Sub

    Private Sub CLOSEToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CLOSEToolStripMenuItem.Click
        Me.Close()
    End Sub

    Private Sub ToolStripMenuItem2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem2.Click
        If Me.lstVoteAmount.Items.Count <= 0 Then
            MsgBox("No item in the list to remove.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.lstVoteAmount.SelectedItems.Count <= 0 Then
            MsgBox("No Selected item/s in the list to remove.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If
        For i = 0 To Me.lstVoteAmount.SelectedItems.Count - 1
            Me.lstVoteAmount.SelectedItems(0).Remove()
        Next
    End Sub

    Private Sub btnClear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClear.Click
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            loadCombos()
            clearTexts()
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub
    Private Sub clearTexts()
        Me.lstVoteAmount.Items.Clear()
        Me.lstClasses.Items.Clear()
        Me.lstVwAssigned.Items.Clear()

        Me.txtVoteAmount.Text = ""
        Me.lstVwAssigned.Tag = Nothing
        Me.txtVoteAmount.Tag = Nothing

        Me.cboVwClass.Items.Clear()
        Me.cboVwClass.Text = ""
        Me.cboVwClass.SelectedIndex = -1

        Me.cboVwStream.Items.Clear()
        Me.cboVwStream.Text = ""
        Me.cboVwStream.SelectedIndex = -1

        Me.lstClasses.Enabled = True
        Me.lstVoteAmount.Enabled = True

        Me.cboTerm.Enabled = True
        Me.cboYear.Enabled = True
        Me.cboFeeCat.Enabled = True
        Me.cboVoteHeadName.Enabled = True

        Me.btnSave.Enabled = True
        Me.btnUpdate.Enabled = False
        Me.btnAdd.Enabled = True
    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        If Me.cboTerm.Text.Trim.Length <= 0 Then
            MsgBox("Select Term.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboFeeCat.Text.Trim.Length <= 0 Then
            MsgBox("Select Fee Category.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.lstClasses.Items.Count <= 0 Then
            MsgBox("No classes To Save Against.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.lstClasses.CheckedItems.Count <= 0 Then
            MsgBox("No checked classes To Save Against.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.lstVoteAmount.Items.Count <= 0 Then
            MsgBox("No vote Heads in the list.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            i = 0
            j = 0
            For i = 0 To Me.lstClasses.CheckedItems.Count - 1
                j = 0
                For j = 0 To Me.lstVoteAmount.Items.Count - 1
                    Dim recordExists As Boolean = checkIfRecordExists(Me.lstClasses.CheckedItems(i).Text.Trim, _
                    Me.lstClasses.CheckedItems(i).SubItems(1).Text.Trim, Me.lstClasses.CheckedItems(i).SubItems(2).Text.Trim,
                    Me.cboFeeCat.Text.Trim, Me.cboTerm.Text.Trim, Me.lstVoteAmount.Items(j).Text)
                    If recordExists = True Then
                        MsgBox("Duplicate Found in the database!", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
                        Exit Sub
                    End If
                Next
            Next

            Dim result As MsgBoxResult = MsgBox("Save Record/s?", MsgBoxStyle.Question + MsgBoxStyle.ApplicationModal + MsgBoxStyle.YesNo, "Confirm Transaction")
            If result = MsgBoxResult.No Then
                Exit Sub
            End If

            i = 0
            j = 0
            For i = 0 To Me.lstClasses.CheckedItems.Count - 1
                j = 0
                For j = 0 To Me.lstVoteAmount.Items.Count - 1
                    Me.cmdFeeSetUp.Connection = conn
                    Me.cmdFeeSetUp.CommandType = CommandType.StoredProcedure
                    Me.cmdFeeSetUp.CommandText = "sprocFinFeeSetUp"
                    Me.cmdFeeSetUp.Parameters.Clear()
                    Me.cmdFeeSetUp.Parameters.AddWithValue("@className", Me.lstClasses.CheckedItems(i).Text.Trim)
                    Me.cmdFeeSetUp.Parameters.AddWithValue("@stream", Me.lstClasses.CheckedItems(i).SubItems(1).Text.Trim)
                    Me.cmdFeeSetUp.Parameters.AddWithValue("@year", Me.lstClasses.CheckedItems(i).SubItems(2).Text.Trim)
                    Me.cmdFeeSetUp.Parameters.AddWithValue("@feeCatName", Me.cboFeeCat.Text.Trim)
                    Me.cmdFeeSetUp.Parameters.AddWithValue("@termName", Me.cboTerm.Text.Trim)
                    Me.cmdFeeSetUp.Parameters.AddWithValue("@voteHeadName", Me.lstVoteAmount.Items(j).Text.Trim)
                    Me.cmdFeeSetUp.Parameters.AddWithValue("@amount", Me.lstVoteAmount.Items(j).SubItems(1).Text.Trim)
                    Me.cmdFeeSetUp.Parameters.AddWithValue("@regBy", userName.Trim)
                    Me.cmdFeeSetUp.Parameters.AddWithValue("@queryType", "INSERT")
                    rec = rec + Me.cmdFeeSetUp.ExecuteNonQuery
                Next
            Next

            If rec > 0 Then
                MsgBox("Record/s Saved Successfully.", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "System Reply")
            End If
            loadCombos()
            clearTexts()
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub
    Private Function checkIfRecordExists(ByVal className As String, ByVal stream As String, ByVal year As Integer, ByVal feeCatName As String, ByVal termName As String, ByVal voteHeadName As String)
        Dim recordExists As Boolean = False
        Me.cmdFeeSetUp.Connection = conn
        Me.cmdFeeSetUp.CommandType = CommandType.Text
        Me.cmdFeeSetUp.CommandText = "SELECT * FROM vwFinFeeSetUp WHERE (className=@className) AND (stream=@stream) " & _
            vbNewLine & " AND (year=@year) AND (feeCatName=@feeCatName) AND (termName=@termName) AND (voteHeadName=@voteHeadName)"
        Me.cmdFeeSetUp.Parameters.Clear()
        Me.cmdFeeSetUp.Parameters.AddWithValue("@className", className.Trim)
        Me.cmdFeeSetUp.Parameters.AddWithValue("@stream", stream.Trim)
        Me.cmdFeeSetUp.Parameters.AddWithValue("@year", year)
        Me.cmdFeeSetUp.Parameters.AddWithValue("@feeCatName", feeCatName.Trim)
        Me.cmdFeeSetUp.Parameters.AddWithValue("@termName", termName.Trim)
        Me.cmdFeeSetUp.Parameters.AddWithValue("@voteHeadName", voteHeadName.Trim)
        reader = Me.cmdFeeSetUp.ExecuteReader
        If reader.HasRows = True Then
            recordExists = True
        ElseIf reader.HasRows = False Then
            recordExists = False
        End If
        reader.Close()
        Return recordExists
    End Function
    Private Function checkIfRecordExistsUpdate(ByVal className As String, ByVal stream As String, ByVal year As Integer, ByVal feeCatName As String, ByVal termName As String, ByVal voteHeadName As String, ByVal amount As Integer)
        Dim recordExists As Boolean = False
        Me.cmdFeeSetUp.Connection = conn
        Me.cmdFeeSetUp.CommandType = CommandType.Text
        Me.cmdFeeSetUp.CommandText = "SELECT * FROM vwFinFeeSetUp WHERE (className=@className) AND (stream=@stream) AND (amount=@amount) " & _
            vbNewLine & " AND (year=@year) AND (feeCatName=@feeCatName) AND (termName=@termName) AND (voteHeadName=@voteHeadName)"
        Me.cmdFeeSetUp.Parameters.Clear()
        Me.cmdFeeSetUp.Parameters.AddWithValue("@className", className.Trim)
        Me.cmdFeeSetUp.Parameters.AddWithValue("@stream", stream.Trim)
        Me.cmdFeeSetUp.Parameters.AddWithValue("@year", year)
        Me.cmdFeeSetUp.Parameters.AddWithValue("@feeCatName", feeCatName.Trim)
        Me.cmdFeeSetUp.Parameters.AddWithValue("@termName", termName.Trim)
        Me.cmdFeeSetUp.Parameters.AddWithValue("@voteHeadName", voteHeadName.Trim)
        Me.cmdFeeSetUp.Parameters.AddWithValue("@amount", amount)
        reader = Me.cmdFeeSetUp.ExecuteReader
        If reader.HasRows = True Then
            recordExists = True
        ElseIf reader.HasRows = False Then
            recordExists = False
        End If
        reader.Close()
        Return recordExists
    End Function

    Private Sub cboViewTerm_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboViewTerm.SelectedIndexChanged
        Me.lstVwAssigned.Items.Clear()
    End Sub

    Private Sub cboVwClass_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboVwClass.SelectedIndexChanged
        Me.lstVwAssigned.Items.Clear()
    End Sub

    Private Sub cboVwStream_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboVwStream.SelectedIndexChanged
        Me.lstVwAssigned.Items.Clear()
    End Sub

    Private Sub cboVwFeeCat_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboVwFeeCat.SelectedIndexChanged
        Me.lstVwAssigned.Items.Clear()
    End Sub

    Private Sub DELETEToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DELETEToolStripMenuItem.Click
        If Me.lstVwAssigned.Items.Count <= 0 Then
            MsgBox("No items in the list to delete.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.lstVwAssigned.CheckedItems.Count <= 0 Then
            MsgBox("No checked items in the list to delete.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
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

            For i = 0 To Me.lstVwAssigned.CheckedItems.Count - 1
                Me.cmdFeeSetUp.Connection = conn
                Me.cmdFeeSetUp.CommandType = CommandType.StoredProcedure
                Me.cmdFeeSetUp.CommandText = "sprocFinFeeSetUp"
                Me.cmdFeeSetUp.Parameters.Clear()
                Me.cmdFeeSetUp.Parameters.AddWithValue("@className", Me.cboVwClass.Text.Trim)
                Me.cmdFeeSetUp.Parameters.AddWithValue("@stream", Me.cboVwStream.Text.Trim)
                Me.cmdFeeSetUp.Parameters.AddWithValue("@year", Me.cboViewYear.Text.Trim)
                Me.cmdFeeSetUp.Parameters.AddWithValue("@feeCatName", Me.cboVwFeeCat.Text.Trim)
                Me.cmdFeeSetUp.Parameters.AddWithValue("@termName", Me.cboViewTerm.Text.Trim)
                Me.cmdFeeSetUp.Parameters.AddWithValue("@voteHeadName", Me.lstVwAssigned.CheckedItems(i).Text.Trim)
                Me.cmdFeeSetUp.Parameters.AddWithValue("@amount", Me.lstVwAssigned.CheckedItems(i).SubItems(2).Text.Trim)
                Me.cmdFeeSetUp.Parameters.AddWithValue("@regBy", userName.Trim)
                Me.cmdFeeSetUp.Parameters.AddWithValue("@feeSetId", Me.lstVwAssigned.CheckedItems(i).Tag)
                Me.cmdFeeSetUp.Parameters.AddWithValue("@queryType", "DELETE")
                rec = rec + Me.cmdFeeSetUp.ExecuteNonQuery
            Next
            If rec > 0 Then
                MsgBox("Record/s Deleted Successfully.", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "System Reply")
            End If
            loadCombos()
            clearTexts()
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub

    Private Sub UPDATEToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UPDATEToolStripMenuItem.Click
        If Me.lstVwAssigned.Items.Count <= 0 Then
            MsgBox("No items in the list to update.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.lstVwAssigned.CheckedItems.Count <= 0 Then
            MsgBox("No checked items in the list to update.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.lstVwAssigned.CheckedItems.Count > 1 Then
            MsgBox("Update One Item At A time.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If
        Try
            Me.cboFeeCat.Text = Me.cboVwFeeCat.Text.Trim
            Me.cboTerm.Text = Me.cboViewTerm.Text.Trim
            Me.cboVoteHeadName.Text = Me.lstVwAssigned.CheckedItems(0).Text.Trim
            Me.cboYear.Text = Me.cboViewYear.Text.Trim
            Me.txtVoteAmount.Text = Me.lstVwAssigned.CheckedItems(0).SubItems(2).Text.Trim
            Me.lstVwAssigned.Tag = Me.lstVwAssigned.CheckedItems(0).Tag
            Me.txtVoteAmount.Tag = Me.lstVwAssigned.CheckedItems(0).SubItems(2).Text.Trim

            Me.cboTerm.Enabled = False
            Me.cboYear.Enabled = False
            Me.cboFeeCat.Enabled = False
            Me.cboVoteHeadName.Enabled = False

            Me.btnUpdate.Enabled = True
            Me.btnSave.Enabled = False
            Me.btnAdd.Enabled = False

            Me.lstVoteAmount.Items.Clear()
            Me.lstClasses.Items.Clear()

            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            Me.cmdFeeSetUp.Connection = conn
            Me.cmdFeeSetUp.CommandType = CommandType.Text
            Me.cmdFeeSetUp.CommandText = "SELECT * FROM tblClasses WHERE (className=@className) AND (stream=@stream) " & _
                vbNewLine & " AND (year=@year) "
            Me.cmdFeeSetUp.Parameters.Clear()
            Me.cmdFeeSetUp.Parameters.AddWithValue("@className", Me.cboVwClass.Text.Trim)
            Me.cmdFeeSetUp.Parameters.AddWithValue("@stream", Me.cboVwStream.Text.Trim)
            Me.cmdFeeSetUp.Parameters.AddWithValue("@year", Me.cboViewYear.Text.Trim)
            reader = Me.cmdFeeSetUp.ExecuteReader
            While reader.Read
                li = Me.lstClasses.Items.Add(IIf(DBNull.Value.Equals(reader!className), "", reader!className))
                li.SubItems.Add(IIf(DBNull.Value.Equals(reader!stream), "", reader!stream))
                li.SubItems.Add(IIf(DBNull.Value.Equals(reader!Year), "", reader!Year))
                li.Tag = IIf(DBNull.Value.Equals(reader!classId), "", reader!classId)
            End While
            reader.Close()

            Me.lstClasses.Items(0).Checked = True

            Me.lstClasses.Enabled = False
            Me.lstVoteAmount.Enabled = False
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub

    Private Sub btnUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdate.Click
        If Me.cboTerm.Text.Trim.Length <= 0 Then
            MsgBox("Term Name Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboYear.Text.Trim.Length <= 0 Then
            MsgBox("Year Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboFeeCat.Text.Trim.Length <= 0 Then
            MsgBox("Fee Category Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboVoteHeadName.Text.Trim.Length <= 0 Then
            MsgBox("VoteHead Name Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.txtVoteAmount.Text.Trim.Length <= 0 Then
            MsgBox("Amount Is Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.lstClasses.CheckedItems.Count <= 0 Then
            MsgBox("Class Details Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            i = 0
            For i = 0 To Me.lstClasses.CheckedItems.Count - 1
                Dim recordExists As Boolean = checkIfRecordExistsUpdate(Me.lstClasses.CheckedItems(0).Text.Trim, _
                Me.lstClasses.CheckedItems(0).SubItems(1).Text.Trim, Me.lstClasses.CheckedItems(0).SubItems(2).Text.Trim,
                Me.cboFeeCat.Text.Trim, Me.cboTerm.Text.Trim, Me.cboVoteHeadName.Text.Trim, Me.txtVoteAmount.Text.Trim)
                    If recordExists = True Then
                        MsgBox("Duplicate Found in the database!", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
                        Exit Sub
                    End If
            Next

            Dim result As MsgBoxResult = MsgBox("Update Record?", MsgBoxStyle.Question + MsgBoxStyle.ApplicationModal + MsgBoxStyle.YesNo, "Confirm Transaction")
            If result = MsgBoxResult.No Then
                Exit Sub
            End If

            Me.cmdFeeSetUp.Connection = conn
            Me.cmdFeeSetUp.CommandType = CommandType.StoredProcedure
            Me.cmdFeeSetUp.CommandText = "sprocFinFeeSetUp"
            Me.cmdFeeSetUp.Parameters.Clear()
            Me.cmdFeeSetUp.Parameters.AddWithValue("@className", Me.lstClasses.CheckedItems(0).Text.Trim)
            Me.cmdFeeSetUp.Parameters.AddWithValue("@stream", Me.lstClasses.CheckedItems(0).SubItems(1).Text.Trim)
            Me.cmdFeeSetUp.Parameters.AddWithValue("@year", Me.lstClasses.CheckedItems(0).SubItems(2).Text.Trim)
            Me.cmdFeeSetUp.Parameters.AddWithValue("@feeCatName", Me.cboFeeCat.Text.Trim)
            Me.cmdFeeSetUp.Parameters.AddWithValue("@termName", Me.cboTerm.Text.Trim)
            Me.cmdFeeSetUp.Parameters.AddWithValue("@voteHeadName", Me.cboVoteHeadName.Text.Trim)
            Me.cmdFeeSetUp.Parameters.AddWithValue("@amount", Me.txtVoteAmount.Text.Trim)
            Me.cmdFeeSetUp.Parameters.AddWithValue("@regBy", userName.Trim)
            Me.cmdFeeSetUp.Parameters.AddWithValue("@feeSetId", Me.lstVwAssigned.Tag)
            Me.cmdFeeSetUp.Parameters.AddWithValue("@fromAmount", Me.txtVoteAmount.Tag)
            Me.cmdFeeSetUp.Parameters.AddWithValue("@queryType", "UPDATE")
            rec = rec + Me.cmdFeeSetUp.ExecuteNonQuery

            If rec > 0 Then
                MsgBox("Record/s Updated Successfully.", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "System Reply")
            End If
            loadCombos()
            clearTexts()
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub

    Private Sub cboVoteHeadName_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboVoteHeadName.SelectedIndexChanged
        Me.txtVoteAmount.Text = ""
    End Sub

    Private Sub lstClasses_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lstClasses.SelectedIndexChanged

    End Sub
End Class