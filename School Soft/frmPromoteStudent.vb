Imports System.Data.SqlClient
Public Class frmPromoteStudent
    Dim queryType As String = Nothing
    Dim newYear As Integer = 0
    Dim cmdPromote As New SqlCommand
    Dim reader As SqlDataReader
    Dim rec As Integer = 0
    Dim recordOk As Boolean = False
    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub frmPromoteStudent_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
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

    Private Sub frmPromoteStudent_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint
        If Me.IsMdiChild Then
            Dim PnlLoc As New Point
            PnlLoc.X = CInt((Me.Width - Me.pnlPromoteStud.Width) / 2)
            PnlLoc.Y = CInt((Me.Height - Me.pnlPromoteStud.Height) / 2.5)
            Me.pnlPromoteStud.Location = PnlLoc
        Else
            Me.pnlPromoteStud.Dock = DockStyle.Fill
        End If
    End Sub
    Private Sub loadCombos()
        Me.cboNameFrom.Items.Clear()
        Me.cboNameTo.Items.Clear()
        Me.cboNameFrom.Text = ""
        Me.cboNameTo.Text = ""

        Me.cboStreamTo.Items.Clear()
        Me.cboStreamFrom.Items.Clear()
        Me.cboStreamFrom.Text = ""
        Me.cboStreamTo.Text = ""

        Me.cboYearFrom.Items.Clear()
        Me.cboYearTo.Items.Clear()
        Me.cboYearFrom.Text = ""
        Me.cboYearTo.Text = ""

        cmdPromote.Connection = conn
        cmdPromote.CommandType = CommandType.Text
        cmdPromote.CommandText = "SELECT DISTINCT className FROM tblClasses WHERE (status='True') ORDER BY ClassName"
        cmdPromote.Parameters.Clear()
        reader = cmdPromote.ExecuteReader
        If reader.HasRows Then
            While reader.Read
                Me.cboNameFrom.Items.Add(IIf(DBNull.Value.Equals(reader!className), "", reader!className))
                Me.cboNameTo.Items.Add(IIf(DBNull.Value.Equals(reader!className), "", reader!className))
            End While
        End If
        reader.Close()

        cmdPromote.CommandText = "SELECT DISTINCT stream FROM tblClasses WHERE (status='True') ORDER BY stream"
        cmdPromote.Parameters.Clear()
        reader = cmdPromote.ExecuteReader
        If reader.HasRows Then
            While reader.Read
                Me.cboStreamFrom.Items.Add(IIf(DBNull.Value.Equals(reader!stream), "", reader!stream))
                Me.cboStreamTo.Items.Add(IIf(DBNull.Value.Equals(reader!stream), "", reader!stream))
            End While
        End If
        reader.Close()

        cmdPromote.CommandText = "SELECT DISTINCT year FROM tblClasses WHERE (status='True') ORDER BY year"
        cmdPromote.Parameters.Clear()
        reader = cmdPromote.ExecuteReader
        If reader.HasRows Then
            While reader.Read
                Me.cboYearFrom.Items.Add(IIf(DBNull.Value.Equals(reader!year), "", reader!year))
                Me.cboYearTo.Items.Add(IIf(DBNull.Value.Equals(reader!year), "", reader!year))
            End While
        End If
        reader.Close()
    End Sub

    Private Sub btnLoad_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLoad.Click
        If Me.cboNameFrom.Text.Trim.Length <= 0 Then
            MsgBox("Missing Class Name", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboYearFrom.Text.Trim.Length <= 0 Then
            MsgBox("Missing Year", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboStreamFrom.Text.Trim.Length <= 0 Then
            MsgBox("Missing Stream", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            Me.lstClassFrom.Items.Clear()
            dbconnection()
            cmdPromote.Connection = conn
            cmdPromote.CommandType = CommandType.Text
            cmdPromote.CommandText = "SELECT * FROM vwStudClass WHERE (studStatus='True') AND (classStatus='True') AND " & _
                vbNewLine & "(classStudStatus='True') AND (className=@className) AND (stream=@stream) AND (year=@year) ORDER BY admNo"
            cmdPromote.Parameters.Clear()
            cmdPromote.Parameters.AddWithValue("@className", Me.cboNameFrom.Text.Trim)
            cmdPromote.Parameters.AddWithValue("@stream", Me.cboStreamFrom.Text.Trim)
            cmdPromote.Parameters.AddWithValue("@year", Me.cboYearFrom.Text.Trim)
            reader = cmdPromote.ExecuteReader
            If reader.HasRows = True Then
                While reader.Read
                    li = Me.lstClassFrom.Items.Add(IIf(DBNull.Value.Equals(reader!admNo), "", reader!admNo))
                    li.SubItems.Add(IIf(DBNull.Value.Equals(reader!FullName), "", reader!FullName))
                    li.SubItems.Add(IIf(DBNull.Value.Equals(reader!year), "", reader!year))
                    li.Tag = (IIf(DBNull.Value.Equals(reader!classStudId), "", reader!classStudId))
                End While
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

    Private Sub btnPromote_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPromote.Click
        If Me.lstClassTo.Items.Count = 0 Then
            MsgBox("No Students In The List To Promote", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboNameFrom.Text.Trim = Me.cboNameTo.Text.Trim Then
            MsgBox("Classes Should not be the Same", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboStreamFrom.Text.Trim <> Me.cboStreamTo.Text.Trim Then
            MsgBox("Streams Should be the Same", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboYearFrom.Text.Trim = Me.cboYearTo.Text.Trim Then
            MsgBox("Years Promoted To Should Not Be the same", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Not (CInt(Me.cboYearTo.Text.Trim) - CInt(Me.cboYearFrom.Text.Trim) = 1) Then
            MsgBox("Year Different should be one", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If
        i = 0
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            Dim result As MsgBoxResult = MsgBox("Update Record?", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.YesNo, "Confirm Transaction")
            If result = MsgBoxResult.No Then
                Exit Sub
            End If
            For i = 0 To Me.lstClassTo.Items.Count - 1
                queryType = "INSERT"
                cmdPromote.Connection = conn
                cmdPromote.CommandType = CommandType.StoredProcedure
                cmdPromote.CommandText = "sprocStudentPromote"
                cmdPromote.Parameters.Clear()
                cmdPromote.Parameters.AddWithValue("@dateOfReg", Date.Now)
                cmdPromote.Parameters.AddWithValue("@userName", userName.Trim)
                cmdPromote.Parameters.AddWithValue("@admNo", Me.lstClassTo.Items(i).Text.Trim)
                cmdPromote.Parameters.AddWithValue("@classNameFrom", Me.cboNameFrom.Text.Trim)
                cmdPromote.Parameters.AddWithValue("@streamFrom", Me.cboStreamFrom.Text.Trim)
                cmdPromote.Parameters.AddWithValue("@yearFrom", Me.cboYearFrom.Text.Trim)
                cmdPromote.Parameters.AddWithValue("@classNameTo", Me.cboNameTo.Text.Trim)
                cmdPromote.Parameters.AddWithValue("@streamTo", Me.cboStreamTo.Text.Trim)
                cmdPromote.Parameters.AddWithValue("@yearTo", Me.cboYearTo.Text.Trim)
                cmdPromote.Parameters.AddWithValue("@queryType", Me.queryType.Trim)
                cmdPromote.ExecuteNonQuery()
            Next
            If Me.lstClassTo.Items.Count = 1 Then
                MsgBox("Record Updated", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "SuccessFul Transaction")
            Else
                MsgBox("Records Updated", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "SuccessFul Transaction")
            End If
            clearTexts()
            loadCombos()
            Me.lstClassTo.Items.Clear()
            Me.lstClassFrom.Items.Clear()
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally

            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub
    Private Sub clearTexts()
        Me.cboYearTo.Items.Clear()
        Me.cboYearTo.SelectedIndex = -1
        Me.cboYearFrom.Items.Clear()
        Me.cboYearFrom.SelectedIndex = -1
        Me.cboStreamTo.Items.Clear()
        Me.cboStreamTo.SelectedIndex = -1
        Me.cboStreamFrom.Items.Clear()
        Me.cboStreamFrom.SelectedIndex = -1
        Me.cboNameTo.Items.Clear()
        Me.cboNameTo.SelectedIndex = -1
        Me.cboNameFrom.Items.Clear()
        Me.cboNameFrom.SelectedIndex = -1
    End Sub

    Private Sub btnView_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnView.Click
        If Me.cboNameTo.Text.Trim.Length <= 0 Then
            MsgBox("Missing Class Name", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboYearTo.Text.Trim.Length <= 0 Then
            MsgBox("Missing Year", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboStreamTo.Text.Trim.Length <= 0 Then
            MsgBox("Missing Stream", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            Me.lstClassTo.Items.Clear()
            dbconnection()
            cmdPromote.Connection = conn
            cmdPromote.CommandType = CommandType.Text
            cmdPromote.CommandText = "SELECT * FROM vwStudClass WHERE (studStatus='True') AND (classStatus='True') AND " & _
                vbNewLine & "(classStudStatus='True') AND (className=@className) AND (stream=@stream) AND (year=@year) ORDER BY admNo"
            cmdPromote.Parameters.Clear()
            cmdPromote.Parameters.AddWithValue("@className", Me.cboNameTo.Text.Trim)
            cmdPromote.Parameters.AddWithValue("@stream", Me.cboStreamTo.Text.Trim)
            cmdPromote.Parameters.AddWithValue("@year", Me.cboYearTo.Text.Trim)
            reader = cmdPromote.ExecuteReader
            If reader.HasRows = True Then
                While reader.Read
                    li = Me.lstClassTo.Items.Add(IIf(DBNull.Value.Equals(reader!admNo), "", reader!admNo))
                    li.SubItems.Add(IIf(DBNull.Value.Equals(reader!FullName), "", reader!FullName))
                    li.SubItems.Add(IIf(DBNull.Value.Equals(reader!year), "", reader!year))
                    li.Tag = (IIf(DBNull.Value.Equals(reader!classStudId), "", reader!classStudId))
                End While
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

    Private Sub cboStreamFrom_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboStreamFrom.SelectedIndexChanged, cboNameFrom.SelectedIndexChanged, cboYearFrom.SelectedIndexChanged
        Me.lstClassFrom.Items.Clear()
        Me.lstClassTo.Items.Clear()
    End Sub

    'Private Sub cboNameTo_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboNameTo.SelectedIndexChanged, cboStreamTo.SelectedIndexChanged, cboYearTo.SelectedIndexChanged
    '    Me.lstClassTo.Items.Clear()
    'End Sub

    Private Sub lstClassFrom_ColumnClick(ByVal sender As Object, ByVal e As System.Windows.Forms.ColumnClickEventArgs) Handles lstClassFrom.ColumnClick
        If (e.Column() = 0) And (Me.lstClassFrom.Items.Count > 0) Then
            For Each Li As ListViewItem In Me.lstClassFrom.Items
                Li.Checked = Not (Li.Checked)
            Next
        End If
    End Sub

    Private Sub btnOneFrom_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOneFrom.Click
        If Me.lstClassFrom.CheckedItems.Count = 1 Then
            i = 0
            For i = 0 To Me.lstClassTo.Items.Count - 1
                If Me.lstClassFrom.CheckedItems(0).Text.Trim = Me.lstClassTo.Items(i).Text.Trim Then
                    MsgBox("Cannot Add The Student. " & vbNewLine & "The Student Already added", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
                    Exit Sub
                End If
            Next

            li = Me.lstClassTo.Items.Add(Me.lstClassFrom.CheckedItems(0).Text.Trim)
            li.SubItems.Add(Me.lstClassFrom.CheckedItems(0).SubItems(1).Text.Trim)
            newYear = CInt(Me.lstClassFrom.CheckedItems(0).SubItems(2).Text.Trim) + 1
            li.SubItems.Add(newYear)
            Me.cboNameTo.SelectedIndex = -1
            Me.cboStreamTo.SelectedIndex = -1
            Me.cboYearTo.SelectedIndex = -1
        Else
            MsgBox("Check One Student Only To Add", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        End If
    End Sub

    Private Sub btnAllFrom_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAllFrom.Click
        Dim addStudent As Boolean = True
        If Me.lstClassFrom.CheckedItems.Count > 1 Then
            i = 0
            j = 0
            maxrec = Me.lstClassFrom.CheckedItems.Count - 1
            For i = 0 To maxrec
                j = 0
                addStudent = True
                maxrec1 = Me.lstClassTo.Items.Count - 1
                For j = 0 To maxrec1
                    If Me.lstClassFrom.CheckedItems(i).Text.Trim = Me.lstClassTo.Items(j).Text.Trim Then
                        MsgBox("Student " & Me.lstClassFrom.CheckedItems(i).SubItems(1).Text.Trim & " Is Already added", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
                        addStudent = False
                    End If
                Next
                If addStudent = True Then
                    li = Me.lstClassTo.Items.Add(Me.lstClassFrom.CheckedItems(i).Text.Trim)
                    li.SubItems.Add(Me.lstClassFrom.CheckedItems(i).SubItems(1).Text.Trim)
                    newYear = CInt(Me.lstClassFrom.CheckedItems(i).SubItems(2).Text.Trim) + 1
                    li.SubItems.Add(newYear)
                End If
                If Me.lstClassTo.Items.Count = 0 Then
                    li = Me.lstClassTo.Items.Add(Me.lstClassFrom.CheckedItems(0).Text.Trim)
                    li.SubItems.Add(Me.lstClassFrom.CheckedItems(0).SubItems(1).Text.Trim)
                    newYear = CInt(Me.lstClassFrom.CheckedItems(0).SubItems(2).Text.Trim) + 1
                    li.SubItems.Add(newYear)
                End If
            Next
            Me.cboNameTo.SelectedIndex = -1
            Me.cboStreamTo.SelectedIndex = -1
            Me.cboYearTo.SelectedIndex = -1
        Else
            MsgBox("Check More Than One Student Only To Add", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        End If

    End Sub

    Private Sub btnClearLists_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClearLists.Click
        Me.lstClassFrom.Items.Clear()
        Me.lstClassTo.Items.Clear()
    End Sub

    Private Sub btnRemove_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRemove.Click
        If Me.lstClassTo.SelectedItems.Count > 0 Then
            For i = 0 To Me.lstClassTo.SelectedItems.Count - 1
                Me.lstClassTo.SelectedItems(0).Remove()
            Next
        Else
            MsgBox("Select Items To Remove", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        End If
    End Sub

    Private Sub lstClassFrom_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lstClassFrom.SelectedIndexChanged

    End Sub

    Private Sub cboNameTo_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboNameTo.SelectedIndexChanged, cboStreamTo.SelectedIndexChanged, cboYearTo.SelectedIndexChanged
        If Not (Me.cboNameTo.Text.Trim.Length = 0) And Not (Me.cboStreamTo.Text.Trim.Length = 0) And Not (Me.cboYearTo.Text.Trim.Length = 0) Then
            Try
                If conn.State = ConnectionState.Closed Then
                    conn.Open()
                End If
                dbconnection()
                cmdPromote.CommandText = "SELECT * FROM tblClasses WHERE (status='True') AND (className=@className) " & _
                    vbNewLine & " AND (stream=@stream) AND (year=@year)"
                cmdPromote.CommandType = CommandType.Text
                cmdPromote.Connection = conn
                cmdPromote.Parameters.Clear()
                cmdPromote.Parameters.AddWithValue("@year", Me.cboYearTo.Text.Trim)
                cmdPromote.Parameters.AddWithValue("@className", Me.cboNameTo.Text.Trim)
                cmdPromote.Parameters.AddWithValue("@stream", Me.cboStreamTo.Text.Trim)
                reader = cmdPromote.ExecuteReader
                If reader.HasRows = True Then
                ElseIf reader.HasRows = False Then
                    MsgBox("The Class Is Not Registered." & vbNewLine & "Contact System Administrator.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
                    Me.cboNameTo.SelectedIndex = -1
                    Me.cboStreamTo.SelectedIndex = -1
                    Me.cboYearTo.SelectedIndex = -1
                    Me.lstClassTo.Items.Clear()
                End If
                reader.Close()
            Catch ex As Exception
                MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Finally
                If conn.State = ConnectionState.Open Then
                    conn.Close()
                End If
            End Try
        End If

    End Sub

    Private Sub CloseToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CloseToolStripMenuItem.Click
        Me.Close()
    End Sub

    Private Sub DeleteStudentToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DeleteStudentToolStripMenuItem.Click
        If Me.cboNameTo.Text.Trim.Length <= 0 Then
            MsgBox("Missing Class Name", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboYearTo.Text.Trim.Length <= 0 Then
            MsgBox("Missing Year", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboStreamTo.Text.Trim.Length <= 0 Then
            MsgBox("Missing Stream", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If
        If Me.lstClassTo.SelectedItems.Count = 1 Then
            Try
                If conn.State = ConnectionState.Closed Then
                    conn.Open()
                End If
                dbconnection()
                Dim result As MsgBoxResult = MsgBox("Delete Record?", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.YesNo, "Confirm Transaction")
                If result = MsgBoxResult.No Then
                    Exit Sub
                End If
                queryType = "DELETE"
                cmdPromote.Connection = conn
                cmdPromote.CommandType = CommandType.StoredProcedure
                cmdPromote.CommandText = "sprocStudentPromote"
                cmdPromote.Parameters.Clear()
                cmdPromote.Parameters.AddWithValue("@dateOfReg", Date.Now)
                cmdPromote.Parameters.AddWithValue("@userName", userName.Trim)
                cmdPromote.Parameters.AddWithValue("@admNo", Me.lstClassTo.SelectedItems(0).Text.Trim)
                cmdPromote.Parameters.AddWithValue("@classNameFrom", Me.cboNameFrom.Text.Trim)
                cmdPromote.Parameters.AddWithValue("@streamFrom", Me.cboStreamFrom.Text.Trim)
                cmdPromote.Parameters.AddWithValue("@yearFrom", Me.cboYearFrom.Text.Trim)
                cmdPromote.Parameters.AddWithValue("@classNameTo", Me.cboNameTo.Text.Trim)
                cmdPromote.Parameters.AddWithValue("@streamTo", Me.cboStreamTo.Text.Trim)
                cmdPromote.Parameters.AddWithValue("@yearTo", Me.cboYearTo.Text.Trim)
                cmdPromote.Parameters.AddWithValue("@queryType", Me.queryType.Trim)
                rec = cmdPromote.ExecuteNonQuery()
                If rec > 1 Then
                    MsgBox("Record Deleted", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "SuccessFul Transaction")
                End If
                clearTexts()
                loadCombos()
                Me.lstClassTo.Items.Clear()
                Me.lstClassFrom.Items.Clear()
            Catch ex As Exception
                MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Finally
                If conn.State = ConnectionState.Open Then
                    conn.Close()
                End If
            End Try
        End If
    End Sub
End Class