Imports System.Data.SqlClient
Public Class frmAcadGrades
    Dim duplicateFound As Boolean = True
    Dim cmdGrades As New SqlCommand
    Dim reader As SqlDataReader
    Dim rec As Integer = 0
    Dim queryType As String = Nothing
    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub
    Private Sub frmAcadGrades_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint
        If Me.IsMdiChild Then
            Dim PnlLoc As New Point
            PnlLoc.X = CInt((Me.Width - Me.pnlAcadGrades.Width) / 2)
            PnlLoc.Y = CInt((Me.Height - Me.pnlAcadGrades.Height) / 2.5)
            Me.pnlAcadGrades.Location = PnlLoc
        Else
            Me.pnlAcadGrades.Dock = DockStyle.Fill
        End If
    End Sub
    Private Sub frmAcadGrades_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
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
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If
        dbconnection()
        Me.cboRemarkEng.Items.Clear()
        Me.cboRemarkEng.SelectedIndex = -1
        Me.cboRemarkEng.Text = ""

        Me.cboRemarkSwa.Items.Clear()
        Me.cboRemarkSwa.SelectedIndex = -1
        Me.cboRemarkSwa.Text = ""

        Me.cboClassName.Items.Clear()
        Me.cboClassName.SelectedIndex = -1
        Me.cboClassName.Text = ""

        cmdGrades.Connection = conn
        cmdGrades.CommandText = "SELECT DISTINCT remarkSwa FROM tblGrades WHERE (Status=1) ORDER BY remarkSwa"
        cmdGrades.CommandType = CommandType.Text
        cmdGrades.Parameters.Clear()
        reader = cmdGrades.ExecuteReader
        If reader.HasRows Then
            While reader.Read
                Me.cboRemarkSwa.Items.Add(IIf(DBNull.Value.Equals(reader!remarkSwa), "", reader!remarkSwa))
            End While
        End If
        reader.Close()

        cmdGrades.CommandText = "SELECT DISTINCT remarkEng FROM tblGrades WHERE (Status=1) ORDER BY remarkEng"
        cmdGrades.Parameters.Clear()
        reader = cmdGrades.ExecuteReader
        If reader.HasRows Then
            While reader.Read
                Me.cboRemarkEng.Items.Add(IIf(DBNull.Value.Equals(reader!remarkEng), "", reader!remarkEng))
            End While
        End If
        reader.Close()

        cmdGrades.CommandText = "SELECT DISTINCT className FROM tblClasses WHERE (Status=1) ORDER BY className"
        cmdGrades.Parameters.Clear()
        reader = cmdGrades.ExecuteReader
        If reader.HasRows Then
            While reader.Read
                Me.cboClassName.Items.Add(IIf(DBNull.Value.Equals(reader!className), "", reader!className))
            End While
        End If
        reader.Close()
    End Sub

    Private Sub loadList()
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If
        dbconnection()
        Me.lstClassNames.Items.Clear()
        Me.lstAcadGrades.Items.Clear()

        cmdGrades.Connection = conn
        cmdGrades.CommandText = "SELECT DISTINCT className FROM tblClasses WHERE (Status=1) ORDER BY className"
        cmdGrades.CommandType = CommandType.Text
        cmdGrades.Parameters.Clear()
        reader = cmdGrades.ExecuteReader
        If reader.HasRows Then
            While reader.Read
                li = Me.lstClassNames.Items.Add(IIf(DBNull.Value.Equals(reader!className), "", reader!className))
            End While
        End If
        reader.Close()
    End Sub

    Private Sub btnLoad_Click(sender As Object, e As EventArgs) Handles btnLoad.Click
        If Me.cboClassName.Text.Trim.Count <= 0 Then
            MsgBox("Select the class to view grades.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            Me.lstViewSavedGrades.Items.Clear()
            cmdGrades.Connection = conn
            cmdGrades.CommandText = "SELECT * FROM tblGrades WHERE (Status=1) AND (className=@className) ORDER BY upperMark DESC"
            cmdGrades.CommandType = CommandType.Text
            cmdGrades.Parameters.Clear()
            cmdGrades.Parameters.AddWithValue("@className", Me.cboClassName.Text.Trim)
            reader = cmdGrades.ExecuteReader
            If reader.HasRows Then
                While reader.Read
                    li = Me.lstViewSavedGrades.Items.Add(IIf(DBNull.Value.Equals(reader!gradeName), "", reader!gradeName))
                    li.SubItems.Add(IIf(DBNull.Value.Equals(reader!upperMark), "", reader!upperMark))
                    li.SubItems.Add(IIf(DBNull.Value.Equals(reader!lowerMark), "", reader!lowerMark))
                    li.SubItems.Add(IIf(DBNull.Value.Equals(reader!points), "", reader!points))
                    li.SubItems.Add(IIf(DBNull.Value.Equals(reader!remarkSwa), "", reader!remarkSwa))
                    li.SubItems.Add(IIf(DBNull.Value.Equals(reader!remarkEng), "", reader!remarkEng))
                    li.Tag = IIf(DBNull.Value.Equals(reader!gradeId), "", reader!gradeId)
                End While
            Else
                MsgBox("No record Found!", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
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

    Private Sub btnAdd_Click(sender As Object, e As EventArgs) Handles btnAdd.Click
        If Me.txtGradeName.Text.Trim.Length <= 0 Then
            MsgBox("Missing Grade Name", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.txtGradePoints.Text.Trim.Length <= 0 Then
            MsgBox("Missing Grade points", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.txtLowerMark.Text.Trim.Length <= 0 Then
            MsgBox("Missing Lower Marks", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.txtUpperMark.Text.Trim.Length <= 0 Then
            MsgBox("Missing Upper Marks", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboRemarkEng.Text.Trim.Length <= 0 Then
            MsgBox("Missing English Remark", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboRemarkSwa.Text.Trim.Length <= 0 Then
            MsgBox("Missing Swahili Remark", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If
        i = 0
        maxrec = Me.lstAcadGrades.Items.Count - 1
        For i = 0 To maxrec
            If Me.lstAcadGrades.Items(i).Text.Trim = Me.txtGradeName.Text.Trim Then
                MsgBox("Grade name already added", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
                Exit Sub
            End If
            If Me.lstAcadGrades.Items(i).SubItems(1).Text.Trim = Me.txtUpperMark.Text.Trim Then
                MsgBox("Upper Mark Already Added", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
                Exit Sub
            End If
            If Me.lstAcadGrades.Items(i).SubItems(2).Text.Trim = Me.txtLowerMark.Text.Trim Then
                MsgBox("Lower Mark Already Added", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
                Exit Sub
            End If
            If Me.lstAcadGrades.Items(i).SubItems(3).Text.Trim = Me.txtGradePoints.Text.Trim Then
                MsgBox("Points Already Added", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
                Exit Sub
            End If
        Next

        li = Me.lstAcadGrades.Items.Add(Me.txtGradeName.Text.Trim)
        li.SubItems.Add(Me.txtUpperMark.Text.Trim)
        li.SubItems.Add(Me.txtLowerMark.Text.Trim)
        li.SubItems.Add(Me.txtGradePoints.Text.Trim)
        li.SubItems.Add(Me.cboRemarkSwa.Text.Trim)
        li.SubItems.Add(Me.cboRemarkEng.Text.Trim)
    End Sub

    Private Sub btnRemove_Click(sender As Object, e As EventArgs) Handles btnRemove.Click
        If Me.lstAcadGrades.Items.Count = 0 Then
            MsgBox("No items in the list", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.lstAcadGrades.SelectedItems.Count <= 0 Then
            MsgBox("Select items to remove from list", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If
        i = 0
        For i = 0 To Me.lstAcadGrades.SelectedItems.Count - 1
            Me.lstAcadGrades.SelectedItems(0).Remove()
        Next
    End Sub

    Private Function CheckIfRecordExistsInTheDatabase(ByVal className As String, ByVal gradeName As String, ByVal upperMark As Double, ByVal lowerMark As Double,
                                                 ByVal points As Double) As Integer
        Dim returnValue As Integer = 0
        Me.cmdGrades.CommandType = CommandType.Text
        Me.cmdGrades.Connection = conn
        Me.cmdGrades.CommandText = "SELECT COUNT(gradeId) AS Count FROM tblGrades WHERE (status=1) AND (gradeName=@gradeName) AND (className=@className)"
        Me.cmdGrades.Parameters.Clear()
        Me.cmdGrades.Parameters.AddWithValue("@gradeName", gradeName)
        Me.cmdGrades.Parameters.AddWithValue("@className", className)
        reader = Me.cmdGrades.ExecuteReader
        While reader.Read
            If (IIf(DBNull.Value.Equals(reader!Count), "", reader!Count) > 0) Then
                returnValue = 1
            End If
        End While
        reader.Close()

        Me.cmdGrades.CommandText = "SELECT COUNT(gradeId) AS Count FROM tblGrades WHERE (status=1) AND (upperMark=@upperMark) AND (className=@className)"
        Me.cmdGrades.Parameters.Clear()
        Me.cmdGrades.Parameters.AddWithValue("@upperMark", upperMark)
        Me.cmdGrades.Parameters.AddWithValue("@className", className)
        reader = Me.cmdGrades.ExecuteReader
        While reader.Read
            If (IIf(DBNull.Value.Equals(reader!Count), "", reader!Count) > 0) Then
                returnValue = 2
            End If
        End While
        reader.Close()

        Me.cmdGrades.CommandText = "SELECT COUNT(gradeId) AS Count FROM tblGrades WHERE (status=1) AND (lowerMark=@lowerMark) AND (className=@className)"
        Me.cmdGrades.Parameters.Clear()
        Me.cmdGrades.Parameters.AddWithValue("@lowerMark", lowerMark)
        Me.cmdGrades.Parameters.AddWithValue("@className", className)
        reader = Me.cmdGrades.ExecuteReader
        While reader.Read
            If (IIf(DBNull.Value.Equals(reader!Count), "", reader!Count) > 0) Then
                returnValue = 3
            End If
        End While
        reader.Close()

        Me.cmdGrades.CommandText = "SELECT COUNT(gradeId) AS Count FROM tblGrades WHERE (status=1) AND (points=@points) AND (className=@className)"
        Me.cmdGrades.Parameters.Clear()
        Me.cmdGrades.Parameters.AddWithValue("@points", points)
        Me.cmdGrades.Parameters.AddWithValue("@className", className)
        reader = Me.cmdGrades.ExecuteReader
        While reader.Read
            If (IIf(DBNull.Value.Equals(reader!Count), "", reader!Count) > 0) Then
                returnValue = 4
            End If
        End While
        reader.Close()
        Return returnValue
    End Function

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        If Me.lstAcadGrades.Items.Count <= 0 Then
            MsgBox("Add items in the list before saving.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.lstClassNames.Items.Count <= 0 Then
            MsgBox("Register classes first before registering the grades.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.lstClassNames.CheckedItems.Count <= 0 Then
            MsgBox("Check the classes you want to save grades for.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If

        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            Dim recordExists As Integer = 0
            Me.Cursor = Cursors.WaitCursor

            For Each liClass As ListViewItem In lstClassNames.CheckedItems
                For Each liGrades As ListViewItem In lstAcadGrades.Items
                    recordExists = CheckIfRecordExistsInTheDatabase(liClass.Text.Trim, liGrades.Text.Trim, liGrades.SubItems(1).Text.Trim, liGrades.SubItems(2).Text.Trim,
                                                     liGrades.SubItems(3).Text.Trim)
                    If recordExists = 1 Then
                        MsgBox("Grade Name for Grade (" + liGrades.Text.Trim + ") for class (" & liClass.Text.Trim + ") is already registered.",
                               MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
                        Exit Sub
                    ElseIf recordExists = 2 Then
                        MsgBox("Upper Mark for Grade (" + liGrades.Text.Trim + ") for class (" & liClass.Text.Trim + ") is already registered.",
                             MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
                        Exit Sub
                    ElseIf recordExists = 3 Then
                        MsgBox("Lower Mark for Grade (" + liGrades.Text.Trim + ") for class (" & liClass.Text.Trim + ") is already registered.",
                             MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
                        Exit Sub
                    ElseIf recordExists = 4 Then
                        MsgBox("Points for Grade (" + liGrades.Text.Trim + ") for class (" & liClass.Text.Trim + ") is already registered.",
                             MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
                        Exit Sub
                    End If
                Next
            Next

            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            Dim result As MsgBoxResult = MsgBox("Save Record/s?", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.YesNo, "Confirm Transaction")
            If result = MsgBoxResult.No Then
                Exit Sub
            End If
            rec = 0
            Me.cmdGrades.Connection = conn
            Me.cmdGrades.CommandType = CommandType.StoredProcedure
            Me.cmdGrades.CommandText = "sprocGrades"
            For Each liClass As ListViewItem In lstClassNames.CheckedItems
                For Each liGrades As ListViewItem In lstAcadGrades.Items
                    Me.cmdGrades.Parameters.Clear()
                    Me.cmdGrades.Parameters.AddWithValue("@queryType", "INSERT")
                    Me.cmdGrades.Parameters.AddWithValue("@regBy", userName.Trim)
                    Me.cmdGrades.Parameters.AddWithValue("@gradeName", liGrades.Text.Trim)
                    Me.cmdGrades.Parameters.AddWithValue("@upperMark", liGrades.SubItems(1).Text.Trim)
                    Me.cmdGrades.Parameters.AddWithValue("@lowerMark", liGrades.SubItems(2).Text.Trim)
                    Me.cmdGrades.Parameters.AddWithValue("@points", liGrades.SubItems(3).Text.Trim)
                    Me.cmdGrades.Parameters.AddWithValue("@remarkSwa", liGrades.SubItems(4).Text.Trim)
                    Me.cmdGrades.Parameters.AddWithValue("@remarkEng", liGrades.SubItems(5).Text.Trim)
                    Me.cmdGrades.Parameters.AddWithValue("@className", liClass.Text.Trim)
                    Me.cmdGrades.Parameters.AddWithValue("@gradeId", 0)
                    rec = rec + Me.cmdGrades.ExecuteNonQuery
                Next
            Next
            If rec > 0 Then
                MsgBox("Record/s Saved", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "SuccessFull Transaction")
            End If
            Me.lstAcadGrades.Items.Clear()
            loadCombos()
            loadList()
            Me.txtGradeName.Text = ""
            Me.txtGradePoints.Text = ""
            Me.txtLowerMark.Text = ""
            Me.txtUpperMark.Text = ""
            Me.lstViewSavedGrades.Items.Clear()
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            Me.Cursor = Cursors.Default
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub

    Private Sub CLOSEToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CLOSEToolStripMenuItem.Click
        Me.Close()
    End Sub

    Private Sub DELETEToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DELETEToolStripMenuItem.Click
        If Me.lstViewSavedGrades.Items.Count <= 0 Then
            MsgBox("Select items in the list before deleting.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.lstViewSavedGrades.CheckedItems.Count <= 0 Then
            MsgBox("Check the items to be deleted.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()

            Dim result As MsgBoxResult = MsgBox("Delete Record/s?", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.YesNo, "Confirm Transaction")
            If result = MsgBoxResult.No Then
                Exit Sub
            End If
            rec = 0
            Me.cmdGrades.Connection = conn
            Me.cmdGrades.CommandType = CommandType.StoredProcedure
            Me.cmdGrades.CommandText = "sprocGrades"

            For Each liGradesVw As ListViewItem In lstViewSavedGrades.CheckedItems
                Me.cmdGrades.Parameters.Clear()
                Me.cmdGrades.Parameters.AddWithValue("@queryType", "DELETE")
                Me.cmdGrades.Parameters.AddWithValue("@regBy", userName.Trim)
                Me.cmdGrades.Parameters.AddWithValue("@gradeName", liGradesVw.Text.Trim)
                Me.cmdGrades.Parameters.AddWithValue("@upperMark", liGradesVw.SubItems(1).Text.Trim)
                Me.cmdGrades.Parameters.AddWithValue("@lowerMark", liGradesVw.SubItems(2).Text.Trim)
                Me.cmdGrades.Parameters.AddWithValue("@points", liGradesVw.SubItems(3).Text.Trim)
                Me.cmdGrades.Parameters.AddWithValue("@remarkSwa", liGradesVw.SubItems(4).Text.Trim)
                Me.cmdGrades.Parameters.AddWithValue("@remarkEng", liGradesVw.SubItems(5).Text.Trim)
                Me.cmdGrades.Parameters.AddWithValue("@className", Me.cboClassName.Text.Trim)
                Me.cmdGrades.Parameters.AddWithValue("@gradeId", liGradesVw.Tag)
                rec = rec + Me.cmdGrades.ExecuteNonQuery
            Next
            If rec > 0 Then
                MsgBox("Record/s Deleted Successfully", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "SuccessFull Transaction")
            End If
            Me.lstViewSavedGrades.Items.Clear()

        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            Me.Cursor = Cursors.Default
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub

    Private Sub UPDATEToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles UPDATEToolStripMenuItem.Click
        If Me.lstViewSavedGrades.Items.Count <= 0 Then
            MsgBox("Load items in the list before editing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.lstViewSavedGrades.CheckedItems.Count <= 0 Then
            MsgBox("Check the item to edit.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.lstViewSavedGrades.CheckedItems.Count > 1 Then
            MsgBox("Edit one item at a time.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If
        Try
            For Each liClassVw As ListViewItem In Me.lstClassNames.Items
                If cboClassName.Text.Trim = liClassVw.Text.Trim Then
                    liClassVw.Checked = True
                End If
            Next

            Me.txtGradeName.Text = Me.lstViewSavedGrades.CheckedItems(0).Text.Trim
            Me.txtGradePoints.Text = Me.lstViewSavedGrades.CheckedItems(0).SubItems(3).Text.Trim
            Me.txtLowerMark.Text = Me.lstViewSavedGrades.CheckedItems(0).SubItems(2).Text.Trim
            Me.txtUpperMark.Text = Me.lstViewSavedGrades.CheckedItems(0).SubItems(1).Text.Trim
            Me.cboRemarkSwa.Text = Me.lstViewSavedGrades.CheckedItems(0).SubItems(4).Text.Trim
            Me.cboRemarkEng.Text = Me.lstViewSavedGrades.CheckedItems(0).SubItems(5).Text.Trim
            Me.txtGradeName.Tag = Me.lstViewSavedGrades.CheckedItems(0).Tag

            Me.txtGradeName.Enabled = False
            Me.txtGradePoints.Enabled = False
            Me.txtLowerMark.Enabled = False
            Me.txtUpperMark.Enabled = False

            Me.lstClassNames.Enabled = False
            Me.lstAcadGrades.Enabled = False
            Me.btnSave.Enabled = False
            Me.btnUpdate.Enabled = True

        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
        End Try
    End Sub

    Private Sub btnUpdate_Click(sender As Object, e As EventArgs) Handles btnUpdate.Click
        If Me.lstClassNames.Items.Count <= 0 Then
            MsgBox("No class names are available in the list.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.lstClassNames.CheckedItems.Count <= 0 Then
            MsgBox("No class name matched the class being edited.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.lstClassNames.CheckedItems.Count > 1 Then
            MsgBox("More than one class selected for editing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.txtGradeName.Text.Trim.Length <= 0 Then
            MsgBox("Missing Grade Name to update.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.txtGradePoints.Text.Trim.Length <= 0 Then
            MsgBox("Missing Grade points to update.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.txtLowerMark.Text.Trim.Length <= 0 Then
            MsgBox("Missing Lower Marks to update.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.txtUpperMark.Text.Trim.Length <= 0 Then
            MsgBox("Missing Upper Marks to update.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboRemarkEng.Text.Trim.Length <= 0 Then
            MsgBox("Missing English Remark to update.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboRemarkSwa.Text.Trim.Length <= 0 Then
            MsgBox("Missing Swahili Remark to update.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.txtGradeName.Tag.ToString().Trim.Length <= 0 Then
            MsgBox("Reselect the grade being edited.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If

        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()

            Dim recordExists As Boolean = CheckIfRecordExistsInTheDatabaseEdit(Me.lstClassNames.CheckedItems(0).Text.Trim, Me.txtGradeName.Text.Trim,
                                                                               Me.txtUpperMark.Text.Trim, Me.txtLowerMark.Text.Trim, Me.txtGradePoints.Text.Trim,
                                                                               Me.cboRemarkSwa.Text.Trim, Me.cboRemarkEng.Text.Trim)
            If recordExists Then
                MsgBox("You have not edited the grade details.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
                Exit Sub
            End If

            Dim result As MsgBoxResult = MsgBox("Update Record?", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.YesNo, "Confirm Transaction")
            If result = MsgBoxResult.No Then
                Exit Sub
            End If
            rec = 0
            Me.cmdGrades.Connection = conn
            Me.cmdGrades.CommandType = CommandType.StoredProcedure
            Me.cmdGrades.CommandText = "sprocGrades"

            For Each liGradesVw As ListViewItem In lstViewSavedGrades.CheckedItems
                Me.cmdGrades.Parameters.Clear()
                Me.cmdGrades.Parameters.AddWithValue("@queryType", "UPDATE")
                Me.cmdGrades.Parameters.AddWithValue("@regBy", userName.Trim)
                Me.cmdGrades.Parameters.AddWithValue("@gradeName", Me.txtGradeName.Text.Trim)
                Me.cmdGrades.Parameters.AddWithValue("@upperMark", Me.txtUpperMark.Text.Trim)
                Me.cmdGrades.Parameters.AddWithValue("@lowerMark", Me.txtLowerMark.Text.Trim)
                Me.cmdGrades.Parameters.AddWithValue("@points", Me.txtGradePoints.Text.Trim)
                Me.cmdGrades.Parameters.AddWithValue("@remarkSwa", Me.cboRemarkSwa.Text.Trim)
                Me.cmdGrades.Parameters.AddWithValue("@remarkEng", Me.cboRemarkEng.Text.Trim)
                Me.cmdGrades.Parameters.AddWithValue("@className", Me.lstClassNames.CheckedItems(0).Text.Trim)
                Me.cmdGrades.Parameters.AddWithValue("@gradeId", Me.txtGradeName.Tag)
                rec = rec + Me.cmdGrades.ExecuteNonQuery
            Next
            If rec > 0 Then
                MsgBox("Record Updated Successfully", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "SuccessFull Transaction")
            End If
            clearTexts()
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            Me.Cursor = Cursors.Default
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub
    Private Sub clearTexts()
        Me.lstViewSavedGrades.Items.Clear()
        Me.lstClassNames.Items.Clear()
        Me.lstAcadGrades.Items.Clear()

        Me.txtGradeName.Text = ""
        Me.txtGradePoints.Text = ""
        Me.txtLowerMark.Text = ""
        Me.txtUpperMark.Text = ""
        Me.cboRemarkSwa.Text = ""
        Me.cboRemarkEng.Text = ""

        Me.txtGradeName.Enabled = True
        Me.txtGradePoints.Enabled = True
        Me.txtLowerMark.Enabled = True
        Me.txtUpperMark.Enabled = True

        Me.lstClassNames.Enabled = True
        Me.lstAcadGrades.Enabled = True
        Me.btnSave.Enabled = True
        Me.btnUpdate.Enabled = False

        loadCombos()
        loadList()
    End Sub
    Private Function CheckIfRecordExistsInTheDatabaseEdit(ByVal className As String, ByVal gradeName As String, ByVal upperMark As Double, ByVal lowerMark As Double,
                                                 ByVal points As Double, ByVal remarksKis As String, ByVal remarksEng As String) As Boolean
        Dim returnValue As Boolean = False
        Me.cmdGrades.CommandType = CommandType.Text
        Me.cmdGrades.Connection = conn
        Me.cmdGrades.CommandText = "SELECT COUNT(gradeId) AS Count FROM tblGrades WHERE (status=1) AND (gradeName=@gradeName) AND (className=@className)" +
            " AND (upperMark=@upperMark) AND (lowerMark=@lowerMark) AND (points=@points) AND (remarkSwa=@remarkSwa) AND (remarkEng=@remarkEng)"
        Me.cmdGrades.Parameters.Clear()
        Me.cmdGrades.Parameters.AddWithValue("@gradeName", gradeName)
        Me.cmdGrades.Parameters.AddWithValue("@className", className)
        Me.cmdGrades.Parameters.AddWithValue("@upperMark", upperMark)
        Me.cmdGrades.Parameters.AddWithValue("@lowerMark", lowerMark)
        Me.cmdGrades.Parameters.AddWithValue("@points", points)
        Me.cmdGrades.Parameters.AddWithValue("@remarkSwa", remarksKis)
        Me.cmdGrades.Parameters.AddWithValue("@remarkEng", remarksEng)
        reader = Me.cmdGrades.ExecuteReader
        While reader.Read
            If (IIf(DBNull.Value.Equals(reader!Count), "", reader!Count) > 0) Then
                returnValue = True
            End If
        End While
        reader.Close()
        Return returnValue
    End Function

    Private Sub btnClear_Click(sender As Object, e As EventArgs) Handles btnClear.Click
        clearTexts()
    End Sub
End Class