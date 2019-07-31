Imports System.Data.SqlClient
Public Class frmStudFormSchDetails
    Dim recordOk As Boolean = False
    Dim queryType As String = Nothing
    Dim recordExists As Boolean = True
    Dim recordNotOk As Boolean = False
    Dim reader As SqlDataReader
    Dim rec As Integer = 0
    Dim cmdFormer As New SqlCommand
    Private Sub frmStudFormSchDetails_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
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
    Private Sub clearTexts()
        Me.txtOutOf.Text = ""
        Me.txtTotalScore.Text = ""
        Me.txtAdmNoRead.Text = ""
        Me.txtDateOfAdm.Text = ""
        Me.txtMarksScored.Text = ""
        Me.txtStudName.Text = ""
        Me.cboClassName.Text = ""
        Me.cboOutOf.Text = ""
        Me.cboSchLevel.Text = ""
        Me.cboSchName.Text = ""
        Me.cboSubject.Text = ""
        Me.lstFormerSchool.Items.Clear()
        Me.btnSave.Enabled = True
        Me.btnUpdate.Enabled = False
    End Sub
    Private Sub frmStudFormSchDetails_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint
        If Me.IsMdiChild Then
            Dim PnlLoc As New Point
            PnlLoc.X = CInt((Me.Width - Me.pnlStudFormSchl.Width) / 2)
            PnlLoc.Y = CInt((Me.Height - Me.pnlStudFormSchl.Height) / 2.5)
            Me.pnlStudFormSchl.Location = PnlLoc
        Else
            Me.pnlStudFormSchl.Dock = DockStyle.Fill
        End If
    End Sub
    Private Sub loadCombos()
        Me.cboClassName.Items.Clear()
        Me.cboClassName.Text = ""
        Me.cboOutOf.Items.Clear()
        Me.cboOutOf.Text = ""
        Me.cboSchLevel.Items.Clear()
        Me.cboSchLevel.Text = ""
        Me.cboSchName.Items.Clear()
        Me.cboSchName.Text = ""
        Me.cboSubject.Items.Clear()
        Me.cboSubject.Text = ""

        cmdFormer.Connection = conn
        cmdFormer.CommandText = "SELECT DISTINCT schoolName FROM vwStudFormerDetails WHERE (formerStatus='True') AND " & _
            vbNewLine & " (studStatus='True') ORDER BY schoolName"
        cmdFormer.CommandType = CommandType.Text
        cmdFormer.Parameters.Clear()
        reader = cmdFormer.ExecuteReader
        If reader.HasRows Then
            While reader.Read
                Me.cboSchName.Items.Add(IIf(DBNull.Value.Equals(reader!schoolName), "", reader!schoolName))
            End While
        End If
        reader.Close()

        cmdFormer.CommandText = "SELECT DISTINCT className FROM vwStudFormerDetails WHERE (formerStatus='True') AND " & _
            vbNewLine & " (studStatus='True') ORDER BY className"
        cmdFormer.Parameters.Clear()
        reader = cmdFormer.ExecuteReader
        If reader.HasRows Then
            While reader.Read
                Me.cboClassName.Items.Add(IIf(DBNull.Value.Equals(reader!className), "", reader!className))
            End While
        End If
        reader.Close()

        cmdFormer.CommandText = "SELECT DISTINCT subject FROM vwStudFormerDetails WHERE (formerStatus='True') AND " & _
            vbNewLine & " (studStatus='True') ORDER BY subject"
        cmdFormer.Parameters.Clear()
        reader = cmdFormer.ExecuteReader
        If reader.HasRows Then
            While reader.Read
                Me.cboSubject.Items.Add(IIf(DBNull.Value.Equals(reader!subject), "", reader!subject))
            End While
        End If
        reader.Close()

        cmdFormer.CommandText = "SELECT DISTINCT outOf FROM vwStudFormerDetails WHERE (formerStatus='True') AND " & _
            vbNewLine & " (studStatus='True') ORDER BY outOf"
        cmdFormer.Parameters.Clear()
        reader = cmdFormer.ExecuteReader
        If reader.HasRows Then
            While reader.Read
                Me.cboOutOf.Items.Add(IIf(DBNull.Value.Equals(reader!outOf), "", reader!outOf))
            End While
        End If
        reader.Close()

        cmdFormer.CommandText = "SELECT DISTINCT schoolLevel FROM vwStudFormerDetails WHERE (formerStatus='True') AND " & _
            vbNewLine & " (studStatus='True') ORDER BY schoolLevel"
        cmdFormer.Parameters.Clear()
        reader = cmdFormer.ExecuteReader
        If reader.HasRows Then
            While reader.Read
                Me.cboSchLevel.Items.Add(IIf(DBNull.Value.Equals(reader!schoolLevel), "", reader!schoolLevel))
            End While
        End If
        reader.Close()

    End Sub

    Private Sub btnLoad_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLoad.Click
        If Me.txtStudNo.Text.Trim.Length <= 0 Then
            MsgBox("Missing Student No", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If
        Me.txtAdmNoRead.Text = ""
        Me.txtStudName.Text = ""
        Me.txtDateOfAdm.Text = ""
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            cmdFormer.Connection = conn
            cmdFormer.CommandText = "SELECT * FROM tblStudDetails WHERE (status='True') AND (admNo=@admNo)"
            cmdFormer.CommandType = CommandType.Text
            cmdFormer.Parameters.Clear()
            cmdFormer.Parameters.AddWithValue("@admNo", Me.txtStudNo.Text)
            reader = cmdFormer.ExecuteReader
            If reader.HasRows Then
                While reader.Read
                    stringDate(IIf(DBNull.Value.Equals(reader!dateOfAdmission), "", reader!dateOfAdmission))
                    Me.txtAdmNoRead.Text = IIf(DBNull.Value.Equals(reader!admNo), "", reader!admNo)
                    Me.txtStudName.Text = IIf(DBNull.Value.Equals(reader!FullName), "", reader!FullName)
                    Me.txtDateOfAdm.Text = Datestring
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

    Private Sub txtStudNo_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtStudNo.TextChanged
        Me.txtAdmNoRead.Text = ""
        Me.txtStudName.Text = ""
        Me.txtDateOfAdm.Text = ""
        Me.lstFormerSchool.Items.Clear()
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub btnAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAdd.Click
        If Me.cboSchName.Text.Trim.Length <= 0 Then
            MsgBox("Missing School Name", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboClassName.Text.Trim.Length <= 0 Then
            MsgBox("Missing class Name", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboSchLevel.Text.Trim.Length <= 0 Then
            MsgBox("Missing school level", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.txtAdmNoRead.Text.Trim.Length <= 0 Then
            MsgBox("Missing Admission Number", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.txtStudName.Text.Trim.Length <= 0 Then
            MsgBox("Missing Student Name", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If
        maxrec = Me.lstFormerSchool.Items.Count - 1
        recordNotOk = False
        If maxrec >= 0 Then
            For i = 0 To maxrec
                If (Me.txtAdmNoRead.Text.Trim = Me.lstFormerSchool.Items(i).Text.Trim) And ( _
                    Me.txtStudName.Text.Trim = Me.lstFormerSchool.Items(i).SubItems(1).Text.Trim) And ( _
                    Me.cboSchName.Text.Trim = Me.lstFormerSchool.Items(i).SubItems(2).Text.Trim) And ( _
                    Me.cboClassName.Text.Trim = Me.lstFormerSchool.Items(i).SubItems(3).Text.Trim) And ( _
                    Me.cboSubject.Text.Trim = Me.lstFormerSchool.Items(i).SubItems(4).Text.Trim) And ( _
                    Me.cboSchLevel.Text.Trim = Me.lstFormerSchool.Items(i).SubItems(7).Text.Trim) Then
                    recordNotOk = True
                End If
            Next
            If recordNotOk = True Then
                MsgBox("Record Exists", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
                Exit Sub
            End If
        End If
        li = Me.lstFormerSchool.Items.Add(Me.txtAdmNoRead.Text.Trim)
        li.SubItems.Add(Me.txtStudName.Text.Trim)
        li.SubItems.Add(Me.cboSchName.Text.Trim)
        li.SubItems.Add(Me.cboClassName.Text.Trim)
        li.SubItems.Add(Me.cboSubject.Text.Trim)
        li.SubItems.Add(Me.txtMarksScored.Text.Trim)
        li.SubItems.Add(Me.cboOutOf.Text.Trim)
        li.SubItems.Add(Me.cboSchLevel.Text.Trim)
    End Sub

    Private Sub btnRemove_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRemove.Click
        If Me.lstFormerSchool.SelectedItems.Count = 1 Then
            Me.lstFormerSchool.SelectedItems(0).Remove()
        End If
    End Sub

    Private Sub btnView_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnView.Click
        If Me.txtStudNo.Text.Trim.Length <= 0 Then
            MsgBox("Missing Student No", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If
        Try
            clearTexts()
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            cmdFormer.Connection = conn
            cmdFormer.CommandText = "SELECT * FROM vwStudFormerDetails WHERE (formerStatus='True') AND " & _
                vbNewLine & " (studStatus='True') AND (admNo=@admNo) ORDER BY className,subject"
            cmdFormer.CommandType = CommandType.Text
            cmdFormer.Parameters.Clear()
            cmdFormer.Parameters.AddWithValue("@admNo", Me.txtStudNo.Text.Trim)
            reader = cmdFormer.ExecuteReader

            If reader.HasRows Then
                While reader.Read
                    li = Me.lstFormerSchool.Items.Add(IIf(DBNull.Value.Equals(reader!admNo), "", reader!admNo))
                    li.SubItems.Add(IIf(DBNull.Value.Equals(reader!FullName), "", reader!FullName))
                    li.SubItems.Add(IIf(DBNull.Value.Equals(reader!schoolName), "", reader!schoolName))
                    li.SubItems.Add(IIf(DBNull.Value.Equals(reader!className), "", reader!className))
                    li.SubItems.Add(IIf(DBNull.Value.Equals(reader!subject), "", reader!subject))
                    li.SubItems.Add(IIf(DBNull.Value.Equals(reader!marks), "", reader!marks))
                    li.SubItems.Add(IIf(DBNull.Value.Equals(reader!outOf), "", reader!outOf))
                    li.SubItems.Add(IIf(DBNull.Value.Equals(reader!schoolLevel), "", reader!schoolLevel))
                    Me.txtOutOf.Text = IIf(DBNull.Value.Equals(reader!totalMark), "", reader!totalMark)
                    Me.txtTotalScore.Text = IIf(DBNull.Value.Equals(reader!totalScore), "", reader!totalScore)
                End While
            Else
                MsgBox("No records Found For the student with the admission Number.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
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

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click

        If Me.lstFormerSchool.Items.Count <= 0 Then
            MsgBox("No items in the list to save", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Missing Records")
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()

            queryType = "INSERT"
            Dim result As MsgBoxResult = MsgBox("Save Record?", MsgBoxStyle.Question + MsgBoxStyle.ApplicationModal + MsgBoxStyle.YesNo, "Confirm Transaction")
            If result = MsgBoxResult.Yes Then
                For i = 0 To Me.lstFormerSchool.Items.Count - 1
                    recordOk = False
                    checkRecord()
                    If recordOk = True Then
                        cmdFormer.Connection = conn
                        cmdFormer.CommandText = "sprocStudentFormer"
                        cmdFormer.CommandType = CommandType.StoredProcedure
                        cmdFormer.Parameters.Clear()
                        cmdFormer.Parameters.AddWithValue("@admNo", Me.lstFormerSchool.Items(i).Text.Trim)
                        cmdFormer.Parameters.AddWithValue("@dateOfReg", Date.Now)
                        cmdFormer.Parameters.AddWithValue("@schoolName", Me.lstFormerSchool.Items(i).SubItems(2).Text.Trim)
                        cmdFormer.Parameters.AddWithValue("@schoolLevel", Me.lstFormerSchool.Items(i).SubItems(7).Text.Trim)
                        cmdFormer.Parameters.AddWithValue("@className", Me.lstFormerSchool.Items(i).SubItems(3).Text.Trim)
                        cmdFormer.Parameters.AddWithValue("@subject", Me.lstFormerSchool.Items(i).SubItems(4).Text.Trim)
                        cmdFormer.Parameters.AddWithValue("@marks", Me.lstFormerSchool.Items(i).SubItems(5).Text.Trim)
                        cmdFormer.Parameters.AddWithValue("@outOf", Me.lstFormerSchool.Items(i).SubItems(6).Text.Trim)
                        cmdFormer.Parameters.AddWithValue("@queryType", queryType.Trim)
                        cmdFormer.Parameters.AddWithValue("@regBy", userName)
                        rec += cmdFormer.ExecuteNonQuery
                    End If
                Next
                If rec > 0 Then
                    MsgBox("Record Saved!", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "SuccessFul Transaction")
                End If
            End If
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            Me.txtStudNo.Text = ""
            clearTexts()
            loadCombos()
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub
    Private Sub checkRecord()
        recordOk = False
        cmdFormer.Connection = conn
        cmdFormer.CommandText = "SELECT * FROM vwStudFormerDetails WHERE (formerStatus='True') AND " & _
            vbNewLine & " (studStatus='True') AND (admNo=@admNo) AND (SchoolName=@SchoolName) AND (className=@className) " & _
            vbNewLine & " AND (subject=@subject) AND (schoolLevel=@schoolLevel)"
        cmdFormer.CommandType = CommandType.Text
        cmdFormer.Parameters.Clear()
        cmdFormer.Parameters.AddWithValue("@admNo", Me.lstFormerSchool.Items(i).Text.Trim)
        cmdFormer.Parameters.AddWithValue("@schoolName", Me.lstFormerSchool.Items(i).SubItems(2).Text.Trim)
        cmdFormer.Parameters.AddWithValue("@schoolLevel", Me.lstFormerSchool.Items(i).SubItems(7).Text.Trim)
        cmdFormer.Parameters.AddWithValue("@className", Me.lstFormerSchool.Items(i).SubItems(3).Text.Trim)
        cmdFormer.Parameters.AddWithValue("@subject", Me.lstFormerSchool.Items(i).SubItems(4).Text.Trim)
        reader = cmdFormer.ExecuteReader
        If reader.HasRows Then
            recordOk = False
            reader.Close()
            Exit Sub
        End If
        reader.Close()
        recordOk = True
    End Sub

    Private Sub CLOSEToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CLOSEToolStripMenuItem.Click
        Me.Close()
    End Sub

    Private Sub EDITToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EDITToolStripMenuItem.Click
        If Me.lstFormerSchool.SelectedItems.Count = 1 Then

            Me.txtStudNo.Enabled = False
            Me.cboClassName.Enabled = False
            Me.cboSchLevel.Enabled = False
            Me.cboSchName.Enabled = False
            Me.cboSubject.Enabled = False
            Me.btnUpdate.Enabled = True
            Me.btnSave.Enabled = False

            Me.txtAdmNoRead.Text = Me.lstFormerSchool.SelectedItems(0).Text.Trim
            Me.txtMarksScored.Text = Me.lstFormerSchool.SelectedItems(0).SubItems(5).Text.Trim
            Me.txtStudName.Text = Me.lstFormerSchool.SelectedItems(0).SubItems(1).Text.Trim
            Me.txtStudNo.Text = Me.lstFormerSchool.SelectedItems(0).Text.Trim
            Me.cboSchName.Text = Me.lstFormerSchool.SelectedItems(0).SubItems(2).Text.Trim
            Me.cboClassName.Text = Me.lstFormerSchool.SelectedItems(0).SubItems(3).Text.Trim
            Me.cboSubject.Text = Me.lstFormerSchool.SelectedItems(0).SubItems(4).Text.Trim
            Me.cboOutOf.Text = Me.lstFormerSchool.SelectedItems(0).SubItems(6).Text.Trim
            Me.cboSchLevel.Text = Me.lstFormerSchool.SelectedItems(0).SubItems(7).Text.Trim

        End If
        Me.lstFormerSchool.Items.Clear()
    End Sub

    Private Sub txtMarksScored_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtMarksScored.TextChanged
        If Me.txtMarksScored.Text.Trim.Length <= 0 Then
            Exit Sub
        End If
        If IsNumeric(Me.txtMarksScored.Text) = False Then
            MsgBox("Non Numeric Values Detected", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Me.txtMarksScored.Text = ""
            Exit Sub
        End If
    End Sub
    Private Sub cboOutOf_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboOutOf.SelectedIndexChanged
        If Me.cboOutOf.Text.Trim.Length <= 0 Then
            Exit Sub
        End If
        If IsNumeric(Me.cboOutOf.Text) = False Then
            MsgBox("Non Numeric Values Detected", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Me.cboOutOf.Text = ""
            Exit Sub
        End If
    End Sub

    Private Sub btnUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdate.Click
        If Me.txtMarksScored.Text.Trim.Length <= 0 Then
            MsgBox("No Marks Entered To Update", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Missing Records")
            Exit Sub
        ElseIf Me.cboOutOf.Text.Trim.Length <= 0 Then
            MsgBox("Enter Mark out of", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Missing Records")
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            queryType = "UPDATE"
            recordOk = False
            checkRecordOne()
            If recordOk = False Then
                MsgBox("Record Already Exists", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Missing Records")
                Exit Sub
            End If
            Dim result As MsgBoxResult = MsgBox("Update Record?", MsgBoxStyle.Question + MsgBoxStyle.ApplicationModal + MsgBoxStyle.YesNo, "Confirm Transaction")
            If result = MsgBoxResult.Yes Then
                cmdFormer.Connection = conn
                cmdFormer.CommandText = "sprocStudentFormer"
                cmdFormer.CommandType = CommandType.StoredProcedure
                cmdFormer.Parameters.Clear()
                cmdFormer.Parameters.AddWithValue("@admNo", Me.txtAdmNoRead.Text.Trim)
                cmdFormer.Parameters.AddWithValue("@schoolName", Me.cboSchName.Text.Trim)
                cmdFormer.Parameters.AddWithValue("@schoolLevel", Me.cboSchLevel.Text.Trim)
                cmdFormer.Parameters.AddWithValue("@className", Me.cboClassName.Text.Trim)
                cmdFormer.Parameters.AddWithValue("@subject", Me.cboSubject.Text.Trim)
                cmdFormer.Parameters.AddWithValue("@marks", Me.txtMarksScored.Text.Trim)
                cmdFormer.Parameters.AddWithValue("@outOf", Me.cboOutOf.Text.Trim)
                cmdFormer.Parameters.AddWithValue("@dateOfReg", Date.Now)
                cmdFormer.Parameters.AddWithValue("@regBy", userName)
                cmdFormer.Parameters.AddWithValue("@queryType", queryType)
                rec = cmdFormer.ExecuteNonQuery
                If rec > 0 Then
                    MsgBox("Record Updated", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "SuccessFul Transaction")
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            clearTexts()
            Me.txtStudNo.Enabled = True
            Me.cboClassName.Enabled = True
            Me.cboSchLevel.Enabled = True
            Me.cboSchName.Enabled = True
            Me.cboSubject.Enabled = True
            Me.btnUpdate.Enabled = False
            Me.btnSave.Enabled = True
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub
    Private Sub checkRecordOne()
        recordOk = False
        cmdFormer.Connection = conn
        cmdFormer.CommandText = "SELECT * FROM vwStudFormerDetails WHERE (formerStatus='True') AND " & _
            vbNewLine & " (studStatus='True') AND (admNo=@admNo) AND (SchoolName=@SchoolName) AND (className=@className) " & _
            vbNewLine & " AND (subject=@subject) AND (schoolLevel=@schoolLevel) AND (marks=@marks) AND (outOf=@outOf)"
        cmdFormer.CommandType = CommandType.Text
        cmdFormer.Parameters.Clear()
        cmdFormer.Parameters.AddWithValue("@admNo", Me.txtAdmNoRead.Text.Trim)
        cmdFormer.Parameters.AddWithValue("@schoolName", Me.cboSchName.Text.Trim)
        cmdFormer.Parameters.AddWithValue("@schoolLevel", Me.cboSchLevel.Text.Trim)
        cmdFormer.Parameters.AddWithValue("@className", Me.cboClassName.Text.Trim)
        cmdFormer.Parameters.AddWithValue("@subject", Me.cboSubject.Text.Trim)
        cmdFormer.Parameters.AddWithValue("@marks", Me.txtMarksScored.Text.Trim)
        cmdFormer.Parameters.AddWithValue("@outOf", Me.cboOutOf.Text.Trim)
        reader = cmdFormer.ExecuteReader
        If reader.HasRows Then
            recordOk = False
            reader.Close()
            Exit Sub
        End If
        reader.Close()
        recordOk = True
    End Sub

    Private Sub DELETEToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DELETEToolStripMenuItem.Click
        If Me.lstFormerSchool.SelectedItems.Count = 1 Then
            Try
                If conn.State = ConnectionState.Closed Then
                    conn.Open()
                End If
                dbconnection()
                queryType = "DELETE"
                Dim result As MsgBoxResult = MsgBox("Delete Record?", MsgBoxStyle.Question + MsgBoxStyle.ApplicationModal + MsgBoxStyle.YesNo, "Confirm Transaction")
                If result = MsgBoxResult.Yes Then
                    cmdFormer.Connection = conn
                    cmdFormer.CommandText = "sprocStudentFormer"
                    cmdFormer.CommandType = CommandType.StoredProcedure
                    cmdFormer.Parameters.Clear()
                    cmdFormer.Parameters.AddWithValue("@admNo", Me.lstFormerSchool.SelectedItems(0).Text.Trim)
                    cmdFormer.Parameters.AddWithValue("@dateOfReg", Date.Now)
                    cmdFormer.Parameters.AddWithValue("@schoolName", Me.lstFormerSchool.SelectedItems(0).SubItems(2).Text.Trim)
                    cmdFormer.Parameters.AddWithValue("@schoolLevel", Me.lstFormerSchool.SelectedItems(0).SubItems(7).Text.Trim)
                    cmdFormer.Parameters.AddWithValue("@className", Me.lstFormerSchool.SelectedItems(0).SubItems(3).Text.Trim)
                    cmdFormer.Parameters.AddWithValue("@subject", Me.lstFormerSchool.SelectedItems(0).SubItems(4).Text.Trim)
                    cmdFormer.Parameters.AddWithValue("@marks", Me.lstFormerSchool.SelectedItems(0).SubItems(5).Text.Trim)
                    cmdFormer.Parameters.AddWithValue("@outOf", Me.lstFormerSchool.SelectedItems(0).SubItems(6).Text.Trim)
                    cmdFormer.Parameters.AddWithValue("@queryType", queryType.Trim)
                    cmdFormer.Parameters.AddWithValue("@regBy", userName)
                    rec = cmdFormer.ExecuteNonQuery
                    If rec > 0 Then
                        MsgBox("Record Deleted", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "SuccessFul Transaction")
                        Exit Sub
                    End If
                End If
            Catch ex As Exception
                MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Finally
                clearTexts()
                Me.txtStudNo.Enabled = True
                Me.cboClassName.Enabled = True
                Me.cboSchLevel.Enabled = True
                Me.cboSchName.Enabled = True
                Me.cboSubject.Enabled = True
                Me.btnUpdate.Enabled = False
                Me.btnSave.Enabled = True
                queryType = Nothing
                If conn.State = ConnectionState.Open Then
                    conn.Close()
                End If
            End Try
        End If
    End Sub
End Class