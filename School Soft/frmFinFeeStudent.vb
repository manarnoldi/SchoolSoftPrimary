Imports System.Data.SqlClient
Public Class frmFinFeeStudent
    Dim cmdStudFee As New SqlCommand
    Dim reader As SqlDataReader
    Dim rec As Integer = 0
    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub loadCombos()
        Me.cboTerm.Items.Clear()
        Me.cboTerm.Text = ""
        Me.cboTerm.SelectedIndex = -1

        Me.cboVwTerm.Items.Clear()
        Me.cboVwTerm.Text = ""
        Me.cboVwTerm.SelectedIndex = -1

        Me.cboYear.Items.Clear()
        Me.cboYear.Text = ""
        Me.cboYear.SelectedIndex = -1

        Me.cboVwYear.Items.Clear()
        Me.cboVwYear.Text = ""
        Me.cboVwYear.SelectedIndex = -1

        Me.cboStream.Items.Clear()
        Me.cboStream.Text = ""
        Me.cboStream.SelectedIndex = -1

        Me.cboVwStream.Items.Clear()
        Me.cboVwStream.Text = ""
        Me.cboVwStream.SelectedIndex = -1

        Me.cboClass.Items.Clear()
        Me.cboClass.Text = ""
        Me.cboClass.SelectedIndex = -1

        Me.cboVwClass.Items.Clear()
        Me.cboVwClass.Text = ""
        Me.cboVwClass.SelectedIndex = -1

        Me.cboFeeCategory.Items.Clear()
        Me.cboFeeCategory.Text = ""
        Me.cboFeeCategory.SelectedIndex = -1

        Me.cboVwFeeCat.Items.Clear()
        Me.cboVwFeeCat.Text = ""
        Me.cboVwFeeCat.SelectedIndex = -1

        Me.cboVoteHeadName.Items.Clear()
        Me.cboVoteHeadName.Text = ""
        Me.cboVoteHeadName.SelectedIndex = -1

        Me.cmdStudFee.Connection = conn
        Me.cmdStudFee.CommandType = CommandType.Text
        Me.cmdStudFee.CommandText = "SELECT DISTINCT termName FROM tblSchoolCalendar WHERE (status=1) ORDER BY termName"
        Me.cmdStudFee.Parameters.Clear()
        reader = Me.cmdStudFee.ExecuteReader
        While reader.Read
            Me.cboTerm.Items.Add(IIf(DBNull.Value.Equals(reader!termName), "", reader!termName))
            Me.cboVwTerm.Items.Add(IIf(DBNull.Value.Equals(reader!termName), "", reader!termName))
        End While
        reader.Close()

        Me.cmdStudFee.CommandText = "SELECT DISTINCT className FROM tblClasses WHERE (status=1) ORDER BY className"
        Me.cmdStudFee.Parameters.Clear()
        reader = Me.cmdStudFee.ExecuteReader
        While reader.Read
            Me.cboClass.Items.Add(IIf(DBNull.Value.Equals(reader!className), "", reader!className))
            Me.cboVwClass.Items.Add(IIf(DBNull.Value.Equals(reader!className), "", reader!className))
        End While
        reader.Close()

        Me.cmdStudFee.CommandText = "SELECT DISTINCT stream FROM tblClasses WHERE (status=1) ORDER BY stream"
        Me.cmdStudFee.Parameters.Clear()
        reader = Me.cmdStudFee.ExecuteReader
        While reader.Read
            Me.cboStream.Items.Add(IIf(DBNull.Value.Equals(reader!stream), "", reader!stream))
            Me.cboVwStream.Items.Add(IIf(DBNull.Value.Equals(reader!stream), "", reader!stream))
        End While
        reader.Close()

        Me.cmdStudFee.CommandText = "SELECT DISTINCT year FROM tblSchoolCalendar WHERE (status=1) ORDER BY year"
        Me.cmdStudFee.Parameters.Clear()
        reader = Me.cmdStudFee.ExecuteReader
        While reader.Read
            Me.cboYear.Items.Add(IIf(DBNull.Value.Equals(reader!year), "", reader!year))
            Me.cboVwYear.Items.Add(IIf(DBNull.Value.Equals(reader!year), "", reader!year))
        End While
        reader.Close()

        Me.cmdStudFee.CommandText = "SELECT DISTINCT feeCatName FROM tblFinFeeCategory ORDER BY feeCatName"
        Me.cmdStudFee.Parameters.Clear()
        reader = Me.cmdStudFee.ExecuteReader
        While reader.Read
            Me.cboFeeCategory.Items.Add(IIf(DBNull.Value.Equals(reader!feeCatName), "", reader!feeCatName))
            Me.cboVwFeeCat.Items.Add(IIf(DBNull.Value.Equals(reader!feeCatName), "", reader!feeCatName))
        End While
        reader.Close()

        Me.cmdStudFee.CommandText = "SELECT DISTINCT voteHeadName FROM tblFinFeeVoteHeads ORDER BY voteHeadName"
        Me.cmdStudFee.Parameters.Clear()
        reader = Me.cmdStudFee.ExecuteReader
        While reader.Read
            Me.cboVoteHeadName.Items.Add(IIf(DBNull.Value.Equals(reader!voteHeadName), "", reader!voteHeadName))
        End While
        reader.Close()
    End Sub

    Private Sub frmFinStudentFee_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
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

    Private Sub frmFinStudentFee_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint
        If Me.IsMdiChild Then
            Dim PnlLoc As New Point
            PnlLoc.X = CInt((Me.Width - Me.pnlSetFee.Width) / 2)
            PnlLoc.Y = CInt((Me.Height - Me.pnlSetFee.Height) / 2.5)
            Me.pnlSetFee.Location = PnlLoc
        Else
            Me.pnlSetFee.Dock = DockStyle.Fill
        End If
    End Sub

    Private Sub btnLoad_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLoad.Click
        If Me.cboClass.Text.Trim.Length <= 0 Then
            MsgBox("Class Name Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboStream.Text.Trim.Length <= 0 Then
            MsgBox("Stream Is Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboYear.Text.Trim.Length <= 0 Then
            MsgBox("Year Is Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()

            Me.lstStudent.Items.Clear()
            Me.cmdStudFee.Connection = conn
            Me.cmdStudFee.CommandType = CommandType.Text
            Me.cmdStudFee.CommandText = "SELECT * FROM vwStudClass WHERE (className=@className) AND (stream=@stream) AND " &
                vbNewLine & " (year=@year) ORDER BY boarder,admNo"
            Me.cmdStudFee.Parameters.Clear()
            Me.cmdStudFee.Parameters.AddWithValue("@className", Me.cboClass.Text.Trim)
            Me.cmdStudFee.Parameters.AddWithValue("@stream", Me.cboStream.Text.Trim)
            Me.cmdStudFee.Parameters.AddWithValue("@year", Me.cboYear.Text.Trim)
            reader = Me.cmdStudFee.ExecuteReader
            While reader.Read
                li = Me.lstStudent.Items.Add(IIf(DBNull.Value.Equals(reader!admNo), "", reader!admNo))
                li.SubItems.Add(IIf(DBNull.Value.Equals(reader!FullName), "", reader!FullName))
                li.Tag = IIf(DBNull.Value.Equals(reader!studId), "", reader!studId)
                li.SubItems(1).Tag = IIf(DBNull.Value.Equals(reader!boarder), "", reader!boarder)
            End While
            reader.Close()

            For i = 0 To Me.lstStudent.Items.Count - 1
                If Me.lstStudent.Items(i).SubItems(1).Tag = "True" Then
                    Me.lstStudent.Items(i).BackColor = Color.LightSkyBlue
                End If
            Next

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
            For i = 0 To Me.lstVoteHeads.Items.Count - 1
                If Me.cboVoteHeadName.Text.Trim = Me.lstVoteHeads.Items(i).Text.Trim Then
                    MsgBox("VoteHead Name Already In The List.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
                    Exit Sub
                End If
            Next
            li = Me.lstVoteHeads.Items.Add(Me.cboVoteHeadName.Text.Trim)
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

    Private Sub ToolStripMenuItem2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem2.Click
        If Me.lstVoteHeads.Items.Count <= 0 Then
            MsgBox("No item in the list to remove.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.lstVoteHeads.SelectedItems.Count <= 0 Then
            MsgBox("No Selected item/s in the list to remove.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If
        For i = 0 To Me.lstVoteHeads.SelectedItems.Count - 1
            Me.lstVoteHeads.SelectedItems(0).Remove()
        Next
    End Sub

    Private Sub cboStream_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboStream.SelectedIndexChanged
        Me.lstStudent.Items.Clear()
    End Sub

    Private Sub cboClass_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboClass.SelectedIndexChanged
        Me.lstStudent.Items.Clear()
    End Sub

    Private Sub cboYear_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboYear.SelectedIndexChanged
        Me.lstStudent.Items.Clear()
    End Sub

    'Private Sub cboFeeCategory_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboFeeCategory.SelectedIndexChanged
    '    Me.lstVoteHeads.Items.Clear()
    '    Try
    '        If conn.State = ConnectionState.Closed Then
    '            conn.Open()
    '        End If
    '        dbconnection()
    '        Me.lstVoteHeads.Items.Clear()
    '        Me.cmdStudFee.Connection = conn
    '        Me.cmdStudFee.CommandType = CommandType.Text
    '        Me.cmdStudFee.CommandText = "SELECT * FROM vwFinFeeSetUp WHERE (feeCatName=@feeCatName) AND (className=@className) AND " &
    '            vbNewLine & " (stream=@stream) AND (year=@year) AND (termName=@termName) ORDER BY priority"
    '        Me.cmdStudFee.Parameters.Clear()
    '        Me.cmdStudFee.Parameters.AddWithValue("@className", Me.cboClass.Text.Trim)
    '        Me.cmdStudFee.Parameters.AddWithValue("@stream", Me.cboStream.Text.Trim)
    '        Me.cmdStudFee.Parameters.AddWithValue("@year", Me.cboYear.Text.Trim)
    '        Me.cmdStudFee.Parameters.AddWithValue("@feeCatName", Me.cboFeeCategory.Text.Trim)
    '        Me.cmdStudFee.Parameters.AddWithValue("@termName", Me.cboTerm.Text.Trim)
    '        reader = Me.cmdStudFee.ExecuteReader
    '        While reader.Read
    '            li = Me.lstVoteHeads.Items.Add(IIf(DBNull.Value.Equals(reader!voteHeadName), "", reader!voteHeadName))
    '            li.SubItems.Add(IIf(DBNull.Value.Equals(reader!amount), "", reader!amount))
    '        End While
    '        reader.Close()
    '    Catch ex As Exception
    '        MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
    '    Finally
    '        If conn.State = ConnectionState.Open Then
    '            conn.Close()
    '        End If
    '    End Try
    'End Sub

    Private Sub cboVwYear_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles _
        cboVwYear.SelectedIndexChanged, cboVwClass.SelectedIndexChanged, cboVwStream.SelectedIndexChanged
        Me.lstVwStudFee.Items.Clear()
        If Me.cboVwClass.Text.Trim.Length > 0 And Me.cboVwStream.Text.Trim.Length > 0 And Me.cboVwYear.Text.Trim.Length Then
            Try
                If conn.State = ConnectionState.Closed Then
                    conn.Open()
                End If
                dbconnection()

                Me.cboVwAdm.Items.Clear()
                Me.cboVwAdm.Text = ""
                Me.cboVwAdm.SelectedIndex = -1

                Me.cmdStudFee.Connection = conn
                Me.cmdStudFee.CommandType = CommandType.Text
                Me.cmdStudFee.CommandText = "SELECT admNo FROM vwStudClass WHERE (className=@className) AND (stream=@stream) AND " &
                    vbNewLine & " (year=@year) ORDER BY admNo"
                Me.cmdStudFee.Parameters.Clear()
                Me.cmdStudFee.Parameters.AddWithValue("@className", Me.cboVwClass.Text.Trim)
                Me.cmdStudFee.Parameters.AddWithValue("@stream", Me.cboVwStream.Text.Trim)
                Me.cmdStudFee.Parameters.AddWithValue("@year", Me.cboVwYear.Text.Trim)
                reader = Me.cmdStudFee.ExecuteReader
                While reader.Read
                    Me.cboVwAdm.Items.Add(IIf(DBNull.Value.Equals(reader!admNo), "", reader!admNo))
                End While
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

    Private Sub btnVwLoad_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnVwLoad.Click
        Me.ToolStripLabel1.Text = "Total Amount"
        Me.lstVwStudFee.Items.Clear()
        If Me.cboVwClass.Text.Trim.Length <= 0 Then
            MsgBox("Class Name Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboVwStream.Text.Trim.Length <= 0 Then
            MsgBox("Stream Is Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboVwYear.Text.Trim.Length <= 0 Then
            MsgBox("Year Is Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboVwTerm.Text.Trim.Length <= 0 Then
            MsgBox("Term Is Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboVwFeeCat.Text.Trim.Length <= 0 Then
            MsgBox("Fee Category Is Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboVwAdm.Text.Trim.Length <= 0 Then
            MsgBox("Admission Number Is Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If

        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()

            Me.cmdStudFee.Connection = conn
            Me.cmdStudFee.CommandType = CommandType.Text
            Me.cmdStudFee.CommandText = "SELECT * FROM vwFinFeeStudent WHERE (feeCatName=@feeCatName) AND (className=@className) AND " &
                vbNewLine & " (stream=@stream) AND (year=@year) AND (termName=@termName) AND (admNo=@admNo) ORDER BY priority"
            Me.cmdStudFee.Parameters.Clear()
            Me.cmdStudFee.Parameters.AddWithValue("@className", Me.cboVwClass.Text.Trim)
            Me.cmdStudFee.Parameters.AddWithValue("@stream", Me.cboVwStream.Text.Trim)
            Me.cmdStudFee.Parameters.AddWithValue("@year", Me.cboVwYear.Text.Trim)
            Me.cmdStudFee.Parameters.AddWithValue("@feeCatName", Me.cboVwFeeCat.Text.Trim)
            Me.cmdStudFee.Parameters.AddWithValue("@termName", Me.cboVwTerm.Text.Trim)
            Me.cmdStudFee.Parameters.AddWithValue("@admNo", Me.cboVwAdm.Text.Trim)
            reader = Me.cmdStudFee.ExecuteReader
            Dim totalAmount As Double = 0
            If reader.HasRows = True Then
                While reader.Read
                    li = Me.lstVwStudFee.Items.Add(IIf(DBNull.Value.Equals(reader!voteHeadName), "", reader!voteHeadName))
                    li.SubItems.Add(IIf(DBNull.Value.Equals(reader!priority), "", reader!priority))
                    li.SubItems.Add(IIf(DBNull.Value.Equals(reader!amount), "", reader!amount))
                    li.Tag = IIf(DBNull.Value.Equals(reader!studFeeId), "", reader!studFeeId)
                    totalAmount = totalAmount + IIf(DBNull.Value.Equals(reader!amount), "", reader!amount)
                End While
            ElseIf reader.HasRows = False Then
                Me.ToolStripLabel1.Text = "Total Amount"
                MsgBox("No Record Found!", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "System Reply")
            End If
            reader.Close()
            If totalAmount > 0 Then
                Me.ToolStripLabel1.Text = totalAmount
            Else
                Me.ToolStripLabel1.Text = "Total Amount"
            End If
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

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        If Me.cboTerm.Text.Trim.Length <= 0 Then
            MsgBox("Select Term.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboClass.Text.Trim.Length <= 0 Then
            MsgBox("Class Name Is Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboStream.Text.Trim.Length <= 0 Then
            MsgBox("Stream Is Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboYear.Text.Trim.Length <= 0 Then
            MsgBox("Year Is Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboFeeCategory.Text.Trim.Length <= 0 Then
            MsgBox("Select Fee Category.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.lstStudent.Items.Count <= 0 Then
            MsgBox("No Student To Save Against.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.lstStudent.CheckedItems.Count <= 0 Then
            MsgBox("No checked student/s To Save Against.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.lstVoteHeads.Items.Count <= 0 Then
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
            For i = 0 To Me.lstStudent.CheckedItems.Count - 1
                j = 0
                For j = 0 To Me.lstVoteHeads.Items.Count - 1
                    Dim feeCatConflict As Boolean = checkStudFeeCatConflict(Me.lstStudent.CheckedItems(i).Text.Trim,
                    Me.cboFeeCategory.Text.Trim, Me.cboClass.Text.Trim, Me.cboStream.Text.Trim, Me.cboYear.Text.Trim, Me.cboTerm.Text.Trim)
                    If feeCatConflict = True Then
                        MsgBox("Student AdmNo " & Me.lstStudent.CheckedItems(i).Text.Trim _
                               & " Is Already Assigned A different Fee Category.", MsgBoxStyle.Exclamation +
                               MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
                        Exit Sub
                    End If


                    Dim recordExists As Boolean = checkIfRecordExists(Me.cboClass.Text.Trim, Me.cboStream.Text.Trim,
                    Me.cboYear.Text.Trim, Me.cboFeeCategory.Text.Trim, Me.cboTerm.Text.Trim, Me.lstVoteHeads.Items(j).Text.Trim,
                    Me.lstStudent.CheckedItems(i).Text.Trim)
                    If recordExists = True Then
                        MsgBox("VoteHead (" + Me.lstVoteHeads.Items(j).Text.Trim +
                               ") is already assigned to the student for the term. Update the amount instead incase the figure is the problem.",
                               MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
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
            For i = 0 To Me.lstStudent.CheckedItems.Count - 1
                j = 0
                For j = 0 To Me.lstVoteHeads.Items.Count - 1
                    Me.cmdStudFee.Connection = conn
                    Me.cmdStudFee.CommandType = CommandType.StoredProcedure
                    Me.cmdStudFee.CommandText = "sprocFinFeeStudent"
                    Me.cmdStudFee.Parameters.Clear()

                    Me.cmdStudFee.Parameters.AddWithValue("@className", Me.cboClass.Text.Trim)
                    Me.cmdStudFee.Parameters.AddWithValue("@stream", Me.cboStream.Text.Trim)
                    Me.cmdStudFee.Parameters.AddWithValue("@year", Me.cboYear.Text.Trim)
                    Me.cmdStudFee.Parameters.AddWithValue("@feeCatName", Me.cboFeeCategory.Text.Trim)
                    Me.cmdStudFee.Parameters.AddWithValue("@termName", Me.cboTerm.Text.Trim)
                    Me.cmdStudFee.Parameters.AddWithValue("@voteHeadName", Me.lstVoteHeads.Items(j).Text.Trim)
                    Me.cmdStudFee.Parameters.AddWithValue("@amount", Me.lstVoteHeads.Items(j).SubItems(1).Text.Trim)
                    Me.cmdStudFee.Parameters.AddWithValue("@regBy", userName.Trim)
                    Me.cmdStudFee.Parameters.AddWithValue("@queryType", "INSERT")
                    Me.cmdStudFee.Parameters.AddWithValue("@admNo", Me.lstStudent.CheckedItems(i).Text.Trim)
                    rec = rec + Me.cmdStudFee.ExecuteNonQuery
                Next
            Next

            If rec > 0 Then
                MsgBox("Record/s Saved Successfully.", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "System Reply")
            End If
            Me.lstVoteHeads.Items.Clear()
            'loadCombos()
            'clearTexts()
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub
    Private Sub clearTexts()
        Me.lstVoteHeads.Items.Clear()
        Me.lstStudent.Items.Clear()
        Me.lstVwStudFee.Items.Clear()

        Me.txtVoteAmount.Text = ""
        Me.lstVwStudFee.Tag = Nothing
        Me.txtVoteAmount.Tag = Nothing

        Me.cboVwAdm.Items.Clear()
        Me.cboVwAdm.Text = ""
        Me.cboVwAdm.SelectedIndex = -1

        Me.lstStudent.Enabled = True
        Me.lstVoteHeads.Enabled = True

        Me.ToolStripLabel1.Text = "Total Amount"
        Me.cboTerm.Enabled = True
        Me.cboYear.Enabled = True
        Me.cboClass.Enabled = True
        Me.cboStream.Enabled = True
        Me.cboFeeCategory.Enabled = True
        Me.cboVoteHeadName.Enabled = True

        Me.btnSave.Enabled = True
        Me.btnUpdate.Enabled = False
        Me.btnAdd.Enabled = True
        Me.btnLoad.Enabled = True
    End Sub
    Private Function checkIfRecordExists(ByVal className As String, ByVal stream As String, ByVal year As Integer,
    ByVal feeCatName As String, ByVal termName As String, ByVal voteHeadName As String, ByVal admNo As String)
        Dim recordExists As Boolean = False
        Me.cmdStudFee.Connection = conn
        Me.cmdStudFee.CommandType = CommandType.Text
        Me.cmdStudFee.CommandText = "SELECT * FROM vwFinFeeStudent WHERE (className=@className) AND (stream=@stream) " &
            vbNewLine & " AND (year=@year) AND (feeCatName=@feeCatName) AND (termName=@termName) " &
            vbNewLine & " AND (admNo=@admNo) AND (voteHeadName=@voteHeadName)"
        Me.cmdStudFee.Parameters.Clear()
        Me.cmdStudFee.Parameters.AddWithValue("@className", className.Trim)
        Me.cmdStudFee.Parameters.AddWithValue("@stream", stream.Trim)
        Me.cmdStudFee.Parameters.AddWithValue("@year", year)
        Me.cmdStudFee.Parameters.AddWithValue("@feeCatName", feeCatName.Trim)
        Me.cmdStudFee.Parameters.AddWithValue("@termName", termName.Trim)
        Me.cmdStudFee.Parameters.AddWithValue("@voteHeadName", voteHeadName.Trim)
        Me.cmdStudFee.Parameters.AddWithValue("@admNo", admNo.Trim)
        reader = Me.cmdStudFee.ExecuteReader
        If reader.HasRows = True Then
            recordExists = True
        ElseIf reader.HasRows = False Then
            recordExists = False
        End If
        reader.Close()
        Return recordExists
    End Function

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

    Private Sub lstStudent_ColumnClick(ByVal sender As Object, ByVal e As System.Windows.Forms.ColumnClickEventArgs) Handles lstStudent.ColumnClick
        If (e.Column() = 0) And (Me.lstStudent.Items.Count > 0) Then
            For Each Li As ListViewItem In Me.lstStudent.Items
                Li.Checked = Not (Li.Checked)
            Next
        End If
    End Sub

    Private Sub DELETEToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DELETEToolStripMenuItem.Click
        If Me.lstVwStudFee.Items.Count <= 0 Then
            MsgBox("No items in the list to delete.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.lstVwStudFee.CheckedItems.Count <= 0 Then
            MsgBox("No checked items in the list to delete.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            If checkIfFeePaymentsHaveStarted(Me.cboVwAdm.Text.Trim, Me.cboVwYear.Text.Trim, Me.cboVwTerm.Text.Trim) Then
                MsgBox("Student Fee payments have started hence you cannot delete the record! You can only update the amount. " +
                       "Ask the administrator to reverse the student fee setup details instead!", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
                Me.btnUpdate.Enabled = False
                Exit Sub
            End If

            Dim result As MsgBoxResult = MsgBox("Delete Record/s?", MsgBoxStyle.Question + MsgBoxStyle.ApplicationModal + MsgBoxStyle.YesNo, "Confirm Transaction")
            If result = MsgBoxResult.No Then
                Exit Sub
            End If

            For i = 0 To Me.lstVwStudFee.CheckedItems.Count - 1
                Me.cmdStudFee.Connection = conn
                Me.cmdStudFee.CommandType = CommandType.StoredProcedure
                Me.cmdStudFee.CommandText = "sprocFinFeeStudent"
                Me.cmdStudFee.Parameters.Clear()
                Me.cmdStudFee.Parameters.AddWithValue("@className", Me.cboVwClass.Text.Trim)
                Me.cmdStudFee.Parameters.AddWithValue("@stream", Me.cboVwStream.Text.Trim)
                Me.cmdStudFee.Parameters.AddWithValue("@year", Me.cboVwYear.Text.Trim)
                Me.cmdStudFee.Parameters.AddWithValue("@feeCatName", Me.cboVwFeeCat.Text.Trim)
                Me.cmdStudFee.Parameters.AddWithValue("@termName", Me.cboVwTerm.Text.Trim)
                Me.cmdStudFee.Parameters.AddWithValue("@admNo", Me.cboVwAdm.Text.Trim)
                Me.cmdStudFee.Parameters.AddWithValue("@voteHeadName", Me.lstVwStudFee.CheckedItems(i).Text.Trim)
                Me.cmdStudFee.Parameters.AddWithValue("@amount", Me.lstVwStudFee.CheckedItems(i).SubItems(2).Text.Trim)
                Me.cmdStudFee.Parameters.AddWithValue("@regBy", userName.Trim)
                Me.cmdStudFee.Parameters.AddWithValue("@studFeeId", Me.lstVwStudFee.CheckedItems(i).Tag)
                Me.cmdStudFee.Parameters.AddWithValue("@queryType", "DELETE")
                rec = rec + Me.cmdStudFee.ExecuteNonQuery
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

    Private Function checkIfFeePaymentsHaveStarted(ByVal admNo As String, ByVal studYear As Integer, ByVal termName As String)
        Dim feePayId As Integer = 0
        Dim votePayCount As Integer = 0
        Dim feePaymentStarted As Boolean = False
        Dim studId As Integer = 0
        Dim termId As Integer = 0

        Me.cmdStudFee.CommandType = CommandType.Text
        Me.cmdStudFee.Connection = conn
        Me.cmdStudFee.CommandText = "SELECT studId FROM tblStudDetails WHERE (admNo=@admNo) And (status=1)"
        Me.cmdStudFee.Parameters.Clear()
        Me.cmdStudFee.Parameters.AddWithValue("@admNo", admNo)
        reader = Me.cmdStudFee.ExecuteReader
        While reader.Read
            studId = IIf(DBNull.Value.Equals(reader!studId), 0, reader!studId)
        End While
        reader.Close()

        Me.cmdStudFee.CommandText = "SELECT termId FROM tblSchoolCalendar WHERE (termName=@termName) AND (year=@year) AND (status=1)"
        Me.cmdStudFee.Parameters.Clear()
        Me.cmdStudFee.Parameters.AddWithValue("@termName", termName)
        Me.cmdStudFee.Parameters.AddWithValue("@year", studYear)
        reader = Me.cmdStudFee.ExecuteReader
        While reader.Read
            termId = IIf(DBNull.Value.Equals(reader!termId), 0, reader!termId)
        End While
        reader.Close()

        Me.cmdStudFee.CommandText = "SELECT feePayId FROM tblFinFeePayDetails WHERE (studId=@studId) AND (year=@year) AND " +
        "(termId=@termId) AND (reversed=0)"
        Me.cmdStudFee.Parameters.Clear()
        Me.cmdStudFee.Parameters.AddWithValue("@studId", studId)
        Me.cmdStudFee.Parameters.AddWithValue("@year", studYear)
        Me.cmdStudFee.Parameters.AddWithValue("@termId", termId)
        reader = Me.cmdStudFee.ExecuteReader
        While reader.Read
            feePayId = IIf(DBNull.Value.Equals(reader!feePayId), 0, reader!feePayId)
        End While
        reader.Close()

        If feePayId > 0 Then
            Me.cmdStudFee.CommandText = "SELECT COUNT(voteFeePayId) AS Count FROM tblFinFeePayVotes WHERE (feePayId=@feePayId) AND (reversed=0)"
            Me.cmdStudFee.Parameters.Clear()
            Me.cmdStudFee.Parameters.AddWithValue("@feePayId", feePayId)
            reader = Me.cmdStudFee.ExecuteReader
            While reader.Read
                votePayCount = IIf(DBNull.Value.Equals(reader!Count), "", reader!Count)
            End While
            reader.Close()
        End If

        If votePayCount > 0 Then
            feePaymentStarted = True
        End If
        Return feePaymentStarted
    End Function

    Private Sub UPDATEToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UPDATEToolStripMenuItem.Click
        If Me.lstVwStudFee.Items.Count <= 0 Then
            MsgBox("No items in the list to update.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.lstVwStudFee.CheckedItems.Count <= 0 Then
            MsgBox("No checked items in the list to update.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.lstVwStudFee.CheckedItems.Count > 1 Then
            MsgBox("Update One Item At A time.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If
        Try
            Me.cboFeeCategory.Text = Me.cboVwFeeCat.Text.Trim
            Me.cboTerm.Text = Me.cboVwTerm.Text.Trim
            Me.cboVoteHeadName.Text = Me.lstVwStudFee.CheckedItems(0).Text.Trim
            Me.cboYear.Text = Me.cboVwYear.Text.Trim
            Me.cboClass.Text = Me.cboVwClass.Text.Trim
            Me.cboStream.Text = Me.cboVwStream.Text.Trim
            Me.cboTerm.Text = Me.cboVwTerm.Text.Trim

            Me.txtVoteAmount.Text = Me.lstVwStudFee.CheckedItems(0).SubItems(2).Text.Trim
            Me.lstVwStudFee.Tag = Me.lstVwStudFee.CheckedItems(0).Tag
            Me.txtVoteAmount.Tag = Me.lstVwStudFee.CheckedItems(0).SubItems(2).Text.Trim

            Me.cboClass.Enabled = False
            Me.cboStream.Enabled = False
            Me.cboTerm.Enabled = False
            Me.cboYear.Enabled = False
            Me.cboFeeCategory.Enabled = False
            Me.cboVoteHeadName.Enabled = False

            Me.btnUpdate.Enabled = True
            Me.btnSave.Enabled = False
            Me.btnAdd.Enabled = False
            Me.btnLoad.Enabled = False

            Me.lstVoteHeads.Items.Clear()
            Me.lstStudent.Items.Clear()

            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If

            dbconnection()
            Me.lstStudent.Items.Clear()
            Me.cmdStudFee.Connection = conn
            Me.cmdStudFee.CommandType = CommandType.Text
            Me.cmdStudFee.CommandText = "SELECT * FROM vwStudClass WHERE (className=@className) AND (stream=@stream) AND " &
                vbNewLine & " (year=@year) AND (admNo=@admNo)"
            Me.cmdStudFee.Parameters.Clear()
            Me.cmdStudFee.Parameters.AddWithValue("@className", Me.cboVwClass.Text.Trim)
            Me.cmdStudFee.Parameters.AddWithValue("@stream", Me.cboVwStream.Text.Trim)
            Me.cmdStudFee.Parameters.AddWithValue("@year", Me.cboVwYear.Text.Trim)
            Me.cmdStudFee.Parameters.AddWithValue("@admNo", Me.cboVwAdm.Text.Trim)
            reader = Me.cmdStudFee.ExecuteReader
            While reader.Read
                li = Me.lstStudent.Items.Add(IIf(DBNull.Value.Equals(reader!admNo), "", reader!admNo))
                li.SubItems.Add(IIf(DBNull.Value.Equals(reader!FullName), "", reader!FullName))
                li.Tag = IIf(DBNull.Value.Equals(reader!studId), "", reader!studId)
            End While
            reader.Close()

            Me.lstStudent.Items(0).Checked = True

            Me.lstStudent.Enabled = False
            Me.lstVoteHeads.Enabled = False
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
            MsgBox("Select Term.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboClass.Text.Trim.Length <= 0 Then
            MsgBox("Class Name Is Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboStream.Text.Trim.Length <= 0 Then
            MsgBox("Stream Is Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboYear.Text.Trim.Length <= 0 Then
            MsgBox("Year Is Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboFeeCategory.Text.Trim.Length <= 0 Then
            MsgBox("Select Fee Category.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.lstStudent.Items.Count <= 0 Then
            MsgBox("No Student To Save Against.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.lstStudent.CheckedItems.Count <= 0 Then
            MsgBox("No checked student/s To Save Against.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboVoteHeadName.Items.Count <= 0 Then
            MsgBox("No votehead Name to Update.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.txtVoteAmount.Text.Trim.Length <= 0 Then
            MsgBox("No votehead Amount to Update.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If

        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            i = 0
            For i = 0 To Me.lstStudent.CheckedItems.Count - 1
                Dim recordExists As Boolean = checkIfRecordExistsUpdate(Me.cboClass.Text.Trim, Me.cboStream.Text.Trim, Me.cboYear.Text.Trim,
                Me.cboFeeCategory.Text.Trim, Me.cboTerm.Text.Trim, Me.cboVoteHeadName.Text.Trim, Me.lstStudent.CheckedItems(0).Text.Trim, Me.txtVoteAmount.Text.Trim)
                If recordExists = True Then
                    MsgBox("Duplicate Found in the database!", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
                    Exit Sub
                End If
            Next

            Dim result As MsgBoxResult = MsgBox("Update Record?", MsgBoxStyle.Question + MsgBoxStyle.ApplicationModal + MsgBoxStyle.YesNo, "Confirm Transaction")
            If result = MsgBoxResult.No Then
                Exit Sub
            End If

            Me.cmdStudFee.Connection = conn
            Me.cmdStudFee.CommandType = CommandType.StoredProcedure
            Me.cmdStudFee.CommandText = "sprocFinFeeStudent"
            Me.cmdStudFee.Parameters.Clear()
            Me.cmdStudFee.Parameters.AddWithValue("@className", Me.cboClass.Text.Trim)
            Me.cmdStudFee.Parameters.AddWithValue("@stream", Me.cboStream.Text.Trim)
            Me.cmdStudFee.Parameters.AddWithValue("@year", Me.cboYear.Text.Trim)
            Me.cmdStudFee.Parameters.AddWithValue("@feeCatName", Me.cboFeeCategory.Text.Trim)
            Me.cmdStudFee.Parameters.AddWithValue("@admNo", Me.lstStudent.CheckedItems(0).Text.Trim)
            Me.cmdStudFee.Parameters.AddWithValue("@termName", Me.cboTerm.Text.Trim)
            Me.cmdStudFee.Parameters.AddWithValue("@voteHeadName", Me.cboVoteHeadName.Text.Trim)
            Me.cmdStudFee.Parameters.AddWithValue("@amount", Me.txtVoteAmount.Text.Trim)
            Me.cmdStudFee.Parameters.AddWithValue("@regBy", userName.Trim)
            Me.cmdStudFee.Parameters.AddWithValue("@studFeeId", Me.lstVwStudFee.Tag)
            Me.cmdStudFee.Parameters.AddWithValue("@fromAmount", Me.txtVoteAmount.Tag)
            Me.cmdStudFee.Parameters.AddWithValue("@queryType", "UPDATE")
            rec = rec + Me.cmdStudFee.ExecuteNonQuery

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
    Private Function checkIfRecordExistsUpdate(ByVal className As String, ByVal stream As String, ByVal year As Integer,
    ByVal feeCatName As String, ByVal termName As String, ByVal voteHeadName As String, ByVal admNo As String, ByVal amount As Integer)
        Dim recordExists As Boolean = False
        Me.cmdStudFee.Connection = conn
        Me.cmdStudFee.CommandType = CommandType.Text
        Me.cmdStudFee.CommandText = "SELECT * FROM vwFinFeeStudent WHERE (className=@className) AND (stream=@stream) " &
            vbNewLine & " AND (year=@year) AND (feeCatName=@feeCatName) AND (termName=@termName) " &
            vbNewLine & " AND (admNo=@admNo) AND (voteHeadName=@voteHeadName) AND (amount=@amount)"
        Me.cmdStudFee.Parameters.Clear()
        Me.cmdStudFee.Parameters.AddWithValue("@className", className.Trim)
        Me.cmdStudFee.Parameters.AddWithValue("@stream", stream.Trim)
        Me.cmdStudFee.Parameters.AddWithValue("@year", year)
        Me.cmdStudFee.Parameters.AddWithValue("@feeCatName", feeCatName.Trim)
        Me.cmdStudFee.Parameters.AddWithValue("@termName", termName.Trim)
        Me.cmdStudFee.Parameters.AddWithValue("@voteHeadName", voteHeadName.Trim)
        Me.cmdStudFee.Parameters.AddWithValue("@admNo", admNo.Trim)
        Me.cmdStudFee.Parameters.AddWithValue("@amount", amount)
        reader = Me.cmdStudFee.ExecuteReader
        If reader.HasRows = True Then
            recordExists = True
        ElseIf reader.HasRows = False Then
            recordExists = False
        End If
        reader.Close()
        Return recordExists
    End Function
    Private Function checkStudFeeCatConflict(ByVal admNo As String, ByVal feeCatName As String, ByVal className As String, ByVal stream As String,
                                             ByVal year As Integer, ByVal termName As String)
        Dim conflictFound As Boolean = False

        Me.cmdStudFee.Connection = conn
        Me.cmdStudFee.CommandType = CommandType.Text
        Me.cmdStudFee.CommandText = "SELECT admNo,feeCatName FROM vwFinFeeStudent WHERE (admNo=@admNo) AND " &
            vbNewLine & " (feeCatName<>@feeCatName) AND (classname=@className) AND (stream=@stream) AND (year=@year) AND (termName=@termName)"
        Me.cmdStudFee.Parameters.Clear()
        Me.cmdStudFee.Parameters.AddWithValue("@admNo", admNo.Trim)
        Me.cmdStudFee.Parameters.AddWithValue("@feeCatName", feeCatName.Trim)
        Me.cmdStudFee.Parameters.AddWithValue("@className", Me.cboClass.Text.Trim)
        Me.cmdStudFee.Parameters.AddWithValue("@stream", Me.cboStream.Text.Trim)
        Me.cmdStudFee.Parameters.AddWithValue("@year", Me.cboYear.Text.Trim)
        Me.cmdStudFee.Parameters.AddWithValue("@termName", Me.cboTerm.Text.Trim)
        reader = Me.cmdStudFee.ExecuteReader
        If reader.HasRows = True Then
            conflictFound = True
        ElseIf reader.HasRows = False Then
            conflictFound = False
        End If
        reader.Close()
        Return conflictFound
    End Function

    Private Sub cboTerm_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboTerm.SelectedIndexChanged
        Me.lstVoteHeads.Items.Clear()
    End Sub

    Private Sub cboVwTerm_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboVwTerm.SelectedIndexChanged
        Me.lstVwStudFee.Items.Clear()
    End Sub

    Private Sub cboVwFeeCat_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboVwFeeCat.SelectedIndexChanged
        Me.lstVwStudFee.Items.Clear()
    End Sub

    Private Sub cboVwAdm_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboVwAdm.SelectedIndexChanged
        Me.lstVwStudFee.Items.Clear()
    End Sub
End Class