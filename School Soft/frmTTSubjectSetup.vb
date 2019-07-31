Imports System.Data.SqlClient
Public Class frmTTSubjectSetup
    Dim reader As SqlDataReader
    Dim rec As Integer
    Dim cmdTTSubSetup As New SqlCommand
    Dim rowNo As Integer
    Dim columnNo As Integer
    Private Sub frmTTSubjectSetup_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
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
    Private Sub loadSubjSetup()
        Me.dgSubjectSetUp.Rows.Clear()

        Me.cmdTTSubSetup.Connection = conn
        Me.cmdTTSubSetup.CommandType = CommandType.Text
        Me.cmdTTSubSetup.CommandText = "SELECT * FROM  tblTTSetUp WHERE (academicYear=@year)"
        Me.cmdTTSubSetup.Parameters.Clear()
        Me.cmdTTSubSetup.Parameters.AddWithValue("@year", Me.cboYear.Text.Trim)
        reader = Me.cmdTTSubSetup.ExecuteReader
        While reader.Read
            rowNo = (IIf(DBNull.Value.Equals(reader!daysPerWeek), "", reader!daysPerWeek))
            columnNo = (IIf(DBNull.Value.Equals(reader!lessonsPerDay), "", reader!lessonsPerDay))
        End While
        reader.Close()
        Me.dgSubjectSetUp.ColumnCount = columnNo + 1
        Me.dgSubjectSetUp.Columns(0).Name = "days"
        Me.dgSubjectSetUp.Columns(0).HeaderText = "Days Name"
        Me.dgSubjectSetUp.Columns(0).Width = 55
        Me.dgSubjectSetUp.Columns(0).ReadOnly = True
        Me.dgSubjectSetUp.Columns(0).SortMode = DataGridViewColumnSortMode.NotSortable
        For i = 1 To columnNo
            Me.dgSubjectSetUp.Columns(i).Name = "Lesson" & i
            Me.dgSubjectSetUp.Columns(i).HeaderText = i
            Me.dgSubjectSetUp.Columns(i).Width = 39
            'Me.dgSubjectSetUp.Columns(i).DefaultCellStyle.BackColor = Color.Green
            Me.dgSubjectSetUp.Columns(i).ReadOnly = True
            Me.dgSubjectSetUp.Columns(i).SortMode = DataGridViewColumnSortMode.NotSortable
        Next
        j = 0
        For j = 0 To rowNo - 1
            If j = 0 Then
                Me.dgSubjectSetUp.Rows.Add("MON")
            ElseIf j = 1 Then
                Me.dgSubjectSetUp.Rows.Add("TUE")
            ElseIf j = 2 Then
                Me.dgSubjectSetUp.Rows.Add("WED")
            ElseIf j = 3 Then
                Me.dgSubjectSetUp.Rows.Add("THU")
            ElseIf j = 4 Then
                Me.dgSubjectSetUp.Rows.Add("FRI")
            ElseIf j = 5 Then
                Me.dgSubjectSetUp.Rows.Add("SAT")
            ElseIf j = 6 Then
                Me.dgSubjectSetUp.Rows.Add("SUN")
            End If
        Next
        For i = 1 To Me.dgSubjectSetUp.ColumnCount - 1
            j = 0
            For j = 0 To Me.dgSubjectSetUp.RowCount - 1
                Me.cmdTTSubSetup.Connection = conn
                Me.cmdTTSubSetup.CommandType = CommandType.Text
                Me.cmdTTSubSetup.CommandText = "SELECT allowed FROM vwTTSubjectSetUp WHERE (acadYear=@acadYear) AND (subName=@subName) " & _
                    vbNewLine & " AND (lessonNo=@lessonNo) AND (weekDayNo=@weekDayNo)"
                Me.cmdTTSubSetup.Parameters.Clear()
                Me.cmdTTSubSetup.Parameters.AddWithValue("@acadYear", Me.cboYear.Text.Trim)
                Me.cmdTTSubSetup.Parameters.AddWithValue("@subName", Me.cboSubjectName.Text.Trim)
                Me.cmdTTSubSetup.Parameters.AddWithValue("@lessonNo", i)
                Me.cmdTTSubSetup.Parameters.AddWithValue("@weekDayNo", j + 1)
                reader = Me.cmdTTSubSetup.ExecuteReader
                While reader.Read
                    Me.dgSubjectSetUp.Item(i, j).Value = (IIf(DBNull.Value.Equals(reader!allowed), "", reader!allowed))
                    If (IIf(DBNull.Value.Equals(reader!allowed), "", reader!allowed)) = "YES" Then
                        Me.dgSubjectSetUp.Item(i, j).Style.BackColor = Color.Green
                    ElseIf (IIf(DBNull.Value.Equals(reader!allowed), "", reader!allowed)) = "NO" Then
                        Me.dgSubjectSetUp.Item(i, j).Style.BackColor = Color.Red
                    End If
                End While
                reader.Close()
            Next
        Next
    End Sub
    Private Sub frmTTSubjectSetup_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint
        If Me.IsMdiChild Then
            Dim PnlLoc As New Point
            PnlLoc.X = CInt((Me.Width - Me.pnlSubjectSetup.Width) / 2)
            PnlLoc.Y = CInt((Me.Height - Me.pnlSubjectSetup.Height) / 2.5)
            Me.pnlSubjectSetup.Location = PnlLoc
        Else
            Me.pnlSubjectSetup.Dock = DockStyle.Fill
        End If
    End Sub
    Private Sub loadCombos()
        Me.cboSubjectGroup.Items.Clear()
        Me.cboSubjectGroup.Text = ""
        Me.cboSubjectGroup.SelectedIndex = -1

        Me.cboSubjectColor.Items.Clear()
        Me.cboSubjectColor.Text = ""
        Me.cboSubjectColor.SelectedIndex = -1

        Me.cboYear.Items.Clear()
        Me.cboYear.Text = ""
        Me.cboYear.SelectedIndex = -1

        Me.cmdTTSubSetup.Connection = conn
        Me.cmdTTSubSetup.CommandType = CommandType.Text
        Me.cmdTTSubSetup.CommandText = "SELECT DISTINCT subGroup FROM tblSubjects WHERE (subStatus=1) ORDER BY subGroup"
        Me.cmdTTSubSetup.Parameters.Clear()
        reader = Me.cmdTTSubSetup.ExecuteReader
        While reader.Read
            Me.cboSubjectGroup.Items.Add(IIf(DBNull.Value.Equals(reader!subGroup), "", reader!subGroup))
        End While
        reader.Close()

        Me.cmdTTSubSetup.CommandText = "SELECT DISTINCT year FROM tblSchoolCalendar WHERE (Status=1) " & _
              vbNewLine & "  ORDER BY year"
        Me.cmdTTSubSetup.Parameters.Clear()
        reader = Me.cmdTTSubSetup.ExecuteReader
        While reader.Read
            Me.cboYear.Items.Add(IIf(DBNull.Value.Equals(reader!year), "", reader!year))
        End While
        reader.Close()

        Dim c As Color
        For Each value As KnownColor In [Enum].GetValues(GetType(KnownColor))
            c = Color.FromKnownColor(value)
            Me.cboSubjectColor.Items.Add(c.Name)
        Next value
    End Sub
    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub cboSubjectColor_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboSubjectColor.SelectedIndexChanged
        If Me.cboSubjectColor.Text.Trim.Length <= 0 Then
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            Me.txtSubColor.BackColor = Color.FromName(cboSubjectColor.Text.Trim)
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub

    Private Sub cboSubjectGroup_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboSubjectGroup.SelectedIndexChanged
        Me.cboSubjectName.Items.Clear()
        Me.cboSubjectName.Text = ""
        Me.cboSubjectName.SelectedIndex = -1

        Me.txtAbbreviation.Text = ""
        Me.txtSubjectCode.Text = ""
        Me.txtSubColor.BackColor = Color.WhiteSmoke
        Me.cboSubjectColor.Text = ""
        Me.cboSubjectColor.SelectedIndex = -1
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            Me.cmdTTSubSetup.Connection = conn
            Me.cmdTTSubSetup.CommandType = CommandType.Text
            Me.cmdTTSubSetup.CommandText = "SELECT DISTINCT subName FROM tblSubjects WHERE (subStatus=1) AND " & _
                vbNewLine & " (subGroup=@subGroup) ORDER BY subName"
            Me.cmdTTSubSetup.Parameters.Clear()
            Me.cmdTTSubSetup.Parameters.AddWithValue("@subGroup", Me.cboSubjectGroup.Text.Trim)
            reader = Me.cmdTTSubSetup.ExecuteReader
            While reader.Read
                Me.cboSubjectName.Items.Add(IIf(DBNull.Value.Equals(reader!subName), "", reader!subName))
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

    Private Sub cboSubjectName_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboSubjectName.SelectedIndexChanged, cboYear.SelectedIndexChanged
        Me.txtAbbreviation.Text = ""
        Me.txtSubjectCode.Text = ""
        Me.txtSubColor.BackColor = Color.WhiteSmoke
        Me.cboSubjectColor.Text = ""
        Me.cboSubjectColor.SelectedIndex = -1

        If Me.cboSubjectGroup.Text.Trim.Length <= 0 Then
            MsgBox("Subject Group is Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            Me.cmdTTSubSetup.Connection = conn
            Me.cmdTTSubSetup.CommandType = CommandType.Text
            Me.cmdTTSubSetup.CommandText = "SELECT * FROM tblSubjects WHERE (subStatus=1) AND " & _
                vbNewLine & " (subName=@subName)"
            Me.cmdTTSubSetup.Parameters.Clear()
            Me.cmdTTSubSetup.Parameters.AddWithValue("@subName", Me.cboSubjectName.Text.Trim)
            reader = Me.cmdTTSubSetup.ExecuteReader
            While reader.Read
                Me.txtSubjectCode.Text = (IIf(DBNull.Value.Equals(reader!subCode), "", reader!subCode))
                Me.txtAbbreviation.Text = (IIf(DBNull.Value.Equals(reader!abbr), "", reader!abbr))
            End While
            reader.Close()
            If Me.cboYear.Text.Trim.Length > 0 Then
                loadSubjSetup()
            End If
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub

    Private Sub dgSubjectSetUp_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgSubjectSetUp.CellClick

        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            If e.ColumnIndex <> 0 Then
                If Me.dgSubjectSetUp.Item(e.ColumnIndex, e.RowIndex).Value = "YES" Then
                    Me.dgSubjectSetUp.Item(e.ColumnIndex, e.RowIndex).Value = "NO"
                    Me.dgSubjectSetUp.Item(e.ColumnIndex, e.RowIndex).Style.BackColor = Color.Red
                ElseIf Me.dgSubjectSetUp.Item(e.ColumnIndex, e.RowIndex).Value = "NO" Then
                    Me.dgSubjectSetUp.Item(e.ColumnIndex, e.RowIndex).Value = "YES"
                    Me.dgSubjectSetUp.Item(e.ColumnIndex, e.RowIndex).Style.BackColor = Color.Green
                End If
            End If
            Me.dgSubjectSetUp.Rows(e.RowIndex).Cells(e.ColumnIndex).Selected = False
        Catch ex As Exception
            'MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub

    Private Sub btnUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdate.Click
        If Me.cboYear.Text.Trim.Length <= 0 Then
            MsgBox("Missing year.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboSubjectName.Text.Trim.Length <= 0 Then
            MsgBox("Missing subject name.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboSubjectColor.Text.Trim.Length <= 0 Then
            MsgBox("Missing subject color.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            Dim colourused As Boolean = checkColourIfUsed(Me.cboSubjectName.Text.Trim, Me.cboSubjectColor.Text.Trim)
            If colourused = True Then
                MsgBox("Colour used for a different subject.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
                Exit Sub
            End If
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            Dim result As MsgBoxResult = MsgBox("Save Record?", MsgBoxStyle.Question + MsgBoxStyle.YesNo + MsgBoxStyle.ApplicationModal, "Confirm Transaction")
            If result = MsgBoxResult.No Then
                Exit Sub
            End If
            i = 1
            j = 1
            For i = 1 To Me.dgSubjectSetUp.ColumnCount - 1
                j = 1
                For j = 1 To Me.dgSubjectSetUp.RowCount
                    Me.cmdTTSubSetup.Connection = conn
                    Me.cmdTTSubSetup.CommandType = CommandType.StoredProcedure
                    Me.cmdTTSubSetup.CommandText = "sprocTTSubjectSetUp"
                    Me.cmdTTSubSetup.Parameters.Clear()
                    Me.cmdTTSubSetup.Parameters.AddWithValue("@subjectName", Me.cboSubjectName.Text.Trim)
                    Me.cmdTTSubSetup.Parameters.AddWithValue("@acadYear", Me.cboYear.Text.Trim)
                    Me.cmdTTSubSetup.Parameters.AddWithValue("@subColourName", Me.cboSubjectColor.Text.Trim)
                    Me.cmdTTSubSetup.Parameters.AddWithValue("@lessonNo", Me.dgSubjectSetUp.Columns(i).HeaderText)
                    Me.cmdTTSubSetup.Parameters.AddWithValue("@weekDayNo", j)
                    Me.cmdTTSubSetup.Parameters.AddWithValue("@userName", userName.Trim)
                    Me.cmdTTSubSetup.Parameters.AddWithValue("@allowed", Me.dgSubjectSetUp.Item(i, j - 1).Value)
                    rec = rec + Me.cmdTTSubSetup.ExecuteNonQuery
                Next
            Next
            If rec > 0 Then
                MsgBox("Records Saved.", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "System Reply")
            End If
            Me.dgSubjectSetUp.Columns.Clear()
            Me.dgSubjectSetUp.Rows.Clear()
            loadCombos()
            Me.txtAbbreviation.Text = ""
            Me.txtSubjectCode.Text = ""
            Me.txtSubColor.BackColor = Color.WhiteSmoke
            Me.cboSubjectColor.Text = ""
            Me.cboSubjectColor.SelectedIndex = -1
            Me.cboSubjectName.Items.Clear()
            Me.cboSubjectName.Text = ""
            Me.cboSubjectName.SelectedIndex = -1
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub
    Private Function checkColourIfUsed(ByVal subname As String, ByVal subjName As String)
        Me.cmdTTSubSetup.Connection = conn
        Me.cmdTTSubSetup.CommandType = CommandType.Text
        Me.cmdTTSubSetup.CommandText = "SELECT DISTINCT subColourName FROM  vwTTSubjectSetUp WHERE (subName<>@subName) " & _
            vbNewLine & " AND (acadYear=@acadYear) AND (subColourName IS NOT NULL) AND (subColourName=@subColourName)"
        Me.cmdTTSubSetup.Parameters.Clear()
        Me.cmdTTSubSetup.Parameters.AddWithValue("@subColourName", subjName.Trim)
        Me.cmdTTSubSetup.Parameters.AddWithValue("@subName", subname.Trim)
        Me.cmdTTSubSetup.Parameters.AddWithValue("@acadYear", Me.cboYear.Text.Trim)
        reader = Me.cmdTTSubSetup.ExecuteReader
        If reader.HasRows = True Then
            Return True
        ElseIf reader.HasRows = False Then
            Return False
        End If
        reader.Close()
        conn.Close()
    End Function

    Private Sub dgSubjectSetUp_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgSubjectSetUp.CellContentClick

    End Sub
End Class