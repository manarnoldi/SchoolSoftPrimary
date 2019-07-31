Imports System.Data.SqlClient
Public Class frmTTStaffSetUp
    Dim reader As SqlDataReader
    Dim rec As Integer
    Dim cmdTTStaffSetup As New SqlCommand
    Dim rowNo As Integer
    Dim columnNo As Integer
    Private Sub frmTTStaffSetUp_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint
        If Me.IsMdiChild Then
            Dim PnlLoc As New Point
            PnlLoc.X = CInt((Me.Width - Me.pnlStaffSetup.Width) / 2)
            PnlLoc.Y = CInt((Me.Height - Me.pnlStaffSetup.Height) / 2.5)
            Me.pnlStaffSetup.Location = PnlLoc
        Else
            Me.pnlStaffSetup.Dock = DockStyle.Fill
        End If
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub
    Private Sub loadCombos()
        Me.cboTeacherNo.Items.Clear()
        Me.cboTeacherNo.Text = ""
        Me.cboTeacherNo.SelectedIndex = -1

        Me.cboAcadYear.Items.Clear()
        Me.cboAcadYear.Text = ""
        Me.cboAcadYear.SelectedIndex = -1

        Me.cmdTTStaffSetup.Connection = conn
        Me.cmdTTStaffSetup.CommandType = CommandType.Text
        Me.cmdTTStaffSetup.CommandText = "SELECT DISTINCT empNo FROM  tblSchoolStaff WHERE (Status=1) AND " & _
            vbNewLine & " (empType='Teaching') ORDER BY empNo"
        Me.cmdTTStaffSetup.Parameters.Clear()
        reader = Me.cmdTTStaffSetup.ExecuteReader
        While reader.Read
            Me.cboTeacherNo.Items.Add(IIf(DBNull.Value.Equals(reader!empNo), "", reader!empNo))
        End While
        reader.Close()

        Me.cmdTTStaffSetup.CommandText = "SELECT DISTINCT year FROM  tblSchoolCalendar WHERE (Status=1) ORDER BY year"
        Me.cmdTTStaffSetup.Parameters.Clear()
        reader = Me.cmdTTStaffSetup.ExecuteReader
        While reader.Read
            Me.cboAcadYear.Items.Add(IIf(DBNull.Value.Equals(reader!year), "", reader!year))
        End While
        reader.Close()
    End Sub
    Private Sub frmTTStaffSetUp_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
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

    Private Sub cboTeacherNo_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboTeacherNo.SelectedIndexChanged, cboAcadYear.SelectedIndexChanged
        Me.txtStaffName.Text = ""
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            Me.cmdTTStaffSetup.Connection = conn
            Me.cmdTTStaffSetup.CommandType = CommandType.Text
            Me.cmdTTStaffSetup.CommandText = "SELECT FullName FROM  tblSchoolStaff WHERE (Status=1) AND " & _
                vbNewLine & " (empType='Teaching') AND (empNo=@empNo)"
            Me.cmdTTStaffSetup.Parameters.Clear()
            Me.cmdTTStaffSetup.Parameters.AddWithValue("empNo", Me.cboTeacherNo.Text.Trim)
            reader = Me.cmdTTStaffSetup.ExecuteReader
            While reader.Read
                Me.txtStaffName.Text = (IIf(DBNull.Value.Equals(reader!FullName), "", reader!FullName))
            End While
            reader.Close()

            If Me.cboAcadYear.Text.Trim.Length > 0 Then
                loadStaffSetup()
            End If
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub
    Private Sub loadStaffSetup()
        Me.dgStaffSetUp.Rows.Clear()

        Me.cmdTTStaffSetup.Connection = conn
        Me.cmdTTStaffSetup.CommandType = CommandType.Text
        Me.cmdTTStaffSetup.CommandText = "SELECT * FROM  tblTTSetUp WHERE (academicYear=@year)"
        Me.cmdTTStaffSetup.Parameters.Clear()
        Me.cmdTTStaffSetup.Parameters.AddWithValue("@year", Me.cboAcadYear.Text.Trim)
        reader = Me.cmdTTStaffSetup.ExecuteReader
        While reader.Read
            rowNo = (IIf(DBNull.Value.Equals(reader!daysPerWeek), "", reader!daysPerWeek))
            columnNo = (IIf(DBNull.Value.Equals(reader!lessonsPerDay), "", reader!lessonsPerDay))
        End While
        reader.Close()
        Me.dgStaffSetUp.ColumnCount = columnNo + 1
        Me.dgStaffSetUp.Columns(0).Name = "days"
        Me.dgStaffSetUp.Columns(0).HeaderText = "Days Name"
        Me.dgStaffSetUp.Columns(0).Width = 55
        Me.dgStaffSetUp.Columns(0).ReadOnly = True
        Me.dgStaffSetUp.Columns(0).SortMode = DataGridViewColumnSortMode.NotSortable
        For i = 1 To columnNo
            Me.dgStaffSetUp.Columns(i).Name = "Lesson" & i
            Me.dgStaffSetUp.Columns(i).HeaderText = i
            Me.dgStaffSetUp.Columns(i).Width = 39
            Me.dgStaffSetUp.Columns(i).ReadOnly = True
            Me.dgStaffSetUp.Columns(i).SortMode = DataGridViewColumnSortMode.NotSortable
        Next
        j = 0
        For j = 0 To rowNo - 1
            If j = 0 Then
                Me.dgStaffSetUp.Rows.Add("MON")
            ElseIf j = 1 Then
                Me.dgStaffSetUp.Rows.Add("TUE")
            ElseIf j = 2 Then
                Me.dgStaffSetUp.Rows.Add("WED")
            ElseIf j = 3 Then
                Me.dgStaffSetUp.Rows.Add("THU")
            ElseIf j = 4 Then
                Me.dgStaffSetUp.Rows.Add("FRI")
            ElseIf j = 5 Then
                Me.dgStaffSetUp.Rows.Add("SAT")
            ElseIf j = 6 Then
                Me.dgStaffSetUp.Rows.Add("SUN")
            End If
        Next
        For i = 1 To Me.dgStaffSetUp.ColumnCount - 1
            j = 0
            For j = 0 To Me.dgStaffSetUp.RowCount - 1
                Me.cmdTTStaffSetup.Connection = conn
                Me.cmdTTStaffSetup.CommandType = CommandType.Text
                Me.cmdTTStaffSetup.CommandText = "SELECT allowed FROM vwTTStaffSetUp WHERE (academicYear=@academicYear) AND (empNo=@empNo) " & _
                    vbNewLine & " AND (lessonNo=@lessonNo) AND (weekDayNo=@weekDayNo)"
                Me.cmdTTStaffSetup.Parameters.Clear()
                Me.cmdTTStaffSetup.Parameters.AddWithValue("@academicYear", Me.cboAcadYear.Text.Trim)
                Me.cmdTTStaffSetup.Parameters.AddWithValue("@empNo", Me.cboTeacherNo.Text.Trim)
                Me.cmdTTStaffSetup.Parameters.AddWithValue("@lessonNo", i)
                Me.cmdTTStaffSetup.Parameters.AddWithValue("@weekDayNo", j + 1)
                reader = Me.cmdTTStaffSetup.ExecuteReader
                While reader.Read
                    Me.dgStaffSetUp.Item(i, j).Value = (IIf(DBNull.Value.Equals(reader!allowed), "", reader!allowed))
                    If (IIf(DBNull.Value.Equals(reader!allowed), "", reader!allowed)) = "YES" Then
                        Me.dgStaffSetUp.Item(i, j).Style.BackColor = Color.Green
                    ElseIf (IIf(DBNull.Value.Equals(reader!allowed), "", reader!allowed)) = "NO" Then
                        Me.dgStaffSetUp.Item(i, j).Style.BackColor = Color.Red
                    End If
                End While
                reader.Close()
            Next
        Next
    End Sub

    Private Sub dgStaffSetUp_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgStaffSetUp.CellClick
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            If e.ColumnIndex <> 0 Then
                If Me.dgStaffSetUp.Item(e.ColumnIndex, e.RowIndex).Value = "YES" Then
                    Me.dgStaffSetUp.Item(e.ColumnIndex, e.RowIndex).Value = "NO"
                    Me.dgStaffSetUp.Item(e.ColumnIndex, e.RowIndex).Style.BackColor = Color.Red
                ElseIf Me.dgStaffSetUp.Item(e.ColumnIndex, e.RowIndex).Value = "NO" Then
                    Me.dgStaffSetUp.Item(e.ColumnIndex, e.RowIndex).Value = "YES"
                    Me.dgStaffSetUp.Item(e.ColumnIndex, e.RowIndex).Style.BackColor = Color.Green
                End If
            End If
            Me.dgStaffSetUp.Rows(e.RowIndex).Cells(e.ColumnIndex).Selected = False
        Catch ex As Exception
       Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub

    Private Sub btnUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdate.Click
        If Me.cboAcadYear.Text.Trim.Length <= 0 Then
            MsgBox("Missing year.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboTeacherNo.Text.Trim.Length <= 0 Then
            MsgBox("Missing staff Number.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If
        Try
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
            For i = 1 To Me.dgStaffSetUp.ColumnCount - 1
                j = 1
                For j = 1 To Me.dgStaffSetUp.RowCount
                    Me.cmdTTStaffSetup.Connection = conn
                    Me.cmdTTStaffSetup.CommandType = CommandType.StoredProcedure
                    Me.cmdTTStaffSetup.CommandText = "sprocTTStaffSetUp"
                    Me.cmdTTStaffSetup.Parameters.Clear()
                    Me.cmdTTStaffSetup.Parameters.AddWithValue("@empNo", Me.cboTeacherNo.Text.Trim)
                    Me.cmdTTStaffSetup.Parameters.AddWithValue("@acadYear", Me.cboAcadYear.Text.Trim)
                    Me.cmdTTStaffSetup.Parameters.AddWithValue("@lessonNo", Me.dgStaffSetUp.Columns(i).HeaderText)
                    Me.cmdTTStaffSetup.Parameters.AddWithValue("@weekDayNo", j)
                    Me.cmdTTStaffSetup.Parameters.AddWithValue("@userName", userName.Trim)
                    Me.cmdTTStaffSetup.Parameters.AddWithValue("@allowed", Me.dgStaffSetUp.Item(i, j - 1).Value)
                    rec = rec + Me.cmdTTStaffSetup.ExecuteNonQuery
                Next
            Next
            If rec > 0 Then
                MsgBox("Records Saved.", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "System Reply")
            End If
            Me.dgStaffSetUp.Columns.Clear()
            Me.dgStaffSetUp.Rows.Clear()
            loadCombos()
            Me.txtStaffName.Text = ""
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub
End Class