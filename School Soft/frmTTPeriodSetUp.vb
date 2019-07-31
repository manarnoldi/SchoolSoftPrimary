Imports System.Data.SqlClient
Public Class frmTTPeriodSetUp
    Dim lessonNo As Integer = 0
    Dim breaksPerDay As Integer = 0
    Dim cmdPeriod As New SqlCommand
    Dim reader As SqlDataReader
    Dim rec As Integer
    Private Sub frmTTPeriodSetUp_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            loadCombos()
            Me.dtpStart.Text = "07:00:00 AM"
            Me.dtpEnd.Text = "07:00:00 AM"
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub
    Private Sub loadCombos()
        Me.cboYear.Items.Clear()
        Me.cboYear.Text = ""
        Me.cboYear.SelectedIndex = -1

        Me.cmdPeriod.Connection = conn
        Me.cmdPeriod.CommandType = CommandType.Text
        Me.cmdPeriod.CommandText = "SELECT DISTINCT Year FROM tblSchoolCalendar ORDER BY Year"
        Me.cmdPeriod.Parameters.Clear()
        reader = Me.cmdPeriod.ExecuteReader
        While reader.Read
            Me.cboYear.Items.Add(IIf(DBNull.Value.Equals(reader!Year), "", reader!Year))
        End While
        reader.Close()

    End Sub
    Private Sub frmTTPeriodSetUp_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint
        If Me.IsMdiChild Then
            Dim PnlLoc As New Point
            PnlLoc.X = CInt((Me.Width - Me.pnlPeriodSetup.Width) / 2)
            PnlLoc.Y = CInt((Me.Height - Me.pnlPeriodSetup.Height) / 2.5)
            Me.pnlPeriodSetup.Location = PnlLoc
        Else
            Me.pnlPeriodSetup.Dock = DockStyle.Fill
        End If
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub cboPeriodType_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboPeriodType.SelectedIndexChanged
        Me.cboPeriodName.Items.Clear()
        Me.cboPeriodName.Text = ""
        Me.cboPeriodName.SelectedIndex = -1
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            If Me.cboPeriodType.Text.Trim = "" Then
                Exit Sub
            ElseIf Me.cboPeriodType.Text.Trim = "Lesson" Then
                Me.cmdPeriod.Connection = conn
                Me.cmdPeriod.CommandType = CommandType.Text
                Me.cmdPeriod.CommandText = "SELECT lessonsPerDay FROM tblTTSetUp"
                Me.cmdPeriod.Parameters.Clear()
                reader = Me.cmdPeriod.ExecuteReader
                While reader.Read
                    lessonNo = (IIf(DBNull.Value.Equals(reader!lessonsPerDay), "", reader!lessonsPerDay))
                End While
                reader.Close()
                If lessonNo > 0 Then
                    For i = 1 To lessonNo
                        Me.cboPeriodName.Items.Add("Lesson " & i)
                    Next
                End If
            ElseIf Me.cboPeriodType.Text = "Break" Then
                Me.cmdPeriod.Connection = conn
                Me.cmdPeriod.CommandType = CommandType.Text
                Me.cmdPeriod.CommandText = "SELECT breaksPerDay FROM tblTTSetUp"
                Me.cmdPeriod.Parameters.Clear()
                reader = Me.cmdPeriod.ExecuteReader
                While reader.Read
                    breaksPerDay = (IIf(DBNull.Value.Equals(reader!breaksPerDay), "", reader!breaksPerDay))
                End While
                reader.Close()
                If breaksPerDay > 0 Then
                    For j = 1 To breaksPerDay
                        Me.cboPeriodName.Items.Add("Break " & j)
                    Next
                End If
            ElseIf Me.cboPeriodType.Text = "Lunch" Then
                Me.cboPeriodName.Items.Add("Lunch")
            End If
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub

    Private Sub btnAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAdd.Click
        If Me.cboYear.Text.Trim.Length <= 0 Then
            MsgBox("Academic Year is missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboPeriodType.Text.Trim.Length <= 0 Then
            MsgBox("Period Type is missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboPeriodNo.Text.Trim.Length <= 0 Then
            MsgBox("Period Number is missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboPeriodName.Text.Trim.Length <= 0 Then
            MsgBox("Period Name is missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.dtpStart.Text = Me.dtpEnd.Text Then
            MsgBox("Start and End date Cannot be the same", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.dtpStart.Value > Me.dtpEnd.Value Then
            MsgBox("Start Date Cannot be Greater than End date", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If
        For i = 0 To Me.lstPeriodSetUp.Items.Count - 1
            If Me.lstPeriodSetUp.Items(i).SubItems(1).Text.Trim = Me.cboPeriodType.Text.Trim And _
                Me.lstPeriodSetUp.Items(i).SubItems(2).Text.Trim = Me.cboPeriodName.Text.Trim And _
                Me.lstPeriodSetUp.Items(i).SubItems(3).Text.Trim = Me.dtpStart.Text.Trim And _
                Me.lstPeriodSetUp.Items(i).SubItems(4).Text.Trim = Me.dtpEnd.Text.Trim Then
                MsgBox("Record already in the list.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
                Exit Sub
            ElseIf Me.lstPeriodSetUp.Items(i).SubItems(1).Text.Trim = Me.cboPeriodName.Text.Trim Then
                MsgBox("Period Name Already Exists.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
                Exit Sub
            End If
        Next
        li = Me.lstPeriodSetUp.Items.Add(Me.cboPeriodNo.Text.Trim)
        li.SubItems.Add(Me.cboPeriodType.Text.Trim)
        li.SubItems.Add(Me.cboPeriodName.Text.Trim)
        li.SubItems.Add(Me.dtpStart.Text)
        li.SubItems.Add(Me.dtpEnd.Text)
    End Sub

    Private Sub btnRemove_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRemove.Click
        If Me.lstPeriodSetUp.Items.Count <= 0 Then
            MsgBox("Missing Items in the list.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.lstPeriodSetUp.SelectedItems.Count <= 0 Then
            MsgBox("Missing Selected Items in the list.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If
        For i = 0 To Me.lstPeriodSetUp.SelectedItems.Count - 1
            Me.lstPeriodSetUp.SelectedItems(0).Remove()
        Next
    End Sub

    Private Sub btnView_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnView.Click
        Me.lstPeriodSetUp.Items.Clear()
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            loadLists()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub
    Private Sub loadLists()
        Me.cmdPeriod.Connection = conn
        Me.cmdPeriod.CommandType = CommandType.Text
        Me.cmdPeriod.CommandText = "SELECT * FROM tblTTPeriods ORDER BY startTime"
        Me.cmdPeriod.Parameters.Clear()
        reader = Me.cmdPeriod.ExecuteReader
        While reader.Read
            Dim startDate As String = Nothing
            Dim endDate As String = Nothing
            If CStr(IIf(DBNull.Value.Equals(reader!startTimeShow), "", reader!startTimeShow)).Length = 6 Then
                startDate = "0" & IIf(DBNull.Value.Equals(reader!startTimeShow), "", reader!startTimeShow)
            ElseIf CStr(IIf(DBNull.Value.Equals(reader!startTimeShow), "", reader!startTimeShow)).Length = 7 Then
                startDate = IIf(DBNull.Value.Equals(reader!startTimeShow), "", reader!startTimeShow)
            End If
            If CStr(IIf(DBNull.Value.Equals(reader!endTimeShow), "", reader!endTimeShow)).Length = 6 Then
                endDate = "0" & IIf(DBNull.Value.Equals(reader!endTimeShow), "", reader!endTimeShow)
            ElseIf CStr(IIf(DBNull.Value.Equals(reader!endTimeShow), "", reader!endTimeShow)).Length = 7 Then
                endDate = IIf(DBNull.Value.Equals(reader!endTimeShow), "", reader!endTimeShow)
            End If


            li = Me.lstPeriodSetUp.Items.Add(IIf(DBNull.Value.Equals(reader!periodNo), "", reader!periodNo))
            li.SubItems.Add(IIf(DBNull.Value.Equals(reader!periodType), "", reader!periodType))
            li.SubItems.Add(IIf(DBNull.Value.Equals(reader!periodName), "", reader!periodName))
            li.SubItems.Add(startDate)
            li.SubItems.Add(endDate)
        End While
        reader.Close()
    End Sub
    
    Private Sub btnUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdate.Click
        If Me.cboYear.Text.Trim.Length <= 0 Then
            MsgBox("Academic Year is missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If
        Dim result As MsgBoxResult = MsgBox("Update Record/s?", MsgBoxStyle.Question + MsgBoxStyle.ApplicationModal + MsgBoxStyle.YesNo, "Confirm Transaction")
        If result = MsgBoxResult.No Then
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            Me.cmdPeriod.Connection = conn
            Me.cmdPeriod.CommandType = CommandType.Text
            Me.cmdPeriod.CommandText = "DELETE FROM tblTTPeriods WHERE (academicYear=@academicYear)"
            Me.cmdPeriod.Parameters.Clear()
            Me.cmdPeriod.Parameters.AddWithValue("@academicYear", Me.cboYear.Text.Trim)
            Me.cmdPeriod.ExecuteNonQuery()
            rec = 0
            i = 0
            For i = 0 To Me.lstPeriodSetUp.Items.Count - 1
                Me.cmdPeriod.Connection = conn
                Me.cmdPeriod.CommandType = CommandType.StoredProcedure
                Me.cmdPeriod.CommandText = "sprocTTPeriodSetUp"
                Me.cmdPeriod.Parameters.Clear()
                Me.cmdPeriod.Parameters.AddWithValue("@periodNo", Me.lstPeriodSetUp.Items(i).Text.Trim)
                Me.cmdPeriod.Parameters.AddWithValue("@periodType", Me.lstPeriodSetUp.Items(i).SubItems(1).Text.Trim)
                Me.cmdPeriod.Parameters.AddWithValue("@periodName", Me.lstPeriodSetUp.Items(i).SubItems(2).Text.Trim)
                Me.cmdPeriod.Parameters.AddWithValue("@startTime", Me.lstPeriodSetUp.Items(i).SubItems(3).Text.Trim)
                Me.cmdPeriod.Parameters.AddWithValue("@endTime", Me.lstPeriodSetUp.Items(i).SubItems(4).Text.Trim)
                Me.cmdPeriod.Parameters.AddWithValue("@academicYear", Me.cboYear.Text.Trim)
                Me.cmdPeriod.Parameters.AddWithValue("@userName", userName.Trim)
                rec = rec + Me.cmdPeriod.ExecuteNonQuery
            Next
            If rec > 0 Then
                MsgBox("Record/s Updated!", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "SuccessFull Transaction")
            End If
            Me.lstPeriodSetUp.Items.Clear()
            loadCombos()
            Me.cboPeriodName.Text = ""
            Me.cboPeriodName.SelectedIndex = -1
            Me.cboPeriodType.Text = ""
            Me.cboPeriodType.SelectedIndex = -1
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub

    Private Sub cboYear_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboYear.SelectedIndexChanged
        Dim totalPeriods As Integer = 0
        If Me.cboYear.Text.Trim.Length <= 0 Then
            Exit Sub
        End If
        Me.cboPeriodNo.Items.Clear()
        Me.cboPeriodNo.Text = ""
        Me.cboPeriodNo.SelectedIndex = -1
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            Me.cmdPeriod.Connection = conn
            Me.cmdPeriod.CommandType = CommandType.Text
            Me.cmdPeriod.CommandText = "SELECT (lessonsPerDay+breaksPerDay)+1 AS totalPeriods FROM tblTTSetUp WHERE (academicYear=@academicYear)"
            Me.cmdPeriod.Parameters.Clear()
            Me.cmdPeriod.Parameters.AddWithValue("@academicYear", Me.cboYear.Text.Trim)
            reader = Me.cmdPeriod.ExecuteReader
            While reader.Read
                totalPeriods = (IIf(DBNull.Value.Equals(reader!totalPeriods), "", reader!totalPeriods))
            End While
            reader.Close()
            For i = 1 To totalPeriods
                Me.cboPeriodNo.Items.Add(i)
            Next
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub

    Private Sub cboPeriodName_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboPeriodName.SelectedIndexChanged

    End Sub
End Class