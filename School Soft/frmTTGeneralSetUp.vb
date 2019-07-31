Imports System.Data.SqlClient
Public Class frmTTGeneralSetUp
    Dim reader As SqlDataReader
    Dim rec As Integer
    Dim queryType As String = Nothing
    Dim cmdGeneral As New SqlCommand
    Private Sub frmTTGeneralSetUp_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint
        If Me.IsMdiChild Then
            Dim PnlLoc As New Point
            PnlLoc.X = CInt((Me.Width - Me.pnlTTGeneralSetUp.Width) / 2)
            PnlLoc.Y = CInt((Me.Height - Me.pnlTTGeneralSetUp.Height) / 2.5)
            Me.pnlTTGeneralSetUp.Location = PnlLoc
        Else
            Me.pnlTTGeneralSetUp.Dock = DockStyle.Fill
        End If
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub frmTTGeneralSetUp_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
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
    Private Sub loadCombos()
        Me.cboAcadYear.Items.Clear()
        Me.cboAcadYear.Text = ""
        Me.cboAcadYear.SelectedIndex = -1


        Me.cmdGeneral.Connection = conn
        Me.cmdGeneral.CommandType = CommandType.Text
        Me.cmdGeneral.CommandText = "SELECT DISTINCT Year FROM tblSchoolCalendar ORDER BY Year"
        Me.cmdGeneral.Parameters.Clear()
        reader = Me.cmdGeneral.ExecuteReader
        While reader.Read
            Me.cboAcadYear.Items.Add(IIf(DBNull.Value.Equals(reader!Year), "", reader!Year))
        End While
        reader.Close()
    End Sub

    Private Sub cboDayWeek_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboDayWeek.SelectedIndexChanged
        Me.txtDaysRange.Text = ""
        Select Case Me.cboDayWeek.Text.Trim
            Case ""
                Exit Sub
            Case "5"
                Me.txtDaysRange.Text = "Monday-Friday"
            Case "6"
                Me.txtDaysRange.Text = "Monday-Saturday"
            Case "7"
                Me.txtDaysRange.Text = "Monday-Sunday"
        End Select
    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        If Me.cboAcadYear.Text.Trim.Length <= 0 Then
            MsgBox("Year Is Missing", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboDayWeek.Text.Trim.Length <= 0 Then
            MsgBox("Days per week Missing", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboLessonDay.Text.Trim.Length <= 0 Then
            MsgBox("Lessons Per Day Missing", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboBreaksPerDay.Text.Trim.Length <= 0 Then
            MsgBox("Breaks Per Day Missing", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.txtDaysRange.Text.Trim.Length <= 0 Then
            MsgBox("Days Range Missing", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            Dim recordExists As Boolean = checkForExistenceSave(Me.cboAcadYear.Text.Trim)
            If recordExists = True Then
                MsgBox("Duplicate Record Found.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
                Exit Sub
            End If
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            Dim result As MsgBoxResult = MsgBox("Save Record?", MsgBoxStyle.Question + MsgBoxStyle.ApplicationModal + MsgBoxStyle.YesNo, "Confirm Transaction")
            If result = MsgBoxResult.No Then
                Exit Sub
            End If
            queryType = "INSERT"
            Me.cmdGeneral.Connection = conn
            Me.cmdGeneral.CommandType = CommandType.StoredProcedure
            Me.cmdGeneral.CommandText = "sprocTTGeneralSetUp"
            Me.cmdGeneral.Parameters.Clear()
            Me.cmdGeneral.Parameters.AddWithValue("@lessonsPerDay", Me.cboLessonDay.Text.Trim)
            Me.cmdGeneral.Parameters.AddWithValue("@daysPerWeek", Me.cboDayWeek.Text.Trim)
            Me.cmdGeneral.Parameters.AddWithValue("@academicYear", Me.cboAcadYear.Text.Trim)
            Me.cmdGeneral.Parameters.AddWithValue("@daysRange", Me.txtDaysRange.Text.Trim)
            Me.cmdGeneral.Parameters.AddWithValue("@userName", userName.Trim)
            Me.cmdGeneral.Parameters.AddWithValue("@dateOfReg", Date.Now)
            Me.cmdGeneral.Parameters.AddWithValue("@queryType", Me.queryType.Trim)
            Me.cmdGeneral.Parameters.AddWithValue("@breaksPerDay", Me.cboBreaksPerDay.Text.Trim)
            rec = Me.cmdGeneral.ExecuteNonQuery

            Me.cmdGeneral.CommandText = "sprocTTStaffGenSetup"
            Me.cmdGeneral.Parameters.Clear()
            Me.cmdGeneral.Parameters.AddWithValue("@lessonsPerDay", Me.cboLessonDay.Text.Trim)
            Me.cmdGeneral.Parameters.AddWithValue("@daysPerWeek", Me.cboDayWeek.Text.Trim)
            Me.cmdGeneral.Parameters.AddWithValue("@academicYear", Me.cboAcadYear.Text.Trim)
            Me.cmdGeneral.Parameters.AddWithValue("@userName", userName.Trim)
            Me.cmdGeneral.Parameters.AddWithValue("@dateOfReg", Date.Now)
            rec = Me.cmdGeneral.ExecuteNonQuery
            If rec > 0 Then
                MsgBox("Record Saved!", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "SuccessFull Transaction")
            End If
            clearTexts()
            loadCombos()
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub
    Private Function checkForExistenceSave(ByVal acadYear As Integer)
        Me.cmdGeneral.Connection = conn
        Me.cmdGeneral.CommandType = CommandType.Text
        Me.cmdGeneral.CommandText = "SELECT * FROM  tblTTSetUp WHERE (academicYear=@academicYear)"
        Me.cmdGeneral.Parameters.Clear()
        Me.cmdGeneral.Parameters.AddWithValue("@academicYear", acadYear)
        reader = Me.cmdGeneral.ExecuteReader
        If reader.HasRows = True Then
            reader.Close()
            conn.Close()
            Return True
        ElseIf reader.HasRows = False Then
            reader.Close()
            conn.Close()
            Return False
        End If
    End Function
    Private Function checkForExistenceUpdate(ByVal acadYear As Integer, ByVal dayWeek As Integer, ByVal lessonDay As Integer, ByVal dayRange As String, ByVal breaksPerDay As Integer)
        Me.cmdGeneral.Connection = conn
        Me.cmdGeneral.CommandType = CommandType.Text
        Me.cmdGeneral.CommandText = "SELECT * FROM  tblTTSetUp WHERE (academicYear=@academicYear) AND (breaksPerDay=@breaksPerDay)" & _
            vbNewLine & " AND (lessonsPerDay=@lessonsPerDay) AND (daysPerWeek=@daysPerWeek) AND (daysRange=@daysRange)"
        Me.cmdGeneral.Parameters.Clear()
        Me.cmdGeneral.Parameters.AddWithValue("@academicYear", acadYear)
        Me.cmdGeneral.Parameters.AddWithValue("@lessonsPerDay", lessonDay)
        Me.cmdGeneral.Parameters.AddWithValue("@daysPerWeek", dayWeek)
        Me.cmdGeneral.Parameters.AddWithValue("@daysRange", dayRange.Trim)
        Me.cmdGeneral.Parameters.AddWithValue("@breaksPerDay", breaksPerDay)
        reader = Me.cmdGeneral.ExecuteReader
        If reader.HasRows = True Then
            reader.Close()
            conn.Close()
            Return True
        ElseIf reader.HasRows = False Then
            reader.Close()
            conn.Close()
            Return False
        End If
    End Function
    Private Sub clearTexts()
        Me.cboDayWeek.Text = ""
        Me.cboDayWeek.SelectedIndex = -1
        Me.cboLessonDay.Text = ""
        Me.cboLessonDay.SelectedIndex = -1
        Me.cboBreaksPerDay.Text = ""
        Me.cboBreaksPerDay.SelectedIndex = -1
        Me.txtDaysRange.Text = ""
    End Sub

    Private Sub btnUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdate.Click
        If Me.cboAcadYear.Text.Trim.Length <= 0 Then
            MsgBox("Year Is Missing", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboDayWeek.Text.Trim.Length <= 0 Then
            MsgBox("Days per week Missing", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboLessonDay.Text.Trim.Length <= 0 Then
            MsgBox("Lessons Per Day Missing", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.txtDaysRange.Text.Trim.Length <= 0 Then
            MsgBox("Days Range Missing", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboBreaksPerDay.Text.Trim.Length <= 0 Then
            MsgBox("Breaks Per Day Missing", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            Dim updateRecordExists As Boolean = checkForExistenceSave(Me.cboAcadYear.Text.Trim)
            If updateRecordExists = False Then
                MsgBox("Year to Update Not there.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
                Exit Sub
            End If

            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()

            Dim recordExists As Boolean = checkForExistenceUpdate(Me.cboAcadYear.Text.Trim, Me.cboDayWeek.Text.Trim, _
                                                                Me.cboLessonDay.Text.Trim, Me.txtDaysRange.Text.Trim, _
                                                                Me.cboBreaksPerDay.Text.Trim)
            If recordExists = True Then
                MsgBox("Duplicate Record Found.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
                Exit Sub
            End If
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            Dim result As MsgBoxResult = MsgBox("Update Record?", MsgBoxStyle.Question + MsgBoxStyle.ApplicationModal + MsgBoxStyle.YesNo, "Confirm Transaction")
            If result = MsgBoxResult.No Then
                Exit Sub
            End If
            queryType = "UPDATE"
            Me.cmdGeneral.Connection = conn
            Me.cmdGeneral.CommandType = CommandType.StoredProcedure
            Me.cmdGeneral.CommandText = "sprocTTGeneralSetUp"
            Me.cmdGeneral.Parameters.Clear()
            Me.cmdGeneral.Parameters.AddWithValue("@lessonsPerDay", Me.cboLessonDay.Text.Trim)
            Me.cmdGeneral.Parameters.AddWithValue("@daysPerWeek", Me.cboDayWeek.Text.Trim)
            Me.cmdGeneral.Parameters.AddWithValue("@academicYear", Me.cboAcadYear.Text.Trim)
            Me.cmdGeneral.Parameters.AddWithValue("@daysRange", Me.txtDaysRange.Text.Trim)
            Me.cmdGeneral.Parameters.AddWithValue("@userName", userName.Trim)
            Me.cmdGeneral.Parameters.AddWithValue("@dateOfReg", Date.Now)
            Me.cmdGeneral.Parameters.AddWithValue("@queryType", Me.queryType.Trim)
            Me.cmdGeneral.Parameters.AddWithValue("@breaksPerDay", Me.cboBreaksPerDay.Text.Trim)
            rec = Me.cmdGeneral.ExecuteNonQuery

            Me.cmdGeneral.CommandText = "sprocTTStaffGenSetup"
            Me.cmdGeneral.Parameters.Clear()
            Me.cmdGeneral.Parameters.AddWithValue("@lessonsPerDay", Me.cboLessonDay.Text.Trim)
            Me.cmdGeneral.Parameters.AddWithValue("@daysPerWeek", Me.cboDayWeek.Text.Trim)
            Me.cmdGeneral.Parameters.AddWithValue("@academicYear", Me.cboAcadYear.Text.Trim)
            Me.cmdGeneral.Parameters.AddWithValue("@userName", userName.Trim)
            Me.cmdGeneral.Parameters.AddWithValue("@dateOfReg", Date.Now)
            rec = Me.cmdGeneral.ExecuteNonQuery
            If rec > 0 Then
                MsgBox("Record Update!", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "SuccessFull Transaction")
            End If
            clearTexts()
            loadCombos()
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub

End Class