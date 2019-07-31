Imports System.Data.SqlClient
Imports System.Globalization
Public Class frmSchTermDates
    Dim rec As Integer = 0
    Dim recordExists As Boolean = False
    Dim queryType As String = Nothing
    Dim reader As SqlDataReader
    Dim cmdTermDates As New SqlCommand
    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub frmSchTermDates_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            loadTCombos()
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub
    Private Sub loadTCombos()
        dbconnection()
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If
        Me.cboTermYear.Items.Clear()
        Me.cboTermName.Items.Clear()
        Me.cboTermYear.Text = ""
        Me.cboTermName.Text = ""

        cmdTermDates.Connection = conn
        cmdTermDates.CommandType = CommandType.Text
        cmdTermDates.CommandText = "SELECT DISTINCT termName FROM tblSchoolCalendar WHERE (status='True') ORDER BY termName"
        cmdTermDates.Parameters.Clear()
        reader = cmdTermDates.ExecuteReader
        If reader.HasRows Then
            While reader.Read
                cboTermName.Items.Add(IIf(DBNull.Value.Equals(reader!termName), "", reader!termName))
            End While
        End If
        reader.Close()

        cmdTermDates.CommandText = "SELECT DISTINCT year FROM tblSchoolCalendar WHERE (status='True') ORDER BY Year"
        cmdTermDates.Parameters.Clear()
        reader = cmdTermDates.ExecuteReader
        If reader.HasRows Then
            While reader.Read
                cboTermYear.Items.Add(IIf(DBNull.Value.Equals(reader!year), "", reader!year))
            End While
        End If
        reader.Close()
    End Sub
    Private Sub frmSchTermDates_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint
        If Me.IsMdiChild Then
            Dim PnlLoc As New Point
            PnlLoc.X = CInt((Me.Width - Me.pnlSchTermDates.Width) / 2)
            PnlLoc.Y = CInt((Me.Height - Me.pnlSchTermDates.Height) / 2.5)
            Me.pnlSchTermDates.Location = PnlLoc
        Else
            Me.pnlSchTermDates.Dock = DockStyle.Fill
        End If
    End Sub

    Private Sub btnView_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnView.Click
        If Me.cboTermYear.Text.Trim.Length <= 0 Then
            MsgBox("Year Missing!", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If
        Try
            loadTermDates(Me.cboTermYear.Text.Trim)
            Me.cboTermName.Text = ""
            Me.cboTermYear.Text = ""
            Me.dtpDateBeg.Value = Date.Now
            Me.dtpDateClosing.Value = Date.Now

        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
            Me.btnSave.Enabled = True
            'Me.btnDelete.Enabled = False
            Me.btnUpdate.Enabled = False
        End Try
    End Sub
    Private Function loadTermDates(ByVal termYear As Integer)
        dbconnection()
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If
        Me.lstTermDates.Items.Clear()
        cmdTermDates.Connection = conn
        cmdTermDates.CommandType = CommandType.Text
        cmdTermDates.CommandText = "SELECT * FROM tblSchoolCalendar WHERE (status='True') AND (year=@year) ORDER BY year,termName"
        cmdTermDates.Parameters.Clear()
        Me.cmdTermDates.Parameters.AddWithValue("@year", termYear)
        reader = cmdTermDates.ExecuteReader
        If reader.HasRows Then
            While reader.Read
                li = Me.lstTermDates.Items.Add(IIf(DBNull.Value.Equals(reader!termName), "", reader!termName))
                li.SubItems.Add(IIf(DBNull.Value.Equals(reader!year), "", reader!year))
                Dim DateOpen As String = CDate(IIf(DBNull.Value.Equals(reader!dateBeginning), "", reader!dateBeginning)).Day.ToString("00") & "-" & _
                 CDate(IIf(DBNull.Value.Equals(reader!dateBeginning), "", reader!dateBeginning)).Month.ToString("00") & "-" & _
                CDate(IIf(DBNull.Value.Equals(reader!dateBeginning), "", reader!dateBeginning)).Year.ToString("0000")
                Dim DateClose As String = CDate(IIf(DBNull.Value.Equals(reader!dateClosing), "", reader!dateClosing)).Day.ToString("00") & "-" & _
                 CDate(IIf(DBNull.Value.Equals(reader!dateClosing), "", reader!dateClosing)).Month.ToString("00") & "-" & _
                CDate(IIf(DBNull.Value.Equals(reader!dateClosing), "", reader!dateClosing)).Year.ToString("0000")
                li.SubItems.Add(DateOpen)
                li.SubItems.Add(DateClose)
                li.SubItems(2).Tag = IIf(DBNull.Value.Equals(reader!dateBeginning), "", reader!dateBeginning)
                li.SubItems(3).Tag = IIf(DBNull.Value.Equals(reader!dateClosing), "", reader!dateClosing)
                li.Tag = IIf(DBNull.Value.Equals(reader!termId), "", reader!termId)
            End While
        End If
        reader.Close()
    End Function

    Private Sub lstTermDates_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lstTermDates.Click
        If Me.lstTermDates.SelectedItems.Count = 1 Then
            Me.cboTermName.Text = Me.lstTermDates.SelectedItems(0).Text
            Me.cboTermYear.Text = Me.lstTermDates.SelectedItems(0).SubItems(1).Text
            Me.dtpDateBeg.Value = Me.lstTermDates.SelectedItems(0).SubItems(2).Tag
            Me.dtpDateClosing.Value = Me.lstTermDates.SelectedItems(0).SubItems(3).Tag
            Me.cboTermName.Tag = Me.lstTermDates.SelectedItems(0).Tag
            Me.btnSave.Enabled = False
            'Me.btnDelete.Enabled = True
            Me.btnUpdate.Enabled = True
        End If
    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        If Me.cboTermName.Text.Trim.Length <= 0 Then
            MsgBox("Term Missing!", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboTermYear.Text.Trim.Length <= 0 Then
            MsgBox("Year Missing!", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If
        Try
            queryType = "INSERT"
            dbconnection()
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            checkForExistence()
            If recordExists = True Then
                MsgBox("Record Exists!", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "UNSuccessFul Transaction")
                Exit Sub
            End If
            Dim result As MsgBoxResult = MsgBox("Save Record?", MsgBoxStyle.Question + MsgBoxStyle.ApplicationModal + MsgBoxStyle.YesNo, "Confirm Transaction")
            If result = MsgBoxResult.Yes Then
                cmdTermDates.Connection = conn
                cmdTermDates.CommandType = CommandType.StoredProcedure
                cmdTermDates.CommandText = "sprocTermDates"
                cmdTermDates.Parameters.Clear()
                cmdTermDates.Parameters.AddWithValue("@termName", Me.cboTermName.Text.Trim)
                cmdTermDates.Parameters.AddWithValue("@year", Me.cboTermYear.Text.Trim)
                cmdTermDates.Parameters.AddWithValue("@queryType", queryType.Trim)
                cmdTermDates.Parameters.AddWithValue("@userName", userName.Trim)
                cmdTermDates.Parameters.AddWithValue("@regDate", Date.Now)
                cmdTermDates.Parameters.AddWithValue("@termBeginning", Me.dtpDateBeg.Value.Date)
                cmdTermDates.Parameters.AddWithValue("@termEnding", Me.dtpDateClosing.Value.Date)
                cmdTermDates.Parameters.AddWithValue("@termId", Me.cboTermName.Tag)
                rec = cmdTermDates.ExecuteNonQuery
                If rec > 0 Then
                    MsgBox("Record Saved!", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "SuccessFul Transaction")
                End If
            End If
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            Me.lstTermDates.Items.Clear()
            loadTermDates(Me.cboTermYear.Text.Trim)
            Me.cboTermName.Text = ""
            Me.cboTermName.Tag = Nothing

            Me.dtpDateClosing.Value = Date.Now
            Me.dtpDateBeg.Value = Date.Now
            queryType = Nothing
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub
    Private Sub checkForExistence()
        If conn.State = ConnectionState.Open Then
            conn.Close()
        End If
        dbconnection()
        cmdTermDates.Connection = conn
        cmdTermDates.CommandType = CommandType.Text
        cmdTermDates.CommandText = "SELECT * FROM tblSchoolCalendar WHERE (status='True') AND (termName=@termName) AND (year=@year)"
        cmdTermDates.Parameters.Clear()
        cmdTermDates.Parameters.AddWithValue("@termName", Me.cboTermName.Text)
        cmdTermDates.Parameters.AddWithValue("@year", Me.cboTermYear.Text)
        reader = cmdTermDates.ExecuteReader
        If reader.HasRows Then
            recordExists = True
        Else
            recordExists = False
        End If
        reader.Close()
    End Sub
    Private Sub checkForExistenceOne()
        If conn.State = ConnectionState.Open Then
            conn.Close()
        End If
        dbconnection()
        cmdTermDates.Connection = conn
        cmdTermDates.CommandType = CommandType.Text
        cmdTermDates.CommandText = "SELECT * FROM tblSchoolCalendar WHERE (status='True') AND " & _
            vbNewLine & "(termName=@termName) AND (year=@year) AND (dateBeginning=@dateBeginning) AND (dateClosing=@dateClosing)"
        cmdTermDates.Parameters.Clear()
        cmdTermDates.Parameters.AddWithValue("@termName", Me.cboTermName.Text)
        cmdTermDates.Parameters.AddWithValue("@year", Me.cboTermYear.Text)
        cmdTermDates.Parameters.AddWithValue("@dateBeginning", Me.dtpDateBeg.Value.Date)
        cmdTermDates.Parameters.AddWithValue("@dateClosing", Me.dtpDateClosing.Value.Date)
        reader = cmdTermDates.ExecuteReader
        If reader.HasRows Then
            recordExists = True
        Else
            recordExists = False
        End If
        reader.Close()
    End Sub
    Private Sub btnUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdate.Click
        If Me.cboTermName.Text.Trim.Length <= 0 Then
            MsgBox("Term Missing!", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboTermYear.Text.Trim.Length <= 0 Then
            MsgBox("Year Missing!", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If
        Try
            queryType = "UPDATE"
            dbconnection()
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            checkForExistenceOne()
            If recordExists = True Then
                MsgBox("Record Exists!", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "UNSuccessFul Transaction")
                Exit Sub
            End If
            Dim result As MsgBoxResult = MsgBox("Update Record?", MsgBoxStyle.Question + MsgBoxStyle.ApplicationModal + MsgBoxStyle.YesNo, "Confirm Transaction")
            If result = MsgBoxResult.Yes Then
                cmdTermDates.Connection = conn
                cmdTermDates.CommandType = CommandType.StoredProcedure
                cmdTermDates.CommandText = "sprocTermDates"
                cmdTermDates.Parameters.Clear()
                cmdTermDates.Parameters.AddWithValue("@termName", Me.cboTermName.Text.Trim)
                cmdTermDates.Parameters.AddWithValue("@year", Me.cboTermYear.Text.Trim)
                cmdTermDates.Parameters.AddWithValue("@queryType", queryType.Trim)
                cmdTermDates.Parameters.AddWithValue("@userName", userName.Trim)
                cmdTermDates.Parameters.AddWithValue("@regDate", Date.Now)
                cmdTermDates.Parameters.AddWithValue("@termBeginning", Me.dtpDateBeg.Value.Date)
                cmdTermDates.Parameters.AddWithValue("@termEnding", Me.dtpDateClosing.Value.Date)
                cmdTermDates.Parameters.AddWithValue("@termId", Me.cboTermName.Tag)
                rec = cmdTermDates.ExecuteNonQuery
                If rec > 0 Then
                    MsgBox("Record Updated!", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "SuccessFul Transaction")
                End If
            End If
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            Me.lstTermDates.Items.Clear()
            loadTermDates(Me.cboTermYear.Text.Trim)
            Me.cboTermName.Text = ""
            Me.cboTermName.Tag = Nothing
            Me.cboTermYear.Text = ""

            Me.dtpDateClosing.Value = Date.Now
            Me.dtpDateBeg.Value = Date.Now
            queryType = Nothing

            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub

    'Private Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete.Click
    '    If Me.cboTermName.Text.Trim.Length <= 0 Then
    '        MsgBox("Term Missing!", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
    '        Exit Sub
    '    ElseIf Me.cboTermYear.Text.Trim.Length <= 0 Then
    '        MsgBox("Year Missing!", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
    '        Exit Sub
    '    End If
    '    Try
    '        queryType = "DELETE"
    '        dbconnection()
    '        If conn.State = ConnectionState.Closed Then
    '            conn.Open()
    '        End If
    '        If recordExists = True Then
    '            MsgBox("Record Exists!", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "UNSuccessFul Transaction")
    '            Exit Sub
    '        End If
    '        Dim result As MsgBoxResult = MsgBox("Delete Record?", MsgBoxStyle.Question + MsgBoxStyle.ApplicationModal + MsgBoxStyle.YesNo, "Confirm Transaction")
    '        If result = MsgBoxResult.Yes Then
    '            cmdTermDates.Connection = conn
    '            cmdTermDates.CommandType = CommandType.StoredProcedure
    '            cmdTermDates.CommandText = "sprocTermDates"
    '            cmdTermDates.Parameters.Clear()
    '            cmdTermDates.Parameters.AddWithValue("@termName", Me.cboTermName.Text.Trim)
    '            cmdTermDates.Parameters.AddWithValue("@year", Me.cboTermYear.Text.Trim)
    '            cmdTermDates.Parameters.AddWithValue("@queryType", queryType.Trim)
    '            cmdTermDates.Parameters.AddWithValue("@userName", userName.Trim)
    '            cmdTermDates.Parameters.AddWithValue("@regDate", Date.Now)
    '            cmdTermDates.Parameters.AddWithValue("@termBeginning", Me.dtpDateBeg.Value.Date)
    '            cmdTermDates.Parameters.AddWithValue("@termEnding", Me.dtpDateClosing.Value.Date)
    '            cmdTermDates.Parameters.AddWithValue("@termId", Me.cboTermName.Tag)
    '            rec = cmdTermDates.ExecuteNonQuery
    '            If rec > 0 Then
    '                MsgBox("Record Deleted!", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "SuccessFul Transaction")
    '            End If
    '        End If
    '    Catch ex As Exception
    '        MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
    '    Finally
    '        Me.cboTermName.Text = ""
    '        Me.cboTermName.Tag = Nothing
    '        Me.cboTermYear.Text = ""
    '        Me.lstTermDates.Items.Clear()
    '        Me.dtpDateClosing.Value = Date.Now
    '        Me.dtpDateBeg.Value = Date.Now
    '        queryType = Nothing
    '        loadTermDates()
    '        If conn.State = ConnectionState.Open Then
    '            conn.Close()
    '        End If
    '    End Try
    'End Sub
End Class