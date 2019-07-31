Imports System.Data.SqlClient
Public Class frmSchoolPeriods
    Dim recordExists As Boolean = True
    Dim reader As SqlDataReader
    Dim cmdSchPeriods As New SqlCommand
    Dim rec As Integer = 0
    Dim queryType As String = Nothing
    Private Sub frmSchoolPeriods_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            loadLists()
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub
    Private Sub loadLists()
        Me.lstSchoolPeriods.Items.Clear()
        Me.cmdSchPeriods.Connection = conn
        Me.cmdSchPeriods.CommandText = "SELECT * FROM tblPeriods WHERE (status=1) ORDER BY periodName"
        Me.cmdSchPeriods.CommandType = CommandType.Text
        Me.cmdSchPeriods.Parameters.Clear()
        reader = Me.cmdSchPeriods.ExecuteReader
        While reader.Read
            Dim dateBeg As String = CDate(IIf(DBNull.Value.Equals(reader!dateBeginning), "", reader!dateBeginning)).ToString("dd-MM-yyyy")
            Dim dateEnd As String = CDate(IIf(DBNull.Value.Equals(reader!dateEnding), "", reader!dateEnding)).ToString("dd-MM-yyyy")
            li = Me.lstSchoolPeriods.Items.Add(IIf(DBNull.Value.Equals(reader!periodName), "", reader!periodName))
            li.SubItems.Add(dateBeg)
            li.SubItems.Add(dateEnd)
            li.SubItems(1).Tag = IIf(DBNull.Value.Equals(reader!dateBeginning), "", reader!dateBeginning)
            li.SubItems(2).Tag = IIf(DBNull.Value.Equals(reader!dateEnding), "", reader!dateEnding)
        End While
        reader.Close()
    End Sub
    Private Sub frmSchoolPeriods_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint
        If Me.IsMdiChild Then
            Dim PnlLoc As New Point
            PnlLoc.X = CInt((Me.Width - Me.pnlSchoolPeriods.Width) / 2)
            PnlLoc.Y = CInt((Me.Height - Me.pnlSchoolPeriods.Height) / 2.5)
            Me.pnlSchoolPeriods.Location = PnlLoc
        Else
            Me.pnlSchoolPeriods.Dock = DockStyle.Fill
        End If
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub
    Private Sub btnView_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnView.Click
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            loadLists()
            clearTexts()
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try

    End Sub

    Private Sub btnSaveChanges_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSaveChanges.Click
        If Me.txtPeriodName.Text.Trim.Length <= 0 Then
            MsgBox("Period Name Missing", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "System Reply")
            Exit Sub
        ElseIf DateDiff(DateInterval.Day, Me.dtpDateBeginning.Value.Date, Me.dtpDateEnding.Value.Date) < 0 Then
            MsgBox("Ending date cannot be before beginning date", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "System Reply")
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            checkIfRecordExists()
            If recordExists = True Then
                MsgBox("Period Already Exists", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "System Reply")
                Exit Sub
            End If
            Dim result As MsgBoxResult = MsgBox("Update Changes?", MsgBoxStyle.Question + MsgBoxStyle.ApplicationModal + MsgBoxStyle.YesNo, "Confirm Transaction")
            If result = MsgBoxResult.No Then
                Exit Sub
            End If
            queryType = "UPDATE"
            Me.cmdSchPeriods.CommandType = CommandType.StoredProcedure
            Me.cmdSchPeriods.Connection = conn
            Me.cmdSchPeriods.CommandText = "sprocPeriods"
            Me.cmdSchPeriods.Parameters.Clear()
            Me.cmdSchPeriods.Parameters.AddWithValue("@periodName", Me.txtPeriodName.Text.Trim)
            Me.cmdSchPeriods.Parameters.AddWithValue("@dateBeginning", Me.dtpDateBeginning.Value.Date)
            Me.cmdSchPeriods.Parameters.AddWithValue("@dateEnding", Me.dtpDateEnding.Value.Date)
            Me.cmdSchPeriods.Parameters.AddWithValue("@queryType", queryType.Trim)
            Me.cmdSchPeriods.Parameters.AddWithValue("@dateOfReg", Date.Now)
            Me.cmdSchPeriods.Parameters.AddWithValue("@regBy", userName.Trim)
            rec = Me.cmdSchPeriods.ExecuteNonQuery
            If rec > 0 Then
                MsgBox("Period Updated", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "SuccessFul Transaction")
            End If
            clearTexts()
            loadLists()
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub
    Private Sub clearTexts()
        Me.txtPeriodName.Text = ""
        Me.dtpDateBeginning.Value = Date.Now
        Me.dtpDateEnding.Value = Date.Now
    End Sub
    Private Sub checkIfRecordExists()
        recordExists = True
        Me.cmdSchPeriods.Connection = conn
        Me.cmdSchPeriods.CommandText = "SELECT * FROM tblPeriods WHERE (periodName=@periodName) AND (dateBeginning=@dateBeginning) " & _
            vbNewLine & " AND (dateEnding=@dateEnding) AND (status=1)"
        Me.cmdSchPeriods.CommandType = CommandType.Text
        Me.cmdSchPeriods.Parameters.Clear()
        Me.cmdSchPeriods.Parameters.AddWithValue("@periodName", Me.txtPeriodName.Text.Trim)
        Me.cmdSchPeriods.Parameters.AddWithValue("@dateBeginning", Me.dtpDateBeginning.Value.Date)
        Me.cmdSchPeriods.Parameters.AddWithValue("@dateEnding", Me.dtpDateEnding.Value.Date)
        reader = Me.cmdSchPeriods.ExecuteReader
        If reader.HasRows = True Then
            recordExists = True
        ElseIf reader.HasRows = False Then
            recordExists = False
        End If
        reader.Close()
    End Sub

    Private Sub lstSchoolPeriods_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lstSchoolPeriods.Click
        If Me.lstSchoolPeriods.Items.Count <= 0 Then
            MsgBox("No items in the list.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "System Reply")
            Exit Sub
        ElseIf Me.lstSchoolPeriods.SelectedItems.Count <= 0 Then
            MsgBox("Select period to Update.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "System Reply")
            Exit Sub
        ElseIf Me.lstSchoolPeriods.SelectedItems.Count > 1 Then
            MsgBox("Select one period at a time.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "System Reply")
            Exit Sub
        End If
        If Me.lstSchoolPeriods.SelectedItems.Count = 1 Then
            Me.txtPeriodName.Text = Me.lstSchoolPeriods.SelectedItems(0).Text.Trim
            Me.dtpDateBeginning.Value = Me.lstSchoolPeriods.SelectedItems(0).SubItems(1).Tag
            Me.dtpDateEnding.Value = Me.lstSchoolPeriods.SelectedItems(0).SubItems(2).Tag
        End If
        Me.lstSchoolPeriods.Items.Clear()
    End Sub

    Private Sub btnClear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClear.Click
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            clearTexts()
            loadLists()
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub
End Class