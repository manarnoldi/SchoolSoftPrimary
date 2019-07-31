Imports System.Data.SqlClient
Public Class frmSchTeachingRooms
    Dim queryType As String = Nothing
    Dim recordExists As Boolean = False
    Dim reader As SqlDataReader
    Dim rec As Integer = 0
    Dim cmdRooms As New SqlCommand
    Private Sub frmSchTeachingRooms_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            loadRooms()
            loadRoomType()
            Me.rbFalse.Checked = False
            Me.rbTrue.Checked = False
        Catch ex As Exception
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub

    Private Sub frmSchTeachingRooms_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint
        If Me.IsMdiChild Then
            Dim PnlLoc As New Point
            PnlLoc.X = CInt((Me.Width - Me.pnlTeachingRooms.Width) / 2)
            PnlLoc.Y = CInt((Me.Height - Me.pnlTeachingRooms.Height) / 2.5)
            Me.pnlTeachingRooms.Location = PnlLoc
        Else
            Me.pnlTeachingRooms.Dock = DockStyle.Fill
        End If
    End Sub
    Private Sub loadRoomType()
        cboRoomType.Items.Clear()
        Me.cboRoomType.Text = ""
        cmdRooms.Connection = conn
        cmdRooms.CommandType = CommandType.Text
        cmdRooms.CommandText = "SELECT DISTINCT roomType FROM tblSchoolRooms WHERE (status='True') ORDER BY roomType"
        cmdRooms.Parameters.Clear()
        reader = cmdRooms.ExecuteReader
        If reader.HasRows Then
            While reader.Read
                Me.cboRoomType.Items.Add(IIf(DBNull.Value.Equals(reader!roomType), "", reader!roomType))
            End While
        End If
        reader.Close()
    End Sub
    Private Sub loadRooms()
        Me.lstTeachingRooms.Items.Clear()
        cmdRooms.Connection = conn
        cmdRooms.CommandType = CommandType.Text
        cmdRooms.CommandText = "SELECT * FROM tblSchoolRooms WHERE (status='True') ORDER BY roomType,roomName"
        cmdRooms.Parameters.Clear()
        reader = cmdRooms.ExecuteReader
        If reader.HasRows Then
            While reader.Read
                li = Me.lstTeachingRooms.Items.Add(IIf(DBNull.Value.Equals(reader!roomType), "", reader!roomType))
                li.SubItems.Add(IIf(DBNull.Value.Equals(reader!roomName), "", reader!roomName))
                li.SubItems.Add(IIf(DBNull.Value.Equals(reader!roomCapacity), "", reader!roomCapacity))
                li.SubItems.Add(IIf(DBNull.Value.Equals(reader!sharable), "", reader!sharable))
                li.Tag = IIf(DBNull.Value.Equals(reader!roomId), "", reader!roomId)
            End While
        End If
        reader.Close()
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub txtRoomCapacity_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtRoomCapacity.LostFocus
        If Me.txtRoomCapacity.Text = "" Then
            Exit Sub
        Else
            Me.txtRoomCapacity.Text = Math.Floor(CDbl(Me.txtRoomCapacity.Text))
        End If
    End Sub

    Private Sub txtRoomCapacity_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtRoomCapacity.TextChanged
        If Me.txtRoomCapacity.Text.Trim.Length <= 0 Then
            Exit Sub
        ElseIf IsNumeric(Me.txtRoomCapacity.Text) = False Then
            MsgBox("Non Numeric Values Detected", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Me.txtRoomCapacity.Text = ""
            Exit Sub
        End If
    End Sub

    Private Sub btnView_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnView.Click
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            loadRooms()
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            loadRoomType()
            Me.cboRoomType.SelectedIndex = -1
            Me.txtRoomCapacity.Text = ""
            Me.txtRoomName.Text = ""
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
            Me.btnDelete.Enabled = False
            Me.btnSave.Enabled = True
            Me.btnUpdate.Enabled = False
            Me.rbFalse.Checked = False
            Me.rbTrue.Checked = False
        End Try
    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        If Me.txtRoomName.Text.Trim.Length <= 0 Then
            MsgBox("Missing Room Name", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboRoomType.Text.Trim.Length <= 0 Then
            MsgBox("Missing Room type", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.rbFalse.Checked = False And Me.rbTrue.Checked = False Then
            MsgBox("Select If Room Is Sharable", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.txtRoomCapacity.Text = "" Or CInt(Me.txtRoomCapacity.Text) = 0 Then
            MsgBox("Enter Room Capacity", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If
        Dim sharable As Boolean = False
        If Me.rbTrue.Checked = True Then
            sharable = True
        ElseIf Me.rbFalse.Checked = True Then
            sharable = False
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            recordExists = False
            checkForExistence()
            If recordExists = True Then
                MsgBox("Record Exists!", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "UNSuccessFul Transaction")
                Exit Sub
            End If
            queryType = "INSERT"
            Dim result As MsgBoxResult = MsgBox("Save Record?", MsgBoxStyle.Question + MsgBoxStyle.ApplicationModal + MsgBoxStyle.YesNo, "Confirm Transaction")
            If result = MsgBoxResult.Yes Then
                cmdRooms.Connection = conn
                cmdRooms.CommandType = CommandType.StoredProcedure
                cmdRooms.CommandText = "sprocSchoolRooms"
                cmdRooms.Parameters.Clear()
                cmdRooms.Parameters.AddWithValue("@regBy", userName.Trim)
                cmdRooms.Parameters.AddWithValue("@roomName", Me.txtRoomName.Text.Trim)
                cmdRooms.Parameters.AddWithValue("@roomCapacity", Me.txtRoomCapacity.Text.Trim)
                cmdRooms.Parameters.AddWithValue("@sharable", sharable)
                cmdRooms.Parameters.AddWithValue("@roomType", Me.cboRoomType.Text)
                cmdRooms.Parameters.AddWithValue("@queryType", queryType.Trim)
                cmdRooms.Parameters.AddWithValue("@dateOfReg", Date.Now)
                rec = cmdRooms.ExecuteNonQuery
                If rec > 0 Then
                    MsgBox("Record Saved!", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "SuccessFul Transaction")
                End If
            End If
            loadRoomType()
            loadRooms()
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
            Me.txtRoomName.Tag = Nothing
            Me.cboRoomType.SelectedIndex = -1
            Me.txtRoomCapacity.Text = ""
            Me.txtRoomName.Text = ""
            Me.btnDelete.Enabled = False
            Me.btnSave.Enabled = True
            Me.btnUpdate.Enabled = False
            Me.rbFalse.Checked = False
            Me.rbTrue.Checked = False
        End Try
    End Sub
    Private Sub checkForExistence()
        If conn.State = ConnectionState.Open Then
            conn.Close()
        End If
        dbconnection()
        cmdRooms.Connection = conn
        cmdRooms.CommandType = CommandType.Text
        cmdRooms.CommandText = "SELECT * FROM tblSchoolRooms WHERE (status='True') AND (roomName=@roomName)"
        cmdRooms.Parameters.Clear()
        cmdRooms.Parameters.AddWithValue("@roomName", Me.txtRoomName.Text.Trim)
        reader = cmdRooms.ExecuteReader
        If reader.HasRows Then
            recordExists = True
        Else
            recordExists = False
        End If
        reader.Close()
    End Sub
    Private Sub checkForExistenceOne()
        Dim sharable As Boolean = False
        If Me.rbTrue.Checked = True Then
            sharable = True
        ElseIf Me.rbFalse.Checked = True Then
            sharable = False
        End If
        If conn.State = ConnectionState.Open Then
            conn.Close()
        End If
        dbconnection()
        cmdRooms.Connection = conn
        cmdRooms.CommandType = CommandType.Text
        cmdRooms.CommandText = "SELECT * FROM tblSchoolRooms WHERE (roomName=@roomName) AND (roomCapacity=@roomCapacity) AND " & _
             vbNewLine & "(sharable=@sharable) AND (roomType=@roomType) AND (status='True')"
        cmdRooms.Parameters.Clear()
        cmdRooms.Parameters.AddWithValue("@roomName", Me.txtRoomName.Text.Trim)
        cmdRooms.Parameters.AddWithValue("@roomCapacity", Me.txtRoomCapacity.Text.Trim)
        cmdRooms.Parameters.AddWithValue("@sharable", sharable)
        cmdRooms.Parameters.AddWithValue("@roomType", Me.cboRoomType.Text.Trim)
        reader = cmdRooms.ExecuteReader
        If reader.HasRows Then
            recordExists = True
        Else
            recordExists = False
        End If
        reader.Close()
    End Sub

    Private Sub lstTeachingRooms_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lstTeachingRooms.Click
        If Me.lstTeachingRooms.SelectedItems.Count = 1 Then
            Me.txtRoomCapacity.Text = Me.lstTeachingRooms.SelectedItems(0).SubItems(2).Text
            Me.txtRoomName.Text = Me.lstTeachingRooms.SelectedItems(0).SubItems(1).Text
            Me.cboRoomType.Text = Me.lstTeachingRooms.SelectedItems(0).Text
            Me.txtRoomName.Tag = Me.lstTeachingRooms.SelectedItems(0).Tag
            If Me.lstTeachingRooms.SelectedItems(0).SubItems(3).Text.Trim = "True" Then
                rbTrue.Checked = True
            ElseIf Me.lstTeachingRooms.SelectedItems(0).SubItems(3).Text.Trim = "False" Then
                rbFalse.Checked = True
            End If
            Me.btnDelete.Enabled = True
            Me.btnSave.Enabled = False
            Me.btnUpdate.Enabled = True
        End If
    End Sub

    Private Sub btnUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdate.Click
        If Me.txtRoomName.Text.Trim.Length <= 0 Then
            MsgBox("Missing Room Name", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboRoomType.Text.Trim.Length <= 0 Then
            MsgBox("Missing Room type", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.rbFalse.Checked = False And Me.rbTrue.Checked = False Then
            MsgBox("Select If Room Is Sharable", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.txtRoomCapacity.Text = "" Or CInt(Me.txtRoomCapacity.Text) = 0 Then
            MsgBox("Enter Room Capacity", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            queryType = "UPDATE"
            recordExists = False
            checkForExistenceOne()
            If recordExists = True Then
                MsgBox("Record Exists!", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "UNSuccessFul Transaction")
                Exit Sub
            End If
            Dim result As MsgBoxResult = MsgBox("Update Record?", MsgBoxStyle.Question + MsgBoxStyle.ApplicationModal + MsgBoxStyle.YesNo, "Confirm Transaction")
            If result = MsgBoxResult.Yes Then
                Dim sharable As Boolean = False
                If Me.rbTrue.Checked = True Then
                    sharable = True
                ElseIf Me.rbFalse.Checked = True Then
                    sharable = False
                End If
                cmdRooms.Connection = conn
                cmdRooms.CommandType = CommandType.StoredProcedure
                cmdRooms.CommandText = "sprocSchoolRooms"
                cmdRooms.Parameters.Clear()
                cmdRooms.Parameters.AddWithValue("@roomId", Me.txtRoomName.Tag)
                cmdRooms.Parameters.AddWithValue("@regBy", userName.Trim)
                cmdRooms.Parameters.AddWithValue("@roomName", Me.txtRoomName.Text.Trim)
                cmdRooms.Parameters.AddWithValue("@roomCapacity", Me.txtRoomCapacity.Text.Trim)
                cmdRooms.Parameters.AddWithValue("@sharable", sharable)
                cmdRooms.Parameters.AddWithValue("@roomType", Me.cboRoomType.Text)
                cmdRooms.Parameters.AddWithValue("@queryType", queryType.Trim)
                cmdRooms.Parameters.AddWithValue("@dateOfReg", Date.Now)
                rec = cmdRooms.ExecuteNonQuery()
                If rec > 0 Then
                    MsgBox("Record Updated!", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "SuccessFul Transaction")
                End If
                loadRooms()
                loadRoomType()
            End If
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            Me.txtRoomName.Tag = Nothing
            Me.cboRoomType.SelectedIndex = -1
            Me.txtRoomCapacity.Text = ""
            Me.txtRoomName.Text = ""
            Me.btnDelete.Enabled = False
            Me.btnSave.Enabled = True
            Me.btnUpdate.Enabled = False
            Me.rbFalse.Checked = False
            Me.rbTrue.Checked = False
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub

    Private Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        If Me.txtRoomName.Text.Trim.Length <= 0 Then
            MsgBox("Missing Room Name", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.txtRoomName.Tag = Nothing Then
            MsgBox("Reselect Room to delete", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            queryType = "DELETE"
            Dim result As MsgBoxResult = MsgBox("Delete Record?", MsgBoxStyle.Question + MsgBoxStyle.ApplicationModal + MsgBoxStyle.YesNo, "Confirm Transaction")
            If result = MsgBoxResult.Yes Then
                cmdRooms.Connection = conn
                cmdRooms.CommandType = CommandType.StoredProcedure
                cmdRooms.CommandText = "sprocSchoolRooms"
                cmdRooms.Parameters.Clear()
                cmdRooms.Parameters.AddWithValue("@roomId", Me.txtRoomName.Tag)
                cmdRooms.Parameters.AddWithValue("@regBy", userName.Trim)
                cmdRooms.Parameters.AddWithValue("@roomName", Me.txtRoomName.Text.Trim)
                cmdRooms.Parameters.AddWithValue("@queryType", queryType.Trim)
                cmdRooms.Parameters.AddWithValue("@dateOfReg", Date.Now)
                rec = cmdRooms.ExecuteNonQuery()
                If rec > 0 Then
                    MsgBox("Record Deleted!", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "SuccessFul Transaction")
                End If
                loadRooms()
                loadRoomType()
            End If
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            Me.txtRoomName.Tag = Nothing
            Me.cboRoomType.SelectedIndex = -1
            Me.txtRoomCapacity.Text = ""
            Me.txtRoomName.Text = ""
            Me.btnDelete.Enabled = False
            Me.btnSave.Enabled = True
            Me.btnUpdate.Enabled = False
            Me.rbFalse.Checked = False
            Me.rbTrue.Checked = False
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub

    Private Sub lstTeachingRooms_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lstTeachingRooms.SelectedIndexChanged

    End Sub
End Class
