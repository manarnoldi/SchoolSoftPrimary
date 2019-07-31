Imports System.Data.SqlClient
Public Class frmSchDepartments
    Dim sqlQuery As String = Nothing
    Dim recordExists As Boolean = False
    Dim cmdDepartment As New SqlCommand
    Dim reader As SqlDataReader
    Dim rec As Integer = 0
    Private Sub frmSchDepartments_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            LoadStaff()
            loadDepartments()
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub

    Private Sub frmSchDepartments_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint
        If Me.IsMdiChild Then
            Dim PnlLoc As New Point
            PnlLoc.X = CInt((Me.Width - Me.pnlSchDepartment.Width) / 2)
            PnlLoc.Y = CInt((Me.Height - Me.pnlSchDepartment.Height) / 2.5)
            Me.pnlSchDepartment.Location = PnlLoc
        Else
            Me.pnlSchDepartment.Dock = DockStyle.Fill
        End If
    End Sub
    Private Sub LoadStaff()
        Me.cboStaffType.Items.Clear()
        Me.cboStaffType.Text = ""
        cmdDepartment.Connection = conn
        cmdDepartment.CommandType = CommandType.Text
        cmdDepartment.CommandText = "SELECT DISTINCT empType FROM tblSchoolStaff WHERE (status='True') ORDER BY empType"
        reader = cmdDepartment.ExecuteReader()
        If reader.HasRows Then
            While reader.Read
                Me.cboStaffType.Items.Add(IIf(DBNull.Value.Equals(reader!empType), "", reader!empType))
            End While
        End If
        reader.Close()
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub cboStaffType_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboStaffType.SelectedIndexChanged
        If Me.cboStaffType.Text.Trim.Length <= 0 Then
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            Me.cboStaffNo.Items.Clear()
            Me.cboStaffNo.Text = ""
            Me.txtStaffName.Text = ""
            cmdDepartment.Connection = conn
            cmdDepartment.CommandType = CommandType.Text
            cmdDepartment.CommandText = "SELECT DISTINCT empNo FROM tblSchoolStaff WHERE (empType=@empType) AND (status='True') ORDER BY empNo"
            Me.cmdDepartment.Parameters.Clear()
            Me.cmdDepartment.Parameters.AddWithValue("@empType", Me.cboStaffType.Text.Trim)
            reader = cmdDepartment.ExecuteReader()
            If reader.HasRows Then
                While reader.Read
                    Me.cboStaffNo.Items.Add(IIf(DBNull.Value.Equals(reader!empNo), "", reader!empNo))
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

    Private Sub cboStaffNo_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboStaffNo.SelectedIndexChanged
        If Me.cboStaffNo.Text.Trim.Length <= 0 Then
            Exit Sub
        ElseIf Me.cboStaffType.Text.Trim.Length <= 0 Then
            MsgBox("Missing Staff Type", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If
        If Me.cboStaffType.Text.Trim.Length <= 0 Then
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            Me.txtStaffName.Text = ""
            cmdDepartment.Connection = conn
            cmdDepartment.CommandType = CommandType.Text
            cmdDepartment.CommandText = "SELECT FullName FROM tblSchoolStaff WHERE (empNo=@empNo) AND (status='True')"
            cmdDepartment.Parameters.Clear()
            cmdDepartment.Parameters.AddWithValue("@empNo", Me.cboStaffNo.Text.Trim)
            reader = cmdDepartment.ExecuteReader()
            If reader.HasRows Then
                While reader.Read
                    Me.txtStaffName.Text = (IIf(DBNull.Value.Equals(reader!FullName), "", reader!FullName))
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
    Private Sub loadDepartments()
        Me.lstDepartMents.Items.Clear()
        cmdDepartment.Connection = conn
        cmdDepartment.CommandType = CommandType.Text
        cmdDepartment.CommandText = "SELECT * FROM  vwSchoolDepts WHERE (status='True') ORDER BY deptCode"
        cmdDepartment.Parameters.Clear()
        reader = cmdDepartment.ExecuteReader()
        If reader.HasRows Then
            While reader.Read
                li = Me.lstDepartMents.Items.Add(IIf(DBNull.Value.Equals(reader!deptCode), "", reader!deptCode))
                li.SubItems.Add(IIf(DBNull.Value.Equals(reader!deptName), "", reader!deptName))
                li.SubItems.Add(IIf(DBNull.Value.Equals(reader!FullName), "", reader!FullName))
                li.Tag = IIf(DBNull.Value.Equals(reader!deptId), "", reader!deptId)
            End While
        End If
        reader.Close()
    End Sub

    Private Sub btnView_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnView.Click
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            loadDepartments()
            LoadStaff()
            Me.cboStaffNo.Items.Clear()
            Me.cboStaffNo.Text = ""
            Me.txtStaffName.Text = ""
            Me.txtDepCode.Text = ""
            Me.txtDepName.Text = ""
            Me.btnSave.Enabled = True
            Me.btnUpdate.Enabled = False
            Me.btnDelete.Enabled = False
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        If Me.txtDepCode.Text.Trim.Length <= 0 Then
            MsgBox("Missing Department Code", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.txtDepName.Text.Trim.Length <= 0 Then
            MsgBox("Missing Department Name", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If
        Try
            sqlQuery = "INSERT"
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
            Dim result As MsgBoxResult = MsgBox("Save Record?", MsgBoxStyle.Question + MsgBoxStyle.ApplicationModal + MsgBoxStyle.YesNo, "Confirm Transaction")
            If result = MsgBoxResult.Yes Then
                cmdDepartment.Connection = conn
                cmdDepartment.CommandType = CommandType.StoredProcedure
                cmdDepartment.CommandText = "sprocSchDepartments"
                cmdDepartment.Parameters.Clear()
                cmdDepartment.Parameters.AddWithValue("@deptCode", Me.txtDepCode.Text.Trim)
                cmdDepartment.Parameters.AddWithValue("@deptName", Me.txtDepName.Text.Trim)
                cmdDepartment.Parameters.AddWithValue("@deptHead", Me.cboStaffNo.Text.Trim)
                cmdDepartment.Parameters.AddWithValue("@userName", userName.Trim)
                cmdDepartment.Parameters.AddWithValue("@regDate", Date.Now)
                cmdDepartment.Parameters.AddWithValue("@queryType", sqlQuery.Trim)
                rec = cmdDepartment.ExecuteNonQuery
                If rec > 0 Then
                    MsgBox("Record Saved!", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "SuccessFul Transaction")
                End If
            End If
            loadDepartments()
            LoadStaff()
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
            Me.cboStaffType.Text = ""
            Me.cboStaffNo.SelectedIndex = -1
            Me.txtDepCode.Text = ""
            Me.txtDepCode.Tag = Nothing
            Me.txtDepName.Text = ""
            Me.txtStaffName.Text = ""
            sqlQuery = Nothing
        End Try
    End Sub

    Private Sub btnUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdate.Click
        If Me.txtDepCode.Text.Trim.Length <= 0 Then
            MsgBox("Missing Department Code", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.txtDepName.Text.Trim.Length <= 0 Then
            MsgBox("Missing Department Name", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If

        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            sqlQuery = "UPDATE"
            recordExists = False
            checkForExistenceOne()
            If recordExists = True Then
                MsgBox("Record Exists!", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "UNSuccessFul Transaction")
                Exit Sub
            End If
            Dim result As MsgBoxResult = MsgBox("Update Record?", MsgBoxStyle.Question + MsgBoxStyle.ApplicationModal + MsgBoxStyle.YesNo, "Confirm Transaction")
            If result = MsgBoxResult.Yes Then
                cmdDepartment.Connection = conn
                cmdDepartment.CommandType = CommandType.StoredProcedure
                cmdDepartment.CommandText = "sprocSchDepartments"
                cmdDepartment.Parameters.Clear()
                cmdDepartment.Parameters.AddWithValue("@deptId", Me.txtDepCode.Tag)
                cmdDepartment.Parameters.AddWithValue("@deptCode", Me.txtDepCode.Text.Trim)
                cmdDepartment.Parameters.AddWithValue("@deptName", Me.txtDepName.Text.Trim)
                cmdDepartment.Parameters.AddWithValue("@deptHead", Me.cboStaffNo.Text.Trim)
                cmdDepartment.Parameters.AddWithValue("@userName", userName.Trim)
                cmdDepartment.Parameters.AddWithValue("@regDate", Date.Now)
                cmdDepartment.Parameters.AddWithValue("@queryType", sqlQuery.Trim)
                rec = cmdDepartment.ExecuteNonQuery
                If rec > 0 Then
                    MsgBox("Record Updated!", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "SuccessFul Transaction")
                End If
            End If
            loadDepartments()
            LoadStaff()
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
            Me.cboStaffType.Text = ""
            Me.cboStaffNo.SelectedIndex = -1
            Me.cboStaffNo.Text = ""
            Me.txtDepCode.Text = ""
            Me.txtDepCode.Tag = Nothing
            Me.txtDepName.Text = ""
            Me.txtStaffName.Text = ""
            sqlQuery = Nothing
            Me.btnSave.Enabled = True
            Me.btnUpdate.Enabled = False
            Me.btnDelete.Enabled = False
        End Try
    End Sub

    Private Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        If Me.txtDepCode.Text.Trim.Length <= 0 Then
            MsgBox("Missing Department Code", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.txtDepName.Text.Trim.Length <= 0 Then
            MsgBox("Missing Department Name", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If

        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            sqlQuery = "DELETE"
            Dim result As MsgBoxResult = MsgBox("Delete Record?", MsgBoxStyle.Question + MsgBoxStyle.ApplicationModal + MsgBoxStyle.YesNo, "Confirm Transaction")
            If result = MsgBoxResult.Yes Then
                cmdDepartment.Connection = conn
                cmdDepartment.CommandType = CommandType.StoredProcedure
                cmdDepartment.CommandText = "sprocSchDepartments"
                cmdDepartment.Parameters.Clear()
                cmdDepartment.Parameters.AddWithValue("@deptId", Me.txtDepCode.Tag)
                cmdDepartment.Parameters.AddWithValue("@deptCode", Me.txtDepCode.Text.Trim)
                cmdDepartment.Parameters.AddWithValue("@deptName", Me.txtDepName.Text.Trim)
                cmdDepartment.Parameters.AddWithValue("@deptHead", Me.cboStaffNo.Text.Trim)
                cmdDepartment.Parameters.AddWithValue("@userName", userName.Trim)
                cmdDepartment.Parameters.AddWithValue("@regDate", Date.Now)
                cmdDepartment.Parameters.AddWithValue("@queryType", sqlQuery.Trim)
                rec = cmdDepartment.ExecuteNonQuery
                If rec > 0 Then
                    MsgBox("Record Deleted!", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "SuccessFul Transaction")
                End If
            End If
            loadDepartments()
            LoadStaff()
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
            Me.cboStaffType.Text = ""
            Me.cboStaffNo.Text = ""
            Me.cboStaffNo.SelectedIndex = -1
            Me.txtDepCode.Text = ""
            Me.txtDepCode.Tag = Nothing
            Me.txtDepName.Text = ""
            Me.txtStaffName.Text = ""
            sqlQuery = Nothing
            Me.btnSave.Enabled = True
            Me.btnUpdate.Enabled = False
            Me.btnDelete.Enabled = False
        End Try
    End Sub

    Private Sub lstDepartMents_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lstDepartMents.Click
        If Me.lstDepartMents.SelectedItems.Count = 1 Then
            Me.txtDepCode.Text = Me.lstDepartMents.SelectedItems(0).Text
            Me.txtDepName.Text = Me.lstDepartMents.SelectedItems(0).SubItems(1).Text
            Me.txtStaffName.Text = Me.lstDepartMents.SelectedItems(0).SubItems(2).Text
            Me.txtDepCode.Tag = Me.lstDepartMents.SelectedItems(0).Tag
            Me.btnSave.Enabled = False
            Me.btnUpdate.Enabled = True
            Me.btnDelete.Enabled = True
        End If
    End Sub
    Private Sub checkForExistence()
        If conn.State = ConnectionState.Open Then
            conn.Close()
        End If
        dbconnection()
        cmdDepartment.Connection = conn
        cmdDepartment.CommandType = CommandType.Text
        cmdDepartment.CommandText = "SELECT * FROM tblSchoolDepts WHERE (status='True') AND (deptName=@deptName)"
        cmdDepartment.Parameters.Clear()
        cmdDepartment.Parameters.AddWithValue("@deptName", Me.txtDepName.Text.Trim)
        reader = cmdDepartment.ExecuteReader
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
        cmdDepartment.Connection = conn
        cmdDepartment.CommandType = CommandType.Text
        cmdDepartment.CommandText = "SELECT * FROM vwSchoolDepts WHERE (status='True') AND (staffState='True') " & _
            vbNewLine & "AND (deptName=@deptName) AND (deptCode=@deptCode) AND (fullName=@fullName)"
        cmdDepartment.Parameters.Clear()
        cmdDepartment.Parameters.AddWithValue("@deptName", Me.txtDepName.Text.Trim)
        cmdDepartment.Parameters.AddWithValue("@deptCode", Me.txtDepCode.Text.Trim)
        cmdDepartment.Parameters.AddWithValue("@fullName", Me.txtStaffName.Text.Trim)
        reader = cmdDepartment.ExecuteReader
        If reader.HasRows Then
            recordExists = True
        Else
            recordExists = False
        End If
        reader.Close()
    End Sub

    Private Sub lstDepartMents_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lstDepartMents.SelectedIndexChanged

    End Sub
End Class