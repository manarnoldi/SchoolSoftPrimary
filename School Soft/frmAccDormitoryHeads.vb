Imports System.Data.SqlClient
Public Class frmAccDormitoryHeads
    Dim reader As SqlDataReader
    Dim rec As Integer = 0
    Dim cmdDormHeads As New SqlCommand

    Private Sub frmAccDormitoryHeads_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
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
        Me.cboDormName.Items.Clear()
        Me.cboDormName.Text = ""
        Me.cboDormName.SelectedIndex = -1

        Me.cboEmpType.Items.Clear()
        Me.cboEmpType.Text = ""
        Me.cboEmpType.SelectedIndex = -1

        Me.cboYear.Items.Clear()
        Me.cboYear.Text = ""
        Me.cboYear.SelectedIndex = -1

        Me.cmdDormHeads.Connection = conn
        Me.cmdDormHeads.CommandType = CommandType.Text
        Me.cmdDormHeads.CommandText = "SELECT DISTINCT dormName FROM tblAccdormRegister ORDER BY dormName"
        Me.cmdDormHeads.Parameters.Clear()
        reader = Me.cmdDormHeads.ExecuteReader
        While reader.Read
            Me.cboDormName.Items.Add(IIf(DBNull.Value.Equals(reader!dormName), "", reader!dormName))
        End While
        reader.Close()

        Me.cmdDormHeads.CommandText = "SELECT DISTINCT empType FROM tblSchoolStaff WHERE (status=1) ORDER BY empType"
        Me.cmdDormHeads.Parameters.Clear()
        reader = Me.cmdDormHeads.ExecuteReader
        While reader.Read
            Me.cboEmpType.Items.Add(IIf(DBNull.Value.Equals(reader!empType), "", reader!empType))
        End While
        reader.Close()

        Me.cmdDormHeads.CommandText = "SELECT DISTINCT year FROM tblSchoolCalendar WHERE (status=1) ORDER BY year"
        Me.cmdDormHeads.Parameters.Clear()
        reader = Me.cmdDormHeads.ExecuteReader
        While reader.Read
            Me.cboYear.Items.Add(IIf(DBNull.Value.Equals(reader!year), "", reader!year))
        End While
        reader.Close()
    End Sub
    Private Sub frmAccDormitoryHeads_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint
        If Me.IsMdiChild Then
            Dim PnlLoc As New Point
            PnlLoc.X = CInt((Me.Width - Me.pnlDormHeads.Width) / 2)
            PnlLoc.Y = CInt((Me.Height - Me.pnlDormHeads.Height) / 2.5)
            Me.pnlDormHeads.Location = PnlLoc
        Else
            Me.pnlDormHeads.Dock = DockStyle.Fill
        End If
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub cboDormName_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboDormName.SelectedIndexChanged
        Me.cboStudNo.Items.Clear()
        Me.cboStudNo.Text = ""
        Me.cboStudNo.SelectedIndex = -1

        Me.txtBorder.Text = ""
        Me.txtStudName.Text = ""

        If Me.cboDormName.Text.Trim.Length <= 0 Then
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            Me.cmdDormHeads.Connection = conn
            Me.cmdDormHeads.CommandType = CommandType.Text
            Me.cmdDormHeads.CommandText = "SELECT DISTINCT admNo FROM vwAccDormStudentDetails WHERE (dormName=@dormName) ORDER BY admNo"
            Me.cmdDormHeads.Parameters.Clear()
            Me.cmdDormHeads.Parameters.AddWithValue("@dormName", Me.cboDormName.Text.Trim)
            reader = Me.cmdDormHeads.ExecuteReader
            While reader.Read
                Me.cboStudNo.Items.Add(IIf(DBNull.Value.Equals(reader!admNo), "", reader!admNo))
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

    Private Sub cboEmpType_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboEmpType.SelectedIndexChanged
        Me.cboEmpNo.Items.Clear()
        Me.cboEmpNo.Text = ""
        Me.cboEmpNo.SelectedIndex = -1

        Me.txtEmpName.Text = ""

        If Me.cboEmpType.Text.Trim.Length <= 0 Then
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            Me.cmdDormHeads.Connection = conn
            Me.cmdDormHeads.CommandType = CommandType.Text
            Me.cmdDormHeads.CommandText = "SELECT DISTINCT empNo FROM tblSchoolStaff WHERE (empType=@empType) AND (status=1) ORDER BY empNo"
            Me.cmdDormHeads.Parameters.Clear()
            Me.cmdDormHeads.Parameters.AddWithValue("@empType", Me.cboEmpType.Text.Trim)
            reader = Me.cmdDormHeads.ExecuteReader
            While reader.Read
                Me.cboEmpNo.Items.Add(IIf(DBNull.Value.Equals(reader!empNo), "", reader!empNo))
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

    Private Sub cboEmpNo_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboEmpNo.SelectedIndexChanged
        Me.txtEmpName.Text = ""
        If Me.cboEmpNo.Text.Trim.Length <= 0 Then
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            Me.cmdDormHeads.Connection = conn
            Me.cmdDormHeads.CommandType = CommandType.Text
            Me.cmdDormHeads.CommandText = "SELECT DISTINCT FullName FROM tblSchoolStaff WHERE (empNo=@empNo) AND (status=1)"
            Me.cmdDormHeads.Parameters.Clear()
            Me.cmdDormHeads.Parameters.AddWithValue("@empNo", Me.cboEmpNo.Text.Trim)
            reader = Me.cmdDormHeads.ExecuteReader
            While reader.Read
                Me.txtEmpName.Text = (IIf(DBNull.Value.Equals(reader!FullName), "", reader!FullName))
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

    Private Sub cboStudNo_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboStudNo.SelectedIndexChanged
        Me.txtBorder.Text = ""
        Me.txtStudName.Text = ""
        If Me.cboStudNo.Text.Trim.Length <= 0 Then
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            Me.cmdDormHeads.Connection = conn
            Me.cmdDormHeads.CommandType = CommandType.Text
            Me.cmdDormHeads.CommandText = "SELECT FullName,boarder FROM vwAccDormStudentDetails WHERE (admNo=@admNo)"
            Me.cmdDormHeads.Parameters.Clear()
            Me.cmdDormHeads.Parameters.AddWithValue("@admNo", Me.cboStudNo.Text.Trim)
            reader = Me.cmdDormHeads.ExecuteReader
            While reader.Read
                Me.txtStudName.Text = (IIf(DBNull.Value.Equals(reader!FullName), "", reader!FullName))
                Me.txtBorder.Text = (IIf(DBNull.Value.Equals(reader!boarder), "", reader!boarder))
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
    Private Sub loadList(ByVal year As Integer)
        Me.lstDormDetails.Items.Clear()

        Me.cmdDormHeads.Connection = conn
        Me.cmdDormHeads.CommandType = CommandType.Text
        Me.cmdDormHeads.CommandText = "SELECT * FROM vwAccdormHeadsDetails WHERE (academicYear=@academicYear) ORDER BY dormName"
        Me.cmdDormHeads.Parameters.Clear()
        Me.cmdDormHeads.Parameters.AddWithValue("@academicYear", year)
        reader = Me.cmdDormHeads.ExecuteReader
        While reader.Read
            li = Me.lstDormDetails.Items.Add(IIf(DBNull.Value.Equals(reader!dormName), "", reader!dormName))
            li.SubItems.Add(IIf(DBNull.Value.Equals(reader!empNo), "", reader!empNo))
            li.SubItems.Add(IIf(DBNull.Value.Equals(reader!staffName), "", reader!staffName))
            li.SubItems.Add(IIf(DBNull.Value.Equals(reader!admNo), "", reader!admNo))
            li.SubItems.Add(IIf(DBNull.Value.Equals(reader!studName), "", reader!studName))
            li.Tag = IIf(DBNull.Value.Equals(reader!empType), "", reader!empType)
            li.SubItems(1).Tag = IIf(DBNull.Value.Equals(reader!boarder), "", reader!boarder)
        End While
        reader.Close()
    End Sub
    Private Sub btnView_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnView.Click
        Me.lstDormDetails.Items.Clear()
        If Me.cboYear.Text.Trim.Length <= 0 Then
            MsgBox("Year Missing", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "System Reply")
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            clearTexts()
            loadList(Me.cboYear.Text.Trim)
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        If Me.cboDormName.Text.Trim.Length <= 0 Then
            MsgBox("Dorm Name Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "System Reply")
            Exit Sub
        ElseIf Me.cboYear.Text.Trim.Length <= 0 Then
            MsgBox("Academic Year Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "System Reply")
            Exit Sub
        ElseIf Me.cboEmpNo.Text.Trim.Length <= 0 Then
            MsgBox("Select Employee Number.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "System Reply")
            Exit Sub
        ElseIf Me.cboStudNo.Text.Trim.Length <= 0 Then
            MsgBox("Select Student Number.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "System Reply")
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()

            Dim result As MsgBoxResult = MsgBox("Save Record?", MsgBoxStyle.Question + MsgBoxStyle.ApplicationModal + MsgBoxStyle.YesNo, "Confirm Transaction")
            If result = MsgBoxResult.No Then
                Exit Sub
            End If

            Me.cmdDormHeads.Connection = conn
            Me.cmdDormHeads.CommandType = CommandType.StoredProcedure
            Me.cmdDormHeads.CommandText = "sprocAccDormHeadsRegister"
            Me.cmdDormHeads.Parameters.Clear()
            Me.cmdDormHeads.Parameters.AddWithValue("@dormName", Me.cboDormName.Text.Trim)
            Me.cmdDormHeads.Parameters.AddWithValue("@studNo", Me.cboStudNo.Text.Trim)
            Me.cmdDormHeads.Parameters.AddWithValue("@staffNo", Me.cboEmpNo.Text.Trim)
            Me.cmdDormHeads.Parameters.AddWithValue("@academicYear", Me.cboYear.Text.Trim)
            Me.cmdDormHeads.Parameters.AddWithValue("@regBy", userName.Trim)
            Me.cmdDormHeads.Parameters.AddWithValue("@queryType", 1)
            rec = Me.cmdDormHeads.ExecuteNonQuery

            If rec > 0 Then
                MsgBox("Record Saved Successfully.", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "System Reply")
            End If
            loadList(Me.cboYear.Text.Trim)
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
    Private Sub clearTexts()
        Me.cboDormName.Enabled = True
        Me.btnSave.Enabled = True
        Me.btnUpdate.Enabled = False

        Me.cboStudNo.Items.Clear()
        Me.cboStudNo.Text = ""
        Me.cboStudNo.SelectedIndex = -1

        Me.cboEmpNo.Items.Clear()
        Me.cboEmpNo.Text = ""
        Me.cboEmpNo.SelectedIndex = -1

        Me.txtEmpName.Text = ""
        Me.txtBorder.Text = ""
        Me.txtStudName.Text = ""
    End Sub

    Private Sub cboYear_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboYear.SelectedIndexChanged
        Me.lstDormDetails.Items.Clear()
    End Sub

    Private Sub CLOSEToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CLOSEToolStripMenuItem.Click
        Me.Close()
    End Sub

    Private Sub UPDATEToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UPDATEToolStripMenuItem.Click
        If Me.lstDormDetails.Items.Count <= 0 Then
            MsgBox("No items Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "System Reply")
            Exit Sub
        ElseIf Me.lstDormDetails.CheckedItems.Count <= 0 Then
            MsgBox("No checked items Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "System Reply")
            Exit Sub
        ElseIf Me.lstDormDetails.CheckedItems.Count > 1 Then
            MsgBox("Update one item at a time Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "System Reply")
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            Me.cboDormName.Text = Me.lstDormDetails.CheckedItems(0).Text.Trim
            Me.cboDormName.Enabled = False

            Me.cboStudNo.Text = Me.lstDormDetails.CheckedItems(0).SubItems(3).Text.Trim
            Me.txtBorder.Text = Me.lstDormDetails.CheckedItems(0).SubItems(1).Tag
            Me.txtStudName.Text = Me.lstDormDetails.CheckedItems(0).SubItems(4).Text.Trim
            Me.cboEmpType.Text = Me.lstDormDetails.CheckedItems(0).Tag
            Me.cboEmpNo.Text = Me.lstDormDetails.CheckedItems(0).SubItems(1).Text.Trim
            Me.txtEmpName.Text = Me.lstDormDetails.CheckedItems(0).SubItems(2).Text.Trim

            Me.btnSave.Enabled = False
            Me.btnUpdate.Enabled = True
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub

    Private Sub btnUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdate.Click
        If Me.cboDormName.Text.Trim.Length <= 0 Then
            MsgBox("Dorm Name Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "System Reply")
            Exit Sub
        ElseIf Me.cboYear.Text.Trim.Length <= 0 Then
            MsgBox("Academic Year Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "System Reply")
            Exit Sub
        ElseIf Me.cboEmpNo.Text.Trim.Length <= 0 Then
            MsgBox("Select Employee Number.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "System Reply")
            Exit Sub
        ElseIf Me.cboStudNo.Text.Trim.Length <= 0 Then
            MsgBox("Select Student Number.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "System Reply")
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()

            Dim result As MsgBoxResult = MsgBox("Update Record?", MsgBoxStyle.Question + MsgBoxStyle.ApplicationModal + MsgBoxStyle.YesNo, "Confirm Transaction")
            If result = MsgBoxResult.No Then
                Exit Sub
            End If

            Me.cmdDormHeads.Connection = conn
            Me.cmdDormHeads.CommandType = CommandType.StoredProcedure
            Me.cmdDormHeads.CommandText = "sprocAccDormHeadsRegister"
            Me.cmdDormHeads.Parameters.Clear()
            Me.cmdDormHeads.Parameters.AddWithValue("@dormName", Me.cboDormName.Text.Trim)
            Me.cmdDormHeads.Parameters.AddWithValue("@studNo", Me.cboStudNo.Text.Trim)
            Me.cmdDormHeads.Parameters.AddWithValue("@staffNo", Me.cboEmpNo.Text.Trim)
            Me.cmdDormHeads.Parameters.AddWithValue("@academicYear", Me.cboYear.Text.Trim)
            Me.cmdDormHeads.Parameters.AddWithValue("@regBy", userName.Trim)
            Me.cmdDormHeads.Parameters.AddWithValue("@queryType", 2)
            rec = Me.cmdDormHeads.ExecuteNonQuery

            If rec > 0 Then
                MsgBox("Record Updated Successfully.", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "System Reply")
            End If
            loadList(Me.cboYear.Text.Trim)
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

    Private Sub DELETEToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DELETEToolStripMenuItem.Click
      If Me.lstDormDetails.Items.Count <= 0 Then
            MsgBox("No items Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "System Reply")
            Exit Sub
        ElseIf Me.lstDormDetails.CheckedItems.Count <= 0 Then
            MsgBox("No checked items Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "System Reply")
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()

            Dim result As MsgBoxResult = MsgBox("Delete Record/s?", MsgBoxStyle.Question + MsgBoxStyle.ApplicationModal + MsgBoxStyle.YesNo, "Confirm Transaction")
            If result = MsgBoxResult.No Then
                Exit Sub
            End If

            i = 0
            rec = 0

            For i = 0 To Me.lstDormDetails.CheckedItems.Count - 1
                Me.cmdDormHeads.Connection = conn
                Me.cmdDormHeads.CommandType = CommandType.StoredProcedure
                Me.cmdDormHeads.CommandText = "sprocAccDormHeadsRegister"
                Me.cmdDormHeads.Parameters.Clear()
                Me.cmdDormHeads.Parameters.AddWithValue("@dormName", Me.lstDormDetails.CheckedItems(i).Text.Trim)
                Me.cmdDormHeads.Parameters.AddWithValue("@studNo", Me.lstDormDetails.CheckedItems(i).SubItems(3).Text.Trim)
                Me.cmdDormHeads.Parameters.AddWithValue("@staffNo", Me.lstDormDetails.CheckedItems(i).SubItems(1).Text.Trim)
                Me.cmdDormHeads.Parameters.AddWithValue("@academicYear", Me.cboYear.Text.Trim)
                Me.cmdDormHeads.Parameters.AddWithValue("@regBy", userName.Trim)
                Me.cmdDormHeads.Parameters.AddWithValue("@queryType", 3)
                rec = rec + Me.cmdDormHeads.ExecuteNonQuery
            Next

            If rec > 0 Then
                MsgBox("Record/s Deleted Successfully.", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "System Reply")
            End If
            loadList(Me.cboYear.Text.Trim)
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
End Class