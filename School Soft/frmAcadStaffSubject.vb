Imports System.Data.SqlClient
Public Class frmAcadStaffSubject
    Dim subjectIsAllocated As Boolean = True
    Dim numberOfSubjects As Integer = 0
    Dim maxNoOfSubjects As Integer = 0
    Dim queryType As String = Nothing
    Dim reader As SqlDataReader
    Dim cmdStaffSubject As New SqlCommand
    Dim rec As Integer = 0
    Dim recordExists As Boolean = True
    Private Sub frmAcadStaffSubject_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            loadCombos()
            loadLists()
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub

    Private Sub frmAcadStaffSubject_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint
        If Me.IsMdiChild Then
            Dim PnlLoc As New Point
            PnlLoc.X = CInt((Me.Width - Me.pnlAcadStaffSubj.Width) / 2)
            PnlLoc.Y = CInt((Me.Height - Me.pnlAcadStaffSubj.Height) / 2.5)
            Me.pnlAcadStaffSubj.Location = PnlLoc
        Else
            Me.pnlAcadStaffSubj.Dock = DockStyle.Fill
        End If
    End Sub
    Private Sub loadCombos()
        Me.cboYear.Items.Clear()
        Me.cboYear.Text = ""
        Me.cboYear.SelectedIndex = -1

        Me.cboEmpNo.Items.Clear()
        Me.cboEmpNo.Text = ""
        Me.cboEmpNo.SelectedIndex = -1

        Me.cboYear.Items.Clear()
        Me.cboYear.Text = ""
        Me.cboYear.SelectedIndex = -1

        Me.cboYearView.Items.Clear()
        Me.cboYearView.Text = ""
        Me.cboYearView.SelectedIndex = -1

        Me.cboContract.Items.Clear()
        Me.cboContract.Text = ""
        Me.cboContract.SelectedIndex = -1

        Me.cboContractView.Items.Clear()
        Me.cboContractView.Text = ""
        Me.cboContractView.SelectedIndex = -1
        Me.txtEmpName.Text = ""
        Me.cmdStaffSubject.Connection = conn
        Me.cmdStaffSubject.CommandType = CommandType.Text
        Me.cmdStaffSubject.CommandText = "SELECT DISTINCT year FROM tblClasses WHERE (status=1) ORDER BY year"
        Me.cmdStaffSubject.Parameters.Clear()
        reader = Me.cmdStaffSubject.ExecuteReader
        If reader.HasRows Then
            While reader.Read
                Me.cboYear.Items.Add(IIf(DBNull.Value.Equals(reader!year), "", (reader!year)))
                Me.cboYearView.Items.Add(IIf(DBNull.Value.Equals(reader!year), "", (reader!year)))
            End While
        End If
        reader.Close()

        Me.cmdStaffSubject.CommandText = "SELECT DISTINCT contractType FROM tblSchoolStaff WHERE (status=1) AND (empType='Teaching') ORDER BY contractType"
        Me.cmdStaffSubject.Parameters.Clear()
        reader = Me.cmdStaffSubject.ExecuteReader
        If reader.HasRows Then
            While reader.Read
                Me.cboContract.Items.Add(IIf(DBNull.Value.Equals(reader!contractType), "", (reader!contractType)))
                Me.cboContractView.Items.Add(IIf(DBNull.Value.Equals(reader!contractType), "", (reader!contractType)))
            End While
        End If
        reader.Close()
    End Sub
    Private Sub loadLists()
        Me.lstSubject.Items.Clear()
        Me.lstClasses.Items.Clear()

        Me.cmdStaffSubject.Connection = conn
        Me.cmdStaffSubject.CommandType = CommandType.Text
        Me.cmdStaffSubject.CommandText = "SELECT * FROM tblSubjects WHERE (subStatus=1) ORDER BY subCode"
        Me.cmdStaffSubject.Parameters.Clear()
        reader = Me.cmdStaffSubject.ExecuteReader
        If reader.HasRows Then
            While reader.Read
                li = Me.lstSubject.Items.Add(IIf(DBNull.Value.Equals(reader!subCode), "", (reader!subCode)))
                li.SubItems.Add(IIf(DBNull.Value.Equals(reader!subName), "", (reader!subName)))
            End While
        End If
        reader.Close()


    End Sub
    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub CLOSEToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Me.Close()
    End Sub

    Private Sub cboContract_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboContract.SelectedIndexChanged
        If Me.cboContract.Text.Trim.Length <= 0 Then
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()

            Me.cboEmpNo.Items.Clear()
            Me.cboEmpNo.Text = ""
            Me.cboEmpNo.SelectedIndex = -1

            Me.txtEmpName.Text = ""

            Me.cmdStaffSubject.Connection = conn
            Me.cmdStaffSubject.CommandType = CommandType.Text
            Me.cmdStaffSubject.CommandText = "SELECT DISTINCT empNo FROM tblSchoolStaff WHERE (status=1) AND (empType='Teaching') AND " & _
                vbNewLine & "(contractType=@contractType) ORDER BY empNo"
            Me.cmdStaffSubject.Parameters.Clear()
            Me.cmdStaffSubject.Parameters.AddWithValue("@contractType", Me.cboContract.Text.Trim)
            reader = Me.cmdStaffSubject.ExecuteReader
            If reader.HasRows Then
                While reader.Read
                    Me.cboEmpNo.Items.Add(IIf(DBNull.Value.Equals(reader!empNo), "", (reader!empNo)))
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

    Private Sub cboEmpNo_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboEmpNo.SelectedIndexChanged
        If Me.cboEmpNo.Text.Trim.Length <= 0 Then
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()

            Me.txtEmpName.Text = ""

            Me.cmdStaffSubject.Connection = conn
            Me.cmdStaffSubject.CommandType = CommandType.Text
            Me.cmdStaffSubject.CommandText = "SELECT FullName FROM tblSchoolStaff WHERE (status=1) AND (empNo=@empNo)"
            Me.cmdStaffSubject.Parameters.Clear()
            Me.cmdStaffSubject.Parameters.AddWithValue("@empNo", Me.cboEmpNo.Text.Trim)
            reader = Me.cmdStaffSubject.ExecuteReader
            If reader.HasRows Then
                While reader.Read
                    Me.txtEmpName.Text = (IIf(DBNull.Value.Equals(reader!FullName), "", (reader!FullName)))
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

    Private Sub cboContractView_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()

            Me.cboEmpNoView.Items.Clear()
            Me.cboEmpNoView.Text = ""
            Me.cboEmpNoView.SelectedIndex = -1

            Me.lstStaffSubjectView.Items.Clear()

            Me.cmdStaffSubject.Connection = conn
            Me.cmdStaffSubject.CommandType = CommandType.Text
            Me.cmdStaffSubject.CommandText = "SELECT DISTINCT empNo FROM tblSchoolStaff WHERE (status=1) AND (empType='Teaching') AND " & _
                vbNewLine & "(contractType=@contractType) ORDER BY empNo"
            Me.cmdStaffSubject.Parameters.Clear()
            Me.cmdStaffSubject.Parameters.AddWithValue("@contractType", Me.cboContractView.Text.Trim)
            reader = Me.cmdStaffSubject.ExecuteReader
            If reader.HasRows Then
                While reader.Read
                    Me.cboEmpNoView.Items.Add(IIf(DBNull.Value.Equals(reader!empNo), "", (reader!empNo)))
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

    Private Sub cboClassView_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboYearView.SelectedIndexChanged
        txtTotalSubjects.Text = ""
        Me.lstStaffSubjectView.Items.Clear()
    End Sub

    Private Sub cboYear_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboYear.SelectedIndexChanged
        If Me.cboYear.Text.Trim.Length <= 0 Then
            MsgBox("Missing year", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            Me.lstClasses.Items.Clear()
            Me.cmdStaffSubject.Connection = conn
            Me.cmdStaffSubject.CommandType = CommandType.Text
            Me.cmdStaffSubject.CommandText = "SELECT * FROM tblClasses WHERE (status=1) AND (year=@year) ORDER BY year,className,stream"
            Me.cmdStaffSubject.Parameters.Clear()
            Me.cmdStaffSubject.Parameters.AddWithValue("@year", Me.cboYear.Text.Trim)
            reader = Me.cmdStaffSubject.ExecuteReader
            If reader.HasRows Then
                While reader.Read
                    li = Me.lstClasses.Items.Add(IIf(DBNull.Value.Equals(reader!className), "", (reader!className)))
                    li.SubItems.Add(IIf(DBNull.Value.Equals(reader!stream), "", (reader!stream)))
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

    Private Sub btnLoad_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLoad.Click
        If Me.txtFullNameView.Text.Trim.Length <= 0 Then
            MsgBox("Employee Name Missing", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboEmpNoView.Text.Trim.Length <= 0 Then
            MsgBox("Employee Number Missing", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboYearView.Text.Trim.Length <= 0 Then
            MsgBox("Year is Missing", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If
        Try
            Me.txtTotalSubjects.Text = ""
            Me.lstStaffSubjectView.Items.Clear()
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            Me.cmdStaffSubject.Connection = conn
            Me.cmdStaffSubject.CommandType = CommandType.Text
            Me.cmdStaffSubject.CommandText = "SELECT * FROM vwAcadStaffSubjects WHERE (year=@year) AND " & _
                vbNewLine & " (empNo=@empNo) ORDER BY subName,year,className,stream"
            Me.cmdStaffSubject.Parameters.Clear()
            Me.cmdStaffSubject.Parameters.AddWithValue("@year", Me.cboYearView.Text.Trim)
            Me.cmdStaffSubject.Parameters.AddWithValue("@empNo", Me.cboEmpNoView.Text.Trim)
            reader = Me.cmdStaffSubject.ExecuteReader
            If reader.HasRows Then
                While reader.Read
                    li = Me.lstStaffSubjectView.Items.Add(IIf(DBNull.Value.Equals(reader!FullName), "", (reader!FullName)))
                    li.SubItems.Add(IIf(DBNull.Value.Equals(reader!subName), "", (reader!subName)))
                    li.SubItems.Add(IIf(DBNull.Value.Equals(reader!className), "", (reader!className)))
                    li.SubItems.Add(IIf(DBNull.Value.Equals(reader!stream), "", (reader!stream)))
                    li.SubItems.Add(IIf(DBNull.Value.Equals(reader!year), "", (reader!year)))
                    li.Tag = IIf(DBNull.Value.Equals(reader!empNo), "", (reader!empNo))
                End While
            End If
            reader.Close()
            Me.txtTotalSubjects.Text = Me.lstStaffSubjectView.Items.Count
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub

    Private Sub cboEmpNoView_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboEmpNoView.SelectedIndexChanged
        If Me.cboEmpNoView.Text.Trim.Length <= 0 Then
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()

            Me.txtFullNameView.Text = ""
            Me.lstStaffSubjectView.Items.Clear()
            Me.txtTotalSubjects.Text = ""

            Me.cmdStaffSubject.Connection = conn
            Me.cmdStaffSubject.CommandType = CommandType.Text
            Me.cmdStaffSubject.CommandText = "SELECT FullName FROM tblSchoolStaff WHERE (status=1) AND (empNo=@empNo)"
            Me.cmdStaffSubject.Parameters.Clear()
            Me.cmdStaffSubject.Parameters.AddWithValue("@empNo", Me.cboEmpNoView.Text.Trim)
            reader = Me.cmdStaffSubject.ExecuteReader
            If reader.HasRows Then
                While reader.Read
                    Me.txtFullNameView.Text = (IIf(DBNull.Value.Equals(reader!FullName), "", (reader!FullName)))
                End While
            End If
            reader.Close()
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            Me.txtTotalSubjects.Text = ""
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub
    Private Sub getMaxNoOfSubjects()
        Me.cmdStaffSubject.Connection = conn
        Me.cmdStaffSubject.CommandType = CommandType.Text
        Me.cmdStaffSubject.CommandText = "SELECT maxNo FROM tblMaxNoOfSubjects WHERE parameter='Teacher' AND class='LESSONS'"
        Me.cmdStaffSubject.Parameters.Clear()
        reader = Me.cmdStaffSubject.ExecuteReader
        If reader.HasRows Then
            While reader.Read
                maxNoOfSubjects = IIf(DBNull.Value.Equals(reader!maxNo), "", (reader!maxNo))
            End While
        Else
            maxNoOfSubjects = 0
        End If
        reader.Close()
    End Sub
    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        Dim diffence As Integer = 0
        If Me.cboEmpNo.Text.Trim.Length <= 0 Then
            MsgBox("Missing Employee Number.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.txtEmpName.Text.Trim.Length <= 0 Then
            MsgBox("Missing Employee Name.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboYear.Text.Trim.Length <= 0 Then
            MsgBox("Missing Year", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.lstClasses.Items.Count <= 0 Then
            MsgBox("Missing classes in the list.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.lstSubject.Items.Count <= 0 Then
            MsgBox("Missing subjects in the list", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.lstSubject.CheckedItems.Count <= 0 Then
            MsgBox("No subject checked in the list", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.lstClasses.CheckedItems.Count <= 0 Then
            MsgBox("No classes checked in the list", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.lstSubject.CheckedItems.Count > 1 Then
            MsgBox("Save one subject at a time", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            numberOfSubjects = 0
            GetTeacherSubjectNo(Me.cboEmpNo.Text.Trim, Me.cboYear.Text.Trim)
            numberOfSubjects = numberOfSubjects + Me.lstClasses.CheckedItems.Count
            maxNoOfSubjects = 0
            getMaxNoOfSubjects()
            diffence = maxNoOfSubjects - numberOfSubjects
            If diffence < 0 Then
                MsgBox("If subjects added teacher will exceed maximum." & _
                       vbNewLine & "The teacher will have " & numberOfSubjects & " Subjects" & _
                       vbNewLine & "Which exceed " & maxNoOfSubjects & " the maximum number of subjects per teacher.", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
                Exit Sub
            End If



            Dim result As MsgBoxResult = MsgBox("Save Record?", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.YesNo, "Confirm Transaction")
            If result = MsgBoxResult.No Then
                Exit Sub
            End If
            rec = 0
            For i = 0 To Me.lstClasses.CheckedItems.Count - 1
                recordExists = True
                checkForExistence()
                If recordExists = True Then
                    MsgBox("Record for Teacher " & Me.txtEmpName.Text.Trim & vbNewLine & "For class " & Me.lstClasses.CheckedItems(i).Text & _
                           Me.lstClasses.CheckedItems(i).SubItems(1).Text & " Not Saved." & vbNewLine & "Duplicate Record Found." & _
                           vbNewLine & "Click Ok To Continue Saving.", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Duplicate Found")
                End If

                subjectIsAllocated = True
                checkIfClassSubjectIsAllocated()
                If subjectIsAllocated = True Then
                    MsgBox("Subject " & Me.lstSubject.CheckedItems(0).SubItems(1).Text.Trim & vbNewLine & "For class " & Me.lstClasses.CheckedItems(i).Text & _
                           Me.lstClasses.CheckedItems(i).SubItems(1).Text & " Not Saved." & vbNewLine & "Its Already Allocated To A different Teacher." & _
                           vbNewLine & "Click Ok To Continue Saving.", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Duplicate Found")
                End If
                If recordExists = False And subjectIsAllocated = False Then
                    Me.queryType = "INSERT"
                    Me.cmdStaffSubject.Connection = conn
                    Me.cmdStaffSubject.CommandType = CommandType.StoredProcedure
                    Me.cmdStaffSubject.CommandText = "sprocStaffSubject"
                    Me.cmdStaffSubject.Parameters.Clear()
                    Me.cmdStaffSubject.Parameters.AddWithValue("@queryType", Me.queryType.Trim)
                    Me.cmdStaffSubject.Parameters.AddWithValue("@empNo", Me.cboEmpNo.Text.Trim)
                    Me.cmdStaffSubject.Parameters.AddWithValue("@subName", Me.lstSubject.CheckedItems(0).SubItems(1).Text)
                    Me.cmdStaffSubject.Parameters.AddWithValue("@year", Me.cboYear.Text.Trim)
                    Me.cmdStaffSubject.Parameters.AddWithValue("@className", Me.lstClasses.CheckedItems(i).Text)
                    Me.cmdStaffSubject.Parameters.AddWithValue("@stream", Me.lstClasses.CheckedItems(i).SubItems(1).Text)
                    Me.cmdStaffSubject.Parameters.AddWithValue("@dateOfReg", Date.Now)
                    Me.cmdStaffSubject.Parameters.AddWithValue("@regBy", userName.Trim)
                    rec = rec + Me.cmdStaffSubject.ExecuteNonQuery
                End If
            Next
            If rec > 0 Then
                MsgBox("Record/s Saved", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "SuccessFull Transactions")
            End If
            clearTexts()
            'enableControls()
            loadCombos()
            loadLists()
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub
    Private Sub checkForExistence()
        Me.cmdStaffSubject.Connection = conn
        Me.cmdStaffSubject.CommandType = CommandType.Text
        Me.cmdStaffSubject.CommandText = "SELECT * FROM vwAcadStaffSubjects WHERE (empNo=@empNo) AND (subName=@subName) " & _
            vbNewLine & " AND (className=@className) AND (year=@year) AND (stream=@stream)"
        Me.cmdStaffSubject.Parameters.Clear()
        Me.cmdStaffSubject.Parameters.AddWithValue("@subName", Me.lstSubject.CheckedItems(0).SubItems(1).Text.Trim)
        Me.cmdStaffSubject.Parameters.AddWithValue("@empNo", Me.cboEmpNo.Text.Trim)
        Me.cmdStaffSubject.Parameters.AddWithValue("@className", Me.lstClasses.CheckedItems(i).Text.Trim)
        Me.cmdStaffSubject.Parameters.AddWithValue("@stream", Me.lstClasses.CheckedItems(i).SubItems(1).Text.Trim)
        Me.cmdStaffSubject.Parameters.AddWithValue("@year", Me.cboYear.Text.Trim)

        reader = Me.cmdStaffSubject.ExecuteReader
        If reader.HasRows Then
            recordExists = True
        Else
            recordExists = False
        End If
        reader.Close()
    End Sub
    Private Sub checkIfClassSubjectIsAllocated()
        Me.cmdStaffSubject.Connection = conn
        Me.cmdStaffSubject.CommandType = CommandType.Text
        Me.cmdStaffSubject.CommandText = "SELECT * FROM vwAcadStaffSubjects WHERE (subName=@subName) AND " & _
            vbNewLine & " (className=@className) AND (year=@year) AND (stream=@stream)"
        Me.cmdStaffSubject.Parameters.Clear()
        Me.cmdStaffSubject.Parameters.AddWithValue("@subName", Me.lstSubject.CheckedItems(0).SubItems(1).Text.Trim)
        Me.cmdStaffSubject.Parameters.AddWithValue("@className", Me.lstClasses.CheckedItems(i).Text.Trim)
        Me.cmdStaffSubject.Parameters.AddWithValue("@stream", Me.lstClasses.CheckedItems(i).SubItems(1).Text.Trim)
        Me.cmdStaffSubject.Parameters.AddWithValue("@year", Me.cboYear.Text.Trim)

        reader = Me.cmdStaffSubject.ExecuteReader
        If reader.HasRows Then
            subjectIsAllocated = True
        Else
            subjectIsAllocated = False
        End If
        reader.Close()
    End Sub
    Private Sub clearTexts()
        Me.cboYear.Items.Clear()
        Me.cboYear.Text = ""
        Me.cboYear.SelectedIndex = -1

        Me.cboContract.Items.Clear()
        Me.cboContract.Text = ""
        Me.cboContract.SelectedIndex = -1

        Me.cboEmpNo.Items.Clear()
        Me.cboEmpNo.Text = ""
        Me.cboEmpNo.SelectedIndex = -1

        Me.cboEmpNoView.Items.Clear()
        Me.cboEmpNoView.Text = ""
        Me.cboEmpNoView.SelectedIndex = -1

        Me.lstClasses.Items.Clear()
        Me.lstStaffSubjectView.Items.Clear()
        Me.lstSubject.Items.Clear()
        Me.txtEmpName.Text = ""
        Me.txtTotalSubjects.Text = ""
        Me.txtSearchTeacher.Text = ""
        Me.txtFullNameView.Text = ""

        Me.cboYearView.Items.Clear()
        Me.cboYearView.Text = ""
        Me.cboYearView.SelectedIndex = -1
    End Sub

    Private Sub btnClear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClear.Click
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            clearTexts()
            loadCombos()
            loadLists()
            'enableControls()
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub
    'Private Sub enableControls()
    '    Me.btnSave.Enabled = True
    '    Me.btnDelete.Enabled = False
    '    Me.cboContract.Enabled = True
    '    Me.cboEmpNo.Enabled = True
    '    Me.lstSubject.Enabled = True
    'End Sub
    'Private Sub DisableControls()
    '    Me.btnSave.Enabled = False
    '    Me.btnDelete.Enabled = True
    '    Me.cboContract.Enabled = False
    '    Me.cboEmpNo.Enabled = False
    '    Me.lstSubject.Enabled = False
    'End Sub

    Private Sub cboContractView_SelectedIndexChanged_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboContractView.SelectedIndexChanged
        If Me.cboContractView.Text.Trim.Length <= 0 Then
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()

            Me.cboEmpNoView.Items.Clear()
            Me.cboEmpNoView.Text = ""
            Me.cboEmpNoView.SelectedIndex = -1

            Me.txtFullNameView.Text = ""
            Me.txtTotalSubjects.Text = ""
            Me.txtSearchTeacher.Text = ""

            Me.lstStaffSubjectView.Items.Clear()

            Me.cmdStaffSubject.Connection = conn
            Me.cmdStaffSubject.CommandType = CommandType.Text
            Me.cmdStaffSubject.CommandText = "SELECT DISTINCT empNo FROM tblSchoolStaff WHERE (status=1) AND (empType='Teaching') AND " & _
                vbNewLine & "(contractType=@contractType) ORDER BY empNo"
            Me.cmdStaffSubject.Parameters.Clear()
            Me.cmdStaffSubject.Parameters.AddWithValue("@contractType", Me.cboContractView.Text.Trim)
            reader = Me.cmdStaffSubject.ExecuteReader
            If reader.HasRows Then
                While reader.Read
                    Me.cboEmpNoView.Items.Add(IIf(DBNull.Value.Equals(reader!empNo), "", (reader!empNo)))
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

    Private Sub txtSearchTeacher_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtSearchTeacher.TextChanged
        If Me.cboYearView.Text.Trim.Length <= 0 Then
            MsgBox("Missing Year", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If
        If Me.txtSearchTeacher.Text.Trim.Length <= 3 Then
            Me.lstStaffSubjectView.Items.Clear()
            Exit Sub
        End If

        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            Try
                Me.txtTotalSubjects.Text = ""
                Me.lstStaffSubjectView.Items.Clear()
                If conn.State = ConnectionState.Closed Then
                    conn.Open()
                End If
                dbconnection()
                Me.cmdStaffSubject.Connection = conn
                Me.cmdStaffSubject.CommandType = CommandType.Text
                Me.cmdStaffSubject.CommandText = "SELECT * FROM vwAcadStaffSubjects WHERE  (year=@year) AND (FullName LIKE @FullName) " & _
                    vbNewLine & " ORDER BY subName,year,className,stream"
                Me.cmdStaffSubject.Parameters.Clear()
                Me.cmdStaffSubject.Parameters.AddWithValue("@year", Me.cboYearView.Text.Trim)
                Me.cmdStaffSubject.Parameters.AddWithValue("@FullName", String.Format("%{0}%", TryCast(Me.txtSearchTeacher.Text.Trim, String).Trim))
                reader = Me.cmdStaffSubject.ExecuteReader
                If reader.HasRows Then
                    While reader.Read
                        li = Me.lstStaffSubjectView.Items.Add(IIf(DBNull.Value.Equals(reader!FullName), "", (reader!FullName)))
                        li.SubItems.Add(IIf(DBNull.Value.Equals(reader!subName), "", (reader!subName)))
                        li.SubItems.Add(IIf(DBNull.Value.Equals(reader!className), "", (reader!className)))
                        li.SubItems.Add(IIf(DBNull.Value.Equals(reader!stream), "", (reader!stream)))
                        li.SubItems.Add(IIf(DBNull.Value.Equals(reader!year), "", (reader!year)))
                    End While
                End If
                reader.Close()
                Me.txtTotalSubjects.Text = ""
            Catch ex As Exception
                MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Finally
                If conn.State = ConnectionState.Open Then
                    conn.Close()
                End If
            End Try
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub
    Private Function GetTeacherSubjectNo(ByVal empNo As String, ByVal yearOf As Integer)
        Me.cmdStaffSubject.Connection = conn
        Me.cmdStaffSubject.CommandText = "SELECT COUNT(empNo) AS subNo FROM vwAcadStaffSubjects WHERE (empNo=@empNo) " & _
            vbNewLine & " AND (year=@year)"
        Me.cmdStaffSubject.CommandType = CommandType.Text
        Me.cmdStaffSubject.Parameters.Clear()
        Me.cmdStaffSubject.Parameters.AddWithValue("@empNo", empNo.Trim)
        Me.cmdStaffSubject.Parameters.AddWithValue("@year", yearOf)
        reader = Me.cmdStaffSubject.ExecuteReader
        If reader.HasRows = True Then
            While reader.Read
                numberOfSubjects = IIf(DBNull.Value.Equals(reader!subNo), "", reader!subNo)
            End While
        ElseIf reader.HasRows = False Then
            numberOfSubjects = 0
        End If
        reader.Close()
    End Function
    Private Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        If Me.lstStaffSubjectView.Items.Count <= 0 Then
            MsgBox("No items to delete.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.lstStaffSubjectView.CheckedItems.Count <= 0 Then
            MsgBox("Check the items to delete.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            Dim result As MsgBoxResult = MsgBox("Delete Record/s?", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.YesNo, "Confirm Transaction")
            If result = MsgBoxResult.No Then
                Exit Sub
            End If
            rec = 0
            i = 0
            For i = 0 To Me.lstStaffSubjectView.CheckedItems.Count - 1
                Me.queryType = "DELETE"
                Me.cmdStaffSubject.Connection = conn
                Me.cmdStaffSubject.CommandType = CommandType.StoredProcedure
                Me.cmdStaffSubject.CommandText = "sprocStaffSubject"
                Me.cmdStaffSubject.Parameters.Clear()
                Me.cmdStaffSubject.Parameters.AddWithValue("@queryType", Me.queryType.Trim)
                Me.cmdStaffSubject.Parameters.AddWithValue("@empNo", Me.lstStaffSubjectView.CheckedItems(i).Tag)
                Me.cmdStaffSubject.Parameters.AddWithValue("@subName", Me.lstStaffSubjectView.CheckedItems(i).SubItems(1).Text)
                Me.cmdStaffSubject.Parameters.AddWithValue("@year", Me.lstStaffSubjectView.CheckedItems(i).SubItems(4).Text)
                Me.cmdStaffSubject.Parameters.AddWithValue("@className", Me.lstStaffSubjectView.CheckedItems(i).SubItems(2).Text)
                Me.cmdStaffSubject.Parameters.AddWithValue("@stream", Me.lstStaffSubjectView.CheckedItems(i).SubItems(3).Text)
                Me.cmdStaffSubject.Parameters.AddWithValue("@dateOfReg", Date.Now)
                Me.cmdStaffSubject.Parameters.AddWithValue("@regBy", userName.Trim)
                rec = rec + Me.cmdStaffSubject.ExecuteNonQuery
            Next
            If rec > 0 Then
                MsgBox("Record/s Deleted", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "SuccessFull Transactions")
            End If
            clearTexts()
            'enableControls()
            loadCombos()
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