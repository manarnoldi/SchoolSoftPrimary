Imports System.Data.SqlClient
Public Class frmSchoolStaff
    Dim passQuery As Boolean = True
    Dim queryType As String = Nothing
    Dim recordExists As Boolean = False
    Dim Rec As Integer = 0
    Dim cmdSchStaff As New SqlCommand
    Dim reader As SqlDataReader
    Private Sub frmSchoolStaff_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            clearTexts()
            loadCombos()
            loadSchoolStaff()
            Me.rbMarriedFalse.Checked = False
            Me.rbMarriedTrue.Checked = False
            Me.rbSexFemale.Checked = False
            Me.rbSexMale.Checked = False
            Me.rbHeadTrue.Checked = False
            Me.rbHeadFalse.Checked = False
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub

    Private Sub frmSchoolStaff_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint
        If Me.IsMdiChild Then
            Dim PnlLoc As New Point
            PnlLoc.X = CInt((Me.Width - Me.pnlSchoolStaff.Width) / 2)
            PnlLoc.Y = CInt((Me.Height - Me.pnlSchoolStaff.Height) / 2.5)
            Me.pnlSchoolStaff.Location = PnlLoc
        Else
            Me.pnlSchoolStaff.Dock = DockStyle.Fill
        End If
    End Sub
    Private Sub loadCombos()
        Me.cboContractType.Items.Clear()
        Me.cboEmpTitle.Items.Clear()
        Me.cboEmpType.Items.Clear()
        Me.cboReligion.Items.Clear()
        Me.cboResidence.Items.Clear()

        Me.cboContractType.SelectedIndex = -1
        Me.cboEmpTitle.SelectedIndex = -1
        Me.cboEmpType.SelectedIndex = -1
        Me.cboReligion.SelectedIndex = -1
        Me.cboResidence.SelectedIndex = -1

        cmdSchStaff.CommandType = CommandType.Text
        cmdSchStaff.Connection = conn
        cmdSchStaff.CommandText = "SELECT DISTINCT title FROM tblSchoolStaff WHERE (status='True') ORDER BY title"
        cmdSchStaff.Parameters.Clear()
        reader = cmdSchStaff.ExecuteReader
        If reader.HasRows Then
            While reader.Read
                Me.cboEmpTitle.Items.Add(IIf(DBNull.Value.Equals(reader!title), "", reader!title))
            End While
        End If
        reader.Close()

        cmdSchStaff.CommandText = "SELECT DISTINCT residence FROM tblSchoolStaff WHERE (status='True') ORDER BY residence"
        cmdSchStaff.Parameters.Clear()
        reader = cmdSchStaff.ExecuteReader
        If reader.HasRows Then
            While reader.Read
                Me.cboResidence.Items.Add(IIf(DBNull.Value.Equals(reader!residence), "", reader!residence))
            End While
        End If
        reader.Close()

        cmdSchStaff.CommandText = "SELECT DISTINCT contractType FROM tblSchoolStaff WHERE (status='True') ORDER BY contractType"
        cmdSchStaff.Parameters.Clear()
        reader = cmdSchStaff.ExecuteReader
        If reader.HasRows Then
            While reader.Read
                Me.cboContractType.Items.Add(IIf(DBNull.Value.Equals(reader!contractType), "", reader!contractType))
            End While
        End If
        reader.Close()

        cmdSchStaff.CommandText = "SELECT DISTINCT empType FROM tblSchoolStaff WHERE (status='True') ORDER BY empType"
        cmdSchStaff.Parameters.Clear()
        reader = cmdSchStaff.ExecuteReader
        If reader.HasRows Then
            While reader.Read
                Me.cboEmpType.Items.Add(IIf(DBNull.Value.Equals(reader!empType), "", reader!empType))
            End While
        End If
        reader.Close()

        cmdSchStaff.CommandText = "SELECT DISTINCT religion FROM tblSchoolStaff WHERE (status='True') ORDER BY religion"
        cmdSchStaff.Parameters.Clear()
        reader = cmdSchStaff.ExecuteReader
        If reader.HasRows Then
            While reader.Read
                Me.cboReligion.Items.Add(IIf(DBNull.Value.Equals(reader!religion), "", reader!religion))
            End While
        End If
        reader.Close()
    End Sub

    Private Sub cboEmpTitle_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboEmpTitle.LostFocus
        If Me.cboEmpTitle.Text = "" Then
            Exit Sub
        End If
        Dim titleString As String = Me.cboEmpTitle.Text
        If titleString.Contains(".") = True Then
            titleString = titleString.Replace(".", "")
            Me.cboEmpTitle.Text = titleString.Trim
        End If
    End Sub

    Private Sub txtEmail_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtEmail.LostFocus
        If Me.txtEmail.Text = "" Then
            Exit Sub
        End If
        If Me.txtEmail.Text.Contains(".") = False Or Me.txtEmail.Text.Contains("@") = False Then
            MsgBox("Enter Correct Email Address", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Missing Details")
            Me.txtEmail.Text = ""
        End If
    End Sub
    Private Sub loadSchoolStaff()
        Me.lstStaff.Items.Clear()
        cmdSchStaff.Connection = conn
        cmdSchStaff.CommandType = CommandType.Text
        cmdSchStaff.CommandText = "SELECT * FROM tblSchoolStaff WHERE (status='True') ORDER BY empNo"
        cmdSchStaff.Parameters.Clear()
        reader = cmdSchStaff.ExecuteReader
        If reader.HasRows Then
            While reader.Read
                li = Me.lstStaff.Items.Add(IIf(DBNull.Value.Equals(reader!empNo), "", reader!empNo))
                li.SubItems.Add(IIf(DBNull.Value.Equals(reader!FullName), "", reader!FullName))
                li.SubItems.Add(IIf(DBNull.Value.Equals(reader!contractType), "", reader!contractType))
                li.SubItems.Add(IIf(DBNull.Value.Equals(reader!empType), "", reader!empType))
                li.SubItems.Add(IIf(DBNull.Value.Equals(reader!phoneNo), "", reader!phoneNo))
                li.SubItems.Add(IIf(DBNull.Value.Equals(reader!idNumber), "", reader!idNumber))
                li.SubItems.Add(IIf(DBNull.Value.Equals(reader!religion), "", reader!religion))
                li.SubItems.Add(IIf(DBNull.Value.Equals(reader!residence), "", reader!residence))
                li.Tag = IIf(DBNull.Value.Equals(reader!empId), "", reader!empId)
            End While
        End If
        reader.Close()
    End Sub
    Private Sub clearTexts()
        Me.DTPDOB.Value = Date.Now
        Me.DTPDOE.Value = Date.Now
        Me.DTPDOR.Value = Date.Now
        Me.rbMarriedFalse.Checked = False
        Me.rbMarriedTrue.Checked = False
        Me.rbSexFemale.Checked = False
        Me.rbSexMale.Checked = False
        Me.rbHeadTrue.Checked = False
        Me.rbHeadFalse.Checked = False
        Me.txtEmpNo.Tag = Nothing
        Me.txtAddress.Text = ""
        Me.txtEmail.Text = ""
        Me.txtEmpNo.Text = ""
        Me.txtFName.Text = ""
        Me.txtIdNo.Text = ""
        Me.txtLName.Text = ""
        Me.txtMName.Text = ""
        Me.txtPhoneNo.Text = ""
        loadCombos()
    End Sub

    Private Sub btnView_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnView.Click
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            loadSchoolStaff()
            clearTexts()
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
            Me.DTPDOB.Enabled = True
            Me.DTPDOE.Enabled = True
            Me.DTPDOR.Enabled = True
            Me.txtEmpNo.Enabled = True
            Me.txtIdNo.Enabled = True
            Me.rbSexFemale.Enabled = True
            Me.rbSexMale.Enabled = True
            Me.txtEmpNo.Tag = Nothing
            Me.cboContractType.Text = ""
            Me.cboEmpTitle.Text = ""
            Me.cboEmpType.Text = ""
            Me.cboReligion.Text = ""
            Me.cboResidence.Text = ""
            Me.btnDelete.Enabled = False
            Me.btnUpdate.Enabled = False
            Me.btnSave.Enabled = True
        End Try
    End Sub
    Private Function checkIfHeadEntered(ByVal empNo As String)
        Dim headFound As Boolean = True
        Me.cmdSchStaff.Connection = conn
        Me.cmdSchStaff.CommandType = CommandType.Text
        Me.cmdSchStaff.CommandText = "SELECT COUNT(empId) FROM tblSchoolStaff WHERE (isHead=1) AND (status=1)"
        Me.cmdSchStaff.Parameters.Clear()
        reader = Me.cmdSchStaff.ExecuteReader
        If reader.HasRows = True Then
            headFound = True
        ElseIf reader.HasRows = False Then
            headFound = False
        End If
        reader.Close()
        Return headFound
    End Function
    Private Sub checkEnteredData()
        passQuery = False
        If Me.cboEmpTitle.Text.Trim.Length <= 0 Then
            MsgBox("Employee Title Missing", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            passQuery = False
            Exit Sub
        ElseIf Me.txtEmpNo.Text.Trim.Length <= 0 Then
            MsgBox("Employee Number Missing", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            passQuery = False
            Exit Sub
        ElseIf (Me.txtFName.Text.Trim.Length <= 0 And Me.txtLName.Text.Length <= 0) Or (Me.txtFName.Text.Trim.Length <= 0 And Me.txtMName.Text.Length <= 0) Or (Me.txtMName.Text.Trim.Length <= 0 And Me.txtLName.Text.Length <= 0) Then
            MsgBox("Enter At Least Two Names", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            passQuery = False
            Exit Sub
        ElseIf Me.cboEmpType.Text.Trim.Length <= 0 Then
            MsgBox("Employee Type Missing", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            passQuery = False
            Exit Sub
        ElseIf Me.cboContractType.Text.Trim.Length <= 0 Then
            MsgBox("Contract Type Missing", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            passQuery = False
            Exit Sub
        ElseIf Me.rbHeadFalse.Checked = False And Me.rbHeadTrue.Checked = False Then
            MsgBox("Select If HeadTeacher", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            passQuery = False
            Exit Sub
        ElseIf Me.rbSexFemale.Checked = False And Me.rbSexMale.Checked = False Then
            MsgBox("Select Staff Sex", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            passQuery = False
            Exit Sub
            'ElseIf Me.txtIdNo.Text.Trim.Length <= 0 Then
            '    MsgBox("Enter Staff Id Number", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            '    passQuery = False
            '    Exit Sub
            'ElseIf DateDiff(DateInterval.Day, Date.Now.Date, Me.DTPDOE.Value.Date) > 0 Then
            '    MsgBox("Date of Employment Cannot be more than Today.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            '    passQuery = False
            '    Exit Sub
            'ElseIf DateDiff(DateInterval.Day, Date.Now.Date, Me.DTPDOR.Value.Date) > 0 Then
            '    MsgBox("Date of Registration Cannot be more than Today.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            '    passQuery = False
            '    Exit Sub
            'ElseIf DateDiff(DateInterval.Year, Me.DTPDOB.Value.Date, Date.Now.Date) < 18 Then
            '    MsgBox("Employee is UnderAge", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            '    passQuery = False
            '    Exit Sub
        End If
        passQuery = True
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub
    Private Sub CheckForExistence()
        dbconnection()
        cmdSchStaff.Connection = conn
        cmdSchStaff.CommandType = CommandType.Text
        cmdSchStaff.CommandText = "SELECT * FROM  tblSchoolStaff WHERE (status='True')AND (empNo=@empNo)" & _
            vbNewLine & " AND (idNumber=@idNumber)"
        cmdSchStaff.Parameters.Clear()
        cmdSchStaff.Parameters.AddWithValue("@empNo", Me.txtEmpNo.Text.Trim)
        cmdSchStaff.Parameters.AddWithValue("@idNumber", Me.txtIdNo.Text.Trim)
        reader = cmdSchStaff.ExecuteReader
        If reader.HasRows Then
            recordExists = True
        Else
            recordExists = False
        End If
        reader.Close()
    End Sub
    Private Sub CheckForExistenceOne()
        Dim MaritalStatus As Boolean = False
        Dim sex As Boolean = True
        If Me.rbMarriedFalse.Checked = True Then
            MaritalStatus = False
        ElseIf Me.rbMarriedTrue.Checked = True Then
            MaritalStatus = True
        End If
        If Me.rbSexFemale.Checked = True Then
            sex = False
        ElseIf Me.rbSexMale.Checked = True Then
            sex = True
        End If
        dbconnection()
        cmdSchStaff.Connection = conn
        cmdSchStaff.CommandType = CommandType.Text
        cmdSchStaff.CommandText = "SELECT * FROM  tblSchoolStaff WHERE (status='True')AND (empNo=@empNo)" & _
            vbNewLine & " AND (idNumber=@idNumber) AND (FName=@FName) AND (MName=@MName) AND (LName=@LName) AND (Email=@Email)" & _
            vbNewLine & " AND (phoneNo=@phoneNo) AND (sex=@sex) AND (empType=@empType) AND (residence=@residence)" & _
            vbNewLine & " AND (contactAddress=@contactAddress) AND (title=@title) AND (dateOfEmployment=@dateOfEmployment)" & _
            vbNewLine & " AND (dateOfBirth=@dateOfBirth) AND (dateOfReg=@dateOfReg) AND (regBy=@regBy)" & _
            vbNewLine & " AND (maritalStatus=@maritalStatus) AND (contractType=@contractType) AND (religion=@religion)"
        cmdSchStaff.Parameters.Clear()
        cmdSchStaff.Parameters.AddWithValue("@empNo", Me.txtEmpNo.Text.Trim)
        cmdSchStaff.Parameters.AddWithValue("@FName", Me.txtFName.Text.Trim)
        cmdSchStaff.Parameters.AddWithValue("@MName", Me.txtMName.Text.Trim)
        cmdSchStaff.Parameters.AddWithValue("@LName", Me.txtLName.Text.Trim)
        cmdSchStaff.Parameters.AddWithValue("@Email", Me.txtEmail.Text.Trim)
        cmdSchStaff.Parameters.AddWithValue("@idNumber", Me.txtIdNo.Text.Trim)
        cmdSchStaff.Parameters.AddWithValue("@phoneNo", Me.txtPhoneNo.Text.Trim)
        cmdSchStaff.Parameters.AddWithValue("@sex", sex)
        cmdSchStaff.Parameters.AddWithValue("@empType", Me.cboEmpType.Text.Trim)
        cmdSchStaff.Parameters.AddWithValue("@residence", Me.cboResidence.Text.Trim)
        cmdSchStaff.Parameters.AddWithValue("@contactAddress", Me.txtAddress.Text.Trim)
        cmdSchStaff.Parameters.AddWithValue("@title", Me.cboEmpTitle.Text.Trim)
        cmdSchStaff.Parameters.AddWithValue("@dateOfEmployment", DTPDOE.Value.Date)
        cmdSchStaff.Parameters.AddWithValue("@dateOfBirth", DTPDOB.Value.Date)
        cmdSchStaff.Parameters.AddWithValue("@dateOfReg", DTPDOE.Value)
        cmdSchStaff.Parameters.AddWithValue("@regBy", userName.Trim)
        cmdSchStaff.Parameters.AddWithValue("@maritalStatus", MaritalStatus)
        cmdSchStaff.Parameters.AddWithValue("@contractType", Me.cboContractType.Text.Trim)
        cmdSchStaff.Parameters.AddWithValue("@religion", Me.cboReligion.Text.Trim)
        reader = cmdSchStaff.ExecuteReader
        If reader.HasRows Then
            recordExists = True
        Else
            recordExists = False
        End If
        reader.Close()
    End Sub
    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        Dim MaritalStatus As Boolean = False
        Dim headFound As Boolean = True
        Dim sex As Boolean = True
        Dim isHead As Boolean = False
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            recordExists = False
            CheckForExistence()

            passQuery = False
            checkEnteredData()
            If passQuery = False Then
                Exit Sub
            End If

            If recordExists = True Then
                MsgBox("Record Exists!", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "UNSuccessFul Transaction")
                Exit Sub
            End If

            If Me.rbHeadTrue.Checked = True Then
                headFound = checkIfHeadEntered(Me.txtEmpNo.Text.Trim)
                If headFound = True Then
                    MsgBox("School head teacher already registered!", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "System Reply")
                    Exit Sub
                End If
            End If
            If Me.rbMarriedFalse.Checked = True Then
                MaritalStatus = False
            ElseIf Me.rbMarriedTrue.Checked = True Then
                MaritalStatus = True
            End If
            If Me.rbSexFemale.Checked = True Then
                sex = False
            ElseIf Me.rbSexMale.Checked = True Then
                sex = True
            End If
            If Me.rbHeadFalse.Checked = True Then
                isHead = False
            ElseIf Me.rbHeadTrue.Checked = True Then
                isHead = True
            End If
            queryType = "INSERT"
            Dim result As MsgBoxResult = MsgBox("Save Record?", MsgBoxStyle.Question + MsgBoxStyle.ApplicationModal + MsgBoxStyle.YesNo, "Confirm Transaction")
            If result = MsgBoxResult.Yes Then
                cmdSchStaff.Connection = conn
                cmdSchStaff.CommandType = CommandType.StoredProcedure
                cmdSchStaff.CommandText = "sprocSchoolStaff"
                cmdSchStaff.Parameters.Clear()
                cmdSchStaff.Parameters.AddWithValue("@queryType", queryType.Trim)
                cmdSchStaff.Parameters.AddWithValue("@empNo", Me.txtEmpNo.Text.Trim)
                cmdSchStaff.Parameters.AddWithValue("@FName", Me.txtFName.Text.Trim)
                cmdSchStaff.Parameters.AddWithValue("@MName", Me.txtMName.Text.Trim)
                cmdSchStaff.Parameters.AddWithValue("@LName", Me.txtLName.Text.Trim)
                cmdSchStaff.Parameters.AddWithValue("@Email", Me.txtEmail.Text.Trim)
                cmdSchStaff.Parameters.AddWithValue("@idNumber", Me.txtIdNo.Text.Trim)
                cmdSchStaff.Parameters.AddWithValue("@phoneNo", Me.txtPhoneNo.Text.Trim)
                cmdSchStaff.Parameters.AddWithValue("@sex", sex)
                cmdSchStaff.Parameters.AddWithValue("@empType", Me.cboEmpType.Text.Trim)
                cmdSchStaff.Parameters.AddWithValue("@residence", Me.cboResidence.Text.Trim)
                cmdSchStaff.Parameters.AddWithValue("@contactAddress", Me.txtAddress.Text.Trim)
                cmdSchStaff.Parameters.AddWithValue("@title", Me.cboEmpTitle.Text.Trim)
                cmdSchStaff.Parameters.AddWithValue("@dateOfEmployment", DTPDOE.Value.Date)
                cmdSchStaff.Parameters.AddWithValue("@dateOfBirth", DTPDOB.Value.Date)
                cmdSchStaff.Parameters.AddWithValue("@dateOfReg", DTPDOE.Value)
                cmdSchStaff.Parameters.AddWithValue("@regBy", userName.Trim)
                cmdSchStaff.Parameters.AddWithValue("@maritalStatus", MaritalStatus)
                cmdSchStaff.Parameters.AddWithValue("@contractType", Me.cboContractType.Text.Trim)
                cmdSchStaff.Parameters.AddWithValue("@religion", Me.cboReligion.Text.Trim)
                cmdSchStaff.Parameters.AddWithValue("@isHead", isHead)
                Rec = cmdSchStaff.ExecuteNonQuery
                If Rec > 0 Then
                    MsgBox("Record Saved!", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "SuccessFul Transaction")
                End If
            End If
            clearTexts()
            loadCombos()
            loadSchoolStaff()
            Me.DTPDOB.Enabled = True
            Me.DTPDOE.Enabled = True
            Me.DTPDOR.Enabled = True
            Me.txtEmpNo.Enabled = True
            Me.txtIdNo.Enabled = True
            Me.rbSexFemale.Enabled = True
            Me.rbSexMale.Enabled = True
            Me.btnSave.Enabled = True
            Me.btnDelete.Enabled = False
            Me.btnUpdate.Enabled = False
            Me.cboContractType.Text = ""
            Me.cboEmpTitle.Text = ""
            Me.cboEmpType.Text = ""
            Me.cboReligion.Text = ""
            Me.cboResidence.Text = ""
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            queryType = Nothing

            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub

    Private Sub lstStaff_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lstStaff.SelectedIndexChanged
        Try
            If Me.lstStaff.SelectedItems.Count = 1 Then
                If conn.State = ConnectionState.Closed Then
                    conn.Open()
                End If
                dbconnection()
                cmdSchStaff.Connection = conn
                cmdSchStaff.CommandType = CommandType.Text
                cmdSchStaff.CommandText = "SELECT * FROM tblSchoolStaff WHERE (status='True') AND (empNo=@empNo)"
                cmdSchStaff.Parameters.Clear()
                cmdSchStaff.Parameters.AddWithValue("@empNo", Me.lstStaff.SelectedItems(0).Text.Trim)
                reader = cmdSchStaff.ExecuteReader
                If reader.HasRows Then
                    While reader.Read
                        Me.txtEmpNo.Tag = IIf(DBNull.Value.Equals(reader!empId), "", reader!empId)
                        Me.cboEmpTitle.Text = IIf(DBNull.Value.Equals(reader!title), "", reader!title)
                        Me.txtEmpNo.Text = IIf(DBNull.Value.Equals(reader!empNo), "", reader!empNo)
                        Me.txtFName.Text = IIf(DBNull.Value.Equals(reader!FName), "", reader!FName)
                        Me.txtMName.Text = IIf(DBNull.Value.Equals(reader!MName), "", reader!MName)
                        Me.txtLName.Text = IIf(DBNull.Value.Equals(reader!LName), "", reader!LName)
                        Me.txtAddress.Text = IIf(DBNull.Value.Equals(reader!contactAddress), "", reader!contactAddress)
                        Me.cboResidence.Text = IIf(DBNull.Value.Equals(reader!residence), "", reader!residence)
                        Me.cboContractType.Text = IIf(DBNull.Value.Equals(reader!contractType), "", reader!contractType)
                        Me.cboEmpType.Text = IIf(DBNull.Value.Equals(reader!empType), "", reader!empType)
                        Me.txtEmail.Text = IIf(DBNull.Value.Equals(reader!Email), "", reader!Email)
                        Me.txtPhoneNo.Text = IIf(DBNull.Value.Equals(reader!phoneNo), "", reader!phoneNo)
                        Me.txtIdNo.Text = IIf(DBNull.Value.Equals(reader!idNumber), "", reader!idNumber)
                        Me.cboReligion.Text = IIf(DBNull.Value.Equals(reader!religion), "", reader!religion)
                        Me.DTPDOB.Value = IIf(DBNull.Value.Equals(reader!dateOfBirth), Date.Now.Date, reader!dateOfBirth)
                        Me.DTPDOE.Value = IIf(DBNull.Value.Equals(reader!dateOfEmployment), Date.Now.Date, reader!dateOfEmployment)
                        Me.DTPDOR.Value = IIf(DBNull.Value.Equals(reader!dateOfReg), Date.Now.Date, reader!dateOfReg)
                        If IIf(DBNull.Value.Equals(reader!maritalStatus), "", reader!maritalStatus) = "True" Then
                            Me.rbMarriedTrue.Checked = True
                        ElseIf IIf(DBNull.Value.Equals(reader!maritalStatus), "", reader!maritalStatus) = "False" Then
                            Me.rbMarriedFalse.Checked = True
                        Else
                            Me.rbMarriedTrue.Checked = False
                            Me.rbMarriedFalse.Checked = False
                        End If
                        If IIf(DBNull.Value.Equals(reader!sex), "", reader!sex) = "True" Then
                            Me.rbSexMale.Checked = True
                        ElseIf IIf(DBNull.Value.Equals(reader!sex), "", reader!sex) = "False" Then
                            Me.rbSexFemale.Checked = True
                        Else
                            Me.rbSexMale.Checked = False
                            Me.rbSexFemale.Checked = False
                        End If
                        If IIf(DBNull.Value.Equals(reader!isHead), "", reader!isHead) = "True" Then
                            Me.rbHeadTrue.Checked = True
                        ElseIf IIf(DBNull.Value.Equals(reader!isHead), "", reader!isHead) = "False" Then
                            Me.rbHeadFalse.Checked = True
                        Else
                            Me.rbHeadFalse.Checked = False
                            Me.rbHeadTrue.Checked = False
                        End If
                    End While
                End If
                reader.Close()
                Me.btnSave.Enabled = False
                Me.btnDelete.Enabled = True
                Me.btnUpdate.Enabled = True
            End If
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            queryType = Nothing
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub

    Private Sub btnUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdate.Click
        Dim headFound As Boolean = True
        Try
            Dim MaritalStatus As Boolean = False
            Dim isHead As Boolean = False
            Dim sex As Boolean = False
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            queryType = "UPDATE"
            recordExists = False
            CheckForExistenceOne()
            If recordExists = True Then
                MsgBox("Record Exists!", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "UNSuccessFul Transaction")
                Exit Sub
            End If

            passQuery = False
            checkEnteredData()
            If passQuery = False Then
                Exit Sub
            End If
            'If Me.rbHeadTrue.Checked = True Then
            '    headFound = checkIfHeadEntered(Me.txtEmpNo.Text.Trim)
            '    If headFound = True Then
            '        MsgBox("School head teacher already registered!", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "System Reply")
            '        Exit Sub
            '    End If
            'End If

            If Me.rbMarriedFalse.Checked = True Then
                MaritalStatus = False
            ElseIf Me.rbMarriedTrue.Checked = True Then
                MaritalStatus = True
            End If
            If Me.rbHeadFalse.Checked = True Then
                isHead = False
            ElseIf Me.rbHeadTrue.Checked = True Then
                isHead = True
            End If

            If Me.rbSexMale.Checked = True Then
                sex = True
            ElseIf Me.rbSexFemale.Checked = True Then
                sex = False
            End If

            Dim result As MsgBoxResult = MsgBox("Update Record?", MsgBoxStyle.Question + MsgBoxStyle.ApplicationModal + MsgBoxStyle.YesNo, "Confirm Transaction")
            If result = MsgBoxResult.Yes Then
                cmdSchStaff.Connection = conn
                cmdSchStaff.CommandType = CommandType.StoredProcedure
                cmdSchStaff.CommandText = "sprocSchoolStaff"
                cmdSchStaff.Parameters.Clear()
                cmdSchStaff.Parameters.AddWithValue("@empId", Me.txtEmpNo.Tag)
                cmdSchStaff.Parameters.AddWithValue("@empNo", Me.txtEmpNo.Text.Trim)
                cmdSchStaff.Parameters.AddWithValue("@queryType", queryType.Trim)
                cmdSchStaff.Parameters.AddWithValue("@FName", Me.txtFName.Text.Trim)
                cmdSchStaff.Parameters.AddWithValue("@MName", Me.txtMName.Text.Trim)
                cmdSchStaff.Parameters.AddWithValue("@LName", Me.txtLName.Text.Trim)
                cmdSchStaff.Parameters.AddWithValue("@Email", Me.txtEmail.Text.Trim)
                cmdSchStaff.Parameters.AddWithValue("@phoneNo", Me.txtPhoneNo.Text.Trim)
                cmdSchStaff.Parameters.AddWithValue("@empType", Me.cboEmpType.Text.Trim)
                cmdSchStaff.Parameters.AddWithValue("@residence", Me.cboResidence.Text.Trim)
                cmdSchStaff.Parameters.AddWithValue("@contactAddress", Me.txtAddress.Text.Trim)
                cmdSchStaff.Parameters.AddWithValue("@title", Me.cboEmpTitle.Text.Trim)
                cmdSchStaff.Parameters.AddWithValue("@regBy", userName.Trim)
                cmdSchStaff.Parameters.AddWithValue("@maritalStatus", MaritalStatus)
                cmdSchStaff.Parameters.AddWithValue("@contractType", Me.cboContractType.Text.Trim)
                cmdSchStaff.Parameters.AddWithValue("@religion", Me.cboReligion.Text.Trim)
                cmdSchStaff.Parameters.AddWithValue("@isHead", isHead)
                cmdSchStaff.Parameters.AddWithValue("@idNumber", Me.txtIdNo.Text.Trim)
                cmdSchStaff.Parameters.AddWithValue("@sex", sex)
                Rec = cmdSchStaff.ExecuteNonQuery
                If Rec > 0 Then
                    MsgBox("Record Updated!", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "SuccessFul Transaction")
                End If
                clearTexts()
                loadCombos()
                loadSchoolStaff()
                Me.DTPDOB.Enabled = True
                Me.DTPDOE.Enabled = True
                Me.DTPDOR.Enabled = True
                Me.txtEmpNo.Enabled = True
                Me.txtIdNo.Enabled = True
                Me.rbSexFemale.Enabled = True
                Me.rbSexMale.Enabled = True
                Me.btnSave.Enabled = True
                Me.btnDelete.Enabled = False
                Me.btnUpdate.Enabled = False
                Me.cboContractType.Text = ""
                Me.cboEmpTitle.Text = ""
                Me.cboEmpType.Text = ""
                Me.cboReligion.Text = ""
                Me.cboResidence.Text = ""
                Me.lstStaff.Items.Clear()
            End If
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            queryType = Nothing
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub

    Private Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            queryType = "DELETE"
            Dim result As MsgBoxResult = MsgBox("Delete Record?", MsgBoxStyle.Question + MsgBoxStyle.ApplicationModal + MsgBoxStyle.YesNo, "Confirm Transaction")
            If result = MsgBoxResult.Yes Then
                cmdSchStaff.Connection = conn
                cmdSchStaff.CommandType = CommandType.StoredProcedure
                cmdSchStaff.CommandText = "sprocSchoolStaff"
                cmdSchStaff.Parameters.Clear()
                cmdSchStaff.Parameters.AddWithValue("@empId", Me.txtEmpNo.Tag)
                cmdSchStaff.Parameters.AddWithValue("@empNo", Me.txtEmpNo.Text.Trim)
                cmdSchStaff.Parameters.AddWithValue("@queryType", queryType.Trim)
                cmdSchStaff.Parameters.AddWithValue("@FName", Me.txtFName.Text.Trim)
                cmdSchStaff.Parameters.AddWithValue("@MName", Me.txtMName.Text.Trim)
                cmdSchStaff.Parameters.AddWithValue("@LName", Me.txtLName.Text.Trim)
                cmdSchStaff.Parameters.AddWithValue("@Email", Me.txtEmail.Text.Trim)
                cmdSchStaff.Parameters.AddWithValue("@phoneNo", Me.txtPhoneNo.Text.Trim)
                cmdSchStaff.Parameters.AddWithValue("@empType", Me.cboEmpType.Text.Trim)
                cmdSchStaff.Parameters.AddWithValue("@residence", Me.cboResidence.Text.Trim)
                cmdSchStaff.Parameters.AddWithValue("@contactAddress", Me.txtAddress.Text.Trim)
                cmdSchStaff.Parameters.AddWithValue("@title", Me.cboEmpTitle.Text.Trim)
                cmdSchStaff.Parameters.AddWithValue("@regBy", userName.Trim)
                cmdSchStaff.Parameters.AddWithValue("@contractType", Me.cboContractType.Text.Trim)
                cmdSchStaff.Parameters.AddWithValue("@religion", Me.cboReligion.Text.Trim)
                Rec = cmdSchStaff.ExecuteNonQuery
                If Rec > 0 Then
                    MsgBox("Record Deleted!", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "SuccessFul Transaction")
                End If
                clearTexts()
                loadCombos()
                loadSchoolStaff()
                Me.DTPDOB.Enabled = True
                Me.DTPDOE.Enabled = True
                Me.DTPDOR.Enabled = True
                Me.txtEmpNo.Enabled = True
                Me.txtIdNo.Enabled = True
                Me.rbSexFemale.Enabled = True
                Me.rbSexMale.Enabled = True
                Me.btnSave.Enabled = True
                Me.btnDelete.Enabled = False
                Me.btnUpdate.Enabled = False
                Me.cboContractType.Text = ""
                Me.cboEmpTitle.Text = ""
                Me.cboEmpType.Text = ""
                Me.cboReligion.Text = ""
                Me.cboResidence.Text = ""
            End If
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            queryType = Nothing
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub

    Private Sub txtEmail_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtEmail.TextChanged

    End Sub
End Class