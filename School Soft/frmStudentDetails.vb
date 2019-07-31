Imports System.Data.SqlClient
Public Class frmStudentDetails
    Dim recordExists As Boolean = True
    Dim passQuery As Boolean = False
    Dim reader As SqlDataReader
    Dim cmdStudDetails As New SqlCommand
    Dim rec As Integer = 0
    Dim queryType As String = Nothing
    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub
    Private Sub radioButtonsUncheck()
        Me.rbBoarderFalse.Checked = False
        Me.rbBoarderTrue.Checked = False
        Me.rbDisabledFalse.Checked = False
        Me.rbDisabledTrue.Checked = False
        Me.rbSexFemale.Checked = False
        Me.rbSexMale.Checked = False
        Me.rbTransportFalse.Checked = False
        Me.rbTransportTrue.Checked = False
    End Sub
    Private Sub frmStudentDetails_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            loadCombos()
            'loadAdmNo()
            'loadStudents()
            radioButtonsUncheck()
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub
    Private Sub ClearTexts()
        Me.lstStudDetails.Tag = Nothing
        Me.cboReligion.Text = ""
        Me.cboReligion.SelectedIndex = -1
        Me.cboResidence.Text = ""
        Me.cboResidence.SelectedIndex = -1
        Me.txtAddress.Text = ""
        Me.txtEmail.Text = ""
        Me.txtFName.Text = ""
        Me.txtIdNo.Text = ""
        Me.txtLName.Text = ""
        Me.txtMName.Text = ""
        Me.txtPhoneNo.Text = ""
        Me.DTPDOA.Value = Date.Now
        Me.DTPDOB.Value = Date.Now
        Me.DTPDOR.Value = Date.Now
        radioButtonsUncheck()
    End Sub
    Private Sub frmStudentDetails_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint
        If Me.IsMdiChild Then
            Dim PnlLoc As New Point
            PnlLoc.X = CInt((Me.Width - Me.pnlStudDetails.Width) / 2)
            PnlLoc.Y = CInt((Me.Height - Me.pnlStudDetails.Height) / 2.5)
            Me.pnlStudDetails.Location = PnlLoc
        Else
            Me.pnlStudDetails.Dock = DockStyle.Fill
        End If
    End Sub
    Private Sub loadAdmNo()
        Me.txtStudNo.Text = ""

        cmdStudDetails.Connection = conn
        cmdStudDetails.CommandType = CommandType.Text
        cmdStudDetails.CommandText = "SELECT ISNULL((max(admNo)),0)+1 AS admNumber FROM tblstudAdmNos"
        cmdStudDetails.Parameters.Clear()
        reader = cmdStudDetails.ExecuteReader
        If reader.HasRows Then
            While reader.Read
                Me.txtStudNo.Text = (IIf(DBNull.Value.Equals(reader!admNumber), "", reader!admNumber))
            End While
        End If
        reader.Close()
    End Sub
    Private Sub loadStudents()
        Dim dateOfBirth As String = Nothing
        Dim dateOfAdmission As String = Nothing
        Dim boarder As String = Nothing
        Dim transPort As String = Nothing
        Dim disabled As String = Nothing
        Me.lstStudDetails.Items.Clear()
        cmdStudDetails.Connection = conn
        cmdStudDetails.CommandType = CommandType.Text
        cmdStudDetails.CommandText = "SELECT * FROM tblStudDetails WHERE (status='True') ORDER BY admNo"
        cmdStudDetails.Parameters.Clear()
        reader = cmdStudDetails.ExecuteReader
        If reader.HasRows Then
            While reader.Read
                dateOfBirth = IIf(DBNull.Value.Equals(reader!dateOfBirth), "", reader!dateOfBirth)
                If Not (dateOfBirth = "") Then
                    dateOfBirth = CDate(dateOfBirth).Day.ToString("00") & "-" & CDate(dateOfBirth).Month.ToString("00") & "-" _
                        & CDate(dateOfBirth).Year.ToString("0000")
                End If

                dateOfAdmission = IIf(DBNull.Value.Equals(reader!dateOfAdmission), "", reader!dateOfAdmission)
                If Not (dateOfAdmission = "") Then
                    dateOfAdmission = CDate(dateOfAdmission).Day.ToString("00") & "-" & CDate(dateOfAdmission).Month.ToString("00") & "-" _
                        & CDate(dateOfAdmission).Year.ToString("0000")
                End If

                If IIf(DBNull.Value.Equals(reader!boarder), "", reader!boarder) = "True" Then
                    boarder = "Boarder"
                Else
                    boarder = "Day Scholar"
                End If

                If IIf(DBNull.Value.Equals(reader!transport), "", reader!transport) = "True" Then
                    transPort = "Transport User"
                Else
                    transPort = "Transport Non-User"
                End If

                If IIf(DBNull.Value.Equals(reader!Disabled), "", reader!Disabled) = "True" Then
                    disabled = "Disabled"
                Else
                    disabled = "Not Disabled"
                End If

                li = Me.lstStudDetails.Items.Add(IIf(DBNull.Value.Equals(reader!admNo), "", reader!admNo))
                li.SubItems.Add(IIf(DBNull.Value.Equals(reader!FName), "", reader!FName))
                li.SubItems.Add(IIf(DBNull.Value.Equals(reader!MName), "", reader!MName))
                li.SubItems.Add(IIf(DBNull.Value.Equals(reader!LName), "", reader!LName))
                li.SubItems.Add(boarder)
                li.SubItems.Add(IIf(DBNull.Value.Equals(reader!residence), "", reader!residence))
                li.SubItems.Add(IIf(DBNull.Value.Equals(reader!religion), "", reader!religion))
                li.SubItems.Add(dateOfBirth)
                li.SubItems.Add(dateOfAdmission)
                li.SubItems.Add(transPort)
                li.SubItems.Add(disabled)
                li.Tag = (IIf(DBNull.Value.Equals(reader!studId), "", reader!studId))
            End While
        End If
        reader.Close()
        'loadAdmNo()
    End Sub
    Private Sub loadCombos()
        cboReligion.Items.Clear()
        cboResidence.Items.Clear()

        cboReligion.Text = ""
        cboResidence.Text = ""

        cmdStudDetails.Connection = conn
        cmdStudDetails.CommandType = CommandType.Text
        cmdStudDetails.CommandText = "SELECT DISTINCT residence FROM tblStudDetails WHERE (status='True') ORDER BY residence"
        cmdStudDetails.Parameters.Clear()
        reader = cmdStudDetails.ExecuteReader
        If reader.HasRows Then
            While reader.Read
                cboResidence.Items.Add(IIf(DBNull.Value.Equals(reader!residence), "", reader!residence))
            End While
        End If
        reader.Close()

        cmdStudDetails.CommandText = "SELECT DISTINCT religion FROM tblStudDetails WHERE (status='True') ORDER BY religion"
        cmdStudDetails.Parameters.Clear()
        reader = cmdStudDetails.ExecuteReader
        If reader.HasRows Then
            While reader.Read
                cboReligion.Items.Add(IIf(DBNull.Value.Equals(reader!religion), "", reader!religion))
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
            loadStudents()
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
            Me.btnSave.Enabled = True
            Me.btnUpdate.Enabled = False
            Me.btnDelete.Enabled = False
            ClearTexts()
        End Try
    End Sub
    Private Sub checkEntry()
        passQuery = False
        'If Me.cboReligion.Text.Trim.Length <= 0 Then
        '    MsgBox("Religion Missing", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        '    passQuery = False
        '    Exit Sub
        If Me.txtStudNo.Text.Trim.Length <= 0 Then
            MsgBox("Admission Number Missing", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            passQuery = False
            'ElseIf Me.cboResidence.Text.Trim.Length <= 0 Then
            'MsgBox("Residence Missing", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            'passQuery = False
            'Exit Sub
        ElseIf (Me.txtFName.Text.Trim.Length <= 0 And Me.txtLName.Text.Length <= 0) Or (Me.txtFName.Text.Trim.Length <= 0 And Me.txtMName.Text.Length <= 0) Or (Me.txtMName.Text.Trim.Length <= 0 And Me.txtLName.Text.Length <= 0) Then
            MsgBox("Enter At Least Two Names", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            passQuery = False
            Exit Sub
        ElseIf Me.rbSexFemale.Checked = False And Me.rbSexMale.Checked = False Then
            MsgBox("Select Student Sex", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            passQuery = False
            Exit Sub
            'ElseIf Me.rbBoarderFalse.Checked = False And Me.rbBoarderTrue.Checked = False Then
            'MsgBox("Select If Border", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            'passQuery = False
            'Exit Sub
            'ElseIf Me.rbDisabledFalse.Checked = False And Me.rbBoarderTrue.Checked = False Then
            'MsgBox("Select If Disabled", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            'passQuery = False
            'Exit Sub
            'ElseIf DateDiff(DateInterval.Day, Date.Now.Date, Me.DTPDOR.Value.Date) > 0 Then
            'MsgBox("Date of Registration Cannot be more than Today.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            'passQuery = False
            'Exit Sub
            'ElseIf DateDiff(DateInterval.Day, Date.Now.Date, Me.DTPDOA.Value.Date) > 0 Then
            'MsgBox("Date of Admission Cannot be more than Today.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            'passQuery = False
            'Exit Sub
            'ElseIf Me.rbTransportFalse.Checked = False And Me.rbTransportTrue.Checked = False Then
            'MsgBox("Select If Requires Transport", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            'passQuery = False
            'Exit Sub
        End If
        passQuery = True
    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        passQuery = False
        checkEntry()
        If passQuery = False Then
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            recordExists = True
            checkForExistence()
            If recordExists = True Then
                MsgBox("Record Exists", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
                Exit Sub
            End If

            Dim result As MsgBoxResult = MsgBox("Save Record?", MsgBoxStyle.Question + MsgBoxStyle.ApplicationModal + MsgBoxStyle.YesNo, "Confirm Transaction")
            If result = MsgBoxResult.Yes Then
                Dim sex As Boolean = Nothing
                Dim boarder As Boolean = Nothing
                Dim transport As Boolean = Nothing
                Dim disabled As Boolean = Nothing

                If Me.rbBoarderTrue.Checked = True Then
                    boarder = True
                ElseIf Me.rbBoarderFalse.Checked = True Then
                    boarder = False
                End If

                If Me.rbSexMale.Checked = True Then
                    sex = True
                ElseIf Me.rbSexFemale.Checked = True Then
                    sex = False
                End If

                If Me.rbTransportTrue.Checked = True Then
                    transport = True
                ElseIf Me.rbTransportFalse.Checked = True Then
                    transport = False
                End If

                If Me.rbDisabledTrue.Checked = True Then
                    disabled = True
                ElseIf Me.rbDisabledFalse.Checked = True Then
                    disabled = False
                End If

                queryType = "INSERT"
                cmdStudDetails.Connection = conn
                cmdStudDetails.CommandType = CommandType.StoredProcedure
                cmdStudDetails.CommandText = "sprocStudentDetails"
                cmdStudDetails.Parameters.Clear()
                cmdStudDetails.Parameters.AddWithValue("@queryType", queryType.Trim)
                cmdStudDetails.Parameters.AddWithValue("@admNo", Me.txtStudNo.Text.Trim)
                cmdStudDetails.Parameters.AddWithValue("@FName", Me.txtFName.Text.Trim)
                cmdStudDetails.Parameters.AddWithValue("@MName", Me.txtMName.Text.Trim)
                cmdStudDetails.Parameters.AddWithValue("@LName", Me.txtLName.Text.Trim)
                cmdStudDetails.Parameters.AddWithValue("@Email", Me.txtEmail.Text.Trim)
                cmdStudDetails.Parameters.AddWithValue("@idNumber", Me.txtIdNo.Text.Trim)
                cmdStudDetails.Parameters.AddWithValue("@phone", Me.txtPhoneNo.Text.Trim)
                cmdStudDetails.Parameters.AddWithValue("@sex", sex)
                cmdStudDetails.Parameters.AddWithValue("@boarder", boarder)
                cmdStudDetails.Parameters.AddWithValue("@residence", Me.cboResidence.Text.Trim)
                cmdStudDetails.Parameters.AddWithValue("@transport", transport)
                cmdStudDetails.Parameters.AddWithValue("@disabled", disabled)
                cmdStudDetails.Parameters.AddWithValue("@dateOfAdmission", DTPDOA.Value.Date)
                cmdStudDetails.Parameters.AddWithValue("@dateOfBirth", DTPDOB.Value.Date)
                cmdStudDetails.Parameters.AddWithValue("@dateOfReg", DTPDOR.Value)
                cmdStudDetails.Parameters.AddWithValue("@regBy", userName.Trim)
                cmdStudDetails.Parameters.AddWithValue("@religion", Me.cboReligion.Text.Trim)
                cmdStudDetails.Parameters.AddWithValue("@address", Me.txtAddress.Text.Trim)
                rec = cmdStudDetails.ExecuteNonQuery
                If rec > 0 Then
                    MsgBox("Record Saved!", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "SuccessFul Transaction")
                End If
            End If
            ClearTexts()
            loadCombos()
            'loadAdmNo()
            'loadStudents()
            radioButtonsUncheck()
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            queryType = Nothing
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
            Me.btnDelete.Enabled = False
            Me.btnSave.Enabled = True
            Me.btnUpdate.Enabled = False
        End Try
    End Sub
    Private Sub checkForExistence()
        cmdStudDetails.Connection = conn
        cmdStudDetails.CommandType = CommandType.Text
        cmdStudDetails.CommandText = "SELECT * FROM tblStudDetails WHERE (status='True') AND (admNo=@admNo)"
        cmdStudDetails.Parameters.Clear()
        cmdStudDetails.Parameters.AddWithValue("@admNo", Me.txtStudNo.Text.Trim)
        reader = cmdStudDetails.ExecuteReader
        If reader.HasRows Then
            recordExists = True
        Else
            recordExists = False
        End If
        reader.Close()
    End Sub
    Private Sub checkForExistenceOne()
        Dim sex As Boolean = Nothing
        Dim boarder As Boolean = Nothing
        Dim transport As Boolean = Nothing
        Dim disabled As Boolean = Nothing

        If Me.rbBoarderTrue.Checked = True Then
            boarder = True
        ElseIf Me.rbBoarderFalse.Checked = True Then
            boarder = False
        End If

        If Me.rbTransportTrue.Checked = True Then
            transport = True
        ElseIf Me.rbTransportFalse.Checked = True Then
            transport = False
        End If

        If Me.rbDisabledTrue.Checked = True Then
            disabled = True
        ElseIf Me.rbDisabledFalse.Checked = True Then
            disabled = False
        End If

        If Me.rbSexMale.Checked = True Then
            sex = True
        ElseIf Me.rbSexFemale.Checked = True Then
            sex = False
        End If

        cmdStudDetails.Connection = conn
        cmdStudDetails.CommandType = CommandType.Text
        cmdStudDetails.CommandText = "SELECT * FROM tblStudDetails WHERE (status='True') AND (admNo=@admNo) AND (FName=@FName)" &
            vbNewLine & "AND (MName=@MName) AND (LName=@LName) AND (phone=@phone) AND (idNumber=@idNumber) AND (boarder=@boarder)" &
            vbNewLine & " AND (transport=@transport) AND (email=@email) AND (residence=@residence) AND (religion=@religion) " &
            vbNewLine & " AND (address=@address) AND (disabled=@disabled) AND (dateOfBirth=@dateOfBirth) AND (sex=@sex)"
        cmdStudDetails.Parameters.Clear()
        cmdStudDetails.Parameters.AddWithValue("@admNo", Me.txtStudNo.Text.Trim)
        cmdStudDetails.Parameters.AddWithValue("@FName", Me.txtFName.Text.Trim)
        cmdStudDetails.Parameters.AddWithValue("@MName", Me.txtMName.Text.Trim)
        cmdStudDetails.Parameters.AddWithValue("@LName", Me.txtLName.Text.Trim)
        cmdStudDetails.Parameters.AddWithValue("@Email", Me.txtEmail.Text.Trim)
        cmdStudDetails.Parameters.AddWithValue("@idNumber", Me.txtIdNo.Text.Trim)
        cmdStudDetails.Parameters.AddWithValue("@phone", Me.txtPhoneNo.Text.Trim)
        cmdStudDetails.Parameters.AddWithValue("@boarder", boarder)
        cmdStudDetails.Parameters.AddWithValue("@residence", Me.cboResidence.Text.Trim)
        cmdStudDetails.Parameters.AddWithValue("@transport", transport)
        cmdStudDetails.Parameters.AddWithValue("@disabled", disabled)
        cmdStudDetails.Parameters.AddWithValue("@dateOfBirth", DTPDOB.Value.Date)
        cmdStudDetails.Parameters.AddWithValue("@religion", Me.cboReligion.Text.Trim)
        cmdStudDetails.Parameters.AddWithValue("@address", Me.txtAddress.Text.Trim)
        cmdStudDetails.Parameters.AddWithValue("@sex", sex)
        reader = cmdStudDetails.ExecuteReader
        If reader.HasRows Then
            recordExists = True
        Else
            recordExists = False
        End If
        reader.Close()
    End Sub

    Private Sub lstStudDetails_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lstStudDetails.Click
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()

            If Me.lstStudDetails.SelectedItems.Count = 1 Then
                cmdStudDetails.Connection = conn
                cmdStudDetails.CommandType = CommandType.Text
                cmdStudDetails.CommandText = "SELECT * FROM tblStudDetails WHERE (status='True') AND (studId=@studId)"
                cmdStudDetails.Parameters.Clear()
                cmdStudDetails.Parameters.AddWithValue("@studId", Me.lstStudDetails.SelectedItems(0).Tag)
                reader = cmdStudDetails.ExecuteReader
                If reader.HasRows Then
                    While reader.Read
                        Me.txtStudNo.Text = IIf(DBNull.Value.Equals(reader!admNo), "", reader!admNo)
                        Me.txtFName.Text = IIf(DBNull.Value.Equals(reader!FName), "", reader!FName)
                        Me.txtMName.Text = IIf(DBNull.Value.Equals(reader!MName), "", reader!MName)
                        Me.txtLName.Text = IIf(DBNull.Value.Equals(reader!LName), "", reader!LName)
                        Me.txtAddress.Text = IIf(DBNull.Value.Equals(reader!address), "", reader!address)
                        Me.txtPhoneNo.Text = IIf(DBNull.Value.Equals(reader!phone), "", reader!phone)
                        Me.txtIdNo.Text = IIf(DBNull.Value.Equals(reader!idNumber), "", reader!idNumber)
                        Me.txtEmail.Text = IIf(DBNull.Value.Equals(reader!email), "", reader!email)
                        Me.cboResidence.Text = IIf(DBNull.Value.Equals(reader!residence), "", reader!residence)
                        Me.cboReligion.Text = IIf(DBNull.Value.Equals(reader!religion), "", reader!religion)
                        Me.DTPDOA.Value = IIf(DBNull.Value.Equals(reader!dateOfAdmission), Date.Now, reader!dateOfAdmission)
                        Me.DTPDOB.Value = IIf(DBNull.Value.Equals(reader!dateOfBirth), Date.Now, reader!dateOfBirth)
                        Me.DTPDOR.Value = IIf(DBNull.Value.Equals(reader!dateOfReg), Date.Now, reader!dateOfReg)
                        Me.lstStudDetails.Tag = Me.lstStudDetails.SelectedItems(0).Tag

                        Dim sex As String = IIf(DBNull.Value.Equals(reader!sex), "", reader!sex)
                        Dim boarder As String = IIf(DBNull.Value.Equals(reader!boarder), "", reader!boarder)
                        Dim disabled As String = IIf(DBNull.Value.Equals(reader!disabled), "", reader!disabled)
                        Dim transport As String = IIf(DBNull.Value.Equals(reader!transport), "", reader!transport)

                        radioButtonsUncheck()
                        If sex = "True" Or sex = "False" Then
                            If sex = "True" Then
                                Me.rbSexMale.Checked = True
                            Else
                                Me.rbSexFemale.Checked = True
                            End If
                        End If
                        If boarder = "True" Or boarder = "False" Then
                            If boarder = "True" Then
                                Me.rbBoarderTrue.Checked = True
                            Else
                                Me.rbBoarderFalse.Checked = True
                            End If
                        End If
                        If disabled = "True" Or disabled = "False" Then
                            If disabled = True Then
                                Me.rbDisabledTrue.Checked = True
                            Else
                                Me.rbDisabledFalse.Checked = True
                            End If
                        End If
                        If transport = "True" Or transport = "False" Then
                            If transport = "True" Then
                                Me.rbTransportTrue.Checked = True
                            Else
                                Me.rbTransportFalse.Checked = True
                            End If
                        End If
                    End While
                End If
                reader.Close()
                Me.btnSave.Enabled = False
                Me.btnUpdate.Enabled = True
                Me.btnDelete.Enabled = True
            End If
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try

    End Sub

    Private Sub btnUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdate.Click
        passQuery = False
        checkEntry()
        If passQuery = False Then
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            recordExists = True
            checkForExistenceOne()
            If recordExists = True Then
                MsgBox("Record Exists", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
                Exit Sub
            End If

            Dim result As MsgBoxResult = MsgBox("Update Record?", MsgBoxStyle.Question + MsgBoxStyle.ApplicationModal + MsgBoxStyle.YesNo, "Confirm Transaction")
            If result = MsgBoxResult.Yes Then
                Dim sex As Boolean = Nothing
                Dim boarder As Boolean = Nothing
                Dim transport As Boolean = Nothing
                Dim disabled As Boolean = Nothing

                If Me.rbBoarderTrue.Checked = True Then
                    boarder = True
                ElseIf Me.rbBoarderFalse.Checked = True Then
                    boarder = False
                End If

                If Me.rbSexMale.Checked = True Then
                    sex = True
                ElseIf Me.rbSexFemale.Checked = True Then
                    sex = False
                End If

                If Me.rbTransportTrue.Checked = True Then
                    transport = True
                ElseIf Me.rbTransportFalse.Checked = True Then
                    transport = False
                End If

                If Me.rbDisabledTrue.Checked = True Then
                    disabled = True
                ElseIf Me.rbDisabledFalse.Checked = True Then
                    disabled = False
                End If

                queryType = "UPDATE"
                cmdStudDetails.Connection = conn
                cmdStudDetails.CommandType = CommandType.StoredProcedure
                cmdStudDetails.CommandText = "sprocStudentDetails"
                cmdStudDetails.Parameters.Clear()
                cmdStudDetails.Parameters.AddWithValue("@queryType", queryType.Trim)
                cmdStudDetails.Parameters.AddWithValue("@admNo", Me.txtStudNo.Text.Trim)
                cmdStudDetails.Parameters.AddWithValue("@FName", Me.txtFName.Text.Trim)
                cmdStudDetails.Parameters.AddWithValue("@MName", Me.txtMName.Text.Trim)
                cmdStudDetails.Parameters.AddWithValue("@LName", Me.txtLName.Text.Trim)
                cmdStudDetails.Parameters.AddWithValue("@Email", Me.txtEmail.Text.Trim)
                cmdStudDetails.Parameters.AddWithValue("@idNumber", Me.txtIdNo.Text.Trim)
                cmdStudDetails.Parameters.AddWithValue("@phone", Me.txtPhoneNo.Text.Trim)
                cmdStudDetails.Parameters.AddWithValue("@sex", sex)
                cmdStudDetails.Parameters.AddWithValue("@boarder", boarder)
                cmdStudDetails.Parameters.AddWithValue("@residence", Me.cboResidence.Text.Trim)
                cmdStudDetails.Parameters.AddWithValue("@transport", transport)
                cmdStudDetails.Parameters.AddWithValue("@disabled", disabled)
                cmdStudDetails.Parameters.AddWithValue("@dateOfAdmission", DTPDOA.Value.Date)
                cmdStudDetails.Parameters.AddWithValue("@dateOfBirth", DTPDOB.Value.Date)
                cmdStudDetails.Parameters.AddWithValue("@dateOfReg", DTPDOR.Value)
                cmdStudDetails.Parameters.AddWithValue("@regBy", userName.Trim)
                cmdStudDetails.Parameters.AddWithValue("@religion", Me.cboReligion.Text.Trim)
                cmdStudDetails.Parameters.AddWithValue("@address", Me.txtAddress.Text.Trim)
                cmdStudDetails.Parameters.AddWithValue("@studId", Me.lstStudDetails.Tag)
                rec = cmdStudDetails.ExecuteNonQuery
                If rec > 0 Then
                    MsgBox("Record Updated!", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "SuccessFul Transaction")
                End If
            End If
            ClearTexts()
            loadCombos()
            'loadAdmNo()
            'loadStudents()
            radioButtonsUncheck()
            Me.btnDelete.Enabled = False
            Me.btnSave.Enabled = True
            Me.btnUpdate.Enabled = False
            Me.lstStudDetails.Items.Clear()
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            queryType = Nothing
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If

        End Try
    End Sub

    Private Sub StudentImageToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        frmStudImages.MdiParent = frmHome
        frmStudImages.Show()
    End Sub

    Private Sub StudentClassToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles StudentClassToolStripMenuItem.Click
        If Me.lstStudDetails.SelectedItems.Count <> 1 Then
            MsgBox("Assign One Student At a time.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        Else
            frmStudClass.MdiParent = frmHome
            frmStudClass.Show()
            frmStudClass.txtAdmNo.Text = Me.lstStudDetails.SelectedItems(0).Text.Trim
            frmStudClass.btnLoad.PerformClick()
            frmStudClass.btnView.PerformClick()
        End If
    End Sub

    Private Sub txtQuickSearchAdmNo_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtQuickSearchAdmNo.TextChanged
        Me.lstStudDetails.Items.Clear()
        If Me.txtQuickSearchAdmNo.Text.Trim.Length <= 0 Then
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            Dim dateOfBirth As String = Nothing
            Dim dateOfAdmission As String = Nothing
            Dim boarder As String = Nothing
            Dim transPort As String = Nothing
            Dim disabled As String = Nothing
            Me.lstStudDetails.Items.Clear()
            cmdStudDetails.Connection = conn
            cmdStudDetails.CommandType = CommandType.Text
            cmdStudDetails.CommandText = "SELECT * FROM tblStudDetails WHERE (status='True') AND (admNo LIKE @admNo) ORDER BY admNo"
            cmdStudDetails.Parameters.Clear()
            Me.cmdStudDetails.Parameters.AddWithValue("admNo", String.Format("%{0}%", TryCast(Me.txtQuickSearchAdmNo.Text.Trim, String).Trim))
            reader = cmdStudDetails.ExecuteReader
            If reader.HasRows Then
                While reader.Read
                    dateOfBirth = IIf(DBNull.Value.Equals(reader!dateOfBirth), "", reader!dateOfBirth)
                    If Not (dateOfBirth = "") Then
                        dateOfBirth = CDate(dateOfBirth).Day.ToString("00") & "-" & CDate(dateOfBirth).Month.ToString("00") & "-" _
                            & CDate(dateOfBirth).Year.ToString("0000")
                    End If

                    dateOfAdmission = IIf(DBNull.Value.Equals(reader!dateOfAdmission), "", reader!dateOfAdmission)
                    If Not (dateOfAdmission = "") Then
                        dateOfAdmission = CDate(dateOfAdmission).Day.ToString("00") & "-" & CDate(dateOfAdmission).Month.ToString("00") & "-" _
                            & CDate(dateOfAdmission).Year.ToString("0000")
                    End If

                    If IIf(DBNull.Value.Equals(reader!boarder), "", reader!boarder) = "True" Then
                        boarder = "Boarder"
                    Else
                        boarder = "Day Scholar"
                    End If

                    If IIf(DBNull.Value.Equals(reader!transport), "", reader!transport) = "True" Then
                        transPort = "Transport User"
                    Else
                        transPort = "Transport Non-User"
                    End If

                    If IIf(DBNull.Value.Equals(reader!Disabled), "", reader!Disabled) = "True" Then
                        disabled = "Disabled"
                    Else
                        disabled = "Not Disabled"
                    End If

                    li = Me.lstStudDetails.Items.Add(IIf(DBNull.Value.Equals(reader!admNo), "", reader!admNo))
                    li.SubItems.Add(IIf(DBNull.Value.Equals(reader!FName), "", reader!FName))
                    li.SubItems.Add(IIf(DBNull.Value.Equals(reader!MName), "", reader!MName))
                    li.SubItems.Add(IIf(DBNull.Value.Equals(reader!LName), "", reader!LName))
                    li.SubItems.Add(boarder)
                    li.SubItems.Add(IIf(DBNull.Value.Equals(reader!residence), "", reader!residence))
                    li.SubItems.Add(IIf(DBNull.Value.Equals(reader!religion), "", reader!religion))
                    li.SubItems.Add(dateOfBirth)
                    li.SubItems.Add(dateOfAdmission)
                    li.SubItems.Add(transPort)
                    li.SubItems.Add(disabled)
                    li.Tag = (IIf(DBNull.Value.Equals(reader!studId), "", reader!studId))
                End While
            End If
            reader.Close()
            'loadAdmNo()
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub

    Private Sub StudentFormerSchoolToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles StudentFormerSchoolToolStripMenuItem.Click
        If Me.lstStudDetails.SelectedItems.Count <> 1 Then
            MsgBox("Assign One Student At a time.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        Else
            frmStudFormSchDetails.MdiParent = frmHome
            frmStudFormSchDetails.Show()
            frmStudFormSchDetails.txtStudNo.Text = Me.lstStudDetails.SelectedItems(0).Text.Trim
            frmStudFormSchDetails.btnView.PerformClick()
            frmStudFormSchDetails.btnLoad.PerformClick()
        End If
    End Sub

    Private Sub ParentDetailsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ParentDetailsToolStripMenuItem.Click
        If Me.lstStudDetails.SelectedItems.Count <> 1 Then
            MsgBox("Assign One Student At a time.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        Else
            frmStudParents.MdiParent = frmHome
            frmStudParents.Show()
            frmStudParents.txtStudNo.Text = Me.lstStudDetails.SelectedItems(0).Text.Trim
            frmStudParents.btnView.PerformClick()
            frmStudParents.btnLoad.PerformClick()
        End If
    End Sub

    Private Sub ExitToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitToolStripMenuItem.Click
        Me.Close()
    End Sub

    Private Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        passQuery = False
        checkEntry()
        If passQuery = False Then
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()

            Dim result As MsgBoxResult = MsgBox("Delete Record?", MsgBoxStyle.Question + MsgBoxStyle.ApplicationModal + MsgBoxStyle.YesNo, "Confirm Transaction")
            If result = MsgBoxResult.Yes Then
                Dim sex As Boolean = Nothing
                Dim boarder As Boolean = Nothing
                Dim transport As Boolean = Nothing
                Dim disabled As Boolean = Nothing

                If Me.rbBoarderTrue.Checked = True Then
                    boarder = True
                ElseIf Me.rbBoarderFalse.Checked = True Then
                    boarder = False
                End If

                If Me.rbSexMale.Checked = True Then
                    sex = True
                ElseIf Me.rbSexFemale.Checked = True Then
                    sex = False
                End If

                If Me.rbTransportTrue.Checked = True Then
                    transport = True
                ElseIf Me.rbTransportFalse.Checked = True Then
                    transport = False
                End If

                If Me.rbDisabledTrue.Checked = True Then
                    disabled = True
                ElseIf Me.rbDisabledFalse.Checked = True Then
                    disabled = False
                End If

                queryType = "DELETE"
                cmdStudDetails.Connection = conn
                cmdStudDetails.CommandType = CommandType.StoredProcedure
                cmdStudDetails.CommandText = "sprocStudentDetails"
                cmdStudDetails.Parameters.Clear()
                cmdStudDetails.Parameters.AddWithValue("@queryType", queryType.Trim)
                cmdStudDetails.Parameters.AddWithValue("@admNo", Me.txtStudNo.Text.Trim)
                cmdStudDetails.Parameters.AddWithValue("@FName", Me.txtFName.Text.Trim)
                cmdStudDetails.Parameters.AddWithValue("@MName", Me.txtMName.Text.Trim)
                cmdStudDetails.Parameters.AddWithValue("@LName", Me.txtLName.Text.Trim)
                cmdStudDetails.Parameters.AddWithValue("@Email", Me.txtEmail.Text.Trim)
                cmdStudDetails.Parameters.AddWithValue("@idNumber", Me.txtIdNo.Text.Trim)
                cmdStudDetails.Parameters.AddWithValue("@phone", Me.txtPhoneNo.Text.Trim)
                cmdStudDetails.Parameters.AddWithValue("@sex", sex)
                cmdStudDetails.Parameters.AddWithValue("@boarder", boarder)
                cmdStudDetails.Parameters.AddWithValue("@residence", Me.cboResidence.Text.Trim)
                cmdStudDetails.Parameters.AddWithValue("@transport", transport)
                cmdStudDetails.Parameters.AddWithValue("@disabled", disabled)
                cmdStudDetails.Parameters.AddWithValue("@dateOfAdmission", DTPDOA.Value.Date)
                cmdStudDetails.Parameters.AddWithValue("@dateOfBirth", DTPDOB.Value.Date)
                cmdStudDetails.Parameters.AddWithValue("@dateOfReg", DTPDOR.Value)
                cmdStudDetails.Parameters.AddWithValue("@regBy", userName.Trim)
                cmdStudDetails.Parameters.AddWithValue("@religion", Me.cboReligion.Text.Trim)
                cmdStudDetails.Parameters.AddWithValue("@address", Me.txtAddress.Text.Trim)
                cmdStudDetails.Parameters.AddWithValue("@studId", Me.lstStudDetails.Tag)
                rec = cmdStudDetails.ExecuteNonQuery
                If rec > 0 Then
                    MsgBox("Record Deleted!", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "SuccessFul Transaction")
                End If
            End If
            ClearTexts()
            loadCombos()
            'loadAdmNo()
            'loadStudents()
            radioButtonsUncheck()
            Me.btnDelete.Enabled = False
            Me.btnSave.Enabled = True
            Me.btnUpdate.Enabled = False
            Me.lstStudDetails.Items.Clear()
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            queryType = Nothing
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If

        End Try
    End Sub

    Private Sub lstStudDetails_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lstStudDetails.SelectedIndexChanged

    End Sub
End Class