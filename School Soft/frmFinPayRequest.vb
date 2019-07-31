Imports System.Data.SqlClient
Public Class frmFinPayRequest
    Dim reader As SqlDataReader
    Dim rec As Integer
    Dim cmdPayRequest As New SqlCommand
    Dim payInitials As String = ""
    Private Sub frmFinPayRequest_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
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

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub loadVoucherNo()
        Dim term As String = Nothing
        If Me.cboPayTerm.Text.Trim = "TERM 1" Then
            term = "01"
        ElseIf Me.cboPayTerm.Text.Trim = "TERM 2" Then
            term = "02"
        ElseIf Me.cboPayTerm.Text.Trim = "TERM 3" Then
            term = "03"
        End If

        Me.txtVoucherNo.Text = ""
        Me.txtVoucherNo.Tag = Nothing

        Me.cmdPayRequest.Connection = conn
        Me.cmdPayRequest.CommandType = CommandType.Text
        Me.cmdPayRequest.CommandText = "SELECT TOP 1 initials FROM tblSchoolDetails"
        Me.cmdPayRequest.Parameters.Clear()
        reader = Me.cmdPayRequest.ExecuteReader
        While reader.Read
            Me.payInitials = IIf(DBNull.Value.Equals(reader!initials), "", reader!initials)
        End While
        reader.Close()


        Me.cmdPayRequest.CommandText = "SELECT ISNULL(MAX(voucherNo),0)+1 AS newRecNo FROM vwFinPayments " &
               vbNewLine & " WHERE (termName=@termName) AND (year=@year)"
        Me.cmdPayRequest.Parameters.Clear()
        Me.cmdPayRequest.Parameters.AddWithValue("@termName", Me.cboPayTerm.Text.Trim)
        Me.cmdPayRequest.Parameters.AddWithValue("@year", Me.cboPayYear.Text.Trim)
        reader = Me.cmdPayRequest.ExecuteReader
        While reader.Read
            Me.txtVoucherNo.Text = payInitials & Me.cboPayYear.Text.Trim.Substring(2, 2) & term.Trim & "/" &
                CInt(IIf(DBNull.Value.Equals(reader!newRecNo), "", reader!newRecNo)).ToString("0000")
            Me.txtVoucherNo.Tag = (IIf(DBNull.Value.Equals(reader!newRecNo), "", reader!newRecNo))
        End While
        reader.Close()
    End Sub

    Private Sub loadCombos()

        Me.cboPayTerm.Items.Clear()
        Me.cboPayTerm.Text = ""
        Me.cboPayTerm.SelectedIndex = -1

        Me.cboPayYear.Items.Clear()
        Me.cboPayYear.Text = ""
        Me.cboPayYear.SelectedIndex = -1

        Me.cboPayMode.Items.Clear()
        Me.cboPayMode.Text = ""
        Me.cboPayMode.SelectedIndex = -1

        Me.cboExpCat.Items.Clear()
        Me.cboExpCat.Text = ""
        Me.cboExpCat.SelectedIndex = -1

        Me.cmdPayRequest.Connection = conn
        Me.cmdPayRequest.CommandType = CommandType.Text
        Me.cmdPayRequest.CommandText = "SELECT DISTINCT termName FROM   tblSchoolCalendar WHERE (status=1) ORDER BY termName"
        Me.cmdPayRequest.Parameters.Clear()
        reader = Me.cmdPayRequest.ExecuteReader
        While reader.Read
            Me.cboPayTerm.Items.Add(IIf(DBNull.Value.Equals(reader!termName), "", reader!termName))
        End While
        reader.Close()

        Me.cmdPayRequest.CommandText = "SELECT DISTINCT year FROM   tblSchoolCalendar WHERE (status=1) ORDER BY year"
        Me.cmdPayRequest.Parameters.Clear()
        reader = Me.cmdPayRequest.ExecuteReader
        While reader.Read
            Me.cboPayYear.Items.Add(IIf(DBNull.Value.Equals(reader!year), "", reader!year))
        End While
        reader.Close()

        Me.cmdPayRequest.CommandText = "SELECT DISTINCT modeName FROM tblFinPayModes ORDER BY modeName"
        Me.cmdPayRequest.Parameters.Clear()
        reader = Me.cmdPayRequest.ExecuteReader
        While reader.Read
            Me.cboPayMode.Items.Add(IIf(DBNull.Value.Equals(reader!modeName), "", reader!modeName))
        End While
        reader.Close()

        Me.cmdPayRequest.CommandText = "SELECT DISTINCT expCategoryName FROM tblFinExpCategory ORDER BY expCategoryName"
        Me.cmdPayRequest.Parameters.Clear()
        reader = Me.cmdPayRequest.ExecuteReader
        While reader.Read
            Me.cboExpCat.Items.Add(IIf(DBNull.Value.Equals(reader!expCategoryName), "", reader!expCategoryName))
        End While
        reader.Close()
    End Sub
    Private Sub frmFinPayRequest_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint
        If Me.IsMdiChild Then
            Dim PnlLoc As New Point
            PnlLoc.X = CInt((Me.Width - Me.pnlPayReq.Width) / 2)
            PnlLoc.Y = CInt((Me.Height - Me.pnlPayReq.Height) / 2.5)
            Me.pnlPayReq.Location = PnlLoc
        Else
            Me.pnlPayReq.Dock = DockStyle.Fill
        End If
    End Sub

    Private Sub cboPayTerm_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboPayTerm.SelectedIndexChanged, cboPayYear.SelectedIndexChanged
        If Me.cboPayTerm.Text.Trim.Length > 0 And Me.cboPayYear.Text.Trim.Length > 0 Then
            Try
                If conn.State = ConnectionState.Closed Then
                    conn.Open()
                End If
                dbconnection()
                loadVoucherNo()
            Catch ex As Exception
                MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Finally
                If conn.State = ConnectionState.Open Then
                    conn.Close()
                End If
            End Try
        End If
    End Sub

    Private Sub cboPayMode_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboPayMode.SelectedIndexChanged
        If Me.cboPayMode.Text.Trim.Length <= 0 Then
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()

            Me.cboAccountNo.Items.Clear()
            Me.cboAccountNo.Text = ""
            Me.cboAccountNo.SelectedIndex = -1

            Me.cmdPayRequest.Connection = conn
            Me.cmdPayRequest.CommandType = CommandType.Text
            Me.cmdPayRequest.CommandText = "SELECT DISTINCT accountNumber FROM tblFinPayAccounts WHERE (modeName=@modeName) ORDER BY accountNumber"
            Me.cmdPayRequest.Parameters.Clear()
            Me.cmdPayRequest.Parameters.AddWithValue("@modeName", Me.cboPayMode.Text.Trim)
            reader = Me.cmdPayRequest.ExecuteReader
            While reader.Read
                Me.cboAccountNo.Items.Add(IIf(DBNull.Value.Equals(reader!accountNumber), "", reader!accountNumber))
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

    Private Sub cboAccountNo_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboAccountNo.SelectedIndexChanged
        If Me.cboPayMode.Text.Trim.Length <= 0 Or Me.cboAccountNo.Text.Trim.Length <= 0 Then
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()

            Me.txtBalance.Text = ""
            Me.txtVendorName.Text = ""

            Me.cmdPayRequest.Connection = conn
            Me.cmdPayRequest.CommandType = CommandType.Text
            Me.cmdPayRequest.CommandText = "SELECT DISTINCT balance,vendorName FROM tblFinPayAccounts WHERE (modeName=@modeName) " &
                vbNewLine & " AND (accountNumber=@accountNumber)"
            Me.cmdPayRequest.Parameters.Clear()
            Me.cmdPayRequest.Parameters.AddWithValue("@modeName", Me.cboPayMode.Text.Trim)
            Me.cmdPayRequest.Parameters.AddWithValue("@accountNumber", Me.cboAccountNo.Text.Trim)
            reader = Me.cmdPayRequest.ExecuteReader
            While reader.Read
                Me.txtBalance.Text = IIf(DBNull.Value.Equals(reader!balance), "", reader!balance)
                Me.txtBalance.Tag = IIf(DBNull.Value.Equals(reader!balance), "", reader!balance)
                Me.txtVendorName.Text = IIf(DBNull.Value.Equals(reader!vendorName), "", reader!vendorName)
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

    Private Sub cboExpCat_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboExpCat.SelectedIndexChanged
        If Me.cboExpCat.Text.Trim.Length <= 0 Then
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()

            Me.cboExpName.Items.Clear()
            Me.cboExpName.Text = ""
            Me.cboExpName.SelectedIndex = -1

            Me.cmdPayRequest.Connection = conn
            Me.cmdPayRequest.CommandType = CommandType.Text
            Me.cmdPayRequest.CommandText = "SELECT DISTINCT expName FROM vwFinExpName WHERE (expCategoryName=@expCategoryName) " &
                vbNewLine & " ORDER BY expName"
            Me.cmdPayRequest.Parameters.Clear()
            Me.cmdPayRequest.Parameters.AddWithValue("@expCategoryName", Me.cboExpCat.Text.Trim)
            reader = Me.cmdPayRequest.ExecuteReader
            While reader.Read
                Me.cboExpName.Items.Add(IIf(DBNull.Value.Equals(reader!expName), "", reader!expName))
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

    Private Sub cboExpName_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboExpName.SelectedIndexChanged
        If Me.cboPayTerm.Text.Trim.Length > 0 And Me.cboPayYear.Text.Trim.Length > 0 And Me.cboExpName.Text.Trim.Length > 0 Then
            Try
                If conn.State = ConnectionState.Closed Then
                    conn.Open()
                End If
                dbconnection()

                Me.cboMoneyPaidTo.Items.Clear()
                Me.cboMoneyPaidTo.Text = ""
                Me.cboMoneyPaidTo.SelectedIndex = -1

                Me.cmdPayRequest.Connection = conn
                Me.cmdPayRequest.CommandType = CommandType.Text
                Me.cmdPayRequest.CommandText = "SELECT DISTINCT paidTo FROM vwFinPayments WHERE (expName=@expName) " &
                    vbNewLine & " AND (termName=@termName) AND (year=@year) ORDER BY paidTo"
                Me.cmdPayRequest.Parameters.Clear()
                Me.cmdPayRequest.Parameters.AddWithValue("@termName", Me.cboPayTerm.Text.Trim)
                Me.cmdPayRequest.Parameters.AddWithValue("@year", Me.cboPayYear.Text.Trim)
                Me.cmdPayRequest.Parameters.AddWithValue("@expName", Me.cboExpName.Text.Trim)
                reader = Me.cmdPayRequest.ExecuteReader
                While reader.Read
                    Me.cboMoneyPaidTo.Items.Add(IIf(DBNull.Value.Equals(reader!paidTo), "", reader!paidTo))
                End While
                reader.Close()
            Catch ex As Exception
                MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Finally
                If conn.State = ConnectionState.Open Then
                    conn.Close()
                End If
            End Try
        End If
    End Sub

    Private Sub txtAmount_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtAmount.TextChanged
        If IsNumeric(Me.txtAmount.Text.Trim) = False And Not (Me.txtAmount.Text.Trim.Length <= 0) Then
            MsgBox("Non Numeric Values Detected.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Me.txtAmount.Text = ""
            Exit Sub
        End If
        If Me.txtBalance.Text.Trim.Length <= 0 Or Me.txtAmount.Text.Trim.Length <= 0 Then
            Exit Sub
        End If
        If Not (Me.txtBalance.Text.Trim.Length <= 0) And Not (Me.txtBalance.Text.Trim = "") Then
            Dim bal As Double = 0
            bal = Me.txtBalance.Text.Trim - Me.txtAmount.Text.Trim
            If bal < 0 Then
                MsgBox("Amount Exceeds what is available.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
                Me.txtAmount.Text = ""
            End If
        End If
    End Sub

    Private Sub btnAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAdd.Click
        If Me.cboPayTerm.Text.Trim.Length <= 0 Then
            MsgBox("Payment Term Is Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboPayYear.Text.Trim.Length <= 0 Then
            MsgBox("Payment Year Is Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboAccountNo.Text.Trim.Length <= 0 Then
            MsgBox("Account Number Is Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.txtVoucherNo.Text.Trim.Length <= 0 Then
            MsgBox("Voucher Number Is Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.txtBalance.Text.Trim.Length <= 0 Or Me.txtBalance.Text.Trim = 0 Then
            MsgBox("Balance Is Missing Or Zero.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboExpCat.Text.Trim.Length <= 0 Then
            MsgBox("Expense Category Is Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboExpName.Text.Trim.Length <= 0 Then
            MsgBox("Account Name Is Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboMoneyPaidTo.Text.Trim.Length <= 0 Then
            MsgBox("Money Paid To Is Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.txtAmount.Text.Trim.Length <= 0 Or Me.txtAmount.Text.Trim = "0" Then
            MsgBox("Amount Is Missing Or Zero.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If

        'j = 0
        'For j = 0 To Me.lstPayRequest.Items.Count - 1
        '    If Me.cboMoneyPaidTo.Text.Trim <> Me.lstPayRequest.Items(j).SubItems(5).Text.Trim Then
        '        MsgBox("One voucher Number For One Individual.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        '        Exit Sub
        '    End If
        'Next
        Dim amountToPayFromAccount As Double = 0
        For i = 0 To Me.lstPayRequest.Items.Count - 1
            If Me.cboAccountNo.Text.Trim = Me.lstPayRequest.Items(i).SubItems(2).Text.Trim Then
                amountToPayFromAccount = amountToPayFromAccount + Me.lstPayRequest.Items(i).SubItems(6).Text.Trim
            End If
        Next
        amountToPayFromAccount = amountToPayFromAccount + Me.txtAmount.Text.Trim
        If amountToPayFromAccount > Me.txtBalance.Text.Trim Then
            MsgBox("Cannot Add The Expense." & vbNewLine & "Amount will exceed the account balance",
                   MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If


        li = Me.lstPayRequest.Items.Add(Me.txtVoucherNo.Text.Trim)
        li.SubItems.Add(Me.cboPayMode.Text.Trim)
        li.SubItems.Add(Me.cboAccountNo.Text.Trim)
        li.SubItems.Add(Me.dtpPayDate.Value.Date)
        li.SubItems.Add(Me.cboExpName.Text.Trim)
        li.SubItems.Add(Me.cboMoneyPaidTo.Text.Trim)
        li.SubItems.Add(Me.txtAmount.Text.Trim)
        li.Tag = Me.txtVoucherNo.Tag

        Me.cboExpCat.Text = ""
        Me.cboExpCat.SelectedIndex = -1

        Me.cboExpName.Text = ""
        Me.cboExpName.SelectedIndex = -1

        Me.cboMoneyPaidTo.Text = ""
        Me.cboMoneyPaidTo.SelectedIndex = -1

        Me.txtAmount.Text = ""
    End Sub

    Private Sub btnUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdate.Click
        If Me.lstPayRequest.Items.Count <= 0 Then
            MsgBox("Missing Details in the List.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            Dim result As MsgBoxResult = MsgBox("Update Record/s?", MsgBoxStyle.Question + MsgBoxStyle.ApplicationModal + MsgBoxStyle.YesNo, "Confirm Transaction")
            If result = MsgBoxResult.No Then
                Exit Sub
            End If


            For i = 0 To Me.lstPayRequest.Items.Count - 1
                Me.cmdPayRequest.Connection = conn
                Me.cmdPayRequest.CommandType = CommandType.StoredProcedure
                Me.cmdPayRequest.CommandText = "sprocFinPayments"
                Me.cmdPayRequest.Parameters.Clear()
                Me.cmdPayRequest.Parameters.AddWithValue("@termName", Me.cboPayTerm.Text.Trim)
                Me.cmdPayRequest.Parameters.AddWithValue("@year", Me.cboPayYear.Text.Trim)
                Me.cmdPayRequest.Parameters.AddWithValue("@voucherNo", Me.lstPayRequest.Items(i).Tag)
                Me.cmdPayRequest.Parameters.AddWithValue("@VoucherNoaLL", Me.lstPayRequest.Items(i).Text.Trim)
                Me.cmdPayRequest.Parameters.AddWithValue("@payMode", Me.lstPayRequest.Items(i).SubItems(1).Text.Trim)
                Me.cmdPayRequest.Parameters.AddWithValue("@accountNumber", Me.lstPayRequest.Items(i).SubItems(2).Text.Trim)
                Me.cmdPayRequest.Parameters.AddWithValue("@expenseName", Me.lstPayRequest.Items(i).SubItems(4).Text.Trim)
                Me.cmdPayRequest.Parameters.AddWithValue("@paidTo", Me.lstPayRequest.Items(i).SubItems(5).Text.Trim)
                Me.cmdPayRequest.Parameters.AddWithValue("@payAmount", Me.lstPayRequest.Items(i).SubItems(6).Text.Trim)
                Me.cmdPayRequest.Parameters.AddWithValue("@userName", fullName.Trim)
                Me.cmdPayRequest.Parameters.AddWithValue("@requestDate", Me.lstPayRequest.Items(i).SubItems(3).Text.Trim)
                rec = rec + Me.cmdPayRequest.ExecuteNonQuery
            Next
            If rec > 0 Then
                MsgBox("Payment Requests Successfully Entered.", MsgBoxStyle.Information +
                       MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "System Reply")
            End If

            Me.lstPayRequest.Items.Clear()
            loadCombos()

            Me.txtVoucherNo.Text = ""
            Me.txtVoucherNo.Tag = Nothing

            Me.cboAccountNo.Items.Clear()
            Me.cboAccountNo.Text = ""
            Me.cboAccountNo.SelectedIndex = -1

            Me.txtBalance.Text = ""
            Me.txtVendorName.Text = ""

            Me.txtAmount.Text = ""

            Me.cboExpName.Items.Clear()
            Me.cboExpName.Text = ""
            Me.cboExpName.SelectedIndex = -1

            Me.cboMoneyPaidTo.Items.Clear()
            Me.cboMoneyPaidTo.Text = ""
            Me.cboMoneyPaidTo.SelectedIndex = -1

        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub
End Class