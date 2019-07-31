Imports System.Data.SqlClient
Public Class frmFinBankBalances
    Dim cmdBankBalances As New SqlCommand
    Dim reader As SqlDataReader
    Dim rec As Integer = 0
    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub
    Private Sub loadList()
        Me.lstBankBalances.Items.Clear()

        Me.cmdBankBalances.Connection = conn
        Me.cmdBankBalances.CommandType = CommandType.Text
        Me.cmdBankBalances.CommandText = "SELECT * FROM  tblFinPayAccounts WHERE (modeName='BANK') ORDER BY vendorName"
        Me.cmdBankBalances.Parameters.Clear()
        reader = Me.cmdBankBalances.ExecuteReader
        While reader.Read
            li = Me.lstBankBalances.Items.Add(IIf(DBNull.Value.Equals(reader!vendorName), "", reader!vendorName))
            li.SubItems.Add(IIf(DBNull.Value.Equals(reader!accountName), "", reader!accountName))
            li.SubItems.Add(IIf(DBNull.Value.Equals(reader!accountNumber), "", reader!accountNumber))
            li.SubItems.Add(IIf(DBNull.Value.Equals(reader!balance), "", reader!balance))
        End While
        reader.Close()
    End Sub
    Private Sub frmFinBankBalances_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            loadCombos()
            loadList()
            loadPayModes()
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub
    Private Sub loadPayModes()
        Me.cboTransferTo.Items.Clear()
        Me.cboTransferTo.Text = ""
        Me.cboTransferTo.SelectedIndex = -1

        Me.cmdBankBalances.Connection = conn
        Me.cmdBankBalances.CommandType = CommandType.Text
        Me.cmdBankBalances.CommandText = "SELECT DISTINCT modeName FROM tblFinPayModes " & _
            vbNewLine & " WHERE (modeName<>'BANK') ORDER BY modeName"
        Me.cmdBankBalances.Parameters.Clear()
        reader = Me.cmdBankBalances.ExecuteReader
        While reader.Read
            Me.cboTransferTo.Items.Add(IIf(DBNull.Value.Equals(reader!modeName), "", reader!modeName))
        End While
        reader.Close()
    End Sub
    Private Sub loadCombos()
        Me.cboBankName.Items.Clear()
        Me.cboBankName.Text = ""
        Me.cboBankName.SelectedIndex = -1

        Me.cmdBankBalances.Connection = conn
        Me.cmdBankBalances.CommandType = CommandType.Text
        Me.cmdBankBalances.CommandText = "SELECT DISTINCT vendorName FROM tblFinPayAccounts " & _
            vbNewLine & " WHERE (modeName='BANK') ORDER BY vendorName"
        Me.cmdBankBalances.Parameters.Clear()
        reader = Me.cmdBankBalances.ExecuteReader
        While reader.Read
            Me.cboBankName.Items.Add(IIf(DBNull.Value.Equals(reader!vendorName), "", reader!vendorName))
        End While
        reader.Close()
    End Sub
    Private Sub frmFinBankBalances_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint
        If Me.IsMdiChild Then
            Dim PnlLoc As New Point
            PnlLoc.X = CInt((Me.Width - Me.pnlBankBalances.Width) / 2)
            PnlLoc.Y = CInt((Me.Height - Me.pnlBankBalances.Height) / 2.5)
            Me.pnlBankBalances.Location = PnlLoc
        Else
            Me.pnlBankBalances.Dock = DockStyle.Fill
        End If
    End Sub

    Private Sub cboBankName_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboBankName.SelectedIndexChanged
        If Me.cboBankName.Text.Trim.Length <= 0 Then
            MsgBox("Bank Name Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            Me.cboAccountName.Items.Clear()
            Me.cboAccountName.Text = ""
            Me.cboAccountName.SelectedIndex = -1

            Me.cmdBankBalances.Connection = conn
            Me.cmdBankBalances.CommandType = CommandType.Text
            Me.cmdBankBalances.CommandText = "SELECT DISTINCT accountName FROM tblFinPayAccounts " & _
                vbNewLine & " WHERE (modeName='Bank') AND (vendorName=@vendorName) ORDER BY accountName"
            Me.cmdBankBalances.Parameters.Clear()
            Me.cmdBankBalances.Parameters.AddWithValue("@vendorName", Me.cboBankName.Text.Trim)
            reader = Me.cmdBankBalances.ExecuteReader
            While reader.Read
                Me.cboAccountName.Items.Add(IIf(DBNull.Value.Equals(reader!accountName), "", reader!accountName))
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

    Private Sub cboAccountName_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboAccountName.SelectedIndexChanged
        If Me.cboBankName.Text.Trim.Length <= 0 Or Me.cboAccountName.Text.Trim.Length <= 0 Then
            MsgBox("Bank Name or Account Name Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
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

            Me.cmdBankBalances.Connection = conn
            Me.cmdBankBalances.CommandType = CommandType.Text
            Me.cmdBankBalances.CommandText = "SELECT DISTINCT accountNumber FROM tblFinPayAccounts  WHERE (modeName='BANK') " & _
                vbNewLine & " AND (vendorName=@vendorName) AND (accountName=@accountName) ORDER BY accountNumber"
            Me.cmdBankBalances.Parameters.Clear()
            Me.cmdBankBalances.Parameters.AddWithValue("@vendorName", Me.cboBankName.Text.Trim)
            Me.cmdBankBalances.Parameters.AddWithValue("@accountName", Me.cboAccountName.Text.Trim)
            reader = Me.cmdBankBalances.ExecuteReader
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
        If Me.cboBankName.Text.Trim.Length <= 0 Or Me.cboAccountName.Text.Trim.Length <= 0 Or Me.cboAccountNo.Text.Trim.Length <= 0 Then
            MsgBox("Bank Name or Account Name or Account Number Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            Me.txtBalance.Text = ""
            Me.txtBalance.Tag = Nothing

            Me.cmdBankBalances.Connection = conn
            Me.cmdBankBalances.CommandType = CommandType.Text
            Me.cmdBankBalances.CommandText = "SELECT DISTINCT accountId,balance FROM tblFinPayAccounts  WHERE (modeName='BANK') " & _
                vbNewLine & " AND (vendorName=@vendorName) AND (accountName=@accountName) AND (accountNumber=@accountNumber)"
            Me.cmdBankBalances.Parameters.Clear()
            Me.cmdBankBalances.Parameters.AddWithValue("@vendorName", Me.cboBankName.Text.Trim)
            Me.cmdBankBalances.Parameters.AddWithValue("@accountName", Me.cboAccountName.Text.Trim)
            Me.cmdBankBalances.Parameters.AddWithValue("@accountNumber", Me.cboAccountNo.Text.Trim)
            reader = Me.cmdBankBalances.ExecuteReader
            While reader.Read
                Me.txtBalance.Text = (IIf(DBNull.Value.Equals(reader!balance), "", reader!balance))
                Me.txtBalance.Tag = (IIf(DBNull.Value.Equals(reader!accountId), "", reader!accountId))
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
    Private Sub cboTransferTo_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboTransferTo.SelectedIndexChanged
        If Me.cboTransferTo.Text.Trim.Length <= 0 Then
            MsgBox("Transfer To Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            Me.txtAmount.Tag = Nothing
            If Me.cboTransferTo.Text = "MOBILE" Then
                Me.cboPhoneNo.Enabled = True

                Me.cboPhoneNo.Items.Clear()
                Me.cboPhoneNo.Text = ""
                Me.cboPhoneNo.SelectedIndex = -1

                Me.cmdBankBalances.Connection = conn
                Me.cmdBankBalances.CommandType = CommandType.Text
                Me.cmdBankBalances.CommandText = "SELECT DISTINCT accountNumber FROM tblFinPayAccounts  WHERE " & _
                    vbNewLine & " (modeName='MOBILE') ORDER BY accountNumber"
                Me.cmdBankBalances.Parameters.Clear()
                reader = Me.cmdBankBalances.ExecuteReader
                While reader.Read
                    Me.cboPhoneNo.Items.Add(IIf(DBNull.Value.Equals(reader!accountNumber), "", reader!accountNumber))
                End While
                reader.Close()
            ElseIf Me.cboTransferTo.Text = "CASH" Then
                Me.cboPhoneNo.Items.Clear()
                Me.cboPhoneNo.Text = ""
                Me.cboPhoneNo.SelectedIndex = -1
                Me.cboPhoneNo.Enabled = False

                Me.cmdBankBalances.Connection = conn
                Me.cmdBankBalances.CommandType = CommandType.Text
                Me.cmdBankBalances.CommandText = "SELECT accountId FROM tblFinPayAccounts  WHERE " & _
                    vbNewLine & " (modeName='CASH')"
                Me.cmdBankBalances.Parameters.Clear()
                reader = Me.cmdBankBalances.ExecuteReader
                While reader.Read
                    Me.txtAmount.Tag = (IIf(DBNull.Value.Equals(reader!accountId), "", reader!accountId))
                End While
                reader.Close()
            End If
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub

    Private Sub txtAmount_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtAmount.TextChanged
        If IsNumeric(Me.txtAmount.Text.Trim) = False And Not (Me.txtAmount.Text.Trim.Length <= 0) Then
            MsgBox("Non Numeric Values Detected.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Me.txtAmount.Text = ""
            Exit Sub
        End If
        If Me.txtAmount.Text.Trim.Length <= 0 Then
            Exit Sub
        End If
        If Me.txtBalance.Text.Trim.Length <= 0 And Me.txtBalance.Text.Trim = "" Then
            MsgBox("Balance Is Not Available.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Me.txtAmount.text=""
            Exit Sub
        End If
        If Me.txtBalance.Text.Trim - Me.txtAmount.Text.Trim < 0 Then
            MsgBox("You Cannot Transfer More Than What Is Available.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Me.txtAmount.Text = ""
            Exit Sub
        End If
        'If Me.txtBalance.Text.Trim.Length > 0 And Me.txtBalance.Text <> "0" Then
        '    Me.txtBalance.Text = Me.txtBalance.Text.Trim - Me.txtAmount.Text.Trim
        'End If
    End Sub

    Private Sub btnUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdate.Click
        If Me.cboBankName.Text.Trim.Length <= 0 Then
            MsgBox("Bank Name Is Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboTransferTo.Text.Trim.Length <= 0 Then
            MsgBox("Transfer To Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboTransferTo.Text = "MOBILE" Then
            If Me.cboPhoneNo.Text.Trim.Length <= 0 Then
                MsgBox("Transfer To Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
                Exit Sub
            End If
        ElseIf Me.txtAmount.Tag = Nothing Then
            MsgBox("Important Data Missing." & vbNewLine & "Reselect Your Entries.", MsgBoxStyle.Exclamation + _
                   MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.txtBalance.Text.Trim - Me.txtAmount.Text.Trim < 0 Then
            MsgBox("You Cannot Transfer More Than What Is Available.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            Dim result As MsgBoxResult = MsgBox("Transfer Funds?", MsgBoxStyle.Question + MsgBoxStyle.ApplicationModal + MsgBoxStyle.YesNo, "Confirm Transaction")
            If result = MsgBoxResult.No Then
                Exit Sub
            End If
            Me.cmdBankBalances.Connection = conn
            Me.cmdBankBalances.CommandType = CommandType.StoredProcedure
            Me.cmdBankBalances.CommandText = "sprocFinTransferBalances"
            Me.cmdBankBalances.Parameters.Clear()
            Me.cmdBankBalances.Parameters.AddWithValue("@accountId", Me.txtBalance.Tag)
            Me.cmdBankBalances.Parameters.AddWithValue("@toAccountId", Me.txtAmount.Tag)
            Me.cmdBankBalances.Parameters.AddWithValue("@amount", Me.txtAmount.Text.Trim)
            Me.cmdBankBalances.Parameters.AddWithValue("@doneBy", userName.Trim)
            rec = Me.cmdBankBalances.ExecuteNonQuery
            If rec > 0 Then
                MsgBox("Funds Transfered Successfully.", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "System Reply")
            End If
            clearTexts()
            loadCombos()
            loadList()
            loadPayModes()
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub
    Private Sub clearTexts()
        Me.txtAmount.Text = ""
        Me.txtAmount.Tag = Nothing
        Me.txtBalance.Text = ""
        Me.txtBalance.Tag = Nothing

        Me.cboAccountName.Items.Clear()
        Me.cboAccountName.Text = ""
        Me.cboAccountName.SelectedIndex = -1

        Me.cboAccountNo.Items.Clear()
        Me.cboAccountNo.Text = ""
        Me.cboAccountNo.SelectedIndex = -1

        Me.cboTransferTo.Items.Clear()
        Me.cboTransferTo.Text = ""
        Me.cboTransferTo.SelectedIndex = -1

        Me.cboPhoneNo.Items.Clear()
        Me.cboPhoneNo.Text = ""
        Me.cboPhoneNo.SelectedIndex = -1
    End Sub
    Private Sub cboPhoneNo_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboPhoneNo.SelectedIndexChanged
        If Me.cboPhoneNo.Text.Trim.Length <= 0 Then
            MsgBox("Phone Number Is Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            Me.cboPhoneNo.Tag = Nothing

            Me.cmdBankBalances.Connection = conn
            Me.cmdBankBalances.CommandType = CommandType.Text
            Me.cmdBankBalances.CommandText = "SELECT DISTINCT accountId,accountNumber FROM tblFinPayAccounts  WHERE " & _
            vbNewLine & " (modeName='MOBILE') AND (accountNumber=@accountNumber) ORDER BY accountNumber"
            Me.cmdBankBalances.Parameters.Clear()
            Me.cmdBankBalances.Parameters.AddWithValue("@accountNumber", Me.cboPhoneNo.Text.Trim)
            reader = Me.cmdBankBalances.ExecuteReader
            While reader.Read
                Me.txtAmount.Tag = (IIf(DBNull.Value.Equals(reader!accountId), "", reader!accountId))
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
End Class