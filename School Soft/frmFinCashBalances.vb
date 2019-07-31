Imports System.Data.SqlClient
Public Class frmFinCashBalances
    Dim cmdCashBalances As New SqlCommand
    Dim reader As SqlDataReader
    Dim rec As Integer = 0
    Private Sub frmFinCashBalances_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
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
        Me.cboAccountName.Items.Clear()
        Me.cboAccountName.Text = ""
        Me.cboAccountName.SelectedIndex = -1

        Me.cmdCashBalances.Connection = conn
        Me.cmdCashBalances.CommandType = CommandType.Text
        Me.cmdCashBalances.CommandText = "SELECT DISTINCT modeName FROM tblFinPayModes WHERE (modeName<>'CASH') ORDER BY modeName"
        Me.cmdCashBalances.Parameters.Clear()
        reader = Me.cmdCashBalances.ExecuteReader
        While reader.Read
            Me.cboTransferTo.Items.Add(IIf(DBNull.Value.Equals(reader!modeName), "", reader!modeName))
        End While
        reader.Close()

        Me.cmdCashBalances.CommandText = "SELECT accountId,balance FROM  tblFinPayAccounts WHERE (modeName='CASH')"
        Me.cmdCashBalances.Parameters.Clear()
        reader = Me.cmdCashBalances.ExecuteReader
        While reader.Read
            Me.txtBalance.Text = (IIf(DBNull.Value.Equals(reader!balance), "", reader!balance))
            Me.txtBalance.Tag = (IIf(DBNull.Value.Equals(reader!accountId), "", reader!accountId))
        End While
        reader.Close()
    End Sub
    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub frmFinCashBalances_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint
        If Me.IsMdiChild Then
            Dim PnlLoc As New Point
            PnlLoc.X = CInt((Me.Width - Me.pnlCashBalances.Width) / 2)
            PnlLoc.Y = CInt((Me.Height - Me.pnlCashBalances.Height) / 2.5)
            Me.pnlCashBalances.Location = PnlLoc
        Else
            Me.pnlCashBalances.Dock = DockStyle.Fill
        End If
    End Sub

    Private Sub cboTrasferTo_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboTransferTo.SelectedIndexChanged
        If Me.cboTransferTo.Text.Trim.Length <= 0 Or Me.cboTransferTo.Text.Trim.Length <= 0 Then
            MsgBox("Transfer To Details Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
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

            Me.cboAccountNo.Items.Clear()
            Me.cboAccountNo.Text = ""
            Me.cboAccountNo.SelectedIndex = -1
            Me.txtVendorName.Text = ""
            Me.cmdCashBalances.Connection = conn
            Me.cmdCashBalances.CommandType = CommandType.Text
            If Me.cboTransferTo.Text.Trim = "BANK" Then
                Me.cmdCashBalances.CommandText = "SELECT DISTINCT accountName FROM tblFinPayAccounts  WHERE (modeName='BANK') " & _
                vbNewLine & "  ORDER BY accountName"
                Me.cmdCashBalances.Parameters.Clear()
            ElseIf Me.cboTransferTo.Text.Trim = "MOBILE" Then
                Me.cmdCashBalances.CommandText = "SELECT DISTINCT accountName FROM tblFinPayAccounts  WHERE (modeName='MOBILE') " & _
                vbNewLine & "  ORDER BY accountName"
                Me.cmdCashBalances.Parameters.Clear()
            End If
            reader = Me.cmdCashBalances.ExecuteReader
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
        If Me.cboAccountName.Text.Trim.Length <= 0 Then
            MsgBox("Account Name Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
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
            Me.txtVendorName.Text = ""
            Dim accountType As String = Nothing
            If Me.cboTransferTo.Text.Trim = "BANK" Then
                accountType = "BANK"
            ElseIf Me.cboTransferTo.Text.Trim = "MOBILE" Then
                accountType = "MOBILE"
            End If
            Me.cmdCashBalances.Connection = conn
            Me.cmdCashBalances.CommandType = CommandType.Text
            Me.cmdCashBalances.CommandText = "SELECT DISTINCT accountNumber FROM tblFinPayAccounts  WHERE (modeName=@modeName) " & _
                vbNewLine & " AND  (accountName=@accountName) ORDER BY accountNumber"
            Me.cmdCashBalances.Parameters.Clear()
            Me.cmdCashBalances.Parameters.AddWithValue("@modeName", accountType.Trim)
            Me.cmdCashBalances.Parameters.AddWithValue("@accountName", Me.cboAccountName.Text.Trim)
            reader = Me.cmdCashBalances.ExecuteReader
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
        If Me.cboAccountNo.Text.Trim.Length <= 0 Then
            MsgBox("Account Number Is Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            Me.txtVendorName.Text = ""

            Dim accountType As String = Nothing
            If Me.cboTransferTo.Text.Trim = "BANK" Then
                accountType = "BANK"
            ElseIf Me.cboTransferTo.Text.Trim = "MOBILE" Then
                accountType = "MOBILE"
            End If
            Me.cmdCashBalances.Connection = conn
            Me.cmdCashBalances.CommandType = CommandType.Text
            Me.cmdCashBalances.CommandText = "SELECT DISTINCT accountId,vendorName FROM tblFinPayAccounts  WHERE (modeName=@modeName) " & _
                vbNewLine & " AND  (accountName=@accountName) AND (accountNumber=@accountNumber)"
            Me.cmdCashBalances.Parameters.Clear()
            Me.cmdCashBalances.Parameters.AddWithValue("@modeName", accountType.Trim)
            Me.cmdCashBalances.Parameters.AddWithValue("@accountName", Me.cboAccountName.Text.Trim)
            Me.cmdCashBalances.Parameters.AddWithValue("@accountNumber", Me.cboAccountNo.Text.Trim)
            reader = Me.cmdCashBalances.ExecuteReader
            While reader.Read
                Me.txtVendorName.Text = (IIf(DBNull.Value.Equals(reader!vendorName), "", reader!vendorName))
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
            Me.txtAmount.Text = ""
            Exit Sub
        End If
        If Me.txtBalance.Text.Trim - Me.txtAmount.Text.Trim < 0 Then
            MsgBox("You Cannot Transfer More Than What Is Available.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Me.txtAmount.Text = ""
            Exit Sub
        End If
    End Sub

    Private Sub btnUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdate.Click
        If Me.txtVendorName.Text.Trim.Length <= 0 Then
            MsgBox("Vendor Name Is Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboAccountName.Text.Trim.Length <= 0 Then
            MsgBox("Account Name Is Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboTransferTo.Text.Trim.Length <= 0 Then
            MsgBox("Transfer To Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboAccountNo.Text.Trim.Length <= 0 Then
            MsgBox("Account Number is Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.txtBalance.Text.Trim = "" Or Me.txtAmount.Text.Trim.Length < 0 Then
            MsgBox("Balance Details Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
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
            Me.cmdCashBalances.Connection = conn
            Me.cmdCashBalances.CommandType = CommandType.StoredProcedure
            Me.cmdCashBalances.CommandText = "sprocFinTransferBalances"
            Me.cmdCashBalances.Parameters.Clear()
            Me.cmdCashBalances.Parameters.AddWithValue("@accountId", Me.txtBalance.Tag)
            Me.cmdCashBalances.Parameters.AddWithValue("@toAccountId", Me.txtAmount.Tag)
            Me.cmdCashBalances.Parameters.AddWithValue("@amount", Me.txtAmount.Text.Trim)
            Me.cmdCashBalances.Parameters.AddWithValue("@doneBy", userName.Trim)
            rec = Me.cmdCashBalances.ExecuteNonQuery
            If rec > 0 Then
                MsgBox("Funds Transfered Successfully.", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "System Reply")
            End If
            clearTexts()
            loadCombos()
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

        Me.txtVendorName.Text = ""

        Me.cboAccountName.Items.Clear()
        Me.cboAccountName.Text = ""
        Me.cboAccountName.SelectedIndex = -1

        Me.cboAccountNo.Items.Clear()
        Me.cboAccountNo.Text = ""
        Me.cboAccountNo.SelectedIndex = -1

        Me.cboTransferTo.Items.Clear()
        Me.cboTransferTo.Text = ""
        Me.cboTransferTo.SelectedIndex = -1
    End Sub
End Class