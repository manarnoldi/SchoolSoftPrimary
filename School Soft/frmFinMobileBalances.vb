Imports System.Data.SqlClient
Public Class frmFinMobileBalances
    Dim cmdMobileBal As New SqlCommand
    Dim reader As SqlDataReader
    Dim rec As Integer = 0
    Private Sub frmFinMobileBalances_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
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
    Private Sub loadList()
        Me.lstMobileBalances.Items.Clear()

        Me.cmdMobileBal.Connection = conn
        Me.cmdMobileBal.CommandType = CommandType.Text
        Me.cmdMobileBal.CommandText = "SELECT * FROM  tblFinPayAccounts WHERE (modeName='MOBILE') ORDER BY vendorName"
        Me.cmdMobileBal.Parameters.Clear()
        reader = Me.cmdMobileBal.ExecuteReader
        While reader.Read
            li = Me.lstMobileBalances.Items.Add(IIf(DBNull.Value.Equals(reader!vendorName), "", reader!vendorName))
            li.SubItems.Add(IIf(DBNull.Value.Equals(reader!accountNumber), "", reader!accountNumber))
            li.SubItems.Add(IIf(DBNull.Value.Equals(reader!balance), "", reader!balance))
        End While
        reader.Close()
    End Sub
    Private Sub loadPayModes()
        Me.cboTransferTo.Items.Clear()
        Me.cboTransferTo.Text = ""
        Me.cboTransferTo.SelectedIndex = -1

        Me.cmdMobileBal.Connection = conn
        Me.cmdMobileBal.CommandType = CommandType.Text
        Me.cmdMobileBal.CommandText = "SELECT DISTINCT modeName FROM tblFinPayModes " & _
            vbNewLine & " WHERE (modeName<>'MOBILE') ORDER BY modeName"
        Me.cmdMobileBal.Parameters.Clear()
        reader = Me.cmdMobileBal.ExecuteReader
        While reader.Read
            Me.cboTransferTo.Items.Add(IIf(DBNull.Value.Equals(reader!modeName), "", reader!modeName))
        End While
        reader.Close()
    End Sub
    Private Sub loadCombos()
        Me.cboBankName.Items.Clear()
        Me.cboBankName.Text = ""
        Me.cboBankName.SelectedIndex = -1

        Me.cmdMobileBal.Connection = conn
        Me.cmdMobileBal.CommandType = CommandType.Text
        Me.cmdMobileBal.CommandText = "SELECT DISTINCT vendorName FROM tblFinPayAccounts " & _
            vbNewLine & " WHERE (modeName='MOBILE') ORDER BY vendorName"
        Me.cmdMobileBal.Parameters.Clear()
        reader = Me.cmdMobileBal.ExecuteReader
        While reader.Read
            Me.cboVendorName.Items.Add(IIf(DBNull.Value.Equals(reader!vendorName), "", reader!vendorName))
        End While
        reader.Close()
    End Sub
    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub frmFinMobileBalances_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint
        If Me.IsMdiChild Then
            Dim PnlLoc As New Point
            PnlLoc.X = CInt((Me.Width - Me.pnlMobileBalances.Width) / 2)
            PnlLoc.Y = CInt((Me.Height - Me.pnlMobileBalances.Height) / 2.5)
            Me.pnlMobileBalances.Location = PnlLoc
        Else
            Me.pnlMobileBalances.Dock = DockStyle.Fill
        End If
    End Sub

    Private Sub cboVendorName_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboVendorName.SelectedIndexChanged
        If Me.cboVendorName.Text.Trim.Length <= 0 Then
            MsgBox("Vendor Name Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            Me.cboPhoneNo.Items.Clear()
            Me.cboPhoneNo.Text = ""
            Me.cboPhoneNo.SelectedIndex = -1

            Me.txtBalance.Text = ""

            Me.cmdMobileBal.Connection = conn
            Me.cmdMobileBal.CommandType = CommandType.Text
            Me.cmdMobileBal.CommandText = "SELECT DISTINCT accountNumber FROM tblFinPayAccounts  WHERE (modeName='MOBILE') " & _
                vbNewLine & " AND (vendorName=@vendorName)  ORDER BY accountNumber"
            Me.cmdMobileBal.Parameters.Clear()
            Me.cmdMobileBal.Parameters.AddWithValue("@vendorName", Me.cboVendorName.Text.Trim)
            reader = Me.cmdMobileBal.ExecuteReader
            While reader.Read
                Me.cboPhoneNo.Items.Add(IIf(DBNull.Value.Equals(reader!accountNumber), "", reader!accountNumber))
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

    Private Sub cboPhoneNo_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboPhoneNo.SelectedIndexChanged
        If Me.cboPhoneNo.Text.Trim.Length <= 0 Or Me.cboVendorName.Text.Trim.Length <= 0 Then
            MsgBox("Phone Number or Vendor Name Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            Me.txtBalance.Text = ""
            Me.txtBalance.Tag = Nothing

            Me.cmdMobileBal.Connection = conn
            Me.cmdMobileBal.CommandType = CommandType.Text
            Me.cmdMobileBal.CommandText = "SELECT DISTINCT accountId,balance FROM tblFinPayAccounts  WHERE (modeName='MOBILE') " & _
                vbNewLine & " AND (vendorName=@vendorName) AND (accountNumber=@accountNumber)"
            Me.cmdMobileBal.Parameters.Clear()
            Me.cmdMobileBal.Parameters.AddWithValue("@vendorName", Me.cboVendorName.Text.Trim)
            Me.cmdMobileBal.Parameters.AddWithValue("@accountNumber", Me.cboPhoneNo.Text.Trim)
            reader = Me.cmdMobileBal.ExecuteReader
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
            If Me.cboTransferTo.Text = "BANK" Then
                Me.grpBank.Enabled = True

                Me.cboBankName.Items.Clear()
                Me.cboBankName.Text = ""
                Me.cboBankName.SelectedIndex = -1

                Me.cboAccountNo.Items.Clear()
                Me.cboAccountNo.Text = ""
                Me.cboAccountNo.SelectedIndex = -1

                Me.cmdMobileBal.Connection = conn
                Me.cmdMobileBal.CommandType = CommandType.Text
                Me.cmdMobileBal.CommandText = "SELECT DISTINCT vendorName FROM tblFinPayAccounts  WHERE " & _
                    vbNewLine & " (modeName='BANK') ORDER BY vendorName"
                Me.cmdMobileBal.Parameters.Clear()
                reader = Me.cmdMobileBal.ExecuteReader
                While reader.Read
                    Me.cboBankName.Items.Add(IIf(DBNull.Value.Equals(reader!vendorName), "", reader!vendorName))
                End While
                reader.Close()
            ElseIf Me.cboTransferTo.Text = "CASH" Then
                Me.cboBankName.Items.Clear()
                Me.cboBankName.Text = ""
                Me.cboBankName.SelectedIndex = -1

                Me.cboAccountNo.Items.Clear()
                Me.cboAccountNo.Text = ""
                Me.cboAccountNo.SelectedIndex = -1

                Me.grpBank.Enabled = False

                Me.cmdMobileBal.Connection = conn
                Me.cmdMobileBal.CommandType = CommandType.Text
                Me.cmdMobileBal.CommandText = "SELECT accountId FROM tblFinPayAccounts  WHERE " & _
                    vbNewLine & " (modeName='CASH')"
                Me.cmdMobileBal.Parameters.Clear()
                reader = Me.cmdMobileBal.ExecuteReader
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
            Me.cboAccountNo.Items.Clear()
            Me.cboAccountNo.Text = ""
            Me.cboAccountNo.SelectedIndex = -1

            Me.cmdMobileBal.Connection = conn
            Me.cmdMobileBal.CommandType = CommandType.Text
            Me.cmdMobileBal.CommandText = "SELECT DISTINCT accountNumber FROM tblFinPayAccounts " & _
                vbNewLine & " WHERE (modeName='BANK') AND (vendorName=@vendorName) ORDER BY accountNumber"
            Me.cmdMobileBal.Parameters.Clear()
            Me.cmdMobileBal.Parameters.AddWithValue("@vendorName", Me.cboBankName.Text.Trim)
            reader = Me.cmdMobileBal.ExecuteReader
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
        If Me.cboBankName.Text.Trim.Length <= 0 Or Me.cboAccountNo.Text.Trim.Length <= 0 Or Me.cboAccountNo.Text.Trim.Length <= 0 Then
            MsgBox("Bank Name or Account Number Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()

            Me.txtAmount.Tag = Nothing

            Me.cmdMobileBal.Connection = conn
            Me.cmdMobileBal.CommandType = CommandType.Text
            Me.cmdMobileBal.CommandText = "SELECT DISTINCT accountId FROM tblFinPayAccounts  WHERE (modeName='BANK') " & _
                vbNewLine & " AND (vendorName=@vendorName) AND (accountNumber=@accountNumber)"
            Me.cmdMobileBal.Parameters.Clear()
            Me.cmdMobileBal.Parameters.AddWithValue("@vendorName", Me.cboBankName.Text.Trim)
            Me.cmdMobileBal.Parameters.AddWithValue("@accountNumber", Me.cboAccountNo.Text.Trim)
            reader = Me.cmdMobileBal.ExecuteReader
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
        If Me.cboVendorName.Text.Trim.Length <= 0 Then
            MsgBox("Vendor Name Is Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboTransferTo.Text.Trim.Length <= 0 Then
            MsgBox("Transfer To Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboTransferTo.Text = "BANK" Then
            If Me.cboBankName.Text.Trim.Length <= 0 Then
                MsgBox("Bank Name Is Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
                Exit Sub
            ElseIf Me.cboAccountNo.Text.Trim.Length <= 0 Then
                MsgBox("Account Number Is Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
                Exit Sub
            End If
        ElseIf Me.txtAmount.Tag = Nothing Then
            MsgBox("Important Data Missing." & vbNewLine & "Reselect Your Entries.", MsgBoxStyle.Exclamation + _
                   MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.txtBalance.Text.Trim - Me.txtAmount.Text.Trim < 0 Then
            MsgBox("You Cannot Transfer More Than What Is Available.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboPhoneNo.Text.Trim.Length <= 0 Then
            MsgBox("Phone Number Is Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
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
            Me.cmdMobileBal.Connection = conn
            Me.cmdMobileBal.CommandType = CommandType.StoredProcedure
            Me.cmdMobileBal.CommandText = "sprocFinTransferBalances"
            Me.cmdMobileBal.Parameters.Clear()
            Me.cmdMobileBal.Parameters.AddWithValue("@accountId", Me.txtBalance.Tag)
            Me.cmdMobileBal.Parameters.AddWithValue("@toAccountId", Me.txtAmount.Tag)
            Me.cmdMobileBal.Parameters.AddWithValue("@amount", Me.txtAmount.Text.Trim)
            Me.cmdMobileBal.Parameters.AddWithValue("@doneBy", userName.Trim)
            rec = Me.cmdMobileBal.ExecuteNonQuery
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

        Me.cboBankName.Items.Clear()
        Me.cboBankName.Text = ""
        Me.cboBankName.SelectedIndex = -1

        Me.cboVendorName.Items.Clear()
        Me.cboVendorName.Text = ""
        Me.cboVendorName.SelectedIndex = -1

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
End Class