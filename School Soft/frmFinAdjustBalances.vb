Imports System.Data.SqlClient
Public Class frmFinAdjustBalances
    Dim cmdAdjustBal As New SqlCommand
    Dim reader As SqlDataReader
    Dim rec As Integer = 0
    Private Sub frmFinAdjustBalances_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            loadCombos()
            loadList()
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub
    Private Sub loadCombos()
        Me.cboAccType.Items.Clear()
        Me.cboAccType.Text = ""
        Me.cboAccType.SelectedIndex = -1

        Me.cboVendorName.Items.Clear()
        Me.cboVendorName.Text = ""
        Me.cboVendorName.SelectedIndex = -1

        Me.cboAccountNo.Items.Clear()
        Me.cboAccountNo.Text = ""
        Me.cboAccountNo.SelectedIndex = -1

        Me.txtBalance.Text = ""
        Me.txtBalance.Tag = Nothing

        Me.cmdAdjustBal.Connection = conn
        Me.cmdAdjustBal.CommandType = CommandType.Text
        Me.cmdAdjustBal.CommandText = "SELECT DISTINCT modeName FROM tblFinPayAccounts ORDER BY modeName"
        Me.cmdAdjustBal.Parameters.Clear()
        reader = Me.cmdAdjustBal.ExecuteReader
        While reader.Read
            Me.cboAccType.Items.Add(IIf(DBNull.Value.Equals(reader!modeName), "", reader!modeName))
        End While
        reader.Close()
    End Sub
    Private Sub loadList()
        Me.lstAccountBalances.Items.Clear()

        Me.cmdAdjustBal.Connection = conn
        Me.cmdAdjustBal.CommandType = CommandType.Text
        Me.cmdAdjustBal.CommandText = "SELECT * FROM  tblFinPayAccounts ORDER BY modeName"
        Me.cmdAdjustBal.Parameters.Clear()
        reader = Me.cmdAdjustBal.ExecuteReader
        While reader.Read
            li = Me.lstAccountBalances.Items.Add(IIf(DBNull.Value.Equals(reader!modeName), "", reader!modeName))
            li.SubItems.Add(IIf(DBNull.Value.Equals(reader!vendorName), "", reader!vendorName))
            li.SubItems.Add(IIf(DBNull.Value.Equals(reader!accountNumber), "", reader!accountNumber))
            li.SubItems.Add(IIf(DBNull.Value.Equals(reader!balance), "", reader!balance))
        End While
        reader.Close()
    End Sub
    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Me.Close()
    End Sub

    Private Sub frmFinAdjustBalances_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint
        If Me.IsMdiChild Then
            Dim PnlLoc As New Point
            PnlLoc.X = CInt((Me.Width - Me.pnlAdjustBal.Width) / 2)
            PnlLoc.Y = CInt((Me.Height - Me.pnlAdjustBal.Height) / 2.5)
            Me.pnlAdjustBal.Location = PnlLoc
        Else
            Me.pnlAdjustBal.Dock = DockStyle.Fill
        End If
    End Sub

    Private Sub btnClose_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub cboAccType_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboAccType.SelectedIndexChanged
        If Me.cboAccType.Text.Trim.Length <= 0 Then
            MsgBox("Account Type Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            Me.cboVendorName.Items.Clear()
            Me.cboVendorName.Text = ""
            Me.cboVendorName.SelectedIndex = -1

            Me.cboAccountNo.Items.Clear()
            Me.cboAccountNo.Text = ""
            Me.cboAccountNo.SelectedIndex = -1

            Me.txtBalance.Text = ""
            Me.txtBalance.Tag = Nothing

            Me.cmdAdjustBal.Connection = conn
            Me.cmdAdjustBal.CommandType = CommandType.Text
            Me.cmdAdjustBal.CommandText = "SELECT DISTINCT vendorName FROM tblFinPayAccounts " & _
                vbNewLine & " WHERE (modeName=@modeName) ORDER BY vendorName"
            Me.cmdAdjustBal.Parameters.Clear()
            Me.cmdAdjustBal.Parameters.AddWithValue("@modeName", Me.cboAccType.Text.Trim)
            reader = Me.cmdAdjustBal.ExecuteReader
            While reader.Read
                Me.cboVendorName.Items.Add(IIf(DBNull.Value.Equals(reader!vendorName), "", reader!vendorName))
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

    Private Sub cboVendorName_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboVendorName.SelectedIndexChanged
        If Me.cboAccType.Text.Trim.Length <= 0 Or Me.cboVendorName.Text.Trim.Length <= 0 Then
            MsgBox("Account Type or Vendor Name Is Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
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

            Me.txtBalance.Text = ""
            Me.txtBalance.Tag = Nothing

            Me.cmdAdjustBal.Connection = conn
            Me.cmdAdjustBal.CommandType = CommandType.Text
            Me.cmdAdjustBal.CommandText = "SELECT DISTINCT accountNumber FROM tblFinPayAccounts " & _
                vbNewLine & " WHERE (modeName=@modeName) AND (vendorName=@vendorName) ORDER BY accountNumber"
            Me.cmdAdjustBal.Parameters.Clear()
            Me.cmdAdjustBal.Parameters.AddWithValue("@modeName", Me.cboAccType.Text.Trim)
            Me.cmdAdjustBal.Parameters.AddWithValue("@vendorName", Me.cboVendorName.Text.Trim)
            reader = Me.cmdAdjustBal.ExecuteReader
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
        If Me.cboAccType.Text.Trim.Length <= 0 Or Me.cboVendorName.Text.Trim.Length <= 0 Or Me.cboAccountNo.Text.Trim.Length <= 0 Then
            MsgBox("Account Type or Vendor Name or Account Number Is Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()

            Me.txtBalance.Tag = Nothing
            Me.txtBalance.Text = ""

            Me.cmdAdjustBal.Connection = conn
            Me.cmdAdjustBal.CommandType = CommandType.Text
            Me.cmdAdjustBal.CommandText = "SELECT accountId,balance FROM tblFinPayAccounts  WHERE (modeName=@modeName) AND " & _
                vbNewLine & " (vendorName=@vendorName) AND (accountNumber=@accountNumber)"
            Me.cmdAdjustBal.Parameters.Clear()
            Me.cmdAdjustBal.Parameters.AddWithValue("@modeName", Me.cboAccType.Text.Trim)
            Me.cmdAdjustBal.Parameters.AddWithValue("@vendorName", Me.cboVendorName.Text.Trim)
            Me.cmdAdjustBal.Parameters.AddWithValue("@accountNumber", Me.cboAccountNo.Text.Trim)
            reader = Me.cmdAdjustBal.ExecuteReader
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

    Private Sub txtAdjustTo_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtAdjustTo.TextChanged
        If IsNumeric(Me.txtAdjustTo.Text.Trim) = False And Not (Me.txtAdjustTo.Text.Trim.Length <= 0) Then
            MsgBox("Non Numeric Values Detected.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Me.txtAdjustTo.Text = ""
            Exit Sub
        End If
    End Sub

    Private Sub btnUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdate.Click
        If Me.cboAccType.Text.Trim.Length <= 0 Then
            MsgBox("Select Account Type.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboVendorName.Text.Trim.Length <= 0 Then
            MsgBox("Vendor Name Is Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboAccountNo.Text.Trim.Length <= 0 Then
            MsgBox("Account Number Is Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.txtBalance.Text.Trim.Length <= 0 Then
            MsgBox("Balance Is Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.txtAdjustTo.Text.Trim.Length <= 0 Then
            MsgBox("Adjust Amount Is Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.txtAppFormNo.Text.Trim.Length <= 0 Then
            MsgBox("Approval Form Number Is Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
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

            Me.cmdAdjustBal.Connection = conn
            Me.cmdAdjustBal.CommandType = CommandType.StoredProcedure
            Me.cmdAdjustBal.CommandText = "sprocFinAdjustTrans"
            Me.cmdAdjustBal.Parameters.Clear()
            Me.cmdAdjustBal.Parameters.AddWithValue("@accountId", Me.txtBalance.Tag)
            Me.cmdAdjustBal.Parameters.AddWithValue("@balBefore", Me.txtBalance.Text.Trim)
            Me.cmdAdjustBal.Parameters.AddWithValue("@balAfter", Me.txtAdjustTo.Text.Trim)
            Me.cmdAdjustBal.Parameters.AddWithValue("@apprFormNo", Me.txtAppFormNo.Text.Trim)
            Me.cmdAdjustBal.Parameters.AddWithValue("@doneBy", userName.Trim)
            rec = Me.cmdAdjustBal.ExecuteNonQuery
            If rec > 0 Then
                MsgBox("Account Adjusted Successfully.", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "System Reply")
            End If
            clearTexts()
            loadCombos()
            loadList()
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub
    Private Sub clearTexts()
        Me.txtAdjustTo.Text = ""
        Me.txtAppFormNo.Text = ""
        Me.txtBalance.Text = ""
        Me.txtBalance.Tag = Nothing

        Me.cboAccType.Items.Clear()
        Me.cboAccType.Text = ""
        Me.cboAccType.SelectedIndex = -1

        Me.cboAccountNo.Items.Clear()
        Me.cboAccountNo.Text = ""
        Me.cboAccountNo.SelectedIndex = -1

        Me.cboVendorName.Items.Clear()
        Me.cboVendorName.Text = ""
        Me.cboVendorName.SelectedIndex = -1
    End Sub
End Class