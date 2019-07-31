Imports System.Data.SqlClient
Public Class frmFinPayAccounts
    Dim cmdPayAccounts As New SqlCommand
    Dim reader As SqlDataReader
    Dim rec As Integer = 0
    Private Sub frmFinPayAccounts_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
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

    Private Sub loadList()
        Me.lstPayAccounts.Items.Clear()

        Me.cmdPayAccounts.Connection = conn
        Me.cmdPayAccounts.CommandType = CommandType.Text
        Me.cmdPayAccounts.CommandText = "SELECT * FROM tblFinPayAccounts ORDER BY modeName,accountNumber"
        Me.cmdPayAccounts.Parameters.Clear()
        reader = Me.cmdPayAccounts.ExecuteReader
        While reader.Read
            li = Me.lstPayAccounts.Items.Add(IIf(DBNull.Value.Equals(reader!modeName), "", reader!modeName))
            li.SubItems.Add(IIf(DBNull.Value.Equals(reader!location), "", reader!location))
            li.SubItems.Add(IIf(DBNull.Value.Equals(reader!vendorName), "", reader!vendorName))
            li.SubItems.Add(IIf(DBNull.Value.Equals(reader!accountName), "", reader!accountName))
            li.SubItems.Add(IIf(DBNull.Value.Equals(reader!accountNumber), "", reader!accountNumber))
            li.Tag = IIf(DBNull.Value.Equals(reader!accountId), "", reader!accountId)
        End While
        reader.Close()
    End Sub

    Private Sub loadCombos()
        Me.cboPayMode.Items.Clear()
        Me.cboPayMode.Text = ""
        Me.cboPayMode.SelectedIndex = -1

        Me.cmdPayAccounts.Connection = conn
        Me.cmdPayAccounts.CommandType = CommandType.Text
        Me.cmdPayAccounts.CommandText = "SELECT DISTINCT modeName FROM tblFinPayModes ORDER BY modeName"
        Me.cmdPayAccounts.Parameters.Clear()
        reader = Me.cmdPayAccounts.ExecuteReader
        While reader.Read
            Me.cboPayMode.Items.Add(IIf(DBNull.Value.Equals(reader!modeName), "", reader!modeName))
        End While
        reader.Close()
    End Sub
    Private Sub frmFinPayAccounts_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint
        If Me.IsMdiChild Then
            Dim PnlLoc As New Point
            PnlLoc.X = CInt((Me.Width - Me.pnlPayAccounts.Width) / 2)
            PnlLoc.Y = CInt((Me.Height - Me.pnlPayAccounts.Height) / 2.5)
            Me.pnlPayAccounts.Location = PnlLoc
        Else
            Me.pnlPayAccounts.Dock = DockStyle.Fill
        End If
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub CLOSEToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CLOSEToolStripMenuItem.Click
        Me.Close()
    End Sub

    Private Sub btnView_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnView.Click
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            Me.cboPayMode.Enabled = True
            Me.txtAccountName.ReadOnly = False
            Me.txtVendorName.ReadOnly = False
            Me.btnUpdate.Enabled = False
            Me.btnSave.Enabled = True
            Me.cboPayMode.Tag = Nothing
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

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        If Me.cboPayMode.Text.Trim.Length <= 0 Then
            MsgBox("Mode Name Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.txtAccountName.Text.Trim.Length <= 0 Then
            MsgBox("Account Name Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.txtAccountNo.Text.Trim.Length <= 0 Then
            MsgBox("Account Number Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.txtLocation.Text.Trim.Length <= 0 Then
            MsgBox("Location Is Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.txtVendorName.Text.Trim.Length <= 0 Then
            MsgBox("Vendor Name Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If
        Try
            Dim recordExists As Boolean
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            If Me.cboPayMode.Text = "CASH" Then
                recordExists = checkIfCashRecordExists(Me.cboPayMode.Text.Trim)
                If recordExists = True Then
                    MsgBox("Cash Account Already Registered!", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
                    Exit Sub
                End If
            End If
            recordExists = checkIfRecordExists(Me.cboPayMode.Text.Trim, Me.txtAccountName.Text.Trim, Me.txtAccountNo.Text.Trim)
            If recordExists = True Then
                MsgBox("Duplicate Found in the database!", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
                Exit Sub
            End If

            Dim result As MsgBoxResult = MsgBox("Save Record?", MsgBoxStyle.Question + MsgBoxStyle.ApplicationModal + MsgBoxStyle.YesNo, "Confirm Transaction")
            If result = MsgBoxResult.No Then
                Exit Sub
            End If

            Me.cmdPayAccounts.Connection = conn
            Me.cmdPayAccounts.CommandType = CommandType.StoredProcedure
            Me.cmdPayAccounts.CommandText = "sprocFinPayAccounts"
            Me.cmdPayAccounts.Parameters.Clear()

            Me.cmdPayAccounts.Parameters.AddWithValue("@modeName", Me.cboPayMode.Text.Trim)
            Me.cmdPayAccounts.Parameters.AddWithValue("@regBy", userName.Trim)
            Me.cmdPayAccounts.Parameters.AddWithValue("@queryType", "INSERT")
            Me.cmdPayAccounts.Parameters.AddWithValue("@vendorName", Me.txtVendorName.Text.Trim)
            Me.cmdPayAccounts.Parameters.AddWithValue("@location", Me.txtLocation.Text.Trim)
            Me.cmdPayAccounts.Parameters.AddWithValue("@accountName", Me.txtAccountName.Text.Trim)
            Me.cmdPayAccounts.Parameters.AddWithValue("@accountNumber", Me.txtAccountNo.Text.Trim)
            rec = Me.cmdPayAccounts.ExecuteNonQuery
            If rec > 0 Then
                MsgBox("Record Saved Successfully.", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "System Reply")
            End If
            loadCombos()
            loadList()
            clearTexts()
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub
    Private Function checkIfRecordExists(ByVal modeName As String, ByVal accountName As String, ByVal accountNumber As String)
        Dim recordExists As Boolean = False
        Me.cmdPayAccounts.Connection = conn
        Me.cmdPayAccounts.CommandType = CommandType.Text
        Me.cmdPayAccounts.CommandText = "SELECT * FROM tblFinPayAccounts WHERE (modeName=@modeName) AND (accountName=@accountName) " & _
            vbNewLine & " AND (accountNumber=@accountNumber)"
        Me.cmdPayAccounts.Parameters.Clear()
        Me.cmdPayAccounts.Parameters.AddWithValue("@modeName", modeName.Trim)
        Me.cmdPayAccounts.Parameters.AddWithValue("@accountName", accountName.Trim)
        Me.cmdPayAccounts.Parameters.AddWithValue("@accountNumber", accountNumber.Trim)
        reader = Me.cmdPayAccounts.ExecuteReader
        If reader.HasRows = True Then
            recordExists = True
        ElseIf reader.HasRows = False Then
            recordExists = False
        End If
        reader.Close()
        Return recordExists
    End Function
    Private Function checkIfCashRecordExists(ByVal modeName As String)
        Dim recordExists As Boolean = False
        Me.cmdPayAccounts.Connection = conn
        Me.cmdPayAccounts.CommandType = CommandType.Text
        Me.cmdPayAccounts.CommandText = "SELECT * FROM tblFinPayAccounts WHERE (modeName=@modeName)"
        Me.cmdPayAccounts.Parameters.Clear()
        Me.cmdPayAccounts.Parameters.AddWithValue("@modeName", modeName.Trim)
        reader = Me.cmdPayAccounts.ExecuteReader
        If reader.HasRows = True Then
            recordExists = True
        ElseIf reader.HasRows = False Then
            recordExists = False
        End If
        reader.Close()
        Return recordExists
    End Function
    Private Sub UPDATEToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UPDATEToolStripMenuItem.Click
        If Me.lstPayAccounts.Items.Count <= 0 Then
            MsgBox("Missing Items In The List.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.lstPayAccounts.CheckedItems.Count <= 0 Then
            MsgBox("Missing Checked Items In The List.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.lstPayAccounts.CheckedItems.Count > 1 Then
            MsgBox("Edit one item At A Time.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()

            Me.cboPayMode.Items.Clear()
            Me.cboPayMode.Text = ""
            Me.cboPayMode.SelectedIndex = -1

            Me.cboPayMode.Items.Add(Me.lstPayAccounts.CheckedItems(0).Text.Trim)
            Me.cboPayMode.SelectedIndex = 0
            Me.cboPayMode.Tag = Me.lstPayAccounts.CheckedItems(0).Tag
            Me.txtAccountName.Text = Me.lstPayAccounts.CheckedItems(0).SubItems(3).Text.Trim
            Me.txtAccountNo.Text = Me.lstPayAccounts.CheckedItems(0).SubItems(4).Text.Trim
            Me.txtLocation.Text = Me.lstPayAccounts.CheckedItems(0).SubItems(1).Text.Trim
            Me.txtVendorName.Text = Me.lstPayAccounts.CheckedItems(0).SubItems(2).Text.Trim

            Me.cboPayMode.Enabled = False
            Me.txtAccountName.ReadOnly = True
            Me.txtVendorName.ReadOnly = True
            Me.btnUpdate.Enabled = True
            Me.btnSave.Enabled = False
            Me.lstPayAccounts.Items.Clear()
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub

    Private Sub btnUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdate.Click
        If Me.cboPayMode.Text.Trim.Length <= 0 Then
            MsgBox("Mode Name Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.txtAccountName.Text.Trim.Length <= 0 Then
            MsgBox("Account Name Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.txtAccountNo.Text.Trim.Length <= 0 Then
            MsgBox("Account Number Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.txtLocation.Text.Trim.Length <= 0 Then
            MsgBox("Location Is Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.txtVendorName.Text.Trim.Length <= 0 Then
            MsgBox("Vendor Name Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            'MSGBOX
            Dim recordExists As Boolean = checkIfRecordExists(Me.cboPayMode.Text.Trim, Me.txtAccountName.Text.Trim, Me.txtAccountNo.Text.Trim)
            If recordExists = True Then
                MsgBox("Duplicate Found in the database!", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
                Exit Sub
            End If

            Dim result As MsgBoxResult = MsgBox("Update Record?", MsgBoxStyle.Question + MsgBoxStyle.ApplicationModal + MsgBoxStyle.YesNo, "Confirm Transaction")
            If result = MsgBoxResult.No Then
                Exit Sub
            End If

            Me.cmdPayAccounts.Connection = conn
            Me.cmdPayAccounts.CommandType = CommandType.StoredProcedure
            Me.cmdPayAccounts.CommandText = "sprocFinPayAccounts"
            Me.cmdPayAccounts.Parameters.Clear()
            Me.cmdPayAccounts.Parameters.AddWithValue("@modeName", Me.cboPayMode.Text.Trim)
            Me.cmdPayAccounts.Parameters.AddWithValue("@regBy", userName.Trim)
            Me.cmdPayAccounts.Parameters.AddWithValue("@queryType", "UPDATE")
            Me.cmdPayAccounts.Parameters.AddWithValue("@vendorName", Me.txtVendorName.Text.Trim)
            Me.cmdPayAccounts.Parameters.AddWithValue("@location", Me.txtLocation.Text.Trim)
            Me.cmdPayAccounts.Parameters.AddWithValue("@accountName", Me.txtAccountName.Text.Trim)
            Me.cmdPayAccounts.Parameters.AddWithValue("@accountNumber", Me.txtAccountNo.Text.Trim)
            Me.cmdPayAccounts.Parameters.AddWithValue("@accountId", Me.cboPayMode.Tag)
            rec = Me.cmdPayAccounts.ExecuteNonQuery
            If rec > 0 Then
                MsgBox("Record Updated Successfully.", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "System Reply")
            End If
            clearTexts()
            Me.cboPayMode.Enabled = True
            Me.txtAccountName.ReadOnly = False
            Me.txtVendorName.ReadOnly = False
            Me.btnUpdate.Enabled = False
            Me.btnSave.Enabled = True
            Me.cboPayMode.Tag = Nothing
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
        Me.txtAccountName.Text = ""
        Me.txtAccountNo.Text = ""
        Me.txtLocation.Text = ""
        Me.txtVendorName.Text = ""
    End Sub

    Private Sub DELETEToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DELETEToolStripMenuItem.Click
        If Me.lstPayAccounts.Items.Count <= 0 Then
            MsgBox("Missing Items In The List.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.lstPayAccounts.CheckedItems.Count <= 0 Then
            MsgBox("Missing Checked Items In The List.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
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

            For i = 0 To Me.lstPayAccounts.CheckedItems.Count - 1
                Me.cmdPayAccounts.Connection = conn
                Me.cmdPayAccounts.CommandType = CommandType.StoredProcedure
                Me.cmdPayAccounts.CommandText = "sprocFinPayAccounts"
                Me.cmdPayAccounts.Parameters.Clear()
                Me.cmdPayAccounts.Parameters.AddWithValue("@modeName", Me.lstPayAccounts.CheckedItems(i).Text.Trim)
                Me.cmdPayAccounts.Parameters.AddWithValue("@regBy", userName.Trim)
                Me.cmdPayAccounts.Parameters.AddWithValue("@queryType", "DELETE")
                Me.cmdPayAccounts.Parameters.AddWithValue("@vendorName", Me.lstPayAccounts.CheckedItems(i).SubItems(2).Text.Trim)
                Me.cmdPayAccounts.Parameters.AddWithValue("@location", Me.lstPayAccounts.CheckedItems(i).SubItems(1).Text.Trim)
                Me.cmdPayAccounts.Parameters.AddWithValue("@accountName", Me.lstPayAccounts.CheckedItems(i).SubItems(3).Text.Trim)
                Me.cmdPayAccounts.Parameters.AddWithValue("@accountNumber", Me.lstPayAccounts.CheckedItems(i).SubItems(4).Text.Trim)
                Me.cmdPayAccounts.Parameters.AddWithValue("@accountId", Me.lstPayAccounts.CheckedItems(i).Tag)
                rec = Me.cmdPayAccounts.ExecuteNonQuery
            Next
            If rec > 0 Then
                MsgBox("Record/s Deleted Successfully.", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "System Reply")
            End If
            clearTexts()
            Me.cboPayMode.Enabled = True
            Me.txtAccountName.ReadOnly = False
            Me.txtVendorName.ReadOnly = False
            Me.btnUpdate.Enabled = False
            Me.btnSave.Enabled = True
            Me.cboPayMode.Tag = Nothing
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

    Private Sub cboPayMode_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboPayMode.SelectedIndexChanged

    End Sub

    Private Sub lstPayAccounts_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lstPayAccounts.SelectedIndexChanged

    End Sub
End Class