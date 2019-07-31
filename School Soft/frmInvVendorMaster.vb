Imports System.Data.SqlClient
Public Class frmInvVendorMaster
    Dim reader As SqlDataReader
    Dim rec As Integer
    Dim cmdInvVendorMaster As New SqlCommand
    Dim vendorNo As Integer = 0
    Dim dateOfInc As String = Nothing
    Private Sub frmInvVendorMaster_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            loadCombos()
            loadVendorNo()
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub
    Private Sub loadCombos()
        Me.cboVendorCategory.Items.Clear()
        Me.cboVendorCategory.Text = ""
        Me.cboVendorCategory.SelectedIndex = -1

        Me.cboVendorLocation.Items.Clear()
        Me.cboVendorLocation.Text = ""
        Me.cboVendorLocation.SelectedIndex = -1

        Me.cmdInvVendorMaster.Connection = conn
        Me.cmdInvVendorMaster.CommandType = CommandType.Text
        Me.cmdInvVendorMaster.CommandText = "SELECT DISTINCT vendorCategory FROM tblInvVendors WHERE (vendorCategory IS NOT NULL) " & _
            vbNewLine & " ORDER BY vendorCategory"
        Me.cmdInvVendorMaster.Parameters.Clear()
        reader = Me.cmdInvVendorMaster.ExecuteReader
        While reader.Read
            Me.cboVendorCategory.Items.Add(IIf(DBNull.Value.Equals(reader!vendorCategory), "", reader!vendorCategory))
        End While
        reader.Close()

        Me.cmdInvVendorMaster.CommandText = "SELECT DISTINCT vendorLocation FROM tblInvVendors WHERE (vendorLocation IS NOT NULL) " & _
                    vbNewLine & " ORDER BY vendorLocation"
        Me.cmdInvVendorMaster.Parameters.Clear()
        reader = Me.cmdInvVendorMaster.ExecuteReader
        While reader.Read
            Me.cboVendorLocation.Items.Add(IIf(DBNull.Value.Equals(reader!vendorLocation), "", reader!vendorLocation))
        End While
        reader.Close()

    End Sub
    Private Sub loadVendorNo()
        Me.cmdInvVendorMaster.Connection = conn
        Me.cmdInvVendorMaster.CommandType = CommandType.Text
        Me.cmdInvVendorMaster.CommandText = "SELECT ISNULL(MAX(vendorNo),0)+1 AS newNo FROM tblInvVendors"
        Me.cmdInvVendorMaster.Parameters.Clear()
        reader = Me.cmdInvVendorMaster.ExecuteReader
        While reader.Read
            vendorNo = IIf(DBNull.Value.Equals(reader!newNo), "", reader!newNo)
            Me.txtVendorCode.Text = "INVEN" & vendorNo.ToString("0000")
        End While
        reader.Close()
    End Sub
    Private Sub frmInvVendorMaster_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint
        If Me.IsMdiChild Then
            Dim PnlLoc As New Point
            PnlLoc.X = CInt((Me.Width - Me.pnlVendorMaster.Width) / 2)
            PnlLoc.Y = CInt((Me.Height - Me.pnlVendorMaster.Height) / 2.5)
            Me.pnlVendorMaster.Location = PnlLoc
        Else
            Me.pnlVendorMaster.Dock = DockStyle.Fill
        End If
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub CLOSEToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CLOSEToolStripMenuItem.Click
        Me.Close()
    End Sub

    Private Sub txtVendorEmail_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtVendorEmail.LostFocus
        If Me.txtVendorEmail.Text.Trim.Contains("@") = False Or Me.txtVendorEmail.Text.Trim.Contains(".") = False Then
            MsgBox("Emai Not Valid!", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Me.txtVendorEmail.Clear()
            Exit Sub
        End If
    End Sub

    Private Sub btnView_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnView.Click
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            clearTexts()
            enableControls()
            loadCombos()
            loadList()
            loadVendorNo()
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub
    Private Sub loadList()
        dateofinc = Nothing
        Me.lstVendorMaster.Items.Clear()

        Me.cmdInvVendorMaster.Connection = conn
        Me.cmdInvVendorMaster.CommandType = CommandType.Text
        Me.cmdInvVendorMaster.CommandText = "SELECT * FROM tblInvVendors ORDER BY vendorCode"
        Me.cmdInvVendorMaster.Parameters.Clear()
        reader = Me.cmdInvVendorMaster.ExecuteReader
        While reader.Read
            dateOfInc = CDate(IIf(DBNull.Value.Equals(reader!dateOfIncorporation), "", reader!dateOfIncorporation)).Day.ToString("00") & "-" & _
            CDate(IIf(DBNull.Value.Equals(reader!dateOfIncorporation), "", reader!dateOfIncorporation)).Month.ToString("00") & "-" & _
            CDate(IIf(DBNull.Value.Equals(reader!dateOfIncorporation), "", reader!dateOfIncorporation)).Year.ToString("0000")

            li = Me.lstVendorMaster.Items.Add(IIf(DBNull.Value.Equals(reader!vendorCode), "", reader!vendorCode))
            li.SubItems.Add(IIf(DBNull.Value.Equals(reader!vendorName), "", reader!vendorName))
            li.SubItems.Add(IIf(DBNull.Value.Equals(reader!vendorCategory), "", reader!vendorCategory))
            li.SubItems.Add(IIf(DBNull.Value.Equals(reader!vendorAddress), "", reader!vendorAddress))
            li.SubItems.Add(IIf(DBNull.Value.Equals(reader!vendorLocation), "", reader!vendorLocation))
            li.SubItems.Add(IIf(DBNull.Value.Equals(reader!vendorContacts), "", reader!vendorContacts))
            li.SubItems.Add(IIf(DBNull.Value.Equals(reader!vendorEmail), "", reader!vendorEmail))
            li.SubItems.Add(dateOfInc)
            li.Tag = IIf(DBNull.Value.Equals(reader!dateOfIncorporation), "", reader!dateOfIncorporation)
        End While
        reader.Close()
    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        If Me.txtVendorCode.Text.Trim.Length <= 0 Then
            MsgBox("Vendor Code Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboVendorCategory.Text.Trim.Length <= 0 Then
            MsgBox("Vendor Category Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.txtVendorAddress.Text.Trim.Length <= 0 Then
            MsgBox("Vendor Address Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.txtVendorName.Text.Trim.Length <= 0 Then
            MsgBox("Vendor Name Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboVendorLocation.Text.Trim.Length <= 0 Then
            MsgBox("Vendor Location Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.txtVendorEmail.Text.Trim.Length <= 0 Then
            MsgBox("Vendor Email Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.txtVendorContacts.Text.Trim.Length <= 0 Then
            MsgBox("Vendor Contacts Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            Dim result As MsgBoxResult = MsgBox("Save Record?", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.YesNo, "Confirm Transaction")
            If result = MsgBoxResult.No Then
                Exit Sub
            End If

            Me.cmdInvVendorMaster.Connection = conn
            Me.cmdInvVendorMaster.CommandType = CommandType.StoredProcedure
            Me.cmdInvVendorMaster.CommandText = "sprocInvVendors"
            Me.cmdInvVendorMaster.Parameters.Clear()
            Me.cmdInvVendorMaster.Parameters.AddWithValue("@vendorNo", vendorNo)
            Me.cmdInvVendorMaster.Parameters.AddWithValue("@vendorCode", Me.txtVendorCode.Text.Trim)
            Me.cmdInvVendorMaster.Parameters.AddWithValue("@vendorName", Me.txtVendorName.Text.Trim)
            Me.cmdInvVendorMaster.Parameters.AddWithValue("@vendorCategory", Me.cboVendorCategory.Text.Trim)
            Me.cmdInvVendorMaster.Parameters.AddWithValue("@vendorAddress", Me.txtVendorAddress.Text.Trim)
            Me.cmdInvVendorMaster.Parameters.AddWithValue("@vendorContacts", Me.txtVendorContacts.Text.Trim)
            Me.cmdInvVendorMaster.Parameters.AddWithValue("@vendorEmail", Me.txtVendorEmail.Text.Trim)
            Me.cmdInvVendorMaster.Parameters.AddWithValue("@vendorLocation", Me.cboVendorLocation.Text.Trim)
            Me.cmdInvVendorMaster.Parameters.AddWithValue("@dateOfIncorporation", dtpRegDate.Value.Date)
            Me.cmdInvVendorMaster.Parameters.AddWithValue("@regBy", userName.Trim)
            Me.cmdInvVendorMaster.Parameters.AddWithValue("@queryType", 1)
            rec = Me.cmdInvVendorMaster.ExecuteNonQuery
            If rec > 0 Then
                MsgBox("Record Saved", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "SuccessFull Transaction")
            End If
            loadCombos()
            clearTexts()
            enableControls()
            Me.lstVendorMaster.Items.Clear()
            loadVendorNo()
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub
    Private Sub clearTexts()
        Me.txtSearchVendorName.Text = ""
        Me.txtVendorAddress.Text = ""
        Me.txtVendorCode.Text = ""
        Me.txtVendorContacts.Text = ""
        Me.txtVendorEmail.Text = ""
        Me.txtVendorName.Text = ""
    End Sub
    Private Sub enableControls()
        Me.btnSave.Enabled = True
        Me.btnUpdate.Enabled = False
        Me.cboVendorCategory.Enabled = True
        Me.dtpRegDate.Enabled = True
    End Sub
    Private Sub disbleControls()
        Me.btnSave.Enabled = False
        Me.btnUpdate.Enabled = True
        Me.cboVendorCategory.Enabled = False
        Me.dtpRegDate.Enabled = False
    End Sub
    Private Sub txtSearchVendorName_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtSearchVendorName.TextChanged
        Me.lstVendorMaster.Items.Clear()
        If Me.txtSearchVendorName.Text.Trim.Length < 3 Then
            Me.lstVendorMaster.Items.Clear()
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            Me.cmdInvVendorMaster.Connection = conn
            Me.cmdInvVendorMaster.CommandType = CommandType.Text
            Me.cmdInvVendorMaster.CommandText = "SELECT * FROM tblInvVendors WHERE (vendorName LIKE @vendorName) ORDER BY vendorCode"
            Me.cmdInvVendorMaster.Parameters.Clear()
            Me.cmdInvVendorMaster.Parameters.AddWithValue("@vendorName", String.Format("%{0}%", Me.txtSearchVendorName.Text.Trim))
            reader = Me.cmdInvVendorMaster.ExecuteReader
            While reader.Read
                dateOfInc = CDate(IIf(DBNull.Value.Equals(reader!dateOfIncorporation), "", reader!dateOfIncorporation)).Day.ToString("00") & "-" & _
                CDate(IIf(DBNull.Value.Equals(reader!dateOfIncorporation), "", reader!dateOfIncorporation)).Month.ToString("00") & "-" & _
                CDate(IIf(DBNull.Value.Equals(reader!dateOfIncorporation), "", reader!dateOfIncorporation)).Year.ToString("0000")

                li = Me.lstVendorMaster.Items.Add(IIf(DBNull.Value.Equals(reader!vendorCode), "", reader!vendorCode))
                li.SubItems.Add(IIf(DBNull.Value.Equals(reader!vendorName), "", reader!vendorName))
                li.SubItems.Add(IIf(DBNull.Value.Equals(reader!vendorCategory), "", reader!vendorCategory))
                li.SubItems.Add(IIf(DBNull.Value.Equals(reader!vendorAddress), "", reader!vendorAddress))
                li.SubItems.Add(IIf(DBNull.Value.Equals(reader!vendorLocation), "", reader!vendorLocation))
                li.SubItems.Add(IIf(DBNull.Value.Equals(reader!vendorContacts), "", reader!vendorContacts))
                li.SubItems.Add(IIf(DBNull.Value.Equals(reader!vendorEmail), "", reader!vendorEmail))
                li.SubItems.Add(dateOfInc)
                li.Tag = IIf(DBNull.Value.Equals(reader!dateOfIncorporation), "", reader!dateOfIncorporation)
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

    Private Sub UPDATEToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UPDATEToolStripMenuItem.Click
        If Me.lstVendorMaster.Items.Count <= 0 Then
            MsgBox("No items in the list.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.lstVendorMaster.CheckedItems.Count <= 0 Then
            MsgBox("No checked items in the list.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.lstVendorMaster.CheckedItems.Count > 1 Then
            MsgBox("Update one item at a time.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            Me.txtVendorCode.Text = Me.lstVendorMaster.CheckedItems(0).Text.Trim
            Me.cboVendorCategory.Text = Me.lstVendorMaster.CheckedItems(0).SubItems(2).Text.Trim
            Me.txtVendorAddress.Text = Me.lstVendorMaster.CheckedItems(0).SubItems(3).Text.Trim
            Me.txtVendorName.Text = Me.lstVendorMaster.CheckedItems(0).SubItems(1).Text.Trim
            Me.cboVendorLocation.Text = Me.lstVendorMaster.CheckedItems(0).SubItems(4).Text.Trim
            Me.txtVendorEmail.Text = Me.lstVendorMaster.CheckedItems(0).SubItems(6).Text.Trim
            Me.txtVendorContacts.Text = Me.lstVendorMaster.CheckedItems(0).SubItems(5).Text.Trim
            Me.dtpRegDate.Value = Me.lstVendorMaster.CheckedItems(0).Tag
            disbleControls()
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub

    Private Sub btnUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdate.Click
        If Me.txtVendorCode.Text.Trim.Length <= 0 Then
            MsgBox("Vendor Code Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboVendorCategory.Text.Trim.Length <= 0 Then
            MsgBox("Vendor Category Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.txtVendorAddress.Text.Trim.Length <= 0 Then
            MsgBox("Vendor Address Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.txtVendorName.Text.Trim.Length <= 0 Then
            MsgBox("Vendor Name Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboVendorLocation.Text.Trim.Length <= 0 Then
            MsgBox("Vendor Location Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.txtVendorEmail.Text.Trim.Length <= 0 Then
            MsgBox("Vendor Email Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.txtVendorContacts.Text.Trim.Length <= 0 Then
            MsgBox("Vendor Contacts Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            Dim result As MsgBoxResult = MsgBox("Update Record?", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.YesNo, "Confirm Transaction")
            If result = MsgBoxResult.No Then
                Exit Sub
            End If

            Me.cmdInvVendorMaster.Connection = conn
            Me.cmdInvVendorMaster.CommandType = CommandType.StoredProcedure
            Me.cmdInvVendorMaster.CommandText = "sprocInvVendors"
            Me.cmdInvVendorMaster.Parameters.Clear()
            Me.cmdInvVendorMaster.Parameters.AddWithValue("@vendorNo", vendorNo)
            Me.cmdInvVendorMaster.Parameters.AddWithValue("@vendorCode", Me.txtVendorCode.Text.Trim)
            Me.cmdInvVendorMaster.Parameters.AddWithValue("@vendorName", Me.txtVendorName.Text.Trim)
            Me.cmdInvVendorMaster.Parameters.AddWithValue("@vendorCategory", Me.cboVendorCategory.Text.Trim)
            Me.cmdInvVendorMaster.Parameters.AddWithValue("@vendorAddress", Me.txtVendorAddress.Text.Trim)
            Me.cmdInvVendorMaster.Parameters.AddWithValue("@vendorContacts", Me.txtVendorContacts.Text.Trim)
            Me.cmdInvVendorMaster.Parameters.AddWithValue("@vendorEmail", Me.txtVendorEmail.Text.Trim)
            Me.cmdInvVendorMaster.Parameters.AddWithValue("@vendorLocation", Me.cboVendorLocation.Text.Trim)
            Me.cmdInvVendorMaster.Parameters.AddWithValue("@dateOfIncorporation", dtpRegDate.Value.Date)
            Me.cmdInvVendorMaster.Parameters.AddWithValue("@regBy", userName.Trim)
            Me.cmdInvVendorMaster.Parameters.AddWithValue("@queryType", 2)
            rec = Me.cmdInvVendorMaster.ExecuteNonQuery
            If rec > 0 Then
                MsgBox("Record Updated", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "SuccessFull Transaction")
            End If
            loadCombos()
            clearTexts()
            enableControls()
            Me.lstVendorMaster.Items.Clear()
            loadVendorNo()
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub

    Private Sub DELETEToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DELETEToolStripMenuItem.Click
        If Me.lstVendorMaster.Items.Count <= 0 Then
            MsgBox("No items in the list.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.lstVendorMaster.CheckedItems.Count <= 0 Then
            MsgBox("No checked items in the list.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
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
            For i = 0 To Me.lstVendorMaster.CheckedItems.Count - 1
                Me.cmdInvVendorMaster.Connection = conn
                Me.cmdInvVendorMaster.CommandType = CommandType.StoredProcedure
                Me.cmdInvVendorMaster.CommandText = "sprocInvVendors"
                Me.cmdInvVendorMaster.Parameters.Clear()
                Me.cmdInvVendorMaster.Parameters.AddWithValue("@vendorNo", vendorNo)
                Me.cmdInvVendorMaster.Parameters.AddWithValue("@vendorCode", Me.lstVendorMaster.CheckedItems(i).Text.Trim)
                Me.cmdInvVendorMaster.Parameters.AddWithValue("@vendorName", Me.lstVendorMaster.CheckedItems(i).SubItems(1).Text.Trim)
                Me.cmdInvVendorMaster.Parameters.AddWithValue("@vendorCategory", Me.lstVendorMaster.CheckedItems(i).SubItems(2).Text.Trim)
                Me.cmdInvVendorMaster.Parameters.AddWithValue("@vendorAddress", Me.lstVendorMaster.CheckedItems(i).SubItems(3).Text.Trim)
                Me.cmdInvVendorMaster.Parameters.AddWithValue("@vendorContacts", Me.lstVendorMaster.CheckedItems(i).SubItems(5).Text.Trim)
                Me.cmdInvVendorMaster.Parameters.AddWithValue("@vendorEmail", Me.lstVendorMaster.CheckedItems(i).SubItems(6).Text.Trim)
                Me.cmdInvVendorMaster.Parameters.AddWithValue("@vendorLocation", Me.lstVendorMaster.CheckedItems(i).SubItems(4).Text.Trim)
                Me.cmdInvVendorMaster.Parameters.AddWithValue("@dateOfIncorporation", Me.lstVendorMaster.CheckedItems(i).Tag)
                Me.cmdInvVendorMaster.Parameters.AddWithValue("@regBy", userName.Trim)
                Me.cmdInvVendorMaster.Parameters.AddWithValue("@queryType", 3)
                rec = rec + Me.cmdInvVendorMaster.ExecuteNonQuery
            Next
            If rec > 0 Then
                MsgBox("Record/s Deleted!", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "SuccessFull Transaction")
            End If
            loadCombos()
            clearTexts()
            enableControls()
            Me.lstVendorMaster.Items.Clear()
            loadVendorNo()
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub
End Class