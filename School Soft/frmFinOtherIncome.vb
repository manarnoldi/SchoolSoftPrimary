Imports System.Data.SqlClient
Public Class frmFinOtherIncome
    Dim rec As Integer
    Dim reader As SqlDataReader
    Dim cmdFinOthInc As New SqlCommand
    Private Sub frmFinOtherIncome_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
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

    Private Sub frmFinOtherIncome_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint
        If Me.IsMdiChild Then
            Dim PnlLoc As New Point
            PnlLoc.X = CInt((Me.Width - Me.pnlOtherInc.Width) / 2)
            PnlLoc.Y = CInt((Me.Height - Me.pnlOtherInc.Height) / 2.5)
            Me.pnlOtherInc.Location = PnlLoc
        Else
            Me.pnlOtherInc.Dock = DockStyle.Fill
        End If
    End Sub
    Private Sub loadCombos()
        Me.cboTermName.Items.Clear()
        Me.cboTermName.Text = ""
        Me.cboTermName.SelectedIndex = -1

        Me.cboYear.Items.Clear()
        Me.cboYear.Text = ""
        Me.cboYear.SelectedIndex = -1

        Me.cbomodeName.Items.Clear()
        Me.cbomodeName.Text = ""
        Me.cbomodeName.SelectedIndex = -1

        Me.cboSourceName.Items.Clear()
        Me.cboSourceName.Text = ""
        Me.cboSourceName.SelectedIndex = -1

        Me.cboRecFrom.Items.Clear()
        Me.cboRecFrom.Text = ""
        Me.cboRecFrom.SelectedIndex = -1

        Me.cmdFinOthInc.Connection = conn
        Me.cmdFinOthInc.CommandType = CommandType.Text
        Me.cmdFinOthInc.CommandText = "SELECT DISTINCT termName FROM tblSchoolCalendar WHERE (status=1) ORDER BY termName"
        Me.cmdFinOthInc.Parameters.Clear()
        reader = Me.cmdFinOthInc.ExecuteReader
        While reader.Read
            Me.cboTermName.Items.Add(IIf(DBNull.Value.Equals(reader!termName), "", reader!termName))
        End While
        reader.Close()

        Me.cmdFinOthInc.CommandText = "SELECT DISTINCT year FROM tblSchoolCalendar WHERE (status=1) ORDER BY year"
        Me.cmdFinOthInc.Parameters.Clear()
        reader = Me.cmdFinOthInc.ExecuteReader
        While reader.Read
            Me.cboYear.Items.Add(IIf(DBNull.Value.Equals(reader!year), "", reader!year))
        End While
        reader.Close()

        Me.cmdFinOthInc.CommandText = "SELECT DISTINCT modeName FROM tblFinPayModes ORDER BY modeName"
        Me.cmdFinOthInc.Parameters.Clear()
        reader = Me.cmdFinOthInc.ExecuteReader
        While reader.Read
            Me.cbomodeName.Items.Add(IIf(DBNull.Value.Equals(reader!modeName), "", reader!modeName))
        End While
        reader.Close()

        Me.cmdFinOthInc.CommandText = "SELECT DISTINCT sourceName FROM tblFinOtherIncome ORDER BY sourceName"
        Me.cmdFinOthInc.Parameters.Clear()
        reader = Me.cmdFinOthInc.ExecuteReader
        While reader.Read
            Me.cboSourceName.Items.Add(IIf(DBNull.Value.Equals(reader!sourceName), "", reader!sourceName))
        End While
        reader.Close()

        Me.cmdFinOthInc.CommandText = "SELECT DISTINCT receiptFrom FROM tblFinOtherIncome ORDER BY receiptFrom"
        Me.cmdFinOthInc.Parameters.Clear()
        reader = Me.cmdFinOthInc.ExecuteReader
        While reader.Read
            Me.cboRecFrom.Items.Add(IIf(DBNull.Value.Equals(reader!receiptFrom), "", reader!receiptFrom))
        End While
        reader.Close()
    End Sub

    Private Sub txtAmount_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtAmount.TextChanged
        If IsNumeric(Me.txtAmount.Text.Trim) = False And Not (Me.txtAmount.Text.Trim.Length <= 0) Then
            MsgBox("Non Numeric Values Detected.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Me.txtAmount.Text = ""
            Exit Sub
        End If
    End Sub

    Private Sub cbomodeName_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbomodeName.SelectedIndexChanged
        If Me.cbomodeName.Text.Trim.Length <= 0 Then
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

            Me.cboAccountNo.Tag = Nothing

            Me.cmdFinOthInc.Connection = conn
            Me.cmdFinOthInc.CommandType = CommandType.Text
            Me.cmdFinOthInc.CommandText = "SELECT accountNumber,accountId FROM tblFinPayAccounts WHERE (modeName=@modeName) ORDER BY accountNumber"
            Me.cmdFinOthInc.Parameters.Clear()
            Me.cmdFinOthInc.Parameters.AddWithValue("@modeName", Me.cbomodeName.Text.Trim)
            reader = Me.cmdFinOthInc.ExecuteReader
            While reader.Read
                Me.cboAccountNo.Items.Add(IIf(DBNull.Value.Equals(reader!accountNumber), "", reader!accountNumber))
                Me.cboAccountNo.Tag = IIf(DBNull.Value.Equals(reader!accountId), "", reader!accountId)
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

    Private Sub btnAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAdd.Click
        If Me.cboYear.Text.Trim.Length <= 0 Then
            MsgBox("Year is Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboTermName.Text.Trim.Length <= 0 Then
            MsgBox("Term is Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboSourceName.Text.Trim.Length <= 0 Then
            MsgBox("Source Name is Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboRecFrom.Text.Trim.Length <= 0 Then
            MsgBox("Receipt From is Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cbomodeName.Text.Trim.Length <= 0 Then
            MsgBox("Deposited To is Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboAccountNo.Text.Trim.Length <= 0 Then
            MsgBox("Account Number is Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.txtAmount.Text.Trim.Length <= 0 Then
            MsgBox("Amount is Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If

        li = Me.lstOtherInc.Items.Add(Me.cboSourceName.Text.Trim)
        li.SubItems.Add(Me.cboRecFrom.Text.Trim)
        li.SubItems.Add(Me.cboTermName.Text.Trim)
        li.SubItems.Add(Me.cboYear.Text.Trim)
        li.SubItems.Add(Me.cboAccountNo.Text.Trim)
        li.SubItems.Add(Me.txtAmount.Text.Trim)
        li.Tag = Me.cboAccountNo.Tag
        li.SubItems(1).Tag = Me.txtAmount.Tag
        li.SubItems(2).Tag = Me.dtpIncDate.Value.Date
    End Sub

    Private Sub cboYear_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboYear.SelectedIndexChanged, cboTermName.SelectedIndexChanged
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()

            Me.txtAmount.Tag = Nothing

            If Me.cboYear.Text.Trim.Length > 0 And Me.cboTermName.Text.Trim.Length > 0 Then
                Me.cmdFinOthInc.Connection = conn
                Me.cmdFinOthInc.CommandType = CommandType.Text
                Me.cmdFinOthInc.CommandText = "SELECT termId FROM tblSchoolCalendar WHERE (status=1) AND (termName=@termName) " & _
                    vbNewLine & " AND (year=@year)"
                Me.cmdFinOthInc.Parameters.Clear()
                Me.cmdFinOthInc.Parameters.AddWithValue("@year", Me.cboYear.Text.Trim)
                Me.cmdFinOthInc.Parameters.AddWithValue("@termName", Me.cboTermName.Text.Trim)
                reader = Me.cmdFinOthInc.ExecuteReader
                While reader.Read
                    Me.txtAmount.Tag = IIf(DBNull.Value.Equals(reader!termId), "", reader!termId)
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

    Private Sub btnRemove_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRemove.Click
        If Me.lstOtherInc.Items.Count <= 0 Then
            MsgBox("No items in the list.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.lstOtherInc.CheckedItems.Count <= 0 Then
            MsgBox("No checked items to remove.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If

        Dim result As MsgBoxResult = MsgBox("Remove Item/s", MsgBoxStyle.Question + MsgBoxStyle.ApplicationModal + MsgBoxStyle.YesNo, "Confirm Transaction")
        If result = MsgBoxResult.No Then
            Exit Sub
        End If

        For i = 0 To Me.lstOtherInc.CheckedItems.Count - 1
            Me.lstOtherInc.CheckedItems(0).Remove()
        Next
    End Sub
    Private Sub clearItems()
        Me.lstOtherInc.Items.Clear()
        Me.txtAmount.Text = ""
        Me.txtAmount.Tag = Nothing
        Me.cboAccountNo.Items.Clear()
        Me.cboAccountNo.Text = ""
        Me.cboAccountNo.SelectedIndex = -1
        Me.cboAccountNo.Tag = Nothing
    End Sub
    Private Sub btnClear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClear.Click
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            loadCombos()
            clearItems()
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub

    Private Sub btnUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdate.Click
        If Me.cboSourceName.Text.Trim.Count <= 0 Then
            MsgBox("Source Name Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboRecFrom.Text.Trim.Count <= 0 Then
            MsgBox("Receipt From Is Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.txtAmount.Text.Trim.Count <= 0 Then
            MsgBox("Amount Is Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.lstOtherInc.Items.Count <= 0 Then
            MsgBox("Missing items in the List.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If

        Dim result As MsgBoxResult = MsgBox("Save Record/s?", MsgBoxStyle.Question + MsgBoxStyle.ApplicationModal + MsgBoxStyle.YesNo, "Confirm Transaction")
        If result = MsgBoxResult.No Then
            Exit Sub
        End If

        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            Me.cmdFinOthInc.Connection = conn
            Me.cmdFinOthInc.CommandType = CommandType.StoredProcedure

            For i = 0 To Me.lstOtherInc.Items.Count - 1
                Me.cmdFinOthInc.CommandText = "sprocFinOtherIncome"
                Me.cmdFinOthInc.Parameters.Clear()
                Me.cmdFinOthInc.Parameters.AddWithValue("@sourceName", Me.lstOtherInc.Items(i).Text.Trim)
                Me.cmdFinOthInc.Parameters.AddWithValue("@receiptFrom", Me.lstOtherInc.Items(i).SubItems(1).Text.Trim)
                Me.cmdFinOthInc.Parameters.AddWithValue("@dateOfRec", Me.lstOtherInc.Items(i).SubItems(2).Tag)
                Me.cmdFinOthInc.Parameters.AddWithValue("@amount", Me.lstOtherInc.Items(i).SubItems(5).Text.Trim)
                Me.cmdFinOthInc.Parameters.AddWithValue("@userName", userName.Trim)
                Me.cmdFinOthInc.Parameters.AddWithValue("@termId", Me.lstOtherInc.Items(i).SubItems(1).Tag)
                Me.cmdFinOthInc.Parameters.AddWithValue("@accountId", Me.lstOtherInc.Items(i).Tag)
                rec = rec + Me.cmdFinOthInc.ExecuteNonQuery
            Next
            If rec > 0 Then
                MsgBox("Record/s Saved", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "SuccessFull Transaction")
            End If
            clearItems()
            loadCombos()
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub
End Class