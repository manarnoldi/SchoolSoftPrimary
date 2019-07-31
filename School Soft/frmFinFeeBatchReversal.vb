Imports System.Data.SqlClient
Public Class frmFinFeeBatchReversal
    Dim cmdFeeReceipt As New SqlCommand
    Dim rec As Integer = 0
    Dim reader As SqlDataReader
    Private Sub frmFinFeeBatchReversal_Load(sender As Object, e As EventArgs) Handles MyBase.Load
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
        Me.cboClass.Items.Clear()
        Me.cboClass.Text = ""
        Me.cboClass.SelectedIndex = -1

        Me.cboStream.Items.Clear()
        Me.cboStream.Text = ""
        Me.cboStream.SelectedIndex = -1

        Me.cboYear.Items.Clear()
        Me.cboYear.Text = ""
        Me.cboYear.SelectedIndex = -1

        Me.cboTerm.Items.Clear()
        Me.cboTerm.Text = ""
        Me.cboTerm.SelectedIndex = -1

        Me.cmdFeeReceipt.Connection = conn
        Me.cmdFeeReceipt.CommandType = CommandType.Text
        Me.cmdFeeReceipt.CommandText = "SELECT DISTINCT className FROM tblClasses WHERE (status=1) ORDER BY className"
        Me.cmdFeeReceipt.Parameters.Clear()
        reader = Me.cmdFeeReceipt.ExecuteReader
        While reader.Read
            Me.cboClass.Items.Add(IIf(DBNull.Value.Equals(reader!className), "", reader!className))
        End While
        reader.Close()

        Me.cmdFeeReceipt.CommandText = "SELECT DISTINCT stream FROM tblClasses WHERE (status=1) ORDER BY stream"
        Me.cmdFeeReceipt.Parameters.Clear()
        reader = Me.cmdFeeReceipt.ExecuteReader
        While reader.Read
            Me.cboStream.Items.Add(IIf(DBNull.Value.Equals(reader!stream), "", reader!stream))
        End While
        reader.Close()

        Me.cmdFeeReceipt.CommandText = "SELECT DISTINCT year FROM tblClasses WHERE (status=1) ORDER BY year"
        Me.cmdFeeReceipt.Parameters.Clear()
        reader = Me.cmdFeeReceipt.ExecuteReader
        While reader.Read
            Me.cboYear.Items.Add(IIf(DBNull.Value.Equals(reader!year), "", reader!year))
        End While
        reader.Close()

        Me.cmdFeeReceipt.CommandText = "SELECT DISTINCT termName FROM  tblSchoolCalendar WHERE (status=1) ORDER BY termName"
        Me.cmdFeeReceipt.Parameters.Clear()
        reader = Me.cmdFeeReceipt.ExecuteReader
        While reader.Read
            Me.cboTerm.Items.Add(IIf(DBNull.Value.Equals(reader!termName), "", reader!termName))
        End While
        reader.Close()
    End Sub
    Private Sub frmFinFeeBatchReversal_Paint(sender As Object, e As PaintEventArgs) Handles Me.Paint
        If Me.IsMdiChild Then
            Dim PnlLoc As New Point
            PnlLoc.X = CInt((Me.Width - Me.pnlFinFeeBatchReversal.Width) / 2)
            PnlLoc.Y = CInt((Me.Height - Me.pnlFinFeeBatchReversal.Height) / 2.5)
            Me.pnlFinFeeBatchReversal.Location = PnlLoc
        Else
            Me.pnlFinFeeBatchReversal.Dock = DockStyle.Fill
        End If
    End Sub

    Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub btnLoad_Click(sender As Object, e As EventArgs) Handles btnLoad.Click
        If Me.cboClass.Text.Trim.Length <= 0 Then
            MsgBox("Class Name Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboStream.Text.Trim.Length <= 0 Then
            MsgBox("Stream Is Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboYear.Text.Trim.Length <= 0 Then
            MsgBox("Year Is Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboTerm.Text.Trim.Length <= 0 Then
            MsgBox("Term Is Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If

        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()

            clearTexts()

            Me.cmdFeeReceipt.Connection = conn
            Me.cmdFeeReceipt.CommandType = CommandType.Text
            Me.cmdFeeReceipt.CommandText = "SELECT * FROM vwStudClass WHERE (className=@className) AND " +
            "(stream=@stream) AND (year=@year) ORDER BY admNo"
            Me.cmdFeeReceipt.Parameters.Clear()
            Me.cmdFeeReceipt.Parameters.AddWithValue("@className", Me.cboClass.Text.Trim)
            Me.cmdFeeReceipt.Parameters.AddWithValue("@stream", Me.cboStream.Text.Trim)
            Me.cmdFeeReceipt.Parameters.AddWithValue("@year", Me.cboYear.Text.Trim)
            reader = Me.cmdFeeReceipt.ExecuteReader
            While reader.Read
                Dim classStudId(1) As Integer
                classStudId(0) = IIf(DBNull.Value.Equals(reader!studId), "", reader!studId)
                classStudId(1) = IIf(DBNull.Value.Equals(reader!classStudId), "", reader!classStudId)
                li = lstStudent.Items.Add(IIf(DBNull.Value.Equals(reader!admNo), "", reader!admNo))
                li.SubItems.Add(IIf(DBNull.Value.Equals(reader!FullName), "", reader!FullName))
                li.Tag = classStudId
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

    Private Sub clearTexts()
        Me.lstStudent.Items.Clear()
        Me.lstFeeReceiptDetails.Items.Clear()
        Me.lstFeeVoteDetails.Items.Clear()
        Me.lstFeeVoteBalancesDetails.Items.Clear()
        Me.txtLastYearBal.Text = ""
        Me.txtLastYearBalPaid.Text = ""
        Me.txtLastYearRem.Text = ""
        Me.txtThisYearBal.Text = ""
        Me.cbRevPay.CheckState = CheckState.Unchecked
        Me.cbRevFeeSetup.CheckState = CheckState.Unchecked
    End Sub

    Private Sub lstStudent_ColumnClick(sender As Object, e As ColumnClickEventArgs) Handles lstStudent.ColumnClick
        If (e.Column() = 0) And (Me.lstStudent.Items.Count > 0) Then
            For Each Li As ListViewItem In Me.lstStudent.Items
                Li.Checked = Not (Li.Checked)
            Next
        End If
    End Sub

    Private Sub lstStudent_Click(sender As Object, e As EventArgs) Handles lstStudent.Click, lstStudent.SelectedIndexChanged
        If Me.lstStudent.Items.Count <= 0 Then
            Exit Sub
        End If
        If Me.lstStudent.SelectedItems.Count = 1 Then
            Try
                If conn.State = ConnectionState.Closed Then
                    conn.Open()
                End If
                dbconnection()
                Dim termId As Integer = 0
                Dim studId As Integer = Me.lstStudent.SelectedItems(0).Tag(0)

                Me.lstFeeReceiptDetails.Items.Clear()
                Me.lstFeeVoteDetails.Items.Clear()
                Me.ToolStripStatusStudentDetails.Text = ""
                Me.lstFeeVoteBalancesDetails.Items.Clear()
                Me.txtLastYearBal.Text = ""
                Me.txtLastYearBalPaid.Text = ""
                Me.txtLastYearRem.Text = ""
                Me.txtThisYearBal.Text = ""

                Me.cmdFeeReceipt.CommandType = CommandType.Text
                Me.cmdFeeReceipt.Connection = conn
                Me.cmdFeeReceipt.CommandText = "SELECT termId FROM tblSchoolCalendar WHERE (termName=@termName) AND (year=@year) AND (status=1)"
                Me.cmdFeeReceipt.Parameters.Clear()
                Me.cmdFeeReceipt.Parameters.AddWithValue("@termName", Me.cboTerm.Text.Trim)
                Me.cmdFeeReceipt.Parameters.AddWithValue("@year", Me.cboYear.Text.Trim)
                reader = Me.cmdFeeReceipt.ExecuteReader
                While reader.Read
                    termId = IIf(DBNull.Value.Equals(reader!termId), "", reader!termId)
                End While
                reader.Close()

                Me.ToolStripStatusStudentDetails.Alignment = ToolStripItemAlignment.Right
                Me.ToolStripStatusStudentDetails.Text = "Adm No: " + Me.lstStudent.SelectedItems(0).Text + " FullName: " +
                    Me.lstStudent.SelectedItems(0).SubItems(1).Text

                Me.cmdFeeReceipt.CommandText = "SELECT * FROM tblFinFeePayDetails WHERE (studId=@studId) AND (year=@year) " +
                    " AND (termId=@termId)"
                Me.cmdFeeReceipt.Parameters.Clear()
                Me.cmdFeeReceipt.Parameters.AddWithValue("@studId", studId)
                Me.cmdFeeReceipt.Parameters.AddWithValue("@year", Me.cboYear.Text.Trim)
                Me.cmdFeeReceipt.Parameters.AddWithValue("@termId", termId)
                reader = Me.cmdFeeReceipt.ExecuteReader
                While reader.Read
                    Dim actualPayDate As Date
                    actualPayDate = IIf(DBNull.Value.Equals(reader!actualPayDate), "", reader!actualPayDate)
                    li = lstFeeReceiptDetails.Items.Add(IIf(DBNull.Value.Equals(reader!receiptNoAll), "", reader!receiptNoAll))
                    li.SubItems.Add(IIf(DBNull.Value.Equals(reader!amountPaid), "", reader!amountPaid))
                    li.SubItems.Add(IIf(DBNull.Value.Equals(reader!lastYearBal), "", reader!lastYearBal))
                    li.SubItems.Add(IIf(DBNull.Value.Equals(reader!lastYearBalPaid), "", reader!lastYearBalPaid))
                    li.SubItems.Add(IIf(DBNull.Value.Equals(reader!lastYearBalRem), "", reader!lastYearBalRem))
                    li.SubItems.Add(actualPayDate.ToString("dd-MM-yyyy"))
                    li.Tag = IIf(DBNull.Value.Equals(reader!feePayId), "", reader!feePayId)
                    If IIf(DBNull.Value.Equals(reader!reversed), "", reader!reversed) Then
                        li.BackColor = Color.Red
                    End If
                End While
                reader.Close()

                Me.cmdFeeReceipt.CommandText = "SELECT studFeeId,b.voteHeadName,amount,amountPaid,a.balance FROM tblFinFeeStudent a INNER JOIN  " +
                "tblFinFeeVoteHeads b ON a.finVoteHeadId = b.finVoteHeadId WHERE (classStudId=@classStudId) AND (termId=@termId)"
                Me.cmdFeeReceipt.Parameters.Clear()
                Me.cmdFeeReceipt.Parameters.AddWithValue("@classStudId", Me.lstStudent.SelectedItems(0).Tag(1))
                Me.cmdFeeReceipt.Parameters.AddWithValue("@termId", termId)
                reader = Me.cmdFeeReceipt.ExecuteReader
                lstFeeVoteBalancesDetails.Items.Clear()

                While reader.Read
                    li = lstFeeVoteBalancesDetails.Items.Add(IIf(DBNull.Value.Equals(reader!voteHeadName), "", reader!voteHeadName))
                    li.SubItems.Add(IIf(DBNull.Value.Equals(reader!amount), "", reader!amount))
                    li.SubItems.Add(IIf(DBNull.Value.Equals(reader!amountPaid), "", reader!amountPaid))
                    li.SubItems.Add(IIf(DBNull.Value.Equals(reader!balance), "", reader!balance))
                    li.Tag = IIf(DBNull.Value.Equals(reader!studFeeId), "", reader!studFeeId)
                End While
                reader.Close()

                Dim amountTotal As Integer = 0
                Dim amountPaidTotal As Integer = 0
                Dim balanceTotal As Integer = 0

                For Each li As ListViewItem In lstFeeVoteBalancesDetails.Items
                    amountTotal += li.SubItems(1).Text
                    amountPaidTotal += li.SubItems(2).Text
                    balanceTotal += li.SubItems(3).Text
                Next
                li = lstFeeVoteBalancesDetails.Items.Add("Total")
                li.SubItems.Add(amountTotal)
                li.SubItems.Add(amountPaidTotal)
                li.SubItems.Add(balanceTotal)
                li.Font = New Font(li.Font, FontStyle.Bold)

                Me.cmdFeeReceipt.CommandText = "SELECT * FROM tblFinFeeBalanceBF WHERE (year=@year) AND (StudId=@StudId)"
                Me.cmdFeeReceipt.Parameters.Clear()
                Me.cmdFeeReceipt.Parameters.AddWithValue("@year", CInt(Me.cboYear.Text.Trim) - 1)
                Me.cmdFeeReceipt.Parameters.AddWithValue("@StudId", studId)
                reader = Me.cmdFeeReceipt.ExecuteReader
                While reader.Read
                    Me.txtLastYearBal.Text = IIf(DBNull.Value.Equals(reader!amount), 0, reader!amount)
                    Me.txtLastYearBalPaid.Text = IIf(DBNull.Value.Equals(reader!amountPaid), 0, reader!amountPaid)
                    Me.txtLastYearRem.Text = IIf(DBNull.Value.Equals(reader!balance), 0, reader!balance)
                End While
                reader.Close()

                Me.cmdFeeReceipt.CommandText = "SELECT amount FROM tblFinFeeBalance WHERE (year=@year) AND (StudId=@StudId)"
                Me.cmdFeeReceipt.Parameters.Clear()
                Me.cmdFeeReceipt.Parameters.AddWithValue("@year", CInt(Me.cboYear.Text.Trim))
                Me.cmdFeeReceipt.Parameters.AddWithValue("@StudId", studId)
                reader = Me.cmdFeeReceipt.ExecuteReader
                While reader.Read
                    Me.txtThisYearBal.Text = IIf(DBNull.Value.Equals(reader!amount), 0, reader!amount)
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

    Private Sub lstFeeReceiptDetails_Click(sender As Object, e As EventArgs) Handles lstFeeReceiptDetails.Click
        If Me.lstFeeReceiptDetails.Items.Count <= 0 Then
            Exit Sub
        End If
        If Me.lstFeeReceiptDetails.SelectedItems.Count = 1 Then
            Try
                If conn.State = ConnectionState.Closed Then
                    conn.Open()
                End If
                dbconnection()
                Me.lstFeeVoteDetails.Items.Clear()

                Me.cmdFeeReceipt.CommandType = CommandType.Text
                Me.cmdFeeReceipt.Connection = conn
                Me.cmdFeeReceipt.CommandText = "SELECT voteHeadName,amount,reversed FROM tblFinFeePayVotes a INNER JOIN " +
                    "tblFinFeeVoteHeads b ON a.voteHeadId = b.finVoteHeadId WHERE (feePayId=@feePayId)"
                Me.cmdFeeReceipt.Parameters.Clear()
                Me.cmdFeeReceipt.Parameters.AddWithValue("@feePayId", Me.lstFeeReceiptDetails.SelectedItems(0).Tag)
                reader = Me.cmdFeeReceipt.ExecuteReader
                While reader.Read
                    li = Me.lstFeeVoteDetails.Items.Add(IIf(DBNull.Value.Equals(reader!voteHeadName), "", reader!voteHeadName))
                    li.SubItems.Add(IIf(DBNull.Value.Equals(reader!amount), "", reader!amount))
                    li.SubItems.Add(IIf(DBNull.Value.Equals(reader!reversed), "", reader!reversed))
                    If (IIf(DBNull.Value.Equals(reader!reversed), "", reader!reversed)) Then
                        li.BackColor = Color.Red
                    End If
                End While
                reader.Close()

                Dim votesTotal As Integer = 0

                For Each li As ListViewItem In lstFeeVoteDetails.Items
                    votesTotal += li.SubItems(1).Text
                Next

                li = lstFeeVoteDetails.Items.Add("Total")
                li.SubItems.Add(votesTotal)
                li.SubItems.Add("")
                li.Font = New Font(li.Font, FontStyle.Bold)
            Catch ex As Exception
                MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Finally
                If conn.State = ConnectionState.Open Then
                    conn.Close()
                End If
            End Try
        End If
    End Sub

    Private Sub lstFeeReceiptDetails_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lstFeeReceiptDetails.SelectedIndexChanged

    End Sub

    Private Sub cboClass_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboClass.SelectedIndexChanged,
            cboStream.SelectedIndexChanged, cboTerm.SelectedIndexChanged, cboYear.SelectedIndexChanged
        clearTexts()
    End Sub

    Private Sub checkIfFeeReversed()

    End Sub

    Private Sub btnUpdate_Click(sender As Object, e As EventArgs) Handles btnUpdate.Click
        If Me.cboClass.Text.Trim.Length <= 0 Then
            MsgBox("Class Name Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboStream.Text.Trim.Length <= 0 Then
            MsgBox("Stream Is Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboYear.Text.Trim.Length <= 0 Then
            MsgBox("Year Is Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboTerm.Text.Trim.Length <= 0 Then
            MsgBox("Term Is Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.lstStudent.Items.Count <= 0 Then
            MsgBox("Click load to select the students to load.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.lstStudent.CheckedItems.Count <= 0 Then
            MsgBox("Check the students to reverse details.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cbRevPay.CheckState = CheckState.Unchecked And Me.cbRevFeeSetup.CheckState = CheckState.Unchecked Then
            MsgBox("Check the operation you want to undertake.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cbRevFeeSetup.CheckState = CheckState.Checked And Me.cbRevPay.CheckState = CheckState.Unchecked Then
            MsgBox("You cannot reverse fee setup before reversing fee payment/s.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If

        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()

            Dim result As MsgBoxResult = MsgBox("Update Record?", MsgBoxStyle.Question + MsgBoxStyle.ApplicationModal + MsgBoxStyle.YesNo, "Confirm Transaction")
            If result = MsgBoxResult.No Then
                Exit Sub
            End If

            rec = 0
            Me.cmdFeeReceipt.Connection = conn
            Me.cmdFeeReceipt.CommandType = CommandType.StoredProcedure

            For Each li As ListViewItem In lstStudent.CheckedItems
                'Reversing the fee payments done
                If Me.cbRevPay.CheckState = CheckState.Checked Then
                    Me.cmdFeeReceipt.CommandText = "sprocFinFeeReverseAllStudentPayment"
                    Me.cmdFeeReceipt.Parameters.Clear()
                    Me.cmdFeeReceipt.Parameters.AddWithValue("@admNo", li.Text)
                    Me.cmdFeeReceipt.Parameters.AddWithValue("@term", Me.cboTerm.Text.Trim)
                    Me.cmdFeeReceipt.Parameters.AddWithValue("@className", Me.cboClass.Text.Trim)
                    Me.cmdFeeReceipt.Parameters.AddWithValue("@year", Me.cboYear.Text.Trim)
                    Me.cmdFeeReceipt.Parameters.AddWithValue("@stream", Me.cboStream.Text.Trim)
                    rec = rec + Me.cmdFeeReceipt.ExecuteNonQuery
                End If
                'Reversing the fee setup done
                If Me.cbRevFeeSetup.CheckState = CheckState.Checked Then
                    Me.cmdFeeReceipt.CommandText = "sprocFinFeeReverseAllStudentSetUp"
                    Me.cmdFeeReceipt.Parameters.Clear()
                    Me.cmdFeeReceipt.Parameters.AddWithValue("@admNo", li.Text)
                    Me.cmdFeeReceipt.Parameters.AddWithValue("@term", Me.cboTerm.Text.Trim)
                    Me.cmdFeeReceipt.Parameters.AddWithValue("@class", Me.cboClass.Text.Trim)
                    Me.cmdFeeReceipt.Parameters.AddWithValue("@year", Me.cboYear.Text.Trim)
                    Me.cmdFeeReceipt.Parameters.AddWithValue("@stream", Me.cboStream.Text.Trim)
                    rec = rec + Me.cmdFeeReceipt.ExecuteNonQuery
                End If
            Next

            If rec > 0 Then
                MsgBox("Record/s Updated Successfully.", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "System Reply")
            End If
            clearTexts()
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub

    Private Sub cbRevPay_CheckedChanged(sender As Object, e As EventArgs) Handles cbRevPay.CheckedChanged, cbRevFeeSetup.CheckedChanged
        If cbRevFeeSetup.CheckState = CheckState.Checked Then
            cbRevPay.Checked = True
        End If
    End Sub
End Class