Imports System.Data.SqlClient
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports System.Drawing.Printing
Public Class frmFinFeeReceipt2
    Dim cmdFeeReceipt As New SqlCommand
    Dim rec As Integer
    Dim reader As SqlDataReader
    Dim receiptInitials As String = ""

    Private Sub frmFinFeeReceipt2_Load(sender As Object, e As EventArgs) Handles MyBase.Load
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

    Private Sub clearTexts()
        Me.txtPrevYearBalEditable.Text = ""
        Me.txtStudName.Text = ""
        Me.txtPrevYearBal.Text = ""
        Me.txtRecNo.Text = ""
        Me.txtTotalBalance.Text = ""
        Me.txtTermBal.Text = ""
        Me.txtOverPayAmount.Text = ""
        Me.txtPayAmount.Text = ""
        Me.txtAmount.Text = ""
        Me.txtFeeBalance.Text = ""
        Me.txtThisYearBal.Text = ""
        Me.cboVoteHead.Items.Clear()
        Me.cboVoteHead.Text = ""
        Me.txtAdmNo.Text = ""
        Me.cboVoteHead.SelectedIndex = -1

        Me.cboPayMode.SelectedIndex = -1

        Me.cboPayAccount.Items.Clear()
        Me.cboPayAccount.Text = ""
        Me.cboPayAccount.SelectedIndex = -1

        Me.cboPaidBy.Items.Clear()
        Me.cboPaidBy.Text = ""
        Me.cboPaidBy.SelectedIndex = -1

        Me.lstPrevPayDetails.Items.Clear()
        Me.lstFeeApportDetails.Items.Clear()

        Me.cbAutoPortion.Checked = False

        Me.cboVoteHead.Enabled = True
        Me.txtAmount.Enabled = True
        Me.btnAdd.Enabled = True

        Me.lstFeeApportDetails.Enabled = True
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

        Me.cboPayMode.Items.Clear()
        Me.cboPayMode.Text = ""
        Me.cboPayMode.SelectedIndex = -1

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

        Me.cmdFeeReceipt.CommandText = "SELECT DISTINCT modeName FROM  tblFinPayModes ORDER BY modeName"
        Me.cmdFeeReceipt.Parameters.Clear()
        reader = Me.cmdFeeReceipt.ExecuteReader
        While reader.Read
            Me.cboPayMode.Items.Add(IIf(DBNull.Value.Equals(reader!modeName), "", reader!modeName))
        End While
        reader.Close()
    End Sub

    Private Sub frmFinFeeReceipt2_Paint(sender As Object, e As PaintEventArgs) Handles Me.Paint
        If Me.IsMdiChild Then
            Dim PnlLoc As New Point
            PnlLoc.X = CInt((Me.Width - Me.pnlFeeReceipt.Width) / 2)
            PnlLoc.Y = CInt((Me.Height - Me.pnlFeeReceipt.Height) / 2.5)
            Me.pnlFeeReceipt.Location = PnlLoc
        Else
            Me.pnlFeeReceipt.Dock = DockStyle.Fill
        End If
    End Sub

    Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub txtSearchStudName_TextChanged(sender As Object, e As EventArgs) Handles txtSearchStudName.TextChanged
        If Me.txtSearchStudName.Text.Trim.Length <= 0 Then
            Exit Sub
        End If
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
        ElseIf Me.txtSearchStudName.Text.Trim.Length <= 1 Then
            clearTexts()
            Me.lstStudentDetails.Items.Clear()
            Exit Sub
        End If

        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            clearTexts()
            Me.lstStudentDetails.Items.Clear()

            Me.cmdFeeReceipt.Connection = conn
            Me.cmdFeeReceipt.CommandType = CommandType.Text
            Me.cmdFeeReceipt.CommandText = "SELECT admNo,FullName FROM vwStudClass WHERE (className=@className) AND (stream=@stream) AND (year=@year) AND (FullName LIKE @FullName)"
            Me.cmdFeeReceipt.Parameters.Clear()
            Me.cmdFeeReceipt.Parameters.AddWithValue("@className", Me.cboClass.Text.Trim)
            Me.cmdFeeReceipt.Parameters.AddWithValue("@stream", Me.cboStream.Text.Trim)
            Me.cmdFeeReceipt.Parameters.AddWithValue("@year", Me.cboYear.Text.Trim)
            Me.cmdFeeReceipt.Parameters.AddWithValue("@FullName", String.Format("%{0}%", TryCast(Me.txtSearchStudName.Text.Trim, String).Trim))

            reader = Me.cmdFeeReceipt.ExecuteReader
            While reader.Read
                li = Me.lstStudentDetails.Items.Add(IIf(DBNull.Value.Equals(reader!admNo), "", reader!admNo))
                li.SubItems.Add(IIf(DBNull.Value.Equals(reader!FullName), "", reader!FullName))
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

    Private Sub lstStudentDetails_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lstStudentDetails.SelectedIndexChanged

    End Sub

    Private Function checkIfStudentFeeSet(ByVal admNo As String) As Boolean
        Dim feeSet = False
        Me.cmdFeeReceipt.Connection = conn
        Me.cmdFeeReceipt.CommandType = CommandType.Text
        Me.cmdFeeReceipt.CommandText = "SELECT COUNT(admNo) AS Count FROM  vwFinFeeStudBalance WHERE (className=@className) " +
            " AND (stream=@stream) AND (year=@year) AND (admNo=@admNo) AND (termName=@termName)"
        Me.cmdFeeReceipt.Parameters.Clear()
        Me.cmdFeeReceipt.Parameters.AddWithValue("@year", Me.cboYear.Text.Trim)
        Me.cmdFeeReceipt.Parameters.AddWithValue("@stream", Me.cboStream.Text.Trim)
        Me.cmdFeeReceipt.Parameters.AddWithValue("@className", Me.cboClass.Text.Trim)
        Me.cmdFeeReceipt.Parameters.AddWithValue("@admNo", admNo)
        Me.cmdFeeReceipt.Parameters.AddWithValue("@termName", Me.cboTerm.Text.Trim)
        reader = Me.cmdFeeReceipt.ExecuteReader
        While reader.Read
            If (IIf(DBNull.Value.Equals(reader!Count), "", reader!Count)) > 0 Then
                feeSet = True
            End If
        End While
        reader.Close()
        Return feeSet
    End Function

    Private Sub lstStudentDetails_Click(sender As Object, e As EventArgs) Handles lstStudentDetails.Click
        If Me.lstStudentDetails.Items.Count < 1 Then
            MsgBox("Search the student to load them in the list.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.lstStudentDetails.SelectedItems.Count > 1 Then
            MsgBox("Select one student at a time!", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.lstStudentDetails.SelectedItems.Count < 1 Then
            MsgBox("Click on the admission number of the student!", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboClass.Text.Trim.Length <= 0 Then
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
            Me.Cursor = Cursors.WaitCursor
            Dim termBalance As Double = 0
            Dim thisYearBalance As Double = 0
            Dim prevYearBalance As Double = 0
            Dim totalBalance As Double = 0
            clearTexts()
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()

            For Each item As ListViewItem In Me.lstStudentDetails.Items
                item.BackColor = Color.Empty
            Next

            If checkIfStudentFeeSet(Me.lstStudentDetails.SelectedItems(0).Text.Trim) = False Then
                MsgBox("Student Fee for the term has not been set!", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
                Me.btnUpdate.Enabled = False
                Exit Sub
            Else
                Me.btnUpdate.Enabled = True
            End If

            Me.lstStudentDetails.SelectedItems(0).BackColor = Color.DarkSeaGreen


            Me.cboVoteHead.Text = ""
            Me.cboVoteHead.SelectedIndex = -1
            Me.txtAmount.Text = ""

            Me.cboVoteHead.Enabled = True
            Me.txtAmount.Enabled = True
            Me.btnAdd.Enabled = True
            Me.lstFeeApportDetails.Enabled = True

            Me.txtStudName.Text = ""

            Me.cmdFeeReceipt.Connection = conn
            Me.cmdFeeReceipt.CommandType = CommandType.Text
            Me.cmdFeeReceipt.CommandText = "SELECT admNo,FullName FROM tblStudDetails WHERE (status=1) AND (admNo=@admNo)"
            Me.cmdFeeReceipt.Parameters.Clear()
            Me.cmdFeeReceipt.Parameters.AddWithValue("@admNo", Me.lstStudentDetails.SelectedItems(0).Text.Trim)
            reader = Me.cmdFeeReceipt.ExecuteReader
            While reader.Read
                Me.txtAdmNo.Text = (IIf(DBNull.Value.Equals(reader!admNo), "", reader!admNo))
                Me.txtStudName.Text = (IIf(DBNull.Value.Equals(reader!FullName), "", reader!FullName))
            End While
            reader.Close()

            Me.cboVoteHead.Items.Clear()
            Me.cboVoteHead.Text = ""
            Me.cboVoteHead.SelectedIndex = -1

            Me.cmdFeeReceipt.CommandText = "SELECT DISTINCT voteHeadName FROM  vwFinFeeStudent WHERE (className=@className)" +
                " AND (stream=@stream) AND (year=@year) AND (admNo=@admNo) AND (termName=@termName) ORDER BY voteHeadName"
            Me.cmdFeeReceipt.Parameters.Clear()
            Me.cmdFeeReceipt.Parameters.AddWithValue("@year", Me.cboYear.Text.Trim)
            Me.cmdFeeReceipt.Parameters.AddWithValue("@stream", Me.cboStream.Text.Trim)
            Me.cmdFeeReceipt.Parameters.AddWithValue("@className", Me.cboClass.Text.Trim)
            Me.cmdFeeReceipt.Parameters.AddWithValue("@admNo", Me.txtAdmNo.Text.Trim)
            Me.cmdFeeReceipt.Parameters.AddWithValue("@termName", Me.cboTerm.Text.Trim)
            reader = Me.cmdFeeReceipt.ExecuteReader
            While reader.Read
                Me.cboVoteHead.Items.Add(IIf(DBNull.Value.Equals(reader!voteHeadName), "", reader!voteHeadName))
            End While
            reader.Close()

            Me.cmdFeeReceipt.CommandText = "SELECT TOP 1 initials FROM tblSchoolDetails"
            Me.cmdFeeReceipt.Parameters.Clear()
            reader = Me.cmdFeeReceipt.ExecuteReader
            While reader.Read
                receiptInitials = IIf(DBNull.Value.Equals(reader!initials), "", reader!initials)
            End While
            reader.Close()

            Me.txtThisYearBal.Text = ""
            Me.txtTermBal.Text = ""
            Me.txtOverPayAmount.Text = ""

            Me.cmdFeeReceipt.CommandText = "SELECT * FROM  vwFinFeeStudBalance WHERE (className=@className) " +
                " AND (stream=@stream) AND (year=@year) AND (admNo=@admNo) AND (termName=@termName)"
            Me.cmdFeeReceipt.Parameters.Clear()
            Me.cmdFeeReceipt.Parameters.AddWithValue("@year", Me.cboYear.Text.Trim)
            Me.cmdFeeReceipt.Parameters.AddWithValue("@stream", Me.cboStream.Text.Trim)
            Me.cmdFeeReceipt.Parameters.AddWithValue("@className", Me.cboClass.Text.Trim)
            Me.cmdFeeReceipt.Parameters.AddWithValue("@admNo", Me.txtAdmNo.Text.Trim)
            Me.cmdFeeReceipt.Parameters.AddWithValue("@termName", Me.cboTerm.Text.Trim)
            reader = Me.cmdFeeReceipt.ExecuteReader
            While reader.Read
                thisYearBalance = (IIf(DBNull.Value.Equals(reader!totalBalance), 0, reader!totalBalance))
                termBalance = (IIf(DBNull.Value.Equals(reader!termBalance), 0, reader!termBalance))
            End While
            reader.Close()

            Me.txtThisYearBal.Text = thisYearBalance
            Me.txtTermBal.Text = termBalance

            Me.cmdFeeReceipt.CommandText = "SELECT balance FROM vwFinFeeBalanceBF WHERE (year=@year) AND (admNo=@admNo)"
            Me.cmdFeeReceipt.Parameters.Clear()
            Me.cmdFeeReceipt.Parameters.AddWithValue("@year", CInt(Me.cboYear.Text.Trim) - 1)
            Me.cmdFeeReceipt.Parameters.AddWithValue("@admNo", Me.txtAdmNo.Text.Trim)
            reader = Me.cmdFeeReceipt.ExecuteReader
            While reader.Read
                prevYearBalance = (IIf(DBNull.Value.Equals(reader!balance), 0, reader!balance))
            End While
            reader.Close()

            Me.txtPrevYearBal.Text = prevYearBalance

            Me.lstPrevPayDetails.Items.Clear()

            Me.cmdFeeReceipt.CommandText = "SELECT * FROM  vwFinFeePrevPayments WHERE (year=@year) AND (admNo=@admNo)"
            Me.cmdFeeReceipt.Parameters.Clear()
            Me.cmdFeeReceipt.Parameters.AddWithValue("@year", Me.cboYear.Text.Trim)
            Me.cmdFeeReceipt.Parameters.AddWithValue("@admNo", Me.txtAdmNo.Text.Trim)
            reader = Me.cmdFeeReceipt.ExecuteReader
            While reader.Read
                Dim actualDate As String = CDate(IIf(DBNull.Value.Equals(reader!actualPayDate), "", reader!actualPayDate)).Day.ToString("00") & "-" &
                CDate(IIf(DBNull.Value.Equals(reader!actualPayDate), "", reader!actualPayDate)).Month.ToString("00") & "-" &
                CDate(IIf(DBNull.Value.Equals(reader!actualPayDate), "", reader!actualPayDate)).Year.ToString("00")
                li = Me.lstPrevPayDetails.Items.Add(IIf(DBNull.Value.Equals(reader!receiptNoAll), "", reader!receiptNoAll))
                li.SubItems.Add(actualDate)
                li.SubItems.Add(IIf(DBNull.Value.Equals(reader!Year), "", reader!Year))
                li.SubItems.Add(IIf(DBNull.Value.Equals(reader!termName), "", reader!termName))
                li.SubItems.Add(IIf(DBNull.Value.Equals(reader!amountPaid), "", reader!amountPaid))
            End While
            reader.Close()

            Me.txtRecNo.Text = ""

            Dim term As String = Nothing
            If Me.cboTerm.Text.Trim = "TERM 1" Then
                term = "01"
            ElseIf Me.cboTerm.Text.Trim = "TERM 2" Then
                term = "02"
            ElseIf Me.cboTerm.Text.Trim = "TERM 3" Then
                term = "03"
            End If
            Me.cmdFeeReceipt.CommandText = "SELECT ISNULL(MAX(receiptNo),0)+1 AS newRecNo FROM tblFinFeePayDetails INNER JOIN tblSchoolCalendar " +
            "ON tblFinFeePayDetails.termId =tblSchoolCalendar.termId WHERE (tblSchoolCalendar.termName=@termName) AND (tblSchoolCalendar.year=@year)" +
            "AND (tblFinFeePayDetails.year=@year)"
            Me.cmdFeeReceipt.Parameters.Clear()
            Me.cmdFeeReceipt.Parameters.AddWithValue("@year", Me.cboYear.Text.Trim)
            Me.cmdFeeReceipt.Parameters.AddWithValue("@termName", Me.cboTerm.Text.Trim)
            reader = Me.cmdFeeReceipt.ExecuteReader
            While reader.Read
                Me.txtRecNo.Text = receiptInitials & Me.cboYear.Text.Trim.Substring(2, 2) & term.Trim & "/" &
                    CInt(IIf(DBNull.Value.Equals(reader!newRecNo), "", reader!newRecNo)).ToString("0000")
                Me.txtRecNo.Tag = (IIf(DBNull.Value.Equals(reader!newRecNo), "", reader!newRecNo))
            End While
            reader.Close()

            Me.lstFeeApportDetails.Items.Clear()

            Me.cmdFeeReceipt.CommandText = "SELECT * FROM  vwFinFeeStudent WHERE (className=@className) AND (stream=@stream) " &
           vbNewLine & " AND (year=@year) AND (admNo=@admNo) AND (termName=@termName) ORDER BY priority"
            Me.cmdFeeReceipt.Parameters.Clear()
            Me.cmdFeeReceipt.Parameters.AddWithValue("@year", Me.cboYear.Text.Trim)
            Me.cmdFeeReceipt.Parameters.AddWithValue("@stream", Me.cboStream.Text.Trim)
            Me.cmdFeeReceipt.Parameters.AddWithValue("@className", Me.cboClass.Text.Trim)
            Me.cmdFeeReceipt.Parameters.AddWithValue("@admNo", Me.txtAdmNo.Text.Trim)
            Me.cmdFeeReceipt.Parameters.AddWithValue("@termName", Me.cboTerm.Text.Trim)
            reader = Me.cmdFeeReceipt.ExecuteReader
            While reader.Read
                li = Me.lstFeeApportDetails.Items.Add(IIf(DBNull.Value.Equals(reader!voteHeadName), "", reader!voteHeadName))
                li.SubItems.Add(IIf(DBNull.Value.Equals(reader!amount), "", reader!amount))
                li.SubItems.Add(IIf(DBNull.Value.Equals(reader!amountPaid), "", reader!amountPaid))
                li.SubItems.Add(IIf(DBNull.Value.Equals(reader!toPayBalance), "", reader!toPayBalance))
                li.SubItems.Add(0)
                li.Tag = (IIf(DBNull.Value.Equals(reader!studFeeId), "", reader!studFeeId))
            End While
            reader.Close()
            totalBalance = thisYearBalance + prevYearBalance
            Me.txtTotalBalance.Text = totalBalance
            Me.txtFeeBalance.Text = CDbl(Me.txtTermBal.Text.Trim) + CDbl(Me.txtPrevYearBal.Text.Trim)
            Me.txtOverPayAmount.Text = 0
            Me.txtPrevYearBalEditable.Text = Me.txtPrevYearBal.Text
            Me.Cursor = Cursors.Default
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            Me.Cursor = Cursors.Default
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub

    Private Function checkIfReceiptNoExists(ByVal receiptNo As String)
        Dim recordExists As Boolean = False
        Me.cmdFeeReceipt.Connection = conn
        Me.cmdFeeReceipt.CommandType = CommandType.Text
        Me.cmdFeeReceipt.CommandText = "SELECT * FROM tblFinFeePayDetails WHERE (receiptNoAll=@receiptNoAll)"
        Me.cmdFeeReceipt.Parameters.Clear()
        Me.cmdFeeReceipt.Parameters.AddWithValue("receiptNoAll", receiptNo.Trim)
        reader = Me.cmdFeeReceipt.ExecuteReader
        If reader.HasRows = True Then
            recordExists = True
        ElseIf reader.HasRows = False Then
            recordExists = False
        End If
        reader.Close()
        Return recordExists
    End Function

    Private Sub btnUpdate_Click(sender As Object, e As EventArgs) Handles btnUpdate.Click
        Dim apportionTotal As Double = 0

        If Me.txtRecNo.Text.Trim.Length <= 0 Then
            MsgBox("Receipt Number Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf CDbl(Me.txtOverPayAmount.Text.Trim) > 0 Then
            MsgBox("Receive amount for Next Term in a different receipt.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.txtPayAmount.Text.Trim.Length <= 0 Then
            MsgBox("Pay Amount Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboYear.Text.Trim.Length <= 0 Then
            MsgBox("Year Is Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboTerm.Text.Trim.Length <= 0 Then
            MsgBox("Term Is Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.txtAdmNo.Text.Trim.Length <= 0 Then
            MsgBox("Admission Number Is Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboPayMode.Text.Trim.Length <= 0 Then
            MsgBox("Payment Mode Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboPayAccount.Text.Trim.Length <= 0 Then
            MsgBox("Payment Account Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.txtStudName.Text.Trim.Length <= 0 Then
            MsgBox("Student Name Is Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.txtTermBal.Text.Trim.Length <= 0 Then
            MsgBox("Term Balance Is Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.lstFeeApportDetails.Items.Count <= 0 Then
            MsgBox("No Fee Set For The Student.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If

        For i = 0 To Me.lstFeeApportDetails.Items.Count - 1
            apportionTotal = apportionTotal + Me.lstFeeApportDetails.Items(i).SubItems(4).Text.Trim
        Next

        If ((CDbl(Me.txtPrevYearBal.Text.Trim) - CDbl(Me.txtPayAmount.Text.Trim)) < 0) Then
            If Not ((CDbl(Me.txtPayAmount.Text.Trim) - apportionTotal - CDbl(Me.txtPrevYearBal.Text.Trim)) = 0) Then
                MsgBox("Apportion the fees before saving!", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
                Exit Sub
            End If
        End If

        If Me.cboPaidBy.Text.Trim.Length <= 0 Then
            Dim result As MsgBoxResult = MsgBox("Save Without Paid By Name?", MsgBoxStyle.Question + MsgBoxStyle.ApplicationModal + MsgBoxStyle.YesNo, "Confirm Transaction")
            If result = MsgBoxResult.No Then
                Exit Sub
            End If
        End If

        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()

            Dim recordExists As Boolean = checkIfReceiptNoExists(Me.txtRecNo.Text.Trim)
            If recordExists = True Then
                MsgBox("Receipt Number Exists In The System!", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
                Exit Sub
            End If

            Dim result As MsgBoxResult = MsgBox("Save Record/s?", MsgBoxStyle.Question + MsgBoxStyle.ApplicationModal + MsgBoxStyle.YesNo, "Confirm Transaction")
            If result = MsgBoxResult.No Then
                Exit Sub
            End If

            Me.cmdFeeReceipt.Connection = conn
            Me.cmdFeeReceipt.CommandType = CommandType.StoredProcedure
            Me.cmdFeeReceipt.CommandText = "sprocFinFeeReceipt"
            Me.cmdFeeReceipt.Parameters.Clear()
            Me.cmdFeeReceipt.Parameters.AddWithValue("@receiptNo", Me.txtRecNo.Tag)
            Me.cmdFeeReceipt.Parameters.AddWithValue("@admNo", Me.txtAdmNo.Text.Trim)
            Me.cmdFeeReceipt.Parameters.AddWithValue("@paidBy", Me.cboPaidBy.Text.Trim)
            Me.cmdFeeReceipt.Parameters.AddWithValue("@year", Me.cboYear.Text.Trim)
            Me.cmdFeeReceipt.Parameters.AddWithValue("@termName", Me.cboTerm.Text.Trim)
            Me.cmdFeeReceipt.Parameters.AddWithValue("@receiptNoAll", Me.txtRecNo.Text.Trim)
            Me.cmdFeeReceipt.Parameters.AddWithValue("@actualPayDate", Me.dtpPayDate.Value.Date)
            Me.cmdFeeReceipt.Parameters.AddWithValue("@amountPaid", Me.txtPayAmount.Text.Trim)
            Me.cmdFeeReceipt.Parameters.AddWithValue("@userName", userName.Trim)
            Me.cmdFeeReceipt.Parameters.AddWithValue("@payMode", Me.cboPayMode.Text.Trim)
            Me.cmdFeeReceipt.Parameters.AddWithValue("@payAccount", Me.cboPayAccount.Text.Trim)
            Me.cmdFeeReceipt.Parameters.AddWithValue("@lastYearBal", CDbl(Me.txtPrevYearBal.Text.Trim))
            Me.cmdFeeReceipt.Parameters.AddWithValue("@lastYearBalPaid", CDbl(Me.txtPrevYearBal.Text.Trim) - CDbl(Me.txtPrevYearBalEditable.Text.Trim))
            rec = Me.cmdFeeReceipt.ExecuteNonQuery

            For i = 0 To Me.lstFeeApportDetails.Items.Count - 1
                If Me.lstFeeApportDetails.Items(i).SubItems(4).Text.Trim > 0 Then
                    Me.cmdFeeReceipt.CommandText = "sprocFinFeeReceiptVotes"
                    Me.cmdFeeReceipt.Parameters.Clear()
                    Me.cmdFeeReceipt.Parameters.AddWithValue("@voteHeadName", Me.lstFeeApportDetails.Items(i).Text.Trim)
                    Me.cmdFeeReceipt.Parameters.AddWithValue("@year", Me.cboYear.Text.Trim)
                    Me.cmdFeeReceipt.Parameters.AddWithValue("@termName", Me.cboTerm.Text.Trim)
                    Me.cmdFeeReceipt.Parameters.AddWithValue("@voteHeadAmount", Me.lstFeeApportDetails.Items(i).SubItems(4).Text.Trim)
                    Me.cmdFeeReceipt.Parameters.AddWithValue("@receiptNoAll", Me.txtRecNo.Text.Trim)
                    Me.cmdFeeReceipt.Parameters.AddWithValue("@studFeeId", Me.lstFeeApportDetails.Items(i).Tag)
                    rec = rec + Me.cmdFeeReceipt.ExecuteNonQuery
                End If
            Next

            If rec > 0 Then
                MsgBox("Fee Received Successfully." & vbNewLine & "Wait For the receipt", MsgBoxStyle.Information +
                       MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "System Reply")
            End If

            Me.Cursor = Cursors.WaitCursor

            Dim RptResultsView As New crtFinFeeReceipt
            'Dim ReportPrintOptions As PrintOptions = RptResultsView.PrintOptions
            'Dim ReportPrinter As New PrinterSettings
            'ReportPrintOptions.PrinterName = ReportPrinter.PrinterName
            'ReportPrintOptions.PaperOrientation = PaperOrientation.DefaultPaperOrientation
            'ReportPrintOptions.PaperSize = CrystalDecisions.Shared.PaperSize.PaperA4
            'ReportPrintOptions.PrinterDuplex = PrinterDuplex.Default
            SetReportLogOn(RptResultsView)
            SetReportLogOn(RptResultsView.Subreports("crtFinFeeReceiptVotes"))
            RptResultsView.SummaryInfo.ReportComments = fullName.Trim
            RptResultsView.RecordSelectionFormula = "({vwFinFeeReceipt.receiptNoAll}=" & Chr(34) & Me.txtRecNo.Text.Trim & Chr(34) & ")"
            RptResultsView.RecordSelectionFormula += "AND ({vwFinFeeReceipt.termName}=" & Chr(34) & Me.cboTerm.Text.Trim & Chr(34) & ")"
            frmResultsViewing.crtViewResultsSummary.ReportSource = RptResultsView
            frmResultsViewing.crtViewResultsSummary.Zoom(100)
            frmResultsViewing.crtViewResultsSummary.RefreshReport()
            frmResultsViewing.MdiParent = frmHome
            frmResultsViewing.Show()
            'RptResultsView.PrintToPrinter(1, True, 1, RptResultsView.FormatEngine.GetLastPageNumber(New CrystalDecisions.Shared.ReportPageRequestContext()))
            clearTexts()
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            Me.lstFeeApportDetails.Items.Clear()
            Me.lstStudentDetails.Items.Clear()
            Me.txtOverPayAmount.Text = ""
            dbconnection()
            loadCombos()
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            Me.Cursor = Cursors.Arrow
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub

    Private Sub cboPayMode_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboPayMode.SelectedIndexChanged
        Try
            Me.cboPayAccount.Items.Clear()
            Me.cboPayAccount.Text = ""
            Me.cboPayAccount.SelectedIndex = -1

            Me.cbAutoPortion.Checked = False

            Me.txtPayAmount.Text = ""

            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()

            Me.cmdFeeReceipt.Connection = conn
            Me.cmdFeeReceipt.CommandType = CommandType.Text
            Me.cmdFeeReceipt.CommandText = "SELECT DISTINCT accountNumber FROM tblFinPayAccounts WHERE (modeName=@modeName) " &
                vbNewLine & " ORDER BY accountNumber"
            Me.cmdFeeReceipt.Parameters.Clear()
            Me.cmdFeeReceipt.Parameters.AddWithValue("@modeName", Me.cboPayMode.Text.Trim)
            reader = Me.cmdFeeReceipt.ExecuteReader
            While reader.Read
                Me.cboPayAccount.Items.Add(IIf(DBNull.Value.Equals(reader!accountNumber), "", reader!accountNumber))
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

    Private Sub cboClass_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboClass.SelectedIndexChanged,
        cboYear.SelectedIndexChanged, cboTerm.SelectedIndexChanged, cboStream.SelectedIndexChanged
        Me.lstStudentDetails.Items.Clear()
        Me.txtSearchStudName.Text = ""
        clearTexts()
    End Sub

    Private Sub txtPayAmount_TextChanged(sender As Object, e As EventArgs) Handles txtPayAmount.TextChanged
        Dim overpayment As Double = 0
        Me.cboVoteHead.Text = ""
        Me.cboVoteHead.SelectedIndex = -1
        Me.txtAmount.Text = ""
        Me.txtOverPayAmount.Text = 0

        Me.cboVoteHead.Enabled = True
        Me.txtAmount.Enabled = True
        Me.btnAdd.Enabled = True

        For i = 0 To Me.lstFeeApportDetails.Items.Count - 1
            Me.lstFeeApportDetails.Items(i).SubItems(4).Text = 0
        Next

        Me.cbAutoPortion.Checked = False

        If IsNumeric(Me.txtPayAmount.Text.Trim) = False And Not (Me.txtPayAmount.Text.Trim.Length <= 0) Then
            MsgBox("Non Numeric Values Detected.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Me.txtPayAmount.Text = ""
            Exit Sub
        ElseIf Me.txtPayAmount.Text.Trim.Length <= 0 Then
            Exit Sub
        End If

        Dim termOverPayment As Double = CDbl(Me.txtPayAmount.Text.Trim) - CDbl(Me.txtPrevYearBal.Text.Trim) - CDbl(Me.txtTermBal.Text.Trim)

        If termOverPayment > 0 Then
            MsgBox("Receive overpayment amount for Next Term in a different receipt.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Me.txtFeeBalance.Text = CDbl(Me.txtTermBal.Text.Trim) + CDbl(Me.txtPrevYearBal.Text.Trim)
            Me.txtOverPayAmount.Text = 0
            Me.txtPrevYearBalEditable.Text = Me.txtPrevYearBal.Text
            Me.txtPayAmount.Text = ""
            Exit Sub
        Else
            Me.txtOverPayAmount.Text = 0
        End If
        Me.txtPrevYearBalEditable.Text = If(CDbl(Me.txtPrevYearBal.Text.Trim) - CDbl(Me.txtPayAmount.Text.Trim) <= 0, 0, CDbl(Me.txtPrevYearBal.Text.Trim) - CDbl(Me.txtPayAmount.Text.Trim))
    End Sub

    Private Sub txtAmount_TextChanged(sender As Object, e As EventArgs) Handles txtAmount.TextChanged
        If IsNumeric(Me.txtAmount.Text.Trim) = False And Not (Me.txtAmount.Text.Trim.Length <= 0) Then
            MsgBox("Non Numeric Values Detected.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Me.txtAmount.Text = ""
            Exit Sub
        End If
    End Sub

    Private Sub cbAutoPortion_CheckedChanged(sender As Object, e As EventArgs) Handles cbAutoPortion.CheckedChanged
        For i = 0 To Me.lstFeeApportDetails.Items.Count - 1
            Me.lstFeeApportDetails.Items(i).SubItems(4).Text = 0
        Next
        If Me.cbAutoPortion.CheckState = CheckState.Unchecked Or Me.cbAutoPortion.CheckState = CheckState.Indeterminate Then
            Me.cboVoteHead.Text = ""
            Me.cboVoteHead.SelectedIndex = -1
            Me.txtAmount.Text = ""
            Me.txtPrevYearBalEditable.Text = Me.txtPrevYearBal.Text.Trim
            Me.cboVoteHead.Enabled = True
            Me.txtAmount.Enabled = True
            Me.btnAdd.Enabled = True
            Me.lstFeeApportDetails.Enabled = True
            'Me.txtPrevYearBalEditable.Text = Me.txtPrevYearBal.Text.Trim - Me.txtPayAmount.Text.Trim
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()

            If Me.txtPayAmount.Text.Trim.Length <= 0 Then
                MsgBox("Enter Payment AMount.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
                Me.cbAutoPortion.Checked = False
                Exit Sub
            ElseIf (Me.txtTermBal.Text.Trim.Length <= 0 Or Me.txtTermBal.Text.Trim = "0") And (Me.txtPrevYearBal.Text.Trim = "0" Or
                Me.txtPrevYearBal.Text.Trim.Length <= 0) Then
                MsgBox("The student does not have any balances to pay for the term selected. Select another term.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
                Me.cbAutoPortion.Checked = False
                Exit Sub
            End If
            Me.txtPrevYearBalEditable.Text = ""
            Me.cboVoteHead.Text = ""
            Me.cboVoteHead.SelectedIndex = -1
            Me.txtAmount.Text = ""
            Me.lstFeeApportDetails.Enabled = False

            Me.cboVoteHead.Enabled = False
            Me.txtAmount.Enabled = False
            Me.btnAdd.Enabled = False

            Dim payAmount As Double = CDbl(Me.txtPayAmount.Text.Trim)
            If Me.txtOverPayAmount.Text.Trim <= 0 Then
                payAmount = Me.txtPayAmount.Text.Trim
            Else
                payAmount = Me.txtPayAmount.Text.Trim - Me.txtOverPayAmount.Text.Trim
            End If

            If (payAmount > Me.txtPrevYearBal.Text.Trim) Then
                Me.txtPrevYearBalEditable.Text = 0
                payAmount = payAmount - Me.txtPrevYearBal.Text.Trim
                While payAmount > 0
                    For i = 0 To Me.lstFeeApportDetails.Items.Count - 1
                        If payAmount >= Me.lstFeeApportDetails.Items(i).SubItems(3).Text.Trim Then
                            Me.lstFeeApportDetails.Items(i).SubItems(4).Text = Me.lstFeeApportDetails.Items(i).SubItems(3).Text.Trim
                            payAmount = payAmount - Me.lstFeeApportDetails.Items(i).SubItems(3).Text.Trim
                        Else
                            Me.lstFeeApportDetails.Items(i).SubItems(4).Text = payAmount
                            payAmount = 0
                        End If
                    Next
                End While
            Else
                Me.txtPrevYearBalEditable.Text = CDbl(Me.txtPrevYearBal.Text.Trim) - payAmount
            End If
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub

    Private Sub btnAdd_Click(sender As Object, e As EventArgs) Handles btnAdd.Click
        If Me.cboVoteHead.Text.Trim.Length <= 0 Then
            MsgBox("VoteHead Name Is Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.txtAmount.Text.Trim.Length <= 0 Then
            MsgBox("VoteHead Amount Is Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cbAutoPortion.CheckState = CheckState.Checked Then
            MsgBox("Auto apportioning fees is enabled.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If

        If ((CDbl(Me.txtPrevYearBal.Text.Trim) - CDbl(Me.txtPayAmount.Text.Trim)) >= 0) Then
            MsgBox("Amount being paid will clear the balance first!", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If

        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            For i = 0 To Me.lstFeeApportDetails.Items.Count - 1
                If Me.lstFeeApportDetails.Items(i).Text.Trim = Me.cboVoteHead.Text.Trim Then
                    If CDbl(Me.txtAmount.Text.Trim) > CDbl(Me.lstFeeApportDetails.Items(i).SubItems(3).Text.Trim) Then
                        MsgBox("Amount is Greater than Required.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
                        Exit Sub
                    End If
                    Me.lstFeeApportDetails.Items(i).SubItems(4).Text = Me.txtAmount.Text.Trim
                End If
            Next
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub
End Class