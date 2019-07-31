Imports System.Data.SqlClient
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports System.Drawing.Printing
Public Class frmFinFeeReceipt1
    Dim cmdFeeReceipt As New SqlCommand
    Dim rec As Integer
    Dim reader As SqlDataReader
    Dim receiptInitials As String = ""
    Private Sub frmFeeReceipt_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
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

    Private Sub frmFinFeeReceipt_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint
        If Me.IsMdiChild Then
            Dim PnlLoc As New Point
            PnlLoc.X = CInt((Me.Width - Me.pnlFeeReceipt.Width) / 2)
            PnlLoc.Y = CInt((Me.Height - Me.pnlFeeReceipt.Height) / 2.5)
            Me.pnlFeeReceipt.Location = PnlLoc
        Else
            Me.pnlFeeReceipt.Dock = DockStyle.Fill
        End If
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub cboClass_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles _
        cboClass.SelectedIndexChanged, cboStream.SelectedIndexChanged, cboYear.SelectedIndexChanged
        If Me.cboYear.Text.Trim.Length > 0 And Me.cboClass.Text.Trim.Length > 0 And Me.cboStream.Text.Trim.Length > 0 Then
            Try
                If conn.State = ConnectionState.Closed Then
                    conn.Open()
                End If
                dbconnection()

                Me.txtStudName.Text = ""

                Me.cboAdmNo.Items.Clear()
                Me.cboAdmNo.Text = ""
                Me.cboAdmNo.SelectedIndex = -1

                Me.cmdFeeReceipt.Connection = conn
                Me.cmdFeeReceipt.CommandType = CommandType.Text
                Me.cmdFeeReceipt.CommandText = "SELECT admNo FROM vwStudClass WHERE (className=@className) " &
                    vbNewLine & " AND (stream=@stream) AND (year=@year) ORDER BY admNo"
                Me.cmdFeeReceipt.Parameters.Clear()
                Me.cmdFeeReceipt.Parameters.AddWithValue("@year", Me.cboYear.Text.Trim)
                Me.cmdFeeReceipt.Parameters.AddWithValue("@stream", Me.cboStream.Text.Trim)
                Me.cmdFeeReceipt.Parameters.AddWithValue("@className", Me.cboClass.Text.Trim)
                reader = Me.cmdFeeReceipt.ExecuteReader
                While reader.Read
                    Me.cboAdmNo.Items.Add(IIf(DBNull.Value.Equals(reader!admNo), "", reader!admNo))
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

    Private Sub btnLoadDetails_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
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
        ElseIf Me.cboAdmNo.Text.Trim.Length <= 0 Then
            MsgBox("Admission Number Is Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()

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
            Me.cmdFeeReceipt.CommandText = "SELECT FullName FROM tblStudDetails WHERE (status=1) AND (admNo=@admNo)"
            Me.cmdFeeReceipt.Parameters.Clear()
            Me.cmdFeeReceipt.Parameters.AddWithValue("@admNo", Me.cboAdmNo.Text.Trim)
            reader = Me.cmdFeeReceipt.ExecuteReader
            While reader.Read
                Me.txtStudName.Text = (IIf(DBNull.Value.Equals(reader!FullName), "", reader!FullName))
            End While
            reader.Close()

            Me.cboVoteHead.Items.Clear()
            Me.cboVoteHead.Text = ""
            Me.cboVoteHead.SelectedIndex = -1

            Me.cmdFeeReceipt.CommandText = "SELECT DISTINCT voteHeadName FROM  vwFinFeeStudent WHERE (className=@className) " &
                vbNewLine & " AND (stream=@stream) AND (year=@year) AND (admNo=@admNo) AND (termName=@termName) ORDER BY voteHeadName"
            Me.cmdFeeReceipt.Parameters.Clear()
            Me.cmdFeeReceipt.Parameters.AddWithValue("@year", Me.cboYear.Text.Trim)
            Me.cmdFeeReceipt.Parameters.AddWithValue("@stream", Me.cboStream.Text.Trim)
            Me.cmdFeeReceipt.Parameters.AddWithValue("@className", Me.cboClass.Text.Trim)
            Me.cmdFeeReceipt.Parameters.AddWithValue("@admNo", Me.cboAdmNo.Text.Trim)
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

            Me.txtTotalBal.Text = ""
            Me.txtTermBal.Text = ""
            Me.txtOverPayAmount.Text = ""

            Me.cmdFeeReceipt.CommandText = "SELECT * FROM  vwFinFeeStudBalance WHERE (className=@className) " &
                vbNewLine & " AND (stream=@stream) AND (year=@year) AND (admNo=@admNo) AND (termName=@termName)"
            Me.cmdFeeReceipt.Parameters.Clear()
            Me.cmdFeeReceipt.Parameters.AddWithValue("@year", Me.cboYear.Text.Trim)
            Me.cmdFeeReceipt.Parameters.AddWithValue("@stream", Me.cboStream.Text.Trim)
            Me.cmdFeeReceipt.Parameters.AddWithValue("@className", Me.cboClass.Text.Trim)
            Me.cmdFeeReceipt.Parameters.AddWithValue("@admNo", Me.cboAdmNo.Text.Trim)
            Me.cmdFeeReceipt.Parameters.AddWithValue("@termName", Me.cboTerm.Text.Trim)
            reader = Me.cmdFeeReceipt.ExecuteReader
            While reader.Read
                Me.txtTotalBal.Text = (IIf(DBNull.Value.Equals(reader!totalBalance), "", reader!totalBalance))
                Me.txtTermBal.Text = (IIf(DBNull.Value.Equals(reader!termBalance), "", reader!termBalance))
                Me.txtFeeBalance.Text = (IIf(DBNull.Value.Equals(reader!totalBalance), "", reader!totalBalance))
                Me.txtFeeBalance.Tag = (IIf(DBNull.Value.Equals(reader!totalBalance), "", reader!totalBalance))
                Me.txtOverPayAmount.Text = "0"
            End While
            reader.Close()

            Me.lstPrevPayDetails.Items.Clear()

            Me.cmdFeeReceipt.CommandText = "SELECT * FROM  vwFinFeePrevPayments WHERE (year=@year) AND (admNo=@admNo)"
            Me.cmdFeeReceipt.Parameters.Clear()
            Me.cmdFeeReceipt.Parameters.AddWithValue("@year", Me.cboYear.Text.Trim)
            Me.cmdFeeReceipt.Parameters.AddWithValue("@admNo", Me.cboAdmNo.Text.Trim)
            reader = Me.cmdFeeReceipt.ExecuteReader
            While reader.Read
                Dim actualDate As String = CDate(IIf(DBNull.Value.Equals(reader!actualPayDate), "", reader!actualPayDate)).Day.ToString("00") & " - " &
                CDate(IIf(DBNull.Value.Equals(reader!actualPayDate), "", reader!actualPayDate)).Month.ToString("00") & " - " &
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
            Me.cmdFeeReceipt.CommandText = "SELECT ISNULL(MAX(receiptNo),0)+1 AS newRecNo FROM vwFinFeePrevPayments " &
                vbNewLine & " WHERE (termName=@termName) AND (year=@year)"
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
            Me.cmdFeeReceipt.Parameters.AddWithValue("@admNo", Me.cboAdmNo.Text.Trim)
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
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try

    End Sub

    Private Sub cboAdmNo_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Me.txtStudName.Text = ""
        Me.txtRecNo.Text = ""
        Me.txtTotalBal.Text = ""
        Me.txtTermBal.Text = ""
        Me.txtOverPayAmount.Text = ""
        Me.txtPayAmount.Text = ""
        Me.txtAmount.Text = ""
        Me.txtFeeBalance.Text = ""

        Me.cboVoteHead.Items.Clear()
        Me.cboVoteHead.Text = ""
        Me.cboVoteHead.SelectedIndex = -1

        Me.cboPayAccount.Items.Clear()
        Me.cboPayAccount.Text = ""
        Me.cboPayAccount.SelectedIndex = -1

        Me.lstPrevPayDetails.Items.Clear()
        Me.lstFeeApportDetails.Items.Clear()

        Me.cbAutoPortion.Checked = False

        Me.cboVoteHead.Enabled = True
        Me.txtAmount.Enabled = True
        Me.btnAdd.Enabled = True

        Me.lstFeeApportDetails.Enabled = True
    End Sub

    Private Sub cboTerm_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboTerm.SelectedIndexChanged
        Me.txtStudName.Text = ""
        Me.txtRecNo.Text = ""
        Me.txtTotalBal.Text = ""
        Me.txtTermBal.Text = ""
        Me.txtOverPayAmount.Text = ""
        Me.txtPayAmount.Text = ""
        Me.txtAmount.Text = ""
        Me.txtFeeBalance.Text = ""

        Me.cboVoteHead.Items.Clear()
        Me.cboVoteHead.Text = ""
        Me.cboVoteHead.SelectedIndex = -1

        Me.cboPayAccount.Items.Clear()
        Me.cboPayAccount.Text = ""
        Me.cboPayAccount.SelectedIndex = -1

        Me.lstPrevPayDetails.Items.Clear()
        Me.lstFeeApportDetails.Items.Clear()

        Me.cbAutoPortion.Checked = False

        Me.cboVoteHead.Enabled = True
        Me.txtAmount.Enabled = True
        Me.btnAdd.Enabled = True

        Me.lstFeeApportDetails.Enabled = True
    End Sub

    Private Sub cboPayMode_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboPayMode.SelectedIndexChanged
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

    Private Sub txtPayAmount_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtPayAmount.TextChanged
        If IsNumeric(Me.txtPayAmount.Text.Trim) = False And Not (Me.txtPayAmount.Text.Trim.Length <= 0) Then
            MsgBox("Non Numeric Values Detected.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Me.txtPayAmount.Text = ""
            Exit Sub
        End If

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
        If Me.txtPayAmount.Text.Trim.Length <= 0 Then
            Me.txtFeeBalance.Text = (Me.txtFeeBalance.Tag)
            Exit Sub
        End If

        Me.txtFeeBalance.Text = (Me.txtFeeBalance.Tag) - Me.txtPayAmount.Text.Trim

        If (Me.txtTermBal.Text.Trim) - (Me.txtPayAmount.Text.Trim) < 0 Then
            Me.txtOverPayAmount.Text = ((Me.txtTermBal.Text.Trim) - (Me.txtPayAmount.Text.Trim)) * -1
        Else
            Me.txtOverPayAmount.Text = 0
        End If
    End Sub

    Private Sub txtAmount_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtAmount.TextChanged
        If IsNumeric(Me.txtAmount.Text.Trim) = False And Not (Me.txtAmount.Text.Trim.Length <= 0) Then
            MsgBox("Non Numeric Values Detected.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Me.txtAmount.Text = ""
            Exit Sub
        End If
    End Sub

    Private Sub cbAutoPortion_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbAutoPortion.CheckedChanged

        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            For i = 0 To Me.lstFeeApportDetails.Items.Count - 1
                Me.lstFeeApportDetails.Items(i).SubItems(4).Text = 0
            Next
            If Me.cbAutoPortion.CheckState = CheckState.Checked Then

                If Me.txtPayAmount.Text.Trim.Length <= 0 Then
                    MsgBox("Enter Payment AMount.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
                    Exit Sub
                ElseIf Me.txtTermBal.Text.Trim.Length <= 0 Or Me.txtTermBal.Text.Trim = "0" Then
                    MsgBox("Term Balance Is Zero Or Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
                    Exit Sub
                End If

                Me.cboVoteHead.Text = ""
                Me.cboVoteHead.SelectedIndex = -1
                Me.txtAmount.Text = ""
                Me.lstFeeApportDetails.Enabled = False

                Me.cboVoteHead.Enabled = False
                Me.txtAmount.Enabled = False
                Me.btnAdd.Enabled = False

                Dim payAmount As Double = Me.txtPayAmount.Text.Trim
                If Me.txtOverPayAmount.Text.Trim <= 0 Then
                    payAmount = Me.txtPayAmount.Text.Trim
                Else
                    payAmount = Me.txtPayAmount.Text.Trim - Me.txtOverPayAmount.Text.Trim
                End If
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

            ElseIf Me.cbAutoPortion.CheckState = CheckState.Unchecked Then

                Me.cboVoteHead.Text = ""
                Me.cboVoteHead.SelectedIndex = -1
                Me.txtAmount.Text = ""

                Me.cboVoteHead.Enabled = True
                Me.txtAmount.Enabled = True
                Me.btnAdd.Enabled = True
                Me.lstFeeApportDetails.Enabled = True
            End If
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub

    Private Sub cboPayAccount_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboPayAccount.SelectedIndexChanged
        Me.txtPayAmount.Text = ""
    End Sub

    Private Sub btnAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAdd.Click
        If Me.cboVoteHead.Text.Trim.Length <= 0 Then
            MsgBox("VoteHead Name Is Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.txtAmount.Text.Trim.Length <= 0 Then
            MsgBox("VoteHead Amount Is Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            For i = 0 To Me.lstFeeApportDetails.Items.Count - 1
                If Me.lstFeeApportDetails.Items(i).Text.Trim = Me.cboVoteHead.Text.Trim Then
                    If Me.txtAmount.Text.Trim > Me.lstFeeApportDetails.Items(i).SubItems(3).Text Then
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

    Private Sub btnUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdate.Click
        Dim apportionTotal As Double = 0
        Dim feeBalance As Double = 0

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
        ElseIf Me.cboAdmNo.Text.Trim.Length <= 0 Then
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
            MsgBox("Term Balance Is Zero Or Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.lstFeeApportDetails.Items.Count <= 0 Then
            MsgBox("No Fee Set For The Student.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If

        If Me.cboPaidBy.Text.Trim.Length <= 0 Then
            Dim result As MsgBoxResult = MsgBox("Save Without Paid By Name?", MsgBoxStyle.Question + MsgBoxStyle.ApplicationModal + MsgBoxStyle.YesNo, "Confirm Transaction")
            If result = MsgBoxResult.No Then
                Exit Sub
            End If
        End If

        For i = 0 To Me.lstFeeApportDetails.Items.Count - 1
            apportionTotal = apportionTotal + Me.lstFeeApportDetails.Items(i).SubItems(4).Text.Trim
            feeBalance = feeBalance + Me.lstFeeApportDetails.Items(i).SubItems(4).Text.Trim
        Next
        If feeBalance = apportionTotal Then
        Else
            If apportionTotal <> Me.txtPayAmount.Text.Trim Then
                MsgBox("Apportion Amount is not Equal to Pay Amount.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
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
            Me.cmdFeeReceipt.Parameters.AddWithValue("@admNo", Me.cboAdmNo.Text.Trim)
            Me.cmdFeeReceipt.Parameters.AddWithValue("@paidBy", Me.cboPaidBy.Text.Trim)
            Me.cmdFeeReceipt.Parameters.AddWithValue("@year", Me.cboYear.Text.Trim)
            Me.cmdFeeReceipt.Parameters.AddWithValue("@termName", Me.cboTerm.Text.Trim)
            Me.cmdFeeReceipt.Parameters.AddWithValue("@receiptNoAll", Me.txtRecNo.Text.Trim)
            Me.cmdFeeReceipt.Parameters.AddWithValue("@actualPayDate", Me.dtpPayDate.Value.Date)
            Me.cmdFeeReceipt.Parameters.AddWithValue("@amountPaid", Me.txtPayAmount.Text.Trim)
            Me.cmdFeeReceipt.Parameters.AddWithValue("@userName", userName.Trim)
            Me.cmdFeeReceipt.Parameters.AddWithValue("@payMode", Me.cboPayMode.Text.Trim)
            Me.cmdFeeReceipt.Parameters.AddWithValue("@payAccount", Me.cboPayAccount.Text.Trim)
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
            RptResultsView.SummaryInfo.ReportComments = fullName.Trim
            RptResultsView.RecordSelectionFormula = "({vwFinFeeReceipt.receiptNoAll}=" & Chr(34) & Me.txtRecNo.Text.Trim & Chr(34) & ")"
            RptResultsView.RecordSelectionFormula += "AND ({vwFinFeeReceipt.termName}=" & Chr(34) & Me.cboTerm.Text.Trim & Chr(34) & ")"
            frmResultsViewing.crtViewResultsSummary.ReportSource = RptResultsView
            frmResultsViewing.crtViewResultsSummary.Zoom(100)
            frmResultsViewing.crtViewResultsSummary.RefreshReport()
            frmResultsViewing.MdiParent = frmHome
            frmResultsViewing.Show()
            'RptResultsView.PrintToPrinter(1, True, 1, RptResultsView.FormatEngine.GetLastPageNumber(New CrystalDecisions.Shared.ReportPageRequestContext()))
            loadCombos()
            clearTexts()
            Me.Cursor = Cursors.Arrow
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub
    Private Sub clearTexts()
        Me.txtStudName.Text = ""
        Me.txtRecNo.Text = ""
        Me.txtTotalBal.Text = ""
        Me.txtTermBal.Text = ""
        Me.txtOverPayAmount.Text = ""
        Me.txtPayAmount.Text = ""
        Me.txtAmount.Text = ""
        Me.txtFeeBalance.Text = ""
        Me.cboVoteHead.Items.Clear()
        Me.cboVoteHead.Text = ""
        Me.cboVoteHead.SelectedIndex = -1

        Me.cboAdmNo.Items.Clear()
        Me.cboAdmNo.Text = ""
        Me.cboAdmNo.SelectedIndex = -1

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

    Private Sub cboSearchStudent_SelectedIndexChanged(sender As Object, e As EventArgs)
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
            Me.cmdFeeReceipt.CommandText = "SELECT FullName FROM tblStudDetails WHERE (status=1) AND (admNo=@admNo)"
            Me.cmdFeeReceipt.Parameters.Clear()
            Me.cmdFeeReceipt.Parameters.AddWithValue("@admNo", Me.cboAdmNo.Text.Trim)
            reader = Me.cmdFeeReceipt.ExecuteReader
            While reader.Read
                Me.txtStudName.Text = (IIf(DBNull.Value.Equals(reader!FullName), "", reader!FullName))
            End While
            reader.Close()

            Me.cboVoteHead.Items.Clear()
            Me.cboVoteHead.Text = ""
            Me.cboVoteHead.SelectedIndex = -1

            Me.cmdFeeReceipt.CommandText = "SELECT DISTINCT voteHeadName FROM  vwFinFeeStudent WHERE (className=@className) " &
                vbNewLine & " AND (stream=@stream) AND (year=@year) AND (admNo=@admNo) AND (termName=@termName) ORDER BY voteHeadName"
            Me.cmdFeeReceipt.Parameters.Clear()
            Me.cmdFeeReceipt.Parameters.AddWithValue("@year", Me.cboYear.Text.Trim)
            Me.cmdFeeReceipt.Parameters.AddWithValue("@stream", Me.cboStream.Text.Trim)
            Me.cmdFeeReceipt.Parameters.AddWithValue("@className", Me.cboClass.Text.Trim)
            Me.cmdFeeReceipt.Parameters.AddWithValue("@admNo", Me.cboAdmNo.Text.Trim)
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

            Me.txtTotalBal.Text = ""
            Me.txtTermBal.Text = ""
            Me.txtOverPayAmount.Text = ""

            Me.cmdFeeReceipt.CommandText = "SELECT * FROM  vwFinFeeStudBalance WHERE (className=@className) " &
                vbNewLine & " AND (stream=@stream) AND (year=@year) AND (admNo=@admNo) AND (termName=@termName)"
            Me.cmdFeeReceipt.Parameters.Clear()
            Me.cmdFeeReceipt.Parameters.AddWithValue("@year", Me.cboYear.Text.Trim)
            Me.cmdFeeReceipt.Parameters.AddWithValue("@stream", Me.cboStream.Text.Trim)
            Me.cmdFeeReceipt.Parameters.AddWithValue("@className", Me.cboClass.Text.Trim)
            Me.cmdFeeReceipt.Parameters.AddWithValue("@admNo", Me.cboAdmNo.Text.Trim)
            Me.cmdFeeReceipt.Parameters.AddWithValue("@termName", Me.cboTerm.Text.Trim)
            reader = Me.cmdFeeReceipt.ExecuteReader
            While reader.Read
                Me.txtTotalBal.Text = (IIf(DBNull.Value.Equals(reader!totalBalance), "", reader!totalBalance))
                Me.txtTermBal.Text = (IIf(DBNull.Value.Equals(reader!termBalance), "", reader!termBalance))
                Me.txtFeeBalance.Text = (IIf(DBNull.Value.Equals(reader!totalBalance), "", reader!totalBalance))
                Me.txtFeeBalance.Tag = (IIf(DBNull.Value.Equals(reader!totalBalance), "", reader!totalBalance))
                Me.txtOverPayAmount.Text = "0"
            End While
            reader.Close()

            Me.lstPrevPayDetails.Items.Clear()

            Me.cmdFeeReceipt.CommandText = "SELECT * FROM  vwFinFeePrevPayments WHERE (year=@year) AND (admNo=@admNo)"
            Me.cmdFeeReceipt.Parameters.Clear()
            Me.cmdFeeReceipt.Parameters.AddWithValue("@year", Me.cboYear.Text.Trim)
            Me.cmdFeeReceipt.Parameters.AddWithValue("@admNo", Me.cboAdmNo.Text.Trim)
            reader = Me.cmdFeeReceipt.ExecuteReader
            While reader.Read
                Dim actualDate As String = CDate(IIf(DBNull.Value.Equals(reader!actualPayDate), "", reader!actualPayDate)).Day.ToString("00") & " - " &
                CDate(IIf(DBNull.Value.Equals(reader!actualPayDate), "", reader!actualPayDate)).Month.ToString("00") & " - " &
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
            Me.cmdFeeReceipt.CommandText = "SELECT ISNULL(MAX(receiptNo),0)+1 AS newRecNo FROM vwFinFeePrevPayments " &
                vbNewLine & " WHERE (termName=@termName) AND (year=@year)"
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
            Me.cmdFeeReceipt.Parameters.AddWithValue("@admNo", Me.cboAdmNo.Text.Trim)
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
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub

    Private Sub frmFinFeeReceipt_MouseEnter(sender As Object, e As EventArgs) Handles Me.MouseEnter

    End Sub
End Class
