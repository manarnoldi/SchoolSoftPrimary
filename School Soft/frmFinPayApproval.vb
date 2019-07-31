Imports System.Data.SqlClient
Public Class frmFinPayApproval
    Dim reader As SqlDataReader
    Dim rec As Integer
    Dim cmdPayAppr As New SqlCommand
    
    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub frmFinPayApproval_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
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
        Me.cboTerm.Items.Clear()
        Me.cboTerm.Text = ""
        Me.cboTerm.SelectedIndex = -1

        Me.cboYear.Items.Clear()
        Me.cboYear.Text = ""
        Me.cboYear.SelectedIndex = -1

        Me.cboEmpNo.Items.Clear()
        Me.cboEmpNo.Text = ""
        Me.cboEmpNo.SelectedIndex = -1

        Me.cmdPayAppr.Connection = conn
        Me.cmdPayAppr.CommandType = CommandType.Text
        Me.cmdPayAppr.CommandText = "SELECT DISTINCT termName FROM tblSchoolCalendar WHERE (status=1) ORDER BY termName"
        Me.cmdPayAppr.Parameters.Clear()
        reader = Me.cmdPayAppr.ExecuteReader
        While reader.Read
            Me.cboTerm.Items.Add(IIf(DBNull.Value.Equals(reader!termName), "", reader!termName))
        End While
        reader.Close()

        Me.cmdPayAppr.CommandText = "SELECT DISTINCT year FROM tblSchoolCalendar WHERE (status=1) ORDER BY year"
        Me.cmdPayAppr.Parameters.Clear()
        reader = Me.cmdPayAppr.ExecuteReader
        While reader.Read
            Me.cboYear.Items.Add(IIf(DBNull.Value.Equals(reader!year), "", reader!year))
        End While
        reader.Close()

        Me.cmdPayAppr.CommandText = "SELECT DISTINCT empNo FROM tblSchoolStaff WHERE (status=1) ORDER BY empNo"
        Me.cmdPayAppr.Parameters.Clear()
        reader = Me.cmdPayAppr.ExecuteReader
        While reader.Read
            Me.cboEmpNo.Items.Add(IIf(DBNull.Value.Equals(reader!empNo), "", reader!empNo))
        End While
        reader.Close()
    End Sub
    Private Sub frmFinPayApproval_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint
        If Me.IsMdiChild Then
            Dim PnlLoc As New Point
            PnlLoc.X = CInt((Me.Width - Me.pnlPayApprove.Width) / 2)
            PnlLoc.Y = CInt((Me.Height - Me.pnlPayApprove.Height) / 2.5)
            Me.pnlPayApprove.Location = PnlLoc
        Else
            Me.pnlPayApprove.Dock = DockStyle.Fill
        End If
    End Sub

    Private Sub cboYear_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboYear.SelectedIndexChanged, cboTerm.SelectedIndexChanged
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            If Me.cboTerm.Text.Trim.Length > 0 And Me.cboYear.Text.Trim.Length > 0 Then

                Me.cboVoucherNo.Items.Clear()
                Me.cboVoucherNo.Text = ""
                Me.cboVoucherNo.SelectedIndex = -1

                Me.cmdPayAppr.Connection = conn
                Me.cmdPayAppr.CommandType = CommandType.Text
                Me.cmdPayAppr.CommandText = "SELECT DISTINCT voucherNoAll FROM vwFinPayments WHERE (approved=0) AND " & _
                    vbNewLine & "(reversed=0) AND (termName=@termName) AND (year=@year) ORDER BY voucherNoAll"
                Me.cmdPayAppr.Parameters.Clear()
                Me.cmdPayAppr.Parameters.AddWithValue("@termName", Me.cboTerm.Text.Trim)
                Me.cmdPayAppr.Parameters.AddWithValue("@year", Me.cboYear.Text.Trim)
                reader = Me.cmdPayAppr.ExecuteReader
                If reader.HasRows = True Then
                    While reader.Read
                        Me.cboVoucherNo.Items.Add(IIf(DBNull.Value.Equals(reader!voucherNoAll), "", reader!voucherNoAll))
                    End While
                ElseIf reader.HasRows = False Then
                    MsgBox("No Voucher Found For the Term and Year.", MsgBoxStyle.Information + _
                       MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "System Reply")
                End If
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

    Private Sub cboEmpNo_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboEmpNo.SelectedIndexChanged
        If Me.cboEmpNo.Text.Trim.Length <= 0 Then
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()

            Me.txtEmpName.Text = ""

            Me.cmdPayAppr.Connection = conn
            Me.cmdPayAppr.CommandType = CommandType.Text
            Me.cmdPayAppr.CommandText = "SELECT DISTINCT FullName FROM tblSchoolStaff WHERE (status=1) AND (empNo=@empNo)"
            Me.cmdPayAppr.Parameters.Clear()
            Me.cmdPayAppr.Parameters.AddWithValue("@empNo", Me.cboEmpNo.Text.Trim)
            reader = Me.cmdPayAppr.ExecuteReader
            While reader.Read
                Me.txtEmpName.Text = IIf(DBNull.Value.Equals(reader!FullName), "", reader!FullName)
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

    Private Sub btnLoad_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLoad.Click
        If Me.cboVoucherNo.Text.Trim.Length <= 0 Then
            MsgBox("Voucher Number Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboEmpNo.Text.Trim.Length <= 0 Then
            MsgBox("Employee Number Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
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
            Me.lstPayApprovals.Items.Clear()

            Me.cmdPayAppr.Connection = conn
            Me.cmdPayAppr.CommandType = CommandType.Text
            Me.cmdPayAppr.CommandText = "SELECT * FROM vwFinPayments WHERE (approved=0) AND (reversed=0) AND (termName=@termName) " & _
                vbNewLine & " AND (year=@year) AND (voucherNoAll=@voucherNoAll) ORDER BY requestDate"
            Me.cmdPayAppr.Parameters.Clear()
            Me.cmdPayAppr.Parameters.AddWithValue("@termName", Me.cboTerm.Text.Trim)
            Me.cmdPayAppr.Parameters.AddWithValue("@year", Me.cboYear.Text.Trim)
            Me.cmdPayAppr.Parameters.AddWithValue("@voucherNoAll", Me.cboVoucherNo.Text.Trim)
            reader = Me.cmdPayAppr.ExecuteReader
            While reader.Read
                Dim dateReq As String = CDate(IIf(DBNull.Value.Equals(reader!requestDate), "", reader!requestDate)).Day.ToString("00") & "-" & _
                CDate(IIf(DBNull.Value.Equals(reader!requestDate), "", reader!requestDate)).Month.ToString("00") & "-" & _
                CDate(IIf(DBNull.Value.Equals(reader!requestDate), "", reader!requestDate)).Year.ToString("0000")

                li = Me.lstPayApprovals.Items.Add(dateReq)
                li.SubItems.Add(IIf(DBNull.Value.Equals(reader!requestBy), "", reader!requestBy))
                li.SubItems.Add(IIf(DBNull.Value.Equals(reader!paidTo), "", reader!paidTo))
                li.SubItems.Add(IIf(DBNull.Value.Equals(reader!expName), "", reader!expName))
                li.SubItems.Add(IIf(DBNull.Value.Equals(reader!accountNumber), "", reader!accountNumber))
                li.SubItems.Add(IIf(DBNull.Value.Equals(reader!payAmount), "", reader!payAmount))
                li.Tag = IIf(DBNull.Value.Equals(reader!payId), "", reader!payId)
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

    Private Sub btnApprove_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnApprove.Click
        If Me.lstPayApprovals.Items.Count <= 0 Then
            MsgBox("No items in the list to save.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.txtEmpName.Text.Trim.Length <= 0 Then
            MsgBox("Employee Name Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()

            Dim result As MsgBoxResult = MsgBox("Approve Voucher?", MsgBoxStyle.Question + MsgBoxStyle.ApplicationModal + MsgBoxStyle.YesNo, "Confirm Transaction")
            If result = MsgBoxResult.No Then
                Exit Sub
            End If

            rec = 0

            For i = 0 To Me.lstPayApprovals.Items.Count - 1
                Me.cmdPayAppr.Connection = conn
                Me.cmdPayAppr.CommandType = CommandType.StoredProcedure
                Me.cmdPayAppr.CommandText = "sprocFinPayApprRev"
                Me.cmdPayAppr.Parameters.Clear()
                Me.cmdPayAppr.Parameters.AddWithValue("@transType", "APPROVE")
                Me.cmdPayAppr.Parameters.AddWithValue("@payId", Me.lstPayApprovals.Items(i).Tag)
                Me.cmdPayAppr.Parameters.AddWithValue("@approvedBy", Me.txtEmpName.Text.Trim)
                rec = rec + Me.cmdPayAppr.ExecuteNonQuery
            Next

            If rec > 0 Then
                MsgBox("Voucher Approved!", MsgBoxStyle.Information + _
                       MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "System Reply")
            End If

            Me.lstPayApprovals.Items.Clear()
            loadCombos()
            Me.txtEmpName.Text = ""
            Me.cboVoucherNo.Items.Clear()
            Me.cboVoucherNo.Text = ""
            Me.cboVoucherNo.SelectedIndex = -1

        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try

    End Sub
End Class