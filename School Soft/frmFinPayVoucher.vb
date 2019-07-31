Imports System.Data.SqlClient
Public Class frmFinPayVoucher
    Dim reader As SqlDataReader
    Dim rec As Integer
    Dim cmdPayVoucher As New SqlCommand

    Private Sub frmFinPayVoucher_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
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

        Me.cmdPayVoucher.Connection = conn
        Me.cmdPayVoucher.CommandType = CommandType.Text
        Me.cmdPayVoucher.CommandText = "SELECT DISTINCT termName FROM tblSchoolCalendar WHERE (status=1) ORDER BY termName"
        Me.cmdPayVoucher.Parameters.Clear()
        reader = Me.cmdPayVoucher.ExecuteReader
        While reader.Read
            Me.cboTerm.Items.Add(IIf(DBNull.Value.Equals(reader!termName), "", reader!termName))
        End While
        reader.Close()

        Me.cmdPayVoucher.CommandText = "SELECT DISTINCT year FROM tblSchoolCalendar WHERE (status=1) ORDER BY year"
        Me.cmdPayVoucher.Parameters.Clear()
        reader = Me.cmdPayVoucher.ExecuteReader
        While reader.Read
            Me.cboYear.Items.Add(IIf(DBNull.Value.Equals(reader!year), "", reader!year))
        End While
        reader.Close()
    End Sub
    Private Sub frmFinPayVoucher_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint
        If Me.IsMdiChild Then
            Dim PnlLoc As New Point
            PnlLoc.X = CInt((Me.Width - Me.pnlPayVoucher.Width) / 2)
            PnlLoc.Y = CInt((Me.Height - Me.pnlPayVoucher.Height) / 2.5)
            Me.pnlPayVoucher.Location = PnlLoc
        Else
            Me.pnlPayVoucher.Dock = DockStyle.Fill
        End If
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
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

                Me.cmdPayVoucher.Connection = conn
                Me.cmdPayVoucher.CommandType = CommandType.Text
                Me.cmdPayVoucher.CommandText = "SELECT DISTINCT voucherNoAll FROM vwFinPayments WHERE (approved=1) AND " & _
                    vbNewLine & "(reversed=0) AND (termName=@termName) AND (year=@year) ORDER BY voucherNoAll"
                Me.cmdPayVoucher.Parameters.Clear()
                Me.cmdPayVoucher.Parameters.AddWithValue("@termName", Me.cboTerm.Text.Trim)
                Me.cmdPayVoucher.Parameters.AddWithValue("@year", Me.cboYear.Text.Trim)
                reader = Me.cmdPayVoucher.ExecuteReader
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

    Private Sub btnLoad_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLoad.Click
        If Me.cboTerm.Text.Trim.Length <= 0 Then
            MsgBox("Term Name is Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboVoucherNo.Text.Trim.Length <= 0 Then
            MsgBox("Voucher Number Is Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboYear.Text.Trim.Length <= 0 Then
            MsgBox("Year Is Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()

            Me.crtVwPayVoucher.ReportSource = Nothing

            Me.Cursor = Cursors.WaitCursor

            Dim RptResultsView As New crtFinPayVoucher
            SetReportLogOn(RptResultsView)
            RptResultsView.SummaryInfo.ReportComments = fullName.Trim
            RptResultsView.RecordSelectionFormula = "({vwFinPayments.voucherNoAll}=" & Chr(34) & Me.cboVoucherNo.Text.Trim & Chr(34) & ")"
            RptResultsView.RecordSelectionFormula += "AND ({vwFinPayments.termName}=" & Chr(34) & Me.cboTerm.Text.Trim & Chr(34) & ")"
            RptResultsView.RecordSelectionFormula += "AND ({vwFinPayments.year}=" & Me.cboYear.Text.Trim & ")"
            RptResultsView.RecordSelectionFormula += "AND ({vwFinPayments.reversed}=False)"
            RptResultsView.RecordSelectionFormula += "AND ({vwFinPayments.approved}=True)"

            Me.crtVwPayVoucher.ReportSource = RptResultsView
            Me.crtVwPayVoucher.Zoom(100)
            Me.crtVwPayVoucher.RefreshReport()

            loadCombos()
            Me.Cursor = Cursors.Arrow
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub
End Class