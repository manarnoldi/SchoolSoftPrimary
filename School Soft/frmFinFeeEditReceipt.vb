Imports System.Data.SqlClient
Public Class frmFinFeeEditReceipt
    Dim reader As SqlDataReader
    Dim rec As Integer
    Dim cmdEditReceipt As New SqlCommand
   
    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub frmFinFeeEditReceipt_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
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

    Private Sub frmFinFeeEditReceipt_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint
        If Me.IsMdiChild Then
            Dim PnlLoc As New Point
            PnlLoc.X = CInt((Me.Width - Me.pnlEditReceipt.Width) / 2)
            PnlLoc.Y = CInt((Me.Height - Me.pnlEditReceipt.Height) / 2.5)
            Me.pnlEditReceipt.Location = PnlLoc
        Else
            Me.pnlEditReceipt.Dock = DockStyle.Fill
        End If
    End Sub
    Private Sub loadCombos()
        Me.cboYear.Items.Clear()
        Me.cboYear.Text = ""
        Me.cboYear.SelectedIndex = -1

        Me.cboTermName.Items.Clear()
        Me.cboTermName.Text = ""
        Me.cboTermName.SelectedIndex = -1

        Me.cmdEditReceipt.Connection = conn
        Me.cmdEditReceipt.CommandType = CommandType.Text
        Me.cmdEditReceipt.CommandText = "SELECT DISTINCT year FROM tblSchoolCalendar WHERE (status=1) ORDER BY year"
        Me.cmdEditReceipt.Parameters.Clear()
        reader = Me.cmdEditReceipt.ExecuteReader
        While reader.Read
            Me.cboYear.Items.Add(IIf(DBNull.Value.Equals(reader!year), "", reader!year))
        End While
        reader.Close()

        Me.cmdEditReceipt.CommandText = "SELECT DISTINCT termName FROM tblSchoolCalendar WHERE (status=1) ORDER BY termName"
        Me.cmdEditReceipt.Parameters.Clear()
        reader = Me.cmdEditReceipt.ExecuteReader
        While reader.Read
            Me.cboTermName.Items.Add(IIf(DBNull.Value.Equals(reader!termName), "", reader!termName))
        End While
        reader.Close()
    End Sub

    Private Sub cboTermName_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles _
        cboTermName.SelectedIndexChanged, cboYear.SelectedIndexChanged
        If Me.cboYear.Text.Trim.Length > 0 And Me.cboTermName.Text.Trim.Length > 0 Then

            
            Try
                If conn.State = ConnectionState.Closed Then
                    conn.Open()
                End If
                dbconnection()
                Me.cboMonth.Items.Clear()
                Me.cboMonth.Text = ""
                Me.cboMonth.SelectedIndex = -1

                Me.cboReceiptNo.Items.Clear()
                Me.cboReceiptNo.Text = ""
                Me.cboReceiptNo.SelectedIndex = -1

                Me.cmdEditReceipt.Connection = conn
                Me.cmdEditReceipt.CommandType = CommandType.Text
                Me.cmdEditReceipt.CommandText = "SELECT DISTINCT month FROM vwFinFeeReceipt WHERE (year=@year) AND (termName=@termName) " & _
                    vbNewLine & " ORDER BY month"
                Me.cmdEditReceipt.Parameters.Clear()
                Me.cmdEditReceipt.Parameters.AddWithValue("@year", Me.cboYear.Text.Trim)
                Me.cmdEditReceipt.Parameters.AddWithValue("@termName", Me.cboTermName.Text.Trim)
                reader = Me.cmdEditReceipt.ExecuteReader
                While reader.Read
                    Me.cboMonth.Items.Add(IIf(DBNull.Value.Equals(reader!month), "", reader!month))
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

    Private Sub cboMonth_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboMonth.SelectedIndexChanged
        If Me.cboYear.Text.Trim.Length > 0 And Me.cboTermName.Text.Trim.Length > 0 And Me.cboMonth.Text.Trim.Length > 0 Then
            Try
                If conn.State = ConnectionState.Closed Then
                    conn.Open()
                End If
                dbconnection()
                Me.cboReceiptNo.Items.Clear()
                Me.cboReceiptNo.Text = ""
                Me.cboReceiptNo.SelectedIndex = -1

                Me.cmdEditReceipt.Connection = conn
                Me.cmdEditReceipt.CommandType = CommandType.Text
                Me.cmdEditReceipt.CommandText = "SELECT DISTINCT receiptNoAll FROM vwFinFeeReceipt WHERE (year=@year) AND (termName=@termName) " & _
                    vbNewLine & " AND (month=@month) ORDER BY receiptNoAll"
                Me.cmdEditReceipt.Parameters.Clear()
                Me.cmdEditReceipt.Parameters.AddWithValue("@year", Me.cboYear.Text.Trim)
                Me.cmdEditReceipt.Parameters.AddWithValue("@termName", Me.cboTermName.Text.Trim)
                Me.cmdEditReceipt.Parameters.AddWithValue("@month", Me.cboMonth.Text.Trim)
                reader = Me.cmdEditReceipt.ExecuteReader
                While reader.Read
                    Me.cboReceiptNo.Items.Add(IIf(DBNull.Value.Equals(reader!receiptNoAll), "", reader!receiptNoAll))
                End While
                reader.Close()
            Catch ex As Exception
                MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Finally
                If conn.State = ConnectionState.Open Then
                    conn.Close()
                End If
            End Try
           
        Else
            MsgBox("Year Or Term Or Month Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If
    End Sub

    Private Sub btnLoad_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLoad.Click
        If Me.cboYear.Text.Trim.Length <= 0 Then
            MsgBox("Year is Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboTermName.Text.Trim.Length <= 0 Then
            MsgBox("Term is Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboMonth.Text.Trim.Length <= 0 Then
            MsgBox("Month is Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboReceiptNo.Text.Trim.Length <= 0 Then
            MsgBox("Receipt Number is Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If

        Me.Cursor = Cursors.WaitCursor

        Me.crtViewResultsSummary.ReportSource = Nothing

        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()

            Me.cmdEditReceipt.Connection = conn
            Me.cmdEditReceipt.CommandType = CommandType.Text
            Me.cmdEditReceipt.CommandText = "SELECT * FROM vwFinFeeReceipt WHERE (receiptNoAll=@receiptNoAll) ORDER BY month"
            Me.cmdEditReceipt.Parameters.Clear()
            Me.cmdEditReceipt.Parameters.AddWithValue("@receiptNoAll", Me.cboReceiptNo.Text.Trim)
            reader = Me.cmdEditReceipt.ExecuteReader
            If reader.HasRows = True Then
                Dim RptResultsView As New crtFinFeeReceipt
                SetReportLogOn(RptResultsView)
                SetReportLogOn(RptResultsView.Subreports("crtFinFeeReceiptVotes"))
                RptResultsView.SummaryInfo.ReportComments = fullName.Trim
                RptResultsView.RecordSelectionFormula = "({vwFinFeeReceipt.receiptNoAll}=" & Chr(34) & Me.cboReceiptNo.Text.Trim & Chr(34) & ")"
                RptResultsView.RecordSelectionFormula += "AND ({vwFinFeeReceipt.termName}=" & Chr(34) & Me.cboTermName.Text.Trim & Chr(34) & ")"
                Me.crtViewResultsSummary.ReportSource = RptResultsView
                Me.crtViewResultsSummary.Zoom(100)
                Me.crtViewResultsSummary.RefreshReport()
            ElseIf reader.Read = False Then
                MsgBox("Receipt Not Found in the System!", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            End If
            reader.Close()

            Me.Cursor = Cursors.Arrow
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
        
    End Sub

    Private Sub btnReverse_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnReverse.Click
        If IsNothing(Me.crtViewResultsSummary.ReportSource) = True Then
            MsgBox("View Crystal Report First before Reversing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboReceiptNo.Text.Trim.Length <= 0 Then
            MsgBox("Receipt Number Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.txtReverseReason.Text.Trim.Length <= 0 Then
            MsgBox("Enter Reverse Reason.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            Dim result As MsgBoxResult = MsgBox("Are you sure you want to reverse receipt Number " & _
            Me.cboReceiptNo.Text.Trim & " ?", MsgBoxStyle.Question + MsgBoxStyle.ApplicationModal + MsgBoxStyle.YesNo, "Confirm Transaction")
            If result = MsgBoxResult.No Then
                Exit Sub
            End If
            Me.cmdEditReceipt.Connection = conn
            Me.cmdEditReceipt.CommandType = CommandType.StoredProcedure
            Me.cmdEditReceipt.CommandText = "sprocFinFeeReceiptReverse"
            Me.cmdEditReceipt.Parameters.Clear()
            Me.cmdEditReceipt.Parameters.AddWithValue("@receiptNo", Me.cboReceiptNo.Text.Trim)
            Me.cmdEditReceipt.Parameters.AddWithValue("@doneBy", userName.Trim)
            Me.cmdEditReceipt.Parameters.AddWithValue("@reason", Me.txtReverseReason.Text.Trim)
            rec = Me.cmdEditReceipt.ExecuteNonQuery

            If rec > 0 Then
                MsgBox("Receipt Reversed Successfully.", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "System Reply")
            End If
            Me.crtViewResultsSummary.ReportSource = Nothing
            Me.txtReverseReason.Text = ""
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