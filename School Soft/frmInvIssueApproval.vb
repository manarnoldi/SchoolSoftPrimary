Imports System.Data.SqlClient
Public Class frmInvIssueApproval
    Dim reader As SqlDataReader
    Dim rec As Integer
    Dim cmdIssueApp As New SqlCommand
    Private Sub frmInvIssueApproval_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
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
        Me.cboApprNo.Items.Clear()
        Me.cboApprNo.Text = ""
        Me.cboApprNo.SelectedIndex = -1

        Me.txtApprName.Clear()

        Me.cmdIssueApp.Connection = conn
        Me.cmdIssueApp.CommandType = CommandType.Text
        Me.cmdIssueApp.CommandText = "SELECT DISTINCT empNo FROM tblSchoolStaff WHERE (status=1)"
        Me.cmdIssueApp.Parameters.Clear()
        reader = Me.cmdIssueApp.ExecuteReader
        While reader.Read
            Me.cboApprNo.Items.Add(IIf(DBNull.Value.Equals(reader!empNo), "", reader!empNo))
        End While
        reader.Close()
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub frmInvIssueApproval_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint
        If Me.IsMdiChild Then
            Dim PnlLoc As New Point
            PnlLoc.X = CInt((Me.Width - Me.pnlIssueApproval.Width) / 2)
            PnlLoc.Y = CInt((Me.Height - Me.pnlIssueApproval.Height) / 2.5)
            Me.pnlIssueApproval.Location = PnlLoc
        Else
            Me.pnlIssueApproval.Dock = DockStyle.Fill
        End If
    End Sub

    Private Sub cboApprNo_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboApprNo.SelectedIndexChanged
        If Me.cboApprNo.Text.Trim.Length <= 0 Then
            MsgBox("Approver Number Missing.", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "System Reply")
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            Me.txtApprName.Clear()

            Me.cmdIssueApp.Connection = conn
            Me.cmdIssueApp.CommandType = CommandType.Text
            Me.cmdIssueApp.CommandText = "SELECT DISTINCT FullName FROM tblSchoolStaff WHERE (status=1) AND (empNo=@empNo)"
            Me.cmdIssueApp.Parameters.Clear()
            Me.cmdIssueApp.Parameters.AddWithValue("@empNo", Me.cboApprNo.Text.Trim)
            reader = Me.cmdIssueApp.ExecuteReader
            While reader.Read
                Me.txtApprName.Text = (IIf(DBNull.Value.Equals(reader!FullName), "", reader!FullName))
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
        Me.dgVwIssueAppr.Rows.Clear()
        Dim dateDifference As Integer = DateDiff(DateInterval.Day, Me.dtpReqFrom.Value.Date, Me.dtpReqTo.Value.Date)
        If Me.cboApprNo.Text.Trim.Length <= 0 Then
            MsgBox("Approver Number Missing.", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "System Reply")
            Exit Sub
        ElseIf Me.txtApprName.Text.Trim.Length <= 0 Then
            MsgBox("Approver Name Missing.", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "System Reply")
            Exit Sub
        ElseIf dateDifference < 0 Then
            MsgBox("First Date Should be less than second Date.", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "System Reply")
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            Me.cmdIssueApp.Connection = conn
            Me.cmdIssueApp.CommandType = CommandType.Text
            Me.cmdIssueApp.CommandText = "SELECT * FROM vwInvIssueRequest WHERE (issueReqDate BETWEEN @fromDate AND @toDate) AND (noteStatus=0)"
            Me.cmdIssueApp.Parameters.Clear()
            Me.cmdIssueApp.Parameters.AddWithValue("@fromDate", Me.dtpReqFrom.Value.Date)
            Me.cmdIssueApp.Parameters.AddWithValue("@toDate", Me.dtpReqTo.Value.Date)
            reader = Me.cmdIssueApp.ExecuteReader()
            If reader.HasRows = True Then
                While reader.Read
                    Dim rowNum As Integer = Me.dgVwIssueAppr.Rows.Count
                    Dim row As String() = New String() {IIf(DBNull.Value.Equals(reader!issueNumber), "", (reader!issueNumber)), _
                                    IIf(DBNull.Value.Equals(reader!itemDescription), "", (reader!itemDescription)), _
                                    IIf(DBNull.Value.Equals(reader!uom), "", (reader!uom)), _
                                    IIf(DBNull.Value.Equals(reader!issueQty), "", (reader!issueQty)), _
                                    IIf(DBNull.Value.Equals(reader!costPerUnit), "", (reader!costPerUnit)), _
                                    IIf(DBNull.Value.Equals(reader!totalCost), "", (reader!totalCost)), _
                                    IIf(DBNull.Value.Equals(reader!doneBy), "", (reader!doneBy)), _
                                    IIf(DBNull.Value.Equals(reader!issueToName), "", (reader!issueToName))}
                    Me.dgVwIssueAppr.Rows.Add(row)
                    Me.dgVwIssueAppr.Rows(rowNum).Tag = IIf(DBNull.Value.Equals(reader!issueId), "", (reader!issueId))
                    Me.dgVwIssueAppr.Rows(rowNum).Cells(1).Tag = IIf(DBNull.Value.Equals(reader!itemCode), "", (reader!itemCode))
                End While
            ElseIf reader.HasRows = False Then
                MsgBox("No record Found For the Selection Done.", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "System Reply")
            End If
            reader.Close()
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub

    Private Sub dgVwIssueAppr_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgVwIssueAppr.CellContentClick
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            Dim cellValue As Boolean = False
            If e.ColumnIndex = 8 Then
                    Me.dgVwIssueAppr.Item(9, e.RowIndex).Value = False
            ElseIf e.ColumnIndex = 9 Then
                    Me.dgVwIssueAppr.Item(8, e.RowIndex).Value = False
            End If
        Catch ex As Exception

        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub

    Private Sub btnUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdate.Click
        If Me.dgVwIssueAppr.RowCount <= 0 Then
            MsgBox("Missing items in the gridview.", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "System Reply")
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()

            Dim result As MsgBoxResult = MsgBox("Submit?", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.YesNo, "Confirm Transaction")
            If result = MsgBoxResult.No Then
                Exit Sub
            End If

            i = 0

            For i = 0 To Me.dgVwIssueAppr.RowCount - 1
                Dim status As Boolean = False
                Dim apprStatus As Boolean = False

                If Me.dgVwIssueAppr.Rows(i).Cells(8).Value = True Or Me.dgVwIssueAppr.Rows(i).Cells(9).Value = True Then
                    status = True
                End If
                If status = True Then
                    If Me.dgVwIssueAppr.Rows(i).Cells(8).Value = True Then
                        apprStatus = 1
                    ElseIf Me.dgVwIssueAppr.Rows(i).Cells(9).Value = True Then
                        apprStatus = 0
                    End If

                    Me.cmdIssueApp.Connection = conn
                    Me.cmdIssueApp.CommandType = CommandType.StoredProcedure
                    Me.cmdIssueApp.CommandText = "sprocInvIssueApprovals"
                    Me.cmdIssueApp.Parameters.Clear()
                    Me.cmdIssueApp.Parameters.AddWithValue("@issueId", Me.dgVwIssueAppr.Rows(i).Tag)
                    Me.cmdIssueApp.Parameters.AddWithValue("@itemCode", Me.dgVwIssueAppr.Rows(i).Cells(1).Tag)
                    Me.cmdIssueApp.Parameters.AddWithValue("@doneBy", userName.Trim)
                    Me.cmdIssueApp.Parameters.AddWithValue("@approved", apprStatus)
                    Me.cmdIssueApp.Parameters.AddWithValue("@issueQty", Me.dgVwIssueAppr.Rows(i).Cells(3).Value)
                    rec = rec + Me.cmdIssueApp.ExecuteNonQuery
                End If
            Next
            If rec > 0 Then
                MsgBox("Appovals Submitted!", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "SuccessFull Transaction")
            End If
            Me.dgVwIssueAppr.Rows.Clear()
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub
End Class