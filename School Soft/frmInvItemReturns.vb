Imports System.Data.SqlClient
Public Class frmInvItemReturns
    Dim cmdIssueApp As New SqlCommand
    Dim reader As SqlDataReader
    Dim rec As Integer = 0
    Private Sub frmInvItemReturns_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub

    Private Sub btnLoad_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLoad.Click
        Me.dgVwItemReturn.Rows.Clear()
        Dim dateDifference As Integer = DateDiff(DateInterval.Day, Me.dtpDateFrom.Value.Date, Me.dtpDateTo.Value.Date)
        If dateDifference < 0 Then
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
            Me.cmdIssueApp.CommandText = "SELECT * FROM vwInvPendingReturn WHERE (issueReqDate BETWEEN @fromDate " & _
                vbNewLine & "AND @toDate) AND (pendingQty>0) ORDER BY issueReqDate"
            Me.cmdIssueApp.Parameters.Clear()
            Me.cmdIssueApp.Parameters.AddWithValue("@fromDate", Me.dtpDateFrom.Value.Date)
            Me.cmdIssueApp.Parameters.AddWithValue("@toDate", Me.dtpDateFrom.Value.Date)
            reader = Me.cmdIssueApp.ExecuteReader()
            If reader.HasRows = True Then
                While reader.Read
                    Dim issueReqDate As String = CDate(IIf(DBNull.Value.Equals(reader!issueReqDate), "", (reader!issueReqDate))).Day.ToString("00") & "-" & _
                        CDate(IIf(DBNull.Value.Equals(reader!issueReqDate), "", (reader!issueReqDate))).Month.ToString("00") & "-" & _
                        CDate(IIf(DBNull.Value.Equals(reader!issueReqDate), "", (reader!issueReqDate))).Year.ToString("0000")

                    Dim rowNum As Integer = Me.dgVwItemReturn.Rows.Count
                    Dim row As String() = New String() {IIf(DBNull.Value.Equals(reader!itemDescription), "", (reader!itemDescription)), _
                                    IIf(DBNull.Value.Equals(reader!issuedToTypeName), "", (reader!issuedToTypeName)), _
                                    IIf(DBNull.Value.Equals(reader!issueToName), "", (reader!issueToName)), _
                                    IIf(DBNull.Value.Equals(reader!uom), "", (reader!uom)), _
                                    issueReqDate, _
                                    IIf(DBNull.Value.Equals(reader!issueQty), "", (reader!issueQty)), _
                                    IIf(DBNull.Value.Equals(reader!returnQty), "", (reader!returnQty)), _
                                    IIf(DBNull.Value.Equals(reader!pendingQty), "", (reader!pendingQty)), _
                                    IIf(DBNull.Value.Equals(reader!pendingCost), "", (reader!pendingCost))}
                    Me.dgVwItemReturn.Rows.Add(row)
                    Me.dgVwItemReturn.Rows(rowNum).Tag = IIf(DBNull.Value.Equals(reader!issueId), "", (reader!issueId))
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

    Private Sub frmInvItemReturns_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint
        If Me.IsMdiChild Then
            Dim PnlLoc As New Point
            PnlLoc.X = CInt((Me.Width - Me.pnlItemReturns.Width) / 2)
            PnlLoc.Y = CInt((Me.Height - Me.pnlItemReturns.Height) / 2.5)
            Me.pnlItemReturns.Location = PnlLoc
        Else
            Me.pnlItemReturns.Dock = DockStyle.Fill
        End If
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub txtSearchName_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtSearchName.TextChanged
        If Me.txtSearchName.Text.Trim.Length <= 2 Then
            Me.dgVwItemReturn.Rows.Clear()
            Exit Sub
        End If
        Me.dgVwItemReturn.Rows.Clear()
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            Me.cmdIssueApp.Connection = conn
            Me.cmdIssueApp.CommandType = CommandType.Text
            Me.cmdIssueApp.CommandText = "SELECT * FROM vwInvPendingReturn WHERE (issueToName LIKE @issueToName) AND (pendingQty>0) ORDER BY issueReqDate"
            Me.cmdIssueApp.Parameters.Clear()
            Me.cmdIssueApp.Parameters.AddWithValue("@issueToName", String.Format("%{0}%", TryCast(Me.txtSearchName.Text.Trim, String).Trim))
            reader = Me.cmdIssueApp.ExecuteReader()
            If reader.HasRows = True Then
                While reader.Read
                    Dim issueReqDate As String = CDate(IIf(DBNull.Value.Equals(reader!issueReqDate), "", (reader!issueReqDate))).Day.ToString("00") & "-" & _
                        CDate(IIf(DBNull.Value.Equals(reader!issueReqDate), "", (reader!issueReqDate))).Month.ToString("00") & "-" & _
                        CDate(IIf(DBNull.Value.Equals(reader!issueReqDate), "", (reader!issueReqDate))).Year.ToString("0000")

                    Dim rowNum As Integer = Me.dgVwItemReturn.Rows.Count
                    Dim row As String() = New String() {IIf(DBNull.Value.Equals(reader!itemDescription), "", (reader!itemDescription)), _
                                    IIf(DBNull.Value.Equals(reader!issuedToTypeName), "", (reader!issuedToTypeName)), _
                                    IIf(DBNull.Value.Equals(reader!issueToName), "", (reader!issueToName)), _
                                    IIf(DBNull.Value.Equals(reader!uom), "", (reader!uom)), _
                                    issueReqDate, _
                                    IIf(DBNull.Value.Equals(reader!issueQty), "", (reader!issueQty)), _
                                    IIf(DBNull.Value.Equals(reader!returnQty), "", (reader!returnQty)), _
                                    IIf(DBNull.Value.Equals(reader!pendingQty), "", (reader!pendingQty)), _
                                    IIf(DBNull.Value.Equals(reader!pendingCost), "", (reader!pendingCost))}
                    Me.dgVwItemReturn.Rows.Add(row)
                    Me.dgVwItemReturn.Rows(rowNum).Tag = IIf(DBNull.Value.Equals(reader!issueId), "", (reader!issueId))
                End While
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

    Private Sub txtSearchItem_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtSearchItem.TextChanged
        If Me.txtSearchItem.Text.Trim.Length <= 2 Then
            Me.dgVwItemReturn.Rows.Clear()
            Exit Sub
        End If
        Me.dgVwItemReturn.Rows.Clear()
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            Me.cmdIssueApp.Connection = conn
            Me.cmdIssueApp.CommandType = CommandType.Text
            Me.cmdIssueApp.CommandText = "SELECT * FROM vwInvPendingReturn WHERE (itemDescription LIKE @itemDescription) AND (pendingQty>0) ORDER BY issueReqDate"
            Me.cmdIssueApp.Parameters.Clear()
            Me.cmdIssueApp.Parameters.AddWithValue("@itemDescription", String.Format("%{0}%", TryCast(Me.txtSearchItem.Text.Trim, String).Trim))
            reader = Me.cmdIssueApp.ExecuteReader()
            If reader.HasRows = True Then
                While reader.Read
                    Dim issueReqDate As String = CDate(IIf(DBNull.Value.Equals(reader!issueReqDate), "", (reader!issueReqDate))).Day.ToString("00") & "-" & _
                        CDate(IIf(DBNull.Value.Equals(reader!issueReqDate), "", (reader!issueReqDate))).Month.ToString("00") & "-" & _
                        CDate(IIf(DBNull.Value.Equals(reader!issueReqDate), "", (reader!issueReqDate))).Year.ToString("0000")

                    Dim rowNum As Integer = Me.dgVwItemReturn.Rows.Count
                    Dim row As String() = New String() {IIf(DBNull.Value.Equals(reader!itemDescription), "", (reader!itemDescription)), _
                                    IIf(DBNull.Value.Equals(reader!issuedToTypeName), "", (reader!issuedToTypeName)), _
                                    IIf(DBNull.Value.Equals(reader!issueToName), "", (reader!issueToName)), _
                                    IIf(DBNull.Value.Equals(reader!uom), "", (reader!uom)), _
                                    issueReqDate, _
                                    IIf(DBNull.Value.Equals(reader!issueQty), "", (reader!issueQty)), _
                                    IIf(DBNull.Value.Equals(reader!returnQty), "", (reader!returnQty)), _
                                    IIf(DBNull.Value.Equals(reader!pendingQty), "", (reader!pendingQty)), _
                                    IIf(DBNull.Value.Equals(reader!pendingCost), "", (reader!pendingCost))}
                    Me.dgVwItemReturn.Rows.Add(row)
                    Me.dgVwItemReturn.Rows(rowNum).Tag = IIf(DBNull.Value.Equals(reader!issueId), "", (reader!issueId))
                End While
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

    Private Sub btnReturn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnReturn.Click
        Dim toUpdate As Boolean = False
        If Me.dgVwItemReturn.RowCount <= 0 Then
            MsgBox("Missing Items To Update.", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "System Reply")
            Exit Sub
        End If
        For i = 0 To Me.dgVwItemReturn.RowCount - 1
            If CDbl(Me.dgVwItemReturn.Rows(i).Cells(9).Value) > 0 Then
                toUpdate = True
            End If
        Next
        If toUpdate = False Then
            MsgBox("No item to update found.", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "System Reply")
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

            For i = 0 To Me.dgVwItemReturn.RowCount - 1
                If CDbl(Me.dgVwItemReturn.Rows(i).Cells(9).Value) > 0 Then
                    Me.cmdIssueApp.Connection = conn
                    Me.cmdIssueApp.CommandType = CommandType.StoredProcedure
                    Me.cmdIssueApp.CommandText = "sprocInvIssueReturns"
                    Me.cmdIssueApp.Parameters.Clear()
                    Me.cmdIssueApp.Parameters.AddWithValue("@issueId", Me.dgVwItemReturn.Rows(i).Tag)
                    Me.cmdIssueApp.Parameters.AddWithValue("@returnQty", Me.dgVwItemReturn.Rows(i).Cells(9).Value)
                    Me.cmdIssueApp.Parameters.AddWithValue("@doneBy", userName.Trim)
                    rec = rec + Me.cmdIssueApp.ExecuteNonQuery
                End If
            Next
            If rec > 0 Then
                MsgBox("Item/s Submitted!", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "SuccessFull Transaction")
            End If
            Me.dgVwItemReturn.Rows.Clear()
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub

    Private Sub dgVwItemReturn_CellValueChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgVwItemReturn.CellValueChanged
        Dim returnError As Double = 0
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            
            If e.ColumnIndex = 9 Then
                If IsNumeric(Me.dgVwItemReturn.Item(e.ColumnIndex, e.RowIndex).Value) = False And Not (Me.dgVwItemReturn.Item(e.ColumnIndex, e.RowIndex).Value = "") Then
                    MsgBox("Non Numeric Value Detected.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "System Reply")
                    Me.dgVwItemReturn.Item(e.ColumnIndex, e.RowIndex).Value = ""
                    Exit Sub
                End If
                returnError = (Me.dgVwItemReturn.Item(5, e.RowIndex).Value) - (Me.dgVwItemReturn.Item(9, e.RowIndex).Value)
                If returnError < 0 Then
                    MsgBox("Cannot Return More Than Issued!", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "System Reply")
                    Me.dgVwItemReturn.Item(e.ColumnIndex, e.RowIndex).Value = ""
                End If
            End If
        Catch ex As Exception
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub
End Class