Imports System.Data.SqlClient
Public Class frmSearchStudentDetails
    Dim reader As SqlDataReader
    Dim cmdSearchStudentDetails As New SqlCommand
    Private Sub txtStudentNo_TextChanged(sender As Object, e As EventArgs) Handles txtStudentNo.TextChanged
        Me.lstStudDetails.Items.Clear()
        If Me.txtStudentNo.Text.Trim.Length <= 1 Then
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            cmdSearchStudentDetails.Connection = conn
            cmdSearchStudentDetails.CommandType = CommandType.Text
            cmdSearchStudentDetails.CommandText = "SELECT * FROM vwStudClass WHERE (admNo LIKE @admNo) ORDER BY admNo"
            cmdSearchStudentDetails.Parameters.Clear()
            cmdSearchStudentDetails.Parameters.AddWithValue("admNo", String.Format("%{0}%", TryCast(Me.txtStudentNo.Text.Trim, String).Trim))
            reader = cmdSearchStudentDetails.ExecuteReader
            If reader.HasRows Then
                While reader.Read
                    li = Me.lstStudDetails.Items.Add(IIf(DBNull.Value.Equals(reader!admNo), "", reader!admNo))
                    li.SubItems.Add(IIf(DBNull.Value.Equals(reader!FullName), "", reader!FullName))
                    li.SubItems.Add(IIf(DBNull.Value.Equals(reader!className), "", reader!className))
                    li.SubItems.Add(IIf(DBNull.Value.Equals(reader!stream), "", reader!stream))
                    li.SubItems.Add(IIf(DBNull.Value.Equals(reader!year), "", reader!year))
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

    Private Sub frmSearchStudentDetails_Paint(sender As Object, e As PaintEventArgs) Handles Me.Paint
        If Me.IsMdiChild Then
            Dim PnlLoc As New Point
            PnlLoc.X = CInt((Me.Width - Me.pnlSearchDetails.Width) / 2)
            PnlLoc.Y = CInt((Me.Height - Me.pnlSearchDetails.Height) / 2.5)
            Me.pnlSearchDetails.Location = PnlLoc
        Else
            Me.pnlSearchDetails.Dock = DockStyle.Fill
        End If
    End Sub

    Private Sub txtStudentName_TextChanged(sender As Object, e As EventArgs) Handles txtStudentName.TextChanged
        Me.lstStudDetails.Items.Clear()
        If Me.txtStudentName.Text.Trim.Length <= 2 Then
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            cmdSearchStudentDetails.Connection = conn
            cmdSearchStudentDetails.CommandType = CommandType.Text
            cmdSearchStudentDetails.CommandText = "SELECT * FROM vwStudClass WHERE (FullName LIKE @FullName) ORDER BY admNo"
            cmdSearchStudentDetails.Parameters.Clear()
            cmdSearchStudentDetails.Parameters.AddWithValue("FullName", String.Format("%{0}%", TryCast(Me.txtStudentName.Text.Trim, String).Trim))
            reader = cmdSearchStudentDetails.ExecuteReader
            If reader.HasRows Then
                While reader.Read
                    li = Me.lstStudDetails.Items.Add(IIf(DBNull.Value.Equals(reader!admNo), "", reader!admNo))
                    li.SubItems.Add(IIf(DBNull.Value.Equals(reader!FullName), "", reader!FullName))
                    li.SubItems.Add(IIf(DBNull.Value.Equals(reader!className), "", reader!className))
                    li.SubItems.Add(IIf(DBNull.Value.Equals(reader!stream), "", reader!stream))
                    li.SubItems.Add(IIf(DBNull.Value.Equals(reader!year), "", reader!year))
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

    Private Sub frmSearchStudentDetails_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class