Imports System.Data.SqlClient
Public Class frmSchMaxSubjectSetUp
    Dim recordExists As Boolean = True
    Dim reader As SqlDataReader
    Dim cmdMaxSubjects As New SqlCommand
    Dim rec As Integer = 0
    Dim queryType As String = Nothing
    Private Sub frmSchMaxSubjectSetUp_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            loadLists()
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub
    Private Sub loadLists()
        Me.lstMaxSubjects.Items.Clear()
        Me.cmdMaxSubjects.Connection = conn
        Me.cmdMaxSubjects.CommandText = "SELECT * FROM tblMaxNoOfSubjects WHERE (status=1) ORDER BY parameter,class,maxNo"
        Me.cmdMaxSubjects.CommandType = CommandType.Text
        Me.cmdMaxSubjects.Parameters.Clear()
        reader = Me.cmdMaxSubjects.ExecuteReader
        While reader.Read
            li = Me.lstMaxSubjects.Items.Add(IIf(DBNull.Value.Equals(reader!parameter), "", reader!parameter))
            li.SubItems.Add(IIf(DBNull.Value.Equals(reader!class), "", reader!class))
            li.SubItems.Add(IIf(DBNull.Value.Equals(reader!maxNo), "", reader!maxNo))
        End While
        reader.Close()
    End Sub
    Private Sub frmSchMaxSubjectSetUp_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint
        If Me.IsMdiChild Then
            Dim PnlLoc As New Point
            PnlLoc.X = CInt((Me.Width - Me.pnlStudSubject.Width) / 2)
            PnlLoc.Y = CInt((Me.Height - Me.pnlStudSubject.Height) / 2.5)
            Me.pnlStudSubject.Location = PnlLoc
        Else
            Me.pnlStudSubject.Dock = DockStyle.Fill
        End If
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub
    Private Sub clearTexts()
        Me.txtDescription.Text = ""
        Me.txtMaxNo.Text = ""
        Me.txtParameterName.Text = ""
    End Sub
    Private Sub btnView_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnView.Click
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            loadLists()
            clearTexts()
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub

    Private Sub btnClear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClear.Click
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            clearTexts()
            loadLists()
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub
    Private Sub checkIfRecordExists()
        recordExists = True
        Me.cmdMaxSubjects.Connection = conn
        Me.cmdMaxSubjects.CommandText = "SELECT * FROM tblMaxNoOfSubjects WHERE (parameter=@parameter) AND (maxNo=@maxNo) AND " & _
            vbNewLine & " (class=@class)  AND (status=1)"
        Me.cmdMaxSubjects.CommandType = CommandType.Text
        Me.cmdMaxSubjects.Parameters.Clear()
        Me.cmdMaxSubjects.Parameters.AddWithValue("@parameter", Me.txtParameterName.Text.Trim)
        Me.cmdMaxSubjects.Parameters.AddWithValue("@maxNo", Me.txtMaxNo.Text.Trim)
        Me.cmdMaxSubjects.Parameters.AddWithValue("@class", Me.txtDescription.Text.Trim)
        reader = Me.cmdMaxSubjects.ExecuteReader
        If reader.HasRows = True Then
            recordExists = True
        ElseIf reader.HasRows = False Then
            recordExists = False
        End If
        reader.Close()
    End Sub
    Private Sub btnSaveChanges_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSaveChanges.Click
        If Me.txtDescription.Text.Trim.Length <= 0 Then
            MsgBox("Description is Missing", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "System Reply")
            Exit Sub
        ElseIf Me.txtMaxNo.Text.Trim.Length <= 0 Then
            MsgBox("Maximum Number is missing", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "System Reply")
            Exit Sub
        ElseIf Me.txtParameterName.Text.Trim.Length <= 0 Then
            MsgBox("Parameter is missing", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "System Reply")
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            checkIfRecordExists()
            If recordExists = True Then
                MsgBox("Record Already Exists", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "System Reply")
                Exit Sub
            End If
            Dim result As MsgBoxResult = MsgBox("Update Changes?", MsgBoxStyle.Question + MsgBoxStyle.ApplicationModal + MsgBoxStyle.YesNo, "Confirm Transaction")
            If result = MsgBoxResult.No Then
                Exit Sub
            End If
            queryType = "UPDATE"
            Me.cmdMaxSubjects.CommandType = CommandType.StoredProcedure
            Me.cmdMaxSubjects.Connection = conn
            Me.cmdMaxSubjects.CommandText = "sprocMaxSubjects"
            Me.cmdMaxSubjects.Parameters.Clear()
            Me.cmdMaxSubjects.Parameters.AddWithValue("@parameter", Me.txtParameterName.Text.Trim)
            Me.cmdMaxSubjects.Parameters.AddWithValue("@maxNo", Me.txtMaxNo.Text.Trim)
            Me.cmdMaxSubjects.Parameters.AddWithValue("@class", Me.txtDescription.Text.Trim)
            Me.cmdMaxSubjects.Parameters.AddWithValue("@queryType", queryType.Trim)
            Me.cmdMaxSubjects.Parameters.AddWithValue("@dateOfReg", Date.Now)
            Me.cmdMaxSubjects.Parameters.AddWithValue("@regBy", userName.Trim)
            rec = Me.cmdMaxSubjects.ExecuteNonQuery
            If rec > 0 Then
                MsgBox("Record Updated", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "SuccessFul Transaction")
            End If
            clearTexts()
            loadLists()
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub

    Private Sub lstMaxSubjects_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lstMaxSubjects.Click
        If Me.lstMaxSubjects.Items.Count <= 0 Then
            MsgBox("No items in the list.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "System Reply")
            Exit Sub
        ElseIf Me.lstMaxSubjects.SelectedItems.Count <= 0 Then
            MsgBox("Select period to Update.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "System Reply")
            Exit Sub
        ElseIf Me.lstMaxSubjects.SelectedItems.Count > 1 Then
            MsgBox("Select one period at a time.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "System Reply")
            Exit Sub
        End If
        If Me.lstMaxSubjects.SelectedItems.Count = 1 Then
            Me.txtParameterName.Text = Me.lstMaxSubjects.SelectedItems(0).Text.Trim
            Me.txtDescription.Text = Me.lstMaxSubjects.SelectedItems(0).SubItems(1).Text
            Me.txtMaxNo.Text = Me.lstMaxSubjects.SelectedItems(0).SubItems(2).Text
        End If
        Me.lstMaxSubjects.Items.Clear()
    End Sub

End Class