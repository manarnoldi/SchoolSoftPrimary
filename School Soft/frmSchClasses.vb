Imports System.Data.SqlClient
Public Class frmSchClasses
    Dim queryType As String = Nothing
    Dim Rec As Integer = 0
    Dim recordExists As Boolean = False
    Dim cmdSchClasses As New SqlCommand
    Dim reader As SqlDataReader
    Private Sub frmSchClasses_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            loadCombos()
            loadClasses()
        Catch ex As Exception
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub

    Private Sub frmSchClasses_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint
        If Me.IsMdiChild Then
            Dim PnlLoc As New Point
            PnlLoc.X = CInt((Me.Width - Me.pnlSchClasses.Width) / 2)
            PnlLoc.Y = CInt((Me.Height - Me.pnlSchClasses.Height) / 2.5)
            Me.pnlSchClasses.Location = PnlLoc
        Else
            Me.pnlSchClasses.Dock = DockStyle.Fill
        End If
    End Sub
    Private Sub loadCombos()
        cboClassName.Items.Clear()
        cboStream.Items.Clear()
        cboYear.Items.Clear()
        cboClassName.SelectedIndex = -1
        cboStream.SelectedIndex = -1
        cboYear.SelectedIndex = -1

        cmdSchClasses.CommandType = CommandType.Text
        cmdSchClasses.Connection = conn
        cmdSchClasses.CommandText = "SELECT DISTINCT year FROM tblClasses WHERE (status='True') ORDER BY Year"
        cmdSchClasses.Parameters.Clear()
        reader = cmdSchClasses.ExecuteReader
        If reader.HasRows Then
            While reader.Read
                cboYear.Items.Add(IIf(DBNull.Value.Equals(reader!year), "", reader!year))
            End While
        End If
        reader.Close()

        cmdSchClasses.CommandText = "SELECT DISTINCT className FROM tblClasses WHERE (status='True') ORDER BY className"
        cmdSchClasses.Parameters.Clear()
        reader = cmdSchClasses.ExecuteReader
        If reader.HasRows Then
            While reader.Read
                cboClassName.Items.Add(IIf(DBNull.Value.Equals(reader!className), "", reader!className))
            End While
        End If
        reader.Close()

        cmdSchClasses.CommandText = "SELECT DISTINCT stream FROM tblClasses WHERE (status='True') ORDER BY stream"
        cmdSchClasses.Parameters.Clear()
        reader = cmdSchClasses.ExecuteReader
        If reader.HasRows Then
            While reader.Read
                cboStream.Items.Add(IIf(DBNull.Value.Equals(reader!stream), "", reader!stream))
            End While
        End If
        reader.Close()
    End Sub
    Private Sub loadClasses()
        Me.lstClasses.Items.Clear()
        cmdSchClasses.Connection = conn
        cmdSchClasses.CommandType = CommandType.Text
        cmdSchClasses.CommandText = "SELECT * FROM tblClasses WHERE (status='True') ORDER BY  year,className,stream"
        cmdSchClasses.Parameters.Clear()
        reader = cmdSchClasses.ExecuteReader
        If reader.HasRows Then
            While reader.Read
                li = Me.lstClasses.Items.Add(IIf(DBNull.Value.Equals(reader!year), "", reader!year))
                li.SubItems.Add(IIf(DBNull.Value.Equals(reader!className), "", reader!className))
                li.SubItems.Add(IIf(DBNull.Value.Equals(reader!stream), "", reader!stream))
                li.Tag = IIf(DBNull.Value.Equals(reader!classId), "", reader!classId)
            End While
        End If
        reader.Close()
    End Sub
    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub btnView_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnView.Click
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            loadClasses()
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
            Me.btnSave.Enabled = True
            Me.btnDelete.Enabled = False
            Me.btnUpdate.Enabled = False
        End Try
    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        If Me.cboClassName.Text.Trim.Length <= 0 Then
            MsgBox("Class Name Missing", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboStream.Text.Trim.Length <= 0 Then
            MsgBox("Stream Missing", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboYear.Text.Trim.Length <= 0 Then
            MsgBox("Year is missing", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            recordExists = False
            checkForExistence()
            If recordExists = True Then
                MsgBox("Record Exists!", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "UNSuccessFul Transaction")
                Exit Sub
            End If
            queryType = "INSERT"
            Dim result As MsgBoxResult = MsgBox("Save Record?", MsgBoxStyle.Question + MsgBoxStyle.ApplicationModal + MsgBoxStyle.YesNo, "Confirm Transaction")
            If result = MsgBoxResult.Yes Then
                cmdSchClasses.Connection = conn
                cmdSchClasses.CommandType = CommandType.StoredProcedure
                cmdSchClasses.CommandText = "sprocClasses"
                cmdSchClasses.Parameters.Clear()
                cmdSchClasses.Parameters.AddWithValue("@className", Me.cboClassName.Text.Trim)
                cmdSchClasses.Parameters.AddWithValue("@stream", Me.cboStream.Text.Trim)
                cmdSchClasses.Parameters.AddWithValue("@year", Me.cboYear.Text.Trim)
                cmdSchClasses.Parameters.AddWithValue("@regBy", userName.Trim)
                cmdSchClasses.Parameters.AddWithValue("@queryType", queryType.Trim)
                cmdSchClasses.Parameters.AddWithValue("@dateOfReg", Date.Now)
                Rec = cmdSchClasses.ExecuteNonQuery
                If Rec > 0 Then
                    MsgBox("Record Saved!", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "SuccessFul Transaction")
                End If
            End If
            loadCombos()
            loadClasses()
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            Me.btnSave.Enabled = True
            Me.btnDelete.Enabled = False
            Me.btnUpdate.Enabled = False
            Me.cboClassName.Text = ""
            Me.cboStream.Text = ""
            Me.cboYear.Text = ""
            Me.cboClassName.SelectedIndex = -1
            Me.cboStream.SelectedIndex = -1
            Me.cboYear.SelectedIndex = -1
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub
    Private Sub checkForExistence()
        If conn.State = ConnectionState.Open Then
            conn.Close()
        End If
        dbconnection()
        cmdSchClasses.Connection = conn
        cmdSchClasses.CommandType = CommandType.Text
        cmdSchClasses.CommandText = "SELECT * FROM tblClasses WHERE (status='True')AND (className=@className)" & _
            vbNewLine & " AND (stream=@stream) AND (year=@year)"
        cmdSchClasses.Parameters.Clear()
        cmdSchClasses.Parameters.AddWithValue("@className", Me.cboClassName.Text.Trim)
        cmdSchClasses.Parameters.AddWithValue("@stream", Me.cboStream.Text.Trim)
        cmdSchClasses.Parameters.AddWithValue("@year", Me.cboYear.Text.Trim)
        reader = cmdSchClasses.ExecuteReader
        If reader.HasRows Then
            recordExists = True
        Else
            recordExists = False
        End If
        reader.Close()
    End Sub

    Private Sub lstClasses_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lstClasses.Click
        If Me.lstClasses.SelectedItems.Count = 1 Then
            Me.cboClassName.Text = Me.lstClasses.SelectedItems(0).SubItems(1).Text.Trim
            Me.cboStream.Text = Me.lstClasses.SelectedItems(0).SubItems(2).Text.Trim
            Me.cboYear.Text = Me.lstClasses.SelectedItems(0).Text.Trim
            Me.cboClassName.Tag = Me.lstClasses.SelectedItems(0).Tag
            Me.btnSave.Enabled = False
            Me.btnDelete.Enabled = True
            Me.btnUpdate.Enabled = True
        End If
    End Sub

    Private Sub btnUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdate.Click
        If Me.cboClassName.Text.Trim.Length <= 0 Then
            MsgBox("Class Name Missing", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboStream.Text.Trim.Length <= 0 Then
            MsgBox("Stream Missing", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboYear.Text.Trim.Length <= 0 Then
            MsgBox("Year is missing", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            queryType = "UPDATE"
            recordExists = False
            checkForExistence()
            If recordExists = True Then
                MsgBox("Record Exists!", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "UNSuccessFul Transaction")
                Exit Sub
            End If
            Dim result As MsgBoxResult = MsgBox("Update Record?", MsgBoxStyle.Question + MsgBoxStyle.ApplicationModal + MsgBoxStyle.YesNo, "Confirm Transaction")
            If result = MsgBoxResult.Yes Then
                cmdSchClasses.Connection = conn
                cmdSchClasses.CommandType = CommandType.StoredProcedure
                cmdSchClasses.CommandText = "sprocClasses"
                cmdSchClasses.Parameters.Clear()
                cmdSchClasses.Parameters.AddWithValue("@classId", Me.cboClassName.Tag)
                cmdSchClasses.Parameters.AddWithValue("@className", Me.cboClassName.Text.Trim)
                cmdSchClasses.Parameters.AddWithValue("@stream", Me.cboStream.Text.Trim)
                cmdSchClasses.Parameters.AddWithValue("@year", Me.cboYear.Text.Trim)
                cmdSchClasses.Parameters.AddWithValue("@regBy", userName.Trim)
                cmdSchClasses.Parameters.AddWithValue("@queryType", queryType.Trim)
                cmdSchClasses.Parameters.AddWithValue("@dateOfReg", Date.Now)
                Rec = cmdSchClasses.ExecuteNonQuery()
                If Rec > 0 Then
                    MsgBox("Record Updated!", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "SuccessFul Transaction")
                End If
            End If
            loadCombos()
            loadClasses()
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            Me.btnSave.Enabled = True
            Me.btnDelete.Enabled = False
            Me.btnUpdate.Enabled = False
            Me.cboClassName.Text = ""
            Me.cboStream.Text = ""
            Me.cboYear.Text = ""
            Me.cboClassName.SelectedIndex = -1
            Me.cboStream.SelectedIndex = -1
            Me.cboYear.SelectedIndex = -1
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub

    Private Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            queryType = "DELETE"
            Dim result As MsgBoxResult = MsgBox("Delete Record?", MsgBoxStyle.Question + MsgBoxStyle.ApplicationModal + MsgBoxStyle.YesNo, "Confirm Transaction")
            If result = MsgBoxResult.Yes Then
                cmdSchClasses.Connection = conn
                cmdSchClasses.CommandType = CommandType.StoredProcedure
                cmdSchClasses.CommandText = "sprocClasses"
                cmdSchClasses.Parameters.Clear()
                cmdSchClasses.Parameters.AddWithValue("@classId", Me.cboClassName.Tag)
                cmdSchClasses.Parameters.AddWithValue("@className", Me.cboClassName.Text.Trim)
                cmdSchClasses.Parameters.AddWithValue("@stream", Me.cboStream.Text.Trim)
                cmdSchClasses.Parameters.AddWithValue("@year", Me.cboYear.Text.Trim)
                cmdSchClasses.Parameters.AddWithValue("@regBy", userName.Trim)
                cmdSchClasses.Parameters.AddWithValue("@queryType", queryType.Trim)
                cmdSchClasses.Parameters.AddWithValue("@dateOfReg", Date.Now)
                Rec = cmdSchClasses.ExecuteNonQuery()
                If Rec > 0 Then
                    MsgBox("Record Deleted!", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "SuccessFul Transaction")
                End If
            End If
            loadCombos()
            loadClasses()
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            Me.btnSave.Enabled = True
            Me.btnDelete.Enabled = False
            Me.btnUpdate.Enabled = False
            Me.cboClassName.Text = ""
            Me.cboStream.Text = ""
            Me.cboYear.Text = ""
            Me.cboClassName.SelectedIndex = -1
            Me.cboStream.SelectedIndex = -1
            Me.cboYear.SelectedIndex = -1
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub
End Class