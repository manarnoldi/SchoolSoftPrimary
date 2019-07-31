Imports System.Data.SqlClient
Public Class frmStudParents
    Dim recordExists As Boolean = True
    Dim rec As Integer = 0
    Dim cmdStudParents As New SqlCommand
    Dim reader As SqlDataReader
    Dim queryType As String = Nothing
    Private Sub frmStudParents_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
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
        Me.cboPRelation.Items.Clear()
        Me.cboPResidence.Items.Clear()
        Me.cboPOccupation.Items.Clear()
        Me.cboPRelation.Text = ""
        Me.cboPResidence.Text = ""
        Me.cboPOccupation.Text = ""
        cmdStudParents.Connection = conn
        cmdStudParents.CommandType = CommandType.Text
        cmdStudParents.CommandText = "SELECT DISTINCT residence FROM vwStudParents WHERE residence IS NOT NULL ORDER BY residence"
        reader = cmdStudParents.ExecuteReader
        cmdStudParents.Parameters.Clear()
        If reader.HasRows Then
            While reader.Read
                Me.cboPResidence.Items.Add(IIf(DBNull.Value.Equals(reader!residence), "", reader!residence))
            End While
        End If
        reader.Close()

        cmdStudParents.CommandText = "SELECT DISTINCT occupation FROM vwStudParents WHERE occupation IS NOT NULL ORDER BY occupation"
        reader = cmdStudParents.ExecuteReader
        cmdStudParents.Parameters.Clear()
        If reader.HasRows Then
            While reader.Read
                Me.cboPOccupation.Items.Add(IIf(DBNull.Value.Equals(reader!occupation), "", reader!occupation))
            End While
        End If
        reader.Close()

        cmdStudParents.CommandText = "SELECT DISTINCT relationShip FROM vwStudParents WHERE relationShip IS NOT NULL ORDER BY relationShip"
        reader = cmdStudParents.ExecuteReader
        cmdStudParents.Parameters.Clear()
        If reader.HasRows Then
            While reader.Read
                Me.cboPRelation.Items.Add(IIf(DBNull.Value.Equals(reader!relationShip), "", reader!relationShip))
            End While
        End If
        reader.Close()

    End Sub
    Private Sub frmStudParents_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint
        If Me.IsMdiChild Then
            Dim PnlLoc As New Point
            PnlLoc.X = CInt((Me.Width - Me.pnlStdParents.Width) / 2)
            PnlLoc.Y = CInt((Me.Height - Me.pnlStdParents.Height) / 2.5)
            Me.pnlStdParents.Location = PnlLoc
        Else
            Me.pnlStdParents.Dock = DockStyle.Fill
        End If
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub btnLoad_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLoad.Click
        If Me.txtStudNo.Text = "" Then
            Me.txtStudName.Text = ""
            clearTexts()
            Me.lstParentDetails.Items.Clear()
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            cmdStudParents.Connection = conn
            cmdStudParents.CommandType = CommandType.Text
            cmdStudParents.CommandText = "SELECT * FROM tblStudDetails WHERE (admNo=@admNo) AND (Status='True')"
            cmdStudParents.Parameters.Clear()
            cmdStudParents.Parameters.AddWithValue("@admNo", Me.txtStudNo.Text.Trim)
            reader = cmdStudParents.ExecuteReader
            If reader.HasRows Then
                While reader.Read
                    Me.txtStudName.Text = (IIf(DBNull.Value.Equals(reader!FullName), "", reader!FullName))
                End While
            Else
                MsgBox("No student Found!", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
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

    Private Sub txtStudNo_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtStudNo.TextChanged
        Me.txtStudName.Text = ""
        Me.lstParentDetails.Items.Clear()
        If Me.txtStudNo.Text = "" Then
            Me.txtStudName.Text = ""
            Exit Sub
        End If
    End Sub

    Private Sub btnAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAdd.Click
        If Me.txtStudNo.Text.Trim.Length <= 0 Then
            MsgBox("Missing Student Number", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.txtStudName.Text.Trim.Length <= 0 Then
            MsgBox("Missing Student Name", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf (Me.txtPFName.Text.Trim.Length <= 0 And Me.txtPMName.Text.Trim.Length <= 0) Or _
            (Me.txtPFName.Text.Trim.Length <= 0 And Me.txtPLName.Text.Trim.Length <= 0) Or _
            (Me.txtPMName.Text.Trim.Length <= 0 And Me.txtPLName.Text.Trim.Length <= 0) Then
            MsgBox("Enter at least two parents Names", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.txtPPhoneNo.Text.Trim.Length <= 0 Then
            MsgBox("Enter Parents Phone Number", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.txtPNationalId.Text.Trim.Length <= 0 Then
            MsgBox("Enter Parents National ID", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboPOccupation.Text.Trim.Length <= 0 Then
            MsgBox("Enter Parents occupation", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboPOccupation.Text.Trim.Length <= 0 Then
            MsgBox("Enter Parents Relationship", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboPResidence.Text.Trim.Length <= 0 Then
            MsgBox("Enter Parents Residence", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If

        For i = 0 To Me.lstParentDetails.Items.Count - 1
            If (Me.txtStudNo.Text.Trim = Me.lstParentDetails.Items(i).Text.Trim) And _
                (Me.txtPNationalId.Text.Trim = Me.lstParentDetails.Items(i).SubItems(6).Text) Then
                MsgBox("Record Already Added to the list", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
                Exit Sub
            End If
        Next
        li = Me.lstParentDetails.Items.Add(Me.txtStudNo.Text.Trim)
        li.SubItems.Add(Me.txtPFName.Text.Trim)
        li.SubItems.Add(Me.txtPMName.Text.Trim)
        li.SubItems.Add(Me.txtPLName.Text.Trim)
        li.SubItems.Add(Me.cboPRelation.Text.Trim)
        li.SubItems.Add(Me.cboPResidence.Text.Trim)
        li.SubItems.Add(Me.txtPNationalId.Text.Trim)
        li.SubItems.Add(Me.cboPOccupation.Text.Trim)
        li.SubItems.Add(Me.txtPPhoneNo.Text.Trim)
        li.SubItems(1).Tag = Me.txtPAddress.Text.Trim
        li.SubItems(2).Tag = Me.txtPEmail.Text.Trim
        li.Tag = Me.txtPNationalId.Tag
        'loadCombos()
        'Me.txtPFName.Text = ""
        'Me.txtPLName.Text = ""
        'Me.txtPMName.Text = ""
        'Me.txtPAddress.Text = ""
        'Me.txtPEmail.Text = ""
        'Me.txtPNationalId.Text = ""
        'Me.txtPPhoneNo.Text = ""

    End Sub

    Private Sub btnRemove_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRemove.Click
        If Me.lstParentDetails.SelectedItems.Count > 0 Then
            For i = 0 To Me.lstParentDetails.SelectedItems.Count - 1
                Me.lstParentDetails.SelectedItems(0).Remove()
            Next
        End If

    End Sub

    Private Sub txtPEmail_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPEmail.LostFocus
        If Me.txtPEmail.Text = "" Then
            Exit Sub
        End If
        If Me.txtPEmail.Text.Contains(".") = False Or Me.txtPEmail.Text.Contains("@") = False Then
            MsgBox("Enter Correct Email Address", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Missing Details")
            Me.txtPEmail.Text = ""
        End If
    End Sub

    Private Sub btnView_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnView.Click
        If Me.txtStudNo.Text.Trim.Length <= 0 Then
            MsgBox("Missing Student Number", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
            conn.Open()
        End If
        Try
            If conn.State = ConnectionState.Closed Then
            End If
            dbconnection()
            Me.lstParentDetails.Items.Clear()
            cmdStudParents.Connection = conn
            cmdStudParents.CommandType = CommandType.Text
            cmdStudParents.CommandText = "SELECT * FROM vwStudParents WHERE (admNo=@admNo) AND (parentStatus='True') AND " & _
                vbNewLine & " (studParentStatus='True')  AND (studStatus='True') ORDER BY FullName"
            cmdStudParents.Parameters.Clear()
            cmdStudParents.Parameters.AddWithValue("@admNo", Me.txtStudNo.Text.Trim)
            reader = cmdStudParents.ExecuteReader
            If reader.HasRows Then
                While reader.Read
                    li = Me.lstParentDetails.Items.Add(IIf(DBNull.Value.Equals(reader!admNo), "", reader!admNo))
                    li.SubItems.Add(IIf(DBNull.Value.Equals(reader!FName), "", reader!FName))
                    li.SubItems.Add(IIf(DBNull.Value.Equals(reader!MName), "", reader!MName))
                    li.SubItems.Add(IIf(DBNull.Value.Equals(reader!LName), "", reader!LName))
                    li.SubItems.Add(IIf(DBNull.Value.Equals(reader!relationShip), "", reader!relationShip))
                    li.SubItems.Add(IIf(DBNull.Value.Equals(reader!residence), "", reader!residence))
                    li.SubItems.Add(IIf(DBNull.Value.Equals(reader!nationalId), "", reader!nationalId))
                    li.SubItems.Add(IIf(DBNull.Value.Equals(reader!occupation), "", reader!occupation))
                    li.SubItems.Add(IIf(DBNull.Value.Equals(reader!PhoneNumber), "", reader!PhoneNumber))
                    li.Tag = (IIf(DBNull.Value.Equals(reader!ParentId), "", reader!parentId))
                    li.SubItems(1).Tag = IIf(DBNull.Value.Equals(reader!contactAddress), "", reader!contactAddress)
                    li.SubItems(2).Tag = IIf(DBNull.Value.Equals(reader!Email), "", reader!Email)
                End While
            Else
                MsgBox("No Registered Parent Found For the student.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            End If
            reader.Close()
            loadCombos()
            Me.txtPFName.Text = ""
            Me.txtPLName.Text = ""
            Me.txtPMName.Text = ""
            Me.txtPAddress.Text = ""
            Me.txtPEmail.Text = ""
            Me.txtPNationalId.Text = ""
            Me.txtPPhoneNo.Text = ""
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            Me.txtPNationalId.Enabled = True
            Me.txtStudNo.Enabled = True
            Me.btnUpdate.Enabled = False
            Me.btnSave.Enabled = True
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try

    End Sub

    Private Sub clearTexts()
        Me.txtPNationalId.Tag = Nothing
        Me.txtPFName.Text = ""
        Me.txtPLName.Text = ""
        Me.txtPMName.Text = ""
        Me.txtPAddress.Text = ""
        Me.txtPEmail.Text = ""
        Me.txtPNationalId.Text = ""
        Me.txtPPhoneNo.Text = ""
        Me.txtStudName.Text = ""
        Me.txtStudNo.Text = ""
        Me.lstParentDetails.Items.Clear()
    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        Dim saveDetails As Integer = 0
        If Me.lstParentDetails.Items.Count <= 0 Then
            MsgBox("No parents in the list to save", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()

            Dim result As MsgBoxResult = MsgBox("Save Record?", MsgBoxStyle.Question + MsgBoxStyle.ApplicationModal + MsgBoxStyle.YesNo, "Confirm Transaction")
            If result = MsgBoxResult.No Then
                Exit Sub
            End If
            For i = 0 To Me.lstParentDetails.Items.Count - 1
                recordExists = True
                checkParent(Me.lstParentDetails.Items(i).Text.Trim, Me.lstParentDetails.Items(i).SubItems(6).Text.Trim)
                If recordExists = False Then
                    queryType = "INSERT"
                    cmdStudParents.CommandText = "sprocStudentParent"
                    cmdStudParents.Connection = conn
                    cmdStudParents.CommandType = CommandType.StoredProcedure
                    cmdStudParents.Parameters.Clear()
                    cmdStudParents.Parameters.AddWithValue("@studNo", Me.lstParentDetails.Items(i).Text.Trim)
                    cmdStudParents.Parameters.AddWithValue("@Fname", Me.lstParentDetails.Items(i).SubItems(1).Text.Trim)
                    cmdStudParents.Parameters.AddWithValue("@MName", Me.lstParentDetails.Items(i).SubItems(2).Text.Trim)
                    cmdStudParents.Parameters.AddWithValue("@LName", Me.lstParentDetails.Items(i).SubItems(3).Text.Trim)
                    cmdStudParents.Parameters.AddWithValue("@Email", Me.lstParentDetails.Items(i).SubItems(2).Tag)
                    cmdStudParents.Parameters.AddWithValue("@contactAddress", Me.lstParentDetails.Items(i).SubItems(1).Tag)
                    cmdStudParents.Parameters.AddWithValue("@PhoneNumber", Me.lstParentDetails.Items(i).SubItems(8).Text.Trim)
                    cmdStudParents.Parameters.AddWithValue("@nationalId", Me.lstParentDetails.Items(i).SubItems(6).Text.Trim)
                    cmdStudParents.Parameters.AddWithValue("@residence", Me.lstParentDetails.Items(i).SubItems(5).Text.Trim)
                    cmdStudParents.Parameters.AddWithValue("@occupation", Me.lstParentDetails.Items(i).SubItems(7).Text.Trim)
                    cmdStudParents.Parameters.AddWithValue("@dateOfReg", Date.Now)
                    cmdStudParents.Parameters.AddWithValue("@regBy", userName)
                    cmdStudParents.Parameters.AddWithValue("@relationship", Me.lstParentDetails.Items(i).SubItems(4).Text.Trim)
                    cmdStudParents.Parameters.AddWithValue("@queryType", queryType.Trim)
                    rec = cmdStudParents.ExecuteNonQuery
                ElseIf recordExists = True Then
                    saveDetails = 1
                End If
            Next
            loadCombos()
            If rec > 0 Then
                MsgBox("Record/s Saved", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Successfull Transaction")
                If saveDetails = 1 Then
                    MsgBox("Some Parents Had already been saved", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Successfull Transaction")
                End If
            End If
            clearTexts()
            Me.txtPNationalId.Enabled = True
            Me.txtStudNo.Enabled = True
            Me.btnUpdate.Enabled = False
            Me.btnSave.Enabled = True
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub
    Private Function checkParent(ByVal studNo As String, ByVal parentNationalId As String)
        recordExists = True
        cmdStudParents.CommandType = CommandType.Text
        cmdStudParents.CommandText = "SELECT * FROM vwStudParents vwStudParents WHERE (admNo=@admNo) AND (nationalId=@nationalId)"
        cmdStudParents.Connection = conn
        cmdStudParents.Parameters.Clear()
        cmdStudParents.Parameters.AddWithValue("@admNo", studNo.Trim)
        cmdStudParents.Parameters.AddWithValue("@nationalId", parentNationalId.Trim)
        reader = cmdStudParents.ExecuteReader
        If reader.HasRows Then
            recordExists = True
        Else
            recordExists = False
        End If
        reader.Close()
    End Function
    Private Sub ExitToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitToolStripMenuItem.Click
        Me.Close()
    End Sub

    Private Sub EditToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EditToolStripMenuItem.Click
        If Me.lstParentDetails.SelectedItems.Count = 1 Then
            Me.txtStudNo.Text = Me.lstParentDetails.SelectedItems(0).Text
            Me.txtPFName.Text = Me.lstParentDetails.SelectedItems(0).SubItems(1).Text
            Me.txtPMName.Text = Me.lstParentDetails.SelectedItems(0).SubItems(2).Text
            Me.txtPLName.Text = Me.lstParentDetails.SelectedItems(0).SubItems(3).Text
            Me.cboPRelation.Text = Me.lstParentDetails.SelectedItems(0).SubItems(4).Text
            Me.cboPResidence.Text = Me.lstParentDetails.SelectedItems(0).SubItems(5).Text
            Me.txtPNationalId.Text = Me.lstParentDetails.SelectedItems(0).SubItems(6).Text
            Me.cboPOccupation.Text = Me.lstParentDetails.SelectedItems(0).SubItems(7).Text
            Me.txtPPhoneNo.Text = Me.lstParentDetails.SelectedItems(0).SubItems(8).Text
            Me.txtPAddress.Text = Me.lstParentDetails.SelectedItems(0).SubItems(1).Tag
            Me.txtPEmail.Text = Me.lstParentDetails.SelectedItems(0).SubItems(2).Tag
            Me.txtPNationalId.Tag = Me.lstParentDetails.SelectedItems(0).Tag
            Me.txtPNationalId.Enabled = False
            Me.txtStudNo.Enabled = False
            Me.btnUpdate.Enabled = True
            Me.btnSave.Enabled = False
            Me.lstParentDetails.SelectedItems(0).Remove()
        End If
    End Sub

    Private Sub btnUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdate.Click
        If Me.lstParentDetails.Items.Count <= 0 Then
            MsgBox("No parents in the list to Update", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            Dim result As MsgBoxResult = MsgBox("Update Record?", MsgBoxStyle.Question + MsgBoxStyle.ApplicationModal + MsgBoxStyle.YesNo, "Confirm Transaction")
            If result = MsgBoxResult.No Then
                Exit Sub
            End If
            For i = 0 To Me.lstParentDetails.Items.Count - 1
                    queryType = "UPDATE"
                    cmdStudParents.CommandText = "sprocStudentParent"
                    cmdStudParents.Connection = conn
                    cmdStudParents.CommandType = CommandType.StoredProcedure
                    cmdStudParents.Parameters.Clear()
                    cmdStudParents.Parameters.AddWithValue("@studNo", Me.lstParentDetails.Items(i).Text.Trim)
                    cmdStudParents.Parameters.AddWithValue("@Fname", Me.lstParentDetails.Items(i).SubItems(1).Text.Trim)
                    cmdStudParents.Parameters.AddWithValue("@MName", Me.lstParentDetails.Items(i).SubItems(2).Text.Trim)
                    cmdStudParents.Parameters.AddWithValue("@LName", Me.lstParentDetails.Items(i).SubItems(3).Text.Trim)
                    cmdStudParents.Parameters.AddWithValue("@Email", Me.lstParentDetails.Items(i).SubItems(2).Tag)
                    cmdStudParents.Parameters.AddWithValue("@contactAddress", Me.lstParentDetails.Items(i).SubItems(1).Tag)
                    cmdStudParents.Parameters.AddWithValue("@PhoneNumber", Me.lstParentDetails.Items(i).SubItems(8).Text.Trim)
                    cmdStudParents.Parameters.AddWithValue("@nationalId", Me.lstParentDetails.Items(i).SubItems(6).Text.Trim)
                    cmdStudParents.Parameters.AddWithValue("@residence", Me.lstParentDetails.Items(i).SubItems(5).Text.Trim)
                    cmdStudParents.Parameters.AddWithValue("@occupation", Me.lstParentDetails.Items(i).SubItems(7).Text.Trim)
                    cmdStudParents.Parameters.AddWithValue("@dateOfReg", Date.Now)
                    cmdStudParents.Parameters.AddWithValue("@regBy", userName)
                    cmdStudParents.Parameters.AddWithValue("@relationship", Me.lstParentDetails.Items(i).SubItems(4).Text.Trim)
                cmdStudParents.Parameters.AddWithValue("@queryType", queryType.Trim)
                cmdStudParents.Parameters.AddWithValue("@parentIdEdit", Me.lstParentDetails.Items(i).Tag)
                rec = cmdStudParents.ExecuteNonQuery
            Next
            loadCombos()
            If rec > 0 Then
                MsgBox("Record/s Updated", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Successfull Transaction")
            End If
            clearTexts()
            Me.btnUpdate.Enabled = False
            Me.btnSave.Enabled = True
            Me.txtPNationalId.Enabled = True
            Me.txtStudNo.Enabled = True
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub

    Private Sub lstParentDetails_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lstParentDetails.SelectedIndexChanged

    End Sub

    Private Sub DeleteToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DeleteToolStripMenuItem.Click
        If Me.lstParentDetails.SelectedItems.Count = 1 Then


            Try
                If conn.State = ConnectionState.Closed Then
                    conn.Open()
                End If
                dbconnection()
                Dim result As MsgBoxResult = MsgBox("Delete Record?", MsgBoxStyle.Question + MsgBoxStyle.ApplicationModal + MsgBoxStyle.YesNo, "Confirm Transaction")
                If result = MsgBoxResult.No Then
                    Exit Sub
                End If
                queryType = "DELETE"
                cmdStudParents.CommandText = "sprocStudentParent"
                cmdStudParents.Connection = conn
                cmdStudParents.CommandType = CommandType.StoredProcedure
                cmdStudParents.Parameters.Clear()
                cmdStudParents.Parameters.AddWithValue("@studNo", Me.lstParentDetails.SelectedItems(0).Text.Trim)
                cmdStudParents.Parameters.AddWithValue("@Fname", Me.lstParentDetails.SelectedItems(0).SubItems(1).Text.Trim)
                cmdStudParents.Parameters.AddWithValue("@MName", Me.lstParentDetails.SelectedItems(0).SubItems(2).Text.Trim)
                cmdStudParents.Parameters.AddWithValue("@LName", Me.lstParentDetails.SelectedItems(0).SubItems(3).Text.Trim)
                cmdStudParents.Parameters.AddWithValue("@Email", Me.lstParentDetails.SelectedItems(0).SubItems(2).Tag)
                cmdStudParents.Parameters.AddWithValue("@contactAddress", Me.lstParentDetails.SelectedItems(0).SubItems(1).Tag)
                cmdStudParents.Parameters.AddWithValue("@PhoneNumber", Me.lstParentDetails.SelectedItems(0).SubItems(8).Text.Trim)
                cmdStudParents.Parameters.AddWithValue("@nationalId", Me.lstParentDetails.SelectedItems(0).SubItems(6).Text.Trim)
                cmdStudParents.Parameters.AddWithValue("@residence", Me.lstParentDetails.SelectedItems(0).SubItems(5).Text.Trim)
                cmdStudParents.Parameters.AddWithValue("@occupation", Me.lstParentDetails.SelectedItems(0).SubItems(7).Text.Trim)
                cmdStudParents.Parameters.AddWithValue("@dateOfReg", Date.Now)
                cmdStudParents.Parameters.AddWithValue("@regBy", userName)
                cmdStudParents.Parameters.AddWithValue("@relationship", Me.lstParentDetails.SelectedItems(0).SubItems(4).Text.Trim)
                cmdStudParents.Parameters.AddWithValue("@queryType", queryType.Trim)
                cmdStudParents.Parameters.AddWithValue("@parentIdEdit", Me.lstParentDetails.SelectedItems(0).Tag)
                rec = cmdStudParents.ExecuteNonQuery
                loadCombos()
                If rec > 0 Then
                    MsgBox("Record/s Deleted", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Successfull Transaction")
                End If
                clearTexts()
                Me.btnUpdate.Enabled = False
                Me.btnSave.Enabled = True
                Me.txtPNationalId.Enabled = True
                Me.txtStudNo.Enabled = True
            Catch ex As Exception
                MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Finally
                If conn.State = ConnectionState.Open Then
                    conn.Close()
                End If
            End Try
        End If
    End Sub
End Class