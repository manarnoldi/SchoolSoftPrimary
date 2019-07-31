Imports System.Data.SqlClient
Public Class frmDomainRights
    Dim available As Boolean = False
    Dim Todelete As Boolean = False
    Dim cmdDomainRights As New SqlCommand
    Dim cmdDomainRights1 As New SqlCommand
    Dim reader As SqlDataReader
    Dim moduleName As String
    Dim rightsId As Integer
    Dim domainId As Integer
    Dim rec As Integer = 0
    Private Sub frmDomainRights_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        loadCombos()
        loadAllModules()
    End Sub
    Private Sub loadCombos()
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            cboDomains.Items.Clear()
            cmdDomainRights.Connection = conn
            cmdDomainRights.CommandType = CommandType.Text
            cmdDomainRights.CommandText = "SELECT domainName FROM tblDomains WHERE (status='True')"
            cmdDomainRights.Parameters.Clear()
            reader = cmdDomainRights.ExecuteReader
            If reader.HasRows Then
                While reader.Read
                    Me.cboDomains.Items.Add(IIf(DBNull.Value.Equals(reader!domainName), "", reader!domainName))
                End While
            End If
            reader.Close()
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly + MsgBoxStyle.ApplicationModal, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub
    Private Sub frmDomainRights_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint
        If Me.IsMdiChild Then
            Dim PnlLoc As New Point
            PnlLoc.X = CInt((Me.Width - Me.pnlDomRights.Width) / 2)
            PnlLoc.Y = CInt((Me.Height - Me.pnlDomRights.Height) / 2.5)
            Me.pnlDomRights.Location = PnlLoc
        Else
            Me.pnlDomRights.Dock = DockStyle.Fill
        End If
    End Sub

    Private Sub btnClose_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub cboDomains_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboDomains.SelectedIndexChanged
        If Me.cboDomains.Text = "" Then
            Exit Sub
        End If
        loadAssignedModules()
        loadAllModules()
    End Sub
    Private Sub loadAllModules()
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            Me.lstModules.Items.Clear()
            dbconnection()
            cmdDomainRights.Connection = conn
            cmdDomainRights.CommandType = CommandType.Text
            cmdDomainRights.CommandText = "SELECT * FROM tblModules WHERE (status='True') ORDER BY modName"
            cmdDomainRights.Parameters.Clear()
            reader = cmdDomainRights.ExecuteReader
            If reader.HasRows Then
                While reader.Read
                    li = Me.lstModules.Items.Add(IIf(DBNull.Value.Equals(reader!modName), "", reader!modName))
                    li.Tag = (IIf(DBNull.Value.Equals(reader!modId), "", reader!modId))
                End While
            End If
            reader.Close()
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly + MsgBoxStyle.ApplicationModal, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub
    Private Sub loadAssignedModules()
        If Me.cboDomains.Text = "" Then
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            Me.lstModuleAlreday.Items.Clear()
            cmdDomainRights.Connection = conn
            cmdDomainRights.CommandType = CommandType.Text
            cmdDomainRights.CommandText = "SELECT * FROM vwDomainRights WHERE (rightAccess='True') AND " & _
                vbNewLine & "(domainName=@domainName) AND (modStatus='True') AND (domainStatus='True') ORDER BY modName"
            cmdDomainRights.Parameters.Clear()
            cmdDomainRights.Parameters.AddWithValue("@domainName", Me.cboDomains.Text.Trim)
            reader = cmdDomainRights.ExecuteReader
            If reader.HasRows Then
                While reader.Read
                    li = Me.lstModuleAlreday.Items.Add(IIf(DBNull.Value.Equals(reader!modName), "", reader!modName))
                    li.Tag = (IIf(DBNull.Value.Equals(reader!modId), "", reader!modId))
                End While
            End If
            reader.Close()
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly + MsgBoxStyle.ApplicationModal, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub

    Private Sub lstModules_ColumnClick(ByVal sender As Object, ByVal e As System.Windows.Forms.ColumnClickEventArgs) Handles lstModules.ColumnClick
        If (e.Column() = 0) And (Me.lstModules.Items.Count > 0) Then
            For Each Li As ListViewItem In Me.lstModules.Items
                Li.Checked = Not (Li.Checked)
            Next
        End If
    End Sub

    Private Sub btnRightAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRightAll.Click
        If Me.cboDomains.Text = "" Then
            MsgBox("Select DomainTo Update First", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly + MsgBoxStyle.ApplicationModal, "Error Detected")
            Exit Sub
        End If
        If Me.lstModules.CheckedItems.Count > 1 Then
            i = 0
            maxrec = Me.lstModules.CheckedItems.Count - 1
            maxrec1 = Me.lstModuleAlreday.Items.Count - 1
            If maxrec >= 0 Then
                For i = 0 To maxrec
                    j = 0
                    If maxrec1 >= 0 Then
                        available = False
                        For j = 0 To maxrec1
                            If Me.lstModules.CheckedItems(i).Text.Trim = Me.lstModuleAlreday.Items(j).Text.Trim Then
                                available = True
                            Else
                            End If
                        Next
                        If available = False Then
                            Me.lstModuleAlreday.Items.Add(Me.lstModules.CheckedItems(i).Text)
                        End If

                    Else
                        Me.lstModuleAlreday.Items.Add(Me.lstModules.CheckedItems(i).Text)
                    End If
                    available = False
                Next
            End If
        End If
    End Sub

    Private Sub btnRightOne_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRightOne.Click
        If Me.cboDomains.Text = "" Then
            MsgBox("Select DomainTo Update First", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly + MsgBoxStyle.ApplicationModal, "Error Detected")
            Exit Sub
        End If
        If Me.lstModules.CheckedItems.Count = 1 Then
            i = 0
            maxrec = Me.lstModuleAlreday.Items.Count - 1
            available = False
            For i = 0 To maxrec
                If Me.lstModules.CheckedItems(0).Text.Trim = Me.lstModuleAlreday.Items(i).Text Then
                    available = True
                End If
            Next
            If available = False Then
                Me.lstModuleAlreday.Items.Add(Me.lstModules.CheckedItems(0).Text)
            End If
            available = False
        End If
    End Sub

    Private Sub btnReload_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnReload.Click
        Me.cboDomains.Text = ""
        Me.lstModuleAlreday.Items.Clear()
        loadCombos()
        loadAllModules()
    End Sub

    Private Sub btnLeftOne_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLeftOne.Click
        If Me.cboDomains.Text = "" Then
            MsgBox("Select DomainTo Update First", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly + MsgBoxStyle.ApplicationModal, "Error Detected")
            Exit Sub
        End If
        If Me.lstModuleAlreday.CheckedItems.Count = 1 Then
            Me.lstModuleAlreday.CheckedItems(0).Remove()
        End If
    End Sub

    Private Sub btnLeftAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLeftAll.Click
        If Me.cboDomains.Text = "" Then
            MsgBox("Select DomainTo Update First", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly + MsgBoxStyle.ApplicationModal, "Error Detected")
            Exit Sub
        End If
        maxrec = Me.lstModuleAlreday.CheckedItems.Count - 1
        If Me.lstModuleAlreday.CheckedItems.Count > 1 Then
            i = 0
            For i = 0 To maxrec
                Me.lstModuleAlreday.CheckedItems(0).Remove()
            Next
        End If
    End Sub

    Private Sub lstModuleAlreday_ColumnClick(ByVal sender As Object, ByVal e As System.Windows.Forms.ColumnClickEventArgs) Handles lstModuleAlreday.ColumnClick
        If (e.Column() = 0) And (Me.lstModuleAlreday.Items.Count > 0) Then
            For Each Li As ListViewItem In Me.lstModuleAlreday.Items
                Li.Checked = Not (Li.Checked)
            Next
        End If
    End Sub

    Private Sub btnUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdate.Click
        If Me.lstModuleAlreday.Items.Count <= 0 Then
            MsgBox("No modules to update", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly + MsgBoxStyle.ApplicationModal, "Error Detected")
            Exit Sub
        ElseIf Me.cboDomains.Text.Trim.Length <= 0 Then
            MsgBox("Select domain to Update", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly + MsgBoxStyle.ApplicationModal, "Error Detected")
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            i = 0
            cmdDomainRights.CommandText = "SELECT * FROM vwDomainRights WHERE (modStatus='True') " & _
                vbNewLine & " AND (domainStatus='True') AND (domainName=@domainName)"
            cmdDomainRights.CommandType = CommandType.Text
            cmdDomainRights.Connection = conn
            cmdDomainRights.Parameters.Clear()
            cmdDomainRights.Parameters.AddWithValue("@domainName", Me.cboDomains.Text.Trim)
            reader = cmdDomainRights.ExecuteReader


            If reader.HasRows Then
                While reader.Read
                    rightsId = (IIf(DBNull.Value.Equals(reader!domRightsId), "", reader!domRightsId))
                        Dim conn2 As New SqlConnection("Server=" & My.Settings.serverName.Trim & ";User ID=" & My.Settings.UserName.Trim & ";Database=" & My.Settings.dbName.Trim & ";Password=" & My.Settings.PassWord.Trim & "")
                        conn2.Open()
                        cmdDomainRights1.CommandText = "UPDATE tblDomainRights SET rightAccess='False' WHERE (domRightsId=@domRightsId)"
                        cmdDomainRights1.CommandType = CommandType.Text
                        cmdDomainRights1.Connection = conn2
                        cmdDomainRights1.Parameters.Clear()
                        cmdDomainRights1.Parameters.AddWithValue("@domRightsId", rightsId)
                        cmdDomainRights1.ExecuteNonQuery()
                        conn2.Close()
                End While
            End If
            reader.Close()
            i = 0
            For i = 0 To Me.lstModuleAlreday.Items.Count - 1
                cmdDomainRights.CommandText = "sprocDomainsRight"
                cmdDomainRights.CommandType = CommandType.StoredProcedure
                cmdDomainRights.Connection = conn
                cmdDomainRights.Parameters.Clear()
                cmdDomainRights.Parameters.AddWithValue("@moduleName", Me.lstModuleAlreday.Items(i).Text.Trim)
                cmdDomainRights.Parameters.AddWithValue("@userName", userName)
                cmdDomainRights.Parameters.AddWithValue("@domainName", Me.cboDomains.Text.Trim)
                cmdDomainRights.Parameters.AddWithValue("@dateReg", Date.Now)
                cmdDomainRights.ExecuteNonQuery()
                rec += 1
            Next
            If rec > 0 Then
                MsgBox("Record Updated", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "SuccessFul Transaction")
            End If
            frmHome.NaviBarHome_ActiveBandChanged(frmHome.NaviBarHome, e)
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly + MsgBoxStyle.ApplicationModal, "Error Detected")
        Finally
            loadAssignedModules()
            loadAllModules()
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub
End Class