Imports System.Data.SqlClient
Public Class frmInvIssueRequest
    Dim reader As SqlDataReader
    Dim rec As Integer
    Dim cmdIssueRequest As New SqlCommand
    Dim issueNo As Integer = 0
    Private Sub frmInvIssueRequest_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            loadIssueNo(Me.dtpReqDate.Value.Month, Me.dtpReqDate.Value.Year)
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
        Me.cboItemCategory.Items.Clear()
        Me.cboItemCategory.Text = ""
        Me.cboItemCategory.SelectedIndex = -1

        Me.cboIssueToType.Items.Clear()
        Me.cboIssueToType.Text = ""
        Me.cboIssueToType.SelectedIndex = -1

        Me.cmdIssueRequest.Connection = conn
        Me.cmdIssueRequest.CommandType = CommandType.Text
        Me.cmdIssueRequest.CommandText = "SELECT categoryName FROM tblInvMasterCategory WHERE (categoryName IS NOT NULL) ORDER BY categoryName"
        Me.cmdIssueRequest.Parameters.Clear()
        reader = Me.cmdIssueRequest.ExecuteReader
        While reader.Read
            Me.cboItemCategory.Items.Add(IIf(DBNull.Value.Equals(reader!categoryName), "", reader!categoryName))
        End While
        reader.Close()

        Me.cmdIssueRequest.CommandText = "SELECT DISTINCT issuedToTypeId,issuedToTypeName FROM tblInvIssueToTypeName WHERE " & _
            vbNewLine & " (issuedToTypeName IS NOT NULL) ORDER BY issuedToTypeId"
        Me.cmdIssueRequest.Parameters.Clear()
        reader = Me.cmdIssueRequest.ExecuteReader
        While reader.Read
            Me.cboIssueToType.Items.Add(IIf(DBNull.Value.Equals(reader!issuedToTypeName), "", reader!issuedToTypeName))
        End While
        reader.Close()
    End Sub
    Private Sub loadIssueNo(ByVal month As Integer, ByVal year As Integer)
        issueNo = 0
        Me.txtIssueNo.Clear()
        Me.cmdIssueRequest.Connection = conn
        Me.cmdIssueRequest.CommandType = CommandType.Text
        Me.cmdIssueRequest.CommandText = "SELECT ISNULL(MAX(issueNo),0)+1 AS issueNo FROM tblInvIssues " & _
            vbNewLine & " WHERE (month=@month) AND (year=@year)"
        Me.cmdIssueRequest.Parameters.Clear()
        Me.cmdIssueRequest.Parameters.AddWithValue("@month", month)
        Me.cmdIssueRequest.Parameters.AddWithValue("@year", year)
        reader = Me.cmdIssueRequest.ExecuteReader
        While reader.Read
            issueNo = IIf(DBNull.Value.Equals(reader!issueNo), "", reader!issueNo)
            Me.txtIssueNo.Text = "INVISS" & year.ToString("0000") & month.ToString("00") & "/" & issueNo.ToString("0000")
        End While
        reader.Close()
    End Sub
    Private Sub frmInvIssueRequest_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint
        If Me.IsMdiChild Then
            Dim PnlLoc As New Point
            PnlLoc.X = CInt((Me.Width - Me.pnlIssueReq.Width) / 2)
            PnlLoc.Y = CInt((Me.Height - Me.pnlIssueReq.Height) / 2.5)
            Me.pnlIssueReq.Location = PnlLoc
        Else
            Me.pnlIssueReq.Dock = DockStyle.Fill
        End If
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub dtpReqDate_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtpReqDate.ValueChanged
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            loadIssueNo(Me.dtpReqDate.Value.Month, Me.dtpReqDate.Value.Year)
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub

    Private Sub cboItemCategory_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboItemCategory.SelectedIndexChanged
        If Me.cboItemCategory.Text.Trim.Length <= 0 Then
            Exit Sub
        End If
        Me.cboItemCode.Items.Clear()
        Me.cboItemCode.Text = ""
        Me.cboItemCode.SelectedIndex = -1
        Me.txtItemName.Text = ""
        Me.txtCostPerUnit.Text = ""
        Me.txtQtyInStock.Text = ""
        Me.txtIssueQty.Text = ""
        Me.txtUOM.Text = ""
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            Me.cmdIssueRequest.Connection = conn
            Me.cmdIssueRequest.CommandType = CommandType.Text
            Me.cmdIssueRequest.CommandText = "SELECT DISTINCT itemCode FROM vwInvMaster WHERE (categoryName=@categoryName) ORDER BY itemCode"
            Me.cmdIssueRequest.Parameters.Clear()
            Me.cmdIssueRequest.Parameters.AddWithValue("@categoryName", Me.cboItemCategory.Text.Trim)
            reader = Me.cmdIssueRequest.ExecuteReader
            While reader.Read
                Me.cboItemCode.Items.Add(IIf(DBNull.Value.Equals(reader!itemCode), "", reader!itemCode))
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

    Private Sub cboItemCode_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboItemCode.SelectedIndexChanged
        If Me.cboItemCode.Text.Trim.Length <= 0 Then
            Exit Sub
        End If
        Me.txtItemName.Text = ""
        Me.txtCostPerUnit.Text = ""
        Me.txtQtyInStock.Text = ""
        Me.txtIssueQty.Text = ""
        Me.txtUOM.Text = ""
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            Me.cmdIssueRequest.Connection = conn
            Me.cmdIssueRequest.CommandType = CommandType.Text
            Me.cmdIssueRequest.CommandText = "SELECT itemDescription,costPerUnit,qtyInStock FROM vwInvMaster WHERE (itemCode=@itemCode)"
            Me.cmdIssueRequest.Parameters.Clear()
            Me.cmdIssueRequest.Parameters.AddWithValue("@itemCode", Me.cboItemCode.Text.Trim)
            reader = Me.cmdIssueRequest.ExecuteReader
            While reader.Read
                Me.txtItemName.Text = (IIf(DBNull.Value.Equals(reader!itemDescription), "", reader!itemDescription))
            End While
            reader.Close()
            Me.txtQtyInStock.Text = loadItemQty(Me.cboItemCode.Text.Trim)
            Me.txtCostPerUnit.Text = loadCostPerUnit(Me.cboItemCode.Text.Trim)
            Me.txtUOM.Text = loadUOM(Me.cboItemCode.Text.Trim)
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub
    Private Function loadUOM(ByVal itemCode As String)
        Dim uom As String = Nothing
        Me.cmdIssueRequest.Connection = conn
        Me.cmdIssueRequest.CommandType = CommandType.Text
        Me.cmdIssueRequest.CommandText = "SELECT unitOfMeasure FROM vwInvMaster WHERE (itemCode=@itemCode)"
        Me.cmdIssueRequest.Parameters.Clear()
        Me.cmdIssueRequest.Parameters.AddWithValue("@itemCode", itemCode)
        reader = Me.cmdIssueRequest.ExecuteReader
        If reader.HasRows = True Then
            While reader.Read
                uom = (IIf(DBNull.Value.Equals(reader!unitOfMeasure), "", reader!unitOfMeasure))
            End While
        ElseIf reader.HasRows = False Then
            uom = 0
        End If
        reader.Close()
        Return uom
    End Function
    Private Function loadItemQty(ByVal itemCode As String)
        Dim itemQty As Double = 0
        Me.cmdIssueRequest.Connection = conn
        Me.cmdIssueRequest.CommandType = CommandType.Text
        Me.cmdIssueRequest.CommandText = "SELECT qtyInStock FROM vwInvMaster WHERE (itemCode=@itemCode)"
        Me.cmdIssueRequest.Parameters.Clear()
        Me.cmdIssueRequest.Parameters.AddWithValue("@itemCode", itemCode)
        reader = Me.cmdIssueRequest.ExecuteReader
        If reader.HasRows = True Then
            While reader.Read
                itemQty = (IIf(DBNull.Value.Equals(reader!qtyInStock), "", reader!qtyInStock))
            End While
        ElseIf reader.HasRows = False Then
            itemQty = 0
        End If
        reader.Close()
        Return itemQty
    End Function
    Private Function loadCostPerUnit(ByVal itemCode As String)
        Dim unitCost As Double = 0
        Me.cmdIssueRequest.Connection = conn
        Me.cmdIssueRequest.CommandType = CommandType.Text
        Me.cmdIssueRequest.CommandText = "SELECT costPerUnit FROM vwInvMaster WHERE (itemCode=@itemCode)"
        Me.cmdIssueRequest.Parameters.Clear()
        Me.cmdIssueRequest.Parameters.AddWithValue("@itemCode", itemCode)
        reader = Me.cmdIssueRequest.ExecuteReader
        If reader.HasRows = True Then
            While reader.Read
                unitCost = (IIf(DBNull.Value.Equals(reader!costperunit), "", reader!costperunit))
            End While
        ElseIf reader.HasRows = False Then
            unitCost = 0
        End If
        reader.Close()
        Return unitCost
    End Function

    Private Sub cboIssueToType_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboIssueToType.SelectedIndexChanged
        If Me.cboIssueToType.Text.Trim.Length <= 0 Then
            Exit Sub
        End If
        Me.cboIssueTypeNo.Items.Clear()
        Me.cboIssueTypeNo.Text = ""
        Me.cboIssueTypeNo.SelectedIndex = -1

        Me.txtEmployeeName.Text = ""
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            If Me.cboIssueToType.Text.Trim = "TEACHER" Then
                Me.txtEmployeeName.ReadOnly = True
                Me.cboIssueTypeNo.DropDownStyle = ComboBoxStyle.DropDownList
                Me.cboIssueTypeNo.Enabled = True
                Me.cmdIssueRequest.Connection = conn
                Me.cmdIssueRequest.CommandType = CommandType.Text
                Me.cmdIssueRequest.CommandText = "SELECT empNo FROM tblSchoolStaff WHERE (empType='Teaching') AND (status=1) ORDER BY empNo"
                Me.cmdIssueRequest.Parameters.Clear()
                reader = Me.cmdIssueRequest.ExecuteReader
                While reader.Read
                    Me.cboIssueTypeNo.Items.Add(IIf(DBNull.Value.Equals(reader!empNo), "", reader!empNo))
                End While
            ElseIf Me.cboIssueToType.Text.Trim = "NON-TEACHER" Then
                Me.txtEmployeeName.ReadOnly = True
                Me.cboIssueTypeNo.DropDownStyle = ComboBoxStyle.DropDownList
                Me.cboIssueTypeNo.Enabled = True
                Me.cmdIssueRequest.Connection = conn
                Me.cmdIssueRequest.CommandType = CommandType.Text
                Me.cmdIssueRequest.CommandText = "SELECT empNo FROM tblSchoolStaff WHERE (empType='Non-Teaching') AND (status=1) ORDER BY empNo"
                Me.cmdIssueRequest.Parameters.Clear()
                reader = Me.cmdIssueRequest.ExecuteReader
                While reader.Read
                    Me.cboIssueTypeNo.Items.Add(IIf(DBNull.Value.Equals(reader!empNo), "", reader!empNo))
                End While
            ElseIf Me.cboIssueToType.Text = "STUDENT" Then
                Me.txtEmployeeName.ReadOnly = True
                Me.cboIssueTypeNo.Enabled = True
                Me.cboIssueTypeNo.DropDownStyle = ComboBoxStyle.Simple
            ElseIf Me.cboIssueToType.Text = "OTHERS" Then
                Me.cboIssueTypeNo.DropDownStyle = ComboBoxStyle.Simple
                Me.cboIssueTypeNo.Enabled = False
                Me.cboIssueTypeNo.Text = "OTHERS"
                Me.txtEmployeeName.ReadOnly = False
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

    Private Sub cboIssueTypeNo_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboIssueTypeNo.LostFocus
        If Me.cboIssueTypeNo.Text.Trim.Length <= 0 Then
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            If Me.cboIssueToType.Text = "STUDENT" Then
                Me.txtEmployeeName.Text = ""
                Me.cmdIssueRequest.Connection = conn
                Me.cmdIssueRequest.CommandType = CommandType.Text
                Me.cmdIssueRequest.CommandText = "SELECT FullName FROM tblStudDetails WHERE (status=1) AND (admNo=@admNo)"
                Me.cmdIssueRequest.Parameters.Clear()
                Me.cmdIssueRequest.Parameters.AddWithValue("@admNo", Me.cboIssueTypeNo.Text.Trim)
                reader = Me.cmdIssueRequest.ExecuteReader
                If reader.HasRows = True Then
                    While reader.Read
                        Me.txtEmployeeName.Text = IIf(DBNull.Value.Equals(reader!FullName), "", reader!FullName)
                    End While
                ElseIf reader.HasRows = False Then
                    Me.txtItemName.Text = ""
                    MsgBox("Student Not In The System", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "System Reply")
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

    Private Sub cboIssueTypeNo_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboIssueTypeNo.SelectedIndexChanged
        If Me.cboIssueToType.Text.Trim.Length <= 0 Then
            Exit Sub
        End If
        Me.txtEmployeeName.Text = ""
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            If Me.cboIssueToType.Text.Trim = "TEACHER" Then
                Me.cmdIssueRequest.Connection = conn
                Me.cmdIssueRequest.CommandType = CommandType.Text
                Me.cmdIssueRequest.CommandText = "SELECT FullName FROM tblSchoolStaff WHERE (empType='Teaching') AND " & _
                    vbNewLine & " (status=1) AND (empNo=@empNo)"
                Me.cmdIssueRequest.Parameters.Clear()
                Me.cmdIssueRequest.Parameters.AddWithValue("@empNo", Me.cboIssueTypeNo.Text.Trim)
                reader = Me.cmdIssueRequest.ExecuteReader
                While reader.Read
                    Me.txtEmployeeName.Text = (IIf(DBNull.Value.Equals(reader!FullName), "", reader!FullName))
                End While
            ElseIf Me.cboIssueToType.Text.Trim = "NON-TEACHER" Then
                Me.cmdIssueRequest.Connection = conn
                Me.cmdIssueRequest.CommandType = CommandType.Text
                Me.cmdIssueRequest.CommandText = "SELECT FullName FROM tblSchoolStaff WHERE (empType='Non-Teaching') AND " & _
                    vbNewLine & " (status=1) AND (empNo=@empNo)"
                Me.cmdIssueRequest.Parameters.Clear()
                Me.cmdIssueRequest.Parameters.AddWithValue("@empNo", Me.cboIssueTypeNo.Text.Trim)
                reader = Me.cmdIssueRequest.ExecuteReader
                While reader.Read
                    Me.txtEmployeeName.Text = (IIf(DBNull.Value.Equals(reader!FullName), "", reader!FullName))
                End While
            ElseIf Me.cboIssueToType.Text = "STUDENT" Then
                Me.txtEmployeeName.Text = ""
            ElseIf Me.cboIssueToType.Text = "OTHERS" Then
                Me.txtEmployeeName.Text = "OTHERS"
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

    Private Sub cboIssueTypeNo_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboIssueTypeNo.TextChanged
        Me.txtEmployeeName.Text = ""
    End Sub

    Private Sub txtIssueQty_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtIssueQty.TextChanged
        If IsNumeric(Me.txtIssueQty.Text.Trim) = False And Not (Me.txtIssueQty.Text.Trim.Length <= 0) Then
            MsgBox("Non Numeric Values Detected.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Me.txtIssueQty.Text = ""
            Exit Sub
        End If
        If Me.txtCostPerUnit.Text.Trim.Length > 0 And Me.txtIssueQty.Text.Trim.Length > 0 Then
            Me.txtTotalItemCost.Text = CDbl(Me.txtIssueQty.Text.Trim) * CDbl(Me.txtCostPerUnit.Text.Trim)
        Else
            Me.txtTotalItemCost.Text = ""
        End If
    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        If Me.lstIssueRequest.Items.Count <= 0 Then
            MsgBox("Missing items in the List.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            Dim result As MsgBoxResult = MsgBox("Issue Request?", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.YesNo, "Confirm Transaction")
            If result = MsgBoxResult.No Then
                Exit Sub
            End If

            i = 0

            For i = 0 To Me.lstIssueRequest.Items.Count - 1
                Me.cmdIssueRequest.Connection = conn
                Me.cmdIssueRequest.CommandType = CommandType.StoredProcedure
                Me.cmdIssueRequest.CommandText = "sprocInvIssueRequest"
                Me.cmdIssueRequest.Parameters.Clear()
                Me.cmdIssueRequest.Parameters.AddWithValue("@itemCode", Me.lstIssueRequest.Items(i).SubItems(3).Tag)
                Me.cmdIssueRequest.Parameters.AddWithValue("@issueNo", issueNo)
                Me.cmdIssueRequest.Parameters.AddWithValue("@issueNumber", Me.lstIssueRequest.Items(i).Text.Trim)
                Me.cmdIssueRequest.Parameters.AddWithValue("@issuedToTypeName", Me.lstIssueRequest.Items(i).SubItems(3).Text.Trim)
                Me.cmdIssueRequest.Parameters.AddWithValue("@issueToName", Me.lstIssueRequest.Items(i).SubItems(5).Text.Trim)
                Me.cmdIssueRequest.Parameters.AddWithValue("@issuedToNo", Me.lstIssueRequest.Items(i).SubItems(4).Text.Trim)
                Me.cmdIssueRequest.Parameters.AddWithValue("@issueQty", Me.lstIssueRequest.Items(i).SubItems(7).Text.Trim)
                Me.cmdIssueRequest.Parameters.AddWithValue("@costPerUnit", Me.lstIssueRequest.Items(i).SubItems(8).Text.Trim)
                Me.cmdIssueRequest.Parameters.AddWithValue("@uom", Me.lstIssueRequest.Items(i).SubItems(2).Text.Trim)
                Me.cmdIssueRequest.Parameters.AddWithValue("@issueReqDate", Me.lstIssueRequest.Items(i).SubItems(2).Tag)
                Me.cmdIssueRequest.Parameters.AddWithValue("@doneBy", userName.Trim)
                rec = rec + Me.cmdIssueRequest.ExecuteNonQuery
            Next
            If rec > 0 Then
                MsgBox("Issue Request Done!", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "SuccessFull Transaction")
            End If
            Me.lstIssueRequest.Items.Clear()
            clearTexts()
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            loadIssueNo(Me.dtpReqDate.Value.Month, Me.dtpReqDate.Value.Year)
            loadCombos()
            Me.cboIssueTypeNo.Enabled = True
            Me.cboIssueTypeNo.DropDownStyle = ComboBoxStyle.DropDownList
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub
    Private Sub clearTexts()
        Me.cboItemCategory.Items.Clear()
        Me.cboItemCategory.Text = ""
        Me.cboItemCategory.SelectedIndex = -1

        Me.cboItemCode.Items.Clear()
        Me.cboItemCode.Text = ""
        Me.cboItemCode.SelectedIndex = -1

        Me.cboIssueToType.Items.Clear()
        Me.cboIssueToType.Text = ""
        Me.cboIssueToType.SelectedIndex = -1

        Me.cboIssueTypeNo.Items.Clear()
        Me.cboIssueTypeNo.Text = ""
        Me.cboIssueTypeNo.SelectedIndex = -1

        Me.txtCostPerUnit.Clear()
        Me.txtItemName.Clear()
        Me.txtIssueQty.Clear()
        Me.txtIssueNo.Clear()
        Me.txtQtyInStock.Clear()
        Me.txtEmployeeName.Clear()
        Me.txtTotalItemCost.Clear()
        Me.txtUOM.Clear()

        Me.dtpReqDate.Value = Date.Now
    End Sub
    Private Sub btnAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAdd.Click
        If Me.cboItemCode.Text.Trim.Length <= 0 Then
            MsgBox("Item Code Missing", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.txtIssueNo.Text.Trim.Length <= 0 Then
            MsgBox("Issue Number Missing", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboIssueToType.Text.Trim.Length <= 0 Then
            MsgBox("Issue To Type Missing", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboIssueTypeNo.Text.Trim.Length <= 0 Then
            MsgBox("Issue Type Number Missing", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.txtIssueQty.Text.Trim.Length <= 0 Then
            MsgBox("Issue Quantity Missing", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.txtCostPerUnit.Text.Trim.Length <= 0 Then
            MsgBox("Cost Per Unit Missing", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.txtUOM.Text.Trim.Length <= 0 Then
            MsgBox("UOM Missing", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            Dim issueDateString As String = Me.dtpReqDate.Value.Day.ToString("00") & "-" & Me.dtpReqDate.Value.Month.ToString("00") & "-" & _
                Me.dtpReqDate.Value.Year.ToString("00")

            li = Me.lstIssueRequest.Items.Add(Me.txtIssueNo.Text.Trim)
            li.SubItems.Add(Me.txtItemName.Text.Trim)
            li.SubItems.Add(Me.txtUOM.Text.Trim)
            li.SubItems.Add(Me.cboIssueToType.Text.Trim)
            li.SubItems.Add(Me.cboIssueTypeNo.Text.Trim)
            li.SubItems.Add(Me.txtEmployeeName.Text.Trim)
            li.SubItems.Add(issueDateString)
            li.SubItems.Add(Me.txtIssueQty.Text.Trim)
            li.SubItems.Add(Me.txtCostPerUnit.Text.Trim)
            li.SubItems.Add(Me.txtTotalItemCost.Text.Trim)
            li.Tag = issueNo
            li.SubItems(2).Tag = Me.dtpReqDate.Value.Date
            li.SubItems(3).Tag = Me.cboItemCode.Text.Trim

            Me.txtIssueQty.Clear()
            Me.txtQtyInStock.Clear()
            Me.txtCostPerUnit.Clear()
            Me.txtTotalItemCost.Clear()
            Me.txtUOM.Clear()

            Me.txtCostPerUnit.Text = loadCostPerUnit(Me.cboItemCode.Text.Trim)
            Me.txtQtyInStock.Text = loadItemQty(Me.cboItemCode.Text.Trim)
            Me.txtUOM.Text = loadUOM(Me.cboItemCode.Text.Trim)
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub

    Private Sub btnRemove_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRemove.Click
        If Me.lstIssueRequest.Items.Count <= 0 Then
            MsgBox("Missing Items In The List.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.lstIssueRequest.CheckedItems.Count <= 0 Then
            MsgBox("Missing Checked Items In The List.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If
        i = 0
        For i = 0 To Me.lstIssueRequest.CheckedItems.Count - 1
            Me.lstIssueRequest.CheckedItems(0).Remove()
        Next
    End Sub
End Class