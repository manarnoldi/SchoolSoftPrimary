Imports System.Data.SqlClient
Public Class frmInvMaster
    Dim reader As SqlDataReader
    Dim rec As Integer
    Dim cmdInvMaster As New SqlCommand
    Dim itemNumber As Integer = 0
    Dim returnable As Boolean
    Private Function loadItemCode(ByVal itemCategory As String)
        Dim itemCatAbbr As String = itemCategory.Substring(0, 3)
        Me.txtItemCode.Text = ""
        Me.cmdInvMaster.Connection = conn
        Me.cmdInvMaster.CommandType = CommandType.Text
        Me.cmdInvMaster.CommandText = "SELECT ISNULL(MAX(itemNumber),0)+1 AS newNumber FROM vwInvMaster WHERE (itemCatAbbr=@itemCatAbbr)"
        Me.cmdInvMaster.Parameters.Clear()
        Me.cmdInvMaster.Parameters.AddWithValue("@itemCatAbbr", itemCatAbbr)
        reader = Me.cmdInvMaster.ExecuteReader
        While reader.Read
            Me.txtItemCode.Text = "INV" & itemCatAbbr.ToUpper & "" & CInt(IIf(DBNull.Value.Equals(reader!newNumber), "", reader!newNumber)).ToString("0000")
            itemNumber = IIf(DBNull.Value.Equals(reader!newNumber), "", reader!newNumber)
        End While
        reader.Close()
    End Function
    Private Sub frmInvMaster_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            Me.cbReturnFalse.Checked = False
            Me.cbReturnTrue.Checked = False
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
    Private Function loadList(ByVal itemCategory As String)
        Me.lstInvMaster.Items.Clear()

        Me.cmdInvMaster.Connection = conn
        Me.cmdInvMaster.CommandType = CommandType.Text
        Me.cmdInvMaster.CommandText = "SELECT * FROM  vwInvMaster WHERE (categoryName=@itemCategory) ORDER BY itemCode"
        Me.cmdInvMaster.Parameters.Clear()
        Me.cmdInvMaster.Parameters.AddWithValue("@itemCategory", itemCategory.Trim)
        reader = Me.cmdInvMaster.ExecuteReader
        While reader.Read
            li = Me.lstInvMaster.Items.Add(IIf(DBNull.Value.Equals(reader!itemCode), "", reader!itemCode))
            li.SubItems.Add(IIf(DBNull.Value.Equals(reader!itemDescription), "", reader!itemDescription))
            li.SubItems.Add(IIf(DBNull.Value.Equals(reader!categoryName), "", reader!categoryName))
            li.SubItems.Add(IIf(DBNull.Value.Equals(reader!unitOfMeasure), "", reader!unitOfMeasure))
            li.SubItems.Add(IIf(DBNull.Value.Equals(reader!returnable), "", reader!returnable))
            li.SubItems.Add(IIf(DBNull.Value.Equals(reader!reorderLevel), "", reader!reorderLevel))
            li.SubItems.Add(IIf(DBNull.Value.Equals(reader!costPerUnit), "", reader!costPerUnit))
            li.SubItems.Add(IIf(DBNull.Value.Equals(reader!qtyInStock), "", reader!qtyInStock))
            li.SubItems.Add(IIf(DBNull.Value.Equals(reader!stockValue), "", reader!stockValue))
            li.Tag = IIf(DBNull.Value.Equals(reader!itemNumber), "", reader!itemNumber)
        End While
        reader.Close()
    End Function
    Private Sub loadCombos()
        Me.cboItemCategory.Items.Clear()
        Me.cboItemCategory.Text = ""
        Me.cboItemCategory.SelectedIndex = -1

        Me.cboUOM.Items.Clear()
        Me.cboUOM.Text = ""
        Me.cboUOM.SelectedIndex = -1

        Me.txtQtyInStock.Text = "0"

        Me.cmdInvMaster.Connection = conn
        Me.cmdInvMaster.CommandType = CommandType.Text
        Me.cmdInvMaster.CommandText = "SELECT DISTINCT categoryName FROM  tblInvMasterCategory ORDER BY categoryName"
        Me.cmdInvMaster.Parameters.Clear()
        reader = Me.cmdInvMaster.ExecuteReader
        While reader.Read
            Me.cboItemCategory.Items.Add(IIf(DBNull.Value.Equals(reader!categoryName), "", reader!categoryName))
        End While
        reader.Close()

        Me.cmdInvMaster.CommandText = "SELECT DISTINCT  unitOfMeasure FROM  tblInvMaster ORDER BY unitOfMeasure"
        Me.cmdInvMaster.Parameters.Clear()
        reader = Me.cmdInvMaster.ExecuteReader
        While reader.Read
            Me.cboUOM.Items.Add(IIf(DBNull.Value.Equals(reader!unitOfMeasure), "", reader!unitOfMeasure))
        End While
        reader.Close()
    End Sub

    Private Sub frmInvMaster_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint
        If Me.IsMdiChild Then
            Dim PnlLoc As New Point
            PnlLoc.X = CInt((Me.Width - Me.pnlInvMaster.Width) / 2)
            PnlLoc.Y = CInt((Me.Height - Me.pnlInvMaster.Height) / 2.5)
            Me.pnlInvMaster.Location = PnlLoc
        Else
            Me.pnlInvMaster.Dock = DockStyle.Fill
        End If
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub btnView_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnView.Click
        If Me.cboItemCategory.Text.Trim.Length <= 0 Then
            MsgBox("Item Category Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            loadList(Me.cboItemCategory.Text.Trim)
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub

    Private Sub CLOSEToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CLOSEToolStripMenuItem.Click
        Me.Close()
    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        If Me.cboItemCategory.Text.Trim.Length <= 0 Then
            MsgBox("Item Category Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.txtItemCode.Text.Trim.Length <= 0 Then
            MsgBox("Item Code Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.txtItemName.Text.Trim.Length <= 0 Then
            MsgBox("Item Name Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboUOM.Text.Trim.Length <= 0 Then
            MsgBox("Unit Of Measure Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.txtReorderLevel.Text.Trim.Length <= 0 Then
            MsgBox("Reorder Level Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.txtCostPerUnit.Text.Trim.Length <= 0 Then
            MsgBox("Cost Per Unit Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.txtQtyInStock.Text.Trim.Length <= 0 Then
            MsgBox("Quantity In Stock Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf (Me.cbReturnFalse.Checked = True) And (Me.cbReturnTrue.Checked = True) Then
            MsgBox("Check If Item Is Returnable Or Not.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf (Me.cbReturnFalse.Checked = False) And (Me.cbReturnTrue.Checked = False) Then
            MsgBox("Select One Option Only (Checked Or Unchecked).", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            Dim result As MsgBoxResult = MsgBox("Save Record?", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.YesNo, "Confirm Transaction")
            If result = MsgBoxResult.No Then
                Exit Sub
            End If
            If Me.cbReturnFalse.Checked = True Then
                returnable = False
            ElseIf Me.cbReturnTrue.Checked = True Then
                returnable = True
            End If
            rec = 0
            Me.cmdInvMaster.Connection = conn
            Me.cmdInvMaster.CommandType = CommandType.StoredProcedure
            Me.cmdInvMaster.CommandText = "sprocInvMaster"
            Me.cmdInvMaster.Parameters.Clear()
            Me.cmdInvMaster.Parameters.AddWithValue("@itemNumber", itemNumber)
            Me.cmdInvMaster.Parameters.AddWithValue("@itemCode", Me.txtItemCode.Text.Trim)
            Me.cmdInvMaster.Parameters.AddWithValue("@itemDescription", Me.txtItemName.Text.Trim)
            Me.cmdInvMaster.Parameters.AddWithValue("@invCategory", Me.cboItemCategory.Text.Trim)
            Me.cmdInvMaster.Parameters.AddWithValue("@unitOfMeasure", Me.cboUOM.Text.Trim)
            Me.cmdInvMaster.Parameters.AddWithValue("@reorderLevel", Me.txtReorderLevel.Text.Trim)
            Me.cmdInvMaster.Parameters.AddWithValue("@costPerUnit", Me.txtCostPerUnit.Text.Trim)
            Me.cmdInvMaster.Parameters.AddWithValue("@qtyInStock", Me.txtQtyInStock.Text.Trim)
            Me.cmdInvMaster.Parameters.AddWithValue("@returnable", returnable)
            Me.cmdInvMaster.Parameters.AddWithValue("@regBy", userName.Trim)
            Me.cmdInvMaster.Parameters.AddWithValue("@queryType", 1)
            rec = Me.cmdInvMaster.ExecuteNonQuery
            If rec > 0 Then
                MsgBox("Record Saved", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "SuccessFull Transaction")
            End If
            clearTexts()
            loadCombos()
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub
    Private Sub clearTexts()
        Me.lstInvMaster.Items.Clear()

        Me.cboItemCategory.Items.Clear()
        Me.cboItemCategory.Text = ""
        Me.cboItemCategory.SelectedIndex = -1

        Me.cboUOM.Items.Clear()
        Me.cboUOM.Text = ""
        Me.cboUOM.SelectedIndex = -1

        Me.cbReturnFalse.Checked = False
        Me.cbReturnTrue.Checked = False

        Me.txtCostPerUnit.Text = ""
        Me.txtItemCode.Text = ""
        Me.txtItemName.Text = ""
        Me.txtQtyInStock.Text = ""
        Me.txtReorderLevel.Text = ""
        Me.txtSearchItemName.Text = ""
    End Sub
    Private Sub cboItemCategory_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboItemCategory.SelectedIndexChanged
        Me.lstInvMaster.Items.Clear()
        If Me.cboItemCategory.Text.Trim.Length <= 0 Then
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            loadItemCode(Me.cboItemCategory.Text.Trim)
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub

    Private Sub txtReorderLevel_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtReorderLevel.TextChanged
        If IsNumeric(Me.txtReorderLevel.Text.Trim) = False And Not (Me.txtReorderLevel.Text.Trim.Length <= 0) Then
            MsgBox("Non Numeric Values Detected.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Me.txtReorderLevel.Text = ""
            Exit Sub
        End If
    End Sub

    Private Sub txtCostPerUnit_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCostPerUnit.TextChanged
        If IsNumeric(Me.txtCostPerUnit.Text.Trim) = False And Not (Me.txtCostPerUnit.Text.Trim.Length <= 0) Then
            MsgBox("Non Numeric Values Detected.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Me.txtCostPerUnit.Text = ""
            Exit Sub
        End If
    End Sub

    Private Sub txtQtyInStock_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtQtyInStock.TextChanged
        If IsNumeric(Me.txtQtyInStock.Text.Trim) = False And Not (Me.txtQtyInStock.Text.Trim.Length <= 0) Then
            MsgBox("Non Numeric Values Detected.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Me.txtQtyInStock.Text = ""
            Exit Sub
        End If
    End Sub

    Private Sub btnUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdate.Click
        If Me.cboItemCategory.Text.Trim.Length <= 0 Then
            MsgBox("Item Category Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.txtItemCode.Text.Trim.Length <= 0 Then
            MsgBox("Item Code Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.txtItemName.Text.Trim.Length <= 0 Then
            MsgBox("Item Name Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboUOM.Text.Trim.Length <= 0 Then
            MsgBox("Unit Of Measure Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.txtReorderLevel.Text.Trim.Length <= 0 Then
            MsgBox("Reorder Level Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.txtCostPerUnit.Text.Trim.Length <= 0 Then
            MsgBox("Cost Per Unit Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.txtQtyInStock.Text.Trim.Length <= 0 Then
            MsgBox("Quantity In Stock Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf (Me.cbReturnFalse.Checked = True) And (Me.cbReturnTrue.Checked = True) Then
            MsgBox("Check If Item Is Returnable Or Not.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf (Me.cbReturnFalse.Checked = False) And (Me.cbReturnTrue.Checked = False) Then
            MsgBox("Select One Option Only (Checked Or Unchecked).", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            Dim result As MsgBoxResult = MsgBox("Update Record?", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.YesNo, "Confirm Transaction")
            If result = MsgBoxResult.No Then
                Exit Sub
            End If
            If Me.cbReturnFalse.Checked = True Then
                returnable = False
            ElseIf Me.cbReturnTrue.Checked = True Then
                returnable = True
            End If
            rec = 0
            Me.cmdInvMaster.Connection = conn
            Me.cmdInvMaster.CommandType = CommandType.StoredProcedure
            Me.cmdInvMaster.CommandText = "sprocInvMaster"
            Me.cmdInvMaster.Parameters.Clear()
            Me.cmdInvMaster.Parameters.AddWithValue("@itemNumber", itemNumber)
            Me.cmdInvMaster.Parameters.AddWithValue("@itemCode", Me.txtItemCode.Text.Trim)
            Me.cmdInvMaster.Parameters.AddWithValue("@itemDescription", Me.txtItemName.Text.Trim)
            Me.cmdInvMaster.Parameters.AddWithValue("@invCategory", Me.cboItemCategory.Text.Trim)
            Me.cmdInvMaster.Parameters.AddWithValue("@unitOfMeasure", Me.cboUOM.Text.Trim)
            Me.cmdInvMaster.Parameters.AddWithValue("@reorderLevel", Me.txtReorderLevel.Text.Trim)
            Me.cmdInvMaster.Parameters.AddWithValue("@costPerUnit", Me.txtCostPerUnit.Text.Trim)
            Me.cmdInvMaster.Parameters.AddWithValue("@qtyInStock", Me.txtQtyInStock.Text.Trim)
            Me.cmdInvMaster.Parameters.AddWithValue("@returnable", returnable)
            Me.cmdInvMaster.Parameters.AddWithValue("@regBy", userName.Trim)
            Me.cmdInvMaster.Parameters.AddWithValue("@queryType", 2)
            rec = Me.cmdInvMaster.ExecuteNonQuery
            If rec > 0 Then
                MsgBox("Record Updated!", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "SuccessFull Transaction")
            End If
            clearTexts()
            loadCombos()
            Me.txtItemCode.Enabled = True
            Me.cboItemCategory.Enabled = True
            Me.btnSave.Enabled = True
            Me.btnUpdate.Enabled = False
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub

    Private Sub txtSearchItemName_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtSearchItemName.TextChanged
        Me.lstInvMaster.Items.Clear()
        If Me.txtSearchItemName.Text.Trim.Length <= 3 Then
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            Me.cmdInvMaster.Connection = conn
            Me.cmdInvMaster.CommandType = CommandType.Text
            Me.cmdInvMaster.CommandText = "SELECT * FROM  vwInvMaster WHERE (itemDescription LIKE @itemDescription) ORDER BY categoryName,itemCode"
            Me.cmdInvMaster.Parameters.Clear()
            Me.cmdInvMaster.Parameters.AddWithValue("@itemDescription", String.Format("%{0}%", Me.txtSearchItemName.Text.Trim))
            reader = Me.cmdInvMaster.ExecuteReader
            While reader.Read
                li = Me.lstInvMaster.Items.Add(IIf(DBNull.Value.Equals(reader!itemCode), "", reader!itemCode))
                li.SubItems.Add(IIf(DBNull.Value.Equals(reader!itemDescription), "", reader!itemDescription))
                li.SubItems.Add(IIf(DBNull.Value.Equals(reader!categoryName), "", reader!categoryName))
                li.SubItems.Add(IIf(DBNull.Value.Equals(reader!unitOfMeasure), "", reader!unitOfMeasure))
                li.SubItems.Add(IIf(DBNull.Value.Equals(reader!returnable), "", reader!returnable))
                li.SubItems.Add(IIf(DBNull.Value.Equals(reader!reorderLevel), "", reader!reorderLevel))
                li.SubItems.Add(IIf(DBNull.Value.Equals(reader!costPerUnit), "", reader!costPerUnit))
                li.SubItems.Add(IIf(DBNull.Value.Equals(reader!qtyInStock), "", reader!qtyInStock))
                li.SubItems.Add(IIf(DBNull.Value.Equals(reader!stockValue), "", reader!stockValue))
                li.Tag = IIf(DBNull.Value.Equals(reader!itemId), "", reader!itemId)
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

    Private Sub UPDATEToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UPDATEToolStripMenuItem.Click
        If Me.lstInvMaster.Items.Count <= 0 Then
            MsgBox("No items in the list.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.lstInvMaster.CheckedItems.Count <= 0 Then
            MsgBox("No checked items in the list.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.lstInvMaster.CheckedItems.Count > 1 Then
            MsgBox("Update one item at a time.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            Me.txtItemCode.Enabled = False
            Me.cboItemCategory.Enabled = False
            Me.btnSave.Enabled = False
            Me.btnUpdate.Enabled = True

            Me.cboItemCategory.Text = Me.lstInvMaster.CheckedItems(0).SubItems(2).Text.Trim
            Me.txtItemCode.Text = Me.lstInvMaster.CheckedItems(0).Text.Trim
            Me.txtItemName.Text = Me.lstInvMaster.CheckedItems(0).SubItems(1).Text.Trim
            Me.cboUOM.Text = Me.lstInvMaster.CheckedItems(0).SubItems(3).Text.Trim
            Me.txtReorderLevel.Text = Me.lstInvMaster.CheckedItems(0).SubItems(5).Text.Trim
            Me.txtCostPerUnit.Text = Me.lstInvMaster.CheckedItems(0).SubItems(6).Text.Trim
            Me.txtQtyInStock.Text = Me.lstInvMaster.CheckedItems(0).SubItems(7).Text.Trim

            Me.cbReturnTrue.Checked = False
            Me.cbReturnFalse.Checked = False
            If Me.lstInvMaster.CheckedItems(0).SubItems(4).Text.Trim = "True" Then
                Me.cbReturnTrue.Checked = True
            Else
                Me.cbReturnFalse.Checked = True
            End If
            Me.txtSearchItemName.Text = ""

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
            Me.txtItemCode.Enabled = True
            Me.cboItemCategory.Enabled = True
            Me.btnSave.Enabled = True
            Me.btnUpdate.Enabled = False
            loadCombos()
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
       
    End Sub

    Private Sub DELETEToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DELETEToolStripMenuItem.Click
        If Me.lstInvMaster.Items.Count <= 0 Then
            MsgBox("No items in the list.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.lstInvMaster.CheckedItems.Count <= 0 Then
            MsgBox("No checked items in the list.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            Dim result As MsgBoxResult = MsgBox("Delete Record/s?", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.YesNo, "Confirm Transaction")
            If result = MsgBoxResult.No Then
                Exit Sub
            End If
            rec = 0
            i = 0
            For i = 0 To Me.lstInvMaster.CheckedItems.Count - 1
                Me.cmdInvMaster.Connection = conn
                Me.cmdInvMaster.CommandType = CommandType.StoredProcedure
                Me.cmdInvMaster.CommandText = "sprocInvMaster"
                Me.cmdInvMaster.Parameters.Clear()
                Me.cmdInvMaster.Parameters.AddWithValue("@itemNumber", Me.lstInvMaster.CheckedItems(i).Tag)
                Me.cmdInvMaster.Parameters.AddWithValue("@itemCode", Me.lstInvMaster.CheckedItems(i).Text.Trim)
                Me.cmdInvMaster.Parameters.AddWithValue("@itemDescription", Me.lstInvMaster.CheckedItems(i).SubItems(1).Text.Trim)
                Me.cmdInvMaster.Parameters.AddWithValue("@invCategory", Me.lstInvMaster.CheckedItems(i).SubItems(2).Text.Trim)
                Me.cmdInvMaster.Parameters.AddWithValue("@unitOfMeasure", Me.lstInvMaster.CheckedItems(i).SubItems(3).Text.Trim)
                Me.cmdInvMaster.Parameters.AddWithValue("@reorderLevel", Me.lstInvMaster.CheckedItems(i).SubItems(5).Text.Trim)
                Me.cmdInvMaster.Parameters.AddWithValue("@costPerUnit", Me.lstInvMaster.CheckedItems(i).SubItems(6).Text.Trim)
                Me.cmdInvMaster.Parameters.AddWithValue("@qtyInStock", Me.lstInvMaster.CheckedItems(i).SubItems(7).Text.Trim)
                Me.cmdInvMaster.Parameters.AddWithValue("@returnable", Me.lstInvMaster.CheckedItems(i).SubItems(4).Text.Trim)
                Me.cmdInvMaster.Parameters.AddWithValue("@regBy", userName.Trim)
                Me.cmdInvMaster.Parameters.AddWithValue("@queryType", 3)
                rec = rec + Me.cmdInvMaster.ExecuteNonQuery
            Next
            If rec > 0 Then
                MsgBox("Record/s Deleted!", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "SuccessFull Transaction")
            End If
            clearTexts()
            loadCombos()
            Me.txtItemCode.Enabled = True
            Me.cboItemCategory.Enabled = True
            Me.cboUOM.Enabled = True
            Me.btnSave.Enabled = True
            Me.btnUpdate.Enabled = False
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub
End Class