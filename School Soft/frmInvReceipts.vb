Imports System.Data.SqlClient
Public Class frmInvReceipts
    Dim reader As SqlDataReader
    Dim rec As Integer
    Dim cmdInvReceipts As New SqlCommand
    Dim recNo As Integer = 0
    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub frmInvReceipts_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            loadRecNo(Me.dtpRecDate.Value.Month, Me.dtpRecDate.Value.Year)
            loadCombos()
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub
    Private Sub loadRecNo(ByVal month As Integer, ByVal year As Integer)
        recNo = 0
        Me.txtRecNo.Clear()
        Me.cmdInvReceipts.Connection = conn
        Me.cmdInvReceipts.CommandType = CommandType.Text
        Me.cmdInvReceipts.CommandText = "SELECT ISNULL(MAX(receiptNo),0)+1 AS receiptNo FROM tblInvReceipts " & _
            vbNewLine & " WHERE (month=@month) AND (year=@year)"
        Me.cmdInvReceipts.Parameters.Clear()
        Me.cmdInvReceipts.Parameters.AddWithValue("@month", month)
        Me.cmdInvReceipts.Parameters.AddWithValue("@year", year)
        reader = Me.cmdInvReceipts.ExecuteReader
        While reader.Read
            recNo = IIf(DBNull.Value.Equals(reader!receiptNo), "", reader!receiptNo)
            Me.txtRecNo.Text = "INVREC" & year.ToString("0000") & month.ToString("00") & "/" & recNo.ToString("0000")
        End While
        reader.Close()
    End Sub
    Private Sub loadCombos()
        Me.cboItemCategory.Items.Clear()
        Me.cboItemCategory.Text = ""
        Me.cboItemCategory.SelectedIndex = -1

        Me.cboSuppCat.Items.Clear()
        Me.cboSuppCat.Text = ""
        Me.cboSuppCat.SelectedIndex = -1

        Me.cmdInvReceipts.Connection = conn
        Me.cmdInvReceipts.CommandType = CommandType.Text
        Me.cmdInvReceipts.CommandText = "SELECT categoryName FROM tblInvMasterCategory WHERE (categoryName IS NOT NULL) ORDER BY categoryName"
        Me.cmdInvReceipts.Parameters.Clear()
        reader = Me.cmdInvReceipts.ExecuteReader
        While reader.Read
            Me.cboItemCategory.Items.Add(IIf(DBNull.Value.Equals(reader!categoryName), "", reader!categoryName))
        End While
        reader.Close()

        Me.cmdInvReceipts.CommandText = "SELECT DISTINCT vendorCategory FROM tblInvVendors WHERE (vendorCategory IS NOT NULL) ORDER BY vendorCategory"
        Me.cmdInvReceipts.Parameters.Clear()
        reader = Me.cmdInvReceipts.ExecuteReader
        While reader.Read
            Me.cboSuppCat.Items.Add(IIf(DBNull.Value.Equals(reader!vendorCategory), "", reader!vendorCategory))
        End While
        reader.Close()
    End Sub
    Private Sub frmInvReceipts_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint
        If Me.IsMdiChild Then
            Dim PnlLoc As New Point
            PnlLoc.X = CInt((Me.Width - Me.pnlInvReceipts.Width) / 2)
            PnlLoc.Y = CInt((Me.Height - Me.pnlInvReceipts.Height) / 2.5)
            Me.pnlInvReceipts.Location = PnlLoc
        Else
            Me.pnlInvReceipts.Dock = DockStyle.Fill
        End If
    End Sub

    Private Sub txtQtyRec_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtQtyRec.TextChanged
        If IsNumeric(Me.txtQtyRec.Text.Trim) = False And Not (Me.txtQtyRec.Text.Trim.Length <= 0) Then
            MsgBox("Non Numeric Values Detected.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Me.txtQtyRec.Text = ""
            Exit Sub
        End If
        If Me.txtCostPerUnit.Text.Trim.Length > 0 And Me.txtQtyRec.Text.Trim.Length > 0 Then
            Me.txtTotalCost.Text = CDbl(Me.txtQtyRec.Text.Trim) * CDbl(Me.txtCostPerUnit.Text.Trim)
        Else
            Me.txtTotalCost.Text = ""
        End If
    End Sub

    Private Sub txtCostPerUnit_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCostPerUnit.TextChanged
        If IsNumeric(Me.txtCostPerUnit.Text.Trim) = False And Not (Me.txtCostPerUnit.Text.Trim.Length <= 0) Then
            MsgBox("Non Numeric Values Detected.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Me.txtCostPerUnit.Text = ""
            Exit Sub
        End If
    End Sub

    Private Sub txtTotalCost_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtTotalCost.TextChanged
        If IsNumeric(Me.txtTotalCost.Text.Trim) = False And Not (Me.txtTotalCost.Text.Trim.Length <= 0) Then
            MsgBox("Non Numeric Values Detected.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Me.txtTotalCost.Text = ""
            Exit Sub
        End If
    End Sub

    Private Sub cboSuppCat_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboSuppCat.SelectedIndexChanged
        If Me.cboSuppCat.Text.Trim.Length <= 0 Then
            Exit Sub
        End If
        Me.cboSuppCode.Items.Clear()
        Me.cboSuppCode.Text = ""
        Me.cboSuppCode.SelectedIndex = -1
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            Me.cmdInvReceipts.Connection = conn
            Me.cmdInvReceipts.CommandType = CommandType.Text
            Me.cmdInvReceipts.CommandText = "SELECT DISTINCT vendorCode FROM tblInvVendors WHERE (vendorCategory=@vendorCategory) ORDER BY vendorCode"
            Me.cmdInvReceipts.Parameters.Clear()
            Me.cmdInvReceipts.Parameters.AddWithValue("@vendorCategory", Me.cboSuppCat.Text.Trim)
            reader = Me.cmdInvReceipts.ExecuteReader
            While reader.Read
                Me.cboSuppCode.Items.Add(IIf(DBNull.Value.Equals(reader!vendorCode), "", reader!vendorCode))
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

    Private Sub cboItemCategory_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboItemCategory.SelectedIndexChanged
        If Me.cboItemCategory.Text.Trim.Length <= 0 Then
            Exit Sub
        End If
        Me.cboItemCode.Items.Clear()
        Me.cboItemCode.Text = ""
        Me.cboItemCode.SelectedIndex = -1
        Me.txtItemName.Text = ""
        Me.txtCostPerUnit.Text = ""
        Me.txtStockQty.Text = ""
        Me.txtQtyRec.Text = ""
        Me.txtUOM.Text = ""
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            Me.cmdInvReceipts.Connection = conn
            Me.cmdInvReceipts.CommandType = CommandType.Text
            Me.cmdInvReceipts.CommandText = "SELECT DISTINCT itemCode FROM vwInvMaster WHERE (categoryName=@categoryName) ORDER BY itemCode"
            Me.cmdInvReceipts.Parameters.Clear()
            Me.cmdInvReceipts.Parameters.AddWithValue("@categoryName", Me.cboItemCategory.Text.Trim)
            reader = Me.cmdInvReceipts.ExecuteReader
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

    Private Sub cboSuppCode_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboSuppCode.SelectedIndexChanged
        If Me.cboSuppCode.Text.Trim.Length <= 0 Then
            Exit Sub
        End If
        Me.txtSuppName.Text = ""
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            Me.cmdInvReceipts.Connection = conn
            Me.cmdInvReceipts.CommandType = CommandType.Text
            Me.cmdInvReceipts.CommandText = "SELECT vendorName FROM tblInvVendors WHERE (vendorCode=@vendorCode)"
            Me.cmdInvReceipts.Parameters.Clear()
            Me.cmdInvReceipts.Parameters.AddWithValue("@vendorCode", Me.cboSuppCode.Text.Trim)
            reader = Me.cmdInvReceipts.ExecuteReader
            While reader.Read
                Me.txtSuppName.Text = (IIf(DBNull.Value.Equals(reader!vendorName), "", reader!vendorName))
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
        Me.txtStockQty.Text = ""
        Me.txtQtyRec.Text = ""
        Me.txtUOM.Text = ""
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            Me.cmdInvReceipts.Connection = conn
            Me.cmdInvReceipts.CommandType = CommandType.Text
            Me.cmdInvReceipts.CommandText = "SELECT itemDescription,costPerUnit,qtyInStock FROM vwInvMaster WHERE (itemCode=@itemCode)"
            Me.cmdInvReceipts.Parameters.Clear()
            Me.cmdInvReceipts.Parameters.AddWithValue("@itemCode", Me.cboItemCode.Text.Trim)
            reader = Me.cmdInvReceipts.ExecuteReader
            While reader.Read
                Me.txtItemName.Text = (IIf(DBNull.Value.Equals(reader!itemDescription), "", reader!itemDescription))
            End While
            reader.Close()
            Me.txtStockQty.Text = loadItemQty(Me.cboItemCode.Text.Trim)
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
    Private Function loadItemQty(ByVal itemCode As String)
        Dim itemQty As Double = 0
        Me.cmdInvReceipts.Connection = conn
        Me.cmdInvReceipts.CommandType = CommandType.Text
        Me.cmdInvReceipts.CommandText = "SELECT qtyInStock FROM vwInvMaster WHERE (itemCode=@itemCode)"
        Me.cmdInvReceipts.Parameters.Clear()
        Me.cmdInvReceipts.Parameters.AddWithValue("@itemCode", itemCode)
        reader = Me.cmdInvReceipts.ExecuteReader
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
        Me.cmdInvReceipts.Connection = conn
        Me.cmdInvReceipts.CommandType = CommandType.Text
        Me.cmdInvReceipts.CommandText = "SELECT costPerUnit FROM vwInvMaster WHERE (itemCode=@itemCode)"
        Me.cmdInvReceipts.Parameters.Clear()
        Me.cmdInvReceipts.Parameters.AddWithValue("@itemCode", itemCode)
        reader = Me.cmdInvReceipts.ExecuteReader
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
    Private Function loadUOM(ByVal itemCode As String)
        Dim uom As String = Nothing
        Me.cmdInvReceipts.Connection = conn
        Me.cmdInvReceipts.CommandType = CommandType.Text
        Me.cmdInvReceipts.CommandText = "SELECT unitOfMeasure FROM vwInvMaster WHERE (itemCode=@itemCode)"
        Me.cmdInvReceipts.Parameters.Clear()
        Me.cmdInvReceipts.Parameters.AddWithValue("@itemCode", itemCode)
        reader = Me.cmdInvReceipts.ExecuteReader
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
    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        If Me.lstInvReceipts.Items.Count <= 0 Then
            MsgBox("Missing items in the List.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            Dim result As MsgBoxResult = MsgBox("Save Record/s?", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.YesNo, "Confirm Transaction")
            If result = MsgBoxResult.No Then
                Exit Sub
            End If
            For i = 0 To Me.lstInvReceipts.Items.Count - 1
                Me.cmdInvReceipts.Connection = conn
                Me.cmdInvReceipts.CommandType = CommandType.StoredProcedure
                Me.cmdInvReceipts.CommandText = "sprocInvReceipts"
                Me.cmdInvReceipts.Parameters.Clear()
                Me.cmdInvReceipts.Parameters.AddWithValue("@vendorCode", Me.lstInvReceipts.Items(i).SubItems(3).Text.Trim)
                Me.cmdInvReceipts.Parameters.AddWithValue("@itemCode", Me.lstInvReceipts.Items(i).SubItems(1).Text.Trim)
                Me.cmdInvReceipts.Parameters.AddWithValue("@qtyRec", Me.lstInvReceipts.Items(i).SubItems(6).Text.Trim)
                Me.cmdInvReceipts.Parameters.AddWithValue("@costPerUnit", Me.lstInvReceipts.Items(i).SubItems(7).Text.Trim)
                Me.cmdInvReceipts.Parameters.AddWithValue("@dateOfReceipt", Me.lstInvReceipts.Items(i).SubItems(1).Tag)
                Me.cmdInvReceipts.Parameters.AddWithValue("@receivedBy", userName.Trim)
                Me.cmdInvReceipts.Parameters.AddWithValue("@uom", Me.lstInvReceipts.Items(i).SubItems(4).Text.Trim)
                Me.cmdInvReceipts.Parameters.AddWithValue("@receiptNo", Me.lstInvReceipts.Items(i).Tag)
                Me.cmdInvReceipts.Parameters.AddWithValue("@receiptNumber", Me.lstInvReceipts.Items(i).Text.Trim)
                rec = rec + Me.cmdInvReceipts.ExecuteNonQuery
            Next
            If rec > 0 Then
                MsgBox("Item/s Received!", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "SuccessFull Transaction")
            End If
            Me.lstInvReceipts.Items.Clear()
            clearTexts()
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            loadRecNo(Me.dtpRecDate.Value.Month, Me.dtpRecDate.Value.Year)

            loadCombos()
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub

    Private Sub dtpRecDate_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtpRecDate.ValueChanged
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            loadRecNo(Me.dtpRecDate.Value.Month, Me.dtpRecDate.Value.Year)
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub

    Private Sub btnAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAdd.Click
        If Me.cboSuppCat.Text.Trim.Length <= 0 Then
            MsgBox("Supplier Category Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboSuppCode.Text.Trim.Length <= 0 Then
            MsgBox("Supplier Code Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboItemCategory.Text.Trim.Length <= 0 Then
            MsgBox("Item Category Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboItemCode.Text.Trim.Length <= 0 Then
            MsgBox("Item Code Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.txtItemName.Text.Trim.Length <= 0 Then
            MsgBox("Item Name Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.txtSuppName.Text.Trim.Length <= 0 Then
            MsgBox("Supplier Name Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.txtQtyRec.Text.Trim.Length <= 0 Then
            MsgBox("Receipt Quantity Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.txtCostPerUnit.Text.Trim.Length <= 0 Then
            MsgBox("Cost Per Unit Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.txtTotalCost.Text.Trim.Length <= 0 Then
            MsgBox("Total Cost Missing.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If
        Try
            Dim dateString As String = Me.dtpRecDate.Value.Day.ToString("00") & "-" & Me.dtpRecDate.Value.Month.ToString("00") & _
                "-" & Me.dtpRecDate.Value.Year.ToString("00")
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()

            li = Me.lstInvReceipts.Items.Add(Me.txtRecNo.Text.Trim)
            li.SubItems.Add(Me.cboItemCode.Text.Trim)
            li.SubItems.Add(Me.txtItemName.Text.Trim)
            li.SubItems.Add(Me.cboSuppCode.Text.Trim)
            li.SubItems.Add(Me.txtUOM.Text.Trim)
            li.SubItems.Add(dateString)
            li.SubItems.Add(Me.txtQtyRec.Text.Trim)
            li.SubItems.Add(Me.txtCostPerUnit.Text.Trim)
            li.SubItems.Add(Me.txtTotalCost.Text.Trim)
            li.Tag = Me.recNo
            li.SubItems(1).Tag = Me.dtpRecDate.Value.Date

            Me.txtQtyRec.Clear()
            Me.txtStockQty.Clear()
            Me.txtCostPerUnit.Clear()
            Me.txtTotalCost.Clear()
            Me.txtUOM.Clear()

            Me.txtCostPerUnit.Text = loadCostPerUnit(Me.cboItemCode.Text.Trim)
            Me.txtStockQty.Text = loadItemQty(Me.cboItemCode.Text.Trim)
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
        If Me.lstInvReceipts.Items.Count <= 0 Then
            MsgBox("Missing Items In The List.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.lstInvReceipts.CheckedItems.Count <= 0 Then
            MsgBox("Missing Checked Items In The List.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If
        i = 0
        For i = 0 To Me.lstInvReceipts.CheckedItems.Count - 1
            Me.lstInvReceipts.CheckedItems(0).Remove()
        Next
    End Sub
    Private Sub clearTexts()
        Me.cboItemCategory.Items.Clear()
        Me.cboItemCategory.Text = ""
        Me.cboItemCategory.SelectedIndex = -1

        Me.cboItemCode.Items.Clear()
        Me.cboItemCode.Text = ""
        Me.cboItemCode.SelectedIndex = -1

        Me.cboSuppCat.Items.Clear()
        Me.cboSuppCat.Text = ""
        Me.cboSuppCat.SelectedIndex = -1

        Me.cboSuppCode.Items.Clear()
        Me.cboSuppCode.Text = ""
        Me.cboSuppCode.SelectedIndex = -1

        Me.txtCostPerUnit.Clear()
        Me.txtItemName.Clear()
        Me.txtQtyRec.Clear()
        Me.txtRecNo.Clear()
        Me.txtStockQty.Clear()
        Me.txtSuppName.Clear()
        Me.txtTotalCost.Clear()
        Me.txtUOM.Clear()

        Me.dtpRecDate.Value = Date.Now
    End Sub
End Class