Imports System.Data.SqlClient
Imports System.IO

Public Class frmStudImages
    Dim imageAvailable As Boolean = False
    Dim bitmap1 As Bitmap
    Private mImageFile As Image
    Private mImageFilePath As String
    Dim rec As Integer = 0
    Dim reader As SqlDataReader
    Dim cmdStudImages As New SqlCommand
    Dim queryType As String = Nothing
    Dim recordExists As Boolean = True
    Dim queryOk As Boolean = False
    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub frmStudImages_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
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

    Private Sub frmStudImages_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint

        If Me.IsMdiChild Then
            Dim PnlLoc As New Point
            PnlLoc.X = CInt((Me.Width - Me.pnlStudImages.Width) / 2)
            PnlLoc.Y = CInt((Me.Height - Me.pnlStudImages.Height) / 2.5)
            Me.pnlStudImages.Location = PnlLoc
        Else
            Me.pnlStudImages.Dock = DockStyle.Fill
        End If

    End Sub

    Private Sub btnUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdate.Click
        If Me.cboClassName.Text.Trim.Length <= 0 Then
            MsgBox("Missing Class Name", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboClassStream.Text.Trim.Length <= 0 Then
            MsgBox("Missing Class Stream", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboClassYear.Text.Trim.Length <= 0 Then
            MsgBox("Missing Class Year", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboStudentNo.Text.Trim.Length <= 0 Then
            MsgBox("Missing Student Number", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.txtStudName.Text.Trim.Length <= 0 Then
            MsgBox("Missing Student Name", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf IsNothing(Me.pbStudImage.Image) = True Then
            MsgBox("Missing Image To Update", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            'recordExists = True
            'CheckForExistence()
            'If recordExists = True Then
            '    MsgBox("Student Already Has An Image Saved against them.Try Update", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            '    Exit Sub
            'End If
            'recordExists = True
            'checkImageForImage()
            'If recordExists = True Then
            '    MsgBox("Image Already In Database For Another Student.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            '    Exit Sub
            'End If
            Dim result As MsgBoxResult = MsgBox("Update Record?", MsgBoxStyle.Question + MsgBoxStyle.ApplicationModal + MsgBoxStyle.YesNo, "Confirm Transaction")
            If result = MsgBoxResult.No Then
                Exit Sub
            End If
            queryType = "UPDATE"
            UploadPhoto()
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub

    Private Sub btnClear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClear.Click
        Me.pbStudImage.Image = Nothing
        Me.txtPicName.Text = ""
        Me.cboStudentNo.Tag = Nothing
        Me.cboClassName.SelectedIndex = -1
        Me.cboClassStream.SelectedIndex = -1
        Me.cboClassYear.SelectedIndex = -1
        Me.cboStudentNo.SelectedIndex = -1
        queryType = Nothing
        Me.pbStudImage.Image = Nothing
        Me.btnSave.Enabled = True
        Me.btnUpdate.Enabled = False
        Me.btnDelete.Enabled = False
    End Sub
    Private Sub btnSelect_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSelect.Click
        Me.txtPicName.Text = ""
        fdlg.Title = "Select Photo"
        fdlg.InitialDirectory = "C:\Users\Public\Documents\School Soft\Student Images\"
        fdlg.Filter = "All Files|*.*|Bitmaps|*.bmp|GIFs|*.gif|JPEGs|*.jpg"
        fdlg.FilterIndex = 4
        fdlg.FileName = ""
        fdlg.RestoreDirectory = True
        If fdlg.ShowDialog() = Windows.Forms.DialogResult.OK Then
            Me.txtPicName.Tag = fdlg.FileName
            Me.txtPicName.Text = System.IO.Path.GetFileName(fdlg.FileName)

        Else
            Exit Sub
        End If

        Dim sFilePath As String
        sFilePath = fdlg.FileName
        If sFilePath = "" Then
            Exit Sub
        End If

        If System.IO.File.Exists(sFilePath) = False Then
            Exit Sub
        Else
            Me.txtPicName.Tag = sFilePath
            mImageFilePath = sFilePath
        End If
        pbStudImage.Image = Image.FromFile(Me.txtPicName.Tag)
        pbStudImage.SizeMode = PictureBoxSizeMode.StretchImage
    End Sub
    Private Sub btnRotate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRotate.Click
        bitmap1 = CType(pbStudImage.Image, Bitmap)
        If bitmap1 IsNot Nothing Then
            bitmap1.RotateFlip(RotateFlipType.Rotate90FlipXY)
            pbStudImage.Image = bitmap1
        End If
    End Sub
    Private Sub loadCombos()
        cboClassName.Items.Clear()
        cboClassStream.Items.Clear()
        cboClassYear.Items.Clear()
        cboClassName.SelectedIndex = -1
        cboClassStream.SelectedIndex = -1
        cboClassYear.SelectedIndex = -1
        cmdStudImages.CommandType = CommandType.Text
        cmdStudImages.Connection = conn
        cmdStudImages.CommandText = "SELECT DISTINCT className FROM tblClasses WHERE (status='True') ORDER BY ClassName"
        cmdStudImages.Parameters.Clear()
        reader = cmdStudImages.ExecuteReader
        If reader.HasRows Then
            While reader.Read
                Me.cboClassName.Items.Add(IIf(DBNull.Value.Equals(reader!className), "", reader!className))
            End While
        End If
        reader.Close()

        cmdStudImages.CommandText = "SELECT DISTINCT  stream FROM tblClasses WHERE (status='True') ORDER BY stream"
        cmdStudImages.Parameters.Clear()
        reader = cmdStudImages.ExecuteReader
        If reader.HasRows Then
            While reader.Read
                Me.cboClassStream.Items.Add(IIf(DBNull.Value.Equals(reader!stream), "", reader!stream))
            End While
        End If
        reader.Close()

        cmdStudImages.CommandText = "SELECT DISTINCT year FROM tblClasses WHERE (status='True') ORDER BY year"
        cmdStudImages.Parameters.Clear()
        reader = cmdStudImages.ExecuteReader
        If reader.HasRows Then
            While reader.Read
                Me.cboClassYear.Items.Add(IIf(DBNull.Value.Equals(reader!year), "", reader!year))
            End While
        End If
        reader.Close()
    End Sub

    Private Sub cboClassName_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboClassName.SelectedIndexChanged, cboClassStream.SelectedIndexChanged, cboClassYear.SelectedIndexChanged
        Me.cboStudentNo.Items.Clear()
        Me.cboStudentNo.SelectedIndex = -1
        Me.txtStudName.Text = ""
        Me.txtPicName.Text = ""
        Me.cboStudentNo.Tag = Nothing
        'Me.pbStudImage.Image = Nothing
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()

            If (Me.cboClassYear.Text.Trim.Length <= 0) Or (Me.cboClassName.Text.Trim.Length <= 0) Or (Me.cboClassStream.Text.Trim.Length <= 0) Then
                Exit Sub
            End If
            cmdStudImages.CommandType = CommandType.Text
            cmdStudImages.Connection = conn
            cmdStudImages.CommandText = "SELECT * FROM  vwStudClass WHERE (className=@className) AND (stream=@stream) AND " & _
                vbNewLine & " (year=@year) AND (classStatus='True') AND (classStudStatus='True') AND (studStatus='True') ORDER BY admNo"
            cmdStudImages.Parameters.Clear()
            cmdStudImages.Parameters.AddWithValue("@className", Me.cboClassName.Text.Trim)
            cmdStudImages.Parameters.AddWithValue("@stream", Me.cboClassStream.Text.Trim)
            cmdStudImages.Parameters.AddWithValue("@year", Me.cboClassYear.Text.Trim)
            reader = cmdStudImages.ExecuteReader
            If reader.HasRows Then
                While reader.Read
                    Me.cboStudentNo.Items.Add(IIf(DBNull.Value.Equals(reader!admNo), "", reader!admNo))
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

    Private Sub cboStudentNo_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboStudentNo.SelectedIndexChanged
        Me.txtStudName.Text = ""
        Me.pbStudImage.Image = Nothing
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()

            If (Me.cboClassYear.Text.Trim.Length <= 0) Or (Me.cboClassName.Text.Trim.Length <= 0) Or (Me.cboClassStream.Text.Trim.Length <= 0) Then
                Exit Sub
            End If
            cmdStudImages.CommandType = CommandType.Text
            cmdStudImages.Connection = conn
            cmdStudImages.CommandText = "SELECT * FROM tblStudDetails WHERE (admNo=@admNo) AND (status='True')"
            cmdStudImages.Parameters.Clear()
            cmdStudImages.Parameters.AddWithValue("@admNo", Me.cboStudentNo.Text.Trim)
            reader = cmdStudImages.ExecuteReader
            If reader.HasRows Then
                While reader.Read
                    Me.txtStudName.Text = (IIf(DBNull.Value.Equals(reader!FullName), "", reader!FullName))
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
    Private Sub CheckForExistence()
        cmdStudImages.Connection = conn
        cmdStudImages.CommandType = CommandType.Text
        cmdStudImages.CommandText = "SELECT * FROM  vwStudImages WHERE (imageStatus='True') " & _
            vbNewLine & "AND (admNo=@admNo) AND (studStatus='True')"
        cmdStudImages.Parameters.Clear()
        cmdStudImages.Parameters.AddWithValue("@admNo", Me.cboStudentNo.Text.Trim)
        reader = cmdStudImages.ExecuteReader
        If reader.HasRows Then
            recordExists = True
        Else
            recordExists = False
        End If
        reader.Close()
    End Sub
    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        If Me.cboClassName.Text.Trim.Length <= 0 Then
            MsgBox("Missing Class Name", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboClassStream.Text.Trim.Length <= 0 Then
            MsgBox("Missing Class Stream", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboClassYear.Text.Trim.Length <= 0 Then
            MsgBox("Missing Class Year", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboStudentNo.Text.Trim.Length <= 0 Then
            MsgBox("Missing Student Number", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.txtStudName.Text.Trim.Length <= 0 Then
            MsgBox("Missing Student Name", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf IsNothing(Me.pbStudImage.Image) = True Then
            MsgBox("Missing Image To Insert", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            recordExists = True
            CheckForExistence()
            If recordExists = True Then
                MsgBox("Student Already Has An Image Saved against them.Try Update", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
                Exit Sub
            End If
            'recordExists = True
            'checkImageForImage()
            'If recordExists = True Then
            '    MsgBox("Image Already In Database For Another Student.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            '    Exit Sub
            'End If
            Dim result As MsgBoxResult = MsgBox("Save Record?", MsgBoxStyle.Question + MsgBoxStyle.ApplicationModal + MsgBoxStyle.YesNo, "Confirm Transaction")
            If result = MsgBoxResult.No Then
                Exit Sub
            End If
            queryType = "INSERT"
            UploadPhoto()
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub
    Private Sub releaseObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub
    Private Function checkImageForImage()
        pbStudImage.Image.Dispose()
        Dim ok As Boolean = False
        Try
            If (Me.txtPicName.Tag = String.Empty) Then
                ok = False
                Return ok
                Exit Function
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message.ToString(), "File Upload Error")
        End Try
        Dim fs As FileStream = New FileStream(mImageFilePath.ToString(), FileMode.Open)
        Dim img As Byte() = New Byte(fs.Length) {}
        fs.Read(img, 0, fs.Length)
        fs.Close()

        mImageFile = Image.FromFile(mImageFilePath.ToString())
        Dim imgType As String = Path.GetExtension(mImageFilePath)
        mImageFile = Nothing

        'Using dt As New DataTable
        '    dt.Load(reader)
        '    Dim row As DataRow = dt.Rows(0)
        '    Using ms As New IO.MemoryStream(CType(row("studImage"), Byte()))
        '        Dim img As Image = Image.FromStream(ms)
        '        pbStudImage.Image = img
        '        pbStudImage.SizeMode = PictureBoxSizeMode.StretchImage
        '    End Using
        'End Using

        cmdStudImages.Connection = conn
        cmdStudImages.CommandType = CommandType.Text
        cmdStudImages.CommandText = "SELECT * FROM  vwStudImages WHERE (imageStatus='True') " & _
            vbNewLine & "AND (studImage=@studImage) AND (studStatus='True')"
        cmdStudImages.Parameters.Clear()
        Dim pic As SqlParameter = New SqlParameter("@studImage", SqlDbType.Image)
        pic.Value = img
        cmdStudImages.Parameters.Add(pic)

        reader = cmdStudImages.ExecuteReader
        If reader.HasRows Then
            recordExists = True
        Else
            recordExists = False
        End If
        reader.Close()
    End Function
    Private Function UploadPhoto()
        pbStudImage.Image.Dispose()
        Dim ok As Boolean = False
        Try
            If (Me.txtPicName.Tag = String.Empty) Then
                ok = False
                Return ok
                Exit Function
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message.ToString(), "File Upload Error")
        End Try
        Try
            Dim fs As FileStream = New FileStream(mImageFilePath.ToString(), FileMode.Open)
            Dim img As Byte() = New Byte(fs.Length) {}
            fs.Read(img, 0, fs.Length)
            fs.Close()

            mImageFile = Image.FromFile(mImageFilePath.ToString())
            Dim imgType As String = Path.GetExtension(mImageFilePath)
            mImageFile = Nothing

            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If

            cmdStudImages.CommandText = "sprocStudentImages"
            cmdStudImages.Connection = conn
            cmdStudImages.CommandType = CommandType.StoredProcedure
            cmdStudImages.Parameters.Clear()
            cmdStudImages.Parameters.AddWithValue("@regBy", userName.Trim)
            cmdStudImages.Parameters.AddWithValue("@admNo", Me.cboStudentNo.Text.Trim)
            cmdStudImages.Parameters.AddWithValue("@queryType", Me.queryType.Trim)
            cmdStudImages.Parameters.AddWithValue("@dateOfreg", Date.Now)
            Dim pic As SqlParameter = New SqlParameter("@studImage", SqlDbType.Image)
            pic.Value = img
            cmdStudImages.Parameters.Add(pic)
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")

        Finally
        End Try

        Try
            rec = cmdStudImages.ExecuteNonQuery()
            If rec > 0 Then
                If queryType = "INSERT" Then
                    MsgBox("Record Saved", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Successfull Transaction")
                ElseIf queryType = "UPDATE" Then
                    MsgBox("Record Updated", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Successfull Transaction")
                End If
                Me.cboStudentNo.Tag = Nothing
                Me.cboClassName.SelectedIndex = -1
                Me.cboClassStream.SelectedIndex = -1
                Me.cboClassYear.SelectedIndex = -1
                Me.cboStudentNo.SelectedIndex = -1
                queryType = Nothing
                Me.pbStudImage.Image = Nothing
                Me.btnSave.Enabled = True
                Me.btnUpdate.Enabled = False
                Me.btnDelete.Enabled = False
            End If
            ok = True
        Catch ex As Exception
            MessageBox.Show(ex.Message.ToString(), "Data Error")
            ok = False
            Return ok
            Exit Function
        End Try

        Return ok

    End Function
    Private Sub btnView_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnView.Click
        Me.pbStudImage.Image = Nothing
        If Me.cboClassName.Text.Trim.Length <= 0 Then
            MsgBox("Missing Class Name", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboClassStream.Text.Trim.Length <= 0 Then
            MsgBox("Missing Class Stream", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboClassYear.Text.Trim.Length <= 0 Then
            MsgBox("Missing Class Year", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboStudentNo.Text.Trim.Length <= 0 Then
            MsgBox("Missing Student Number", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.txtStudName.Text.Trim.Length <= 0 Then
            MsgBox("Missing Student Name", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If
        RetrieveImgAC()
        If imageAvailable = True Then
            Me.btnSave.Enabled = False
            Me.btnUpdate.Enabled = True
            Me.btnDelete.Enabled = True
        ElseIf imageAvailable = False Then
            Me.btnSave.Enabled = True
            Me.btnUpdate.Enabled = False
            Me.btnDelete.Enabled = False
        End If
    End Sub
    Private Sub RetrieveImgAC()
        Try

            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            cmdStudImages.Connection = conn
            cmdStudImages.CommandType = CommandType.Text
            cmdStudImages.CommandText = "SELECT  * FROM  vwStudImages WHERE (admNo=@admNo) " & _
                vbNewLine & " AND (studStatus='True') AND (imageStatus='True')"
            cmdStudImages.Parameters.Clear()
            cmdStudImages.Parameters.AddWithValue("@admNo", Me.cboStudentNo.Text.Trim)
            reader = cmdStudImages.ExecuteReader

            Using dt As New DataTable
                dt.Load(reader)
                maxrec = dt.Rows.Count
                If maxrec > 0 Then
                    Dim row As DataRow = dt.Rows(0)
                    Using ms As New IO.MemoryStream(CType(row("studImage"), Byte()))
                        Dim img As Image = Image.FromStream(ms)
                        pbStudImage.Image = img
                        pbStudImage.SizeMode = PictureBoxSizeMode.StretchImage
                    End Using
                Else
                    MsgBox("No image in the database For the student", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
                    Exit Sub
                End If
            End Using
            If Not (reader.IsClosed = True) Then
                reader.Close()
            End If
            cmdStudImages.CommandText = "SELECT  * FROM  vwStudImages WHERE (admNo=@admNo) " & _
                vbNewLine & " AND (studStatus='True') AND (imageStatus='True')"
            cmdStudImages.Parameters.Clear()
            cmdStudImages.Parameters.AddWithValue("@admNo", Me.cboStudentNo.Text.Trim)
            reader = cmdStudImages.ExecuteReader
            If reader.HasRows Then
                While reader.Read
                    Me.cboStudentNo.Tag = IIf(DBNull.Value.Equals(reader!imageId), "", reader!imageId)
                End While


            Else
                MsgBox("No image in the database", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
                Exit Sub
            End If
            reader.Close()
            '    End Using
            'End Using
        Catch ex As Exception
            MessageBox.Show(ex.Message.ToString(), "Data Error")
        End Try
    End Sub

    Private Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        If Me.cboClassName.Text.Trim.Length <= 0 Then
            MsgBox("Missing Class Name", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboClassStream.Text.Trim.Length <= 0 Then
            MsgBox("Missing Class Stream", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboClassYear.Text.Trim.Length <= 0 Then
            MsgBox("Missing Class Year", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboStudentNo.Text.Trim.Length <= 0 Then
            MsgBox("Missing Student Number", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.txtStudName.Text.Trim.Length <= 0 Then
            MsgBox("Missing Student Name", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf IsNothing(Me.pbStudImage.Image) = True Then
            MsgBox("Missing Image To Delete", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If
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
            cmdStudImages.CommandText = "sprocStudentImages"
            cmdStudImages.Connection = conn
            cmdStudImages.CommandType = CommandType.StoredProcedure
            cmdStudImages.Parameters.Clear()
            cmdStudImages.Parameters.AddWithValue("@regBy", userName.Trim)
            cmdStudImages.Parameters.AddWithValue("@admNo", Me.cboStudentNo.Text.Trim)
            cmdStudImages.Parameters.AddWithValue("@queryType", Me.queryType.Trim)
            cmdStudImages.Parameters.AddWithValue("@dateOfreg", Date.Now)
            cmdStudImages.Parameters.AddWithValue("@imageId", Me.cboStudentNo.Tag)
            rec = cmdStudImages.ExecuteNonQuery
            If rec > 0 Then
                MsgBox("Record Deleted", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Successfull Transaction")
            End If
            Me.cboStudentNo.Tag = Nothing
            Me.cboClassName.SelectedIndex = -1
            Me.cboClassStream.SelectedIndex = -1
            Me.cboClassYear.SelectedIndex = -1
            Me.cboStudentNo.SelectedIndex = -1
            queryType = Nothing
            Me.pbStudImage.Image = Nothing
            Me.btnSave.Enabled = True
            Me.btnUpdate.Enabled = False
            Me.btnDelete.Enabled = False
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub
End Class