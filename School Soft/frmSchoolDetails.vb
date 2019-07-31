Imports System.Windows.Forms
Imports System.IO
Imports System.Data.SqlClient
Public Class frmSchoolDetails
    Dim reader As SqlDataReader
    Dim exists As Boolean = True
    Dim rec As Integer = 0
    Dim cmdSchDetails As New SqlCommand
    Dim fdlg As OpenFileDialog = New OpenFileDialog()
    Private mImageFile As Image
    Private mImageFilePath As String

    Private Sub FindSavedRecord()
        dbconnection()
        cmdSchDetails.Connection = conn
        cmdSchDetails.CommandType = CommandType.Text
        cmdSchDetails.CommandText = "SELECT TOP 1 name,address,town,tel,emailAddress,schoolMotto,initials FROM tblSchoolDetails"
        cmdSchDetails.Parameters.Clear()
        reader = cmdSchDetails.ExecuteReader
        If reader.HasRows Then
            While reader.Read
                Me.txtSchName.Text = IIf(DBNull.Value.Equals(reader!name), "", reader!name)
                Me.txtSchAddress.Text = IIf(DBNull.Value.Equals(reader!address), "", reader!address)
                Me.txtTownName.Text = IIf(DBNull.Value.Equals(reader!town), "", reader!town)
                Me.txtTelNumber.Text = IIf(DBNull.Value.Equals(reader!tel), "", reader!tel)
                Me.txtEmailAddress.Text = IIf(DBNull.Value.Equals(reader!emailAddress), "", reader!emailAddress)
                Me.txtInitials.Text = IIf(DBNull.Value.Equals(reader!initials), "", reader!initials)
                Me.txtSchoolMotto.Text = IIf(DBNull.Value.Equals(reader!schoolMotto), "", reader!schoolMotto)
            End While
        End If
        reader.Close()
    End Sub

    Private Sub frmSchoolDetails_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        RetrieveImgAC()
        FindSavedRecord()
    End Sub

    Private Sub btnBrowse_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBrowse.Click
        fdlg.Title = "Set Photo"
        fdlg.InitialDirectory = "C:\Users\ARNOLD\Desktop\PROJECT ICONS\School Soft\"
        fdlg.Filter = "All Files|*.*|Bitmaps|*.bmp|GIFs|*.gif|JPEGs|*.jpg"
        fdlg.FilterIndex = 4
        fdlg.FileName = ""
        fdlg.RestoreDirectory = True
        If fdlg.ShowDialog() = Windows.Forms.DialogResult.OK Then
            Me.txtEmailAddress.Tag = fdlg.FileName
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
            Me.txtEmailAddress.Tag = sFilePath
            mImageFilePath = sFilePath
        End If
        pbImage.Image = Image.FromFile(Me.txtEmailAddress.Tag)
        pbImage.SizeMode = PictureBoxSizeMode.StretchImage
    End Sub
    Private Function UploadPhoto(ByVal updateinsert As String)
        pbImage.Image.Dispose()
        Dim ok As Boolean = False
        Try
            If (Me.txtEmailAddress.Tag = String.Empty) Then
                MsgBox("Select the logo file", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
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
            cmdSchDetails.CommandText = "sprocSchDetails"
            cmdSchDetails.Connection = conn
            cmdSchDetails.CommandType = CommandType.StoredProcedure
            cmdSchDetails.Parameters.Clear()
            cmdSchDetails.Parameters.AddWithValue("@userName", userName.Trim)
            cmdSchDetails.Parameters.AddWithValue("@name", Me.txtSchName.Text.Trim)
            cmdSchDetails.Parameters.AddWithValue("@address", Me.txtSchAddress.Text.Trim)
            cmdSchDetails.Parameters.AddWithValue("@town", Me.txtTownName.Text.Trim)
            cmdSchDetails.Parameters.AddWithValue("@tel", Me.txtTelNumber.Text.Trim)
            cmdSchDetails.Parameters.AddWithValue("@emailAddress", Me.txtEmailAddress.Text.Trim)
            cmdSchDetails.Parameters.AddWithValue("@initials", Me.txtInitials.Text.Trim)
            cmdSchDetails.Parameters.AddWithValue("@dateOfreg", Date.Now)
            cmdSchDetails.Parameters.AddWithValue("@updateInsert", updateinsert)
            cmdSchDetails.Parameters.AddWithValue("@schoolMotto", Me.txtSchoolMotto.Text.Trim)
            Dim pic As SqlParameter = New SqlParameter("@logo", SqlDbType.Image)
            pic.Value = img
            cmdSchDetails.Parameters.Add(pic)
        Catch ex As Exception
        Finally
        End Try

        Try
            rec = cmdSchDetails.ExecuteNonQuery()
            If rec > 0 Then
                MsgBox("Record Saved", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Successfull Transaction")
                Me.txtTownName.Text = ""
                Me.txtTelNumber.Text = ""
                Me.txtSchName.Text = ""
                Me.txtSchAddress.Text = ""
                Me.txtEmailAddress.Text = ""
                Me.txtInitials.Text = ""
                Me.txtSchoolMotto.Text = ""
                Me.txtEmailAddress.Tag = Nothing
                Me.pbImage.Image = Nothing
                RetrieveImgAC()
                FindSavedRecord()
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

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        If Me.txtSchAddress.Text.Trim.Length <= 0 Then
            MsgBox("Missing School Address", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.txtSchName.Text.Trim.Length <= 0 Then
            MsgBox("Missing School Name", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.txtTelNumber.Text.Trim.Length <= 0 Then
            MsgBox("Missing School Telephone Number", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.txtTownName.Text.Trim.Length <= 0 Then
            MsgBox("Missing Town Name", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.txtSchoolMotto.Text.Trim.Length <= 0 Then
            MsgBox("Missing School Motto", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If
        dbconnection()
        checkExistence()
        Dim updateinsert As String = If(checkExistence(), "UPDATE", "INSERT")
        Dim result As MsgBoxResult = MsgBox("Save Record?", MsgBoxStyle.Question + MsgBoxStyle.ApplicationModal + MsgBoxStyle.YesNo, "Confirm Transaction")
        If result = MsgBoxResult.No Then
            Exit Sub
        End If
        UploadPhoto(updateinsert)
    End Sub
    Private Function checkExistence() As Boolean
        cmdSchDetails.Connection = conn
        cmdSchDetails.CommandType = CommandType.Text
        cmdSchDetails.CommandText = "SELECT COUNT (*) AS count FROM tblSchoolDetails"
        cmdSchDetails.Parameters.Clear()
        reader = cmdSchDetails.ExecuteReader
        If reader.HasRows Then
            While reader.Read
                i = IIf(DBNull.Value.Equals(reader!count), 0, reader!count)
            End While
        End If
        If i > 0 Then
            exists = True
        Else
            exists = False
        End If
        reader.Close()
        Return exists
    End Function

    Private Sub txtEmailAddress_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtEmailAddress.LostFocus
        If Me.txtEmailAddress.Text = "" Then
            Exit Sub
        End If
        If Me.txtEmailAddress.Text.Contains(".") = False Or Me.txtEmailAddress.Text.Contains("@") = False Then
            MsgBox("Enter Correct Email Address", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Missing Details")
            Me.txtEmailAddress.Text = ""
        End If
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub frmSchoolDetails_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint
        If Me.IsMdiChild Then
            Dim PnlLoc As New Point
            PnlLoc.X = CInt((Me.Width - Me.pnlSchDetails.Width) / 2)
            PnlLoc.Y = CInt((Me.Height - Me.pnlSchDetails.Height) / 2.5)
            Me.pnlSchDetails.Location = PnlLoc
        Else
            Me.pnlSchDetails.Dock = DockStyle.Fill
        End If
    End Sub
    Private Sub RetrieveImgAC()
        Try

            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            cmdSchDetails.Connection = conn
            cmdSchDetails.CommandType = CommandType.Text
            cmdSchDetails.CommandText = "SELECT  Top 1 logo,name FROM tblSchoolDetails"
            cmdSchDetails.Parameters.Clear()
            reader = cmdSchDetails.ExecuteReader

            'Using cmd As New SqlClient.SqlCommand("SELECT  logo,name FROM  tblSchoolDetails")
            '    cmd.Parameters.AddWithValue("ID", SelectedItem)
            '    Using dr As SqlClient.SqlDataReader = cmd.ExecuteReader()
            Using dt As New DataTable
                dt.Load(reader)
                Dim row As DataRow = dt.Rows(0)
                Using ms As New IO.MemoryStream(CType(row("logo"), Byte()))
                    Dim img As Image = Image.FromStream(ms)
                    pbImage.Image = img
                    pbImage.SizeMode = PictureBoxSizeMode.StretchImage
                End Using
            End Using
            '    End Using
            'End Using
        Catch ex As Exception

            MessageBox.Show(ex.Message.ToString(), "Data Error")

        End Try
    End Sub
End Class