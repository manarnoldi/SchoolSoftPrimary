Imports System.Data.SqlClient
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports System.Drawing.Printing
Imports Excel = Microsoft.Office.Interop.Excel
Public Class frmRptstaffsubj
    Dim xlApp As Excel.Application
    Dim xlWorkBook As Excel.Workbook
    Dim xlWorkSheet As Excel.Worksheet
    Dim range As Excel.Range
    Dim fileName As String = Nothing
    Dim sheetName As String = Nothing
    Dim reader As SqlDataReader
    Dim rec As Integer = 0
    Dim cmdClassList As New SqlCommand
    Private Sub frmRptClassLists_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
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
        Me.cboClassName.Items.Clear()
        Me.cboYear.Items.Clear()

        Me.cboClassName.SelectedIndex = -1
        Me.cboYear.SelectedIndex = -1

        Me.cboClassName.Text = ""
        Me.cboYear.Text = ""

        Me.cmdClassList.Connection = conn
        Me.cmdClassList.CommandType = CommandType.Text
        Me.cmdClassList.CommandText = "SELECT DISTINCT className FROM tblClasses WHERE (status=1) ORDER BY className"
        Me.cmdClassList.Parameters.Clear()
        reader = Me.cmdClassList.ExecuteReader
        While reader.Read
            Me.cboClassName.Items.Add(IIf(DBNull.Value.Equals(reader!className), "", reader!className))
        End While
        reader.Close()

        Me.cmdClassList.CommandText = "SELECT DISTINCT year FROM tblClasses WHERE (status=1) ORDER BY year"
        Me.cmdClassList.Parameters.Clear()
        reader = Me.cmdClassList.ExecuteReader
        While reader.Read
            Me.cboYear.Items.Add(IIf(DBNull.Value.Equals(reader!year), "", reader!year))
        End While
        reader.Close()
    End Sub
    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub frmRptClassLists_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint
        If Me.IsMdiChild Then
            Dim PnlLoc As New Point
            PnlLoc.X = CInt((Me.Width - Me.pnlClassLists.Width) / 2)
            PnlLoc.Y = CInt((Me.Height - Me.pnlClassLists.Height) / 2.5)
            Me.pnlClassLists.Location = PnlLoc
        Else
            Me.pnlClassLists.Dock = DockStyle.Fill
        End If
    End Sub

    Private Sub btnLoad_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLoad.Click
        If Me.cboYear.Text.Trim.Length <= 0 Then
            MsgBox("Year is Missing", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly + MsgBoxStyle.ApplicationModal, "Missing Details")
            Exit Sub
        ElseIf Me.cboClassName.Text.Trim.Length <= 0 Then
            MsgBox("Class Name is Missing", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly + MsgBoxStyle.ApplicationModal, "Missing Details")
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()

            Me.Cursor = Cursors.WaitCursor
            Dim RptResultsView As New crtRptStaffSubjectAllocations
            SetReportLogOn(RptResultsView)
            RptResultsView.SummaryInfo.ReportComments = fullName.Trim
            RptResultsView.RecordSelectionFormula = "({vwAcadStaffSubjects.className}=" & Chr(34) & Me.cboClassName.Text.Trim & Chr(34) & ")"
            RptResultsView.RecordSelectionFormula += "AND ({vwAcadStaffSubjects.year}=" & Me.cboYear.Text.Trim & ")"
            Me.crtVwClassLists.ReportSource = RptResultsView
            Me.crtVwClassLists.Zoom(100)
            Me.crtVwClassLists.RefreshReport()
            Me.Cursor = Cursors.Arrow
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub

    Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        If Me.cboYear.Text.Trim.Length <= 0 Then
            MsgBox("Year is Missing", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly + MsgBoxStyle.ApplicationModal, "Missing Details")
            Exit Sub
        ElseIf Me.cboClassName.Text.Trim.Length <= 0 Then
            MsgBox("Class Name is Missing", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly + MsgBoxStyle.ApplicationModal, "Missing Details")
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()

            Me.Cursor = Cursors.WaitCursor
            Dim RptResultsView As New crtRptStaffSubjectAllocations
            SetReportLogOn(RptResultsView)
            RptResultsView.SummaryInfo.ReportComments = fullName.Trim
            RptResultsView.RecordSelectionFormula = "({vwStaffSubjects.className}=" & Chr(34) & Me.cboClassName.Text.Trim & Chr(34) & ")"
            RptResultsView.RecordSelectionFormula += "AND ({vwStaffSubjects.year}=" & Me.cboYear.Text.Trim & ")"
            RptResultsView.PrintToPrinter(1, True, 1, RptResultsView.FormatEngine.GetLastPageNumber(New CrystalDecisions.Shared.ReportPageRequestContext()))
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub

    Private Sub btnExport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExport.Click
        If Me.cboYear.Text.Trim.Length <= 0 Then
            MsgBox("Select year.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Missing Details")
            Exit Sub
        End If
        Try

        Dim result As MsgBoxResult = MsgBox("Export Records To Excel?", MsgBoxStyle.Question + MsgBoxStyle.ApplicationModal + MsgBoxStyle.YesNo, "Confirm Transaction")
        If result = MsgBoxResult.No Then
            Exit Sub
        End If
        Me.Cursor = Cursors.WaitCursor

        Dim xlApp As Excel.Application
        Dim xlWorkBook As Excel.Workbook
        Dim xlWorkSheet As Excel.Worksheet
        Dim misValue As Object = System.Reflection.Missing.Value
        Dim i As Integer

        xlApp = New Excel.Application
        xlWorkBook = xlApp.Workbooks.Add(misValue)
        xlWorkSheet = xlWorkBook.Sheets("Sheet1")
        
        xlWorkSheet.Cells(1, 1) = "Year"
        xlWorkSheet.Cells(1, 2) = "Class"
        xlWorkSheet.Cells(1, 3) = "Stream"
        xlWorkSheet.Cells(1, 4) = "StaffNo"
        xlWorkSheet.Cells(1, 5) = "Staff Name"
        xlWorkSheet.Cells(1, 6) = "Subject Name"

            xlWorkSheet.Cells.NumberFormat = "@"
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If
        dbconnection()

            Me.cmdClassList.Connection = conn
            Me.cmdClassList.CommandType = CommandType.Text
            Me.cmdClassList.CommandText = "SELECT * FROM vwStaffSubjects WHERE (subStatus=1) AND (staffSubjStatus=1) AND (staffStatus=1) " & _
                vbNewLine & " ORDER BY year,className,stream,subCode"
            Me.cmdClassList.Parameters.Clear()
            reader = Me.cmdClassList.ExecuteReader
            i = 2
            While reader.Read
                xlWorkSheet.Cells(i, 1) = IIf(DBNull.Value.Equals(reader!year), "", reader!year)
                xlWorkSheet.Cells(i, 2) = IIf(DBNull.Value.Equals(reader!className), "", reader!className)
                xlWorkSheet.Cells(i, 3) = IIf(DBNull.Value.Equals(reader!stream), "", reader!stream)
                xlWorkSheet.Cells(i, 4) = IIf(DBNull.Value.Equals(reader!empNo), "", reader!empNo)
                xlWorkSheet.Cells(i, 5) = IIf(DBNull.Value.Equals(reader!FullName), "", reader!FullName)
                xlWorkSheet.Cells(i, 6) = IIf(DBNull.Value.Equals(reader!subName), "", reader!subName)
                i = i + 1
            End While
        reader.Close()
        Dim dlg As New SaveFileDialog
        dlg.Filter = "Excel Files (*.xlsx)|*.xlsx"
        dlg.FilterIndex = 1
        dlg.InitialDirectory = My.Application.Info.DirectoryPath & "\EXCEL\\EICHER\BILLS\"
            dlg.FileName = "Staff Subject"
        Dim ExcelFile As String = ""
        If dlg.ShowDialog = Windows.Forms.DialogResult.OK Then
            ExcelFile = dlg.FileName
            xlWorkSheet.SaveAs(ExcelFile)
        End If
        xlWorkBook.Close()

        xlApp.Quit()

        releaseObject(xlApp)
        releaseObject(xlWorkBook)
        releaseObject(xlWorkSheet)
        Me.Cursor = Cursors.Arrow
            MsgBox("Excel file created successfully", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Successful Transaction")
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
End Class