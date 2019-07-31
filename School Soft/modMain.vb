Imports System.Data.SqlClient
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.ReportSource
Imports CrystalDecisions.Shared
Imports System.Data.DataTable
Module modMain
    Public conn As New SqlConnection
    Public Function dbconnection() As Boolean
        Try
            Dim connectionstring As String = "Server=" & My.Settings.serverName.Trim & ";User ID=" & My.Settings.UserName.Trim & ";Database=" & My.Settings.dbName.Trim & ";Password=" & My.Settings.PassWord.Trim & ";Connect Timeout=30000"
            conn = New SqlConnection(connectionstring)
            conn.Open()
        Catch
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            frmSettings.Show()
        End Try
    End Function
    Public conn1 As New SqlConnection
    Public Function dbconnection1() As Boolean
        Try
            Dim connectionstring As String = "Server=" & My.Settings.serverName.Trim & ";User ID=" & My.Settings.UserName.Trim & ";Database=" & My.Settings.dbName.Trim & ";Password=" & My.Settings.PassWord.Trim & ";Connect Timeout=30000"
            conn1 = New SqlConnection(connectionstring)
            conn1.Open()
        Catch
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            frmSettings.Show()
        End Try
    End Function
    Public maxrec As Integer
    Public activeband As String
    Public maxrec1 As Integer
    Public i As Integer = 0
    Public j As Integer = 0
    Public k As Integer = 0
    Public Datestring As String = ""
    Public userName As String
    Public domainName As String
    Public fullName As String
    Public empNo As String
    Public sql As String
    Public userId As Integer
    Public li As New ListViewItem
    Public dateToday As String = Date.Now.Day.ToString("00") & "-" & Date.Now.Month.ToString("00") & "-" & Date.Now.Year.ToString("0000")
    Public Function stringDate(ByVal sqlDate As Date)
        Datestring = ""
        Datestring = sqlDate.Day.ToString("00") & "-" & sqlDate.Month.ToString("00") & "-" & sqlDate.Year.ToString("0000")
    End Function

    Public Sub SetReportLogOn(ByRef docReport As ReportDocument)
        My.Settings.Reload()
        Dim ConnInfo As New ConnectionInfo
        ConnInfo.ServerName = My.Settings.serverName
        ConnInfo.UserID = My.Settings.userName
        ConnInfo.DatabaseName = My.Settings.dbName
        ConnInfo.Password = My.Settings.passWord
        Dim RptDB As Database = docReport.Database
        Dim RptTbl As Tables = RptDB.Tables
        Dim TblLogInfo As New TableLogOnInfo()
        For Each Tbl As Table In RptTbl
            TblLogInfo = Tbl.LogOnInfo
            TblLogInfo.ConnectionInfo = ConnInfo
            Tbl.ApplyLogOnInfo(TblLogInfo)
            Tbl.Location = (My.Settings.dbName + ".dbo." + Tbl.Name)
        Next
    End Sub
End Module
