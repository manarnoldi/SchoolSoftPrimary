Imports System.Data.SqlClient
Imports CrystalDecisions.CrystalReports.Engine

Public Class frmAcadResultsSummary
    Dim reader As SqlDataReader
    Dim cmdResultsSummary As New SqlCommand
    Dim rec As Integer = 0
    Dim returnedAverages As List(Of String) = New List(Of String)
    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub frmAcadResultsSummary_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
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
        Me.cboClass.Items.Clear()
        Me.cboClass.Text = ""
        Me.cboClass.SelectedIndex = -1

        Me.cboStream.Items.Clear()
        Me.cboStream.Text = ""
        Me.cboStream.SelectedIndex = -1

        Me.cboTerm.Items.Clear()
        Me.cboTerm.Text = ""
        Me.cboTerm.SelectedIndex = -1

        Me.cboYear.Items.Clear()
        Me.cboYear.Text = ""
        Me.cboYear.SelectedIndex = -1

        Me.cboExamType.Items.Clear()
        Me.cboExamType.Text = ""
        Me.cboExamType.SelectedIndex = -1

        Me.cboExamName.Items.Clear()
        Me.cboExamName.Text = ""
        Me.cboExamName.SelectedIndex = -1

        Me.cmdResultsSummary.Connection = conn
        Me.cmdResultsSummary.CommandType = CommandType.Text
        Me.cmdResultsSummary.CommandText = "SELECT DISTINCT className FROM tblClasses WHERE (status=1) ORDER BY className"
        Me.cmdResultsSummary.Parameters.Clear()
        reader = Me.cmdResultsSummary.ExecuteReader
        If reader.HasRows Then
            While reader.Read
                Me.cboClass.Items.Add(IIf(DBNull.Value.Equals(reader!className), "", reader!className))
            End While
        End If
        reader.Close()

        Me.cmdResultsSummary.CommandText = "SELECT DISTINCT stream FROM tblClasses WHERE (status=1) ORDER BY stream"
        Me.cmdResultsSummary.Parameters.Clear()
        reader = Me.cmdResultsSummary.ExecuteReader
        If reader.HasRows Then
            While reader.Read
                Me.cboStream.Items.Add(IIf(DBNull.Value.Equals(reader!stream), "", reader!stream))
            End While
        End If
        reader.Close()

        Me.cmdResultsSummary.CommandText = "SELECT DISTINCT termName FROM tblSchoolCalendar WHERE (status=1) ORDER BY termName"
        Me.cmdResultsSummary.Parameters.Clear()
        reader = Me.cmdResultsSummary.ExecuteReader
        If reader.HasRows Then
            While reader.Read
                Me.cboTerm.Items.Add(IIf(DBNull.Value.Equals(reader!termName), "", reader!termName))
            End While
        End If
        reader.Close()

        Me.cmdResultsSummary.CommandText = "SELECT DISTINCT year FROM tblSchoolCalendar WHERE (status=1) ORDER BY year"
        Me.cmdResultsSummary.Parameters.Clear()
        reader = Me.cmdResultsSummary.ExecuteReader
        If reader.HasRows Then
            While reader.Read
                Me.cboYear.Items.Add(IIf(DBNull.Value.Equals(reader!year), "", reader!year))
            End While
        End If
        reader.Close()

        Me.cmdResultsSummary.CommandText = "SELECT DISTINCT examType FROM tblExamNames WHERE (status=1) ORDER BY examType"
        Me.cmdResultsSummary.Parameters.Clear()
        reader = Me.cmdResultsSummary.ExecuteReader
        If reader.HasRows Then
            While reader.Read
                Me.cboExamType.Items.Add(IIf(DBNull.Value.Equals(reader!examType), "", reader!examType))
            End While
        End If
        reader.Close()
    End Sub
    Private Sub frmAcadResultsSummary_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint
        If Me.IsMdiChild Then
            Dim PnlLoc As New Point
            PnlLoc.X = CInt((Me.Width - Me.pnlResultSummary.Width) / 2)
            PnlLoc.Y = CInt((Me.Height - Me.pnlResultSummary.Height) / 2.5)
            Me.pnlResultSummary.Location = PnlLoc
        Else
            Me.pnlResultSummary.Dock = DockStyle.Fill
        End If
    End Sub

    Private Function CalculateMeans(ByVal tableName As String, ByVal className As String, ByVal reportType As Integer) As List(Of String)
        Dim returnVal As List(Of String) = New List(Of String)
        Dim meanMark As String = ""
        Dim meanGrade As String = ""

        Me.cmdResultsSummary.Connection = conn
        Me.cmdResultsSummary.CommandType = CommandType.Text

        If reportType = 1 Then
            Me.cmdResultsSummary.CommandText = "SELECT meanMark,gradeName,meanScore FROM (SELECT  ROUND(SUM(TTMARK)/SUM(SUBJREQ),2) AS meanMark," +
            "ROUND(SUM(TTPOINTS)/SUM(SUBJREQ),4) AS meanScore,classId FROM " +
            tableName + " GROUP BY classId) t INNER JOIN tblClasses ON t.classId = tblClasses.classId INNER JOIN tblGrades ON " +
            "tblClasses.className = tblGrades.className WHERE t.meanMark BETWEEN tblGrades.lowerMark And tblGrades.upperMark And " +
            "tblGrades.className='" + className + "' AND tblGrades.status=1"
        ElseIf reportType = 2 Then
            Me.cmdResultsSummary.CommandText = "SELECT meanMark,gradeName,meanScore FROM (SELECT  ROUND(SUM(totalMarks)/SUM(totalStud),2) AS meanMark," +
            "ROUND(SUM(totalPoints)/SUM(totalStud),4) AS meanScore,classId FROM " +
            tableName + " GROUP BY classId) t INNER JOIN tblClasses ON t.classId = tblClasses.classId INNER JOIN tblGrades ON " +
            "tblClasses.className = tblGrades.className WHERE t.meanMark BETWEEN tblGrades.lowerMark And tblGrades.upperMark And " +
            "tblGrades.className='" + className + "' AND tblGrades.status=1"
        ElseIf reportType = 3 Then
            Me.cmdResultsSummary.CommandText = "SELECT meanMark,gradeName,meanScore FROM (SELECT  ROUND(SUM(meanMark)/COUNT(meanMark),2) AS meanMark," +
            "ROUND(SUM(meanScore)/SUM(meanScore),4) AS meanScore,classId FROM " +
            tableName + " GROUP BY classId) t INNER JOIN tblClasses ON t.classId = tblClasses.classId INNER JOIN tblGrades ON " +
            "tblClasses.className = tblGrades.className WHERE t.meanMark BETWEEN tblGrades.lowerMark And tblGrades.upperMark And " +
            "tblGrades.className='" + className + "' AND tblGrades.status=1"
        End If

        Me.cmdResultsSummary.Parameters.Clear()
        reader = Me.cmdResultsSummary.ExecuteReader
        While reader.Read
            returnVal.Add(IIf(DBNull.Value.Equals(reader!meanMark), "", reader!meanMark))
            returnVal.Add(IIf(DBNull.Value.Equals(reader!meanScore), "", reader!meanScore))
            returnVal.Add(IIf(DBNull.Value.Equals(reader!gradeName), "", reader!gradeName))
        End While
        reader.Close()
        Return returnVal
    End Function

    Private Sub btnPreview_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPreview.Click
        If Me.cboYear.Text.Trim.Length <= 0 Then
            MsgBox("Missing Year.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.cboTerm.Text.Trim.Length <= 0 Then
            MsgBox("Missing Term.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        ElseIf Me.rbnbroadclass.Checked = False And Me.rbnbroadgradeclass.Checked = False And Me.rbnbroadgradestream.Checked = False _
            And Me.rbnbroadstream.Checked = False And Me.rbnPointClass.Checked = False And Me.rbnPointStream.Checked = False And
            Me.rbnSchSummaryGradeWise.Checked = False And Me.rbnSchSummaryMarkWise.Checked = False And
            Me.rbnsubjectclass.Checked = False And Me.rbnsubjectstream.Checked = False And Me.rbnbroadclassExam.Checked = False And
            Me.rbnbroadstreamExam.Checked = False Then
            MsgBox("Select Analysis Report.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
            Exit Sub
        End If

        Try
            Dim meanMark As Double = 0
            Dim meanGrade As String = ""

            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()
            Me.cmdResultsSummary.Connection = conn
            Me.cmdResultsSummary.CommandType = CommandType.Text
            Me.cmdResultsSummary.CommandText = "SELECT COUNT(*) AS count FROM tblSchoolCalendar WHERE (termName=@termName) AND (year=@year) AND (status=1)"
            Me.cmdResultsSummary.Parameters.Clear()
            Me.cmdResultsSummary.Parameters.AddWithValue("@termName", Me.cboTerm.Text.Trim)
            Me.cmdResultsSummary.Parameters.AddWithValue("@year", Me.cboYear.Text.Trim)
            reader = Me.cmdResultsSummary.ExecuteReader
            i = 0
            While reader.Read
                i = (IIf(DBNull.Value.Equals(reader!count), "", reader!count))
            End While
            reader.Close()

            If i <= 0 Then
                MsgBox("Term Details not registered.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
                Exit Sub
            End If

            Me.Cursor = Cursors.WaitCursor
            If Me.rbnbroadstream.Checked = True Then
                If Me.cboStream.Text.Trim.Length <= 0 Then
                    MsgBox("Missing Stream Name.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
                    Exit Sub
                ElseIf Me.cboClass.Text.Trim.Length <= 0 Then
                    MsgBox("Missing Class.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
                    Exit Sub
                End If

                Me.cmdResultsSummary.Connection = conn
                Me.cmdResultsSummary.CommandType = CommandType.Text
                Me.cmdResultsSummary.CommandText = "SELECT COUNT(*) AS count FROM tblTempERSumMarkStream1"
                Me.cmdResultsSummary.Parameters.Clear()
                reader = Me.cmdResultsSummary.ExecuteReader
                i = 0
                While reader.Read
                    i = (IIf(DBNull.Value.Equals(reader!count), "", reader!count))
                End While
                reader.Close()


                If i > 0 Then
                    MsgBox("Some printing on the reporting criteria going on.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
                    Exit Sub
                End If

                Me.cmdResultsSummary.Connection = conn
                Me.cmdResultsSummary.CommandType = CommandType.StoredProcedure
                Me.cmdResultsSummary.CommandText = "sprocTempERSumMarkStream"
                Me.cmdResultsSummary.Parameters.Clear()
                Me.cmdResultsSummary.Parameters.AddWithValue("@term", Me.cboTerm.Text.Trim)
                Me.cmdResultsSummary.Parameters.AddWithValue("@year", Me.cboYear.Text.Trim)
                Me.cmdResultsSummary.Parameters.AddWithValue("@class", Me.cboClass.Text.Trim)
                Me.cmdResultsSummary.Parameters.AddWithValue("@stream", Me.cboStream.Text.Trim)
                Me.cmdResultsSummary.ExecuteNonQuery()

                returnedAverages.Clear()
                returnedAverages = CalculateMeans("tblTempERSumMarkStream1", Me.cboClass.Text.Trim, 1)

                Dim RptResultsView As New crtAcadResultsMarkClassStream
                SetReportLogOn(RptResultsView)
                RptResultsView.SummaryInfo.ReportComments = fullName.Trim
                RptResultsView.SummaryInfo.ReportTitle = "BROAD SHEET MARK WISE CLASS WISE STREAM WISE"
                RptResultsView.RecordSelectionFormula = "({tblTempERSumMarkStream.className}=" & Chr(34) & Me.cboClass.Text.Trim & Chr(34) & ")"
                RptResultsView.RecordSelectionFormula += "AND ({tblTempERSumMarkStream.year}=" & Me.cboYear.Text.Trim & ")"
                RptResultsView.RecordSelectionFormula += "AND ({tblTempERSumMarkStream.stream}=" & Chr(34) & Me.cboStream.Text.Trim & Chr(34) & ")"
                RptResultsView.RecordSelectionFormula += "AND ({tblTempERSumMarkStream.termName}=" & Chr(34) & Me.cboTerm.Text.Trim & Chr(34) & ")"
                DirectCast(RptResultsView.ReportDefinition.ReportObjects("meanMark"), TextObject).Text = If(returnedAverages.Count > 0, returnedAverages.Item(0), "")
                DirectCast(RptResultsView.ReportDefinition.ReportObjects("meanGrade"), TextObject).Text = If(returnedAverages.Count > 0, returnedAverages.Item(2), "")


                frmResultsViewing.crtViewResultsSummary.ReportSource = RptResultsView
                frmResultsViewing.crtViewResultsSummary.Zoom(100)
                frmResultsViewing.crtViewResultsSummary.RefreshReport()
                frmResultsViewing.MdiParent = frmHome
                frmResultsViewing.Show()

                Me.cmdResultsSummary.Connection = conn
                Me.cmdResultsSummary.CommandType = CommandType.Text
                Me.cmdResultsSummary.CommandText = "DELETE FROM tblTempERSumMarkStream1"
                Me.cmdResultsSummary.Parameters.Clear()
                Me.cmdResultsSummary.ExecuteNonQuery()

            ElseIf Me.rbnbroadclass.Checked = True Then
                If Me.cboClass.Text.Trim.Length <= 0 Then
                    MsgBox("Missing Class.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
                    Exit Sub
                End If
                Me.cmdResultsSummary.Connection = conn
                Me.cmdResultsSummary.CommandType = CommandType.Text
                Me.cmdResultsSummary.CommandText = "SELECT COUNT(*) AS count FROM tblTempERSumMarkClass1"
                Me.cmdResultsSummary.Parameters.Clear()
                reader = Me.cmdResultsSummary.ExecuteReader
                i = 0
                While reader.Read
                    i = (IIf(DBNull.Value.Equals(reader!count), "", reader!count))
                End While
                reader.Close()

                If i > 0 Then
                    MsgBox("Some printing on the reporting criteria going on.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
                    Exit Sub
                End If

                Me.cmdResultsSummary.Connection = conn
                Me.cmdResultsSummary.CommandType = CommandType.StoredProcedure
                Me.cmdResultsSummary.CommandText = "sprocTempERSumMarkClass"
                Me.cmdResultsSummary.Parameters.Clear()
                Me.cmdResultsSummary.Parameters.AddWithValue("@term", Me.cboTerm.Text.Trim)
                Me.cmdResultsSummary.Parameters.AddWithValue("@year", Me.cboYear.Text.Trim)
                Me.cmdResultsSummary.Parameters.AddWithValue("@class", Me.cboClass.Text.Trim)
                Me.cmdResultsSummary.ExecuteNonQuery()

                returnedAverages.Clear()
                returnedAverages = CalculateMeans("tblTempERSumMarkClass1", Me.cboClass.Text.Trim, 1)

                Dim RptResultsView As New crtAcadResultsMarkClass
                SetReportLogOn(RptResultsView)
                RptResultsView.SummaryInfo.ReportComments = fullName.Trim
                RptResultsView.SummaryInfo.ReportTitle = "BROAD SHEET MARK WISE CLASS WISE"
                RptResultsView.RecordSelectionFormula = "({tblTempERSumMarkClass.className}=" & Chr(34) & Me.cboClass.Text.Trim & Chr(34) & ")"
                RptResultsView.RecordSelectionFormula += "AND ({tblTempERSumMarkClass.year}=" & Me.cboYear.Text.Trim & ")"
                RptResultsView.RecordSelectionFormula += "AND ({tblTempERSumMarkClass.termName}=" & Chr(34) & Me.cboTerm.Text.Trim & Chr(34) & ")"
                DirectCast(RptResultsView.ReportDefinition.ReportObjects("meanMark"), TextObject).Text = If(returnedAverages.Count > 0, returnedAverages.Item(0), "")
                DirectCast(RptResultsView.ReportDefinition.ReportObjects("meanGrade"), TextObject).Text = If(returnedAverages.Count > 0, returnedAverages.Item(2), "")

                frmResultsViewing.crtViewResultsSummary.ReportSource = RptResultsView
                frmResultsViewing.crtViewResultsSummary.Zoom(100)
                frmResultsViewing.crtViewResultsSummary.RefreshReport()
                frmResultsViewing.MdiParent = frmHome
                frmResultsViewing.Show()

                Me.cmdResultsSummary.Connection = conn
                Me.cmdResultsSummary.CommandType = CommandType.Text
                Me.cmdResultsSummary.CommandText = "DELETE FROM tblTempERSumMarkClass1"
                Me.cmdResultsSummary.Parameters.Clear()
                Me.cmdResultsSummary.ExecuteNonQuery()
            ElseIf Me.rbnbroadstreamExam.Checked = True Then
                If Me.cboStream.Text.Trim.Length <= 0 Then
                    MsgBox("Missing Stream Name.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
                    Exit Sub
                ElseIf Me.cboClass.Text.Trim.Length <= 0 Then
                    MsgBox("Missing Class.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
                    Exit Sub
                ElseIf Me.cboExamType.Text.Trim.Length <= 0 Then
                    MsgBox("Missing Exam Type.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
                    Exit Sub
                ElseIf Me.cboExamName.Text.Trim.Length <= 0 Then
                    MsgBox("Missing Exam Name.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
                    Exit Sub
                End If

                Me.cmdResultsSummary.Connection = conn
                Me.cmdResultsSummary.CommandType = CommandType.Text
                Me.cmdResultsSummary.CommandText = "SELECT COUNT(*) AS count FROM tblTempERSumMarkStream1"
                Me.cmdResultsSummary.Parameters.Clear()
                reader = Me.cmdResultsSummary.ExecuteReader
                i = 0
                While reader.Read
                    i = (IIf(DBNull.Value.Equals(reader!count), "", reader!count))
                End While
                reader.Close()


                If i > 0 Then
                    MsgBox("Some printing on the reporting criteria going on.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
                    Exit Sub
                End If

                Me.cmdResultsSummary.Connection = conn
                Me.cmdResultsSummary.CommandType = CommandType.StoredProcedure
                Me.cmdResultsSummary.CommandText = "sprocTempERSumExamMarkStream"
                Me.cmdResultsSummary.Parameters.Clear()
                Me.cmdResultsSummary.Parameters.AddWithValue("@term", Me.cboTerm.Text.Trim)
                Me.cmdResultsSummary.Parameters.AddWithValue("@year", Me.cboYear.Text.Trim)
                Me.cmdResultsSummary.Parameters.AddWithValue("@class", Me.cboClass.Text.Trim)
                Me.cmdResultsSummary.Parameters.AddWithValue("@stream", Me.cboStream.Text.Trim)
                Me.cmdResultsSummary.Parameters.AddWithValue("@examType", Me.cboExamType.Text.Trim)
                Me.cmdResultsSummary.Parameters.AddWithValue("@examName", Me.cboExamName.Text.Trim)
                Me.cmdResultsSummary.ExecuteNonQuery()

                returnedAverages.Clear()
                returnedAverages = CalculateMeans("tblTempERSumMarkStream1", Me.cboClass.Text.Trim, 1)

                Dim RptResultsView As New crtAcadResultsMarkClassStream
                SetReportLogOn(RptResultsView)
                RptResultsView.SummaryInfo.ReportComments = fullName.Trim
                RptResultsView.SummaryInfo.ReportTitle = "BROAD SHEET MARK WISE CLASS WISE STREAM WISE EXAM WISE (" +
                    Me.cboExamType.Text.Trim + ": " + Me.cboExamName.Text.Trim + ")"
                RptResultsView.RecordSelectionFormula = "({tblTempERSumMarkStream.className}=" & Chr(34) & Me.cboClass.Text.Trim & Chr(34) & ")"
                RptResultsView.RecordSelectionFormula += "AND ({tblTempERSumMarkStream.year}=" & Me.cboYear.Text.Trim & ")"
                RptResultsView.RecordSelectionFormula += "AND ({tblTempERSumMarkStream.stream}=" & Chr(34) & Me.cboStream.Text.Trim & Chr(34) & ")"
                RptResultsView.RecordSelectionFormula += "AND ({tblTempERSumMarkStream.termName}=" & Chr(34) & Me.cboTerm.Text.Trim & Chr(34) & ")"
                DirectCast(RptResultsView.ReportDefinition.ReportObjects("meanMark"), TextObject).Text = If(returnedAverages.Count > 0, returnedAverages.Item(0), "")
                DirectCast(RptResultsView.ReportDefinition.ReportObjects("meanGrade"), TextObject).Text = If(returnedAverages.Count > 0, returnedAverages.Item(2), "")

                frmResultsViewing.crtViewResultsSummary.ReportSource = RptResultsView
                frmResultsViewing.crtViewResultsSummary.Zoom(100)
                frmResultsViewing.crtViewResultsSummary.RefreshReport()
                frmResultsViewing.MdiParent = frmHome
                frmResultsViewing.Show()

                Me.cmdResultsSummary.Connection = conn
                Me.cmdResultsSummary.CommandType = CommandType.Text
                Me.cmdResultsSummary.CommandText = "DELETE FROM tblTempERSumMarkStream1"
                Me.cmdResultsSummary.Parameters.Clear()
                Me.cmdResultsSummary.ExecuteNonQuery()

            ElseIf Me.rbnbroadclassExam.Checked = True Then
                If Me.cboClass.Text.Trim.Length <= 0 Then
                    MsgBox("Missing Class.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
                    Exit Sub
                ElseIf Me.cboExamType.Text.Trim.Length <= 0 Then
                    MsgBox("Missing Exam Type.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
                    Exit Sub
                ElseIf Me.cboExamName.Text.Trim.Length <= 0 Then
                    MsgBox("Missing Exam Name.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
                    Exit Sub
                End If
                Me.cmdResultsSummary.Connection = conn
                Me.cmdResultsSummary.CommandType = CommandType.Text
                Me.cmdResultsSummary.CommandText = "SELECT COUNT(*) AS count FROM tblTempERSumMarkClass1"
                Me.cmdResultsSummary.Parameters.Clear()
                reader = Me.cmdResultsSummary.ExecuteReader
                i = 0
                While reader.Read
                    i = (IIf(DBNull.Value.Equals(reader!count), "", reader!count))
                End While
                reader.Close()

                If i > 0 Then
                    MsgBox("Some printing on the reporting criteria going on.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
                    Exit Sub
                End If

                Me.cmdResultsSummary.Connection = conn
                Me.cmdResultsSummary.CommandType = CommandType.StoredProcedure
                Me.cmdResultsSummary.CommandText = "sprocTempERSumExamMarkClass"
                Me.cmdResultsSummary.Parameters.Clear()
                Me.cmdResultsSummary.Parameters.AddWithValue("@term", Me.cboTerm.Text.Trim)
                Me.cmdResultsSummary.Parameters.AddWithValue("@year", Me.cboYear.Text.Trim)
                Me.cmdResultsSummary.Parameters.AddWithValue("@class", Me.cboClass.Text.Trim)
                Me.cmdResultsSummary.Parameters.AddWithValue("@examType", Me.cboExamType.Text.Trim)
                Me.cmdResultsSummary.Parameters.AddWithValue("@examName", Me.cboExamName.Text.Trim)
                Me.cmdResultsSummary.ExecuteNonQuery()

                returnedAverages.Clear()
                returnedAverages = CalculateMeans("tblTempERSumMarkClass1", Me.cboClass.Text.Trim, 1)

                Dim RptResultsView As New crtAcadResultsMarkClass
                SetReportLogOn(RptResultsView)
                RptResultsView.SummaryInfo.ReportComments = fullName.Trim
                RptResultsView.SummaryInfo.ReportTitle = "BROAD SHEET MARK WISE CLASS WISE EXAM WISE (" +
                    Me.cboExamType.Text.Trim + ": " + Me.cboExamName.Text.Trim + ")"
                RptResultsView.RecordSelectionFormula = "({tblTempERSumMarkClass.className}=" & Chr(34) & Me.cboClass.Text.Trim & Chr(34) & ")"
                RptResultsView.RecordSelectionFormula += "AND ({tblTempERSumMarkClass.year}=" & Me.cboYear.Text.Trim & ")"
                RptResultsView.RecordSelectionFormula += "AND ({tblTempERSumMarkClass.termName}=" & Chr(34) & Me.cboTerm.Text.Trim & Chr(34) & ")"
                DirectCast(RptResultsView.ReportDefinition.ReportObjects("meanMark"), TextObject).Text = If(returnedAverages.Count > 0, returnedAverages.Item(0), "")
                DirectCast(RptResultsView.ReportDefinition.ReportObjects("meanGrade"), TextObject).Text = If(returnedAverages.Count > 0, returnedAverages.Item(2), "")

                frmResultsViewing.crtViewResultsSummary.ReportSource = RptResultsView
                frmResultsViewing.crtViewResultsSummary.Zoom(100)
                frmResultsViewing.crtViewResultsSummary.RefreshReport()
                frmResultsViewing.MdiParent = frmHome
                frmResultsViewing.Show()

                Me.cmdResultsSummary.Connection = conn
                Me.cmdResultsSummary.CommandType = CommandType.Text
                Me.cmdResultsSummary.CommandText = "DELETE FROM tblTempERSumMarkClass1"
                Me.cmdResultsSummary.Parameters.Clear()
                Me.cmdResultsSummary.ExecuteNonQuery()
            ElseIf Me.rbnbroadgradestream.Checked = True Then
                If Me.cboStream.Text.Trim.Length <= 0 Then
                    MsgBox("Missing Stream Name.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
                    Exit Sub
                ElseIf Me.cboClass.Text.Trim.Length <= 0 Then
                    MsgBox("Missing Class.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
                    Exit Sub
                End If
                Me.cmdResultsSummary.Connection = conn
                Me.cmdResultsSummary.CommandType = CommandType.Text
                Me.cmdResultsSummary.CommandText = "SELECT COUNT(*) AS count FROM tblTempERSumGradeStream1"
                Me.cmdResultsSummary.Parameters.Clear()
                reader = Me.cmdResultsSummary.ExecuteReader
                i = 0
                While reader.Read
                    i = (IIf(DBNull.Value.Equals(reader!count), "", reader!count))
                End While
                reader.Close()

                If i > 0 Then
                    MsgBox("Some printing on the reporting criteria going on.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
                    Exit Sub
                End If

                Me.cmdResultsSummary.Connection = conn
                Me.cmdResultsSummary.CommandType = CommandType.StoredProcedure
                Me.cmdResultsSummary.CommandText = "sprocTempERSumGradeStream"
                Me.cmdResultsSummary.Parameters.Clear()
                Me.cmdResultsSummary.Parameters.AddWithValue("@term", Me.cboTerm.Text.Trim)
                Me.cmdResultsSummary.Parameters.AddWithValue("@year", Me.cboYear.Text.Trim)
                Me.cmdResultsSummary.Parameters.AddWithValue("@class", Me.cboClass.Text.Trim)
                Me.cmdResultsSummary.Parameters.AddWithValue("@stream", Me.cboStream.Text.Trim)
                Me.cmdResultsSummary.ExecuteNonQuery()

                returnedAverages.Clear()
                returnedAverages = CalculateMeans("tblTempERSumGradeStream1", Me.cboClass.Text.Trim, 1)

                Dim RptResultsView As New crtAcadResultsGradeClassStream
                SetReportLogOn(RptResultsView)
                RptResultsView.SummaryInfo.ReportComments = fullName.Trim
                RptResultsView.RecordSelectionFormula = "({tblTempERSumGradeStream.className}=" & Chr(34) & Me.cboClass.Text.Trim & Chr(34) & ")"
                RptResultsView.RecordSelectionFormula += "AND ({tblTempERSumGradeStream.year}=" & Me.cboYear.Text.Trim & ")"
                RptResultsView.RecordSelectionFormula += "AND ({tblTempERSumGradeStream.stream}=" & Chr(34) & Me.cboStream.Text.Trim & Chr(34) & ")"
                RptResultsView.RecordSelectionFormula += "AND ({tblTempERSumGradeStream.termName}=" & Chr(34) & Me.cboTerm.Text.Trim & Chr(34) & ")"
                DirectCast(RptResultsView.ReportDefinition.ReportObjects("meanMark"), TextObject).Text = If(returnedAverages.Count > 0, returnedAverages.Item(0), "")
                DirectCast(RptResultsView.ReportDefinition.ReportObjects("meanGrade"), TextObject).Text = If(returnedAverages.Count > 0, returnedAverages.Item(2), "")

                frmResultsViewing.crtViewResultsSummary.ReportSource = RptResultsView
                frmResultsViewing.crtViewResultsSummary.Zoom(100)
                frmResultsViewing.crtViewResultsSummary.RefreshReport()
                frmResultsViewing.MdiParent = frmHome
                frmResultsViewing.Show()

                Me.cmdResultsSummary.Connection = conn
                Me.cmdResultsSummary.CommandType = CommandType.Text
                Me.cmdResultsSummary.CommandText = "DELETE FROM tblTempERSumGradeStream1"
                Me.cmdResultsSummary.Parameters.Clear()
                Me.cmdResultsSummary.ExecuteNonQuery()
            ElseIf Me.rbnbroadgradeclass.Checked = True Then
                If Me.cboClass.Text.Trim.Length <= 0 Then
                    MsgBox("Missing Class.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
                    Exit Sub
                End If
                Me.cmdResultsSummary.Connection = conn
                Me.cmdResultsSummary.CommandType = CommandType.Text
                Me.cmdResultsSummary.CommandText = "SELECT COUNT(*) AS count FROM tblTempERSumGradeClass1"
                Me.cmdResultsSummary.Parameters.Clear()
                reader = Me.cmdResultsSummary.ExecuteReader
                i = 0
                While reader.Read
                    i = (IIf(DBNull.Value.Equals(reader!count), "", reader!count))
                End While
                reader.Close()

                If i > 0 Then
                    MsgBox("Some printing on the reporting criteria going on.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
                    Exit Sub
                End If

                Me.cmdResultsSummary.Connection = conn
                Me.cmdResultsSummary.CommandType = CommandType.StoredProcedure
                Me.cmdResultsSummary.CommandText = "sprocTempERSumGradeClass"
                Me.cmdResultsSummary.Parameters.Clear()
                Me.cmdResultsSummary.Parameters.AddWithValue("@term", Me.cboTerm.Text.Trim)
                Me.cmdResultsSummary.Parameters.AddWithValue("@year", Me.cboYear.Text.Trim)
                Me.cmdResultsSummary.Parameters.AddWithValue("@class", Me.cboClass.Text.Trim)
                Me.cmdResultsSummary.ExecuteNonQuery()

                returnedAverages.Clear()
                returnedAverages = CalculateMeans("tblTempERSumGradeClass1", Me.cboClass.Text.Trim, 1)

                Dim RptResultsView As New crtAcadResultsGradeClass
                SetReportLogOn(RptResultsView)
                RptResultsView.SummaryInfo.ReportComments = fullName.Trim
                RptResultsView.RecordSelectionFormula = "({tblTempERSumGradeClass.className}=" & Chr(34) & Me.cboClass.Text.Trim & Chr(34) & ")"
                RptResultsView.RecordSelectionFormula += "AND ({tblTempERSumGradeClass.year}=" & Me.cboYear.Text.Trim & ")"
                RptResultsView.RecordSelectionFormula += "AND ({tblTempERSumGradeClass.termName}=" & Chr(34) & Me.cboTerm.Text.Trim & Chr(34) & ")"
                DirectCast(RptResultsView.ReportDefinition.ReportObjects("meanMark"), TextObject).Text = If(returnedAverages.Count > 0, returnedAverages.Item(0), "")
                DirectCast(RptResultsView.ReportDefinition.ReportObjects("meanGrade"), TextObject).Text = If(returnedAverages.Count > 0, returnedAverages.Item(2), "")

                frmResultsViewing.crtViewResultsSummary.ReportSource = RptResultsView
                frmResultsViewing.crtViewResultsSummary.Zoom(100)
                frmResultsViewing.crtViewResultsSummary.RefreshReport()
                frmResultsViewing.MdiParent = frmHome
                frmResultsViewing.Show()

                Me.cmdResultsSummary.Connection = conn
                Me.cmdResultsSummary.CommandType = CommandType.Text
                Me.cmdResultsSummary.CommandText = "DELETE FROM tblTempERSumGradeClass1"
                Me.cmdResultsSummary.Parameters.Clear()
                Me.cmdResultsSummary.ExecuteNonQuery()
            ElseIf Me.rbnsubjectstream.Checked = True Then
                If Me.cboStream.Text.Trim.Length <= 0 Then
                    MsgBox("Missing Stream Name.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
                    Exit Sub
                ElseIf Me.cboClass.Text.Trim.Length <= 0 Then
                    MsgBox("Missing Class.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
                    Exit Sub
                End If

                Me.cmdResultsSummary.Connection = conn
                Me.cmdResultsSummary.CommandType = CommandType.Text
                Me.cmdResultsSummary.CommandText = "SELECT COUNT(*) AS count FROM tblTempERSubSummaryStream1"
                Me.cmdResultsSummary.Parameters.Clear()
                reader = Me.cmdResultsSummary.ExecuteReader
                i = 0
                While reader.Read
                    i = (IIf(DBNull.Value.Equals(reader!count), "", reader!count))
                End While
                reader.Close()

                If i > 0 Then
                    MsgBox("Some printing on the reporting criteria going on.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
                    Exit Sub
                End If

                Me.cmdResultsSummary.Connection = conn
                Me.cmdResultsSummary.CommandType = CommandType.StoredProcedure
                Me.cmdResultsSummary.CommandTimeout = 0
                Me.cmdResultsSummary.CommandText = "sprocTempERSubSummaryStream"
                Me.cmdResultsSummary.Parameters.Clear()
                Me.cmdResultsSummary.Parameters.AddWithValue("@term", Me.cboTerm.Text.Trim)
                Me.cmdResultsSummary.Parameters.AddWithValue("@year", Me.cboYear.Text.Trim)
                Me.cmdResultsSummary.Parameters.AddWithValue("@class", Me.cboClass.Text.Trim)
                Me.cmdResultsSummary.Parameters.AddWithValue("@stream", Me.cboStream.Text.Trim)
                Me.cmdResultsSummary.ExecuteNonQuery()

                returnedAverages.Clear()
                returnedAverages = CalculateMeans("tblTempERSubSummaryStream1", Me.cboClass.Text.Trim, 2)

                Dim RptResultsView As New crtAcadResultsSubjectSummaryStream
                SetReportLogOn(RptResultsView)
                RptResultsView.SummaryInfo.ReportComments = fullName.Trim
                RptResultsView.RecordSelectionFormula = "({tblTempERSubSummaryStream.className}=" & Chr(34) & Me.cboClass.Text.Trim & Chr(34) & ")"
                RptResultsView.RecordSelectionFormula += "AND ({tblTempERSubSummaryStream.year}=" & Me.cboYear.Text.Trim & ")"
                RptResultsView.RecordSelectionFormula += "AND ({tblTempERSubSummaryStream.stream}=" & Chr(34) & Me.cboStream.Text.Trim & Chr(34) & ")"
                RptResultsView.RecordSelectionFormula += "AND ({tblTempERSubSummaryStream.termName}=" & Chr(34) & Me.cboTerm.Text.Trim & Chr(34) & ")"
                DirectCast(RptResultsView.ReportDefinition.ReportObjects("meanMark"), TextObject).Text = If(returnedAverages.Count > 0, returnedAverages.Item(0), "")
                DirectCast(RptResultsView.ReportDefinition.ReportObjects("meanGrade"), TextObject).Text = If(returnedAverages.Count > 0, returnedAverages.Item(2), "")

                frmResultsViewing.crtViewResultsSummary.ReportSource = RptResultsView
                frmResultsViewing.crtViewResultsSummary.Zoom(100)
                frmResultsViewing.crtViewResultsSummary.RefreshReport()
                frmResultsViewing.MdiParent = frmHome
                frmResultsViewing.Show()

                Me.cmdResultsSummary.Connection = conn
                Me.cmdResultsSummary.CommandType = CommandType.Text
                Me.cmdResultsSummary.CommandText = "DELETE FROM tblTempERSubSummaryStream1"
                Me.cmdResultsSummary.Parameters.Clear()
                Me.cmdResultsSummary.ExecuteNonQuery()
            ElseIf Me.rbnsubjectclass.Checked = True Then
                If Me.cboClass.Text.Trim.Length <= 0 Then
                    MsgBox("Missing Class.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
                    Exit Sub
                End If
                Me.cmdResultsSummary.Connection = conn
                Me.cmdResultsSummary.CommandType = CommandType.Text
                Me.cmdResultsSummary.CommandText = "SELECT COUNT(*) AS count FROM tblTempERSubSummaryClass1"
                Me.cmdResultsSummary.Parameters.Clear()
                reader = Me.cmdResultsSummary.ExecuteReader
                i = 0
                While reader.Read
                    i = (IIf(DBNull.Value.Equals(reader!count), "", reader!count))
                End While
                reader.Close()

                If i > 0 Then
                    MsgBox("Some printing on the reporting criteria going on.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
                    Exit Sub
                End If

                Me.cmdResultsSummary.Connection = conn
                Me.cmdResultsSummary.CommandType = CommandType.StoredProcedure
                Me.cmdResultsSummary.CommandTimeout = 0
                Me.cmdResultsSummary.CommandText = "sprocTempERSubSummaryClass"
                Me.cmdResultsSummary.Parameters.Clear()
                Me.cmdResultsSummary.Parameters.AddWithValue("@term", Me.cboTerm.Text.Trim)
                Me.cmdResultsSummary.Parameters.AddWithValue("@year", Me.cboYear.Text.Trim)
                Me.cmdResultsSummary.Parameters.AddWithValue("@class", Me.cboClass.Text.Trim)
                Me.cmdResultsSummary.ExecuteNonQuery()

                returnedAverages.Clear()
                returnedAverages = CalculateMeans("tblTempERSubSummaryClass1", Me.cboClass.Text.Trim, 2)

                Dim RptResultsView As New crtAcadResultsSubjectSummaryClass
                SetReportLogOn(RptResultsView)
                RptResultsView.SummaryInfo.ReportComments = fullName.Trim
                RptResultsView.RecordSelectionFormula = "({tblTempERSubSummaryClass.className}=" & Chr(34) & Me.cboClass.Text.Trim & Chr(34) & ")"
                RptResultsView.RecordSelectionFormula += "AND ({tblTempERSubSummaryClass.year}=" & Me.cboYear.Text.Trim & ")"
                RptResultsView.RecordSelectionFormula += "AND ({tblTempERSubSummaryClass.termName}=" & Chr(34) & Me.cboTerm.Text.Trim & Chr(34) & ")"
                DirectCast(RptResultsView.ReportDefinition.ReportObjects("meanMark"), TextObject).Text = If(returnedAverages.Count > 0, returnedAverages.Item(0), "")
                DirectCast(RptResultsView.ReportDefinition.ReportObjects("meanGrade"), TextObject).Text = If(returnedAverages.Count > 0, returnedAverages.Item(2), "")

                frmResultsViewing.crtViewResultsSummary.ReportSource = RptResultsView
                frmResultsViewing.crtViewResultsSummary.Zoom(100)
                frmResultsViewing.crtViewResultsSummary.RefreshReport()
                frmResultsViewing.MdiParent = frmHome
                frmResultsViewing.Show()

                Me.cmdResultsSummary.Connection = conn
                Me.cmdResultsSummary.CommandType = CommandType.Text
                Me.cmdResultsSummary.CommandText = "DELETE FROM tblTempERSubSummaryClass1"
                Me.cmdResultsSummary.Parameters.Clear()
                Me.cmdResultsSummary.ExecuteNonQuery()
            ElseIf Me.rbnSchSummaryMarkWise.Checked = True Then
                Me.cmdResultsSummary.Connection = conn
                Me.cmdResultsSummary.CommandType = CommandType.Text
                Me.cmdResultsSummary.CommandText = "SELECT COUNT(*) AS count FROM tblTempERSchSummaryMarks1"
                Me.cmdResultsSummary.Parameters.Clear()
                reader = Me.cmdResultsSummary.ExecuteReader
                i = 0
                While reader.Read
                    i = (IIf(DBNull.Value.Equals(reader!count), "", reader!count))
                End While
                reader.Close()

                If i > 0 Then
                    MsgBox("Some printing on the reporting criteria going on.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
                    Exit Sub
                End If

                Me.cmdResultsSummary.Connection = conn
                Me.cmdResultsSummary.CommandType = CommandType.StoredProcedure
                Me.cmdResultsSummary.CommandTimeout = 0
                Me.cmdResultsSummary.CommandText = "sprocTempERSchSummaryMarks"
                Me.cmdResultsSummary.Parameters.Clear()
                Me.cmdResultsSummary.Parameters.AddWithValue("@term", Me.cboTerm.Text.Trim)
                Me.cmdResultsSummary.Parameters.AddWithValue("@year", Me.cboYear.Text.Trim)
                Me.cmdResultsSummary.ExecuteNonQuery()

                returnedAverages.Clear()
                returnedAverages = CalculateMeans("tblTempERSchSummaryMarks1", "GRADE 8", 3)

                Dim RptResultsView As New crtAcadResultsSchoolSummaryMarkWise
                SetReportLogOn(RptResultsView)
                RptResultsView.SummaryInfo.ReportComments = fullName.Trim
                RptResultsView.SummaryInfo.ReportTitle = ("SCHOOL RESULTS SUMMARY MARK WISE FOR " &
                    Me.cboTerm.Text.Trim & " " & Me.cboYear.Text.Trim).ToUpper
                RptResultsView.RecordSelectionFormula = "({tblTempERSchSummaryMarks.termName}=" & Chr(34) & Me.cboTerm.Text.Trim & Chr(34) & ")"
                RptResultsView.RecordSelectionFormula += "AND ({tblTempERSchSummaryMarks.year}=" & Me.cboYear.Text.Trim & ")"
                DirectCast(RptResultsView.ReportDefinition.ReportObjects("meanMark"), TextObject).Text = If(returnedAverages.Count > 0, returnedAverages.Item(0), "")
                DirectCast(RptResultsView.ReportDefinition.ReportObjects("meanGrade"), TextObject).Text = If(returnedAverages.Count > 0, returnedAverages.Item(2), "")

                frmResultsViewing.crtViewResultsSummary.ReportSource = RptResultsView
                frmResultsViewing.crtViewResultsSummary.Zoom(100)
                frmResultsViewing.crtViewResultsSummary.RefreshReport()
                frmResultsViewing.MdiParent = frmHome
                frmResultsViewing.Show()

                Me.cmdResultsSummary.Connection = conn
                Me.cmdResultsSummary.CommandType = CommandType.Text
                Me.cmdResultsSummary.CommandText = "DELETE FROM tblTempERSchSummaryMarks1"
                Me.cmdResultsSummary.Parameters.Clear()
                Me.cmdResultsSummary.ExecuteNonQuery()

            ElseIf Me.rbnSchSummaryGradeWise.Checked = True Then
                Me.cmdResultsSummary.Connection = conn
                Me.cmdResultsSummary.CommandType = CommandType.Text
                Me.cmdResultsSummary.CommandText = "SELECT COUNT(*) AS count FROM tblTempERSchSummaryMarks1"
                Me.cmdResultsSummary.Parameters.Clear()
                reader = Me.cmdResultsSummary.ExecuteReader
                i = 0
                While reader.Read
                    i = (IIf(DBNull.Value.Equals(reader!count), "", reader!count))
                End While
                reader.Close()

                If i > 0 Then
                    MsgBox("Some printing on the reporting criteria going on.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
                    Exit Sub
                End If

                Me.cmdResultsSummary.Connection = conn
                Me.cmdResultsSummary.CommandType = CommandType.StoredProcedure
                Me.cmdResultsSummary.CommandTimeout = 0
                Me.cmdResultsSummary.CommandText = "sprocTempERSchSummaryMarks"
                Me.cmdResultsSummary.Parameters.Clear()
                Me.cmdResultsSummary.Parameters.AddWithValue("@term", Me.cboTerm.Text.Trim)
                Me.cmdResultsSummary.Parameters.AddWithValue("@year", Me.cboYear.Text.Trim)
                Me.cmdResultsSummary.ExecuteNonQuery()

                returnedAverages.Clear()
                returnedAverages = CalculateMeans("tblTempERSchSummaryMarks1", "GRADE 8", 3)

                Dim RptResultsView As New crtAcadResultsSchoolSummaryGradeWise
                SetReportLogOn(RptResultsView)
                RptResultsView.SummaryInfo.ReportComments = fullName.Trim
                RptResultsView.SummaryInfo.ReportTitle = ("SCHOOL RESULTS SUMMARY GRADE WISE FOR " &
                    Me.cboTerm.Text.Trim & " " & Me.cboYear.Text.Trim).ToUpper
                RptResultsView.RecordSelectionFormula = "({tblTempERSchSummaryGrades.termName}=" & Chr(34) & Me.cboTerm.Text.Trim & Chr(34) & ")"
                RptResultsView.RecordSelectionFormula += "AND ({tblTempERSchSummaryGrades.year}=" & Me.cboYear.Text.Trim & ")"
                DirectCast(RptResultsView.ReportDefinition.ReportObjects("meanMark"), TextObject).Text = If(returnedAverages.Count > 0, returnedAverages.Item(0), "")
                DirectCast(RptResultsView.ReportDefinition.ReportObjects("meanGrade"), TextObject).Text = If(returnedAverages.Count > 0, returnedAverages.Item(2), "")

                frmResultsViewing.crtViewResultsSummary.ReportSource = RptResultsView
                frmResultsViewing.crtViewResultsSummary.Zoom(100)
                frmResultsViewing.crtViewResultsSummary.RefreshReport()
                frmResultsViewing.MdiParent = frmHome
                frmResultsViewing.Show()

                Me.cmdResultsSummary.Connection = conn
                Me.cmdResultsSummary.CommandType = CommandType.Text
                Me.cmdResultsSummary.CommandText = "DELETE FROM tblTempERSchSummaryMarks1"
                Me.cmdResultsSummary.Parameters.Clear()
                Me.cmdResultsSummary.ExecuteNonQuery()

            ElseIf Me.rbnPointStream.Checked = True Then
                If Me.cboStream.Text.Trim.Length <= 0 Then
                    MsgBox("Missing Stream Name.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
                    Exit Sub
                ElseIf Me.cboClass.Text.Trim.Length <= 0 Then
                    MsgBox("Missing Class.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
                    Exit Sub
                End If
                Me.cmdResultsSummary.Connection = conn
                Me.cmdResultsSummary.CommandType = CommandType.Text
                Me.cmdResultsSummary.CommandText = "SELECT COUNT(*) AS count FROM tblTempERSumPointsStream1"
                Me.cmdResultsSummary.Parameters.Clear()
                reader = Me.cmdResultsSummary.ExecuteReader
                i = 0
                While reader.Read
                    i = (IIf(DBNull.Value.Equals(reader!count), "", reader!count))
                End While
                reader.Close()

                If i > 0 Then
                    MsgBox("Some printing on the reporting criteria going on.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
                    Exit Sub
                End If

                Me.cmdResultsSummary.Connection = conn
                Me.cmdResultsSummary.CommandType = CommandType.StoredProcedure
                Me.cmdResultsSummary.CommandText = "sprocTempERSumPointsStream"
                Me.cmdResultsSummary.Parameters.Clear()
                Me.cmdResultsSummary.Parameters.AddWithValue("@term", Me.cboTerm.Text.Trim)
                Me.cmdResultsSummary.Parameters.AddWithValue("@year", Me.cboYear.Text.Trim)
                Me.cmdResultsSummary.Parameters.AddWithValue("@class", Me.cboClass.Text.Trim)
                Me.cmdResultsSummary.Parameters.AddWithValue("@stream", Me.cboStream.Text.Trim)
                Me.cmdResultsSummary.ExecuteNonQuery()

                returnedAverages.Clear()
                returnedAverages = CalculateMeans("tblTempERSumPointsStream1", Me.cboClass.Text.Trim, 1)

                Dim RptResultsView As New crtAcadResultsPointsStream
                SetReportLogOn(RptResultsView)
                RptResultsView.SummaryInfo.ReportComments = fullName.Trim
                RptResultsView.RecordSelectionFormula = "({tblTempERSumPointsStream.className}=" & Chr(34) & Me.cboClass.Text.Trim & Chr(34) & ")"
                RptResultsView.RecordSelectionFormula += "AND ({tblTempERSumPointsStream.year}=" & Me.cboYear.Text.Trim & ")"
                RptResultsView.RecordSelectionFormula += "AND ({tblTempERSumPointsStream.stream}=" & Chr(34) & Me.cboStream.Text.Trim & Chr(34) & ")"
                RptResultsView.RecordSelectionFormula += "AND ({tblTempERSumPointsStream.termName}=" & Chr(34) & Me.cboTerm.Text.Trim & Chr(34) & ")"
                DirectCast(RptResultsView.ReportDefinition.ReportObjects("meanMark"), TextObject).Text = If(returnedAverages.Count > 0, returnedAverages.Item(0), "")
                DirectCast(RptResultsView.ReportDefinition.ReportObjects("meanGrade"), TextObject).Text = If(returnedAverages.Count > 0, returnedAverages.Item(2), "")

                frmResultsViewing.crtViewResultsSummary.ReportSource = RptResultsView
                frmResultsViewing.crtViewResultsSummary.Zoom(100)
                frmResultsViewing.crtViewResultsSummary.RefreshReport()
                frmResultsViewing.MdiParent = frmHome
                frmResultsViewing.Show()

                Me.cmdResultsSummary.Connection = conn
                Me.cmdResultsSummary.CommandType = CommandType.Text
                Me.cmdResultsSummary.CommandText = "DELETE FROM tblTempERSumPointsStream1"
                Me.cmdResultsSummary.Parameters.Clear()
                Me.cmdResultsSummary.ExecuteNonQuery()
            ElseIf Me.rbnPointClass.Checked = True Then
                If Me.cboClass.Text.Trim.Length <= 0 Then
                    MsgBox("Missing Class.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
                    Exit Sub
                End If
                Me.cmdResultsSummary.Connection = conn
                Me.cmdResultsSummary.CommandType = CommandType.Text
                Me.cmdResultsSummary.CommandText = "SELECT COUNT(*) AS count FROM tblTempERSumPointsClass1"
                Me.cmdResultsSummary.Parameters.Clear()
                reader = Me.cmdResultsSummary.ExecuteReader
                i = 0
                While reader.Read
                    i = (IIf(DBNull.Value.Equals(reader!count), "", reader!count))
                End While
                reader.Close()

                If i > 0 Then
                    MsgBox("Some printing on the reporting criteria going on.", MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
                    Exit Sub
                End If

                Me.cmdResultsSummary.Connection = conn
                Me.cmdResultsSummary.CommandType = CommandType.StoredProcedure
                Me.cmdResultsSummary.CommandText = "sprocTempERSumPointsClass"
                Me.cmdResultsSummary.Parameters.Clear()
                Me.cmdResultsSummary.Parameters.AddWithValue("@term", Me.cboTerm.Text.Trim)
                Me.cmdResultsSummary.Parameters.AddWithValue("@year", Me.cboYear.Text.Trim)
                Me.cmdResultsSummary.Parameters.AddWithValue("@class", Me.cboClass.Text.Trim)
                Me.cmdResultsSummary.ExecuteNonQuery()

                returnedAverages.Clear()
                returnedAverages = CalculateMeans("tblTempERSumPointsClass1", Me.cboClass.Text.Trim, 1)

                Dim RptResultsView As New crtAcadResultsPointsClass
                SetReportLogOn(RptResultsView)
                RptResultsView.SummaryInfo.ReportComments = fullName.Trim
                RptResultsView.RecordSelectionFormula = "({tblTempERSumPointsClass.className}=" & Chr(34) & Me.cboClass.Text.Trim & Chr(34) & ")"
                RptResultsView.RecordSelectionFormula += "AND ({tblTempERSumPointsClass.year}=" & Me.cboYear.Text.Trim & ")"
                RptResultsView.RecordSelectionFormula += "AND ({tblTempERSumPointsClass.termName}=" & Chr(34) & Me.cboTerm.Text.Trim & Chr(34) & ")"
                DirectCast(RptResultsView.ReportDefinition.ReportObjects("meanMark"), TextObject).Text = If(returnedAverages.Count > 0, returnedAverages.Item(0), "")
                DirectCast(RptResultsView.ReportDefinition.ReportObjects("meanGrade"), TextObject).Text = If(returnedAverages.Count > 0, returnedAverages.Item(2), "")

                frmResultsViewing.crtViewResultsSummary.ReportSource = RptResultsView
                frmResultsViewing.crtViewResultsSummary.Zoom(100)
                frmResultsViewing.crtViewResultsSummary.RefreshReport()
                frmResultsViewing.MdiParent = frmHome
                frmResultsViewing.Show()

                Me.cmdResultsSummary.Connection = conn
                Me.cmdResultsSummary.CommandType = CommandType.Text
                Me.cmdResultsSummary.CommandText = "DELETE FROM tblTempERSumPointsClass1"
                Me.cmdResultsSummary.Parameters.Clear()
                Me.cmdResultsSummary.ExecuteNonQuery()

            ElseIf Me.rbnbroadstreamExam.Checked = True Then


            ElseIf Me.rbnbroadclassExam.Checked = True Then
                Me.cboExamType.Enabled = True
                Me.cboExamName.Enabled = True

            End If
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            Me.Cursor = Cursors.Arrow
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub
    Private Sub rbnbroadstream_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbnbroadstream.CheckedChanged
        DisableExamSelection()
        If Me.rbnbroadstream.Checked = True Then
            Me.cboStream.Enabled = True
        ElseIf Me.rbnbroadstream.Checked = False Then
            Me.cboStream.Enabled = True
        End If
    End Sub

    Private Sub rbnbroadclass_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbnbroadclass.CheckedChanged
        DisableExamSelection()
        If Me.rbnbroadclass.Checked = True Then
            Me.cboStream.SelectedIndex = -1
            Me.cboStream.Enabled = False
        ElseIf Me.rbnbroadclass.Checked = False Then
            Me.cboStream.Enabled = True
        End If
    End Sub

    Private Sub rbnbroadgradestream_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbnbroadgradestream.CheckedChanged
        DisableExamSelection()
        If Me.rbnbroadgradestream.Checked = True Then
            Me.cboStream.Enabled = True
        ElseIf Me.rbnbroadgradestream.Checked = False Then
            Me.cboStream.Enabled = True
        End If
    End Sub

    Private Sub rbnbroadgradeclass_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbnbroadgradeclass.CheckedChanged
        DisableExamSelection()
        If Me.rbnbroadgradeclass.Checked = True Then
            Me.cboStream.SelectedIndex = -1
            Me.cboStream.Enabled = False
        ElseIf Me.rbnbroadgradeclass.Checked = False Then
            Me.cboStream.Enabled = True
        End If
    End Sub

    Private Sub rbnsubjectstream_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbnsubjectstream.CheckedChanged
        DisableExamSelection()
        If Me.rbnsubjectstream.Checked = True Then
            Me.cboStream.Enabled = True
        ElseIf Me.rbnsubjectstream.Checked = False Then
            Me.cboStream.Enabled = True
        End If
    End Sub

    Private Sub rbnsubjectclass_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbnsubjectclass.CheckedChanged
        DisableExamSelection()
        If Me.rbnsubjectclass.Checked = True Then
            Me.cboStream.SelectedIndex = -1
            Me.cboStream.Enabled = False
        ElseIf Me.rbnsubjectclass.Checked = False Then
            Me.cboStream.Enabled = True
        End If
    End Sub

    Private Sub rbSchSummaryMarkWise_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbnSchSummaryMarkWise.CheckedChanged
        DisableExamSelection()
        If Me.rbnSchSummaryMarkWise.Checked = True Then
            Me.cboStream.SelectedIndex = -1
            Me.cboClass.SelectedIndex = -1
            Me.cboStream.Enabled = False
            Me.cboClass.Enabled = False
        ElseIf Me.rbnSchSummaryMarkWise.Checked = False Then
            Me.cboStream.Enabled = True
            Me.cboClass.Enabled = True
        End If
    End Sub

    Private Sub rbSchSummaryGradeWise_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbnSchSummaryGradeWise.CheckedChanged
        DisableExamSelection()
        If Me.rbnSchSummaryGradeWise.Checked = True Then
            Me.cboStream.SelectedIndex = -1
            Me.cboClass.SelectedIndex = -1
            Me.cboStream.Enabled = False
            Me.cboClass.Enabled = False
        ElseIf Me.rbnSchSummaryGradeWise.Checked = False Then
            Me.cboStream.Enabled = True
            Me.cboClass.Enabled = True
        End If
    End Sub
    Private Sub rbPointStream_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbnPointStream.CheckedChanged
        DisableExamSelection()
        If Me.rbnPointStream.Checked = True Then
            Me.cboStream.Enabled = True
        ElseIf Me.rbnPointStream.Checked = False Then
            Me.cboStream.Enabled = True
        End If
    End Sub
    Private Sub rbPointClass_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbnPointClass.CheckedChanged
        DisableExamSelection()
        If Me.rbnPointClass.Checked = True Then
            Me.cboStream.SelectedIndex = -1
            Me.cboStream.Enabled = False
        ElseIf Me.rbnPointClass.Checked = False Then
            Me.cboStream.Enabled = True
        End If
    End Sub

    Private Sub btnClearPrinting_Click(sender As Object, e As EventArgs) Handles btnClearPrinting.Click
        Dim result As MsgBoxResult = MsgBox("Are you sure you want to clear all printing going on?", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.YesNo, "Confirm Transaction")
        If result = MsgBoxResult.No Then
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()

            Me.Cursor = Cursors.WaitCursor
            Me.cmdResultsSummary.Connection = conn
            Me.cmdResultsSummary.CommandType = CommandType.StoredProcedure
            Me.cmdResultsSummary.CommandText = "SprocClearPrinting"
            Me.cmdResultsSummary.Parameters.Clear()
            Me.cmdResultsSummary.ExecuteNonQuery()

            MsgBox("Printing Cleared Successfully!", MsgBoxStyle.Information + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "SuccessFull Transaction")

        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            Me.Cursor = Cursors.Arrow
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub

    Private Sub cboExamType_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboExamType.SelectedIndexChanged
        If Me.cboExamType.Text.Trim.Length <= 0 Then
            Exit Sub
        End If
        Try
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If
            dbconnection()

            Me.cboExamName.Items.Clear()
            Me.cboExamName.Text = ""
            Me.cboExamName.SelectedIndex = -1

            Me.cmdResultsSummary.Connection = conn
            Me.cmdResultsSummary.CommandType = CommandType.Text
            Me.cmdResultsSummary.CommandText = "SELECT DISTINCT examName FROM tblExamNames WHERE (status=1) AND (examType=@examType) ORDER BY examName"
            Me.cmdResultsSummary.Parameters.Clear()
            Me.cmdResultsSummary.Parameters.AddWithValue("@examType", Me.cboExamType.Text.Trim)
            reader = Me.cmdResultsSummary.ExecuteReader
            If reader.HasRows Then
                While reader.Read
                    Me.cboExamName.Items.Add(IIf(DBNull.Value.Equals(reader!examName), "", reader!examName))
                End While
            End If
            reader.Close()
        Catch ex As Exception
            MsgBox(Err.GetException.Message, MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal + MsgBoxStyle.OkOnly, "Error Detected")
        Finally
            Me.Cursor = Cursors.Arrow
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub

    Private Sub rbnbroadstreamExam_CheckedChanged(sender As Object, e As EventArgs) Handles rbnbroadstreamExam.CheckedChanged
        'DisableExamSelection()
        If Me.rbnbroadstreamExam.Checked = True Then
            Me.cboExamType.Enabled = True
            Me.cboExamName.Enabled = True
            Me.cboStream.Enabled = True
        ElseIf Me.rbnbroadstreamExam.Checked = False Then
            Me.cboExamType.Enabled = False
            Me.cboExamName.Enabled = False
            Me.cboStream.Enabled = True
        End If
    End Sub

    Private Sub DisableExamSelection()
        Me.cboExamType.SelectedIndex = -1
        Me.cboExamName.SelectedIndex = -1
        Me.cboExamType.Enabled = False
        Me.cboExamName.Enabled = False
    End Sub
    Private Sub rbnbroadclassExam_CheckedChanged(sender As Object, e As EventArgs) Handles rbnbroadclassExam.CheckedChanged
        'DisableExamSelection()
        If Me.rbnbroadclassExam.Checked = True Then
            Me.cboExamType.Enabled = True
            Me.cboExamName.Enabled = True
            Me.cboStream.SelectedIndex = -1
            Me.cboStream.Enabled = False
        ElseIf Me.rbnbroadclassExam.Checked = False Then
            Me.cboExamType.Enabled = False
            Me.cboExamName.Enabled = False
            Me.cboStream.Enabled = True
        End If
    End Sub
End Class