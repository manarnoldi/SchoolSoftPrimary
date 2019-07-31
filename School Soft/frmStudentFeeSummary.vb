Public Class frmStudentFeeSummary

    Private Sub frmStudentFeeSummary_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub

    Private Sub frmStudentFeeSummary_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint
        If Me.IsMdiChild Then
            Dim PnlLoc As New Point
            PnlLoc.X = CInt((Me.Width - Me.pnlStudFeeSummary.Width) / 2)
            PnlLoc.Y = CInt((Me.Height - Me.pnlStudFeeSummary.Height) / 2.5)
            Me.pnlStudFeeSummary.Location = PnlLoc
        Else
            Me.pnlStudFeeSummary.Dock = DockStyle.Fill
        End If
    End Sub
End Class