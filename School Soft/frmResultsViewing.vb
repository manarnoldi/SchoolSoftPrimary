Public Class frmResultsViewing

    Private Sub frmResultsViewing_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub

    Private Sub frmResultsViewing_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint
        If Me.IsMdiChild Then
            Dim PnlLoc As New Point
            PnlLoc.X = CInt((Me.Width - Me.pnlResultsViewing.Width) / 2)
            PnlLoc.Y = CInt((Me.Height - Me.pnlResultsViewing.Height) / 2.5)
            Me.pnlResultsViewing.Location = PnlLoc
        Else
            Me.pnlResultsViewing.Dock = DockStyle.Fill
        End If
    End Sub
End Class