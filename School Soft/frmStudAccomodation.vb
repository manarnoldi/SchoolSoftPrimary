Public Class frmStudAccomodation

    Private Sub frmStudAccomodation_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub

    Private Sub frmStudAccomodation_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint
        If Me.IsMdiChild Then
            Dim PnlLoc As New Point
            PnlLoc.X = CInt((Me.Width - Me.pnlStudAccom.Width) / 2)
            PnlLoc.Y = CInt((Me.Height - Me.pnlStudAccom.Height) / 2.5)
            Me.pnlStudAccom.Location = PnlLoc
        Else
            Me.pnlStudAccom.Dock = DockStyle.Fill
        End If
    End Sub
End Class