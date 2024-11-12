Public Class Inventory

    Dim btmItemsCount As Object

    Private Property item As Object

    Private Sub Label5_Click(sender As System.Object, e As System.EventArgs) Handles Label5.Click, Label16.Click

    End Sub

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs)
        PHOTO_GALLERY.ShowDialog()
    End Sub

    Private Sub Button1_Click_1(sender As System.Object, e As System.EventArgs) Handles Button1.Click, Button4.Click
        DashBoard.Show()
        Me.Close()
    End Sub

    Private Sub btmDisplayGroceries_Click(sender As System.Object, e As System.EventArgs)

    End Sub

    Private Sub ListBox1_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles ListBox1.SelectedIndexChanged

    End Sub
End Class

