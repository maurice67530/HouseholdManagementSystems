Public Class Photo
    Private toolTip1 As New ToolTip()
    Private Sub Photo_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        toolTip1.SetToolTip(Button1, "Dashboard")
        toolTip1.SetToolTip(Button2, "Save Album")
    End Sub
End Class