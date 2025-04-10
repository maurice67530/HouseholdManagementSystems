Public Class Personnel

    ' Create a ToolTip object
    Private toolTip As New ToolTip()
    Private toolTip1 As New ToolTip()
    Private Sub Personnel_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        ' Initialize ToolTip properties (optional)
        ToolTip.AutoPopDelay = 5000
        ToolTip.InitialDelay = 500
        ToolTip.ReshowDelay = 200
        ToolTip.ShowAlways = True

        toolTip1.SetToolTip(Button3, "Back")
        toolTip1.SetToolTip(Button4, "Add a Picture")
        toolTip1.SetToolTip(Button2, "Edit")
        toolTip1.SetToolTip(Button5, "Delete")
        toolTip1.SetToolTip(Button6, "Clear")
        toolTip1.SetToolTip(Button7, "Daily tasks")
        toolTip1.SetToolTip(Button1, "Save")
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click

    End Sub
End Class