Public Class Chores
    Private Sub Chores_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim tooltip As New ToolTip
        tooltip.SetToolTip(Button1, "Dashboard")
        tooltip.SetToolTip(Button2, "Mark All as Complete")
        tooltip.SetToolTip(Button3, "Refresh")
        tooltip.SetToolTip(Button4, "Delete")
        tooltip.SetToolTip(Button5, "Edit")
        tooltip.SetToolTip(Button6, "Submit")
        tooltip.SetToolTip(Button7, "Highlight")
        tooltip.SetToolTip(Button8, "Filter")
        tooltip.SetToolTip(Button9, "Sort")
    End Sub
End Class