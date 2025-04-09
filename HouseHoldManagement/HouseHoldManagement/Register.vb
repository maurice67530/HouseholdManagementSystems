Public Class Register
    Private Sub Register_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        ComboBox1.Items.AddRange(New String() {"Admin", "Member", "Finance", "Chef"})

        Dim tooltip As New ToolTip()
        tooltip.SetToolTip(Button1, "Register")
        tooltip.SetToolTip(Button2, "Cancel")

    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        End
    End Sub
End Class