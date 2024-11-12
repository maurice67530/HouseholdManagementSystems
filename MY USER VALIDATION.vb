Public Class MY_USER_VALIDATION

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        'If MaskedTextBox1.Text = "0768075898" Then
        '    MsgBox("Welcome to my household system")
        'Else
        '    MsgBox("I don't recogonize this number")
        'End If
        If MaskedTextBox2.Text = "20030" Then
            MsgBox("welcome to my household system")
            DashBoard.ShowDialog()
        Else
            MsgBox("I don't recogonize this number")

        End If
    End Sub
End Class