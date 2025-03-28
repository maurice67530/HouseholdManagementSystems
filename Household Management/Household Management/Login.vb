Imports System.IO
Imports System.Data.OleDb
Public Class Login
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Try
            'conn.Open()
            Dim cmd As New OleDbCommand("SELECT Role FROM Users WHERE Username = ? AND Password = ?", conn)
            cmd.Parameters.AddWithValue("?", TextBox1.Text)
            cmd.Parameters.AddWithValue("?", TextBox2.Text) ' Ideally, hash and compare passwords

            Dim role As Object = cmd.ExecuteScalar()

            If role IsNot Nothing Then
                MessageBox.Show("Login Successful!", "Welcome", MessageBoxButtons.OK, MessageBoxIcon.Information)

                '' Redirect based on role
                'Select Case role.ToString()
                '    Case "Admin"
                '        AdminDashboard.Show()
                '    Case "User"
                '        UserDashboard.Show()
                '    Case Else
                '        MessageBox.Show("Unknown role.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                'End Select

                Me.Hide() ' Hide Login Form
            Else
                MessageBox.Show("Invalid Username or Password.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If

        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message)
        Finally
            ' conn.Close()
        End Try
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

    End Sub
End Class