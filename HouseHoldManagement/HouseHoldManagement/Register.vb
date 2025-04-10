Imports System.Data.OleDb
Public Class Register
    Private Sub Register_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ComboBox1.Items.AddRange(New String() {"Admin", "Member", "Finance", "Chef"})
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click


        Try
            conn.Open()
            ' Check if Username or Email already exists
            Dim checkCmd As New OleDbCommand("SELECT COUNT(*) FROM Users WHERE UserName = ? OR Password = ?", conn)
            checkCmd.Parameters.AddWithValue("?", TextBox3.Text)
            checkCmd.Parameters.AddWithValue("?", TextBox2.Text)
            Dim userExists As Integer = Convert.ToInt32(checkCmd.ExecuteScalar())

            If userExists > 0 Then
                MessageBox.Show("UserName or Password already exists.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If

            ' Insert new user
            Dim cmd As New OleDbCommand("INSERT INTO Users ([FullNames], [UserName], [Email], [Password], [Role], [DateCreated]) VALUES (?, ?, ?, ?, ?, ?)", conn)
            cmd.Parameters.AddWithValue("?", TextBox1.Text)
            cmd.Parameters.AddWithValue("?", TextBox3.Text)
            cmd.Parameters.AddWithValue("?", TextBox4.Text)
            cmd.Parameters.AddWithValue("?", TextBox2.Text)
            cmd.Parameters.AddWithValue("?", ComboBox1.SelectedItem.ToString)
            cmd.Parameters.AddWithValue("?", DateTime.Now)

            cmd.ExecuteNonQuery()
            MessageBox.Show("Registration successful!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)

            Me.Hide()
            Login.Show()

        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message)
        Finally
            conn.Close()
        End Try
    End Sub
End Class
