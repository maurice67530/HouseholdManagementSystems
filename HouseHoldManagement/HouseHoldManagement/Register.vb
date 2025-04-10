Imports System.Data.OleDb
Public Class Register
      Public Property conn As New OleDbConnection(connectionString)

    Public Const connectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\Xiluva\Source\Repos\maurice67530\HouseholdManagementSystems\HMS.accdb;Persist Security Info=False;"

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        Try
            conn.Open()

            Dim checkCmd As New OleDbCommand("SELECT COUNT(*) FROM Users WHERE UserName = ? OR Password = ?", conn)
            checkCmd.Parameters.AddWithValue("?", TextBox3.Text)
            checkCmd.Parameters.AddWithValue("?", TextBox2.Text)
            Dim userExists As Integer = Convert.ToInt32(checkCmd.ExecuteScalar())

            If userExists > 0 Then
                MessageBox.Show("Username or Password already exists.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If


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

    Private Sub Register_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ComboBox1.Items.AddRange(New String() {"Admin", "Member", "Finance", "Chef"})
    End Sub

End Class