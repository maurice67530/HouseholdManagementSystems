
Imports System.Data.OleDb


Public Class Login

    Public Property conn As New OleDbConnection(HouseHoldManagment_Module.connectionString)

    Public Const connectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\Mudzunga\Source\Repos\maurice67530\HouseholdManagementSystems\HMS.accdb"


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim username = TextBox1.Text.Trim()
        Dim password = TextBox2.Text.Trim()

        Dim cmd As New OleDbCommand("SELECT Family, Role FROM Users WHERE Username = ? AND [Password] = ?", conn)
        cmd.Parameters.AddWithValue("?", username)
        cmd.Parameters.AddWithValue("?", password)

        conn.Open()
        Dim reader As OleDbDataReader = cmd.ExecuteReader()

        If reader.Read() Then
            Dim family As String = reader("Family").ToString()
            Dim role As String = reader("Role").ToString()

            MessageBox.Show("Login successful. Family: " & family, "Welcome!", MessageBoxButtons.OK, MessageBoxIcon.Information)

            Dashboard.TextBox3.Text = username
            Dashboard.TextBox2.Text = role
            Dashboard.Label12.Text = family

            Dashboard.Show()
            Me.Hide()
        Else
            MessageBox.Show("Invalid username or password.")
        End If

        conn.Close()
    End Sub


    Private Sub Login_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim tooltip As New ToolTip()
        tooltip.SetToolTip(Button1, "Login")
        tooltip.SetToolTip(Button2, "Register")
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Register.ShowDialog()
    End Sub
End Class