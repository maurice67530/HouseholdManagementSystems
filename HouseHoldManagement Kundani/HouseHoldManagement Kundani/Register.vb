Imports System.Data.OleDb
Public Class Register
    Public Property conn As New OleDbConnection(HouseHoldManagment_Module.connectionString)

    Public Const connectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\Mudzunga\Source\Repos\maurice67530\HouseholdManagementSystems\HMS.accdb"


    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim FullNames = TextBox1.Text.Trim()
        Dim username = TextBox3.Text.Trim()
        Dim password = TextBox2.Text.Trim()
        Dim Email = TextBox4.Text.Trim()
        Dim family = TextBox5.Text.Trim()
        Dim Role = ComboBox1.SelectedItem.ToString().Trim()
        Dim DateCreated = DateTimePicker1.Text.Trim()

        ' Check if the Family already exists
        Dim checkCmd As New OleDbCommand("SELECT COUNT(*) FROM Users WHERE Family = ?", conn)
        checkCmd.Parameters.AddWithValue("?", family)
        conn.Open()
        Dim count As Integer = Convert.ToInt32(checkCmd.ExecuteScalar())
        conn.Close()

        If count > 0 Then
            MessageBox.Show("This family already exists. Please use a different family name.", "Duplicate Family", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        Else
            ' Proceed with saving only if family doesn't exist
            Dim cmd As New OleDbCommand("INSERT INTO Users ([FullNames], [UserName], [Email], [Password], [Role], [DateCreated], [Family]) VALUES (?, ?, ?, ?, ?, ?, ?)", conn)
            cmd.Parameters.AddWithValue("?", FullNames)
            cmd.Parameters.AddWithValue("?", username)
            cmd.Parameters.AddWithValue("?", Email)
            cmd.Parameters.AddWithValue("?", password)
            cmd.Parameters.AddWithValue("?", Role)
            cmd.Parameters.AddWithValue("?", DateCreated)
            cmd.Parameters.AddWithValue("?", family)

            conn.Open()
            cmd.ExecuteNonQuery()
            conn.Close()

            MessageBox.Show("Registered successfully with Family: " & family, "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If
    End Sub

    Private Sub Register_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ComboBox1.Items.AddRange(New String() {"Admin", "Member", "Finance", "Chef"})
    End Sub
End Class