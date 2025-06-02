Imports System.Data.OleDb
Public Class Register
    Public Property conn As New OleDbConnection(connectionString)

    'Public Const connectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\Mudzunga\Source\Repos\maurice67530\HouseholdManagementSystems\HMS.accdb"


    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim FullNames = TextBox1.Text.Trim()
        Dim username = TextBox3.Text.Trim()
        Dim password = TextBox2.Text.Trim()
        Dim Email = TextBox4.Text.Trim()
        Dim family = TextBox5.Text.Trim()
        Dim Role = ComboBox1.SelectedItem.ToString().Trim()
        Dim DateCreated = DateTimePicker1.Text.Trim()
        Dim Age = TextBox7.Text.Trim()
        ''Dim Picture = PictureBox1.ImageLocation.Trim()
        Dim Preference = ComboBox2.Text.Trim()

        ' Check if the Family already exists
        Dim checkCmd As New OleDbCommand("SELECT COUNT(*) FROM Users WHERE Family = ?", conn)
        checkCmd.Parameters.AddWithValue("?", family)
        conn.Open()
        ' Dim count As Integer = Convert.ToInt32(checkCmd.ExecuteScalar())
        conn.Close()

        'If count > 0 Then
        '    MessageBox.Show("This family already exists. Please use a different family name.", "Duplicate Family", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        'Else
        ' Proceed with saving only if family doesn't exist
        Dim cmd As New OleDbCommand("INSERT INTO Users ([FullNames], [UserName], [Email], [Password], [Role], [DateCreated]) VALUES (?, ?, ?, ?, ?, ?)", conn)
        cmd.Parameters.AddWithValue("?", FullNames)
        cmd.Parameters.AddWithValue("?", username)
            cmd.Parameters.AddWithValue("?", Email)
            cmd.Parameters.AddWithValue("?", password)
            cmd.Parameters.AddWithValue("?", Role)
        cmd.Parameters.AddWithValue("?", DateCreated)


        cmd.Parameters.AddWithValue("?", Preference)

        conn.Open()
            cmd.ExecuteNonQuery()
            conn.Close()

            MessageBox.Show("Registered successfully with Family: " & family, "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
        '   End If
    End Sub




    Private Sub Register_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ComboBox1.Items.AddRange(New String() {"Admin", "Member", "Finance", "Chef"})
        CheckDatabaseConnection(StatusLabel)
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim OpenFileDialog As New OpenFileDialog()
        OpenFileDialog.Filter = "Bitmaps (*.jfif)|*.jfif"
        If OpenFileDialog.ShowDialog() = DialogResult.OK Then
            PictureBox1.ImageLocation = OpenFileDialog.FileName
            'TextBox6.Text = OpenFileDialog.FileName
        End If
    End Sub

    Private Sub TextBox7_TextChanged(sender As Object, e As EventArgs) Handles TextBox7.TextChanged
        CalculateAge()
    End Sub
    Private Sub CalculateAge()
        If TextBox7.Text.Length = 13 Then
            Dim idNumber As String = TextBox7.Text
            Dim birthYear As Integer = CInt(idNumber.Substring(0, 2))
            Dim birthMonth As Integer = CInt(idNumber.Substring(2, 2))
            Dim birthDay As Integer = CInt(idNumber.Substring(4, 2))

            ' Adjust century
            If birthYear < 50 Then
                birthYear += 2000
            Else
                birthYear += 1900
            End If

            Dim birthDate As DateTime = New DateTime(birthYear, birthMonth, birthDay)
            Dim age As Integer = DateTime.Now.Year - birthDate.Year

            If DateTime.Now.Month < birthDate.Month OrElse (DateTime.Now.Month = birthDate.Month AndAlso DateTime.Now.Day < birthDate.Day) Then
                age -= 1
            End If

            Label11.Text = age.ToString()
        Else
            Label11.Text = "Invalid ID"
        End If
    End Sub
End Class