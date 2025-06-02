
Imports System.Data.OleDb


Public Class Login

    Public Property conn As New OleDbConnection(connectionString)

    ' Public Const connectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\Mudzunga\Source\Repos\maurice67530\HouseholdManagementSystems\HMS.accdb"



    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        'Dim username = TextBox1.Text.Trim()
        'Dim password = TextBox2.Text.Trim()

        'Dim cmd As New OleDbCommand("SELECT FullNames, Role FROM Users WHERE Username = ? AND [Password] = ?", conn)
        'cmd.Parameters.AddWithValue("?", username)
        'cmd.Parameters.AddWithValue("?", password)

        'conn.Open()
        'Dim reader As OleDbDataReader = cmd.ExecuteReader()

        'If reader.Read() Then
        '    Dim FullNames As String = reader("FullNames").ToString()
        '    Dim role As String = reader("Role").ToString()

        '    MessageBox.Show("Login successful. :" & FullNames & role, "Welcome!", MessageBoxButtons.OK, MessageBoxIcon.Information)

        '    Dashboard.TextBox3.Text = username
        '    Dashboard.TextBox2.Text = role
        '    Dashboard.Label12.Text = FullNames

        '    Budget.TextBox6.Text = FullNames
        '    Budget.TextBox7.Text = role


        '    In_App_Message.TextBox1.Text = username
        '    In_App_Message.Label6.Text = "Logged in as:" & username
        '    Me.Hide()


        '    Dashboard.Show()
        '    '   Budget.Show()
        '    Me.Hide()
        'Else
        '    MessageBox.Show("Invalid username or password.")
        'End If

        'conn.Close()



        conn.Open()
        Dim username = TextBox1.Text.Trim()
        Dim password = TextBox2.Text.Trim()
        Dim chore As New Logging() With {.username = TextBox1.Text}

        Dim cmd As New OleDbCommand("SELECT FullNames, Role FROM Users WHERE Username = ? AND [Password] = ?", conn)
        Dim cmd2 As New OleDbCommand("INSERT INTO Login ([userName]) VALUES (?)", conn)

        cmd2.Parameters.Clear()
        cmd2.Parameters.AddWithValue("@userName", chore.username)
        cmd2.ExecuteNonQuery()

        cmd.Parameters.AddWithValue("?", username)
        cmd.Parameters.AddWithValue("?", password)


        Dim reader As OleDbDataReader = cmd.ExecuteReader()

        If reader.Read() Then
            Dim FullNames As String = reader("FullNames").ToString()
            Dim role As String = reader("Role").ToString()


            '''''

            ' After successful login
            Dim dashboard As New Dashboard()
            dashboard.CurrentUserName = TextBox1.Text.Trim() ' Pass the username from the TextBox


            MessageBox.Show("Login successful. :" & FullNames & role, "Welcome!", MessageBoxButtons.OK, MessageBoxIcon.Information)

            Dashboard.TextBox3.Text = username
            Dashboard.TextBox2.Text = role
            Dashboard.Label12.Text = FullNames

            Budget.TextBox6.Text = FullNames
            Budget.TextBox7.Text = role


            In_App_Message.TextBox1.Text = username
            In_App_Message.Label6.Text = "Logged in as:" & username


            Household_Document.TextBox5.Text = username
            Me.Hide()


            Dashboard.Show()
            '   Budget.Show()
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