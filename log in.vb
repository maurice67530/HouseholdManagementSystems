
Imports System.Data.OleDb
Imports System.Windows.Forms
Imports System.Drawing
Imports System.IO
Public Class MY_USER_VALIDATION



    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click


        If TextBox1.Text = "Austin" Then
            MsgBox("Welcome to my household system")
        Else
            MsgBox("I don't recogonize this password")
        End If
        If TextBox2.Text = "Ramuhashi" Then
            MsgBox("welcome to my household system")



            DashBoard.ShowDialog()
        Else
            MsgBox("I don't recogonize this password")

        End If

        'Try
        '    Debug.WriteLine("Entering BtnLogin")

        '    Dim UserName As String = TextBox1.Text.Trim()
        '    Dim Password As String = TextBox2.Text.Trim()
        '    'Dim HashedPassword As String = HashPassword(Password) ' Hash password before checking

        '    Dim conn As New OleDbConnection(connectionString)
        '    conn.Open()

        '    ' Fetch user role from database
        '    Dim query As String = "SELECT Role FROM LoginDetails WHERE UserName = @UserName AND Password = @Password"

        '    Using cmd As New OleDbCommand(query, conn)
        '        cmd.Parameters.AddWithValue("@UserName", UserName)
        '        cmd.Parameters.AddWithValue("@Password", Password)

        '        Dim cmbxRole As Object = cmd.ExecuteScalar()

        '        If cmbxRole IsNot Nothing Then
        '            MessageBox.Show("Login Successful!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)

        '            ' Open dashboard and set user role
        '            'Dim DashBoard As New Dashboard()
        '            'DashBoard.CurrentUserRole = userRole.ToString() ' Assign role
        '            'DashBoard.TextBox1.Text = UserName & "" & Password & userRole.ToString() ' Assign role
        '            DashBoard.TextBox1.Text = UserName & "   " & Password & "   "
        '            DashBoard.TextBox2.Text = cmbxRole.ToString()
        '            DashBoard.Show()
        '            Hide()
        '        Else
        '            MessageBox.Show("Invalid Username or Password.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        '        End If
        '    End Using

        'Catch ex As OleDbException
        '    Debug.WriteLine($"Error: {ex.Message}")
        '    Debug.WriteLine($"Stack Trace: {ex.StackTrace}")
        '    MessageBox.Show("A database error occurred. Please try again later.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)

        '    Debug.WriteLine("Exiting BtnLogin")
        '    Debug.WriteLine("leaving Authentication")
        '    Visible = True
        'End Try
        '    Using conn As New OleDbConnection(Module1.connectionString)


        '        conn.Open()
        '        Dim query As String = "SELECT COUNT(*) FROM Logindetails WHERE FirstName = ? AND LastName = ? AND Role = ?"
        '        Using cmd As New OleDbCommand(query, conn)

        '            cmd.Parameters.AddWithValue("@FirstName", TextBox1.Text)
        '            cmd.Parameters.AddWithValue("@LastName", TextBox2.Text)


        '            ''Dim userExists As Integer = Convert.ToInt32(cmd.ExecuteScalar())

        '            If TextBox1.Text = "Austin" Then
        '                MsgBox("Welcome To my household system")
        '            Else
        '                If TextBox2.Text = "Ramuhashi" Then
        '                    MsgBox("welcome To my household system")
        '                Else

        '                End If
        '                DashBoard.TextBox1.Text = TextBox1.Text & " " & TextBox2.Text
        '                MsgBox("I don't recogonize this loggins")



        '                MessageBox.Show("Login Successful!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
        '            End If

        '            conn.Close()
        '            regis.ShowDialog()
        '            DashBoard.ShowDialog()
        '        End Using
        '    End Using

    End Sub
    Private Sub MY_USER_VALIDATION_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Set tooltips For buttons
        ToolTip1.SetToolTip(Button1, "Log in")
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        regis.ShowDialog()
    End Sub
End Class