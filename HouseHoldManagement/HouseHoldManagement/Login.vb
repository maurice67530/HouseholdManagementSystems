Imports System.IO
Imports System.Data.OleDb
Public Class Login
    Public Property conn As New OleDbConnection(Xiluva.connectionString)

    Public Const connectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\Xiluva\Source\Repos\maurice67530\HouseholdManagementSystems\HMS.accdb"

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Dim Username As String = TextBox1.Text

        Using conn As New OleDbConnection(Xiluva.connectionString)
            Try
                conn.Open()

                Dim cmd As New OleDbCommand("SELECT Role FROM Users WHERE Username = ? AND Password = ?", conn)

                cmd.Parameters.AddWithValue("?", TextBox1.Text)
                cmd.Parameters.AddWithValue("?", TextBox2.Text)

                Dim Role As Object = cmd.ExecuteScalar()

                If Role IsNot Nothing Then

                    MessageBox.Show("Login Successful!", "Welcome", MessageBoxButtons.OK, MessageBoxIcon.Information)

                    Select Case Role.ToString()

                        Case "Admin"

                            Form1.Show()

                        Case "Member"

                            Form1.Show()

                        Case "Finance"

                            Form1.Show()

                        Case "Chef"

                            Form1.Show()

                        Case Else
                            MessageBox.Show("Unknown role.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    End Select

                    Form1.TextBox1.Text = $"{Username}"

                    Form1.ShowDialog()
                    Me.Hide()
                Else
                    MessageBox.Show("Invalid Username or Password.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End If

            Catch ex As Exception

                MessageBox.Show("Error: " & ex.Message)
            Finally
                conn.Close()
            End Try

        End Using

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