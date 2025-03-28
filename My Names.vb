Imports System.IO
Imports System.Windows.Forms
Imports System.Data.OleDb
Imports System.Drawing
Public Class MY_NAME_AND_SURNAME



    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Dim Names As String
        Dim surname As String
        Names = TextBox1.Text
        surname = TextBox2.Text

        Dim conn As New OleDbConnection(Module1.connectionString)

        conn.Open()

        Dim tablename As String = "MyNames"


        Dim cmd As New OleDbCommand("INSERT INTO MyNames([Names],[Surname]) VALUES (@Names,@Surname)", conn)
        'parameters
        cmd.Parameters.AddWithValue("@Names", Names)
        cmd.Parameters.AddWithValue("Surname", surname)

        Dim RowsAffected As Integer = cmd.ExecuteNonQuery()
        MessageBox.Show(RowsAffected.ToString() & "Inserted succesfully")


    End Sub

    Private Sub MY_NAME_AND_SURNAME_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class

