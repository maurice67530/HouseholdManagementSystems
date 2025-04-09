Imports System.IO
Imports System.Data.OleDb
Public Class Personnel
    Private conn As New OleDbConnection(HouseHoldManagment_Module.connectionString)
    ' Create a ToolTip object
    Private toolTip As New ToolTip()
    Private toolTip1 As New ToolTip()
    Private Sub Personnel_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        LoadData()
        ' Initialize ToolTip properties (optional)
        toolTip.AutoPopDelay = 5000
        toolTip.InitialDelay = 500
        toolTip.ReshowDelay = 200
        toolTip.ShowAlways = True

        toolTip1.SetToolTip(Button3, "Back")
        toolTip1.SetToolTip(Button4, "Add a Picture")
        toolTip1.SetToolTip(Button2, "Edit")
        toolTip1.SetToolTip(Button5, "Delete")
        toolTip1.SetToolTip(Button6, "Clear")
        toolTip1.SetToolTip(Button7, "Daily tasks")
        toolTip1.SetToolTip(Button1, "Save")
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        ' Connection to the database
        Dim conn As New OleDbConnection(HouseHoldManagment_Module.connectionString)

        ' Variables to hold user inputs
        Dim FirstName As String
        Dim LastName As String
        Dim DateOfBirth As Date
        Dim Email As String
        Dim Contact As String
        Dim Age As String
        Dim Role As String
        Dim Gender As String
        Dim PostalCode As String
        Dim HealthStatus As String
        Dim Deleter As String


        ' Get user input from TextBoxes
        FirstName = TextBox1.Text
        LastName = TextBox2.Text
        DateOfBirth = DateTimePicker1.Value
        Email = TextBox4.Text
        Contact = TextBox3.Text
        Age = TextBox5.Text
        Role = ComboBox1.Text
        Gender = ComboBox3.Text
        PostalCode = TextBox6.Text
        HealthStatus = TextBox9.Text
        Deleter = TextBox7.Text
        ' Open the connection
        Try
                conn.Open()

            ' SQL query to insert the data
            Dim query As String = "INSERT INTO PersonalDetails (FirstName, LastName, DateOfBirth, Email, Contact, Age, Role, Gender, PostalCode, HealthStatus, Deleter) " &
                                  "VALUES (@FirstName, @LastName, @DateOfBirth, @Email, @Contact, @Age, @Role, @Gender, @PostalCode, @HealthStatus, @Deleter)"

            ' Create the command and add parameters
            Dim cmd As New OleDbCommand(query, conn)
            cmd.Parameters.AddWithValue("@FirstName", FirstName)
            cmd.Parameters.AddWithValue("@LastName", LastName)
            cmd.Parameters.AddWithValue("@DateOfBirth", DateOfBirth)
            cmd.Parameters.AddWithValue("@Email", Email)
            cmd.Parameters.AddWithValue("@Contact", Contact)
            cmd.Parameters.AddWithValue("@Age", Age)
            cmd.Parameters.AddWithValue("@Role", Role)
            cmd.Parameters.AddWithValue("@Gender", Gender)
            cmd.Parameters.AddWithValue("@PostalCode", PostalCode)
            cmd.Parameters.AddWithValue("@HealthStatus", HealthStatus)
            cmd.Parameters.AddWithValue("@Deleter", Deleter)
            ' Execute the query
            Dim rowsAffected As Integer = cmd.ExecuteNonQuery()

                ' Show confirmation message
                If rowsAffected > 0 Then
                    MessageBox.Show(rowsAffected.ToString() & " record inserted successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Else
                    MessageBox.Show("No records were inserted.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End If
            Catch ex As Exception
                ' Handle any errors
                MessageBox.Show("An error occurred: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Finally
                ' Ensure the connection is closed even if an error occurs
                If conn.State = ConnectionState.Open Then
                    conn.Close()
                End If
            End Try
        End Sub

    ' Method to load data into the DataGridView
    Private Sub LoadData()
        Try
            conn.Open()
            Dim query As String = "SELECT * FROM PersonalDetails"
            Dim cmd As New OleDbCommand(query, conn)
            Dim adapter As New OleDbDataAdapter(cmd)
            Dim dt As New DataTable()
            adapter.Fill(dt)
            DataGridView1.DataSource = dt
        Catch ex As Exception
            MessageBox.Show("Failed to load data: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            conn.Close()
        End Try
    End Sub


End Class