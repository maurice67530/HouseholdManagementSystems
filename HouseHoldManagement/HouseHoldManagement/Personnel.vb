Imports System.IO
Imports System.Data.OleDb
Public Class Personnel

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
    Dim MaritalStatus As String


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

        toolTip1.SetToolTip(BtnBack, "Back")
        toolTip1.SetToolTip(BtnAddpicture, "Add a Picture")
        toolTip1.SetToolTip(BtnEdit, "Edit")
        toolTip1.SetToolTip(BtnDelete, "Delete")
        toolTip1.SetToolTip(BtnClear, "Clear")
        toolTip1.SetToolTip(BtnDailyTasks, "Daily tasks")
        toolTip1.SetToolTip(BtnSave, "Save")
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles BtnSave.Click
        ' Get user input from TextBoxes
        FirstName = TextBox1.Text
        LastName = TextBox2.Text
        DateOfBirth = DateTimePicker1.Value
        Email = TextBox4.Text
        Contact = TextBox3.Text
        Age = TextBox5.Text
        Role = ComboBox1.SelectedItem.ToString
        PostalCode = TextBox6.Text
        Gender = ComboBox3.SelectedItem.ToString
        MaritalStatus = ComboBox2.SelectedItem.ToString



        ' Open the connection
        Try
            conn.Open()

            ' SQL query to insert the data
            Dim query As String = "INSERT INTO PersonalDetails (FirstName, LastName, DateOfBirth, Email, Contact, Age, Role, Gender, PostalCode, MaritalStatus) " &
                                  "VALUES (@FirstName, @LastName, @DateOfBirth, @Email, @Contact, @Age, @Role, @Gender, @PostalCode, @MaritalStatus)"

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
            cmd.Parameters.AddWithValue("@MaritalStatus", MaritalStatus)

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
        LoadData()
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

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        ' When a row is clicked in the DataGridView

        ' Check if a valid row is clicked (not header)
        If e.RowIndex >= 0 Then
            Dim row As DataGridViewRow = DataGridView1.Rows(e.RowIndex)

            ' Fill the form fields with the values from the selected row
            TextBox1.Text = row.Cells("FirstName").Value.ToString()
            TextBox2.Text = row.Cells("LastName").Value.ToString()
            DateTimePicker1.Value = Convert.ToDateTime(row.Cells("DateOfBirth").Value)
            TextBox4.Text = row.Cells("Email").Value.ToString()
            TextBox3.Text = row.Cells("Contact").Value.ToString()
            TextBox5.Text = row.Cells("Age").Value.ToString()
            ComboBox1.Text = row.Cells("Role").Value.ToString()
            ComboBox3.Text = row.Cells("Gender").Value.ToString()
            TextBox6.Text = row.Cells("PostalCode").Value.ToString()
            ComboBox2.SelectedItem = row.Cells("MaritalStatus").Value.ToString()
            'TextBox7.Text = row.Cells("Deleter").Value.ToString()
        End If

    End Sub

    Private Sub BtnEdit_Click(sender As Object, e As EventArgs) Handles BtnEdit.Click
        Try
            conn.Open()
            Dim query As String = "UPDATE PersonalDetails SET FirstName = ?, LastName = ?, DateOfBirth = ?, Email = ?, Contact = ?, Age = ?, Role = ?, Gender = ?, PostalCode = ?, MaritalStatus = ? WHERE ID = ?"
            Using cmd As New OleDbCommand(query, conn)
                cmd.Parameters.AddWithValue("@FirstName", TextBox1.Text)
                cmd.Parameters.AddWithValue("@LastName", TextBox2.Text)
                cmd.Parameters.AddWithValue("@DateOfBirth", DateTimePicker1.Value.ToString)
                cmd.Parameters.AddWithValue("@Email", TextBox4.Text)
                cmd.Parameters.AddWithValue("@Contact", TextBox3.Text)
                cmd.Parameters.AddWithValue("@Age", TextBox5.Text)
                cmd.Parameters.AddWithValue("@Role", ComboBox1.SelectedItem.ToString)
                cmd.Parameters.AddWithValue("@Gender", ComboBox3.SelectedItem.ToString)
                cmd.Parameters.AddWithValue("@PostalCode", TextBox6.Text)
                cmd.Parameters.AddWithValue("@MaritalStatus", ComboBox2.SelectedItem.ToString)
                'cmd.Parameters.AddWithValue("@Deleter", TextBox7.Text)
                cmd.Parameters.AddWithValue("@ID", CInt(TextBox8.Text))
                cmd.ExecuteNonQuery()
            End Using
            MessageBox.Show("Member updated successfully.")
            LoadData()

        Catch ex As Exception
            MessageBox.Show("Update failed: " & ex.Message)
        Finally
            conn.Close()
        End Try
    End Sub

    Private Sub BtnDelete_Click(sender As Object, e As EventArgs) Handles BtnDelete.Click
        Try
            conn.Open()
            Dim query As String = "DELETE FROM PersonalDetails WHERE ID = ?"
            Using cmd As New OleDbCommand(query, conn)
                cmd.Parameters.AddWithValue("@ID", CInt(TextBox8.Text))
                cmd.ExecuteNonQuery()
            End Using
            MessageBox.Show("Member deleted successfully.")
            LoadData()

        Catch ex As Exception
            MessageBox.Show("Delete failed: " & ex.Message)
        Finally
            conn.Close()
        End Try

    End Sub

    Private Sub BtnBack_Click(sender As Object, e As EventArgs) Handles BtnBack.Click

    End Sub
End Class