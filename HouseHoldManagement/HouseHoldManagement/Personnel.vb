Imports System.IO
Imports System.Data.OleDb
Public Class Personnel
    ' Connection to the database
    Dim conn As New OleDbConnection(HouseHoldManagment_Module.connectionString)
    Dim connec As New OleDbConnection(Rasta.connectionString)

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
        ClearForm()
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
        ' Check if a member is selected
        If String.IsNullOrWhiteSpace(TextBox8.Text) Then
            MessageBox.Show("Please select a member from the table before updating.", "Selection Required", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If

        Try
            conn.Open()
            Dim query As String = "UPDATE PersonalDetails SET FirstName = ?, LastName = ?, DateOfBirth = ?, Contact = ?, Email = ?, Age = ? , Role = ?, Gender = ?, PostalCode = ?, MaritalStatus = ? WHERE ID = ?"

            Using cmd As New OleDbCommand(query, conn)
                cmd.Parameters.AddWithValue("@FirstName", TextBox1.Text)
                cmd.Parameters.AddWithValue("@LastName", TextBox2.Text)
                cmd.Parameters.AddWithValue("@DateOfBirth", DateTimePicker1.Value)
                cmd.Parameters.AddWithValue("@Contact", TextBox3.Text)
                cmd.Parameters.AddWithValue("@Email", TextBox4.Text)
                cmd.Parameters.AddWithValue("@Age", TextBox5.Text)
                cmd.Parameters.AddWithValue("@Role", ComboBox1.SelectedItem)
                cmd.Parameters.AddWithValue("@Gender", ComboBox3.SelectedItem)
                cmd.Parameters.AddWithValue("@PostalCode", TextBox6.Text)
                cmd.Parameters.AddWithValue("@MaritalStatus", ComboBox2.SelectedItem)
                cmd.Parameters.AddWithValue("@ID", CInt(TextBox8.Text))
                cmd.ExecuteNonQuery()
            End Using

            MessageBox.Show("Member updated successfully.")
            LoadData()
            ClearForm()

        Catch ex As Exception
            MessageBox.Show("Update failed: " & ex.Message)
        Finally
            conn.Close()
        End Try
    End Sub

    Private Sub BtnDelete_Click(sender As Object, e As EventArgs) Handles BtnDelete.Click
        ' Check if a member is selected
        If String.IsNullOrWhiteSpace(TextBox8.Text) Then
            MessageBox.Show("Please select a member from the table before deleting.", "Selection Required", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If

        ' Confirm before deletion
        Dim confirmDelete As DialogResult = MessageBox.Show("Are you sure you want to delete this member?", "Confirm Delete", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If confirmDelete = DialogResult.No Then
            Exit Sub
        End If

        Try
            conn.Open()
            Dim query As String = "DELETE FROM PersonalDetails WHERE ID = ?"

            Using cmd As New OleDbCommand(query, conn)
                cmd.Parameters.AddWithValue("@ID", CInt(TextBox8.Text))
                cmd.ExecuteNonQuery()
            End Using

            MessageBox.Show("Member deleted successfully.")
            LoadData()
            ClearForm()

        Catch ex As Exception
            MessageBox.Show("Delete failed: " & ex.Message)
        Finally
            conn.Close()
        End Try
    End Sub



    Private Sub BtnBack_Click(sender As Object, e As EventArgs) Handles BtnBack.Click
        Me.Visible = False
    End Sub

    Private Sub DataGridView1_SelectionChanged(sender As Object, e As EventArgs) Handles DataGridView1.SelectionChanged
        If DataGridView1.SelectedRows.Count > 0 Then
            Dim row As DataGridViewRow = DataGridView1.SelectedRows(0)
            If Not row.IsNewRow Then
                TextBox8.Text = row.Cells("ID").Value.ToString()
                TextBox1.Text = row.Cells("FirstName").Value.ToString()
                TextBox2.Text = row.Cells("LastName").Value.ToString()
                TextBox3.Text = row.Cells("Contact").Value.ToString()
                TextBox4.Text = row.Cells("Email").Value.ToString()
                TextBox5.Text = row.Cells("Age").Value.ToString()
                ComboBox2.SelectedItem = row.Cells("MaritalStatus").Value.ToString()
                ComboBox1.SelectedItem = row.Cells("Role").Value.ToString()
                ComboBox3.SelectedItem = row.Cells("Gender").Value.ToString()
                TextBox6.Text = row.Cells("PostalCode").Value.ToString()
                DateTimePicker1.Value = row.Cells("DateOfBirth").Value.ToString()
            End If
        End If
    End Sub

    Private Sub ClearForm()
        TextBox8.Clear()
        TextBox1.Clear()
        TextBox2.Clear()
        'DateTimePicker1.CLEAR
        TextBox4.Clear()
        TextBox3.Clear()
        TextBox5.Clear()
        'ComboBox1.CLEAR
        'ComboBox3.CLEAR
        TextBox6.Clear()
        'ComboBox2.CLEAR

    End Sub

    Private Sub Label13_Click(sender As Object, e As EventArgs) Handles Label13.Click

    End Sub

    Private Sub BtnAddpicture_Click(sender As Object, e As EventArgs) Handles BtnAddpicture.Click
        OpenFileDialog1.Filter = "Bitmaps (*.bmp)|*.bmp| (*.jpg)|*.jpg"
        If OpenFileDialog1.ShowDialog() = DialogResult.OK Then
            PictureBox1.Image = System.Drawing.Image.FromFile(OpenFileDialog1.FileName)
        End If
    End Sub

    Private Sub BtnDailyTasks_Click(sender As Object, e As EventArgs) Handles BtnDailyTasks.Click
        Chores.ShowDialog()
    End Sub

    Private Sub BtnClear_Click(sender As Object, e As EventArgs) Handles BtnClear.Click
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""
        TextBox8.Text = ""
        ComboBox1.Text = ""
        ComboBox2.Text = ""
        ComboBox3.Text = ""
    End Sub
End Class