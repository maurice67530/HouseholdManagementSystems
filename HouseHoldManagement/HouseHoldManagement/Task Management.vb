Imports System.Data.OleDb
Imports HouseHoldManagement
Public Class Task_Management
    Public Property conn As New OleDbConnection(Ndivhuwo.connectionString)
    Public Const connectionString As String = " Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\Delicious\Source\Repos\maurice67530\HouseholdManagementSystems\HMS.accdb"
    'Public Property conn As New OleDbConnection(Rinae.connectionString)
    Private Status As String
    Private Tasks As Object
    Private Dashboard As Object
    Private query As String
    Public Property Priority As String
    Public Sub LoadTasksDataFromDatabase()

        Debug.WriteLine("LoadTasksDataFromDatabase()")
        'Using conn As New OleDbConnection(Rinae.connectionString)
        Using conn As New OleDbConnection(Ndivhuwo.connectionString)
                conn.Open()

                ' Update the table name if necessary  
                Dim tableName As String = "Tasks"

                ' Create an OleDbCommand to select the data from the database  
                Dim cmd As New OleDbCommand($"SELECT * FROM {tableName}", conn)

                ' Create a DataAdapter and fill a DataTable  
                Dim da As New OleDbDataAdapter(cmd)
                Dim dt As New DataTable()
                da.Fill(dt)

                ' Bind the DataTable to the DataGridView  
                DataGridView1.DataSource = dt
                'HighlightExpiredItemss()
            End Using
    End Sub
    Private Sub Task_Management_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        LoadTasksDataFromDatabase()
        Dim tooltip As New ToolTip
        tooltip.SetToolTip(Button1, "Submit")
        tooltip.SetToolTip(Button4, "Refresh")
        tooltip.SetToolTip(Button3, "Delete")
        tooltip.SetToolTip(Button2, "Edit")
        tooltip.SetToolTip(Button8, "Dashboard")
        tooltip.SetToolTip(Button6, "Sort")

        ComboBox1.Items.AddRange(New String() {"Low", "Medium", "High"})
        ComboBox2.Items.AddRange(New String() {"Not started", "In progress", "Completed"})
        PopulateComboboxFromDatabase(ComboBox3)
    End Sub
    Private Sub ComboBox3_click(sender As Object, e As EventArgs) Handles ComboBox3.Click
        PopulateComboboxFromDatabase(ComboBox3)
    End Sub
    Public Sub PopulateComboboxFromDatabase(ByRef comboBox As ComboBox)
        Dim conn As New OleDbConnection(Ndivhuwo.connectionString)
        'Dim conn As New OleDbConnection(Rinae.connectionString)
        Try
            Debug.WriteLine("populating combobox from database successfully!")
            ' 1. Open the database connection  
            conn.Open()

            ' 2. Retrieve the FirstName and LastName columns from the PersonalDetails table  
            Dim query As String = "SELECT FirstName, LastName FROM PersonalDetails"
            Dim cmd As New OleDbCommand(query, conn)
            Dim reader As OleDbDataReader = cmd.ExecuteReader()

            ' 3. Bind the retrieved data to the combobox  
            comboBox.Items.Clear()
            While reader.Read()
                comboBox.Items.Add($"{reader("FirstName")} {reader("LastName")}")
            End While

            ' 4. Close the database connection  
            reader.Close()

        Catch ex As Exception
            Debug.WriteLine($"form loaded unsuccessful")
            Debug.WriteLine($"Stack Trace: {ex.StackTrace}")
        Finally
            ' Close the database connection  
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        ' Check if there are any selected rows in the DataGridView for personal  
        If DataGridView1.SelectedRows.Count > 0 Then
            ' Get the selected row  
            Dim selectedRow As DataGridViewRow = DataGridView1.SelectedRows(0)

            ' Assuming there is an "ID" column which is the primary key in the database  
            Dim ID As Integer = Convert.ToInt32(selectedRow.Cells("ID").Value) ' Replace "ID" with your actual column name  

            ' Confirm deletion  
            Dim confirmationResult As DialogResult = MessageBox.Show("Are you sure you want to delete this e?", "Confirm Deletion", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)
            If confirmationResult = DialogResult.Yes Then
                ' Proceed with deletion  
                Try
                    Using conn As New OleDbConnection(Ndivhuwo.connectionString)
                        conn.Open()

                        ' Create the delete command  
                        Dim cmd As New OleDbCommand("DELETE FROM [Tasks] WHERE [ID] = ?", conn)
                        cmd.Parameters.AddWithValue("@ID", ID) ' Primary key for matching record  

                        ' Execute the delete command  
                        Dim rowsAffected As Integer = cmd.ExecuteNonQuery()

                        If rowsAffected > 0 Then
                            MessageBox.Show("task deleted successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
                            ' Optionally refresh DataGridView or reload from database  
                            'PopulateDataGridView1()
                        Else
                            MessageBox.Show("No Task was deleted. Please check if the ID exists.", "Deletion Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                        End If
                    End Using

                Catch ex As Exception
                    MessageBox.Show($"An error occurred while deleting the Task:  {ex.Message}", "Deletion Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
            End If
        Else
            MessageBox.Show("Please select an Task to delete.", "Deletion Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End If
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try
            Debug.WriteLine("entering btnsave")

            If TextBox2.Text = " " Then
                MsgBox("please enter a value")
                TextBox2.Focus()

            End If
            If ComboBox1.Text = "" Then
                MsgBox("please enter a value")
                ComboBox1.Focus()

            End If
            If DateTimePicker1.Text = "" Then
                MsgBox("please enter a value")
                DateTimePicker1.Focus()

            End If
            Using conn As New OleDbConnection(Ndivhuwo.connectionString)
                conn.Open()

                Dim tableName As String = "Tasks"
                Dim cmd As New OleDbCommand("INSERT INTO Tasks ([Title],[Description],[DueDate],[Priority],[Status],[AssignedTo]) VALUES (@Title, @Description, @DueDate, @Priority, @Status, @AssignedTo)", conn)

                Dim Task As New DailyTask() With {
             .Title = TextBox1.Text,
              .Description = TextBox2.Text,
            .DueDate = DateTimePicker1.Text,
            .Priority = ComboBox1.Text,
            .Status = ComboBox2.Text,
            .AssignedTo = ComboBox3.Text}


                'parameters

                cmd.Parameters.AddWithValue("@Title", Task.Title)
                cmd.Parameters.AddWithValue("@Description", Task.Description)
                cmd.Parameters.AddWithValue("@DueDate", Task.DueDate)
                cmd.Parameters.AddWithValue("@Priority", Task.Priority)
                cmd.Parameters.AddWithValue("@Status", Task.Status)
                cmd.Parameters.AddWithValue("@AssignedTo", Task.AssignedTo)

                cmd.ExecuteNonQuery()

                MsgBox("tasks Added!" & vbCrLf &
                   "Title: " & Task.Title.ToString() & vbCrLf &
                   "Description: " & Task.Description & vbCrLf &
                   "DueDate: " & Task.DueDate & vbCrLf &
                   "Priority: " & Task.Priority & vbCrLf &
                   "Status: " & Task.Status & vbCrLf &
                   "Assignedto: " & Task.AssignedTo, vbInformation, "Tasks Confirmation")


            End Using
        Catch ex As OleDbException
            Debug.WriteLine($"database error: {ex.Message}")
            MessageBox.Show("error saving test to database. please check the connection.", "Database error", MessageBoxButtons.OK, MessageBoxIcon.Error)

        Catch ex As Exception
            Debug.WriteLine($"general error in btnsave:{ex.Message}")
            MessageBox.Show("An unexpected error occurred.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try
        ''Auto generate reminder in calender
        'Ndamu.AddTaskReminder(TextBox1.Text, ComboBox3.Text, DateTimePicker1.Text)
    End Sub
    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        DataGridView1.Sort(DataGridView1.Columns("DueDate"), System.ComponentModel.ListSortDirection.Ascending)
    End Sub
    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        '    Dashboard.Show()
        Me.Close()
    End Sub
    Public Sub LoadTaskDataFromDatabase()

        Debug.WriteLine(" Task load has been initialised!")
        Using conn As New OleDbConnection(Ndivhuwo.connectionString)
            conn.Open()

            Dim TableName As String = "Tasks"

            Dim cmd As New OleDbCommand($"SELECT*FROM {TableName}", conn)

            Dim da As New OleDbDataAdapter(cmd)
            Dim dt As New DataTable
            da.Fill(dt)

            DataGridView1.DataSource = dt

        End Using
    End Sub
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        LoadTaskDataFromDatabase()
        ClearControls(Me)
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        'Ensure a row Is selected in the DataGridView  
        If DataGridView1.SelectedRows.Count = 0 Then
            MessageBox.Show("Please select a record to update.", "No Selection", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If


        Try
            Dim Title As String = TextBox1.Text
            Dim Description As String = TextBox2.Text
            Dim DueDate As DateTime = DateTimePicker1.Text
            Dim Priority As String = ComboBox1.Text
            Dim Status As String = ComboBox2.Text
            Dim AssignedTo As String = ComboBox3.Text

            Using conn As New OleDbConnection(Ndivhuwo.connectionString)

                conn.Open()

                ' Get the ID of the selected row (assuming your table has a primary key named "ID")  
                Dim selectedRow As DataGridViewRow = DataGridView1.SelectedRows(0)
                Dim ID As Integer = Convert.ToInt32(selectedRow.Cells("ID").Value) ' Change "ID" to your primary key column name  

                ' Create an OleDbCommand to update the expense data in the database  
                Dim cmd As New OleDbCommand("UPDATE [Tasks] SET [Title] = ?, [Description] =?, [DueDate] = ?, [Priority] = ?, [Status] = ?, [AssignedTo] = ? WHERE [ID] = ?", conn)

                ' Set the parameter values from the UI controls  
                cmd.Parameters.AddWithValue("@Title", TextBox1.Text)
                cmd.Parameters.AddWithValue("@Description", TextBox2.Text)
                cmd.Parameters.AddWithValue("@DueDate", DateTimePicker1.Text)
                cmd.Parameters.AddWithValue("@Priority", ComboBox1.SelectedItem.ToString())
                cmd.Parameters.AddWithValue("@Status", ComboBox2.SelectedItem.ToString())
                cmd.Parameters.AddWithValue("@AssignedTo", ComboBox3.SelectedItem.ToString())
                cmd.Parameters.AddWithValue("@ID", ID)
                cmd.ExecuteNonQuery()

                MsgBox("Task Updated Successfuly!", vbInformation, "Update Confirmation")

            End Using

        Catch ex As OleDbException
            MessageBox.Show("please ensure all fields are filled correctly. ", "input error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Catch ex As Exception
            Debug.WriteLine($"General error in btnEdit_Click: {ex.Message}")
            Debug.WriteLine($"Stack Trace: {ex.StackTrace}")
            MessageBox.Show("An unexpected error occurred.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        'TextBox4.Text = ("expense updated at{DateTime.Now:hh:mm:ss}")
        If DataGridView1.SelectedRows.Count > 0 Then
            Debug.WriteLine("A row is selected for update.")
        Else
            Debug.WriteLine("No row selected, exiting btnEdit_click.")
            MessageBox.Show("Please select task to update.", "update Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End If
    End Sub
    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub
    Private Sub DataGridView1_SelectionChanged(sender As Object, e As EventArgs) Handles DataGridView1.SelectionChanged
        Try

            Debug.WriteLine("selecting data in the datagridview")
            If DataGridView1.SelectedRows.Count > 0 Then
                Dim selectedRow As DataGridViewRow = DataGridView1.SelectedRows(0)

                ' Load the data from the selected row into UI controls  
                TextBox1.Text = selectedRow.Cells("Title").Value.ToString()
                TextBox2.Text = selectedRow.Cells("Description").Value.ToString()
                DateTimePicker1.Text = selectedRow.Cells("DueDate").Value.ToString()
                ComboBox1.Text = selectedRow.Cells("Priority").Value.ToString()
                ComboBox2.Text = selectedRow.Cells("Status").Value.ToString()
                ComboBox3.Text = selectedRow.Cells("AssignedTo").Value.ToString()

            End If
        Catch ex As Exception
            Debug.WriteLine("error selection data in the database")
            Debug.WriteLine($"Stack Trace: {ex.StackTrace}")
            MessageBox.Show("Error saving inventory to database: " & ex.Message & vbNewLine & ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub
    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged

    End Sub


End Class