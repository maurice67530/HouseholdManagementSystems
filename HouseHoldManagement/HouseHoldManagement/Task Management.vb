Imports System.Data.OleDb

Public Class Task_Management
    Public Const connectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\khodani\Source\Repos\maurice67530\HouseholdManagementSystems\HMS.accdb"
    Dim conn As New OleDbConnection(HouseHoldManagment_Module.connectionString)
    Private Sub Task_Management_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        LoadTaskDataFromDatabase()
        Dim tooltip As New ToolTip
        tooltip.SetToolTip(Button1, "Submit")
        tooltip.SetToolTip(Button4, "Refresh")
        tooltip.SetToolTip(Button3, "Delete")
        tooltip.SetToolTip(Button2, "Edit")
        tooltip.SetToolTip(Button8, "Dashboard")
        tooltip.SetToolTip(Button5, "Filter")
        tooltip.SetToolTip(Button6, "Sort")

        ComboBox1.Items.AddRange(New String() {"Low", "Medium", "High"})
        ComboBox2.Items.AddRange(New String() {"Not started", "In progress", "Completed"})
    End Sub
    Private Sub ComboBox3_click(sender As Object, e As EventArgs) Handles ComboBox3.Click
        PopulateComboboxFromDatabase(ComboBox3)
    End Sub
    Public Sub PopulateComboboxFromDatabase(ByRef comboBox As ComboBox)
        Dim conn As New OleDbConnection(HouseHoldManagment_Module.connectionString)
        Try
            Debug.WriteLine("populating combobox from database successfully!")
            ' 1. Open the database connection  
            conn.Open()

            ' 2. Retrieve the FirstName and LastName columns from the PersonalDetails table  
            Dim query As String = "SELECT FirstName, LastName FROM Person"
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
                    Using conn As New OleDbConnection(HouseHoldManagment_Module.connectionString)
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
            Using conn As New OleDbConnection(connectionString)
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
                   "Assignedto: " & Task.Assignedto, vbInformation, "inventory Confirmation")


            End Using
        Catch ex As OleDbException
            Debug.WriteLine($"database error: {ex.Message}")
            MessageBox.Show("error saving test to database. please check the connection.", "Database error", MessageBoxButtons.OK, MessageBoxIcon.Error)

        Catch ex As Exception
            Debug.WriteLine($"general error in btnsave:{ex.Message}")
            MessageBox.Show("An unexpected error occurred.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub
    Public Sub LoadTaskDataFromDatabase()

        Debug.WriteLine("LoadTaskDataFromDatabase")
        Using connect As New OleDbConnection(connectionString)
            connect.Open()

            ' Update the table name if necessary  
            Dim tableName As String = "Tasks"

            ' Create an OleDbCommand to select the data from the database  
            Dim cmd As New OleDbCommand($"SELECT * FROM {tableName}", connect)

            ' Create a DataAdapter and fill a DataTable  
            Dim da As New OleDbDataAdapter(cmd)
            Dim dt As New DataTable()
            da.Fill(dt)

            ' Bind the DataTable to the DataGridView  
            DataGridView1.DataSource = dt
            'HighlightExpiredItemss()
        End Using

    End Sub
End Class