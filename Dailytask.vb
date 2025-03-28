Imports System.Windows.Forms
Imports System.Data.OleDb
Imports System.Data

Public Class TaskManagement

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        DashBoard.Show()
        Me.Close()

    End Sub


    Public Sub LoadTaskManagementFromDatabase()
        Try
            Debug.WriteLine("populated connected succesfully")
            Using conn As New OleDbConnection(Module1.connectionString)
                conn.Open()

                Dim tableName As String = "TaskManagement"
                Dim cmd As New OleDbCommand($"SELECT * FROM {tableName}", conn)

                Dim da As New OleDbDataAdapter(cmd)
                Dim dt As New DataTable
                da.Fill(dt)

                DataGridView1.DataSource = dt

            End Using

        Catch ex As Exception

            Debug.WriteLine("populated fail to initialize succesfully")
            MessageBox.Show($"Error Loading  Expenses data from database: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub


    Private Sub DataGridView1_SelectionChanged(sender As Object, e As EventArgs) Handles DataGridView1.SelectionChanged
        'Load the data from  the selected row into UI controls 
        If DataGridView1.SelectedRows.Count > 0 Then
            Dim selectedRow As DataGridViewRow = DataGridView1.SelectedRows(0)
            TextBox1.Text = selectedRow.Cells("TaskID").Value.ToString()
            TextBox2.Text = selectedRow.Cells("Title").Value.ToString()
            TextBox3.Text = selectedRow.Cells("Description").Value.ToString()
            DateTimePicker1.Text = selectedRow.Cells("DueDate").Value.ToString()
            ComboBox1.Text = selectedRow.Cells("Priority").Value.ToString()
            TextBox4.Text = selectedRow.Cells("Status").Value.ToString()
            ComboBox2.Text = selectedRow.Cells("AssignedTo").Value.ToString()

        End If
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        Try

            Debug.WriteLine("Entering btnSubmit")
            If TextBox2.Text = "" Then
                MsgBox("please enter a value")
                TextBox1.Focus()
            End If

            If ComboBox2.Text = "" Then
                MsgBox("Please enter a value")
                ComboBox2.Focus()
            End If
            If DateTimePicker1.Text = "" Then
                MsgBox("Please enter a value")
                DateTimePicker1.Focus()
            End If
            If TextBox4.Text = "" Then
                MsgBox("Please enter a value ")
            End If

            Dim task As New TaskManagement With {
           .TaskId = TextBox1.Text,
            .Title = TextBox2.Text,
           .Description = TextBox3.Text,
           .DueDate = DateTimePicker1.Text,
            .priority = ComboBox1.Text,
           .Status = TextBox4.Text,
           .AssignedTo = ComboBox2.Text}


            Using connect As New OleDbConnection(Module1.connectionString)
                Module1.conn.Open()

                Dim tablename As String = "TaskManagement"

                Dim cmd As New OleDbCommand("INSERT INTO [TaskManagement] ([TaskID],[Title],[Description],[DueDate],[Priority],[Status],[AssignedTo]) VALUES (@TaskID ,@Title ,@Description, @DueDate, @Priority, @Status, @AssignedTo)", Module1.conn)
                'set the parameter values from the UI controls
                'class declaration  


                'params
                cmd.Parameters.AddWithValue("@TaskID", task.TaskId)
                cmd.Parameters.AddWithValue("@Title", task.Title)
                cmd.Parameters.AddWithValue("@Description", task.Description)
                cmd.Parameters.AddWithValue("@DueDate", task.DueDate)
                cmd.Parameters.AddWithValue("@Priority", task.priority)
                cmd.Parameters.AddWithValue("@Status", task.Status)
                cmd.Parameters.AddWithValue("@AssignedTo", task.AssignedTo)


                cmd.ExecuteNonQuery()

            End Using



            MsgBox("Task Management Added!" & vbCrLf &
                 "TaskID: " & task.TaskId & vbCrLf &
                   "Title: " & task.Title & vbCrLf &
                   "Description: " & task.Description & vbCrLf &
                   "DueDate: " & task.DueDate & vbCrLf &
                   "Priority: " & task.priority & vbCrLf &
                   "Status: " & task.Status & vbCrLf &
                   "AssignedTo: " & AssignedTo, vbInformation, "TaskManagement Confirmation")




        Catch ex As OleDbException
            Debug.WriteLine($"Database error: {ex.Message}")
            'MessageBox.Show("Error saving test to database. Please check the connection.", "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Catch ex As Exception
            Debug.WriteLine($"General error: {ex.Message}")
            'MessageBox.Show("An unexpected error occurred.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            conn.Close()
        Finally

        End Try
    End Sub

    Public Sub PopulateComboBoxFromDatabase(ByRef comboBox As ComboBox)


        Dim conn As New OleDbConnection(Module1.connectionString)
        Try
            Debug.WriteLine("Combobox populated from database")
            'open the database connection
            conn.Open()

            'retrieve the firstname and surname columns from the personaldetails tabel
            Dim query As String = "SELECT firstname, lastName FROM personaldetails"
            Dim cmd As New OleDbCommand(query, conn)
            Dim reader As OleDbDataReader = cmd.ExecuteReader()

            'bind the retrieved data to the combobox 
            ComboBox2.Items.Clear()
            While reader.Read()
                ComboBox2.Items.Add($"{reader("firstname")}{reader("lastname")}")
            End While

            'close the database 
            reader.Close()

        Catch ex As Exception
            Debug.WriteLine("failed to populate combobox")
            Debug.WriteLine($"Stack Trace:{ex.StackTrace}")
            'handle any exeptions that may occur
            MessageBox.Show($"Error:{ex.Message}")


        Finally
            'close the database connection
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
        Debug.WriteLine("Done With populating combobox  from database")
    End Sub

    Private Sub Dailytask_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        LoadTaskManagementFromDatabase()
        'PopulateComboBoxFromDatabase()

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If DataGridView1.SelectedRows.Count = 0 Then
            MessageBox.Show("Please select a record to update.", "No Selection", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        Try
            Debug.WriteLine("Entering btnEdit")
            Dim TaskID As Integer = TextBox1.Text
            Dim Title As String = TextBox2.Text
            Dim Description As String = TextBox3.Text
            Dim DueDate As DateTime = DateTimePicker1.Text
            Dim Priority As Integer = ComboBox1.Text
            Dim Status As Integer = TextBox4.Text
            Dim AssignedTo As String = ComboBox2.Text




            Using conn As New OleDbConnection(Module1.connectionString)
                conn.Open()

                ' Get the ID of the selected row (assuming your table has a primary key named "ID")  
                Dim selectedRow As DataGridViewRow = DataGridView1.SelectedRows(0)
                Dim ID As Integer = Convert.ToInt32(selectedRow.Cells("ID").Value) ' Change "ID" to your primary key column name  

                ' Create an OleDbCommand to update the expense data in the database  

                Dim cmd As New OleDbCommand("UPDATE TaskManagement SET [TaskID] = ?, [Title]  = ?, [Description] = ?, [DueDate] = ?, [Priority] = ?, [Status] = ?,[AssignedTo]=?WHERE [ID] = ?", conn)
                'Set the parameter values from the UI controls 
                cmd.Parameters.AddWithValue("@TaskID", TextBox1.Text)
                cmd.Parameters.AddWithValue("@Title", TextBox2.Text)
                cmd.Parameters.AddWithValue("@Description", TextBox3.Text)
                cmd.Parameters.AddWithValue("@DueDate", DateTimePicker1.Text)
                cmd.Parameters.AddWithValue("@Priority", ComboBox1.Text)
                cmd.Parameters.AddWithValue("@Status", TextBox4.Text)
                cmd.Parameters.AddWithValue("@AssignedTo", ComboBox2.Text)


                'cmd.Parameters.AddWithValue("@ID", TaskManagementID) ' Primary key for matching record  
                'cmd.ExecuteNonQuery()


                MsgBox("Task Management Updated Successfuly!", vbInformation, "Update Confirmation")
                LoadTaskManagementFromDatabase()

            End Using
        Catch ex As OleDbException
            MessageBox.Show($"Error updating TaskManagement in database: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Catch ex As Exception
            MessageBox.Show($"Unexpected error: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        conn.Close()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click

        ' Check if there are any selected rows in the DataGridView for expenses  
        If DataGridView1.SelectedRows.Count > 0 Then
            ' Get the selected row  
            Dim selectedRow As DataGridViewRow = DataGridView1.SelectedRows(0)

            ' Assuming there is an "ID" column which is the primary key in the database  
            Dim TaskManagementId As Integer = Convert.ToInt32(selectedRow.Cells("ID").Value) ' Replace "ID" with your actual column name  

            ' Confirm deletion  
            Dim confirmationResult As DialogResult = MessageBox.Show("Are you sure you want to delete this expense?", "Confirm Deletion", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)
            If confirmationResult = DialogResult.Yes Then
                ' Proceed with deletion  
                Try
                    Using conn As New OleDbConnection(Module1.connectionString)
                        conn.Open()

                        ' Create the delete command  
                        Dim cmd As New OleDbCommand("DELETE FROM [TaskManagement] WHERE [ID] = ?", conn)
                        cmd.Parameters.AddWithValue("@ID", TaskManagementId) ' Primary key for matching record  

                        ' Execute the delete command  
                        Dim rowsAffected As Integer = cmd.ExecuteNonQuery()

                        If rowsAffected > 0 Then
                            MessageBox.Show("TaskManagement deleted successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
                            ' Optionally refresh DataGridView or reload from database  
                            'PopulateDataGridView()
                        Else
                            MessageBox.Show("No expense was deleted. Please check if the ID exists.", "Deletion Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                        End If
                    End Using

                Catch ex As Exception
                    MessageBox.Show($"An error occurred while deleting the TaskManagement: {ex.Message}", "Deletion Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
            End If
        Else
            MessageBox.Show("Please select an TaskManagement to delete.", "Deletion Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            conn.Close()
        End If
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles BtnRefresh.Click
        LoadTaskManagementFromDatabase()

    End Sub

    Private Sub Button5_Click_1(sender As Object, e As EventArgs) Handles Button5.Click

    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        DataGridView1.Sort(DataGridView1.Columns(0), System.ComponentModel.ListSortDirection.Ascending)
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click

    End Sub
End Class
