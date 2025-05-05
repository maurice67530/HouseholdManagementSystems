Imports System.Data.OleDb
Public Class Family_Schedule
    Public Property conn As New OleDbConnection(connectionString)
    Public Const connectionString As String = " Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\Nedzamba\Source\Repos\maurice67530\HouseholdManagementSystems\HMS.accdb"

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        Try

            Dim schedule As New FamilySchedule() With {
    .Title = (TextBox1.Text),
    .Notes = TextBox2.Text,
    .DateOfEvent = DateTimePicker1.Text,
    .StartTime = DateTimePicker2.Text,
    .EndTime = DateTimePicker3.Text,
    .AssignedTo = ComboBox1.SelectedItem.ToString,
    .EventType = ComboBox2.SelectedItem.ToString
    }


            Using conn As New OleDbConnection(Ndamu.connectionstring)
                conn.Open()

                ' Update the table name if necessary  


                ' Create an OleDbCommand to insert the Expense data into the database  
                Dim cmd As New OleDbCommand("INSERT INTO FamilySchedule ([Title], [Notes], [DateOfEvent], [StartTime], [EndTime], [AssignedTo], [EventType]) VALUES (?, ?, ?, ?, ?, ?, ?)", conn)

                ' Set the parameter values from the UI controls  
                'For Each exp As Expenses In expenses
                cmd.Parameters.Clear()


                'cmd.Parameters.AddWithValue("@ID", Person.ID)
                cmd.Parameters.AddWithValue("@Title", TextBox1.Text)
                cmd.Parameters.AddWithValue("@Notes", TextBox2.Text)
                cmd.Parameters.AddWithValue("@DateOfEvent", DateTimePicker1.Text)
                cmd.Parameters.AddWithValue("@StartTime", DateTimePicker2.Text)
                cmd.Parameters.AddWithValue("@EndTime", DateTimePicker3.Text)
                cmd.Parameters.AddWithValue("@AssignedTo", ComboBox1.SelectedItem.ToString())
                cmd.Parameters.AddWithValue("@EventType", ComboBox2.SelectedItem.ToString())


                Dim rowsAffected As Integer = cmd.ExecuteNonQuery()
                'Display a confirmation messageBox  
                MsgBox("Schedule Information Added!" & vbCrLf &
                "Title: " & schedule.Title.ToString & vbCrLf &
                "Notes: " & schedule.Notes.ToString & vbCrLf &
                "DateOfEvent: " & schedule.DateOfEvent & vbCrLf &
                "StartTime: " & schedule.StartTime & vbCrLf &
                "EndTime: " & schedule.EndTime & vbCrLf &
                "AssignedTo: " & schedule.AssignedTo & vbCrLf &
                "EventType: " & schedule.EventType.ToString(), vbInformation, "Schedule confirmation")

            End Using

        Catch ex As Exception
            Debug.WriteLine($"Database error in btnAdd_Click: {ex.Message}")
            MessageBox.Show("Error saving Schedule to database: " & ex.Message & vbNewLine & ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Debug.WriteLine("Exiting btnAdd")

        End Try
        LoadScheduleFromDatabase()
    End Sub
    Public Sub LoadScheduleFromDatabase()
        Try
            Debug.WriteLine("DataGridview populated successfully ChoresForm_Load")
            Using conn As New OleDbConnection(Ndamu.connectionstring)
                conn.Open()

                ' Update the table name if necessary  
                Dim tableName As String = "FamilySchedule"


                ' Create an OleDbCommand to select the data from the database  
                Dim cmd As New OleDbCommand($"SELECT * FROM {tableName}", conn)

                ' Create a DataAdapter and fill a DataTable  
                Dim da As New OleDbDataAdapter(cmd)
                Dim dt As New DataTable()
                da.Fill(dt)

                ' Bind the DataTable to the DataGridView  
                DataGridView1.DataSource = dt
            End Using

        Catch ex As Exception
            Debug.WriteLine($"DataGridView population failed")
            Debug.WriteLine($"Unexpected error in DataGridView: {ex.Message}")
            Debug.WriteLine($"Error in PopulateDataGridView: {ex.Message}")
            Debug.WriteLine($"Stack Trace: {ex.StackTrace}")
            'MessageBox.Show("An error occurred while loading data into the grid.", "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub Family_Schedule_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim tooltip As New ToolTip
        tooltip.SetToolTip(btnSave, "Submit")
        tooltip.SetToolTip(btnUpdate, "Update")
        tooltip.SetToolTip(btnDelete, "Delete")

        PopulateComboboxFromDatabase(ComboBox1)
        LoadScheduleFromDatabase()
        AutoCreateChoreEvents()
        AutoAddMealTimes()
        AutoCreateTaskReminders()
        MarkPhotoDayEvents()
        'Timer1.Interval = 60000 ' 1 minute
        'Timer1.Start()
    End Sub
    Private Sub btnDelete_Click(sender As Object, e As EventArgs) Handles btnDelete.Click
        If DataGridView1.SelectedRows.Count > 0 Then
            ' Get the selected row  
            Dim selectedRow As DataGridViewRow = DataGridView1.SelectedRows(0)

            ' Assuming there is an "ID" column which is the primary key in the database  
            Dim ID As Integer = Convert.ToInt32(selectedRow.Cells("ID").Value) ' Replace "ID" with your actual column name  

            ' Confirm deletion  

            ' Proceed with deletion  
            Try
                If DataGridView1.SelectedRows.Count > 0 Then
                    Debug.WriteLine("A row is selected for deletion")
                    Dim confirmationResults As DialogResult = MessageBox.Show("Are you sure you want o delete this Schedule?", "Confirm Deletion", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)
                    If confirmationResults = DialogResult.Yes Then
                        Debug.WriteLine("User confirmed deletion.")

                    Else
                        Debug.WriteLine("User cancelled deletion.")
                    End If

                Else
                    Debug.WriteLine("No row selected, Exiting Button delete")

                    MessageBox.Show("Please Select chore to delete", "Deletion Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                End If

                Using conn As New OleDbConnection(Ndamu.connectionstring)
                    conn.Open()


                    ' Create the delete command  
                    Dim cmd As New OleDbCommand("DELETE FROM [FamilySchedule] WHERE [ID] = ?", conn)
                    cmd.Parameters.AddWithValue("@ID", ID) ' Primary key for matching record  

                    ' Execute the delete command  
                    Dim rowsAffected As Integer = cmd.ExecuteNonQuery()

                    If rowsAffected > 0 Then
                        MessageBox.Show("Task deleted successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        ' Optionally refresh DataGridView or reload from database  
                        LoadScheduleFromDatabase()
                    Else
                        MessageBox.Show("No chores deleted. Please check if the ID exists.", "Deletion Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    End If

                End Using

            Catch ex As Exception
                MessageBox.Show($"An error occurred while deleting the expense: {ex.Message}", "Deletion Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        Else

            MessageBox.Show("Please select chores to delete.", "Deletion Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)

            LoadScheduleFromDatabase()
        End If
    End Sub

    Private Sub btnUpdate_Click(sender As Object, e As EventArgs) Handles btnUpdate.Click

        'Ensure a row Is selected in the DataGridView  
        If DataGridView1.SelectedRows.Count = 0 Then
            MessageBox.Show("Please select a record to update.", "No Selection", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If


        Try
            Dim Title As String = TextBox1.Text
            Dim Notes As String = TextBox2.Text
            Dim DateOfEvent As DateTime = DateTimePicker1.Text
            Dim StartTime As String = DateTimePicker2.Text
            Dim EndTime As String = DateTimePicker3.Text
            Dim AssignedTo As String = ComboBox1.Text
            Dim EventType As String = ComboBox2.Text

            Using conn As New OleDbConnection(Ndamu.connectionstring)

                conn.Open()

                ' Get the ID of the selected row (assuming your table has a primary key named "ID")  
                Dim selectedRow As DataGridViewRow = DataGridView1.SelectedRows(0)
                Dim ID As Integer = Convert.ToInt32(selectedRow.Cells("ID").Value) ' Change "ID" to your primary key column name  

                ' Create an OleDbCommand to update the expense data in the database  
                Dim cmd As New OleDbCommand("UPDATE [FamilySchedule] SET [Title] = ?, [Notes] =?, [DateOfEvent] = ?, [StartTime] = ?, [EndTime] = ?, [AssignedTo] = ?, [EventType] = ? WHERE [ID] = ?", conn)

                ' Set the parameter values from the UI controls  
                cmd.Parameters.AddWithValue("@Title", TextBox1.Text)
                cmd.Parameters.AddWithValue("@Notes", TextBox2.Text)
                cmd.Parameters.AddWithValue("@DateOfEvent", DateTimePicker1.Text)
                cmd.Parameters.AddWithValue("@StartTime", DateTimePicker2.Text)
                cmd.Parameters.AddWithValue("@EndTime", DateTimePicker3.Text)
                cmd.Parameters.AddWithValue("@AssignedTo", ComboBox1.SelectedItem.ToString())
                cmd.Parameters.AddWithValue("@EventType", ComboBox2.SelectedItem.ToString())
                cmd.Parameters.AddWithValue("@ID", ID)
                cmd.ExecuteNonQuery()

                MsgBox("Schedule Updated Successfuly!", vbInformation, "Update Confirmation")

            End Using

        Catch ex As OleDbException
            MessageBox.Show("please ensure all fields are filled correctly. ", "input error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show("Error saving Chores to database: " & ex.Message & vbNewLine & ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Catch ex As Exception
            Debug.WriteLine($"General error in btnEdit_Click: {ex.Message}")
            Debug.WriteLine($"Stack Trace: {ex.StackTrace}")
            MessageBox.Show("An unexpected error occurred.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show("Error saving Chores to database: " & ex.Message & vbNewLine & ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try
        'TextBox4.Text = ("expense updated at{DateTime.Now:hh:mm:ss}")
        If DataGridView1.SelectedRows.Count > 0 Then
            Debug.WriteLine("A row is selected for update.")
        Else
            Debug.WriteLine("No row selected, exiting btnEdit_click.")
            MessageBox.Show("Please select task to update.", "update Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End If
        LoadScheduleFromDatabase()
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
                TextBox2.Text = selectedRow.Cells("Notes").Value.ToString()
                DateTimePicker1.Text = selectedRow.Cells("DateOfEvent").Value.ToString()
                DateTimePicker2.Text = selectedRow.Cells("StartTime").Value.ToString()
                DateTimePicker3.Text = selectedRow.Cells("EndTime").Value.ToString()
                ComboBox1.Text = selectedRow.Cells("AssignedTo").Value.ToString()
                ComboBox2.Text = selectedRow.Cells("EventType").Value.ToString()
            End If
        Catch ex As Exception
            Debug.WriteLine("error selection data in the database")
            Debug.WriteLine($"Stack Trace: {ex.StackTrace}")
            MessageBox.Show("Error saving inventory to database: " & ex.Message & vbNewLine & ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub

    Public Sub PopulateComboboxFromDatabase(ByRef comboBox As ComboBox)
        Dim conn As New OleDbConnection(Ndamu.connectionstring)
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
    Private Sub AutoCreateChoreEvents()
        Dim con As New OleDbConnection(Ndamu.connectionstring)
        Dim da As New OleDbDataAdapter("SELECT Tasks, DueDate FROM Tasks", con)
        Dim dt As New DataTable
        da.Fill(dt)

        Dim count As Integer = 0
        For Each row As DataRow In dt.Rows
            Dim cmd As New OleDbCommand("INSERT INTO FamilySchedule (Title, Notes, DateOfEvent, StartTime, EndTime, AssignedTo, EventType) VALUES (?, ?, ?, ?, ?, ?, ?)", conn)
            cmd.Parameters.AddWithValue("?", row("TaskName").ToString())
            cmd.Parameters.AddWithValue("?", "Auto-scheduled chore")
            cmd.Parameters.AddWithValue("?", CDate(row("NextDueDate")))
            cmd.Parameters.AddWithValue("?", #9:00:00 AM#)
            cmd.Parameters.AddWithValue("?", #10:00:00 AM#)
            cmd.Parameters.AddWithValue("?", "Family")
            cmd.Parameters.AddWithValue("?", "Chore")

            con.Open()
            cmd.ExecuteNonQuery()
            con.Close()
            count += 1
        Next

        MessageBox.Show(count.ToString() & " chore event(s) added to the schedule.", "Chore Integration Complete")
    End Sub


    Private Sub AutoAddMealTimes()
        Dim conn As New OleDbConnection(Ndamu.connectionstring)
        Dim da As New OleDbDataAdapter("SELECT MealName, StartDate FROM MealPlans", conn)
        Dim dt As New DataTable
        da.Fill(dt)

        Dim count As Integer = 0
        For Each row As DataRow In dt.Rows
            Dim cmd As New OleDbCommand("INSERT INTO FamilySchedule (Title, Notes, DateOfEvent, StartTime, EndTime, AssignedTo, EventType) VALUES (?, ?, ?, ?, ?, ?, ?)", conn)
            cmd.Parameters.AddWithValue("?", row("MealName").ToString())
            cmd.Parameters.AddWithValue("?", "Scheduled Meal")
            cmd.Parameters.AddWithValue("?", CDate(row("MealDate")))
            cmd.Parameters.AddWithValue("?", #1:00:00 PM#)
            cmd.Parameters.AddWithValue("?", #2:00:00 PM#)
            cmd.Parameters.AddWithValue("?", "Family")
            cmd.Parameters.AddWithValue("?", "Meal")

            conn.Open()
            cmd.ExecuteNonQuery()
            conn.Close()
            count += 1
        Next

        MessageBox.Show(count.ToString() & " meal(s) added to the family calendar.", "Meal Plan Integration Complete")
    End Sub


    Private Sub AutoCreateTaskReminders()
        Dim conn As New OleDbConnection(Ndamu.connectionstring)
        Dim da As New OleDbDataAdapter("SELECT Title, DueDate FROM Tasks WHERE Completed = False", conn)
        Dim dt As New DataTable
        da.Fill(dt)

        Dim count As Integer = 0
        For Each row As DataRow In dt.Rows
            Dim cmd As New OleDbCommand("INSERT INTO FamilySchedule (Title, Notes, DateOfEvent, StartTime, EndTime, AssignedTo, EventType) VALUES (?, ?, ?, ?, ?, ?, ?)", conn)
            cmd.Parameters.AddWithValue("?", row("TaskTitle").ToString())
            cmd.Parameters.AddWithValue("?", "Task due soon")
            cmd.Parameters.AddWithValue("?", CDate(row("DueDate")).AddDays(-1)) ' Reminder 1 day before
            cmd.Parameters.AddWithValue("?", #8:00:00 AM#)
            cmd.Parameters.AddWithValue("?", #8:30:00 AM#)
            cmd.Parameters.AddWithValue("?", "Family")
            cmd.Parameters.AddWithValue("?", "Task")

            conn.Open()
            cmd.ExecuteNonQuery()
            conn.Close()
            count += 1
        Next

        MessageBox.Show(count.ToString() & " task reminder(s) scheduled.", "Task Reminder Integration")
    End Sub


    Private Sub MarkPhotoDayEvents()
        Dim count As Integer = 0

        For Each row As DataGridViewRow In DataGridView1.Rows
            If Not row.IsNewRow AndAlso
           row.Cells("Notes").Value IsNot Nothing AndAlso
           Not IsDBNull(row.Cells("Notes").Value) AndAlso
           row.Cells("Notes").Value.ToString().ToLower().Contains("photo day") Then

                row.DefaultCellStyle.BackColor = Color.LightGoldenrodYellow
                count += 1
            End If
        Next

        If count > 0 Then
            MessageBox.Show(count.ToString() & " Photo Day event(s) highlighted!", "Photo Album Integration")
        Else
            MessageBox.Show("No Photo Day events found.", "Photo Album Integration")
        End If
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        'AutoCreateChoreEvents()
        'AutoAddMealTimes()
        'AutoCreateTaskReminders()
        'MarkPhotoDayEvents()
    End Sub
End Class


