Imports System.Data.OleDb

Public Class Family_Schedule
    Public Const connectionstring As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\Nedzamba\Source\Repos\HouseholdManagementSystems\HMS.accdb"
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        Try

            Dim schedule As New FamilySchedule() With {
        .Title = (TextBox1.Text),
        .Notes = TextBox2.Text,
        .DateOfEvent = DateTimePicker1.Text,
        .StartTime = DateTimePicker2.Text,
        .EndTime = DateTimePicker3.Text,
        .AssignedTo = ComboBox1.SelectedItem.ToString,
        .EventType = ComboBox1.SelectedItem.ToString
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


        Catch ex As OleDbException
            Debug.WriteLine($"Database error in btnSAdd_Click: {ex.Message}")
            Debug.WriteLine($"Stack Trace: {ex.StackTrace}")
            MessageBox.Show("Error saving Chores to database. Please check the connection.", "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Catch ex As Exception

            Debug.WriteLine($"General error in btnAdd_Click: {ex.Message}")
            Debug.WriteLine($"Stack Trace: {ex.StackTrace}")
            MessageBox.Show("An unexpected error occurred.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        Debug.WriteLine("Exiting btnAdd")
        LoadScheduleFromDatabase()
    End Sub
    Public Sub LoadScheduleFromDatabase()
        Try
            Debug.WriteLine("DataGridview populated successfully ChoresForm_Load")
            Using conn As New OleDbConnection(Ndamu.connectionstring)
                conn.Open()

                ' Update the table name if necessary  
                Dim tableName As String = "Chores"

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
            MessageBox.Show("An error occurred while loading data into the grid.", "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
End Class