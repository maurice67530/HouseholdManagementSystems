Imports System.Data.OleDb
Public Class Family_Schedule
    Public Property conn As New OleDbConnection(connectionString)
    'Public Const connectionString As String = " Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\Nedzamba\Source\Repos\maurice67530\HouseholdManagementSystems\HMS.accdb"
    Dim eventType As String = "Chores"
    Dim cmd As OleDbCommand
    Dim da As OleDbDataAdapter
    Dim dt As DataTable
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        Try

            Dim schedule As New FamilySchedule() With {
    .Title = (TextBox1.Text),
    .Notes = TextBox2.Text,
    .DateOfEvent = DateTimePicker1.Text,
    .StartTime = DateTimePicker2.Text,
    .EndTime = DateTimePicker3.Text,
    .AssignedTo = ComboBox1.SelectedItem.ToString,
    .EventType = ComboBox3.SelectedItem.ToString
    }


            Using conn As New OleDbConnection(HouseHoldManagment_Module.connectionString)
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
                cmd.Parameters.AddWithValue("@EventType", ComboBox3.SelectedItem.ToString())


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
            Using conn As New OleDbConnection(HouseHoldManagment_Module.connectionString)
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

        'AutoCreateChoreEvents()
        'AutoAddMealTimes()
        'AutoCreateTaskReminders()
        'MarkPhotoDayEvents()

        ' IntegrateChores()
        LoadFamilyCalendar()
    End Sub
    Private Sub LoadFamilyCalendar()
        Dim conStr As String = (HouseHoldManagment_Module.connectionString)
        Dim con As New OleDbConnection(conStr)
        Dim dt As New DataTable()

        Try
            Dim query As String = "SELECT Title, DateOfEvent, EventType, AssignedTo FROM FamilySchedule"
            Dim da As New OleDbDataAdapter(query, con)
            da.Fill(dt)

            ' Create lists to store dates
            Dim eventDates As New List(Of Date)()
            Dim birthdayDates As New List(Of Date)()

            For Each row As DataRow In dt.Rows
                Dim eventDate As Date = CDate(row("DateOfEvent"))

                If row("EventType").ToString.ToLower() = "birthday" Then
                    ' Add birthdays as annually bolded
                    If Not birthdayDates.Contains(eventDate) Then
                        birthdayDates.Add(eventDate)
                    End If
                Else
                    ' Add other events
                    If Not eventDates.Contains(eventDate) Then
                        eventDates.Add(eventDate)
                    End If
                End If
            Next

            ' Set bolded dates for regular events
            MonthCalendar1.BoldedDates = eventDates.ToArray()

            ' Set annually bolded dates for birthdays
            MonthCalendar1.AnnuallyBoldedDates = birthdayDates.ToArray()

            ' Store the DataTable globally if needed
            Me.Tag = dt

        Catch ex As Exception
            MessageBox.Show("Error loading calendar events: " & ex.Message)
        End Try

        LoadScheduleFromDatabase()
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

                Using conn As New OleDbConnection(HouseHoldManagment_Module.connectionString)
                    conn.Open()


                    ' Create the delete command  
                    Dim cmd As New OleDbCommand("DELETE FROM [FamilySchedule] WHERE [ID] = ?", conn)
                    cmd.Parameters.AddWithValue("@ID", ID) ' Primary key for matching record  

                    ' Execute the delete command  
                    Dim rowsAffected As Integer = cmd.ExecuteNonQuery()

                    If rowsAffected > 0 Then
                        MessageBox.Show("Schedule deleted successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
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
            Dim EventType As String = ComboBox3.Text

            Using conn As New OleDbConnection(HouseHoldManagment_Module.connectionString)

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
                cmd.Parameters.AddWithValue("@EventType", ComboBox3.SelectedItem.ToString())
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
                ComboBox3.Text = selectedRow.Cells("EventType").Value.ToString()
            End If
        Catch ex As Exception
            Debug.WriteLine("error selection data in the database")
            Debug.WriteLine($"Stack Trace: {ex.StackTrace}")
            MessageBox.Show("Error saving inventory to database: " & ex.Message & vbNewLine & ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub

    Public Sub PopulateComboboxFromDatabase(ByRef comboBox As ComboBox)
        Dim conn As New OleDbConnection(HouseHoldManagment_Module.connectionString)
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
    'Private Sub AutoCreateChoreEvents()
    '    Dim conn As New OleDbConnection(HouseHoldManagment_Module.connectionString)
    '    Dim da As New OleDbDataAdapter("SELECT Title, DueDate, AssignedTo FROM Chores", conn)
    '    Dim dt As New DataTable
    '    da.Fill(dt)

    '    Dim count As Integer = 0
    '    For Each row As DataRow In dt.Rows
    '        Dim cmd As New OleDbCommand("INSERT INTO FamilySchedule (Title, Notes, DateOfEvent, StartTime, EndTime, AssignedTo, EventType) VALUES (?, ?, ?, ?, ?, ?, ?)", conn)
    '        cmd.Parameters.AddWithValue("?", row("Title").ToString())
    '        cmd.Parameters.AddWithValue("?", "Auto-scheduled chore")
    '        cmd.Parameters.AddWithValue("?", CDate(row("DueDate")))
    '        cmd.Parameters.AddWithValue("?", #9:00:00 AM#)
    '        cmd.Parameters.AddWithValue("?", #10:00:00 AM#)
    '        cmd.Parameters.AddWithValue("?", row("AssignedTo").ToString())
    '        cmd.Parameters.AddWithValue("?", "Chore")

    '        conn.Open()
    '        cmd.ExecuteNonQuery()
    '        conn.Close()
    '        count += 1
    '    Next

    '    MessageBox.Show(count.ToString() & " chore event(s) added to the schedule.", "Chore Integration Complete")
    'End Sub
    Private Sub AutoCreateChoreEvents()
        Dim conn As New OleDbConnection(HouseHoldManagment_Module.connectionString)
        Dim da As New OleDbDataAdapter("SELECT Title, DueDate, AssignedTo FROM Chores", conn)
        Dim dt As New DataTable
        da.Fill(dt)

        Dim count As Integer = 0
        For Each row As DataRow In dt.Rows
            Dim choreDate As Date = CDate(row("DueDate"))
            Dim startTime As Date = choreDate.Date.AddHours(9)    ' 9:00 AM
            Dim endTime As Date = choreDate.Date.AddHours(10)     ' 10:00 AM

            Dim cmd As New OleDbCommand("INSERT INTO FamilySchedule (Title, Notes, DateOfEvent, StartTime, EndTime, AssignedTo, EventType) VALUES (?, ?, ?, ?, ?, ?, ?)", conn)
            cmd.Parameters.AddWithValue("?", row("Title").ToString())
            cmd.Parameters.AddWithValue("?", "Auto-scheduled chore")
            cmd.Parameters.AddWithValue("?", choreDate.ToString("dddd, MMMM dd, yyyy"))
            cmd.Parameters.AddWithValue("?", startTime.ToString("dddd, MMMM dd, yyyy hh:mm tt"))
            cmd.Parameters.AddWithValue("?", endTime.ToString("dddd, MMMM dd, yyyy hh:mm tt"))
            cmd.Parameters.AddWithValue("?", row("AssignedTo").ToString())
            cmd.Parameters.AddWithValue("?", "Chore")

            conn.Open()
            cmd.ExecuteNonQuery()
            conn.Close()
            count += 1
        Next
        MessageBox.Show(count.ToString() & " chore event(s) added to the schedule.", "Chore Integration Complete")
        LoadScheduleFromDatabase()
    End Sub

    Private Sub AutoAddMealTimes()
        Dim conn As New OleDbConnection(HouseHoldManagment_Module.connectionString)
        Dim da As New OleDbDataAdapter("SELECT MealName, StartDate, Description FROM MealPlans", conn)
        Dim dt As New DataTable
        da.Fill(dt)

        Dim count As Integer = 0

        For Each row As DataRow In dt.Rows
            Dim cmd As New OleDbCommand("INSERT INTO FamilySchedule (Title, Notes, DateOfEvent, StartTime, EndTime, AssignedTo, EventType) " &
                                    "VALUES (?, ?, ?, ?, ?, ?, ?)", conn)

            Dim mealDate As Date = CDate(row("StartDate"))
            Dim startTime As Date = mealDate.Date.AddHours(13) ' 1:00 PM
            Dim endTime As Date = mealDate.Date.AddHours(14)   ' 2:00 PM

            cmd.Parameters.AddWithValue("?", row("MealName").ToString())
            cmd.Parameters.AddWithValue("?", "Scheduled Meal")
            cmd.Parameters.AddWithValue("?", mealDate)
            cmd.Parameters.AddWithValue("?", startTime.ToString("dddd, MMMM dd, yyyy hh:mm tt"))
            cmd.Parameters.AddWithValue("?", endTime.ToString("dddd, MMMM dd, yyyy hh:mm tt"))
            cmd.Parameters.AddWithValue("?", row("Description").ToString())
            cmd.Parameters.AddWithValue("?", "Meal")

            conn.Open()
            cmd.ExecuteNonQuery()
            conn.Close()

            count += 1
        Next

        MessageBox.Show(count.ToString() & " meal(s) added to the family calendar.")
    End Sub
    Private Sub AutoCreateTaskReminders()
        Dim conn As New OleDbConnection(HouseHoldManagment_Module.connectionString)
        Dim da As New OleDbDataAdapter("SELECT Title, DueDate, AssignedTo FROM Tasks", conn)
        Dim dt As New DataTable
        da.Fill(dt)

        Dim count As Integer = 0
        For Each row As DataRow In dt.Rows
            Dim reminderDate As Date = CDate(row("DueDate")).AddDays(-1)
            Dim startTime As Date = reminderDate.Date.AddHours(8)    ' 8:00 AM
            Dim endTime As Date = reminderDate.Date.AddHours(8.5)    ' 8:30 AM

            Dim cmd As New OleDbCommand("INSERT INTO FamilySchedule (Title, Notes, DateOfEvent, StartTime, EndTime, AssignedTo, EventType) " &
                                    "VALUES (?, ?, ?, ?, ?, ?, ?)", conn)

            cmd.Parameters.AddWithValue("?", row("Title").ToString())
            cmd.Parameters.AddWithValue("?", "Task due soon")
            cmd.Parameters.AddWithValue("?", reminderDate.ToString("dddd, MMMM dd, yyyy"))
            cmd.Parameters.AddWithValue("?", startTime.ToString("dddd, MMMM dd, yyyy hh:mm tt"))
            cmd.Parameters.AddWithValue("?", endTime.ToString("dddd, MMMM dd, yyyy hh:mm tt"))
            cmd.Parameters.AddWithValue("?", row("AssignedTo"))
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
    Private Sub MonthCalendar1_DateChanged(sender As Object, e As DateRangeEventArgs) Handles MonthCalendar1.DateChanged
        Dim selectedDate As Date = e.Start

        Dim dt As DataTable = TryCast(Me.Tag, DataTable)
        If dt Is Nothing Then
            'MessageBox.Show("No events loaded.")
            Return
        End If

        ' Filter events for the selected date
        Dim eventsOnDate = dt.AsEnumerable().
        Where(Function(r) CDate(r("DateOfEvent")).Date = selectedDate.Date).
        Select(Function(r) r("EventType").ToString() & ": " & r("Title").ToString() & " (" & r("AssignedTo").ToString() & ")").
        ToList()

        'If eventsOnDate.Count = 0 Then
        '    'MessageBox.Show("No events for " & selectedDate.ToShortDateString(), "No Events")
        'Else
        '    Dim message As String = "Events on " & selectedDate.ToShortDateString() & ":" & vbCrLf & String.Join(vbCrLf, eventsOnDate)
        '    MessageBox.Show(message, "Family Calendar")
        'End If
        LoadScheduleFromDatabase()

        'Dim selectedDate As Date = e.Start

        ' Dim dt As DataTable = TryCast(Me.Tag, DataTable)
        ' Dim eventsOnDate As New List(Of String)

        ' Existing schedule events
        If dt IsNot Nothing Then
            eventsOnDate.AddRange(
            dt.AsEnumerable().
            Where(Function(r) CDate(r("DateOfEvent")).Date = selectedDate.Date).
            Select(Function(r) r("EventType").ToString() & ": " & r("Title").ToString() & " (" & r("AssignedTo").ToString() & ")")
        )
        End If

        ' Add birthdays
        Try
            conn.Open()
            Dim query As String = "SELECT FirstName, DateOfBirth FROM PersonalDetails WHERE MONTH(DateOfBirth) = ? AND DAY(DateOfBirth) = ?"
            Dim cmd As New OleDbCommand(query, conn)
            cmd.Parameters.AddWithValue("?", selectedDate.Month)
            cmd.Parameters.AddWithValue("?", selectedDate.Day)

            Dim reader As OleDbDataReader = cmd.ExecuteReader()

            While reader.Read()
                Dim firstName As String = reader("FirstName").ToString()
                Dim birthDate As Date = CDate(reader("DateOfBirth"))
                eventsOnDate.Add("Birthday: " & firstName & " (" & birthDate.ToShortDateString() & ")")
            End While

        Catch ex As Exception
            MessageBox.Show("Error retrieving birthdays: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            conn.Close()
        End Try

        If eventsOnDate.Count = 0 Then
            'MessageBox.Show("No events or birthdays for " & selectedDate.ToShortDateString(), "No Events")
        Else
            Dim message As String = "Events  " & selectedDate.ToShortDateString() & ":" & vbCrLf & String.Join(vbCrLf, eventsOnDate)
            MessageBox.Show(message, "Family Calendar")
        End If

        LoadScheduleFromDatabase()

    End Sub

    Private Sub ListView1_ItemDrag(sender As Object, e As ItemDragEventArgs) Handles ListView1.ItemDrag
        DoDragDrop(e.Item, DragDropEffects.Move)
    End Sub

    Private Sub MonthCalendar1_DragEnter(sender As Object, e As DragEventArgs) Handles MonthCalendar1.DragEnter
        If e.Data.GetDataPresent(GetType(ListViewItem)) Then
            e.Effect = DragDropEffects.Move
        End If
    End Sub

    Private Sub MonthCalendar1_DragDrop(sender As Object, e As DragEventArgs) Handles MonthCalendar1.DragDrop
        Dim item As ListViewItem = CType(e.Data.GetData(GetType(ListViewItem)), ListViewItem)
        Dim newDate = MonthCalendar1.SelectionStart

        ' Example: You can now update the due date of a chore or appointment in DB
        MessageBox.Show("Item dropped on " & newDate.ToShortDateString() & ". Update DB logic here.")
    End Sub

    Private Sub LoadAllEvents()
        ListView1.Items.Clear()
        Dim dt As New DataTable()
        Dim messageLines As New List(Of String)()



        Using conn As New OleDbConnection(HouseHoldManagment_Module.connectionString)
            conn.Open()
            Dim query As String = ""
            Dim cmd As OleDbCommand

            If eventType = "Chores" Then
                query = "SELECT AssignedTo, DueDate, Status FROM Chores ORDER BY DueDate"
                cmd = New OleDbCommand(query, conn)
            ElseIf eventType = "Birthdays" Then
                query = "SELECT FirstName, LastName, DateOfBirth FROM PersonalDetails ORDER BY DateOfBirth"
                cmd = New OleDbCommand(query, conn)
            End If

            Dim da As New OleDbDataAdapter(cmd)
            da.Fill(dt)
        End Using

        For Each row As DataRow In dt.Rows
            Dim item As New ListViewItem()

            If eventType = "Chores" Then
                Dim assignedTo As String = row("AssignedTo").ToString()
                Dim dueDate As Date = Convert.ToDateTime(row("DueDate"))
                Dim status As String = row("Status").ToString()

                item.Text = assignedTo
                item.SubItems.Add(dueDate.ToShortDateString())
                item.SubItems.Add(status)

                ' Optional coloring for overdue chores
                If dueDate < Date.Today AndAlso status.ToLower() <> "completed" Then
                    item.ForeColor = Color.Blue
                End If

                ListView1.Items.Add(item)

                ' Add to messagebox lines
                messageLines.Add($"Assigned To: {assignedTo} | Due Date: {dueDate.ToShortDateString()} | Status: {status}")

            ElseIf eventType = "Birthdays" Then
                Dim firstName As String = row("FirstName").ToString()
                Dim lastName As String = row("LastName").ToString()
                Dim dob As Date = Convert.ToDateTime(row("DateOfBirth"))
                Dim fullName As String = $"{firstName} {lastName}"

                item.Text = fullName
                item.SubItems.Add(dob.ToShortDateString())

                ' Optional coloring for today's birthdays
                If dob.Month = Date.Today.Month AndAlso dob.Day = Date.Today.Day Then
                    item.ForeColor = Color.DeepPink
                Else
                    item.ForeColor = Color.Blue
                End If

                ListView1.Items.Add(item)

                ' Add to messagebox lines
                messageLines.Add($"Name: {firstName} {lastName} | Date of Birth: {dob.ToShortDateString()}")
            End If

        Next

        ' Display all records in a messagebox
        If messageLines.Count > 0 Then
            Dim message As String = String.Join(Environment.NewLine, messageLines)
            If eventType = "Chores" Then
                MessageBox.Show("All Chores:" & Environment.NewLine & message, "All Chores", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Else
                MessageBox.Show("All Birthdays:" & Environment.NewLine & message, "All Birthdays", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
        End If
    End Sub

    Private Sub HighlightEventsOnCalendar()
        Dim eventDates As New List(Of Date)()

        Using conn As New OleDbConnection(HouseHoldManagment_Module.connectionString)
            conn.Open()
            Dim query As String = ""
            Dim cmd As OleDbCommand

            If eventType = "Chores" Then
                query = "SELECT DueDate, Status FROM Chores WHERE Status <> 'Completed'"
                cmd = New OleDbCommand(query, conn)
            ElseIf eventType = "Birthdays" Then
                query = "SELECT DateOfBirth FROM PersonalDetails"
                cmd = New OleDbCommand(query, conn)
            End If

            Dim reader = cmd.ExecuteReader()
            While reader.Read()
                Dim rawDate As Date = Convert.ToDateTime(reader(0))

                If eventType = "Birthdays" Then
                    ' Bold birthday for current year
                    Dim birthdayThisYear As Date = New Date(Date.Today.Year, rawDate.Month, rawDate.Day)
                    eventDates.Add(birthdayThisYear)
                ElseIf eventType = "Chores" Then
                    ' Only overdue chores
                    If rawDate.Date < Date.Today Then
                        eventDates.Add(rawDate.Date)
                    End If
                End If
            End While
        End Using

        MonthCalendar1.BoldedDates = eventDates.Distinct().ToArray()
    End Sub

    Private Sub ComboBox3_SelectedIndexChanged_1(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged
        eventType = ComboBox3.SelectedItem.ToString()
        HighlightEventsOnCalendar()
        LoadAllEvents()  ' <-- instead of LoadEventsByDate

        If ComboBox3.SelectedItem.ToString() = "Birthdays" Then
            SetupListView()
            LoadPersonalDetails()
        End If

        If ComboBox3.SelectedItem.ToString() = "Chores" Then
            SetupChoresListView()
            LoadChores()
        End If
    End Sub

    Private Sub SetupListView()
        ListView1.Clear()
        ListView1.View = View.Details
        ListView1.FullRowSelect = True
        ListView1.GridLines = True

        ListView1.Columns.Add("First Name", 120)
        ListView1.Columns.Add("Last Name", 120)
        ListView1.Columns.Add("Date of Birth", 120)
    End Sub
    Private Sub LoadPersonalDetails()
        ListView1.Items.Clear()
        Dim dt As New DataTable()

        Using conn As New OleDbConnection(HouseHoldManagment_Module.connectionString)
            conn.Open()
            Dim query As String = "SELECT FirstName, LastName, DateOfBirth FROM PersonalDetails ORDER BY DateOfBirth"
            Dim cmd As New OleDbCommand(query, conn)
            Dim da As New OleDbDataAdapter(cmd)
            da.Fill(dt)
        End Using

        For Each row As DataRow In dt.Rows
            Dim item As New ListViewItem(row("FirstName").ToString())
            item.SubItems.Add(row("LastName").ToString())
            Dim dob As Date = Convert.ToDateTime(row("DateOfBirth"))
            item.SubItems.Add(dob.ToShortDateString())

            ' Check if birthday is today
            If dob.Month = Date.Today.Month AndAlso dob.Day = Date.Today.Day Then
                item.ForeColor = Color.DeepPink  ' Color for birthdays today
            Else
                item.ForeColor = Color.Blue      ' Normal color for other birthdays
            End If

            ListView1.Items.Add(item)
        Next
    End Sub
    Private Sub SetupChoresListView()
        ListView1.Clear()
        ListView1.View = View.Details
        ListView1.FullRowSelect = True
        ListView1.GridLines = True

        ListView1.Columns.Add("Assigned To", 120)
        ListView1.Columns.Add("Due Date", 100)
        ListView1.Columns.Add("Status", 100)
    End Sub
    Private Sub LoadChores()
        ListView1.Items.Clear()
        Dim dt As New DataTable()

        Using conn As New OleDbConnection(HouseHoldManagment_Module.connectionString)
            conn.Open()
            Dim query As String = "SELECT AssignedTo, DueDate, Status FROM Chores ORDER BY DueDate"
            Dim cmd As New OleDbCommand(query, conn)
            Dim da As New OleDbDataAdapter(cmd)
            da.Fill(dt)
        End Using

        For Each row As DataRow In dt.Rows
            Dim item As New ListViewItem(row("AssignedTo").ToString())
            item.SubItems.Add(Convert.ToDateTime(row("DueDate")).ToShortDateString())
            item.SubItems.Add(row("Status").ToString())

            ' Optional: highlight overdue uncompleted chores
            If Convert.ToDateTime(row("DueDate")) < Date.Today AndAlso row("Status").ToString().ToLower() <> "completed" Then
                item.ForeColor = Color.Blue
            End If

            ListView1.Items.Add(item)
        Next
    End Sub
    Private Sub IntegrateChores()
        Using conn As New OleDbConnection(HouseHoldManagment_Module.connectionString)
            conn.Open()
            Dim query As String = "SELECT AssignedTo, DueDate FROM Chores WHERE Status <> 'Completed'"
            Using cmd As New OleDbCommand(query, conn)
                Using reader As OleDbDataReader = cmd.ExecuteReader()
                    While reader.Read()
                        Dim title As String = "Chore: " & reader("AssignedTo").ToString()
                        Dim dateOfEvent As Date = Convert.ToDateTime(reader("DueDate"))
                        Dim existsQuery As String = "SELECT COUNT(*) FROM FamilySchedule WHERE Title=? AND DateOfEvent=?"
                        Using checkCmd As New OleDbCommand(existsQuery, conn)
                            checkCmd.Parameters.AddWithValue("?", title)
                            checkCmd.Parameters.AddWithValue("?", dateOfEvent)
                            Dim exists As Integer = Convert.ToInt32(checkCmd.ExecuteScalar())
                            If exists = 0 Then
                                Dim insertQuery As String = "INSERT INTO FamilySchedule (Title, Notes, DateOfEvent, StartTime, EndTime, AssignedTo, EventType) VALUES (?, ?, ?, ?, ?, ?, ?)"
                                Using insertCmd As New OleDbCommand(insertQuery, conn)
                                    insertCmd.Parameters.AddWithValue("?", title)
                                    insertCmd.Parameters.AddWithValue("?", "Auto-created from Chores")
                                    insertCmd.Parameters.AddWithValue("?", dateOfEvent)
                                    insertCmd.Parameters.AddWithValue("?", #8:00 AM#)
                                    insertCmd.Parameters.AddWithValue("?", #9:00 AM#)
                                    insertCmd.Parameters.AddWithValue("?", reader("AssignedTo").ToString())
                                    insertCmd.Parameters.AddWithValue("?", "Chore")
                                    insertCmd.ExecuteNonQuery()
                                End Using
                            End If
                        End Using
                    End While
                End Using
            End Using
        End Using
    End Sub
End Class


