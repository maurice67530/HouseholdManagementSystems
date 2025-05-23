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
            HighlightScheduleEventsByType()
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

                HighlightScheduleEventsByType()
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
        tooltip.SetToolTip(btnRefresh, "Refresh")
        tooltip.SetToolTip(btnFilte, "Filter")


        PopulateComboboxFromDatabase(ComboBox1)
        LoadScheduleFromDatabase()

        DataGridView1.Columns("DateOfEvent").DefaultCellStyle.Format = "dd, MMMM yyyy"
        DataGridView1.Columns("StartTime").DefaultCellStyle.Format = "dd, MMMM yyyy hh:mm tt"
        DataGridView1.Columns("EndTime").DefaultCellStyle.Format = "dd, MMMM yyyy hh:mm tt"
        AutoIntegrateAllEvents()

        LoadFamilyCalendar()
        LoadScheduleFromDatabase()
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

                ' Enable/disable the buttons based on the selected person  
                btnSave.Enabled = False
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
    Private Sub MonthCalendar1_DateChanged(sender As Object, e As DateRangeEventArgs) Handles MonthCalendar1.DateChanged

        Dim selectedDate As Date = e.Start
        Dim dt As DataTable = TryCast(Me.Tag, DataTable)


        ' Refresh the schedule data first
        LoadScheduleFromDatabase()
        dt = TryCast(Me.Tag, DataTable)
        If dt Is Nothing Then Return

        Dim eventsOnDate As New List(Of String)

        ' Existing FamilySchedule events
        eventsOnDate.AddRange(
        dt.AsEnumerable().
        Where(Function(r) CDate(r("DateOfEvent")).Date = selectedDate.Date).
        Select(Function(r) r("EventType").ToString() & ": " & r("Title").ToString() & " (" & r("AssignedTo").ToString() & ")")
    )

        ' Add birthdays
        Try
            Using conn As New OleDbConnection(HouseHoldManagment_Module.connectionString)
                conn.Open()
                Dim query As String = "SELECT FirstName, DateOfBirth FROM PersonalDetails WHERE MONTH(DateOfBirth) = ? AND DAY(DateOfBirth) = ?"
                Dim cmd As New OleDbCommand(query, conn)
                cmd.Parameters.AddWithValue("?", selectedDate.Month)
                cmd.Parameters.AddWithValue("?", selectedDate.Day)

                Using reader As OleDbDataReader = cmd.ExecuteReader()
                    While reader.Read()
                        Dim firstName As String = reader("FirstName").ToString()
                        Dim birthDate As Date = CDate(reader("DateOfBirth"))
                        eventsOnDate.Add("Birthday: " & firstName & " (" & birthDate.ToShortDateString() & ")")
                    End While
                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show("Error retrieving birthdays: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        ' Display result once
        If eventsOnDate.Count > 0 Then
            Dim message As String = "Events on " & selectedDate.ToShortDateString() & ":" & vbCrLf & String.Join(vbCrLf, eventsOnDate)
            MessageBox.Show(message, "Family Calendar")
        End If
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
    Private Sub ComboBox3_SelectedIndexChanged_1(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged
        eventType = ComboBox3.SelectedItem.ToString()
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
    Private Sub btnRefresh_Click(sender As Object, e As EventArgs) Handles btnRefresh.Click
        LoadScheduleFromDatabase()
        TextBox1.Text = ""
        TextBox2.Text = ""
        ComboBox1.Text = ""
        ComboBox3.Text = ""
        ListView1.Items.Clear()
    End Sub
    Private Sub AutoIntegrateAllEvents()
        Dim conn As New OleDbConnection(HouseHoldManagment_Module.connectionString)
        conn.Open()

        Dim countChores As Integer = 0
        Dim countMeals As Integer = 0
        Dim countTasks As Integer = 0

        ' --- CHORES ---
        Dim choreCmd As New OleDbCommand("SELECT Title, DueDate, AssignedTo FROM Chores", conn)
        Dim choreReader As OleDbDataReader = choreCmd.ExecuteReader()
        While choreReader.Read()
            Dim title As String = choreReader("Title").ToString()
            Dim dueDate As Date = CDate(choreReader("DueDate"))
            Dim startTime As Date = dueDate.Date.AddHours(9)
            Dim endTime As Date = dueDate.Date.AddHours(10)
            Dim assignedTo As String = choreReader("AssignedTo").ToString()

            If Not EventExists(conn, title, dueDate) Then
                InsertEvent(conn, title, "Auto-scheduled chore", dueDate, startTime, endTime, assignedTo, "Chore")
                countChores += 1
            End If
        End While
        choreReader.Close()

        ' --- MEAL PLANS ---
        Dim mealCmd As New OleDbCommand("SELECT MealName, StartDate, Description FROM MealPlans", conn)
        Dim mealReader As OleDbDataReader = mealCmd.ExecuteReader()
        While mealReader.Read()
            Dim title As String = mealReader("MealName").ToString()
            Dim startDate As Date = CDate(mealReader("StartDate"))
            Dim startTime As Date = startDate.Date.AddHours(13)
            Dim endTime As Date = startDate.Date.AddHours(14)
            Dim description As String = mealReader("Description").ToString()

            If Not EventExists(conn, title, startDate) Then
                InsertEvent(conn, title, "Scheduled Meal", startDate, startTime, endTime, description, "Meal")
                countMeals += 1
            End If
        End While
        mealReader.Close()

        ' --- TASK REMINDERS ---
        Dim taskCmd As New OleDbCommand("SELECT Title, DueDate, AssignedTo FROM Tasks", conn)
        Dim taskReader As OleDbDataReader = taskCmd.ExecuteReader()
        While taskReader.Read()
            Dim title As String = taskReader("Title").ToString()
            Dim reminderDate As Date = CDate(taskReader("DueDate")).AddDays(-1)
            Dim startTime As Date = reminderDate.Date.AddHours(8)
            Dim endTime As Date = reminderDate.Date.AddHours(8.5)
            Dim assignedTo As String = taskReader("AssignedTo").ToString()

            If Not EventExists(conn, title, reminderDate) Then
                InsertEvent(conn, title, "Task due soon", reminderDate, startTime, endTime, assignedTo, "Task")
                countTasks += 1
            End If
        End While
        taskReader.Close()

        conn.Close()
        MessageBox.Show(
    countChores & " chore(s)" & vbCrLf &
    countMeals & " meal(s)" & vbCrLf &
    countTasks & " task(s) added.",
    "All Events Integrated",
    MessageBoxButtons.OK,
    MessageBoxIcon.Information)

        LoadScheduleFromDatabase()
    End Sub
    Private Sub HighlightScheduleEventsByType()
        For Each row As DataGridViewRow In DataGridView1.Rows
            If Not row.IsNewRow AndAlso row.Cells("EventType").Value IsNot Nothing Then
                Dim eventType As String = row.Cells("EventType").Value.ToString().Trim()

                Select Case eventType
                    Case "Chore"
                        row.DefaultCellStyle.BackColor = Color.LightPink
                    Case "Meal"
                        row.DefaultCellStyle.BackColor = Color.LightGreen
                    Case "Task"
                        row.DefaultCellStyle.BackColor = Color.LightBlue
                    Case "Birthday"
                        row.DefaultCellStyle.BackColor = Color.LightPink
                    Case Else
                        row.DefaultCellStyle.BackColor = Color.White
                End Select
                row.DefaultCellStyle.SelectionBackColor = row.DefaultCellStyle.BackColor
                row.DefaultCellStyle.SelectionForeColor = Color.Black
            End If
        Next
    End Sub

    ' Declare this at the top of your form
    Private currentEventIndex As Integer = 0
    Private eventTypes As String() = {"Chore", "Meal", "Task"}
    Private Sub btnFilte_Click(sender As Object, e As EventArgs) Handles btnFilte.Click
        Try
            ' Get current event type to show
            Dim eventTypeToShow As String = eventTypes(currentEventIndex)

            ' Build query
            Dim query As String = "SELECT Title, Notes, DateOfEvent, StartTime, EndTime, AssignedTo, EventType " &
                              "FROM FamilySchedule WHERE EventType = @EventType"

            Using conn As New OleDbConnection(HouseHoldManagment_Module.connectionString)
                conn.Open()
                Using cmd As New OleDbCommand(query, conn)
                    cmd.Parameters.AddWithValue("@EventType", eventTypeToShow)
                    Dim adapter As New OleDbDataAdapter(cmd)
                    Dim dt As New DataTable()
                    adapter.Fill(dt)
                    DataGridView1.DataSource = dt

                    ' Show info about what was viewed
                    MessageBox.Show("Now viewing: " & eventTypeToShow & " Events" & vbCrLf &
                                "Total Events Found: " & dt.Rows.Count, "Event Filter", MessageBoxButtons.OK, MessageBoxIcon.Information)
                End Using
            End Using

            ' Highlight event rows
            HighlightScheduleEventsByType()

            ' Check if this was the LAST event in the cycle
            If currentEventIndex = eventTypes.Length - 1 Then
                MessageBox.Show("You have reached the end of the event list." & vbCrLf &
                            "All events (Chores, Meals, Tasks) have been displayed.",
                            "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                btnFilte.Enabled = False
            Else
                ' Prepare for next event type
                currentEventIndex += 1
                'btnFilte.Text = "View: " & eventTypes(currentEventIndex) & " Events"
            End If

        Catch ex As Exception
            MessageBox.Show("Error displaying events: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
End Class


