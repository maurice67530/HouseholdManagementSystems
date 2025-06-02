Imports System.Data.OleDb
Public Class Family_Schedule
    Private EventDictionary As New Dictionary(Of Date, List(Of String))
    Dim budget As Decimal = 7000
    Dim budgetlabel As New Label()
    Dim hasUnsavedChanges As Boolean = False
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
    .EventType = ComboBox3.SelectedItem.ToString,
    .IsBudgetRequired = ComboBox2.SelectedItem.ToString,
    .Amount = TextBox3.Text
    }

            Using conn As New OleDbConnection(HouseHoldManagment_Module.connectionString)
                conn.Open()

                ' Update the table name if necessary  


                ' Create an OleDbCommand to insert the Expense data into the database  
                Dim cmd As New OleDbCommand("INSERT INTO FamilySchedule ([Title], [Notes], [DateOfEvent], [StartTime], [EndTime], [AssignedTo], [EventType], [IsBudgetRequired], [Amount]) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)", conn)

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
                cmd.Parameters.AddWithValue("@IsBudgetRequired", ComboBox2.SelectedItem.ToString())
                cmd.Parameters.AddWithValue("@Amount", TextBox3.Text)

                Dim rowsAffected As Integer = cmd.ExecuteNonQuery()
                'Display a confirmation messageBox  
                MsgBox("Schedule Information Added!" & vbCrLf &
                "Title: " & schedule.Title.ToString & vbCrLf &
                "Notes: " & schedule.Notes.ToString & vbCrLf &
                "DateOfEvent: " & schedule.DateOfEvent & vbCrLf &
                "StartTime: " & schedule.StartTime & vbCrLf &
                "EndTime: " & schedule.EndTime & vbCrLf &
                "AssignedTo: " & schedule.AssignedTo & vbCrLf &
                 "EventType: " & schedule.EventType & vbCrLf &
                "IsBudgetRequired: " & schedule.IsBudgetRequired & vbCrLf &
                "Amount: " & schedule.Amount.ToString(), vbInformation, "Schedule confirmation")
            End Using

        Catch ex As Exception
            Debug.WriteLine($"Database error in btnAdd_Click: {ex.Message}")
            MessageBox.Show("Error saving Schedule to database: " & ex.Message & vbNewLine & ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Debug.WriteLine("Exiting btnAdd")

        End Try
        LoadScheduleFromDatabase()
        SubtractEventAmount()
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

        LoadBudgetAmount()
        'ComboBox3.Items.Clear()
        If Not ComboBox3.Items.Contains("Doctor's Visit") Then
            ComboBox3.Items.Add("Doctor's Visit")
        End If
        If Not ComboBox3.Items.Contains("School Trip") Then
            ComboBox3.Items.Add("School Trip")
        End If
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
            Dim IsBudgetRequired As String = ComboBox2.Text
            Dim Amount As Integer = TextBox3.Text

            Using conn As New OleDbConnection(HouseHoldManagment_Module.connectionString)

                conn.Open()

                ' Get the ID of the selected row (assuming your table has a primary key named "ID")  
                Dim selectedRow As DataGridViewRow = DataGridView1.SelectedRows(0)
                Dim ID As Integer = Convert.ToInt32(selectedRow.Cells("ID").Value) ' Change "ID" to your primary key column name  

                ' Create an OleDbCommand to update the expense data in the database  
                Dim cmd As New OleDbCommand("UPDATE [FamilySchedule] SET [Title] = ?, [Notes] =?, [DateOfEvent] = ?, [StartTime] = ?, [EndTime] = ?, [AssignedTo] = ?, [EventType] = ?, [IsBudgetRequired] = ?, [Amount] = ? WHERE [ID] = ?", conn)

                ' Set the parameter values from the UI controls  
                cmd.Parameters.AddWithValue("@Title", TextBox1.Text)
                cmd.Parameters.AddWithValue("@Notes", TextBox2.Text)
                cmd.Parameters.AddWithValue("@DateOfEvent", DateTimePicker1.Text)
                cmd.Parameters.AddWithValue("@StartTime", DateTimePicker2.Text)
                cmd.Parameters.AddWithValue("@EndTime", DateTimePicker3.Text)
                cmd.Parameters.AddWithValue("@AssignedTo", ComboBox1.SelectedItem.ToString())
                cmd.Parameters.AddWithValue("@EventType", ComboBox3.SelectedItem.ToString())
                cmd.Parameters.AddWithValue("@isBudgetRequired", ComboBox2.SelectedItem.ToString())
                cmd.Parameters.AddWithValue("@Amount", TextBox3.Text)
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
                ComboBox2.Text = selectedRow.Cells("IsBudgetRequired").Value.ToString()
                TextBox3.Text = selectedRow.Cells("Amount").Value.ToString()

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
        Dim selectedDate As Date = e.Start.Date

        If EventDictionary.ContainsKey(selectedDate) Then
            Dim messages As String = String.Join(vbCrLf, EventDictionary(selectedDate))
            MessageBox.Show("Events on " & selectedDate.ToLongDateString() & ":" & vbCrLf & messages)
        End If

        'Dim selectedDate As Date = e.Start
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
            ' MessageBox.Show("Error retrieving birthdays: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            'MessageBox.Show("Error saving Schedule to database: " & ex.Message & vbNewLine & ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        ' Display result once
        If eventsOnDate.Count > 0 Then
            Dim message As String = "Events on " & selectedDate.ToShortDateString() & ":" & vbCrLf & String.Join(vbCrLf, eventsOnDate)
            MessageBox.Show(message, "Family Calendar")
        End If

        ' Dim selectedDate As Date = e.Start

        For Each item As ListViewItem In ListView1.Items
            If item.SubItems(0).Text = selectedDate.ToShortDateString() Then
                MessageBox.Show("There is a " & item.SubItems(1).Text & " on " & selectedDate.ToShortDateString(), "Event Reminder")
                Exit Sub
            End If
        Next
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
            Dim cmd As OleDbCommand = Nothing

            Select Case eventType
                Case "Chores"
                    query = "SELECT AssignedTo, DueDate, Status FROM Chores ORDER BY DueDate"
                    cmd = New OleDbCommand(query, conn)

                Case "Birthdays"
                    query = "SELECT FirstName, LastName, DateOfBirth FROM PersonalDetails ORDER BY DateOfBirth"
                    cmd = New OleDbCommand(query, conn)

                Case "School Trip"
                    query = "SELECT startDate FROM Budget ORDER BY startDate DESC"
                    cmd = New OleDbCommand(query, conn)

                Case "Doctors Visit"
                    query = "SELECT dateOfExpense FROM Expense ORDER BY dateOfExpense DESC"
                    cmd = New OleDbCommand(query, conn)
            End Select

            If cmd IsNot Nothing Then
                Dim da As New OleDbDataAdapter(cmd)
                da.Fill(dt)
            End If
        End Using

        For Each row As DataRow In dt.Rows
            Dim item As New ListViewItem()

            Select Case eventType
                Case "Chores"
                    Dim assignedTo As String = row("AssignedTo").ToString()
                    Dim dueDate As Date = Convert.ToDateTime(row("DueDate"))
                    Dim status As String = row("Status").ToString()

                    item.Text = assignedTo
                    item.SubItems.Add(dueDate.ToShortDateString())
                    item.SubItems.Add(status)

                    If dueDate < Date.Today AndAlso status.ToLower() <> "completed" Then
                        item.ForeColor = Color.Blue
                    End If

                    messageLines.Add($"Assigned To: {assignedTo} | Due Date: {dueDate.ToShortDateString()} | Status: {status}")

                Case "Birthdays"
                    Dim firstName As String = row("FirstName").ToString()
                    Dim lastName As String = row("LastName").ToString()
                    Dim dob As Date = Convert.ToDateTime(row("DateOfBirth"))
                    Dim fullName As String = $"{firstName} {lastName}"

                    item.Text = fullName
                    item.SubItems.Add(dob.ToShortDateString())

                    If dob.Month = Date.Today.Month AndAlso dob.Day = Date.Today.Day Then
                        item.ForeColor = Color.DeepPink
                    Else
                        item.ForeColor = Color.Blue
                    End If

                    messageLines.Add($"Name: {firstName} {lastName} | Date of Birth: {dob.ToShortDateString()}")

                Case "School Trip"
                    Dim startDate As Date = Convert.ToDateTime(row("startDate"))
                    item.Text = "School Trip"
                    item.SubItems.Add(startDate.ToShortDateString())
                    item.ForeColor = Color.DarkGreen

                    messageLines.Add($"School Trip Start Date: {startDate.ToShortDateString()}")

                Case "Doctors Visit"
                    Dim visitDate As Date = Convert.ToDateTime(row("dateOfExpense"))
                    item.Text = "Doctor's Visit"
                    item.SubItems.Add(visitDate.ToShortDateString())
                    item.ForeColor = Color.Maroon

                    messageLines.Add($"Doctor's Visit Date: {visitDate.ToShortDateString()}")
            End Select

            ListView1.Items.Add(item)
        Next

        ' Show all records in a MessageBox
        If messageLines.Count > 0 Then
            Dim message As String = String.Join(Environment.NewLine, messageLines)
            Select Case eventType

                Case "School Trip"
                    MessageBox.Show("All School Trips:" & Environment.NewLine & message, "School Trip Dates", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Case "Doctors Visit"
                    MessageBox.Show("All Doctor's Visits:" & Environment.NewLine & message, "Doctor Visit Dates", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End Select
        End If
    End Sub
    Private Sub ComboBox3_SelectedIndexChanged_1(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged
        Dim selectedEvent As String = ComboBox3.SelectedItem.ToString()

        If selectedEvent = "Doctor's Visit" Then
            LoadDoctorVisits()
        ElseIf selectedEvent = "School Trip" Then
            LoadSchoolTrips()
        End If

        UpdateBoldedDates()

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
        If ComboBox3.SelectedItem.ToString() = "Doctor's Visit" Then
            SetupDoctorsVisitListView()
            LoadDoctorsVisitData()
        End If
        If ComboBox3.SelectedItem.ToString() = "School Trip" Then
            SetupSchoolTripListView()
            LoadBudgetStartDates()
        End If
        SubtractEventAmount()
    End Sub
    Private Sub AddEvent([date] As Date, description As String)
        If Not EventDictionary.ContainsKey([date]) Then
            EventDictionary([date]) = New List(Of String)
        End If

        If Not EventDictionary([date]).Contains(description) Then
            EventDictionary([date]).Add(description)
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
    Private Sub SetupDoctorsVisitListView()
        ListView1.Clear()
        ListView1.View = View.Details
        ListView1.FullRowSelect = True
        ListView1.GridLines = True

        ListView1.Columns.Add("Person", 120)
        ListView1.Columns.Add("Date of Visit", 100)
    End Sub
    Private Sub LoadDoctorsVisitData()
        ListView1.Items.Clear()
        Dim dt As New DataTable()

        Using conn As New OleDbConnection(HouseHoldManagment_Module.connectionString)
            conn.Open()
            ' Select only entries where Category is 'Doctor Visit'
            Dim query As String = "SELECT Person, DateOfexpenses FROM Expense"
            Dim cmd As New OleDbCommand(query, conn)
            Dim da As New OleDbDataAdapter(cmd)
            da.Fill(dt)
        End Using

        ' Add each doctor's visit record to the ListView
        For Each row As DataRow In dt.Rows
            Dim person As String = row("Person").ToString()
            Dim visitDate As Date = Convert.ToDateTime(row("DateOfexpenses"))

            Dim item As New ListViewItem(person)
            item.SubItems.Add(visitDate.ToShortDateString()) ' Display visit date

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
    Private Sub btnFilte_Click_1(sender As Object, e As EventArgs) Handles btnFilte.Click
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
    Private Sub SubtractEventAmount()
        Dim amountUsed As Decimal
        Dim selectedEvent As String = ComboBox3.SelectedItem?.ToString()

        ' Ensure the event type is valid
        If selectedEvent = "Doctor's Visit" OrElse selectedEvent = "School Trip" OrElse selectedEvent = "Birthdays" Then
            ' Check if the amount entered is numeric and valid
            If Decimal.TryParse(TextBox3.Text, amountUsed) Then
                If amountUsed <= budget Then
                    budget -= amountUsed
                    SaveUpdatedBudget(budget)
                    Label12.Text = "Budget: R" & budget.ToString("N2")
                Else
                    MessageBox.Show("Amount exceeds current budget!", "Budget Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                End If
            Else
                MessageBox.Show("Please enter a valid numeric amount.", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        End If
    End Sub
    Private Sub LoadDoctorVisits()
        Try
            EventDictionary.Clear() ' Clear previous events if needed

            Using conn As New OleDbConnection(HouseHoldManagment_Module.connectionString)
                conn.Open()
                Dim cmd As New OleDbCommand("SELECT DateOfexpenses, Person FROM Expense", conn)
                Dim reader As OleDbDataReader = cmd.ExecuteReader()

                While reader.Read()
                    If Not IsDBNull(reader("DateOfexpenses")) AndAlso Not IsDBNull(reader("Person")) Then
                        Dim visitDate As Date = CDate(reader("DateOfexpenses")).Date ' Remove time portion
                        Dim personName As String = reader("Person").ToString()
                        AddEvent(visitDate, "Doctor's Visit - " & personName)
                    End If
                End While
            End Using
        Catch ex As Exception
            MessageBox.Show("Error loading Doctor's Visit: " & ex.Message)
        End Try
    End Sub
    Private Sub LoadSchoolTrips()
        Try
            Using conn As New OleDbConnection(connectionString)
                conn.Open()
                Dim cmd As New OleDbCommand("SELECT StartDate FROM Budget", conn)
                Dim reader As OleDbDataReader = cmd.ExecuteReader()

                While reader.Read()
                    If Not IsDBNull(reader("StartDate")) Then
                        Dim tripDate As Date = CDate(reader("StartDate"))
                        AddEvent(tripDate, "School Trip")
                    End If
                End While
            End Using
        Catch ex As Exception
            MessageBox.Show("Error loading School Trips: " & ex.Message)
        End Try
    End Sub
    Private Sub UpdateBoldedDates()
        MonthCalendar1.RemoveAllBoldedDates()

        For Each evtDate As Date In EventDictionary.Keys
            MonthCalendar1.AddBoldedDate(evtDate)
        Next

        MonthCalendar1.UpdateBoldedDates()
    End Sub
    Private Sub LoadBudgetStartDates()
        ListView1.Items.Clear()
        Dim dt As New DataTable()

        Using conn As New OleDbConnection(HouseHoldManagment_Module.connectionString)
            conn.Open()
            ' Only selecting StartDate from Budget table
            Dim query As String = "SELECT StartDate FROM Budget ORDER BY StartDate DESC"
            Dim cmd As New OleDbCommand(query, conn)
            Dim da As New OleDbDataAdapter(cmd)
            da.Fill(dt)
        End Using

        ' Display each StartDate in the ListView
        For Each row As DataRow In dt.Rows
            Dim startDate As Date = Convert.ToDateTime(row("StartDate"))
            Dim item As New ListViewItem(startDate.ToShortDateString())
            ListView1.Items.Add(item)
        Next
    End Sub
    Private Sub SetupSchoolTripListView()
        ListView1.Clear()
        ListView1.View = View.Details
        ListView1.FullRowSelect = True
        ListView1.GridLines = True

        'ListView1.Columns.Add("Student", 120)
        ListView1.Columns.Add("Trip Date", 100)
    End Sub
    Private Sub LoadBudgetAmount()
        Try
            Using conn As New OleDbConnection(HouseHoldManagment_Module.connectionString)
                conn.Open()

                Dim cmd As New OleDbCommand("SELECT BudgetAmount FROM Budget", conn)
                Dim reader As OleDbDataReader = cmd.ExecuteReader()

                If reader.Read() Then
                    budget = Convert.ToDecimal(reader("BudgetAmount"))
                    Label12.Text = "Budget: R" & budget.ToString("N2")
                Else
                    Label12.Text = "Budget: R0.00"
                End If

                reader.Close()
            End Using
        Catch ex As Exception
            MessageBox.Show("Error loading budget amount: " & ex.Message, "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub SaveUpdatedBudget(updatedBudget As Decimal)
        Try
            Using conn As New OleDbConnection(HouseHoldManagment_Module.connectionString)
                conn.Open()

                ' IMPORTANT: Using correct field and primary key name from your table (BudgetAmou and ID)
                Dim sql As String = "UPDATE Budget SET BudgetAmount = @NewBudget WHERE ID = (SELECT MAX(ID) FROM Budget)"
                Dim cmd As New OleDbCommand(sql, conn)

                cmd.Parameters.AddWithValue("@NewBudget", updatedBudget)

                cmd.ExecuteNonQuery()
            End Using
        Catch ex As Exception
            MessageBox.Show("Error saving updated budget: " & ex.Message, "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
End Class


