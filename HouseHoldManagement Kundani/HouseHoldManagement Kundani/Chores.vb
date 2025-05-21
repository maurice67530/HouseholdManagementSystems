Imports System.Drawing
Imports System.Data.OleDb
Imports System.IO
Imports System.Windows.Forms
Imports System.Net.Mail
Public Class chores
    Private toolTip1 As New ToolTip()
    Public Property conn As New OleDbConnection(connectionString)

    Private Sub chores_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        cmbChore.Visible = False

        Dim conn As New OleDbConnection(connectionString)

        toolTip1.SetToolTip(Button1, "Save")
        toolTip1.SetToolTip(Button2, "Edit")
        toolTip1.SetToolTip(Button4, "Delete")
        toolTip1.SetToolTip(Button5, "Refresh")
        toolTip1.SetToolTip(Button9, "Mark All As complete")
        toolTip1.SetToolTip(Button3, "Dashboard")
        toolTip1.SetToolTip(Button11, "Assign Chores")
        toolTip1.SetToolTip(Button7, "Filter")

        toolTip1.SetToolTip(Button6, "Sort")
        toolTip1.SetToolTip(Button12, "Auto Assignment")
        toolTip1.SetToolTip(Button10, "Check chores")
        toolTip1.SetToolTip(Button13, "Conflicts")

        toolTip1.SetToolTip(Button8, "Highlight")


        ' Check database connectivity 

        Try
            Debug.WriteLine("form load initialized successfully")

            ' Create a new OleDbConnection object and open the connection  

            conn.Open()

            ' Display the connection status on a button with a green background  
            Label11.Text = "Connected"
            Label11.BackColor = Color.Green
            Label11.ForeColor = Color.White



        Catch ex As Exception

            ' Display an error message  
            Debug.WriteLine($"Failed to initialize components from database: {ex.Message}")
            Debug.WriteLine($"Stack Trace: {ex.StackTrace}")
            MessageBox.Show("Error connecting to the database" & ex.Message, "Database Connection Error", MessageBoxButtons.OK, MessageBoxIcon.Error)

            ' Display the connection status on a button with a red background  
            Label11.Text = "Not Connected"
            Label11.BackColor = Color.Red
            Label11.ForeColor = Color.White



        Finally


            PopulateComboboxFromDatabase(CmbASS)
            loadChoresFromDatabase()
            ' Close the database connection  
            conn.Close()
        End Try

        'chores
        LoadBudgetDataFromDatabase()
        CheckRecurringChores()
        CheckPendingChores() ' Check pending chores when the form opens

        'Label16.Text = DateTime.Now.ToString(" HH:mm:ss")
        'Timer1.Interval = 1000
        'Timer1.Enabled = True

    End Sub


    Private Sub LinkLabel1_LinkClicked(sender As System.Object, e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        Expense.ShowDialog()

    End Sub

    Private Sub Button3_Click(sender As System.Object, e As System.EventArgs) Handles Button3.Click

        Close()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click



        If TxtTitle.Text = "" Then
            MsgBox("Please enter tittle on provided space")
            TxtTitle.Focus()
        End If
        Try
            Dim chore As New chores_() With {
           .Title = TxtTitle.Text,
           .AssignedTo = CmbASS.SelectedItem,
           .Priority = cmbpriority.SelectedItem,
           .Status = CMBstatus.SelectedItem,
           .Frequency = Cmbfre.SelectedItem,
           .DueDate = DateTimePicker1.Value,
           .Recurring = ComboBox1.SelectedItem,
           .Description = TxtDes.Text,
            .StartTime = DateTimePicker2.Value,
                .EndTime = DateTimePicker3.Value}


            '            Dim conn As New OleDbConnection(connectionString)

            conn.Open()
            Dim tablename As String = "Chores"
            Dim Cmd As New OleDbCommand($"INSERT INTO {tablename} ([Title], [AssignedTo], [Priority], [Status], [Frequency], [DueDate], [Recurring], [Description], [StartTime], [EndTime]) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)", conn)

            Cmd.Parameters.Clear()

            Cmd.Parameters.AddWithValue("@Title", chore.Title)
            Cmd.Parameters.AddWithValue("@AssignedTo", chore.AssignedTo)
            Cmd.Parameters.AddWithValue("@Priority", chore.Priority)
            Cmd.Parameters.AddWithValue("@Status", chore.Status)
            Cmd.Parameters.AddWithValue("@Frequency", chore.Frequency)
            Cmd.Parameters.AddWithValue("@DueDate", chore.DueDate)
            Cmd.Parameters.AddWithValue("@Recurring", chore.Recurring)
            Cmd.Parameters.AddWithValue("@Description", chore.Description)
            Cmd.Parameters.AddWithValue("@StartTime", chore.StartTime)
            Cmd.Parameters.AddWithValue("@EndTime", chore.EndTime)

            MsgBox("Chores Information Addded!" & vbCrLf &
              "Title: " & chore.Title & vbCrLf &
              "AssignedTo:" & chore.Description & vbCrLf &
              "Priority: " & chore.Priority & vbCrLf &
              "Status : " & chore.Status & vbCrLf &
              "Frequency: " & chore.AssignedTo & vbCrLf &
                "Recurring: " & chore.Recurring & vbCrLf &
              "Description: " & chore.Description & vbCrLf &
              "DueDate: " & chore.DueDate & vbCrLf & vbCrLf &
              "StartTime: " & chore.StartTime & vbCrLf &
              "EndTime: " & chore.EndTime & vbCrLf & vbCrLf, vbInformation, "Chores confirmation")

            '   MessageBox.Show("Chores Information saved to Database successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Cmd.ExecuteNonQuery()
            conn.Close()
        Catch ex As OleDbException
            '    Debug.WriteLine($"General error in button Save: {ex.Message}")
            '    Debug.WriteLine($"Stack Trace: {ex.StackTrace}")
            '    MessageBox.Show("Error saving control texts: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show($"Error Saving To database: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show("Error saving Chores to database: " & ex.Message & vbNewLine & ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)

        Catch ex As Exception
            MessageBox.Show($"Unexpected error: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        loadChoresFromDatabase()
        Debug.WriteLine("Existing button Save")
        CheckPendingChores()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Try


            Debug.WriteLine("entering button update")

            If DGVChores.SelectedRows.Count = 0 Then
                MessageBox.Show("Please select a record to update.", "No Selection", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            If DGVChores.SelectedRows.Count > 0 Then
                Debug.WriteLine("A row is selected for update")

            Else
                MessageBox.Show("Please select a Chore to update.", "update Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Debug.WriteLine("No row selected, exiting Button Update")
            End If
            Debug.WriteLine("Exiting Button Update")

            Dim Title As String = TxtTitle.Text
            Dim AssignedTo As String = CmbASS.SelectedItem
            Dim Priority As String = cmbpriority.SelectedItem
            Dim Status As String = CMBstatus.SelectedItem
            Dim Frequency As String = Cmbfre.SelectedItem
            Dim DueDate As String = DateTimePicker1.Value
            Dim Recurring As String = ComboBox1.SelectedItem
            Dim Description As String = TxtDes.Text
            Dim StartTime As String = DateTimePicker2.Value
            Dim EndTime As String = DateTimePicker3.Value

            Using conn As New OleDbConnection(connectionString)
                conn.Open()

                'Get the ID of the selected row (assuming your table has a primary key named "ID")  
                Dim selectedRow As DataGridViewRow = DGVChores.SelectedRows(0)
                Dim ID As Integer = Convert.ToInt32(selectedRow.Cells("ID").Value) ' Change "ID" to your primary key column name  

                'Create an OleDbCommand to update the personnel data in the database  
                Dim cmd As New OleDbCommand("UPDATE Chores SET [Title] = ?, [AssignedTo] = ?, [Priority] = ?, [Status] = ?, [Frequency] = ?, [DueDate] = ?, [Recurring]= ?, [Description] = ?, [StartTime] = ?, [EndTime] = ? WHERE [ID] = ?", conn)

                'Set the parameter values from the UI controls  


                cmd.Parameters.AddWithValue("@Title", TxtTitle.Text)
                cmd.Parameters.AddWithValue("@AssignedTo", CmbASS.SelectedItem)
                cmd.Parameters.AddWithValue("@Priority", cmbpriority.SelectedItem)
                cmd.Parameters.AddWithValue("@Status", CMBstatus.SelectedItem)
                cmd.Parameters.AddWithValue("@Frequency", Cmbfre.SelectedItem)
                cmd.Parameters.AddWithValue("@DueDate", Task_Management.DateTimePicker1.Value)
                cmd.Parameters.AddWithValue("@Recurring", ComboBox1.SelectedItem)
                cmd.Parameters.AddWithValue("@Description", TxtDes.Text)
                cmd.Parameters.AddWithValue("@StartTime", DateTimePicker2.Value)
                cmd.Parameters.AddWithValue("@EndTime", DateTimePicker3.Value)
                cmd.Parameters.AddWithValue("@ID", ID)
                MsgBox("Chores Updated Successfuly!", vbInformation, "Update Confirmation")

                cmd.ExecuteNonQuery()

            End Using
        Catch ex As FormatException
            Debug.WriteLine($"Format Error in Button Update: {ex.Message}")
            Debug.WriteLine($"Stack Trace: {ex.StackTrace}")
            MessageBox.Show($"Error updating chores in database: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show("Error saving chores to database: " & ex.Message & vbNewLine & ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Catch ex As Exception
            Debug.WriteLine($"unexpected error Button Update: {ex.Message}")
            Debug.WriteLine($"Stack Trace: {ex.StackTrace}")
            MessageBox.Show($"Unexpected error: {ex.Message}", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show("Error saving chores to database: " & ex.Message & vbNewLine & ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        loadChoresFromDatabase()
        Debug.WriteLine("exited button edit")
        CheckPendingChores()
    End Sub
    Public Sub loadChoresFromDatabase()
        Try

            Debug.WriteLine("Datagridview: populated successfully.")

            Using conn As New OleDbConnection(connectionString)
                conn.Open()

                Dim tableName As String = "Chores"
                Dim cmd As New OleDbCommand($"SELECT*FROM {tableName}", conn)
                Dim da As New OleDbDataAdapter(cmd)
                Dim dt As New DataTable
                da.Fill(dt)

                DGVChores.DataSource = dt

            End Using

        Catch ex As Exception
            Debug.WriteLine($"Datagridview: Failed to Populate {ex.Message}")
            Debug.WriteLine($"Stack Trace: {ex.StackTrace}")
            MessageBox.Show($"Error Loading chores data from database: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub DGVChores_SelectionChanged(sender As Object, e As EventArgs) Handles DGVChores.SelectionChanged
        Try
            Button1.Enabled = False

            If DGVChores.SelectedRows.Count > 0 Then
                Dim selectedRow As DataGridViewRow = DGVChores.SelectedRows(0)
                Debug.WriteLine("row selected on dgv")

                TxtTitle.Text = selectedRow.Cells("Title").Value.ToString()
                CmbASS.SelectedItem = selectedRow.Cells("AssignedTo").Value.ToString()
                cmbpriority.SelectedItem = selectedRow.Cells("Priority").Value.ToString()
                CMBstatus.SelectedItem = selectedRow.Cells("Status").Value.ToString()
                Cmbfre.SelectedItem = selectedRow.Cells("Frequency").Value.ToString()
                ComboBox1.SelectedItem = selectedRow.Cells("Recurring").Value.ToString()
                TxtDes.Text = selectedRow.Cells("Description").Value.ToString()
                DateTimePicker1.Value = selectedRow.Cells("DueDate").Value.ToString()
                DateTimePicker2.Value = selectedRow.Cells("StartTime").Value.ToString()
                DateTimePicker3.Value= selectedRow.Cells("EndTime").Value.ToString()

            End If
            Button1.Enabled = True
        Catch ex As Exception
            Debug.WriteLine("Data not selected: Error")
            Debug.WriteLine($"Stack Trace: {ex.StackTrace}")

        End Try
    End Sub
    Public Sub PopulateComboboxFromDatabase(ByRef comboBox As ComboBox)
        Dim conn As New OleDbConnection(connectionString)
        Try
            Debug.WriteLine("populate combobox successful")
            'open the database connection
            conn.Open()

            'retrieve the firstname and surname columns from the personaldetails tabel
            Dim query As String = "SELECT FirstName, LastName FROM PersonalDetails"
            Dim cmd As New OleDbCommand(query, conn)
            Dim reader As OleDbDataReader = cmd.ExecuteReader()

            'bind the retrieved data to the combobox
            CmbASS.Items.Clear()
            While reader.Read()
                CmbASS.Items.Add($"{reader("FirstName")} {reader("LastName")}")
            End While

            'close the database
            reader.Close()

        Catch ex As Exception
            'handle any exeptions that may occur  
            Debug.WriteLine("failed to populate combobox")
            Debug.WriteLine($"Stack Trace: {ex.StackTrace}")
            MessageBox.Show($"Error: {ex.StackTrace}")

        Finally
            'close the database connection
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Debug.WriteLine("Entering button delete")
        ' Check if there are any selected rows in the DataGridView for expenses  
        If DGVChores.SelectedRows.Count > 0 Then

            ' Get the selected row  
            Dim selectedRow As DataGridViewRow = DGVChores.SelectedRows(0)

            ' Assuming there is an "ID" column which is the primary key in the database  
            Dim PhotoID As Integer = Convert.ToInt32(selectedRow.Cells("ID").Value) ' Replace "ID" with your actual column name  
            'Dim DeletedBy As String

            ' Confirm deletion  
            Dim confirmationResult As DialogResult = MessageBox.Show("Are you sure you want to delete this Chores?", "Confirm Deletion", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)
            If confirmationResult = DialogResult.Yes Then
                Debug.WriteLine("User confirmation deletion.")
                ' Proceed with deletion  
                Try
                    Debug.WriteLine("Format errors in button delete")
                    Debug.WriteLine("Deleting data: Data delected")
                    Debug.WriteLine("Stack Trace: {ex.StackTrace}")
                    Using conn As New OleDbConnection(HouseHoldManagment_Module.connectionString)
                        conn.Open()

                        ' Create the delete command  
                        Dim cmd As New OleDbCommand("DELETE FROM [Chores] WHERE [ID] = ?", conn)
                        cmd.Parameters.AddWithValue("@ID", PhotoID) ' Primary key for matching record  

                        ' Execute the delete command  
                        Dim rowsAffected As Integer = cmd.ExecuteNonQuery()

                        If rowsAffected > 0 Then
                            MessageBox.Show("Chores deleted successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
                            ' Optionally refresh DataGridView or reload from database  
                            ' PopulateDataGridView()
                        Else
                            MessageBox.Show("No Chores was deleted. Please check if the ID exists.", "Deletion Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                        End If
                    End Using
                    loadChoresFromDatabase()
                Catch ex As Exception
                    Debug.WriteLine("Failed to delete data")
                    Debug.Write($"Stack Trace: {ex.StackTrace}")
                    MessageBox.Show($"An error occurred while deleting the Chores: {ex.Message}", "Deletion Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
            End If
        Else
            Debug.WriteLine("User cancelled deletion")
            MessageBox.Show("Please select Chores to delete.", "Deletion Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End If
        Debug.WriteLine("Exiting button delete")
    End Sub
    Private Sub Button5_Click(sender As Object, e As EventArgs)
        'Dim itemCount As Integer = 0

        ''loop through the inventory array
        'For i As Integer = 0 To -1
        '    'prompt the user to continue or exit
        '    Dim UserInput As String = InputBox("Current item: " & (i) & vbCrLf & "Enter 'Done' to quit or press OK to continue.", "Count Items")

        '    'check if the user wants to quit
        '    If UserInput.ToLower = "done" Then
        '        Exit For
        '    End If
        '    'increment the counter
        '    itemCount += 1
        'Next
        ''display the total count in the label
        'lblitemCount.Text = "Total Items counted: " & itemCount
    End Sub
    Private Sub Button5_Click_1(sender As Object, e As EventArgs) Handles Button5.Click
        loadChoresFromDatabase()
    End Sub
    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        DGVChores.Sort(DGVChores.Columns("DueDate"), System.ComponentModel.ListSortDirection.Ascending)
    End Sub
    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Dim selectedFrequency As String = If(Cmbfre.SelectedItem IsNot Nothing, Cmbfre.SelectedItem.ToString(), "")
        Dim selectedPriority As String = If(cmbpriority.SelectedItem IsNot Nothing, cmbpriority.SelectedItem.ToString(), "")

        HouseHoldManagment_Module.FilterChores(selectedFrequency, selectedPriority)
    End Sub
    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        Try
            For Each row As DataGridViewRow In DGVChores.Rows
                If row.Cells("DueDate").Value IsNot Nothing Then
                    Dim DueDate As DateTime = Convert.ToDateTime(row.Cells("DueDate").Value)
                    Dim Status As String = row.Cells("Status").Value.ToString

                    If DueDate < DateTime.Now AndAlso Status <> "Completed" Then
                        row.DefaultCellStyle.BackColor = Color.Gray
                    Else
                        row.DefaultCellStyle.BackColor = Color.WhiteSmoke
                    End If
                End If
            Next
            Dim InclompletedCount As Integer = 0
            For Each row As DataGridViewRow In DGVChores.Rows
                If row.Cells("Status").Value IsNot Nothing AndAlso row.Cells("Status").Value.ToString() <> "Completed " Then
                    InclompletedCount += 1
                End If
            Next
            Label12.Text = "Incompleted chores: " & InclompletedCount.ToString
        Catch ex As Exception
            MessageBox.Show("Error highligting overdue chores")
        End Try
    End Sub
    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        For Each row As DataGridViewRow In
                DGVChores.SelectedRows
            row.Cells("Status").Value = "Completed"
        Next
    End Sub
    Private Sub Button10_Click(sender As Object, e As EventArgs)
        If ComboBox1.SelectedItem Then
            Select Case Cmbfre.Text
                Case "Daily"
                    DateTimePicker1.Value = DateTimePicker1.Value.AddDays(1)
                Case "Weekly"
                    DateTimePicker1.Value = DateTimePicker1.Value.AddDays(7)
                Case "Monthly"
                    DateTimePicker1.Value = DateTimePicker1.Value.AddDays(1)
            End Select
        End If
    End Sub
    Private Sub Button12_Click(sender As Object, e As EventArgs)


    End Sub
    'chores
    Public Sub LoadBudgetDataFromDatabase()
        Try
            '  Dim dataTable As DataTable = HouseHold.GetData("SELECT * FROM Budget")
            ' DataGridView1.DataSource = DataTable
            Debug.WriteLine("Populate Datagridview: Datagridview populated successfully.")
            Using conn As New OleDbConnection(connectionString)
                conn.Open()

                Dim tableName As String = "Chores"

                Dim cmd As New OleDbCommand($"SELECT*FROM {tableName}", conn)

                Dim da As New OleDbDataAdapter(cmd)
                Dim dt As New DataTable
                da.Fill(dt)
                DGVChores.DataSource = dt
            End Using
        Catch ex As Exception
            Debug.WriteLine("Failed to populate Gatagridview")
            Debug.Write($"Stack Trace: {ex.StackTrace}")
            MessageBox.Show($"Error Loading chores data from database: {ex.Message}", "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        If CmbASS.SelectedIndex = -1 OrElse cmbChore.SelectedIndex = -1 OrElse Cmbfre.SelectedIndex = -1 Then
            MessageBox.Show("Please select a person, chore, and frequency.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If

        Dim personID As String = CmbASS.SelectedItem
        Dim choreID As String = cmbChore.SelectedValue

        ' Check if the person already has a chore assigned
        If PersonAlreadyHasChore(personID) Then
            MessageBox.Show("This person already has a chore assigned!", "Duplicate Assignment", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If

        Dim frequency As String = Cmbfre.SelectedItem.ToString()
        AssignChore(personID, choreID, frequency)
    End Sub
    Private Sub AssignChore(personID As String, choreID As String, frequency As String)
        Using conn As New OleDbConnection(connectionString)
            Dim query As String = "INSERT INTO AssignedChores (PersonID, choreToDo, Frequency, AssignedDate) VALUES (@PersonID, @ChoreID, @Frequency, @AssignedDate)"
            Dim cmd As New OleDbCommand(query, conn)

            cmd.Parameters.AddWithValue("@PersonID", personID)
            cmd.Parameters.AddWithValue("@choreToDo", choreID)
            cmd.Parameters.AddWithValue("@Frequency", frequency)
            cmd.Parameters.AddWithValue("@AssignedDate", DateTimePicker1.Value)

            conn.Open()
            cmd.ExecuteNonQuery()
            conn.Close()

            MessageBox.Show("Chore assigned successfully.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Using
    End Sub
    Private Function PersonAlreadyHasChore(personID As Integer) As Boolean
        Using conn As New OleDbConnection(connectionString)
            Dim query As String = "SELECT COUNT(*) FROM Chores WHERE PersonID = @PersonID"
            Dim cmd As New OleDbCommand(query, conn)
            cmd.Parameters.AddWithValue("@PersonID", personID)

            conn.Open()
            Dim count As Integer = CInt(cmd.ExecuteScalar())
            conn.Close()

            Return count > 0 ' If count > 0, person already has a chore
        End Using
    End Function
    Private Sub CheckRecurringChores()

        Using con As New OleDbConnection(connectionString)
            Dim query As String = "SELECT Title FROM Chores WHERE Frequency <> 'One-Time' AND Frequency IS NOT NULL"

            Using cmd As New OleDbCommand(query, con)
                Try
                    con.Open()
                    Dim reader As OleDbDataReader = cmd.ExecuteReader()
                    Dim recurringChores As New List(Of String)
                    While reader.Read()
                        recurringChores.Add(reader("Title").ToString())
                    End While
                    reader.Close()
                    ' Display message if there are recurring chores
                    If recurringChores.Count > 0 Then
                        Dim message As String = "Recurring Chores Found:" & vbCrLf & String.Join(vbCrLf, recurringChores)
                        MessageBox.Show(message, "Recurring Chores", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    End If

                Catch ex As Exception
                    MessageBox.Show("Error: " & ex.Message)
                End Try
            End Using
        End Using
    End Sub
    Private Sub CheckPendingChores()
        Using con As New OleDbConnection(connectionString)
            Dim query As String = "SELECT Title, DueDate FROM Chores WHERE Status = 'In Progress' AND DueDate"

            Using cmd As New OleDbCommand(query, con)
                Try
                    con.Open()
                    Dim reader As OleDbDataReader = cmd.ExecuteReader()

                    Dim pendingChores As New List(Of String)

                    While reader.Read()
                        Dim choreName As String = reader("Title").ToString()
                        Dim dueDate As DateTime = Convert.ToDateTime(reader("DueDate"))
                        pendingChores.Add(choreName & " (Due: " & dueDate.ToString("yyyy-MM-dd") & ")")
                    End While

                    reader.Close()

                    ' Display reminder if there are pending chores
                    If pendingChores.Count > 0 Then
                        Dim message As String = "Pending Chores Reminder:" & vbCrLf & String.Join(vbCrLf, pendingChores)
                        MessageBox.Show(message, "Chore Reminder", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    End If

                Catch ex As Exception
                    MessageBox.Show("Error: " & ex.Message)
                End Try
            End Using
        End Using
    End Sub
    Private Sub Button12_Click_1(sender As Object, e As EventArgs) Handles Button12.Click
        Dim choreID As Integer = GetSelectedChoreID()

        If choreID = 0 Then

            MsgBox("Please select a chore first.", MsgBoxStyle.Exclamation, "Selection Required")

            Return

        End If

        CompleteChore(choreID)

        ' Refresh DataGridView after updating
        loadChoresFromDatabase()
        CheckPendingChores()
    End Sub
    Private Function GetSelectedChoreID() As Integer
        If DGVChores.SelectedRows.Count > 0 Then
            Return Convert.ToInt32(DGVChores.SelectedRows(0).Cells("ID").Value)
        End If
        Return 0 ' Return 0 if no row is selected
    End Function
    ' Method to complete the selected chore and auto-assign the next available person
    Private Sub CompleteChore(choreID As Integer)
        Dim connString As String = HouseHoldManagment_Module.connectionString
        Using conn As New OleDb.OleDbConnection(connString)
            Try
                conn.Open()

                ' Get chore details
                Dim query As String = "SELECT Title, AssignedTo, Frequency, Recurring, DueDate FROM Chores WHERE ID = @ID"
                Using cmd As New OleDb.OleDbCommand(query, conn)
                    cmd.Parameters.AddWithValue("@ID", choreID)
                    Using reader As OleDb.OleDbDataReader = cmd.ExecuteReader()
                        If reader.Read() Then
                            Dim title As String = reader("Title").ToString()
                            Dim assignedTo As String = reader("AssignedTo").ToString()
                            Dim frequency As String = reader("Frequency").ToString()
                            Dim recurring As String = reader("Recurring").ToString()
                            Dim dueDate As Date = Convert.ToDateTime(reader("DueDate"))

                            ' Mark as completed
                            UpdateChoreStatus(choreID, "Completed")

                            If recurring.ToLower() = "yes" Then
                                ' Get the next available person
                                Dim nextPerson As String = GetNextAvailablePerson()

                                ' Calculate the next due date
                                Dim nextDueDate As Date = CalculateNextDueDate(frequency, dueDate)

                                ' Update only the selected chore with the next person and due date
                                UpdateRecurringChore(choreID, nextPerson, nextDueDate)

                                MsgBox($"Following Chore: '{title}' completed. Assigned to {nextPerson} with new due date {nextDueDate.ToShortDateString()}.", MsgBoxStyle.Information, "Chore Updated")
                            Else
                                MsgBox($"Following Chore: '{title}' completed and not recurring.", MsgBoxStyle.Information, "Chore Completed")
                            End If
                        End If
                    End Using
                End Using
            Catch ex As Exception
                MsgBox("Error: " & ex.Message, MsgBoxStyle.Critical, "Error")
            End Try
        End Using
    End Sub
    ' Update only the selected chore status to "Completed"
    Private Sub UpdateChoreStatus(choreID As Integer, status As String)
        Dim connString As String = HouseHoldManagment_Module.connectionString
        Using conn As New OleDb.OleDbConnection(connString)
            Try
                conn.Open()
                Dim query As String = "UPDATE Chores SET Status = @Status WHERE ID = @ID"
                Using cmd As New OleDb.OleDbCommand(query, conn)
                    cmd.Parameters.AddWithValue("@Status", status)
                    cmd.Parameters.AddWithValue("@ID", choreID)
                    cmd.ExecuteNonQuery()
                End Using
            Catch ex As Exception
                MsgBox("Error updating chore status: " & ex.Message, MsgBoxStyle.Critical, "Error")
            End Try
        End Using
    End Sub
    ' Get the next available person who is not already assigned in DataGridView or database  
    Private Function GetNextAvailablePerson() As String
        Dim connString As String = HouseHoldManagment_Module.connectionString
        Dim availablePerson As String = String.Empty
        Dim assignedPeople As New List(Of String)

        ' Get already assigned people from DataGridView  
        For Each row As DataGridViewRow In DGVChores.Rows
            If row.Cells("AssignedTo").Value IsNot Nothing Then
                assignedPeople.Add(row.Cells("AssignedTo").Value.ToString())
            End If
        Next

        ' Get an available person from the database who is NOT in DataGridView or already assigned  
        Using conn As New OleDb.OleDbConnection(connString)
            Try
                conn.Open()
                Dim query As String = "SELECT TOP 1 AssignedTo, DateOfEvent FROM FamilySchedule " &
                                  "WHERE AssignedTo + ' ' + DateOfEvent NOT IN (" &
                                  String.Join(",", assignedPeople.Select(Function(p) "'" & p & "'")) & ") " &
                                  "ORDER BY AssignedTo, DateOfEvent"

                Using cmd As New OleDb.OleDbCommand(query, conn)
                    Dim reader As OleDb.OleDbDataReader = cmd.ExecuteReader()
                    If reader.Read() Then
                        availablePerson = reader("AssignedTo").ToString() & " " & reader("LastName").ToString()
                    End If
                End Using
            Catch ex As Exception
                MsgBox("Error retrieving next available person: " & ex.Message, MsgBoxStyle.Critical, "Error")
            End Try
        End Using

        Return availablePerson
    End Function
    ' Calculate the next due date based on frequency (Daily, Weekly, Monthly)
    Private Function CalculateNextDueDate(frequency As String, lastDueDate As Date) As Date
        Select Case frequency.ToLower()
            Case "daily"
                Return lastDueDate.AddDays(1)
            Case "weekly"
                Return lastDueDate.AddDays(7)
            Case "monthly"
                Return lastDueDate.AddMonths(1)
            Case Else
                Return lastDueDate
        End Select
    End Function
    ' Update only the selected chore with the next person and new due date
    Private Sub UpdateRecurringChore(choreID As Integer, nextPerson As String, nextDueDate As Date)
        Dim connString As String = HouseHoldManagment_Module.connectionString
        Using conn As New OleDb.OleDbConnection(connString)
            Try
                conn.Open()
                Dim query As String = "UPDATE Chores SET AssignedTo = @AssignedTo, DueDate = @DueDate WHERE ID = @ID"
                Using cmd As New OleDb.OleDbCommand(query, conn)
                    cmd.Parameters.AddWithValue("@AssignedTo", nextPerson)
                    cmd.Parameters.AddWithValue("@DueDate", nextDueDate)
                    cmd.Parameters.AddWithValue("@ID", choreID)
                    cmd.ExecuteNonQuery()
                End Using
            Catch ex As Exception
                MsgBox("Error updating recurring chore: " & ex.Message, MsgBoxStyle.Critical, "Error")
            End Try
        End Using
    End Sub
    Private Sub Button10_Click_1(sender As Object, e As EventArgs) Handles Button10.Click
        'Start the task timer when the button Is clicked

        Timer1.Start()
        TextBox1.AppendText("Schedules started ." & vbCrLf)
        'Define a list of chores
        Label13.Text = ($"chores checked  {DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")}" & vbCrLf)

        'Check If any chores are due (Pending) for the selected frequency (daily, weekly, monthly)
        ' Dim selectedFrequency As String = ComboBox2.SelectedItem.ToString()
        'LoadChoresByFrequency(selectedFrequency)
        'Update chores based on frequency
        ' Dim choresToUpdate As List(Of Integer) = GetPendingChoresToUpdate(selectedFrequency)
        'For Each choreId As String In choresToUpdate
        '  Next

        Dim quiry = "Select AssignedTo & Title  FROM chores "

        Using conn As New OleDbConnection(HouseHoldManagment_Module.connectionString)
            Using cmd As New OleDbCommand(quiry, conn)
                conn.Open()
                '   MessageBox.Show("  : " & cmd.ExecuteScalar.ToString & MessageBoxIcon.Warning)
            End Using
        End Using

        Dim connString As String = HouseHoldManagment_Module.connectionString
        Dim selectedChoreID As Integer
        Dim selectedChore As String = ""
        Dim currentPerson As String = ""
        Dim assignedPerson As String = ""

        Static lastAssignedPerson As String = ""

        ' Check if a row is selected in DataGridView   
        If DGVChores.SelectedRows.Count > 0 Then
            selectedChoreID = CInt(DGVChores.SelectedRows(0).Cells("ID").Value)
            selectedChore = DGVChores.SelectedRows(0).Cells("Title").Value.ToString()

            If DGVChores.SelectedRows(0).Cells("AssignedTo").Value IsNot Nothing Then
                currentPerson = DGVChores.SelectedRows(0).Cells("AssignedTo").Value.ToString()
            End If
        Else
            MsgBox("Please select a chore to assign.", MsgBoxStyle.Exclamation, "Chore Assignment")
            Exit Sub
        End If

        ' Get all available people except the current person and the last assigned person
        Dim availablePeople As New List(Of String)

        For Each person As String In CmbASS.Items
            If person <> currentPerson AndAlso person <> lastAssignedPerson Then
                availablePeople.Add(person)
            End If
        Next

        ' Validate if there are available people
        If availablePeople.Count = 0 Then
            MsgBox("No different person available for assignment.", MsgBoxStyle.Exclamation, "Chore Assignment")
            Exit Sub
        End If

        ' Randomly select a different person from the available people
        Dim rnd As New Random()
        assignedPerson = availablePeople(rnd.Next(0, availablePeople.Count))

        ' Get the frequency of the selected chore from the database (Daily, Weekly, Monthly)
        Dim choreFrequency As String = ""
        Dim nextDueDate As DateTime

        Using conn As New OleDb.OleDbConnection(connString)
            Try
                conn.Open()
                Dim query As String = "SELECT Frequency, DueDate FROM Chores WHERE ID = ?"
                Using cmd As New OleDb.OleDbCommand(query, conn)
                    cmd.Parameters.AddWithValue("?", selectedChoreID)

                    Using reader As OleDb.OleDbDataReader = cmd.ExecuteReader()
                        If reader.Read() Then
                            choreFrequency = reader("Frequency").ToString()
                            nextDueDate = Convert.ToDateTime(reader("DueDate"))
                        End If
                    End Using
                End Using

                ' Calculate the new due date based on the frequency         
                Select Case choreFrequency
                    Case "Daily"
                        nextDueDate = nextDueDate.AddDays(1)
                    Case "Weekly"
                        nextDueDate = nextDueDate.AddDays(7)
                    Case "Monthly"
                        nextDueDate = nextDueDate.AddMonths(1)
                End Select

                ' Assign the chore in the database and update the next due date
                Dim updateQuery As String = "UPDATE Chores SET AssignedTo = ?, DueDate = ? WHERE ID = ?"

                Using cmd As New OleDb.OleDbCommand(updateQuery, conn)
                    cmd.Parameters.AddWithValue("?", assignedPerson)
                    cmd.Parameters.AddWithValue("?", nextDueDate)
                    cmd.Parameters.AddWithValue("?", selectedChoreID)
                    cmd.ExecuteNonQuery()
                End Using

                ' Update DataGridView with new assignment and due date
                For Each row As DataGridViewRow In DGVChores.Rows
                    If row.Cells("ID").Value = selectedChoreID Then
                        row.Cells("AssignedTo").Value = assignedPerson
                        row.Cells("DueDate").Value = nextDueDate.ToShortDateString()

                        Exit For
                    End If
                Next

                ' Show confirmation
                MsgBox("Chore Reassigned: " & selectedChore & " → " & assignedPerson & vbCrLf & "Next Due Date: " & nextDueDate.ToShortDateString(), MsgBoxStyle.Information, "Success")

                ' Update lastAssignedPerson for the next assignment
                lastAssignedPerson = assignedPerson

                ' Refresh the available persons list
            Catch ex As Exception
                MsgBox("Error assigning chore: " & ex.Message, MsgBoxStyle.Critical, "Error")
            End Try
        End Using
    End Sub

    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click
        CheckDailyChoreOverload()
    End Sub
    Private Sub cmbAssignedTo_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CmbASS.SelectedIndexChanged
        Dim selectedPerson As String = CmbASS.SelectedItem.ToString()
        CheckTimeOverlapForPerson(selectedPerson)
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick

    End Sub

    ''maaano
    Private Sub CheckTimeOverlapForPerson(person As String)
        Dim conflictRows As New HashSet(Of Integer)

        ' Reset all row colors
        For Each row As DataGridViewRow In DGVChores.Rows
            row.DefaultCellStyle.BackColor = Color.White
        Next

        ' Collect rows for selected person
        Dim personChores As New List(Of Tuple(Of Integer, DateTime)) ' (rowIndex, hour)

        For i = 0 To DGVChores.Rows.Count - 2 ' skip new row
            Dim row = DGVChores.Rows(i)
            Dim rowPerson = row.Cells("AssignedTo").Value?.ToString().Trim()
            Dim hourStr = row.Cells("StartTime").Value?.ToString().Trim()

            If rowPerson = person AndAlso Not String.IsNullOrEmpty(hourStr) Then
                Dim hourVal As DateTime
                If DateTime.TryParse(hourStr, hourVal) Then
                    personChores.Add(Tuple.Create(i, hourVal))
                End If
            End If
        Next

        ' Check for hour overlap
        For i = 0 To personChores.Count - 2
            For j = i + 1 To personChores.Count - 1
                If personChores(i).Item2.Hour = personChores(j).Item2.Hour Then
                    conflictRows.Add(personChores(i).Item1)
                    conflictRows.Add(personChores(j).Item1)
                End If
            Next
        Next

        ' Highlight conflicts
        For Each idx In conflictRows
            DGVChores.Rows(idx).DefaultCellStyle.BackColor = Color.LightCoral
        Next

        ' Show message if conflict found
        If conflictRows.Count > 0 Then
            Button1.Visible = False
            MessageBox.Show("This person has multiple chores at the same hour. Please fix the conflicts.", "Time Conflict", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        Else
            Button1.Visible = True
        End If

    End Sub
    Private Sub CheckDailyChoreOverload()
        Dim choreCounts As New Dictionary(Of String, Dictionary(Of Date, Integer))

        For Each row As DataGridViewRow In DGVChores.Rows
            If Not row.IsNewRow Then
                Dim person = row.Cells("AssignedTo").Value?.ToString().Trim()
                Dim freq = row.Cells("Frequency").Value?.ToString().Trim()
                Dim dateStr = row.Cells("DueDate").Value?.ToString().Trim()

                If Not String.IsNullOrEmpty(person) AndAlso freq = "Daily" AndAlso Not String.IsNullOrEmpty(dateStr) Then
                    Dim choreDate As Date
                    If Date.TryParse(dateStr, choreDate) Then
                        If Not choreCounts.ContainsKey(person) Then
                            choreCounts(person) = New Dictionary(Of Date, Integer)
                        End If
                        If Not choreCounts(person).ContainsKey(choreDate.Date) Then
                            choreCounts(person)(choreDate.Date) = 0
                        End If
                        choreCounts(person)(choreDate.Date) += 1
                    End If
                End If
            End If
        Next

        ' Check for overloads and show message
        For Each person In choreCounts.Keys
            For Each choreDate In choreCounts(person).Keys
                If choreCounts(person)(choreDate) >= 3 Then
                    MessageBox.Show($"{person} has {choreCounts(person)(choreDate)} chores on {choreDate:d}. Please review the schedule.",
                                    "Daily Chore Overload", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    Exit Sub ' Show message once for the first overload found
                End If
            Next
        Next
    End Sub

    Private Sub Panel1_Paint(sender As Object, e As PaintEventArgs) Handles Panel1.Paint

    End Sub
End Class