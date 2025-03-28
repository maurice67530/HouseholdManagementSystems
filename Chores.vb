Imports System.Windows.Forms
Imports System.Data.OleDb
Imports System.Drawing
Imports System.Net.Mail
Imports System.Net
Public Class chores
    Private row As Object


    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs)
        Budjet.ShowDialog()

    End Sub

    Private Sub chores_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        LoadChoresFromDatabase()


        ' Set tooltips For buttons
        ToolTip1.SetToolTip(Button1, "Dashboard")
        ToolTip1.SetToolTip(Button2, "Save")
        ToolTip1.SetToolTip(Button3, "Edit")
        ToolTip1.SetToolTip(Button4, "Delete")
        ToolTip1.SetToolTip(Button5, "Sort")
        ToolTip1.SetToolTip(Button6, "Filter")
        ToolTip1.SetToolTip(Button7, "Highlight")
        ToolTip1.SetToolTip(Button8, "Calculate")

        ' Get pending chores from the database
        '    Dim pendingChores As List(Of String) = GetPendingChores()

        '    ' Send an email if there are pending chores
        '    If pendingChores.Count > 0 Then
        '        Dim messageBody As String = "The following chores are still In progress:" & vbCrLf & String.Join(vbCrLf, pendingChores)
        '        SendEmail("austinmulalo113@gmail.com", "Chores Alert: Pending Tasks", messageBody)
        '    Else
        '        MessageBox.Show("No pending chores found.", "Chores Check", MessageBoxButtons.OK, MessageBoxIcon.Information)
        '    End If
        'End Sub

        'Private Function GetPendingChores() As List(Of String)
        '    Dim pendingChores As New List(Of String)

        '    Try
        '        Using conn As New OleDbConnection(connectionString)
        '            conn.Open()
        '            Dim query As String = "SELECT Title FROM CHORES WHERE Status = 'In Progress'"
        '            Using cmd As New OleDbCommand(query, conn)
        '                Using reader As OleDbDataReader = cmd.ExecuteReader()
        '                    While reader.Read()
        '                        pendingChores.Add(reader("Title").ToString())
        '                    End While
        '                End Using
        '            End Using
        '        End Using
        '    Catch ex As Exception
        '        MessageBox.Show("Database error: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        '    End Try

        '    Return pendingChores
        'End Function

        'Private Sub SendEmail(toEmail As String, subject As String, body As String)
        '    Try
        '        Dim smtpClient As New SmtpClient("smtp.gmail.com") With {
        '                        .Port = 587,
        '                        .Credentials = New Net.NetworkCredential("austinmulalo113@gmail.com", "oqsa qwqa bhjc nzoe"), ' Use App Password for security
        '                        .EnableSsl = True
        '                    }

        '        Dim mail As New MailMessage() With {
        '            .From = New MailAddress("austinmulalo113gmail.com;"),
        '            .Subject = subject,
        '            .Body = body
        '        }
        '        mail.To.Add(toEmail)

        '        smtpClient.Send(mail)
        '        MessageBox.Show("Chores alert email sent successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
        '    Catch ex As Exception
        '        MessageBox.Show("Failed to send email: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        '    End Try

        ' Get pending chores from the database
        Dim pendingChores As List(Of String) = GetPendingChores()

        ' Send an email if there are pending chores
        If pendingChores.Count > 0 Then
            Dim messageBody As String = "The following chores are still In progress:" & vbCrLf & String.Join(vbCrLf, pendingChores)
            SendEmail("austinmulalo113@gmail.com", "Chores Alert: Pending Tasks", messageBody)
        Else
            MessageBox.Show("No pending chores found.", "Chores Check", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If
    End Sub

    Private Function GetPendingChores() As List(Of String)
        Dim pendingChores As New List(Of String)

        Try
            Using conn As New OleDbConnection(connectionString)
                conn.Open()
                Dim query As String = "SELECT Title FROM CHORES WHERE Status = 'In Progress'"
                Using cmd As New OleDbCommand(query, conn)
                    Using reader As OleDbDataReader = cmd.ExecuteReader()
                        While reader.Read()
                            pendingChores.Add(reader("Title").ToString())
                        End While
                    End Using
                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show("Database error: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        Return pendingChores
    End Function

    Private Sub SendEmail(toEmail As String, subject As String, body As String)
        Try
            Dim smtpClient As New SmtpClient("smtp.gmail.com") With {
                            .Port = 587,
                            .Credentials = New Net.NetworkCredential("austinmulalo113@gmail.com", "oqsa qwqa bhjc nzoe"), ' Use App Password for security
                            .EnableSsl = True
                        }

            Dim mail As New MailMessage() With {
                .From = New MailAddress("austinmulalo113@gmail.com"),
                .Subject = subject,
                .Body = body
            }
            mail.To.Add(toEmail)

            smtpClient.Send(mail)
            MessageBox.Show("Chores alert email sent successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            MessageBox.Show("Failed to send email: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Try
            Dim Chore As New Chores With {
            .choresID = TextBox1.Text,
            .Assigned = ComboBox1.Text,
            .Title = TextBox2.Text,
            .Priority = CmbPriority.Text,
            .Status = ComboBox2.Text,
            .Frequency = Cmbfre.Text,
            .DueDate = DateTimePicker1.Text,
            .Recurring = TextBox7.Text}

            Using conn As New OleDbConnection(Module1.connectionString)
                conn.Open()

                Dim tablename As String = "CHORES"

                Dim cmd As New OleDbCommand("INSERT INTO [CHORES] ([ChoresID],[Assigned],[Title],[Priority],[Status],[Frequency],[DueDate],[Recurring]) VALUES (@ChoresID ,@Assigned ,@Title,@Priority, @Status, @Frequency,@DueDate,@Recurring)", conn)

                'set the parameter values from the UI controls

                'class declaration  

                ''params
                cmd.Parameters.AddWithValue("@ChoresID", Chore.choresID)
                cmd.Parameters.AddWithValue("@Assigned", Chore.Assigned)
                cmd.Parameters.AddWithValue("@Title", Chore.Title)
                cmd.Parameters.AddWithValue("@Priority", Chore.Priority)
                cmd.Parameters.AddWithValue("@Status", Chore.Status)
                cmd.Parameters.AddWithValue("@Frequency", Chore.Frequency)
                cmd.Parameters.AddWithValue("@DueDate", Chore.DueDate)
                cmd.Parameters.AddWithValue("@Recurring", Chore.Recurring)
                cmd.ExecuteNonQuery()
            End Using


            MsgBox("chores Added!" & vbCrLf &
                 "ChoresID: " & Chore.choresID & vbCrLf &
                   "Assigned: " & Chore.Assigned & vbCrLf &
                   "Title: " & Chore.Title & vbCrLf &
                   "Priority: " & Chore.Priority & vbCrLf &
                   "Status: " & Chore.Status & vbCrLf &
                   "Frequency: " & Chore.Frequency & vbCrLf &
                    "DueDate: " & Chore.DueDate & vbCrLf &
                   "Recurring: " & Recurring, vbInformation, "Chores Confirmation")



        Catch ex As OleDbException
            Debug.WriteLine($"Database error: {ex.Message}")
            Debug.WriteLine($"stack Trace:{ex.StackTrace}")
            'MessageBox.Show("Error saving test to database. Please check the connection.", "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Catch ex As Exception
            Debug.WriteLine($"General error: {ex.Message}")
            'MessageBox.Show("An unexpected error occurred.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            conn.Close()
        Finally

        End Try

    End Sub
    Public Sub LoadChoresFromDatabase()
        Try
            Debug.WriteLine("populated connected succesfully")
            Using conn As New OleDbConnection(Module1.connectionString)
                conn.Open()

                Dim tableName As String = "CHORES"
                Dim cmd As New OleDbCommand($"SELECT * FROM {tableName}", conn)

                Dim da As New OleDbDataAdapter(cmd)
                Dim dt As New DataTable
                da.Fill(dt)

                DataGridView1.DataSource = dt

            End Using

        Catch ex As Exception

            Debug.WriteLine("populated fail to initialize succesfully")
            MessageBox.Show($"Error Loading  Expenses data from database: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Debug.WriteLine($"stack Trace:{ex.StackTrace}")
        End Try
    End Sub
    Private Sub DataGridView1_SelectionChanged(sender As Object, e As EventArgs)
        'Load the data from  the selected row into UI controls 
        If DataGridView1.SelectedRows.Count > 0 Then
            Dim selectedRow As DataGridViewRow = DataGridView1.SelectedRows(0)
            TextBox1.Text = selectedRow.Cells("ChoresID").Value.ToString()
            ComboBox1.Text = selectedRow.Cells("Assigned").Value.ToString()
            TextBox2.Text = selectedRow.Cells("Title").Value.ToString()
            CmbPriority.Text = selectedRow.Cells("Priority").Value.ToString()
            ComboBox2.Text = selectedRow.Cells("Status").Value.ToString()
            Cmbfre.Text = selectedRow.Cells("Frequency").Value.ToString()
            DateTimePicker1.Text = selectedRow.Cells("DueDate").Value.ToString()
            TextBox7.Text = selectedRow.Cells("Recurring").Value.ToString()
        End If

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If DataGridView1.SelectedRows.Count = 0 Then
            MessageBox.Show("Please select a record to update.", "No Selection", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        Try
            Debug.WriteLine("Entering btnEdit")
            Dim ChoresID As Integer = TextBox1.Text
            Dim Assigned As String = ComboBox1.Text
            Dim Title As String = TextBox2.Text
            Dim Priority As String = CmbPriority.Text
            Dim Status As String = ComboBox2.Text
            Dim Frequency As String = Cmbfre.Text
            Dim Recurring As String = TextBox7.Text




            Using conn As New OleDbConnection(Module1.connectionString)
                conn.Open()

                ' Get the ID of the selected row (assuming your table has a primary key named "ID")  
                Dim selectedRow As DataGridViewRow = DataGridView1.SelectedRows(0)
                Dim ID As Integer = Convert.ToInt32(selectedRow.Cells("ID").Value) ' Change "ID" to your primary key column name  

                ' Create an OleDbCommand to update the expense data in the database  

                Dim cmd As New OleDbCommand("UPDATE TO [CHORES] SET [ChoresID] = ?, [Assigned]  = ?, [Title] = ?, [Priority] = ?, [Status] = ?, [Frequency] = ?,[Duedate]=?,[Recurring]=?WHERE [ID] = ?", conn)
                'Set the parameter values from the UI controls 
                cmd.Parameters.AddWithValue("@ChoresID", TextBox1.Text)
                cmd.Parameters.AddWithValue("@Assigned", ComboBox1.Text)
                cmd.Parameters.AddWithValue("@Title", TextBox2.Text)
                cmd.Parameters.AddWithValue("@Priority", CmbPriority.Text)
                cmd.Parameters.AddWithValue("@Satus", ComboBox2.Text)
                cmd.Parameters.AddWithValue("@Frequency", Cmbfre.Text)
                cmd.Parameters.AddWithValue("@DueDate", DateTimePicker1.Text)
                cmd.Parameters.AddWithValue("@Recurring", TextBox7.Text)


                'cmd.Parameters.AddWithValue("@ID", TaskManagementID) ' Primary key for matching record  
                'cmd.ExecuteNonQuery()


                MsgBox("Task Management Updated Successfuly!", vbInformation, "Update Confirmation")


            End Using
        Catch ex As OleDbException
            Debug.WriteLine($"stack Trace:{ex.StackTrace}")
            MessageBox.Show($"Error updating TaskManagement in database: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Catch ex As Exception
            MessageBox.Show($"Unexpected error: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        conn.Close()
    End Sub

    'Private persons As List(Of String) = New List(Of String)()
    'Private chores As List(Of String) = New List(Of String)()

    'Private Const NotificationDays As Integer = 3 ' Days before due date to notify

    'Dim currentTaskID As Integer = 1 ' This should be set to the task ID you want to load.

    'Private Sub LoadData()
    '    Using conn As New OleDbConnection(connectionString)
    '        conn.Open()

    '        Dim peopleCommand As New OleDbCommand("SELECT ID, FirstName FROM  personaldetails", conn)
    '        Dim choresCommand As New OleDbCommand("SELECT ID, Title, Frequency FROM CHORES", conn)

    '        Dim peopleAdapter As New OleDbDataAdapter(peopleCommand)
    '        Dim choresAdapter As New OleDbDataAdapter(choresCommand)

    '        Dim peopleTable As New DataTable()
    '        Dim choresTable As New DataTable()

    '        peopleAdapter.Fill(peopleTable)
    '        choresAdapter.Fill(choresTable)

    '        ' Now we can populate the ListBox with chore assignments 
    '        AssignChores(peopleTable, choresTable)
    '    End Using
    'End Sub
    'Private Sub AssignChores(Personnel As DataTable, chores As DataTable)
    '    ListBox1.Items.Clear()

    '    Dim rng As New Random()
    '    Dim currentDate As DateTime = DateTime.Now
    '    Dim assignments As New Dictionary(Of String, List(Of Integer))() ' Tracks assignments per person

    '    For Each choreRow As DataRow In chores.Rows
    '        Dim choreID As Integer = CInt(choreRow("Id"))
    '        'Dim choreDescription As String = choreRow("Descriptiopn").ToString()
    '        Dim choreFrequency As String = choreRow("Frequency").ToString() ' Get the frequency
    '        Dim dueDate As DateTime = DateTime.Parse(choreRow("DueDate").ToString()) ' Get the due date
    '        Dim personRow As DataRow = Personnel.Rows(rng.Next(Personnel.Rows.Count))
    '        Dim personName As String = personRow("FirstName").ToString()


    '        ' Ensure that we are only assigning chores that are due
    '        If IsChoreDue(currentDate, dueDate, choreFrequency) Then
    '            ' Get a random person and check if they can be assigned the chore
    '            Dim personAssigned As Boolean = False
    '            Dim attempts As Integer = 0

    '            While Not personAssigned AndAlso attempts < Personnel.Rows.Count
    '                Dim personRo As DataRow = Personnel.Rows(rng.Next(Personnel.Rows.Count))
    '                Dim personNam As String = personRow("FirstName").ToString()

    '                ' Check if this person already has this chore assigned
    '                If Not assignments.ContainsKey(choreID) Then
    '                    assignments(choreID) = New List(Of Integer)()
    '                End If

    '                'Dim alreadyAssigned As Boolean = assignments(choreID).Contains(personName)

    '                'If Not assignments(personName).Contains(choreID) Then
    '                '    'PersonAlreadyHasChore()

    '                ListBox1.Items.Add($"Chore: {choreID} | Assigned To: {personName} | Frequency: {choreFrequency}| Due Date: {dueDate.ToShortDateString()}")

    '                '' Check if due date is approaching
    '                'If (dueDate - currentDate).TotalDays <= NotificationDays Then
    '                '    NotifyUser(choreDescription, personName, dueDate)
    '                'End If

    '                ' Check if due date is approaching and if the chore is still recurring
    '                If IsChoreDue(currentDate, dueDate, choreFrequency) Then
    '                    NotifyUser(choreID, personName, dueDate)
    '                End If

    '                '' Check if the chore should be rescheduled based on its frequency
    '                'Dim nextDueDate As DateTime = GetNextDueDate(dueDate, choreFrequency)
    '                'If nextDueDate > currentDate Then
    '                '    ' Update the due date in the database if you're storing it, 
    '                '    ' or just handle the rescheduling logic here
    '                'End If

    '                ' Calculate the next due date
    '                Dim nextDueDate As DateTime = GetNextDueDate(dueDate, choreFrequency)
    '                If nextDueDate > currentDate Then
    '                    conn.Open()
    '                    ' Save assignment to the database
    '                    Dim insertCommand As New OleDbCommand("INSERT INTO AssignedChores (FirstName, Frequency, DueDate) VALUES (@FirstName, @Frequency, @DueDate)", Module1)
    '                    insertCommand.Parameters.AddWithValue("@FirstName", personName)
    '                    'insertCommand.Parameters.AddWithValue("@Descriptiopn", choreID)
    '                    insertCommand.Parameters.AddWithValue("@Frequency", choreFrequency)
    '                    insertCommand.Parameters.AddWithValue("@DueDate", dueDate)

    '                    insertCommand.ExecuteNonQuery() ' Execute the insert

    '                    conn.Close()
    '                    personAssigned = True
    '                End If
    '                'End If
    '                attempts += 1
    '            End While

    '            ' Optional: Notify the user if the chore is due soon
    '            If IsChoreDue(currentDate, dueDate, choreFrequency) Then
    '                NotifyUser(choreID, personName, dueDate)
    '            End If
    '        End If
    '    Next
    'End Sub
    'Private Function IsChoreDue(currentDate As DateTime, dueDate As DateTime, frequency As String) As Boolean
    '    ' Check if the chore is due based on the frequency
    '    Dim daysUntilDue = (dueDate - currentDate).TotalDays
    '    If daysUntilDue <= NotificationDays Then
    '        Return True
    '    End If
    '    Return False
    'End Function

    'Private Function GetNextDueDate(dueDate As DateTime, frequency As String) As DateTime
    '    ' Calculate the next due date based on frequency
    '    Select Case frequency.ToLower()
    '        Case "daily"
    '            Return dueDate.AddDays(1)
    '        Case "weekly"
    '            Return dueDate.AddDays(7)
    '        Case "monthly"
    '            Return dueDate.AddMonths(1)
    '        Case Else
    '            Return dueDate ' Return the same date if frequency is unknown
    '    End Select
    'End Function
    'Private Sub NotifyUser(choreDescription As String, personName As String, dueDate As DateTime)
    '    ' Notify user about the upcoming due chore
    '    MessageBox.Show($"Reminder: The chore '{choreDescription}' assigned to {personName} is due on {dueDate.ToShortDateString()}.", "Chore Due Reminder", MessageBoxButtons.OK, MessageBoxIcon.Information)
    'End Sub
    Private Sub LoadData()
        Using conn As New OleDbConnection(connectionString)
            conn.Open()

            Dim peopleCommand As New OleDbCommand("SELECT ID, FirstName FROM  personaldetails", conn)
            Dim choresCommand As New OleDbCommand("SELECT ID, Title, Frequency FROM CHORES", conn)

            Dim peopleAdapter As New OleDbDataAdapter(peopleCommand)
            Dim choresAdapter As New OleDbDataAdapter(choresCommand)

            Dim peopleTable As New DataTable()
            Dim choresTable As New DataTable()

            peopleAdapter.Fill(peopleTable)
            choresAdapter.Fill(choresTable)

            ' Now we can populate the ListBox with chore assignments 
            AssignChores(peopleTable, choresTable)
        End Using
    End Sub
    Private Sub AssignChores(Personnel As DataTable, chores As DataTable)
        ListBox1.Items.Clear()

        Dim rng As New Random()

        For Each choreRow As DataRow In chores.Rows
            Dim choreID As Integer = CInt(choreRow("Id"))
            'Dim choreDescription As String = choreRow("Description").ToString()
            Dim choreFrequency As String = choreRow("Frequency").ToString() ' Get the frequency
            Dim personRow As DataRow = Personnel.Rows(rng.Next(Personnel.Rows.Count))
            Dim personName As String = personRow("FirstName").ToString()


            ListBox1.Items.Add($"Chore: {choreID} | Assigned To: {personName} | Frequency: {choreFrequency}")
        Next
    End Sub
    Public Sub ScheduleNextChore(choresID As Integer)
        Try
            Dim query As String = "SELECT * FROM Chores WHERE ChoreID=@ChoreID"
            Using conn As New OleDbConnection(Module1.connectionString)
                Using cmd As New OleDbCommand(query, conn)
                    cmd.Parameters.AddWithValue("@ChoresID", choresID)

                    conn.Open()
                    Dim reader = cmd.ExecuteReader()
                    If reader.Read() Then
                        Dim isRecurring As Boolean = Convert.ToBoolean(reader("Recurring"))
                        Dim frequency As String = reader("Frequency").ToString()
                        Dim assignedTo As String = reader("AssignedTo").ToString()

                        If isRecurring Then
                            ' Calculate next due date
                            Dim nextDueDate As DateTime = Date.Today
                            Select Case frequency
                                Case "Daily"
                                    nextDueDate = Date.Today.AddDays(1)
                                Case "Weekly"
                                    nextDueDate = Date.Today.AddDays(7)
                                Case "Monthly"
                                    nextDueDate = Date.Today.AddMonths(1)
                            End Select

                            ' Assign next person in rotation
                            Dim nextPerson As String = GetNextPerson(assignedTo)

                            ' Update the chore for the next cycle
                            Dim updateQuery As String = "UPDATE CHORES SET DueDate=@NextDueDate, AssignedTo=@NextPerson, Status='Pending', LastCompleted=@LastCompleted WHERE ChoreID=@ChoreID"
                            Using updateCmd As New OleDbCommand(updateQuery, conn)
                                updateCmd.Parameters.AddWithValue("@NextDueDate", nextDueDate)
                                updateCmd.Parameters.AddWithValue("@NextPerson", nextPerson)
                                updateCmd.Parameters.AddWithValue("@LastCompleted", Date.Today)
                                updateCmd.Parameters.AddWithValue("@ChoreID", choresID)

                                updateCmd.ExecuteNonQuery()
                            End Using
                        End If
                    End If
                End Using
            End Using
        Catch ex As Exception
            MsgBox("Error scheduling next chore: " & ex.Message, MsgBoxStyle.Critical, "Database Error")
        End Try
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
                        Dim cmd As New OleDbCommand("DELETE FROM [Chores] WHERE [ID] = ?", conn)
                        cmd.Parameters.AddWithValue("@ID", choresID) ' Primary key for matching record  

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
                    Debug.WriteLine($"stack Trace:{ex.StackTrace}")
                    conn.Close()
                End Try
            End If
        Else
            MessageBox.Show("Please select an chores to delete.", "Deletion Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)

        End If
    End Sub
    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click

        DataGridView1.Sort(DataGridView1.Columns(0), System.ComponentModel.ListSortDirection.Ascending)
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Try

            For Each Rows As DataGridView In DataGridView1.Rows
                If row.Cells("DueDate").Value IsNot Nothing Then
                    Dim DueDate As DateTime = Convert.ToDateTime(row.Cells("DueDate").Value)
                    Dim Status As String = row.Cells("Status").Value.ToString

                    If DueDate < DateTime.Now AndAlso Status <> "Completed" Then
                        Rows.DefaultCellStyle.BackColor = Color.Red
                    Else

                        Rows.DefaultCellStyle.BackColor = Color.WhiteSmoke

                    End If

                End If
            Next

            Dim IncompletedCount As Integer = 0
            For Each row As DataGridView In DataGridView1.Rows
                If row.SelectedCells("Status").Value IsNot Nothing AndAlso row.SelectedCells("Status").Value.ToString() <> "Completed" Then

                    IncompletedCount += 1

                End If
            Next

            Label7.Text = "IncompletedCount chores:" & IncompletedCount.ToString
        Catch ex As Exception
            MessageBox.Show("Error highlighting overdue chores")
            Debug.WriteLine($"stack Trace:{ex.StackTrace}")

        End Try

    End Sub
    Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Close()
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        LoadData()
    End Sub
    Public Function GetNextPerson(currentPerson As String) As String
        ' Convert ComboBox items to a List of Strings  
        Dim people As New List(Of String)()
        For Each item In ComboBox1.Items
            people.Add(item.ToString())
        Next

        ' Check if the currentPerson exists in the list  
        Dim index As Integer = people.IndexOf(currentPerson)

        If index = -1 Then
            ' Person not found, return the first person or handle as needed  
            If people.Count > 0 Then
                Return people(0)
            Else
                Return String.Empty ' Return an empty string if the list is empty  
            End If
        Else
            ' Return the next person, wrapping around to the first if at the end  
            Return people((index + 1) Mod people.Count)
        End If
    End Function
    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        If DataGridView1.SelectedRows.Count = 0 Then
            MsgBox("Please select a chore to complete.", MsgBoxStyle.Exclamation, "No Selection")
            Return
        End If

        Dim selectedRow = DataGridView1.SelectedRows(0)
        Dim choresID As Integer

        ' Check if ChoreID is DBNull before conversion  
        If IsDBNull(selectedRow.Cells("ChoresID").Value) Then
            MsgBox("ChoreID is not available. Please select a valid chore.", MsgBoxStyle.Exclamation, "Invalid Chore")
            Return
        Else
            choresID = Convert.ToInt32(selectedRow.Cells("ChoresID").Value)
        End If

        Try
            ' Mark the task as completed  
            'UpdateChoreStatus(choreID, Date.Today)

            ' Automatically schedule the next instance of this chore  
            ScheduleNextChore(choresID)

            MsgBox("Chore completed & next task scheduled!", MsgBoxStyle.Information, "Success")
            'LoadChores()
        Catch ex As Exception
            MsgBox("Error updating chore: " & ex.Message, MsgBoxStyle.Critical, "Database Error")
        End Try
    End Sub


    'Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
    '    '    ' Get pending chores from the database
    '    Dim pendingChores As List(Of String) = GetPendingChores()

    '    ' Send an email if there are pending chores
    '    If pendingChores.Count > 0 Then
    '        Dim messageBody As String = "The following chores are still In progress:" & vbCrLf & String.Join(vbCrLf, pendingChores)
    '        SendEmail("austinmulalo113@gmail.com", "Chores Alert: Pending Tasks", messageBody)
    '    Else
    '        MessageBox.Show("No pending chores found.", "Chores Check", MessageBoxButtons.OK, MessageBoxIcon.Information)
    '    End If
    'End Sub

    'Private Function GetPendingChores() As List(Of String)
    '    Dim pendingChores As New List(Of String)

    '    Try
    '        Using conn As New OleDbConnection(connectionString)
    '            conn.Open()
    '            Dim query As String = "SELECT Title FROM CHORES WHERE Status = 'In Progress'"
    '            Using cmd As New OleDbCommand(query, conn)
    '                Using reader As OleDbDataReader = cmd.ExecuteReader()
    '                    While reader.Read()
    '                        pendingChores.Add(reader("Title").ToString())
    '                    End While
    '                End Using
    '            End Using
    '        End Using
    '    Catch ex As Exception
    '        MessageBox.Show("Database error: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
    '    End Try

    '    Return pendingChores
    'End Function

    'Private Sub SendEmail(toEmail As String, subject As String, body As String)
    '    Try
    '        Dim smtpClient As New SmtpClient("smtp.gmail.com") With {
    '                        .Port = 587,
    '                        .Credentials = New Net.NetworkCredential("austinmulalo113@gmail.com", "oqsa qwqa bhjc nzoe"), ' Use App Password for security
    '                        .EnableSsl = True
    '                    }

    '        Dim mail As New MailMessage() With {
    '            .From = New MailAddress("austinmulalo113gmail.com;"),
    '            .Subject = subject,
    '            .Body = body
    '        }
    '        mail.To.Add(toEmail)

    '        smtpClient.Send(mail)
    '        MessageBox.Show("Chores alert email sent successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
    '    Catch ex As Exception
    '        MessageBox.Show("Failed to send email: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
    '    End Try

    '    ' Get pending chores from the database
    '    Dim pendingChores As List(Of String) = GetPendingChores()

    '    ' Send an email if there are pending chores
    '    If pendingChores.Count > 0 Then
    '        Dim messageBody As String = "The following chores are still In progress:" & vbCrLf & String.Join(vbCrLf, pendingChores)
    '        SendEmail("austinmulalo113@gmail.com", "Chores Alert: Pending Tasks", messageBody)
    '    Else
    '        MessageBox.Show("No pending chores found.", "Chores Check", MessageBoxButtons.OK, MessageBoxIcon.Information)
    '    End If
    'End Sub

    'Private Function GetPendingChores() As List(Of String)
    '    Dim pendingChores As New List(Of String)

    '    Try
    '        Using conn As New OleDbConnection(connectionString)
    '            conn.Open()
    '            Dim query As String = "SELECT Title FROM CHORES WHERE Status = 'In Progress'"
    '            Using cmd As New OleDbCommand(query, conn)
    '                Using reader As OleDbDataReader = cmd.ExecuteReader()
    '                    While reader.Read()
    '                        pendingChores.Add(reader("Title").ToString())
    '                    End While
    '                End Using
    '            End Using
    '        End Using
    '    Catch ex As Exception
    '        MessageBox.Show("Database error: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
    '    End Try

    '    Return pendingChores
    'End Function

    'Private Sub SendEmail(toEmail As String, subject As String, body As String)
    '    Try
    '        Dim smtpClient As New SmtpClient("smtp.gmail.com") With {
    '                        .Port = 587,
    '                        .Credentials = New Net.NetworkCredential("austinmulalo113@gmail.com", "oqsa qwqa bhjc nzoe"), ' Use App Password for security
    '                        .EnableSsl = True
    '                    }

    '        Dim mail As New MailMessage() With {
    '            .From = New MailAddress("austinmulalo113@gmail.com"),
    '            .Subject = subject,
    '            .Body = body
    '        }
    '        mail.To.Add(toEmail)

    '        smtpClient.Send(mail)
    '        MessageBox.Show("Chores alert email sent successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
    '    Catch ex As Exception
    '        MessageBox.Show("Failed to send email: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
    '    End Try

    'End Sub
End Class