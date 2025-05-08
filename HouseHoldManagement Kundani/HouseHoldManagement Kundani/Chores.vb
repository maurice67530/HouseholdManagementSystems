Imports System.Data.OleDb
Public Class Chores
    Public Property conn As New OleDbConnection(Rinae.connectionString)
    Dim tooltip As New ToolTip
    'Public Property connn As New OleDbConnection(Masindi.connectionString)
    'Public Property conn As New OleDbConnection(Murangi.connectionString)
    Private Sub chores_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Dim connection As New OleDbConnection(Murangi.connectionString)

        tooltip.SetToolTip(Button1, "Save")
        tooltip.SetToolTip(Button2, "Edit")
        tooltip.SetToolTip(Button4, "Delete")
        tooltip.SetToolTip(Button5, "Refresh")
        tooltip.SetToolTip(Button9, "Mark All As complete")
        tooltip.SetToolTip(Button3, "Dashboard")

        ' Check database connectivity 

        Try
            Debug.WriteLine("form load initialized successfully")

            ' Create a new OleDbConnection object and open the connection  

            connection.Open()

            ' Display the connection status on a button with a green background  
            'Label11.Text = "Connected"
            'Label11.BackColor = Color.Green
            'Label11.ForeColor = Color.White



        Catch ex As Exception

            ' Display an error message  
            Debug.WriteLine($"Failed to initialize components from database: {ex.Message}")
            Debug.WriteLine($"Stack Trace: {ex.StackTrace}")
            MessageBox.Show("Error connecting to the database" & ex.Message, "Database Connection Error", MessageBoxButtons.OK, MessageBoxIcon.Error)

            ' Display the connection status on a button with a red background  
            'Label11.Text = "Not Connected"
            'Label11.BackColor = Color.Red
            'Label11.ForeColor = Color.White



        Finally


            PopulateComboboxFromDatabase(cmbassi)
            loadChoresFromDatabase()
            ' Close the database connection  
            connection.Close()
        End Try

        'chores
        LoadChores()
        CheckRecurringChores()
        CheckPendingChores() ' Check pending chores when the form opens



    End Sub
    Private Sub LoadChores()
        Using conn As New OleDbConnection(Murangi.connectionString)
            Dim query As String = "SELECT Title FROM Chores"
            Dim cmd As New OleDbCommand(query, conn)
            Dim adapter As New OleDbDataAdapter(cmd)
            Dim table As New DataTable()
            adapter.Fill(table)

            cmbChore.DataSource = table
            cmbChore.ValueMember = "Title"


        End Using
    End Sub
    Private Sub CheckPendingChores()

        Using con As New OleDbConnection(Murangi.connectionString)
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
    Private Sub CheckRecurringChores()


        Using con As New OleDbConnection(Murangi.connectionString)
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
    Private Sub Button6_Click(sender As Object, e As EventArgs)
        conn.Open()

        'Try
        Dim chores As New Chores_() With {
           .Title = TXTtitle.Text,
           .AssignedTo = cmbassi.SelectedItem,
            .Priority = cmbpri.SelectedItem,
            .Status = cmbstatus.SelectedItem,
            .Frequency = cmbfre.SelectedItem,
            .DueDate = DateTimePicker1.Text,
            .Recurring = NumericUpDown1.Value,
            .Description = txtdes.Text}

        Dim tablename As String = "Chores"
        Dim Cmd As New OleDbCommand($"INSERT INTO {tablename} ([Title], [AssignedTo], [Priority], [Status], [Frequency], [DueDate], [Recurring], [Description]) VALUES (?, ?, ?, ?, ?, ?, ?, ?)", conn)

        Cmd.Parameters.Clear()

        Cmd.Parameters.AddWithValue("@Title", chores.Title)
        Cmd.Parameters.AddWithValue("@AssignedTo", chores.AssignedTo)
        Cmd.Parameters.AddWithValue("@Priority", chores.Priority)
        Cmd.Parameters.AddWithValue("@Status", chores.Status)
        Cmd.Parameters.AddWithValue("@Frequency", chores.Frequency)
        Cmd.Parameters.AddWithValue("@DueDate", chores.DueDate)
        Cmd.Parameters.AddWithValue("@Recurring", chores.Recurring)
        Cmd.Parameters.AddWithValue("@Description", chores.Description)

        MsgBox("chores Information Addded!" & vbCrLf &
              "Title: " & chores.Title & vbCrLf &
              "AssignedTo:" & chores.Description & vbCrLf &
              "Priority: " & chores.Priority & vbCrLf &
              "Status : " & chores.Status & vbCrLf &
              "Frequency: " & chores.AssignedTo & vbCrLf &
              "Recurring: " & chores.Recurring & vbCrLf &
              "Description: " & chores.Description & vbCrLf &
              "DueDate: " & chores.DueDate & vbCrLf & vbCrLf, vbInformation, "Chores confirmation")

        MessageBox.Show("Chores Information saved to Database successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Cmd.ExecuteNonQuery()

        'Catch ex As OleDbException
        '    Debug.WriteLine($"General error in button Save: {ex.Message}")
        '    Debug.WriteLine($"Stack Trace: {ex.StackTrace}")
        '    MessageBox.Show("Error saving control texts: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        '    MessageBox.Show($"Error Saving To database: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        '    MessageBox.Show("Error saving Chores to database: " & ex.Message & vbNewLine & ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)

        'Catch ex As Exception
        '    MessageBox.Show($"Unexpected error: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        'End Try
        'Debug.WriteLine("Existing button Save")


        'Add to FamilySchedule
        ' Ndamu.AddChoreEvent(TXTtitle.Text, cmbassi.SelectedItem.ToString, DateTimePicker1.Text)
    End Sub
    Private Sub Button7_Click(sender As Object, e As EventArgs)
        Try
            For Each row As DataGridViewRow In DataGridView1.Rows
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

            For Each row As DataGridViewRow In DataGridView1.Rows
                If row.Cells("Status").Value IsNot Nothing AndAlso row.Cells("Status").Value.ToString() <> "Completed " Then

                    InclompletedCount += 1

                End If
            Next
            Label10.Text = "Incompleted chores: " & InclompletedCount.ToString
        Catch ex As Exception
            MessageBox.Show("Error highligting overdue chores")
        End Try
    End Sub
    Private Sub Button4_Click(sender As Object, e As EventArgs)
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
                    Dim confirmationResults As DialogResult = MessageBox.Show("Are you sure you want o delete this chore?", "Confirm Deletion", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)
                    If confirmationResults = DialogResult.Yes Then
                        Debug.WriteLine("User confirmed deletion.")

                    Else
                        Debug.WriteLine("User cancelled deletion.")
                    End If

                Else
                    Debug.WriteLine("No row selected, Exiting Button delete")

                    MessageBox.Show("Please Select chore to delete", "Deletion Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                End If

                Using conn As New OleDbConnection(Rinae.connectionString)
                    conn.Open()


                    ' Create the delete command  
                    Dim cmd As New OleDbCommand("DELETE FROM [Chores] WHERE [ID] = ?", conn)
                    cmd.Parameters.AddWithValue("@ID", ID) ' Primary key for matching record  

                    ' Execute the delete command  
                    Dim rowsAffected As Integer = cmd.ExecuteNonQuery()

                    If rowsAffected > 0 Then
                        MessageBox.Show("Task deleted successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        ' Optionally refresh DataGridView or reload from database  
                        loadChoresFromDatabase()
                    Else
                        MessageBox.Show("No chores deleted. Please check if the ID exists.", "Deletion Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    End If

                End Using

            Catch ex As Exception
                MessageBox.Show($"An error occurred while deleting the expense: {ex.Message}", "Deletion Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        Else

            MessageBox.Show("Please select chores to delete.", "Deletion Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)

            loadChoresFromDatabase()
        End If
    End Sub
    Public Sub loadChoresFromDatabase()
        Try

            Debug.WriteLine("Datagridview: populated successfully.")

            Using conn As New OleDbConnection(Rinae.connectionString)
                conn.Open()

                Dim tableName As String = "Chores"

                Dim cmd As New OleDbCommand($"SELECT*FROM {tableName}", conn)

                Dim da As New OleDbDataAdapter(cmd)
                Dim dt As New DataTable
                da.Fill(dt)

                DataGridView1.DataSource = dt

            End Using

        Catch ex As Exception
            Debug.WriteLine($"Datagridview: Failed to Populate {ex.Message}")
            Debug.WriteLine($"Stack Trace: {ex.StackTrace}")
            MessageBox.Show($"Error Loading Inventory data from database: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub Button3_Click(sender As Object, e As EventArgs)
        loadChoresFromDatabase()
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs)
        For Each row As DataGridViewRow In
            DataGridView1.SelectedRows
            row.Cells("Status").Value = "Completed"
        Next
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs)
        Close()
    End Sub
    Private Sub Button5_Click(sender As Object, e As EventArgs)
        Try

            Debug.WriteLine("entering button update")

            If DataGridView1.SelectedRows.Count = 0 Then
                MessageBox.Show("Please select a record to update.", "No Selection", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            If DataGridView1.SelectedRows.Count > 0 Then
                Debug.WriteLine("A row is selected for update")

            Else
                MessageBox.Show("Please select a Chore to update.", "update Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Debug.WriteLine("No row selected, exiting Button Update")
            End If
            Debug.WriteLine("Exiting Button Update")

            Dim Title As String = TXTtitle.Text
            Dim AssignedTo As String = cmbassi.SelectedItem
            Dim Priority As String = cmbpri.SelectedItem
            Dim Status As String = cmbstatus.SelectedItem
            Dim Frequency As String = cmbfre.SelectedItem
            Dim DueDate As String = DateTimePicker1.Value
            Dim Recurring As String = NumericUpDown1.Value
            Dim Description As String = txtdes.Text

            Using conn As New OleDbConnection(Rinae.connectionString)
                conn.Open()

                'Get the ID of the selected row (assuming your table has a primary key named "ID")  
                Dim selectedRow As DataGridViewRow = DataGridView1.SelectedRows(0)
                Dim ID As Integer = Convert.ToInt32(selectedRow.Cells("ID").Value) ' Change "ID" to your primary key column name  

                'Create an OleDbCommand to update the personnel data in the database  
                Dim cmd As New OleDbCommand("UPDATE Chores SET [Title] = ?, [AssignedTo] = ?, [Priority] = ?, [Status] = ?, [Frequency] = ?, [DueDate] = ?, [Recurring]= ?, [Description] = ? WHERE [ID] = ?", conn)

                'Set the parameter values from the UI controls  


                cmd.Parameters.AddWithValue("@Title", TXTtitle.Text)
                cmd.Parameters.AddWithValue("@AssignedTo", cmbassi.SelectedItem)
                cmd.Parameters.AddWithValue("@Priority", cmbpri.SelectedItem)
                cmd.Parameters.AddWithValue("@Status", cmbstatus.SelectedItem)
                cmd.Parameters.AddWithValue("@Frequency", cmbfre.SelectedItem)
                cmd.Parameters.AddWithValue("@DueDate", DateTimePicker1.Value)
                cmd.Parameters.AddWithValue("@Recurring", NumericUpDown1.Value)
                cmd.Parameters.AddWithValue("@Description", txtdes.Text)
                cmd.Parameters.AddWithValue("@ID", ID)
                MsgBox("Chores Updated Successfuly!", vbInformation, "Update Confirmation")

                cmd.ExecuteNonQuery()

            End Using
        Catch ex As FormatException
            Debug.WriteLine($"Format Error in Button Update: {ex.Message}")
            Debug.WriteLine($"Stack Trace: {ex.StackTrace}")
            MessageBox.Show($"Error updating Tasks in database: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show("Error saving inventory to database: " & ex.Message & vbNewLine & ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Catch ex As Exception
            Debug.WriteLine($"unexpected error Button Update: {ex.Message}")
            Debug.WriteLine($"Stack Trace: {ex.StackTrace}")
            MessageBox.Show($"Unexpected error: {ex.Message}", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show("Error saving inventory to database: " & ex.Message & vbNewLine & ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub Button9_Click(sender As Object, e As EventArgs)
        DataGridView1.Sort(DataGridView1.Columns("DueDate"), System.ComponentModel.ListSortDirection.Ascending)
    End Sub
    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs)

    End Sub
    Private Sub DataGridView1_SelectionChanged(sender As Object, e As EventArgs)
        Try

            'Button1.Enabled = False

            If DataGridView1.SelectedRows.Count > 0 Then
                Dim selectedRow As DataGridViewRow = DataGridView1.SelectedRows(0)
                Debug.WriteLine("row selected on dgv")


                TXTtitle.Text = selectedRow.Cells("Title").Value.ToString()
                cmbassi.SelectedItem = selectedRow.Cells("AssignedTo").Value.ToString()
                cmbpri.SelectedItem = selectedRow.Cells("Priority").Value.ToString()
                cmbstatus.SelectedItem = selectedRow.Cells("Status").Value.ToString()
                cmbfre.SelectedItem = selectedRow.Cells("Frequency").Value.ToString()
                NumericUpDown1.Value = selectedRow.Cells("Recurring").Value.ToString()
                txtdes.Text = selectedRow.Cells("Description").Value.ToString()
                DateTimePicker1.Value = selectedRow.Cells("DueDate").Value.ToString()


            End If
            'Button1.Enabled = True
        Catch ex As Exception
            Debug.WriteLine("Data not selected: Error")
            Debug.WriteLine($"Stack Trace: {ex.StackTrace}")

        End Try
    End Sub
    Public Sub PopulateComboboxFromDatabase(ByRef comboBox As ComboBox)
        Dim conn As New OleDbConnection(Rinae.connectionString)
        Try
            Debug.WriteLine("populate combobox successful")
            'open the database connection
            conn.Open()

            'retrieve the firstname and surname columns from the personaldetails tabel
            Dim query As String = "SELECT FirstName, LastName FROM PersonalDetails"
            Dim cmd As New OleDbCommand(query, conn)
            Dim reader As OleDbDataReader = cmd.ExecuteReader()

            'bind the retrieved data to the combobox
            cmbassi.Items.Clear()
            While reader.Read()
                cmbassi.Items.Add($"{reader("FirstName")} {reader("LastName")}")
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

    Private Sub Button8_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub cmbassi_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbassi.SelectedIndexChanged

    End Sub

    Private Sub Button9_Click_1(sender As Object, e As EventArgs) Handles Button9.Click

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
        If DataGridView1.SelectedRows.Count > 0 Then
            Return Convert.ToInt32(DataGridView1.SelectedRows(0).Cells("ID").Value)
        End If
        Return 0 ' Return 0 if no row is selected
    End Function
    Private Sub UpdateChoreStatus(choreID As Integer, status As String)
        Dim connString As String = Murangi.connectionString
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
    ' Method to complete the selected chore and auto-assign the next available person
    Private Sub CompleteChore(choreID As Integer)
        Dim connString As String = Murangi.connectionString
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



    Private Function GetNextAvailablePerson() As String
        Dim connString As String = Murangi.connectionString
        Dim availablePerson As String = String.Empty
        Dim assignedPeople As New List(Of String)

        ' Get already assigned people from DataGridView  
        For Each row As DataGridViewRow In DataGridView1.Rows
            If row.Cells("AssignedTo").Value IsNot Nothing Then
                assignedPeople.Add(row.Cells("AssignedTo").Value.ToString())
            End If
        Next

        ' Get an available person from the database who is NOT in DataGridView or already assigned  
        Using conn As New OleDb.OleDbConnection(connString)
            Try
                conn.Open()
                Dim query As String = "SELECT TOP 1 FirstName, LastName FROM PersonalDetails " &
                                  "WHERE FirstName + ' ' + LastName NOT IN (" &
                                  String.Join(",", assignedPeople.Select(Function(p) "'" & p & "'")) & ") " &
                                  "ORDER BY FirstName, LastName"

                Using cmd As New OleDb.OleDbCommand(query, conn)
                    Dim reader As OleDb.OleDbDataReader = cmd.ExecuteReader()
                    If reader.Read() Then
                        availablePerson = reader("FirstName").ToString() & " " & reader("LastName").ToString()
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
        Dim connString As String = Murangi.connectionString
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

End Class