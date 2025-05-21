Imports System.IO
Imports System.Net.Mail
Imports System.Net
Imports System.Data.OleDb
Public Class Expense
    Dim conn As New OleDbConnection(HouseHoldManagment_Module.connectionString)
    ' Create a ToolTip object
    Private toolTip As New ToolTip()
    Private toolTip1 As New ToolTip()
    Private mealPlanData As DataTable

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Debug.WriteLine("Entering btnSubmit")
        Try


            Debug.WriteLine("User confirmed btnSubmit")
            Using conn As New OleDbConnection(HouseHoldManagment_Module.connectionString)
                conn.Open()


                'Define your unique criteria for an existing expense.
                'Adjust these fields based on what makes an expense unique.
                Dim checkQuery As String = "SELECT COUNT(*) FROM [Expense] WHERE [Amount] = ? AND [Description] = ? AND [DateOfexpenses] = ? AND [BillName] = ?"
                Dim checkCmd As New OleDbCommand(checkQuery, conn)
                checkCmd.Parameters.AddWithValue("@Amount", TextBox2.Text)
                checkCmd.Parameters.AddWithValue("@Description", TextBox6.Text)
                checkCmd.Parameters.AddWithValue("@DateOfexpenses", DateTimePicker1.Value)
                'checkCmd.Parameters.AddWithValue("@Person", ComboBox3.SelectedItem.ToString)
                checkCmd.Parameters.AddWithValue("@BillName", TextBox8.Text)

                Dim count As Integer = Convert.ToInt32(checkCmd.ExecuteScalar())

                If count > 0 Then
                    MessageBox.Show("This expense has already been saved.", "Duplicate Entry", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Return ' Exit without saving
                End If

                ' Update the table name if necessary  
                Dim tableName As String = "Expense"

                ' Create an OleDbCommand to insert the Expense data into the database 
                Dim cmd As New OleDbCommand("INSERT INTO [Expense] ([Amount], [Description], [Tags], [Category], [Paymentmethod], [Frequency], [DateOfexpenses], [Person], [BillName], [StartDate], [Recurring], [Paid]) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)", conn)

                ' Set the parameter values from the UI controls 
                'Class declaretions

                Dim expense As New Expensetracking With {
                    .Amount = TextBox2.Text,
                    .Description = TextBox6.Text,
                    .Tags = TextBox4.Text,
                    .Category = ComboBox7.SelectedItem.ToString,
                    .Paymentmethod = ComboBox1.SelectedItem.ToString,
                .Frequency = ComboBox5.SelectedItem.ToString(),
                    .DateOfexpenses = DateTimePicker1.Value,
                     .Person = ComboBox3.SelectedItem.ToString(),
                    .BillName = TextBox8.Text,
                    .StartDate = DateTimePicker2.Value,
                    .Recurring = CheckBox1.Checked,
                    .paid = ComboBox6.SelectedItem.ToString}

                'txtRecentUpdate.Text = $" Expense updated at {DateTime.Now:HH:MM}"

                cmd.Parameters.Clear()

                'cmd.Parameters.AddWithValue("@ExpenseID", expense.ExpenseID)
                cmd.Parameters.AddWithValue("@Amount", expense.Amount)
                cmd.Parameters.AddWithValue("@Description", expense.Description)
                cmd.Parameters.AddWithValue("@Tags", expense.Tags)
                cmd.Parameters.AddWithValue("@Category", expense.Category)
                cmd.Parameters.AddWithValue("@PaymentMethod", expense.Paymentmethod)
                cmd.Parameters.AddWithValue("@Frequency", expense.Frequency)
                cmd.Parameters.AddWithValue("@DateOfexpenses", expense.DateOfexpenses)
                cmd.Parameters.AddWithValue("@Person ", expense.Person)
                cmd.Parameters.AddWithValue("@BillName", expense.BillName)
                cmd.Parameters.AddWithValue("@StartDate", expense.StartDate)
                cmd.Parameters.AddWithValue("@Recurring", expense.Recurring)
                cmd.Parameters.AddWithValue("@Paid", expense.paid)


                MsgBox("Expense Information Saved!" & vbCrLf &
                        "ExpenseID: " & expense.ExpenseID & vbCrLf &
                        "Amount: " & expense.Amount & vbCrLf &
                        "Description: " & expense.Description & vbCrLf &
                        "Tags: " & expense.Tags & vbCrLf &
                        "Category: " & expense.Category & vbCrLf &
                        "PaymentMethod: " & expense.Paymentmethod & vbCrLf &
                         "Frequency: " & expense.Frequency & vbCrLf &
                          "Receiver: " & expense.Person & vbCrLf &
                            "DateOfExpense: " & expense.DateOfexpenses.ToString & vbCrLf &
                            "BillName: " & expense.BillName & vbCrLf &
                             "Recurring: " & expense.Recurring & vbCrLf &
                             "Paid: " & expense.paid & vbCrLf &
                          "StartDate: " & expense.StartDate.ToString, vbInformation, "Expense Confirmation")

                Dim ID As Integer
                Dim Amount As Integer
                SubtractFromBudget(ID, Amount)

                ' Execute the SQL command to insert the data 
                ' Log the SQL statement and parameter values  

                'MessageBox.Show("Expense information saved to Database successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)

                ' Execute the SQL command to insert the data  
                cmd.ExecuteNonQuery()
                'Next

            End Using
        Catch ex As OleDbException
            'MessageBox.Show($"Error saving Expense to database: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)

            Debug.WriteLine($" Database error in Button_Click: {ex.Message}")
            Debug.WriteLine($"Stack Trace : {ex.StackTrace}")
            MessageBox.Show("Error saving Expense to database: " & ex.Message & vbNewLine & ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)

            'MessageBox.Show("Error saving Expense to database.Please Check the connection.", "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Catch ex As Exception
            'MessageBox.Show("Unexpected Error: " & ex.Message & vbNewLine & ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Debug.WriteLine($" General error in Button: {ex.Message}")
            Debug.WriteLine($"Stack Trace : {ex.StackTrace}")
            Debug.WriteLine(" Failed to save")
            MessageBox.Show("An Unexpected error occured.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        conn.Close()
        LoadExpenseDataFromDatabase()
        Debug.WriteLine("Exiting btnSubmit")
    End Sub

    'Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
    '    Debug.WriteLine("Entering btnSubmit")
    '    Try
    '        Using conn As New OleDbConnection(HouseHoldManagment_Module.connectionString)
    '            conn.Open()

    '            ' Define your unique criteria for an existing expense.
    '            ' Adjust these fields based on what makes an expense unique.
    '            Dim checkQuery As String = "SELECT COUNT(*) FROM [Expense] WHERE [Amount] = ? AND [Description] = ? AND [DateOfexpenses] = ? AND [BillName] = ?"
    '            Dim checkCmd As New OleDbCommand(checkQuery, conn)
    '            checkCmd.Parameters.AddWithValue("@Amount", TextBox2.Text)
    '            checkCmd.Parameters.AddWithValue("@Description", TextBox6.Text)
    '            checkCmd.Parameters.AddWithValue("@DateOfexpenses", DateTimePicker1.Value)
    '            'checkCmd.Parameters.AddWithValue("@Person", ComboBox3.SelectedItem.ToString)
    '            checkCmd.Parameters.AddWithValue("@BillName", TextBox8.Text)

    '            Dim count As Integer = Convert.ToInt32(checkCmd.ExecuteScalar())

    '            If count > 0 Then
    '                MessageBox.Show("This expense has already been saved.", "Duplicate Entry", MessageBoxButtons.OK, MessageBoxIcon.Information)
    '                Return ' Exit without saving
    '            End If

    '            ' Proceed with insert if no duplicate found
    '            Dim cmd As New OleDbCommand("INSERT INTO [Expense] ([Amount], [Description], [Tags], [Category], [Paymentmethod], [Frequency], [DateOfexpenses], [Person], [BillName], [StartDate], [Recurring], [Paid]) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)", conn)

    '            Dim expense As New Expensetracking With {
    '            .Amount = TextBox2.Text,
    '            .Description = TextBox6.Text,
    '            .Tags = TextBox4.Text,
    '            .Category = ComboBox7.SelectedItem.ToString,
    '            .Paymentmethod = ComboBox1.SelectedItem.ToString,
    '            .Frequency = ComboBox5.SelectedItem.ToString(),
    '            .DateOfexpenses = DateTimePicker1.Value,
    '            .Person = ComboBox3.SelectedItem.ToString(),
    '            .BillName = TextBox8.Text,
    '            .StartDate = DateTimePicker2.Value,
    '            .Recurring = CheckBox1.Checked,
    '            .paid = ComboBox6.SelectedItem.ToString}

    '            cmd.Parameters.Clear()
    '            cmd.Parameters.AddWithValue("@Amount", expense.Amount)
    '            cmd.Parameters.AddWithValue("@Description", expense.Description)
    '            cmd.Parameters.AddWithValue("@Tags", expense.Tags)
    '            cmd.Parameters.AddWithValue("@Category", expense.Category)
    '            cmd.Parameters.AddWithValue("@PaymentMethod", expense.Paymentmethod)
    '            cmd.Parameters.AddWithValue("@Frequency", expense.Frequency)
    '            cmd.Parameters.AddWithValue("@DateOfexpenses", expense.DateOfexpenses)
    '            cmd.Parameters.AddWithValue("@Person ", expense.Person)
    '            cmd.Parameters.AddWithValue("@BillName", expense.BillName)
    '            cmd.Parameters.AddWithValue("@StartDate", expense.StartDate)
    '            cmd.Parameters.AddWithValue("@Recurring", expense.Recurring)
    '            cmd.Parameters.AddWithValue("@Paid", expense.paid)

    '            MsgBox("Expense Information Saved!", vbInformation, "Expense Confirmation")
    '            cmd.ExecuteNonQuery()

    '        End Using
    '    Catch ex As OleDbException
    '        Debug.WriteLine($" Database error in Button_Click: {ex.Message}")
    '        Debug.WriteLine($"Stack Trace : {ex.StackTrace}")
    '        MessageBox.Show("Error saving Expense to database: " & ex.Message & vbNewLine & ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
    '    Catch ex As Exception
    '        Debug.WriteLine($" General error in Button: {ex.Message}")
    '        Debug.WriteLine($"Stack Trace : {ex.StackTrace}")
    '        MessageBox.Show("An Unexpected error occurred.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
    '    End Try
    '    LoadExpenseDataFromDatabase()
    '    Debug.WriteLine("Exiting btnSubmit")
    'End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Debug.WriteLine("Entering btnEdit")

        ' Ensure a row is selected in the DataGridView  
        If DataGridView1.SelectedRows.Count = 0 Then

            Debug.WriteLine("User canceled btnEdit")
            MessageBox.Show("Please select a record to update.", "No Selection", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        Try

            Debug.WriteLine("User confirmed btnEdit")
            'Dim expenseId As Integer = Convert.ToInt32(selectedRow.Cells("ID").Value) ' Replace "ID" with your actual column name  
            Using conn As New OleDbConnection(HouseHoldManagment_Module.connectionString)
                conn.Open()

                ' Get the ID of the selected row (assuming your table has a primary key named "ID")  

                'Dim ExpenseID As String = TextBox1.Text
                Dim Amount As String = TextBox2.Text
                Dim Description As String = TextBox6.Text
                Dim Tags As String = TextBox4.Text
                Dim Category As String = ComboBox7.SelectedItem.ToString
                Dim Paymentmethod As String = ComboBox1.SelectedItem.ToString
                Dim Frequency As String = ComboBox5.SelectedItem.ToString()
                Dim Person As String = ComboBox3.SelectedItem.ToString()
                Dim DateOfexpenses As String = DateTimePicker1.Value
                Dim BillName As String = TextBox8.Text
                Dim StartDate As String = DateTimePicker2.Value
                Dim Recurring As String = CheckBox1.Checked
                Dim Paid As String = ComboBox6.SelectedItem.ToString

                ' Get the ID of the selected row (assuming your table has a primary key named "ID")  
                Dim selectedRow As DataGridViewRow = DataGridView1.SelectedRows(0)
                Dim ID As Integer = Convert.ToInt32(selectedRow.Cells("ID").Value) ' Change "ID" to your primary key column name  

                ' Create an OleDbCommand to update the Expense data in the database  
                Dim cmd As New OleDbCommand("UPDATE [Expense] SET [Amount] = ?, [Description] = ?, [Tags] = ?, [Category] = ?, [Paymentmethod] = ?, [DateOfexpenses] = ?, [Person] = ?, [BillName] = ?, [StartDate] = ?, [Recurring] = ?, [Paid] = ? WHERE [ID] = ?", conn)

                'cmd.Parameters.AddWithValue("@ExpenseID", ExpenseID)
                cmd.Parameters.AddWithValue("@Amount", Amount)
                cmd.Parameters.AddWithValue("@Description", Description)
                cmd.Parameters.AddWithValue("@Tags", Tags)
                cmd.Parameters.AddWithValue("@Category", Category)
                cmd.Parameters.AddWithValue("@PaymentMethod", Paymentmethod)
                cmd.Parameters.AddWithValue("Frequency", Frequency)
                cmd.Parameters.AddWithValue("Person ", Person)
                cmd.Parameters.AddWithValue("DateOfexpense", DateOfexpenses)
                cmd.Parameters.AddWithValue("BillName", BillName)
                cmd.Parameters.AddWithValue("StartDate", StartDate)
                cmd.Parameters.AddWithValue("Recurring", Recurring)
                cmd.Parameters.AddWithValue("Paid", Paid)
                cmd.Parameters.AddWithValue("ID", ID)

                cmd.ExecuteNonQuery()

                MsgBox("Expense information updated!")

                'Module1.ClearControls(Me)
                'LoadExpenseDataFromDatabase()

            End Using
        Catch ex As OleDbException

            'MessageBox.Show($"Error updating Expense in database: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Debug.WriteLine($"Stack Trace : {ex.StackTrace}")
            MessageBox.Show("Error saving Expense to database: " & ex.Message & vbNewLine & ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Catch ex As FormatException
            Debug.WriteLine($"Format error in btnEdit:{ex.Message}")
            Debug.WriteLine($"Stack Trace : {ex.StackTrace}")
            MessageBox.Show("Ensure all feilds are filled correctly.", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Catch ex As Exception
            Debug.WriteLine("Failed entering btnEdit_")
            MessageBox.Show("Error saving expense to database: " & ex.Message & vbNewLine & ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Debug.WriteLine($"Stack Trace : {ex.StackTrace}")
            Debug.WriteLine($"An  Error has occured when Editing data from Database")
        End Try
        LoadExpenseDataFromDatabase()
        Debug.WriteLine("Exited btnEdit")
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Debug.WriteLine("Entering btnDelete_")

        ' Check if there are any selected rows in the DataGridView for expenses  
        If DataGridView1.SelectedRows.Count > 0 Then

            ' Get the selected row  
            Dim selectedRow As DataGridViewRow = DataGridView1.SelectedRows(0)

            ' Assuming there is an "ID" column which is the primary key in the database  
            Dim ID As Integer = Convert.ToInt32(selectedRow.Cells("ID").Value) ' Replace "ID" with your actual column name  

            ' Confirm deletion  
            Dim confirmationResult As DialogResult = MessageBox.Show("Are you sure you want to delete this expense?", "Confirm Deletion", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)
            If confirmationResult = DialogResult.Yes Then
                ' Proceed with deletion  
                Try
                    TextBox7.Text = $" Expense updated at {DateTime.Now:HH:MM}"

                    Using conn As New OleDbConnection(HouseHoldManagment_Module.connectionString)
                        conn.Open()
                        Debug.WriteLine("User confirmed deletion")
                        ' Create the delete command  
                        Dim cmd As New OleDbCommand("DELETE FROM [Expense] WHERE [ID] = ?", conn)
                        cmd.Parameters.AddWithValue("@ID", ID) ' Primary key for matching record  

                        ' Execute the delete command  
                        Dim rowsAffected As Integer = cmd.ExecuteNonQuery()

                        If rowsAffected > 0 Then
                            MessageBox.Show("Expense deleted successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
                            ' Optionally refresh DataGridView or reload from database  

                            'LoadExpenseDataFromDatabase()

                            TextBox7.Text = $" Expense Delete at {DateTime.Now:HH:MM}"

                        Else
                            Debug.WriteLine(" User canceled deletion")
                            MessageBox.Show("No expense was deleted. Please check if the ID exists.", "Deletion Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                        End If
                    End Using

                Catch ex As FormatException
                    Debug.WriteLine($"Format error in btnDelete: {ex.Message}")
                    Debug.WriteLine($"Stack Trace : {ex.StackTrace}")
                    MessageBox.Show("Ensure you have selected a row correctly")
                Catch ex As Exception
                    MessageBox.Show($"An error occurred while deleting the expense: {ex.Message}", "Deletion Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Debug.WriteLine($"An error has occured when deleteing from database")
                    Debug.WriteLine($"Stack Trace : {ex.StackTrace}")
                End Try
            End If
        Else
            Debug.WriteLine($" No row  selected, exiting btnDelete")
            MessageBox.Show("Please select an expense to delete.", "Deletion Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End If
        LoadExpenseDataFromDatabase()

        Debug.WriteLine("Exiting deletion")
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        DataGridView1.Sort(DataGridView1.Columns("Amount"), System.ComponentModel.ListSortDirection.Descending)
        DataGridView1.Sort(DataGridView1.Columns("Currency"), System.ComponentModel.ListSortDirection.Ascending)

    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs)

        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox6.Text = ""
        TextBox4.Text = ""
        ComboBox7.SelectedItem = ""
        ComboBox1.SelectedItem = ""
        ComboBox5.SelectedItem = ""
        TextBox7.Text = ""
        ComboBox3.SelectedItem = ""
        TextBox8.Text = ""
        CheckBox1.Checked = ""
        ComboBox6.SelectedItem = ""
    End Sub

    Public Sub LoadExpenseDataFromDatabase()


        Try
            Debug.WriteLine("LoadExpenseDataFromDatabase")
            Using conn As New OleDbConnection(HouseHoldManagment_Module.connectionString)
                conn.Open()

                ' Update the table name if necessary  
                Dim tableName As String = "Expense"

                ' Create an OleDbCommand to select the data from the database  
                Dim cmd As New OleDbCommand($"SELECT * FROM {tableName}", conn)

                ' Create a DataAdapter and fill a DataTable  
                Dim da As New OleDbDataAdapter(cmd)
                Dim dt As New DataTable()
                da.Fill(dt)

                ' Bind the DataTable to the DataGridView  
                DataGridView1.DataSource = dt
            End Using
        Catch ex As OleDbException

            Debug.WriteLine($"Stack Trace : {ex.StackTrace}")
            Debug.WriteLine($"Error loading ExpenseDataFromDatabase : {ex.Message}")
            MessageBox.Show($"Error loading Expense data from database: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Catch ex As Exception
            Debug.WriteLine($"Stack Trace : {ex.StackTrace}")
            Debug.WriteLine($" General error in loading ExpenseDataFromDatabase: {ex.Message}")
            MessageBox.Show($"Unexpected error: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try


    End Sub
    Private Sub Expense_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Timer1.Interval = 3000
        Timer1.Enabled = True

        ' Initialize ToolTip properties (optional)
        toolTip.AutoPopDelay = 5000
        toolTip.InitialDelay = 500
        toolTip.ReshowDelay = 200
        toolTip.ShowAlways = True

        toolTip1.SetToolTip(Button5, "Sort")
        'toolTip1.SetToolTip(Button6, "Clear controls")
        toolTip1.SetToolTip(Button3, "Edit")
        toolTip1.SetToolTip(Button4, "Delete")
        toolTip1.SetToolTip(Button7, "Calculate Budget")
        toolTip1.SetToolTip(Button9, "Print to PDF")
        toolTip1.SetToolTip(Button1, "Save")
        toolTip1.SetToolTip(Button10, "Search")

        Try
            ' Dynamic label creation
            Dim lblDateTime As New Label()
            lblDateTime.Name = "lblDateTime"
            lblDateTime.Text = "Today: " & DateTime.Now.ToString("dddd, dd MMMM yyyy HH:mm")
            lblDateTime.Location = New Point(10, 10) ' Adjust position as needed
            lblDateTime.AutoSize = True
            lblDateTime.Font = New Font("Arial", 10, FontStyle.Bold)
            lblDateTime.ForeColor = Color.Blue


            ' Add the label to the forms controls
            Me.Controls.Add(lblDateTime)
            Debug.WriteLine("Expense Load: Datetime label added.")


            Debug.WriteLine("Form loaded successfully")
            ' Create a new OleDbConnection object and open the connection  

            conn.Open()

            ' Display the connection status on a button with a green background  
            Label17.Text = "Connected to Database"
            Label17.BackColor = Color.Green
            Label17.ForeColor = Color.White
        Catch ex As Exception
            ' Display the connection status on a button with a red background  
            Label17.Text = "Not Connected to Database"
            Label17.BackColor = Color.Red
            Label17.ForeColor = Color.White

            ' Display an error message  
            Debug.WriteLine(" Failed loading the Expense data  from DataBase")
            Debug.WriteLine($"Stack Trace : {ex.StackTrace}")
            MessageBox.Show("Error connecting to the database" & ex.Message, "Database Connection Error", MessageBoxButtons.OK, MessageBoxIcon.Error)

        Finally
            ' Close the database connection  
            conn.Close()
        End Try
        'PopulateMessagesFromDatabase()
        LoadExpenseDataFromDatabase()
        PopulateComboboxFromDatabase(ComboBox3)
    End Sub
    Public Sub PopulatelistboxFromDatabase(ByRef listbox As ListBox)

        Dim connect As New OleDbConnection(HouseHoldManagment_Module.connectionString)

        Try
            Debug.WriteLine("listbox populated successfully")
            ' 1. Open the database connection  
            connect.Open()

            ' 2. Retrieve the FirstName and LastName columns from the Personnel table  
            Dim query As String = "SELECT BillName, Amount, StartDate FROM ExpenseLogs"
            Dim cmd As New OleDbCommand(query, connect)
            Dim reader As OleDbDataReader = cmd.ExecuteReader()

            ' 3. Bind the retrieved data to the combobox  
            ListBox1.Items.Clear()
            While reader.Read()
                ListBox1.Items.Add($"{reader("BillName")} {reader("Amount")} {reader("StartDate")}")
            End While

            'increment = increment + 10
            'If increment > ProgressBar1.Maximum Then
            '    increment = ProgressBar1.Maximum
            'End If
            'ProgressBar1.Value = increment

            ' 4. Close the database connection  
            reader.Close()
        Catch ex As Exception
            ' Handle any exceptions that may occur  
            Debug.WriteLine("ComboBox population failed")
            Debug.WriteLine($" An error has occured when PopulateComboboxFromDatabase: {ex.Message}")
            Debug.WriteLine($"Stack Trace : {ex.StackTrace}")
            MessageBox.Show($"Error: {ex.Message}")
        Finally
            ' Close the database connection  
            If connect.State = ConnectionState.Open Then
                connect.Close()
            End If
        End Try
    End Sub
    Sub Mainn()
        Using conn As New OleDbConnection(connectionString)
            Try
                conn.Open()
                ' SQL query to select tasks where due date is today or earlier and status not yet updated
                Dim query As String = "SELECT ID, StartDate FROM Expense WHERE StartDate = ? AND Paid = No"
                Using command As New OleDbCommand(query, conn)
                    command.Parameters.AddWithValue("?", DateTime.Today)
                    'command.Parameters.AddWithValue("@OverdueStatus", "Overdue")

                    Using reader As OleDbDataReader = command.ExecuteReader()
                        Dim tasksToUpdate As New List(Of Integer)

                        While reader.Read()
                            Dim taskId As Integer = reader.GetInt32(0)
                            tasksToUpdate.Add(taskId)
                        End While

                        reader.Close()

                        ' Update each task's status to "Overdue"
                        For Each taskId As Integer In tasksToUpdate
                            Dim updateQuery As String = "UPDATE Expense SET Paid =  Yes WHERE ID = ?"
                            Using updateCmd As New OleDbCommand(updateQuery, conn)
                                updateCmd.Parameters.AddWithValue("Paid", "Yes")
                                updateCmd.Parameters.AddWithValue("?", taskId)
                                updateCmd.ExecuteNonQuery()

                                MessageBox.Show("Tasks updated successfully.")
                            End Using
                        Next
                    End Using
                End Using
            Catch ex As Exception
                MessageBox.Show("Error: " & ex.Message)
            End Try
        End Using
    End Sub
    Public Sub PopulateMessagesFromDatabase()
        Dim connect As New OleDbConnection(HouseHoldManagment_Module.connectionString)
        Try
            'Debug.WriteLine("listbox populated successfully")
            ' 1. Open the database connection  
            connect.Open()

            ' 2. Retrieve the FirstName and LastName columns from the Expense table  
            Dim query As String = "SELECT Amount, BillName, StartDate FROM Expense"
            Dim cmd As New OleDbCommand(query, connect)
            Dim reader As OleDbDataReader = cmd.ExecuteReader()
            ' 3. Bind the retrieved data to the combobox  
            While reader.Read()
                MessageBox.Show($"{reader("Amount")} {reader("BillName")} {reader("StartDate")}")
            End While
            ' 4. Close the database connection  
            reader.Close()
        Catch ex As Exception
            ' Handle any exceptions that may occur  
            Debug.WriteLine("MessageBox population failed")
            Debug.WriteLine($" An error has occured when PopulateMessageBoxFromDatabase: {ex.Message}")
            Debug.WriteLine($"Stack Trace : {ex.StackTrace}")
            MessageBox.Show($"Error: {ex.Message}")
        Finally
            ' Close the database connection  
            If connect.State = ConnectionState.Open Then
                connect.Close()
            End If
        End Try
    End Sub
    Public Sub PopulateComboboxFromDatabase(ByRef comboBox As ComboBox)
        Dim conn As New OleDbConnection(HouseHoldManagment_Module.connectionString)
        Try
            Debug.WriteLine("populate combobox successful")
            'open the database connection
            conn.Open()
            'retrieve the firstname and surname columns from the personaldetails tabel
            Dim query As String = "SELECT FirstName, LastName FROM PersonalDetails"
            Dim cmd As New OleDbCommand(query, conn)
            Dim reader As OleDbDataReader = cmd.ExecuteReader()
            'bind the retrieved data to the combobox
            ComboBox3.Items.Clear()
            While reader.Read()
                ComboBox3.Items.Add($"{reader("FirstName")} {reader("LastName")}")
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
    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Debug.WriteLine("Entering btnCalculate")

        ' Calculate total expenses  
        Dim totalExpenses As Decimal = Decimal.Parse(TextBox2.Text)
        Dim averageExpenses As Decimal = 0
        Try
            Debug.WriteLine("User confirmed btnCalculate")
            ' Calculate average expenses based on frequency  

            If ComboBox5.SelectedItem IsNot Nothing Then

                Select Case ComboBox5.SelectedItem.ToString()
                    Case "Monthly"
                        averageExpenses = totalExpenses / 12
                    Case "Weekly"
                        averageExpenses = totalExpenses / 52
                    Case "Daily"
                        averageExpenses = totalExpenses / 365
                    Case Else
                        ' Handle other frequencies as needed  
                End Select

            End If
            ' Display total and average expenses on the form  
            Label15.Text = $"Total Expense: R {totalExpenses:N2}"
            'Label16.Text = $" {ComboBox5.SelectedItem}: R {averageExpenses:N2}"

        Catch ex As FormatException
            Debug.WriteLine("Invalid format in Button6_Click: Amount should be in numbers.")
            Debug.WriteLine($"Stack Trace : {ex.StackTrace}")
            MessageBox.Show("Please enter a valid numeriv value for the Amount.", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)

        Catch ex As Exception
            Debug.WriteLine("Failed to calculate")
            Debug.WriteLine($"Stack Trace : {ex.StackTrace}")
            Debug.WriteLine($"Unexpected error in Button6_Click: {ex.Message}")
            MessageBox.Show("An unexpected error occured during calculations.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        Debug.WriteLine($"Calculation complete. Total:{totalExpenses},Avarage:{averageExpenses}")
        Debug.WriteLine("Exiting btnCalculate")

    End Sub
    Private Sub DataGridView1_SelectionChanged(sender As Object, e As EventArgs) Handles DataGridView1.SelectionChanged
        Try
            Debug.WriteLine("selecting data in the datagridview")
            If DataGridView1.SelectedRows.Count > 0 Then
                Dim selectedRow As DataGridViewRow = DataGridView1.SelectedRows(0)

                Button1.Enabled = False

                ' Load the data from the selected row into UI controls  
                TextBox2.Text = selectedRow.Cells("Amount").Value.ToString()
                TextBox6.Text = selectedRow.Cells("Description").Value.ToString()
                TextBox4.Text = selectedRow.Cells("Tags").Value.ToString()
                ComboBox7.SelectedItem = selectedRow.Cells("Category").Value.ToString()
                ComboBox1.SelectedItem = selectedRow.Cells("Paymentmethod").Value.ToString()
                ComboBox5.SelectedItem = selectedRow.Cells("Frequency").Value.ToString()
                DateTimePicker1.Text = selectedRow.Cells("Dateofexpenses").Value.ToString()
                ComboBox3.SelectedItem = selectedRow.Cells("Person").Value.ToString()
                TextBox8.Text = selectedRow.Cells("BillName").Value.ToString()
                DateTimePicker2.Text = selectedRow.Cells("StartDate").Value.ToString()
                CheckBox1.Text = selectedRow.Cells("Recurring").Value.ToString()
                ComboBox6.SelectedItem = selectedRow.Cells("Paid").Value.ToString()


            End If
        Catch ex As Exception
            Debug.WriteLine("error selection data in the database")
            Debug.WriteLine($"Stack Trace: {ex.StackTrace}")
        End Try
    End Sub
    Public Sub ClearControlss(ByVal FORM As Form)
        ' Clear TextBoxes  
        For Each ctrl As Control In FORM.Controls
            If TypeOf ctrl Is TextBox Then
                CType(ctrl, TextBox).Clear()
            End If
        Next
        ' Clear ComboBoxes  
        For Each ctrl As Control In FORM.Controls
            If TypeOf ctrl Is ComboBox Then
                CType(ctrl, ComboBox).ResetText()
            End If
        Next

        ' Clear DateTimePickers  
        For Each ctrl As Control In FORM.Controls
            If TypeOf ctrl Is DateTimePicker Then
                CType(ctrl, DateTimePicker).Value = DateTimePicker.MinimumDateTime ' or set to a specific date  
            End If
        Next
    End Sub
    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox6.Text = ""
        TextBox4.Text = ""
        ComboBox7.SelectedItem = ""
        ComboBox1.SelectedItem = ""
        ComboBox5.SelectedItem = ""
        TextBox7.Text = ""
        ComboBox3.SelectedItem = ""
        TextBox8.Text = ""
        'CheckBox1.Checked = ""
        ComboBox6.SelectedItem = ""
        HouseHoldManagment_Module.ClearControls(Me)

        LoadExpenseDataFromDatabase()
        Button1.Enabled = True
    End Sub
    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        PrintDialog1.Document = PrintDocument1
        If PrintDialog1.ShowDialog() = DialogResult.OK Then
            LoadExpenseDataFromDatabase() ' Load filtered data based on selected frequency
            If mealPlanData.Rows.Count > 0 Then
                PrintDocument1.Print()
            Else
                MessageBox.Show("No Expense found for the selected period.", "Print Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End If
        End If
    End Sub

    Private Sub ProcessRecurringExpenses()
        Dim conn As New OleDbConnection(connectionString)
        Try
            conn.Open()

            ' 1. Get all due recurring expenses
            Dim selectCmd As New OleDbCommand(
            "SELECT BillName, Amount, Frequency, StartDate, Description FROM Expense WHERE Recurring = TRUE AND Paid = 'No'", conn)

            Using reader As OleDbDataReader = selectCmd.ExecuteReader()
                While reader.Read()
                    Dim billName As String = reader("BillName").ToString()
                    Dim amount As Decimal = CDec(reader("Amount"))
                    Dim frequency As String = reader("Frequency").ToString()
                    Dim startDate As DateTime = CDate(reader("StartDate"))
                    Dim description As String = reader("Description").ToString()

                    ' Determine if expense is due based on frequency and start date
                    Dim nextPaymentDate As DateTime = startDate
                    Dim isDue As Boolean = False

                    Select Case frequency
                        Case "Daily"
                            isDue = (startDate <= DateTime.Today)
                            nextPaymentDate = startDate.AddDays(1)
                        Case "Weekly"
                            isDue = (startDate <= DateTime.Today)
                            nextPaymentDate = startDate.AddDays(7)
                        Case "Monthly"
                            isDue = (startDate <= DateTime.Today)
                            nextPaymentDate = startDate.AddMonths(1)
                        Case "Annually"
                            isDue = (startDate <= DateTime.Today)
                            nextPaymentDate = startDate.AddYears(1)
                        Case Else
                            ' Unknown frequency, skip
                            Continue While
                    End Select

                    If isDue Then
                        ' 2. Insert a record into ExpenseLogs for the payment
                        Dim insertCmd As New OleDbCommand(
                        "INSERT INTO ExpenseLogs (BillName, Amount, Recurring, Frequency, StartDate, DateOfExpenses, Description, Paid) VALUES (?, ?, ?, ?, ?, ?, ?, ?)", conn)

                        insertCmd.Parameters.AddWithValue("?", billName)
                        insertCmd.Parameters.AddWithValue("?", amount)
                        insertCmd.Parameters.AddWithValue("?", False) ' Paid - false since it's just paid now
                        insertCmd.Parameters.AddWithValue("?", frequency)
                        insertCmd.Parameters.AddWithValue("?", startDate)
                        insertCmd.Parameters.AddWithValue("?", DateTime.Today)
                        insertCmd.Parameters.AddWithValue("?", "Auto-paid: " & description)
                        insertCmd.Parameters.AddWithValue("?", "Yes") ' Mark as paid

                        insertCmd.ExecuteNonQuery()

                        ' 3. Update the original expense's StartDate to the next payment date
                        Dim updateCmd As New OleDbCommand(
                        "UPDATE Expense SET StartDate = ? WHERE BillName = ? AND Recurring = TRUE AND Paid = 'No'", conn)
                        updateCmd.Parameters.AddWithValue("?", nextPaymentDate)
                        updateCmd.Parameters.AddWithValue("?", billName)

                        updateCmd.ExecuteNonQuery()

                        ' Optional: Add feedback or logging here
                    End If
                End While
            End Using

            MessageBox.Show("Recurring expenses processed successfully.")

        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message)
        Finally
            If conn.State = ConnectionState.Open Then conn.Close()
        End Try
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Timer1.Enabled = False
        DisplayDataInMessageBox()
        'Main()
        'SaveChangedDateToAnotherTable()
        'PopulatelistboxFromDatabase(ListBox1)
    End Sub
    Sub Main()
        Using connection As New OleDbConnection(connectionString)
            Try
                connection.Open()

                ' SQL query to select tasks where StartDate is today or earlier and not yet paid
                Dim query As String = "SELECT ID, StartDate FROM Expense WHERE StartDate <= ? AND Paid = 'No'" 'Corrected condition

                Using command As New OleDbCommand(query, connection)
                    command.Parameters.AddWithValue("?", DateTime.Today)

                    Using reader As OleDbDataReader = command.ExecuteReader()
                        Dim tasksToUpdate As New List(Of Integer)

                        While reader.Read()
                            tasksToUpdate.Add(reader.GetInt32(reader.GetOrdinal("ID")))
                        End While

                        reader.Close() 'Close the reader after the loop

                        If tasksToUpdate.Count > 0 Then
                            ' SQL to update the StartDate in another table (e.g., UpdatedExpenses)
                            Dim updateQuery As String = "UPDATE UpdatedExpenses SET StartDate = ?" &
                                                    " WHERE ExpenseID IN (" & String.Join(",", tasksToUpdate.Select(Function(x) x).ToArray()) & ")"

                            Using updateCommand As New OleDbCommand(updateQuery, connection)
                                'Crucial: Add a parameter for the new StartDate
                                updateCommand.Parameters.AddWithValue("?", DateTime.Today)
                                updateCommand.ExecuteNonQuery()
                            End Using
                        End If
                    End Using
                End Using
            Catch ex As Exception
                Console.WriteLine("Error: " & ex.Message)
                ' Important: Log the error for debugging
                ' Add more robust error handling, such as logging to a file
                ' or showing a user-friendly message.
            Finally
                If connection.State = ConnectionState.Open Then connection.Close()
            End Try
        End Using
    End Sub
    'Private Sub SaveChangedDateToAnotherTable()
    '    Dim selectQuery As String = "SELECT * FROM Expense" ' Your source table
    '    Dim dt As New DataTable()

    '    Using conn As New OleDbConnection(connectionString)
    '        Try
    '            conn.Open()

    '            ' Fill DataTable with source data
    '            Using cmd As New OleDbCommand(selectQuery, conn)
    '                Using adapter As New OleDbDataAdapter(cmd)
    '                    adapter.Fill(dt)
    '                End Using
    '            End Using
    '            ' Loop through each row to modify and insert into target table
    '            For Each row As DataRow In dt.Rows
    '                ' Read the 'Paid' status
    '                Dim paidStatus As String = Convert.ToString(row("Recurring")).Trim().ToLower()

    '                ' Proceed only if 'Paid' is "no"
    '                If paidStatus = "True" Then
    '                    ' Prepare the new change date for next payment
    '                    Dim currentStartDate As DateTime = Convert.ToDateTime(row("StartDate"))
    '                    Dim nextPaymentDate As DateTime = currentStartDate.AddMonths(1) ' or your logic for next payment

    '                    ' Insert into the target table, e.g., ExpenseLogs
    '                    Dim insertQuery As String = "INSERT INTO ExpenseLogs (BillName, Amount, StartDate, Paid, Recurring) VALUES (?, ?, ?, ?, ?)"
    '                    Using insertCmd As New OleDbCommand(insertQuery, conn)
    '                        insertCmd.Parameters.AddWithValue("?", row("BillName"))
    '                        insertCmd.Parameters.AddWithValue("?", row("Amount"))
    '                        insertCmd.Parameters.AddWithValue("?", nextPaymentDate)
    '                        insertCmd.Parameters.AddWithValue("?", "Yes") ' or your logic for setting Paid
    '                        insertCmd.Parameters.AddWithValue("?", row("Recurring")) ' assuming same recurring

    '                        Try
    '                            insertCmd.ExecuteNonQuery()
    '                        Catch ex As Exception
    '                            MessageBox.Show("Error inserting row: " & ex.Message)
    '                        End Try
    '                    End Using
    '                Else
    '                    ' Optionally, you can handle the case where Paid = "Yes"
    '                    ' For example, log or ignore
    '                    'MessageBox.Show("Payments that are not Recurring were not paid " & DateTime.Now.ToString())
    '                End If
    '            Next
    '            'MessageBox.Show("Payments with updated dates saved successfully at " & DateTime.Now.ToString())
    '            'UpdateDatesBasedOnFrequency()
    '            'Mainm()
    '        Catch ex As Exception
    '            MessageBox.Show("Error fetching data: " & ex.Message)
    '        End Try
    '    End Using
    'End Sub

    Private Sub SaveChangedDateToAnotherTable()
        Dim getBudgetQuery As String = "SELECT CurrentAmount FROM Budget"
        Dim updateBudgetQuery As String = "UPDATE Budget SET CurrentAmount = ?"
        Dim dt As New DataTable()
        ' Start a transaction to make sure all changes succeed or fail together
        Dim transaction As OleDbTransaction = conn.BeginTransaction()

        Try
            ' Get current budget
            Dim currentBudget As Decimal
            Using budgetCmd As New OleDbCommand(getBudgetQuery, conn, transaction)
                Dim result = budgetCmd.ExecuteScalar()
                If result IsNot Nothing Then
                    currentBudget = Convert.ToDecimal(result)
                Else
                    Throw New Exception("Budget not found.")
                End If
            End Using

            ' Loop through each row to modify and insert into target table
            For Each row As DataRow In dt.Rows
                Dim paidStatus As String = Convert.ToString(row("Recurring")).Trim().ToLower()
                If paidStatus = "true" Then
                    Dim amount As Decimal = Convert.ToDecimal(row("Amount"))

                    ' Deduct from budget
                    currentBudget -= amount
                    Using updateCmd As New OleDbCommand(updateBudgetQuery, conn, transaction)
                        updateCmd.Parameters.AddWithValue("?", currentBudget)
                        updateCmd.ExecuteNonQuery()
                    End Using

                    ' Insert into ExpenseLogs
                    Dim currentStartDate As DateTime = Convert.ToDateTime(row("StartDate"))
                    Dim nextPaymentDate As DateTime = currentStartDate.AddMonths(1)

                    Dim insertQuery As String = "INSERT INTO ExpenseLogs (BillName, Amount, StartDate, Paid, Recurring) VALUES (?, ?, ?, ?, ?)"
                    Using insertCmd As New OleDbCommand(insertQuery, conn, transaction)
                        insertCmd.Parameters.AddWithValue("?", row("BillName"))
                        insertCmd.Parameters.AddWithValue("?", amount)
                        insertCmd.Parameters.AddWithValue("?", nextPaymentDate)
                        insertCmd.Parameters.AddWithValue("?", "Yes")
                        insertCmd.Parameters.AddWithValue("?", row("Recurring"))

                        insertCmd.ExecuteNonQuery()
                    End Using
                End If
            Next

            ' Commit all changes
            transaction.Commit()
            MessageBox.Show("Payments saved and budget updated successfully.")
        Catch ex As Exception
            transaction.Rollback()
            MessageBox.Show("Transaction failed: " & ex.Message)
        End Try
    End Sub
    Private Sub DisplayDataInMessageBox()
        Dim query As String = "SELECT * FROM Expense" ' Fetch all records

        Using conn As New OleDbConnection(connectionString)
            Dim command As New OleDbCommand(query, conn)
            Try
                conn.Open()
                Using reader As OleDbDataReader = command.ExecuteReader()
                    Dim today As Date = DateTime.Now.Date
                    Dim dataList As New List(Of String)
                    While reader.Read()
                        Dim BillName As Object = reader("BillName")
                        Dim Amount As Object = reader("Amount")
                        Dim StartDateObj As Object = reader("StartDate")
                        Dim StartDate As Date

                        ' Attempt to parse StartDate
                        If Date.TryParse(StartDateObj.ToString(), StartDate) Then
                            If StartDate = today Then
                                dataList.Add($"This Payments is Due to be Paid : BillName: {BillName}, Amount: {Amount}, StartDate: {StartDate.ToShortDateString()}")
                            End If
                        End If
                    End While

                    If dataList.Count = 0 Then
                        MessageBox.Show("No Payments are scheduled for today.")
                        Return
                    End If

                    ' Display each record in a message box
                    For Each Datas In dataList
                        Dim result As DialogResult
                        result = MessageBox.Show(Datas, "Confirmation", MessageBoxButtons.YesNo)

                        If result = DialogResult.Yes Then
                            Mainm()
                            SaveChangedDateToAnotherTable()
                            PopulatelistboxFromDatabase(ListBox1)
                            LoadExpenseDataFromDatabase()
                            Dim ID As Integer
                            Dim Amount As Integer
                            SubtractFromBudget(ID, Amount)
                            MessageBox.Show("Payments with updated dates saved successfully at " & DateTime.Now.ToString())
                        Else
                            MessageBox.Show("Payments were cancelled.")
                        End If
                    Next
                End Using
            Catch ex As Exception
                MessageBox.Show("Error: " & ex.Message)
            End Try
        End Using
    End Sub
    'Private Sub DisplayDataInMessageBox()
    '    Dim query As String = "SELECT * FROM Expense " ' Replace with your table name

    '    Using conn As New OleDbConnection(connectionString)
    '        Dim command As New OleDbCommand(query, conn)
    '        Try
    '            conn.Open()
    '            Using reader As OleDbDataReader = command.ExecuteReader()
    '                Dim dataList As New List(Of String)
    '                While reader.Read()
    '                    ' Assuming your table has columns named "ID" and "Name"
    '                    Dim BillName As Object = reader("BillName")
    '                    Dim Amount As Object = reader("Amount")
    '                    Dim StartDate As Object = reader("StartDate")
    '                    dataList.Add($"This item is Due to be Paid : BillName: {BillName}, Amount: {Amount}, StartDate: {StartDate}")
    '                End While

    '                ' Display each record in a message box
    '                For Each Datas In dataList
    '                    Dim result As DialogResult

    '                    ' Create a custom message box 
    '                    result = MsgBox(Datas, MessageBoxButtons.YesNo)

    '                    If result = DialogResult.Yes Then ' Perform saving actions
    '                        Mainm()
    '                        SaveChangedDateToAnotherTable()
    '                        PopulatelistboxFromDatabase(ListBox1)
    '                        LoadExpenseDataFromDatabase()
    '                        'UpdateDatesBasedOnFrequency()
    '                        MessageBox.Show("Payments with updated dates saved successfully at " & DateTime.Now.ToString())
    '                        'MessageBox.Show("Payments with updated dates saved successfully") 'Example

    '                    Else ' Perform actions for No 
    '                        MessageBox.Show("Payments were cancelled.")
    '                        'Example
    '                    End If

    '                Next
    '            End Using
    '        Catch ex As Exception
    '            MessageBox.Show("Error: " & ex.Message)
    '        End Try
    '    End Using
    'End Sub
    Sub Mainm()
        Using conn As New OleDbConnection(connectionString)
            conn.Open()

            Dim selectAllCmd As New OleDbCommand("SELECT ID, StartDate, Frequency FROM Expense", conn)

            Using reader As OleDbDataReader = selectAllCmd.ExecuteReader()
                While reader.Read()
                    Dim recordId As Integer = reader("ID")
                    Dim currentDateObj As Object = reader("StartDate")
                    Dim frequencyStrObj As Object = reader("Frequency")

                    If currentDateObj IsNot Nothing AndAlso Not Convert.IsDBNull(currentDateObj) AndAlso
                   frequencyStrObj IsNot Nothing AndAlso Not Convert.IsDBNull(frequencyStrObj) Then

                        Dim dateValue As DateTime
                        Dim frequencyStr As String = frequencyStrObj.ToString()

                        If DateTime.TryParse(currentDateObj.ToString(), dateValue) Then
                            Dim daysToAdd As Integer = GetDaysFromFrequency(frequencyStr)

                            If daysToAdd > 0 Then
                                Dim newDate As DateTime = dateValue.AddDays(daysToAdd)

                                ' Update the record
                                Dim updateCmd As New OleDbCommand("UPDATE Expense SET StartDate = ? WHERE ID = ?", conn)
                                updateCmd.Parameters.AddWithValue("?", newDate)
                                updateCmd.Parameters.AddWithValue("?", recordId)
                                updateCmd.ExecuteNonQuery()
                            Else
                                MessageBox.Show($"Unknown frequency '{frequencyStr}' for record ID {recordId}. Skipping.")
                            End If
                        Else
                            MessageBox.Show($"Invalid StartDate for record ID {recordId}. Skipping.")
                        End If
                    End If
                End While
            End Using
        End Using
    End Sub

    ' Helper function to convert string frequency to days
    Function GetDaysFromFrequency(freq As String) As Integer
        Select Case freq.Trim().ToLower()
            Case "daily"
                Return 1
            Case "weekly"
                Return 7
            Case "monthly"
                Return 30
            Case "yearly"
                Return 365
                ' Add more cases as needed
            Case Else
                Return 0 ' Unknown frequency
        End Select
    End Function

    'Sub Mainm()
    '    ' The ID or unique identifier for the record you want to update
    '    Dim recordId As Integer = "?"

    '    Using conn As New OleDbConnection(connectionString)
    '        conn.Open()

    '        ' Step 1: Retrieve current date and frequency from the database
    '        Dim selectCmd As New OleDbCommand("SELECT StartDate, Frequency FROM Expense WHERE ID = ?", conn)
    '        selectCmd.Parameters.AddWithValue("?", recordId)

    '        Using reader As OleDbDataReader = selectCmd.ExecuteReader()
    '            If reader.Read() Then
    '                Dim currentDateObj As Object = reader("StartDate")
    '                Dim frequencyObj As Object = reader("Frequency")

    '                If currentDateObj IsNot Nothing AndAlso Not Convert.IsDBNull(currentDateObj) AndAlso
    '               frequencyObj IsNot Nothing AndAlso Not Convert.IsDBNull(frequencyObj) Then

    '                    Dim dateValue As DateTime = Convert.ToDateTime(currentDateObj)
    '                    Dim frequencyDays As Integer = Convert.ToInt32(frequencyObj)

    '                    ' Step 2: Add days based on frequency
    '                    Dim newDate As DateTime = dateValue.AddDays(frequencyDays)

    '                    ' Step 3: Update the database with the new date
    '                    Dim updateCmd As New OleDbCommand("UPDATE Expense SET StartDate = ? WHERE ID = ?", conn)
    '                    updateCmd.Parameters.AddWithValue("?", newDate)
    '                    updateCmd.Parameters.AddWithValue("?", recordId)

    '                    Dim rowsAffected As Integer = updateCmd.ExecuteNonQuery()

    '                    If rowsAffected > 0 Then
    '                        MessageBox.Show("Date updated successfully.")
    '                    Else
    '                        MessageBox.Show("No record updated.")
    '                    End If
    '                Else
    '                    MessageBox.Show("Record not found or date/frequency is null.")
    '                End If
    '            Else
    '                MessageBox.Show("Record not found.")
    '            End If
    '        End Using
    '    End Using
    'End Sub
    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        Try
            Dim SearchTerm As String = TextBox9.Text.ToLower

            For Each row As DataGridViewRow In DataGridView1.Rows
                If Not row.IsNewRow Then
                    Dim Amount As String = row.Cells("Amount").Value.ToString().ToLower
                    Dim TotalIncome As String = row.Cells("TotalIncome").Value.ToString().ToLower
                    Dim Description As String = row.Cells("Description").Value.ToString().ToLower
                    Dim Tags As String = row.Cells("Tags").Value.ToString().ToLower
                    Dim Category As String = row.Cells("Category").Value.ToString().ToLower
                    Dim Paymentmethod As String = row.Cells("Paymentmethod").Value.ToString().ToLower
                    Dim Frequency As String = row.Cells("Frequency").Value.ToString().ToLower
                    Dim Person As String = row.Cells("Person").Value.ToString().ToLower
                    Dim DateOfexpenses As String = row.Cells("DateOfexpenses").Value.ToString.ToLower
                    Dim BillName As String = row.Cells("BillName").Value.ToString().ToLower
                    Dim StartDate As String = row.Cells("StartDate").Value.ToString().ToLower
                    Dim Recurring As String = row.Cells("Recurring").Value.ToString().ToLower
                    Dim paid As String = row.Cells("Paid").Value.ToString.ToLower

                    row.Visible = BillName.Contains(SearchTerm) & Category.Contains(SearchTerm) & Amount.Contains(SearchTerm) & Description.Contains(SearchTerm)
                End If
            Next
        Catch ex As Exception
            Debug.WriteLine($" Error in Search Functionality: {ex.Message}")
            MessageBox.Show("Try to Search Again.", "Search Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Button6_Click_1(sender As Object, e As EventArgs) Handles Button6.Click
        ' Example: get ID and Amount from TextBoxes
        Dim ID As Integer
        Dim Amount As Integer

        SubtractFromBudget(ID, Amount)


    End Sub
    Public Sub SubtractFromBudget(ID As Integer, Amount As Integer)
        ' SQL query to update the Amount by subtracting the specified value
        Dim query As String = "UPDATE Budget SET Amount = Amount - ? WHERE ID = ?"

        Using conn As New OleDbConnection(connectionString)
            Using command As New OleDbCommand(query, conn)
                ' Add parameters to prevent SQL injection
                command.Parameters.AddWithValue("?", Amount)
                command.Parameters.AddWithValue("?", ID)

                Try
                    conn.Open()
                    Dim rowsAffected As Integer = command.ExecuteNonQuery()
                    If rowsAffected > 0 Then
                        MessageBox.Show("Budget updated successfully.")
                    Else
                        'MessageBox.Show("No record found with the specified ID.")
                    End If
                Catch ex As Exception
                    MessageBox.Show("Error updating budget: " & ex.Message)
                End Try
            End Using
        End Using
    End Sub
End Class

